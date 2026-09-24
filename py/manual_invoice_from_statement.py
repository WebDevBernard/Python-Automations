#!/usr/bin/env python3
"""
Fill "Manual Invoice Template.pdf" (Policy + Transactions layout) from the
Horizon West CUSTOMER STATEMENT PDF(s) found among the last 2 modified PDFs
in Downloads.

A PDF is only treated as a statement if the words CUSTOMER STATEMENT appear
inside STATEMENT_RECT on page 1.

Mapping (statement -> invoice):
    Date            -> Policy "Term From" (Term To = Term From + 1 year);
                       shown on the first policy row only when all match
    Transaction     -> Transaction
    Description     -> Description
    Amount          -> Amount
    Policy          -> Policy Number (one policy row per distinct policy)
    Invoice/Cheque  -> Invoice Number (comma-separated if rows differ)
    Statement date  -> Invoice Date

Output: Desktop if it exists, otherwise the current working directory.

    pip install pymupdf

Usage:
    python manual_invoice_from_statement.py                # normal run
    python manual_invoice_from_statement.py --list-fields  # dump the template's form fields
    python manual_invoice_from_statement.py --debug        # also print parsed data
    python manual_invoice_from_statement.py --statement S.pdf [--template T.pdf] [--out DIR]
"""

import argparse
import re
import sys
from datetime import date
from pathlib import Path

try:
    import pymupdf
except ImportError:  # older PyMuPDF versions
    import fitz as pymupdf

from constants import get_insurer
from utils import unique_file_name

TEMPLATE = Path(r"E:\dev\Python-Automations\py\assets\Manual Invoice Template.pdf")
DOWNLOADS = Path.home() / "Downloads"
SCAN_COUNT = 2

# (x0, y0, x1, y1) in points, top-left origin (PyMuPDF convention)
STATEMENT_RECT = (
    179.52000427246094,
    104.70088195800781,
    396.876953125,
    124.75080108642578,
)
RECT_PAD = 1.5  # small tolerance so glyphs on the edge aren't clipped

# Statement dates look like 05/11/2026. "DMY" reads that as 5 November 2026
# (a renewal due after the September statement date); use "MDY" for May 11.
DATE_ORDER = "DMY"

# Template capacity (must match the template's buttons)
PMAX, TMAX = 5, 10
KMAX = PMAX + TMAX - 1  # transaction row "slots"

PHONE_RE = re.compile(r"\(\d{3}\)\s*\d{3}-\d{4}")
# 'VANCOUVER, BC V5R 3J8 WONC05 BY NT PT' -> city line + customer code line
CITY_CODE_RE = re.compile(
    r"^(?P<city>.+?[A-Z]\d[A-Z]\s?\d[A-Z]\d)"
    r"(?:\s+(?P<ref>(?P<code>[A-Z]{2,}\d{2,})\b.*))?"
)
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
MONEY_RE = re.compile(r"^-?\$?[\d,]+\.\d{2}$")

# Statement column headers -> field; data starts a little left of each header
HEADERS = [  # (header word, key, left margin)
    ("Transaction", "transaction", 8),
    ("Invoice", "invoice", 8),
    ("Policy", "policy", 12),
    ("Description", "description", 6),
]
DEFAULT_BOUNDS = {"transaction": 55, "invoice": 130, "policy": 176, "description": 290}


# --------------------------------------------------------------------------- #
# Statement detection + parsing
# --------------------------------------------------------------------------- #
def group_lines(words, ytol: float = 3.0) -> list:
    """Group words into visual lines (lists of words, left to right)."""
    words = sorted(words, key=lambda w: (w[1] + w[3]) / 2)
    lines, cur, ref = [], [], 0.0
    for w in words:
        yc = (w[1] + w[3]) / 2
        if cur and yc - ref > ytol:
            lines.append(cur)
            cur = []
        if not cur:
            ref = yc
        cur.append(w)
    if cur:
        lines.append(cur)
    return [sorted(l, key=lambda w: w[0]) for l in lines]


def statement_title_text(path: Path) -> str:
    x0, y0, x1, y1 = STATEMENT_RECT
    clip = pymupdf.Rect(x0 - RECT_PAD, y0 - RECT_PAD, x1 + RECT_PAD, y1 + RECT_PAD)
    with pymupdf.open(path) as doc:
        if doc.page_count == 0:
            return ""
        return doc[0].get_text("text", clip=clip)


def is_customer_statement(path: Path) -> bool:
    text = re.sub(r"\s+", " ", statement_title_text(path)).strip().upper()
    return "CUSTOMER STATEMENT" in text


def format_name(s: str) -> str:
    """Client name, formatted like the renewal letter: single spaces, no
    colons, title case ('SMITH, JOHN & JANE' -> 'Smith, John & Jane')."""
    return re.sub(r"\s+", " ", s).strip().replace(":", "").title()


def money(s: str) -> float:
    return float(s.replace("$", "").replace(",", ""))


def parse_date(s: str) -> date:
    a, b, y = (int(x) for x in s.split("/"))
    return date(y, b, a) if DATE_ORDER == "DMY" else date(y, a, b)


def long_date(d: date) -> str:
    return f"{d:%B} {d.day}, {d.year}"  # 'November 5, 2026'


def plus_one_year(d: date) -> date:
    try:
        return d.replace(year=d.year + 1)
    except ValueError:  # Feb 29 -> Feb 28
        return d.replace(year=d.year + 1, day=28)


def column_bounds(words) -> dict:
    """Left edge of each column, taken from the table headers."""
    bounds = dict(DEFAULT_BOUNDS)
    for word, key, margin in HEADERS:
        hits = [w for w in words if w[4] == word and 180 < w[1] < 240]
        if hits:
            bounds[key] = min(h[0] for h in hits) - margin
    return bounds


def parse_rows(words) -> list:
    """Transaction rows: any line starting with a dd/dd/yyyy date."""
    b = column_bounds(words)
    items = []
    for line in group_lines(words):
        if not DATE_RE.match(line[0][4]) or line[0][0] >= b["transaction"]:
            continue
        amt_words = [
            w for w in line if MONEY_RE.match(w[4]) and w[0] > b["description"]
        ]
        if not amt_words:
            continue
        amt = amt_words[-1]
        cols = {"transaction": [], "invoice": [], "policy": [], "description": []}
        for w in line[1:]:
            if w is amt:
                continue
            x = w[0]
            key = (
                "description"
                if x >= b["description"]
                else (
                    "policy"
                    if x >= b["policy"]
                    else "invoice" if x >= b["invoice"] else "transaction"
                )
            )
            cols[key].append(w[4])
        items.append(
            dict(
                date=parse_date(line[0][4]),
                amount=money(amt[4]),
                **{k: " ".join(v) for k, v in cols.items()},
            )
        )
    return items


def parse_statement(path: Path) -> dict:
    with pymupdf.open(path) as doc:
        items = [i for page in doc for i in parse_rows(page.get_text("words"))]
        lines = [
            " ".join(w[4] for w in l) for l in group_lines(doc[0].get_text("words"))
        ]
        flat = "\n".join(lines)

    m = re.search(r"\b([A-Z][a-z]+ \d{1,2}, \d{4})\b", flat)
    statement_date = m.group(1) if m else ""

    name = phone = street = city = code = ref = ""
    for i, l in enumerate(lines):
        if not l.startswith("To:"):
            continue
        to_line = l[3:]
        m = PHONE_RE.search(to_line)
        if m:
            phone = m.group(0)
            to_line = to_line[: m.start()] + to_line[m.end() :]
        name = format_name(to_line)
        if i + 1 < len(lines):
            street = lines[i + 1].rstrip(",")
        if i + 2 < len(lines):
            m = CITY_CODE_RE.match(lines[i + 2])
            if m:
                city, code = m["city"], m["code"] or ""
                ref = (m["ref"] or "").strip()  # 'SETS01 BY NT PT'
        break

    m = re.search(r"Customer Code:\s*(\S+)", flat)
    if m and not code:
        code = m.group(1)

    m = re.search(r"Outstanding Balance\s*:\s*(-?\$?[\d,]+\.\d{2})", flat)
    calc = round(sum(i["amount"] for i in items), 2)
    if m and abs(money(m.group(1)) - calc) > 0.005:
        print(
            f"  ! Warning: rows sum to ${calc:,.2f} but statement balance is {m.group(1)}"
        )

    return dict(
        name=name,
        phone=phone,
        street=street,
        city=city,
        customer_code=code,
        customer_ref=ref or code,
        statement_date=statement_date,
        items=items,
        total=calc,
    )


# --------------------------------------------------------------------------- #
# Invoice data
# --------------------------------------------------------------------------- #
def company_for(policy: str) -> str:
    """Company Name isn't printed on the statement, so derive it from the
    policy number prefix (rules live in constants.get_insurer). No exact
    match -> most likely of Wawanesa / Intact / Aviva."""
    return get_insurer(policy, guess=True)


def build_invoice(data: dict) -> dict:
    items = data["items"]
    if len(items) > TMAX:
        print(
            f"  ! Template holds {TMAX} transactions; statement has {len(items)} - extra lines NOT written."
        )
        items = items[:TMAX]

    policies = {}  # policy number -> earliest due date, in order of appearance
    for it in items:
        p = it["policy"]
        if p and (p not in policies or it["date"] < policies[p]):
            policies[p] = it["date"]
    pol_rows = [
        dict(company=company_for(p), number=p, term_from=d, term_to=plus_one_year(d))
        for p, d in policies.items()
    ]
    if len(pol_rows) > PMAX:
        print(
            f"  ! Template holds {PMAX} policies; statement has {len(pol_rows)} - extra policies NOT written."
        )
        pol_rows = pol_rows[:PMAX]

    invoices = list(dict.fromkeys(it["invoice"] for it in items if it["invoice"]))
    return dict(policies=pol_rows, items=items, invoice_number=", ".join(invoices))


def field_values(data: dict, inv: dict) -> dict:
    """{field: (stored value, text shown)} - numbers are stored bare so the
    form's own $ formatting and totals keep working in Acrobat."""
    p, t = max(len(inv["policies"]), 1), max(len(inv["items"]), 1)
    amt = lambda x: (f"{x:.2f}", f"${x:,.2f}")
    same = lambda s: (s, s)
    total = round(sum(i["amount"] for i in inv["items"]), 2)
    stub_code = data["customer_ref"].split()[0] if data["customer_ref"] else ""

    v = {
        "inv_num": same(inv["invoice_number"]),
        "inv_date": same(data["statement_date"]),
        "cust_name": same(data["name"]),
        "cust_addr1": same(data["street"]),
        "cust_addr2": same(data["city"]),
        "cust_phone": same(data["phone"]),
        "cust_code": same(data["customer_ref"]),
        "st_pol": same(str(p)),
        "st_txn": same(str(t)),
        f"tot_{p + t - 1}": amt(total),
        # stub fields are calculated by the form in Acrobat; set them too so
        # they show in every viewer
        "stub_name": same(data["name"]),
        "stub_addr1": same(data["street"]),
        "stub_addr2": same(data["city"]),
        "stub_code": same(stub_code),
        "stub_invnum": same(inv["invoice_number"]),
        "stub_date": same(data["statement_date"]),
        "stub_policy": same(", ".join(r["number"] for r in inv["policies"])),
        "stub_due": amt(total),
    }
    # every policy on the same term -> dates only on the first row
    one_term = len({(r["term_from"], r["term_to"]) for r in inv["policies"]}) == 1
    for r, pol in enumerate(inv["policies"], 1):
        v[f"co_{r}"] = same(pol["company"])
        v[f"pnum_{r}"] = same(pol["number"])
        if r == 1 or not one_term:
            v[f"tfrom_{r}"] = same(long_date(pol["term_from"]))
            v[f"tto_{r}"] = same(long_date(pol["term_to"]))
    for j, it in enumerate(inv["items"], 1):
        k = p + j - 1  # transaction row j sits in slot k (see template buttons)
        v[f"tx_type_{k}"] = same(it["transaction"])
        v[f"tx_desc_{k}"] = same(it["description"])
        v[f"tx_amt_{k}"] = amt(it["amount"])
    return {k: val for k, val in v.items() if val[0] != ""}


def visibility(p: int, t: int) -> dict:
    """{field: visible?} - the same layout the template's buttons draw."""
    vis = {}
    for r in range(1, PMAX + 1):
        for f in (
            "plL",
            "plR",
            "plB",
            "pdiv1",
            "pdiv2",
            "pdiv3",
            "co",
            "pnum",
            "tfrom",
            "tto",
        ):
            vis[f"{f}_{r}"] = r <= p
        for f in (
            "txh_bg",
            "txhT",
            "txhB",
            "txhL",
            "txhR",
            "txh_l1",
            "txh_l3",
            "txh_l4",
        ):
            vis[f"{f}_{r}"] = r == p
    for k in range(1, KMAX + 1):
        for f in ("tlL", "tlR", "tlB", "tx_type", "tx_desc", "tx_amt"):
            vis[f"{f}_{k}"] = p <= k <= p + t - 1
        vis[f"tot_lbl_{k}"] = vis[f"tot_{k}"] = k == p + t - 1
    return vis


# --------------------------------------------------------------------------- #
# Filling
# --------------------------------------------------------------------------- #
def fill(template: Path, out_path: Path, data: dict, inv: dict):
    values = field_values(data, inv)
    vis = visibility(max(len(inv["policies"]), 1), max(len(inv["items"]), 1))
    with pymupdf.open(template) as doc:
        cat = doc.pdf_catalog()
        hebo = doc.xref_get_key(cat, "AcroForm/DR/Font/HeBo")[1]
        for page in doc:
            # 1) values: generate the visible text, then store the bare value
            for w in page.widgets():
                if w.field_name not in values:
                    continue
                stored, shown = values[w.field_name]
                da = doc.xref_get_key(w.xref, "DA")[1]
                w.field_value = shown
                w.update()
                doc.xref_set_key(w.xref, "V", pymupdf.get_pdf_str(stored))
                if "/HeBo" in da:  # update() switches to plain Helv; restore bold
                    doc.xref_set_key(w.xref, "DA", pymupdf.get_pdf_str(da))
                    ap = int(doc.xref_get_key(w.xref, "AP/N")[1].split()[0])
                    doc.update_stream(
                        ap, doc.xref_stream(ap).replace(b"/Helv", b"/HeBo")
                    )
                    doc.xref_set_key(ap, "Resources/Font/HeBo", hebo)
            # 2) show/hide rows by annotation flag only (keeps the drawn table lines intact)
            for w in page.widgets():
                if w.field_name in vis:
                    doc.xref_set_key(w.xref, "F", "4" if vis[w.field_name] else "2")
        doc.save(out_path, garbage=3, deflate=True)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def list_fields(template: Path):
    with pymupdf.open(template) as doc:
        for pno, page in enumerate(doc, 1):
            for w in sorted(
                page.widgets(), key=lambda w: (round(w.rect.y0), w.rect.x0)
            ):
                print(
                    f"p{pno}  x={w.rect.x0:6.1f} y={w.rect.y0:6.1f}  "
                    f"{w.field_type_string:<9} {w.field_name!r}  value={w.field_value!r}"
                )


def output_dir() -> Path:
    for d in (Path.home() / "Desktop", Path.home() / "OneDrive" / "Desktop"):
        if d.is_dir():
            return d
    return Path.cwd()


def manual_invoice(
    config_data=None,
    debug: bool = False,
    statements=None,
    template: Path = TEMPLATE,
    out_dir: Path = None,
) -> int:
    """Entry point for file_completion_tool. Returns the number of invoices made."""
    if not template.is_file():
        print(f"Template not found: {template}")
        return 0

    if statements:
        pdfs = [Path(s) for s in statements]
    else:
        pdfs = sorted(
            DOWNLOADS.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True
        )[:SCAN_COUNT]
    if not pdfs:
        print(f"No PDFs in {DOWNLOADS}")
        return 0

    out_dir = out_dir or output_dir()
    made = 0
    for pdf in pdfs:
        if debug:
            print(
                f"----- {pdf.name} -----\nTitle region text: {statement_title_text(pdf)!r}"
            )
        if not is_customer_statement(pdf):
            print(f"Skipping {pdf.name} (no 'CUSTOMER STATEMENT' in title area)")
            continue

        print(f"Processing {pdf.name}")
        data = parse_statement(pdf)
        if not data["items"]:
            print("  ! No transaction lines parsed - try --debug")
            continue
        inv = build_invoice(data)
        if debug:
            for k in (
                "name",
                "phone",
                "street",
                "city",
                "customer_ref",
                "statement_date",
            ):
                print(f"  {k}: {data[k]!r}")
            for pol in inv["policies"]:
                print(f"  policy: {pol}")
            for it in inv["items"]:
                print(f"  item: {it}")

        client = re.sub(r'[\\/:*?"<>|]', "", data["name"]).strip()
        out = Path(
            unique_file_name(
                str(Path(out_dir) / f"{client or 'Unknown Client'} Invoice.pdf")
            )
        )
        fill(template, out, data, inv)
        print(f"  -> {out}")
        made += 1

    if made:
        print(f"\nCreated {made} invoice(s) in {out_dir}")
    else:
        print(
            f"No customer statement found among the last {SCAN_COUNT} modified PDFs in {DOWNLOADS}"
        )
    return made


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--list-fields", action="store_true")
    ap.add_argument("--debug", action="store_true")
    ap.add_argument(
        "--statement", nargs="+", help="statement PDF(s) instead of scanning Downloads"
    )
    ap.add_argument("--template", type=Path, default=TEMPLATE)
    ap.add_argument("--out", type=Path, help="output folder (default: Desktop)")
    args = ap.parse_args()

    if args.list_fields:
        return list_fields(args.template)
    if not manual_invoice(
        debug=args.debug,
        statements=args.statement,
        template=args.template,
        out_dir=args.out,
    ):
        sys.exit(1)


if __name__ == "__main__":
    main()
