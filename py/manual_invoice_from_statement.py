#!/usr/bin/env python3
"""
Fill "Manual Invoice Template.pdf" from the Horizon West CUSTOMER STATEMENT
PDF(s) found among the last 2 modified PDFs in Downloads.

A PDF is only treated as a statement if the words CUSTOMER STATEMENT appear
inside STATEMENT_RECT on page 1.

Output: Desktop if it exists, otherwise the current working directory.

    pip install pymupdf

Usage:
    python manual_invoice_from_statement.py                # normal run
    python manual_invoice_from_statement.py --list-fields  # dump the template's form fields
    python manual_invoice_from_statement.py --debug        # also print title-region and extracted text
"""

import argparse
import re
import sys
from pathlib import Path

try:
    import pymupdf
except ImportError:  # older PyMuPDF versions
    import fitz as pymupdf

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

# Statement columns, left to right:
#   Date | Transaction | Invoice/Cheque | Policy | Description | Amount
ROW_RE = re.compile(
    r"^(?P<date>\d{2}/\d{2}/\d{4})\s+"
    r"(?P<transaction>[A-Za-z][A-Za-z ]*?)\s+"
    r"(?P<invoice>\d{4,7})\s+"
    r"(?P<policy>[A-Z][A-Z0-9]{6,})\s+"
    r"(?P<description>.*?)\s*"
    r"(?P<amount>-?\$?[\d,]+\.\d{2})$"
)

PHONE_RE = re.compile(r"\(\d{3}\)\s*\d{3}-\d{4}")
# 'VANCOUVER, BC V5R 3J8 WONC05 BY NT PT' -> city line + customer code
CITY_CODE_RE = re.compile(
    r"^(?P<city>.+?[A-Z]\d[A-Z]\s?\d[A-Z]\d)"
    r"(?:\s+(?P<ref>(?P<code>[A-Z]{2,}\d{2,})\b.*))?"
)

# Template form layout: detail rows are date_N, txn_N, ... (N = 1..MAX_ROWS).
# Rows 2+ start hidden; each row count N has its own 'Invoice Total' group
# (ob_box_N / ob_label_N / total_N) and only the one for the last row shows.
MAX_ROWS = 18
DISPLAY_VISIBLE, DISPLAY_HIDDEN = 0, 1  # Widget.field_display values
ROW_FIELDS = {
    "date": "date",
    "txn": "transaction",
    "inv": "invoice",
    "pol": "policy",
    "desc": "description",
    "amt": "amount",
}


# --------------------------------------------------------------------------- #
# Statement detection + parsing
# --------------------------------------------------------------------------- #
def page_lines(page, ytol: float = 3.0) -> list:
    """Rebuild visual text lines from words so each table row is one string."""
    words = sorted(page.get_text("words"), key=lambda w: (w[1] + w[3]) / 2)
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
    return [" ".join(w[4] for w in sorted(l, key=lambda w: w[0])) for l in lines]


def read_text(path: Path) -> str:
    with pymupdf.open(path) as doc:
        return "\n".join("\n".join(page_lines(p)) for p in doc)


def statement_title_text(path: Path) -> str:
    """Text found inside STATEMENT_RECT on page 1."""
    x0, y0, x1, y1 = STATEMENT_RECT
    clip = pymupdf.Rect(x0 - RECT_PAD, y0 - RECT_PAD, x1 + RECT_PAD, y1 + RECT_PAD)
    with pymupdf.open(path) as doc:
        if doc.page_count == 0:
            return ""
        return doc[0].get_text("text", clip=clip)


def is_customer_statement(path: Path) -> bool:
    text = re.sub(r"\s+", " ", statement_title_text(path)).strip().upper()
    return "CUSTOMER STATEMENT" in text


def money(s: str) -> float:
    return float(s.replace("$", "").replace(",", ""))


def parse_statement(text: str) -> dict:
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    flat = "\n".join(lines)

    m = re.search(r"\b([A-Z][a-z]+ \d{1,2}, \d{4})\b", flat)
    statement_date = m.group(1) if m else ""

    # 'To: <name> (phone)' / street / 'CITY, PR POSTAL CODE ...'
    name = phone = street = city = code = ref = ""
    for i, l in enumerate(lines):
        if not l.startswith("To:"):
            continue
        to_line = l[3:]
        m = PHONE_RE.search(to_line)
        if m:
            phone = m.group(0)
            to_line = to_line[: m.start()] + to_line[m.end() :]
        name = to_line.strip()
        if i + 1 < len(lines):
            street = lines[i + 1].rstrip(",")
        if i + 2 < len(lines):
            m = CITY_CODE_RE.match(lines[i + 2])
            if m:
                city, code = m["city"], m["code"] or ""
                ref = (m["ref"] or "").strip()  # 'WONC05 BY NT PT'
        break

    m = re.search(r"Customer Code:\s*(\S+)", flat)
    if m:
        code = m.group(1)

    items = []
    for l in lines:
        r = ROW_RE.match(l)
        if r:
            amt = r["amount"] if r["amount"].startswith("$") else "$" + r["amount"]
            items.append(
                dict(
                    date=r["date"],
                    transaction=r["transaction"].strip(),
                    invoice=r["invoice"],
                    policy=r["policy"],
                    description=r["description"].strip(),
                    amount=amt,
                )
            )

    m = re.search(r"Outstanding Balance\s*:\s*(-?\$?[\d,]+\.\d{2})", flat)
    balance = m.group(1) if m else ""
    calc = sum(money(i["amount"]) for i in items)
    if not balance:
        balance = f"${calc:,.2f}"
    elif abs(money(balance) - calc) > 0.005:
        print(
            f"  ! Warning: rows sum to ${calc:,.2f} but statement balance is {balance}"
        )

    return dict(
        name=name,
        phone=phone,
        street=street,
        city=city,
        customer_code=code,
        customer_ref=ref or code,
        invoice_date=statement_date,
        total=balance,
        amount_due=balance,
        items=items,
    )


# --------------------------------------------------------------------------- #
# Form field mapping
# --------------------------------------------------------------------------- #
def number(s: str) -> str:
    """'$1,182.00' -> '1182.00'. The template's amount/total fields run
    AFNumber_Format, which shows NaN ('$1.#R') for text containing $ or ,"""
    return f"{money(s):.2f}"


def build_values(data: dict):
    """Return ({field_name: value}, number_of_rows_used)."""
    items = data["items"]
    if len(items) > MAX_ROWS:
        print(
            f"  ! Template only has {MAX_ROWS} rows but statement has "
            f"{len(items)} lines - extra lines were NOT written."
        )
        items = items[:MAX_ROWS]
    n_rows = max(len(items), 1)

    values = {
        "stmt_date": data["invoice_date"],
        "cust_name": data["name"],
        "cust_phone": data["phone"],
        "cust_addr1": data["street"],
        "cust_addr2": data["city"],
        "cust_code": data["customer_ref"],
        # stub fields are normally copied by JavaScript in Acrobat; set them
        # directly so they show in every viewer
        "stub_name": data["name"],
        "stub_addr1": data["street"],
        "stub_addr2": data["city"],
        "stub_code": data["customer_code"],
        "stub_date": data["invoice_date"],
        "stub_due": number(data["amount_due"]),
        f"total_{n_rows}": number(data["total"]),
    }
    for r, item in enumerate(items, 1):
        for prefix, key in ROW_FIELDS.items():
            v = item[key]
            values[f"{prefix}_{r}"] = number(v) if key == "amount" else v
    return values, n_rows


def row_visibility(n_rows: int) -> dict:
    """{field_name: visible?} mirroring what the '+ Add Row' button does."""
    vis = {}
    for r in range(2, MAX_ROWS + 1):
        for name in [f"rowbox_{r}", *(f"{p}_{r}" for p in ROW_FIELDS)]:
            vis[name] = r <= n_rows
    for r in range(1, MAX_ROWS + 1):
        for name in (f"ob_box_{r}", f"ob_label_{r}", f"total_{r}"):
            vis[name] = r == n_rows
    return vis


def use_standard_form_fonts(doc):
    """Point the form's Helv/HeBo at the standard Helvetica fonts.

    The template's /DR fonts are embedded Arial subsets, and with
    NeedAppearances on, viewers redraw fields with them - any letter missing
    from the subset vanishes (e.g. 'WONC05' losing letters)."""
    cat = doc.pdf_catalog()
    refs = []
    for name, base in (("Helv", "Helvetica"), ("HeBo", "Helvetica-Bold")):
        xref = doc.get_new_xref()
        doc.update_object(
            xref,
            f"<</Type/Font/Subtype/Type1/BaseFont/{base}/Encoding/WinAnsiEncoding>>",
        )
        refs.append(f"/{name} {xref} 0 R")
    doc.xref_set_key(cat, "AcroForm/DR/Font", f"<<{''.join(refs)}>>")
    # let the viewer redraw fields so its $ formatting scripts apply
    doc.xref_set_key(cat, "AcroForm/NeedAppearances", "true")


def scoped_js(js: str) -> str:
    """Run a form script in its own function scope.

    Some viewers share one global scope between scripts, so the Remove Row
    loop's 'i' got overwritten by the total calculation (also 'var i') that
    fires when a cell is cleared - only one cell was removed per click."""
    body = re.sub(r"\bthis\b", "doc", js)
    return f"(function(doc){{{body}}})(this);"


def fill(template: Path, out_path: Path, data: dict):
    values, n_rows = build_values(data)
    vis = row_visibility(n_rows)
    with pymupdf.open(template) as doc:
        for page in doc:
            for w in page.widgets():
                name = w.field_name
                changed = False
                if name in ("btn_add", "btn_remove") and w.script:
                    w.script = scoped_js(w.script)
                    changed = True
                if w.script_calc and "for(" in w.script_calc:
                    w.script_calc = scoped_js(w.script_calc)
                    changed = True
                if name in vis:
                    w.field_display = DISPLAY_VISIBLE if vis[name] else DISPLAY_HIDDEN
                    changed = True
                if name in values:
                    w.field_value = values[name]
                    changed = True
                if changed:
                    # update() forces the font to plain Helv (PyMuPDF has no
                    # bold field font) - put the original bold DA back
                    da = doc.xref_get_key(w.xref, "DA")[1]
                    w.update()
                    if "/HeBo" in da:
                        doc.xref_set_key(w.xref, "DA", pymupdf.get_pdf_str(da))
        use_standard_form_fonts(doc)
        doc.save(out_path)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def list_fields():
    with pymupdf.open(TEMPLATE) as doc:
        for pno, page in enumerate(doc, 1):
            widgets = sorted(
                page.widgets(), key=lambda w: (round(w.rect.y0), w.rect.x0)
            )
            for w in widgets:
                print(
                    f"p{pno}  x={w.rect.x0:6.1f} y={w.rect.y0:6.1f}  "
                    f"{w.field_type_string:<9} {w.field_name!r}  value={w.field_value!r}"
                )


def output_dir() -> Path:
    for d in (Path.home() / "Desktop", Path.home() / "OneDrive" / "Desktop"):
        if d.is_dir():
            return d
    return Path.cwd()


def unique_path(p: Path) -> Path:
    if not p.exists():
        return p
    i = 2
    while (q := p.with_name(f"{p.stem} ({i}){p.suffix}")).exists():
        i += 1
    return q


def manual_invoice(config_data=None, debug: bool = False) -> int:
    """Entry point for file_completion_tool. Returns the number of invoices made."""
    if not TEMPLATE.is_file():
        print(f"Template not found: {TEMPLATE}")
        return 0

    pdfs = sorted(
        DOWNLOADS.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True
    )
    pdfs = pdfs[:SCAN_COUNT]
    if not pdfs:
        print(f"No PDFs in {DOWNLOADS}")
        return 0

    out_dir = output_dir()
    made = 0
    for pdf in pdfs:
        if debug:
            print(f"----- {pdf.name} -----")
            print(f"Title region text: {statement_title_text(pdf)!r}")
        if not is_customer_statement(pdf):
            print(f"Skipping {pdf.name} (no 'CUSTOMER STATEMENT' in title area)")
            continue

        text = read_text(pdf)
        if debug:
            print(f"{text}\n")

        print(f"Processing {pdf.name}")
        data = parse_statement(text)
        if not data["items"]:
            print("  ! No transaction lines parsed - try --debug")
            continue

        client = re.sub(r'[\\/:*?"<>|]', "", data["name"]).strip()
        out = unique_path(out_dir / f"{client or 'Unknown Client'} Invoice.pdf")
        fill(TEMPLATE, out, data)
        print(f"  -> {out}")
        made += 1

    if made:
        print(f"\nCreated {made} invoice(s) in {out_dir}")
    else:
        print(
            "No customer statement found among the last "
            f"{SCAN_COUNT} modified PDFs in {DOWNLOADS}"
        )
    return made


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--list-fields", action="store_true")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()

    if args.list_fields:
        return list_fields()
    if not manual_invoice(debug=args.debug):
        sys.exit(1)


if __name__ == "__main__":
    main()
