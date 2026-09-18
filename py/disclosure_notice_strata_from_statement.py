import math
import re
import time
import traceback
from datetime import datetime
from pathlib import Path

import fitz

from utils import write_to_new_docx
from statement_to_excel import extract_table_from_pdf
from constants import REGEX_PATTERNS
from py.disclosure_notice_from_excel_ws import _split_mailing_address
from py.disclosure_notice_strata_from_excel_ws import (
    format_amount,
    get_insurer,
    load_producer_mapping,
    parse_date,
    parse_date_to_dt,
    smart_title,
    to_float,
)

DATE_FORMAT = "%B %d, %Y"

TABLE_KEYWORDS = ["Date", "Transaction", "Invoice", "Policy", "Description", "Amount"]

# Address lines start with a street number, "Unit", or "PO Box" (same rule as
# the auto renewal letter). Postal codes mark the end of the address.
_ADDRESS_RE = REGEX_PATTERNS["address"]
_POSTAL_RE = REGEX_PATTERNS["postal_code"]

# Producer line, e.g. "EPS530 BY NT" or "BCS358 BY NT NT".
_PRODUCER_RE = re.compile(r"^\S+\s+BY\s+\S+", re.IGNORECASE)


# Full-line labels that end a name/address block.
_STOP_LABELS = (
    "To",
    "From",
    "Re",
    "Date",
    "Phone",
    "Fax",
    "Policy",
    "Account",
    "Invoice",
    "Amount",
    "Total",
    "Page",
    "Attention",
    "Attn",
    "Enclosed",
)
_STOP_PATTERN = re.compile(rf"^\s*(?:{'|'.join(_STOP_LABELS)})\s*:", re.IGNORECASE)


def _is_label_line(line, label):
    """True if the line is the label itself (e.g. 'To', 'To:') or starts the block."""
    return bool(
        re.match(rf"^\s*{label}\s*:?\s*$", line, re.IGNORECASE)
        or re.match(rf"^\s*{label}\s*:\s*(.+)$", line, re.IGNORECASE)
    )


def extract_label_block(page, label):
    """Extract all text lines following '<label>:' on a page.

    Collects lines until the next label (e.g. 'From:', 'Date:') so names can
    span multiple lines without being cut off. The caller splits the result
    into name / address / producer using regex rules.
    """
    blocks = sorted(page.get_text("blocks"), key=lambda b: (b[1], b[0]))

    for i, block in enumerate(blocks):
        lines = [ln.strip() for ln in block[4].split("\n")]
        label_index = -1
        for idx, ln in enumerate(lines):
            if _is_label_line(ln, label):
                label_index = idx
                break
        if label_index == -1:
            continue

        content = []
        for ln in lines[label_index:]:
            if re.match(rf"^\s*{label}\s*:?\s*$", ln, re.IGNORECASE):
                continue
            m = re.match(rf"^\s*{label}\s*:\s*(.*)$", ln, re.IGNORECASE)
            if m is not None:
                rest = m.group(1).strip()
                if rest:
                    content.append(rest)
                continue
            if _STOP_PATTERN.match(ln):
                break
            if ln:
                content.append(ln)

        # Content continues in the following blocks until the next label/gap.
        for nxt in blocks[i + 1 :]:
            nxt_lines = [l.strip() for l in nxt[4].split("\n") if l.strip()]
            if not nxt_lines:
                continue
            if _STOP_PATTERN.match(nxt_lines[0]):
                break
            for ln in nxt_lines:
                if _STOP_PATTERN.match(ln):
                    break
                content.append(ln)

        return content if content else None

    return None


def extract_name_address(doc):
    """Find the 'To:' block (primary) and 'From:' block (verification)."""
    to_lines = None
    from_lines = None

    for page in doc:
        to_lines = extract_label_block(page, "To")
        if to_lines:
            break

    for page in doc:
        from_lines = extract_label_block(page, "From")
        if from_lines:
            break

    return to_lines, from_lines


def _normalize(value):
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def extract_producer_code(producer_line):
    """Producer code is the 3rd whitespace-separated token in the producer line.

    Handles short lines like "EPS530 BY NT" (3 tokens) and even "BY NT" (2).
    """
    if not producer_line:
        return None
    parts = producer_line.split()
    if len(parts) >= 3:
        return parts[2]
    if len(parts) == 2:
        return parts[1]
    return None


def _split_block_lines(lines):
    """Split raw 'To:'/'From:' lines into (name, address, producer_line).

    - Address starts at the first line that looks like a street address
      (begins with a number, 'Unit', or 'PO Box' - same rule as the auto
      renewal letter), so extra name lines before it are kept as the name.
    - The producer line (e.g. 'EPS530 BY NT') ends the address block.
    """
    producer_line = None
    producer_index = -1
    for i in range(len(lines) - 1, -1, -1):
        if _PRODUCER_RE.match(lines[i]):
            producer_index = i
            producer_line = lines[i]
            break

    body = lines[:producer_index] if producer_index != -1 else lines

    address_start = -1
    for i, ln in enumerate(body):
        if _ADDRESS_RE.search(ln):
            address_start = i
            break

    if address_start == -1:
        return body[:1], body[1:], producer_line

    # Address ends at the line containing the postal code (inclusive), matching
    # the auto renewal letter; otherwise it runs to the end of the block.
    address_end = len(body)
    for i in range(address_start, len(body)):
        if _POSTAL_RE.search(body[i]):
            address_end = i + 1
            break

    return body[:address_start], body[address_start:address_end], producer_line


def _build_block(lines):
    name_lines, address_lines, producer_line = _split_block_lines(lines)
    return {
        "name": " ".join(name_lines) if name_lines else "",
        "address": [l for l in address_lines if l],
        "producer_line": producer_line,
    }


def parse_blocks(to_lines, from_lines):
    """Build the name/address dict from the 'To:' block, verified by 'From:'."""
    if not to_lines:
        return None, "Could not find a 'To:' name/address block in the statement."

    to_block = _build_block(to_lines)

    if not from_lines:
        return (
            to_block,
            "Warning: 'From:' block not found - using 'To:' block without verification.",
        )

    from_block = _build_block(from_lines)

    if _normalize(to_block["name"]) == _normalize(from_block["name"]) and (
        _normalize(" ".join(to_block["address"]))
        == _normalize(" ".join(from_block["address"]))
    ):
        message = f"Verified: 'To:' block matches 'From:' block ({to_block['name']})."
        return to_block, message

    message = (
        "Warning: 'From:' block does not match 'To:' block.\n"
        f"    To:   {to_block['name']} | {' | '.join(to_block['address'])}\n"
        f"    From: {from_block['name']} | {' | '.join(from_block['address'])}\n"
        "    Using 'To:' block for the letter."
    )
    return to_block, message


def _find_column(df, keyword):
    for col in df.columns:
        if keyword in str(col).lower():
            return col
    return None


def _earliest_date(frame, date_col):
    if not date_col:
        return None
    dates = []
    for value in frame[date_col]:
        dt = parse_date_to_dt(value)
        if dt:
            dates.append(dt)
    return min(dates) if dates else None


def _policy_row(policy_number, prem_total, effective_dt):
    return {
        "policy_number": policy_number,
        "prem_amt": format_amount(prem_total) if prem_total else "",
        "effective_date": (
            parse_date(effective_dt)
            if effective_dt
            else datetime.today().strftime(DATE_FORMAT)
        ),
        "insurer": get_insurer(policy_number),
    }


def build_policy_rows(df):
    """Build one letter-row per policy from the extracted transaction table.

    Returns a list of dicts, a single dict with statement totals when no policy
    column exists, or None when there is no usable table.
    """
    if df is None or df.empty:
        return None

    policy_col = _find_column(df, "policy")
    date_col = _find_column(df, "date")
    amount_col = _find_column(df, "amount")

    if policy_col:
        rows = []
        for policy, group in df.groupby(policy_col, dropna=False):
            if isinstance(policy, float) and math.isnan(policy):
                continue
            policy_number = str(policy).strip()
            if not policy_number:
                continue
            prem_total = group[amount_col].apply(to_float).sum() if amount_col else 0.0
            effective_dt = _earliest_date(group, date_col)
            rows.append(_policy_row(policy_number, prem_total, effective_dt))
        return rows or None

    prem_total = df[amount_col].apply(to_float).sum() if amount_col else 0.0
    effective_dt = _earliest_date(df, date_col)
    return [_policy_row("", prem_total, effective_dt)]


def generate_letters(pdf_path, template_path, output_dir, producer_mapping=None):
    """Extract name/address + transactions from a statement and write letters."""
    producer_mapping = producer_mapping or {}
    print(f"Using statement: {pdf_path.name}")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    to_lines, from_lines = extract_name_address(doc)
    doc.close()

    block, message = parse_blocks(to_lines, from_lines)
    print(message)
    if not block:
        return 0

    named_insured = smart_title(block["name"])
    mailing_address = _split_mailing_address(", ".join(block["address"]))

    producer_code = extract_producer_code(block.get("producer_line"))
    producer_name = (
        producer_mapping.get(producer_code.lower(), "") if producer_code else ""
    )

    print(f"Named insured: {named_insured}")
    print(f"Policy address: {mailing_address}")
    print(
        f"Producer: {producer_name or 'unknown'} " f"({producer_code})"
        if producer_code
        else f"Producer: {producer_name or 'not found'}"
    )

    df = extract_table_from_pdf(pdf_path, TABLE_KEYWORDS)
    rows = build_policy_rows(df)

    if not rows:
        rows = [
            {
                "policy_number": "",
                "prem_amt": "",
                "effective_date": datetime.today().strftime(DATE_FORMAT),
                "insurer": "",
            }
        ]
        print(
            "Warning: no transaction rows extracted; generating a single letter "
            "from name/address only."
        )

    success = 0
    for row in rows:
        row["named_insured"] = named_insured
        row["mailing_address"] = mailing_address
        row["producer_name"] = producer_name
        suffix = (
            f"Strata Disclosure Notice - {row['policy_number']}"
            if row["policy_number"]
            else "Strata Disclosure Notice"
        )
        try:
            if write_to_new_docx(
                template_path=template_path,
                data=row,
                output_dir=output_dir,
                output_suffix=suffix,
            ):
                success += 1
        except Exception as e:
            print(f"Failed for {row.get('policy_number', 'unknown')}: {e}")
            traceback.print_exc()

    return success


def _exit_countdown(seconds=3):
    print("\nExiting in ", end="")
    for i in range(seconds, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)
    print()


def _find_statements(downloads_dir, count=2):
    pdfs = sorted(
        downloads_dir.glob("*.pdf"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return pdfs[:count]


def strata_disclosure_from_statement(config_data=None):
    downloads_dir = Path.home() / "Downloads"
    pdf_paths = _find_statements(downloads_dir)
    if not pdf_paths:
        print("No PDF files found in Downloads folder.")
        _exit_countdown()
        return

    template_path = Path.cwd() / "assets" / "Strata Disclosure Notice.docx"
    if not template_path.exists():
        print(f"Template not found: {template_path}")
        _exit_countdown()
        return

    output_dir = Path.home() / "Desktop"
    producer_mapping = load_producer_mapping("config.xlsx")

    success = 0
    for pdf_path in pdf_paths:
        success += generate_letters(
            pdf_path, template_path, output_dir, producer_mapping
        )

    print(
        f"******** Strata Disclosure Letter from Statement ran successfully: "
        f"{success} letter(s) generated ********"
    )
    _exit_countdown()


if __name__ == "__main__":
    strata_disclosure_from_statement()
