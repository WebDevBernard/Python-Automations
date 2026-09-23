from pathlib import Path

import openpyxl
from docxtpl import DocxTemplate

from utils import (
    _escape_xml_values,
    _split_mailing_address,
    load_producer_mapping,
    parse_date,
    progressbar,
    safe_filename,
    safe_strip,
    unique_file_name,
    write_to_new_docx,
)

DOWNLOADS_DIR = Path.home() / "Downloads"
OUTPUT_DIR = Path.home() / "Desktop" / "Disclosure Notices"


def load_excluded_ccodes(mapping_path):
    """Reads ccodes to exclude from disclosure letters. They live in the
    'File Completion Tool' worksheet, column D, starting at row 27."""
    wb = openpyxl.load_workbook(mapping_path, data_only=True)
    if "File Completion Tool" not in wb.sheetnames:
        return set()
    ws = wb["File Completion Tool"]
    ccodes = set()
    for row in range(27, ws.max_row + 1):
        value = safe_strip(ws.cell(row=row, column=4).value)
        if value:
            ccodes.add(value.upper())
    return ccodes


def load_cancelled_ccodes(mapping_path):
    """Reads cancelled ccodes from the 'no active' worksheet in config.xlsx."""
    wb = openpyxl.load_workbook(mapping_path, data_only=True)
    if "no active" not in wb.sheetnames:
        return set()
    ws = wb["no active"]
    headers = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        header = safe_strip(cell.value).lower()
        if header:
            headers[header] = col_idx
    if "ccode" not in headers:
        return set()
    ccode_col = headers["ccode"]
    ccodes = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        ccode = safe_strip(row[ccode_col - 1]) if len(row) >= ccode_col else ""
        if ccode:
            ccodes.add(ccode.upper())
    return ccodes


def is_row_highlighted(ws, row_idx, check_col=1):
    """Returns True if a row is marked cancelled (red solid fill)."""
    cell = ws.cell(row=row_idx, column=check_col)
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return False
    fg = fill.fgColor
    if fg is None or fg.rgb is None:
        return False
    return str(fg.rgb).upper().endswith("FF6666")


def read_filtered_transactions_sheet(mapping_path, producer_mapping, cancelled_ccodes=None):
    """Reads rows from the 'Filtered Transactions' sheet in config.xlsx.

    Rows highlighted red by mark_cancelled_ccodes.py are marked cancelled.
    """
    cancelled_ccodes = cancelled_ccodes or set()
    wb = openpyxl.load_workbook(mapping_path, data_only=True)
    if "Filtered Transactions" not in wb.sheetnames:
        print("\u274c 'Filtered Transactions' sheet not found in config.xlsx")
        return []

    ws = wb["Filtered Transactions"]

    headers = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        header = safe_strip(cell.value).lower()
        if header:
            headers[header] = col_idx

    required = [
        "policy_number",
        "named_insured",
        "mailing_address",
        "effective_date",
        "insurer",
    ]
    missing = [c for c in required if c not in headers]
    if missing:
        print(f"\u274c Missing required columns in Filtered Transactions sheet: {missing}")
        return []

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        policy_number = safe_strip(row[headers["policy_number"] - 1].value)
        named_insured = safe_strip(row[headers["named_insured"] - 1].value)
        if not policy_number and not named_insured:
            continue

        pcode_raw = ""
        if "producer_name" in headers:
            pcode_raw = safe_strip(row[headers["producer_name"] - 1].value)
        elif "pcode" in headers:
            pcode_raw = safe_strip(row[headers["pcode"] - 1].value)

        ccode_val = ""
        cancelled = is_row_highlighted(ws, row[0].row)
        if "ccode" in headers:
            ccode_cell = row[headers["ccode"] - 1]
            ccode_val = safe_strip(ccode_cell.value)
            cancelled = cancelled or ccode_val.upper() in cancelled_ccodes

        entry = {
            "ccode": ccode_val,
            "policy_number": policy_number,
            "named_insured": named_insured,
            "mailing_address": _split_mailing_address(
                row[headers["mailing_address"] - 1].value
            ),
            "effective_date": parse_date(row[headers["effective_date"] - 1].value),
            "insurer": safe_strip(row[headers["insurer"] - 1].value),
            "producer_name": producer_mapping.get(pcode_raw.lower(), ""),
            "cancelled": cancelled,
        }

        rows.append(entry)

    return rows


def _find_latest_filtered_file():
    """Picks the most recently created '*_filtered*.xlsx' in Downloads
    (i.e. the output of filter_earliest_transaction.py)."""
    candidates = sorted(
        DOWNLOADS_DIR.glob("*_filtered*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def read_sheet_rows(source_path, producer_mapping):
    wb = openpyxl.load_workbook(source_path, data_only=True)
    ws = wb.active

    headers = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        header = safe_strip(cell.value).lower()
        if header:
            headers[header] = col_idx

    required = [
        "policy_number",
        "named_insured",
        "mailing_address",
        "effective_date",
        "insurer",
        "producer_name",
    ]
    missing = [c for c in required if c not in headers]
    if missing:
        print(f"\u274c Missing required columns in source sheet: {missing}")
        return []

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        policy_number = safe_strip(row[headers["policy_number"] - 1].value)
        named_insured = safe_strip(row[headers["named_insured"] - 1].value)
        if not policy_number and not named_insured:
            continue

        pcode_raw = safe_strip(row[headers["producer_name"] - 1].value)

        entry = {
            "policy_number": policy_number,
            "named_insured": named_insured,
            "mailing_address": _split_mailing_address(
                row[headers["mailing_address"] - 1].value
            ),
            "effective_date": parse_date(row[headers["effective_date"] - 1].value),
            "insurer": safe_strip(row[headers["insurer"] - 1].value),
            "producer_name": producer_mapping.get(pcode_raw.lower(), ""),
        }

        rows.append(entry)

    return rows


def disclosure_notice(config_data=None):
    """Reads from the 'Filtered Transactions' sheet in config.xlsx."""
    mapping_path = "config.xlsx"
    if not Path(mapping_path).exists():
        print("Config file not found: config.xlsx")
        return

    producer_mapping = load_producer_mapping(mapping_path)
    cancelled_ccodes = load_cancelled_ccodes(mapping_path)
    excluded_ccodes = load_excluded_ccodes(mapping_path)
    rows = read_filtered_transactions_sheet(
        mapping_path, producer_mapping, cancelled_ccodes
    )
    if not rows:
        print("No data found in Filtered Transactions sheet")
        return

    if excluded_ccodes:
        excluded = [r for r in rows if r.get("ccode", "").upper() in excluded_ccodes]
        rows = [r for r in rows if r.get("ccode", "").upper() not in excluded_ccodes]
        if excluded:
            print(
                f"\u26a0\ufe0f Skipping {len(excluded)} excluded disclosure letter(s): "
                f"{sorted({r['ccode'] for r in excluded})}"
            )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    template = Path.cwd() / "assets" / "Disclosure Notice.docx"
    if not template.exists():
        print(f"Template not found: {template}")
        return

    success_count = 0
    cancelled_count = 0
    total_expected = len(rows)

    for row_data in progressbar(rows, prefix="Generating: ", size=40):
        try:
            ccode = row_data.get("ccode", "")
            name = row_data.get("named_insured", "Unnamed Client")

            doc = DocxTemplate(template)
            doc.render(_escape_xml_values(row_data))

            if row_data.get("cancelled"):
                output_dir = OUTPUT_DIR / "Cancelled Policies"
                output_dir.mkdir(parents=True, exist_ok=True)
                cancelled_count += 1
            else:
                output_dir = OUTPUT_DIR

            safe_name = safe_filename(f"{ccode} - {name}")
            output_path = unique_file_name(output_dir / f"{safe_name}.docx")
            doc.save(output_path)
            success_count += 1
        except Exception as e:
            import traceback

            print(
                f"\u274c Failed for {row_data.get('named_insured', 'unknown')} "
                f"({row_data.get('policy_number', 'unknown')}): {e}"
            )
            traceback.print_exc()

    print(
        f"******** Disclosure Notice completed: {success_count}/{total_expected} letters generated ********"
    )
    print(f"Output folder: {OUTPUT_DIR}")
    if cancelled_count:
        print(
            f"Cancelled notices ({cancelled_count}) saved to: {OUTPUT_DIR / 'Cancelled Policies'}"
        )


def disclosure_notice_from_filtered(source_path=None):
    mapping_path = "config.xlsx"
    producer_mapping = load_producer_mapping(mapping_path)

    source_path = Path(source_path) if source_path else _find_latest_filtered_file()

    if not source_path or not source_path.exists():
        print(
            "Could not find a *_filtered.xlsx file in Downloads. "
            "Run Filter Earliest Transaction first, or pass the file path directly."
        )
        return

    print(f"Reading: {source_path.name}")
    rows = read_sheet_rows(source_path, producer_mapping)
    if not rows:
        print("No data found in source sheet")
        return

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    success_count = 0
    total_expected = len(rows)

    for row_data in rows:
        try:
            template = Path.cwd() / "assets" / "Disclosure Notice.docx"
            if write_to_new_docx(
                template_path=template,
                data=row_data,
                output_dir=OUTPUT_DIR,
                output_suffix=f"Disclosure Notice - {row_data['policy_number']}",
            ):
                success_count += 1
        except Exception as e:
            import traceback

            print(
                f"\u274c Failed for {row_data.get('named_insured', 'unknown')} "
                f"({row_data.get('policy_number', 'unknown')}): {e}"
            )
            traceback.print_exc()

    print(
        f"******** Disclosure Notice completed: {success_count}/{total_expected} letters generated ********"
    )
    print(f"Output folder: {OUTPUT_DIR}")


if __name__ == "__main__":
    disclosure_notice()
