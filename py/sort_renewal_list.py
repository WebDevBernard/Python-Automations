import os
import re
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.page import PageMargins
from utils import safe_filename, unique_file_name

input_dir = Path.home() / "Downloads"
output_dir = Path.home() / "Desktop"


def get_sort_key(date_value):
    """Get sort key from date value without parsing."""
    if date_value is None or date_value == "":
        return "9999"

    if isinstance(date_value, datetime):
        return date_value.strftime("%m%d")

    return str(date_value)


def sort_renewal_list():
    """Process the most recent Excel file, clean, and format it as a renewal list."""

    # Find Excel files
    xlsx_files = list(Path(input_dir).glob("*.xlsx"))
    xls_files = list(Path(input_dir).glob("*.xls"))
    files = xlsx_files + xls_files

    if not files:
        print(
            f"No Excel files found in {input_dir}. Please place your renewal lists in Downloads."
        )
        return

    # Pick the two most recently modified Excel files
    sorted_files = sorted(files, key=lambda f: f.stat().st_mtime, reverse=True)
    recent_files = sorted_files[:2]

    if len(recent_files) < 2:
        print(f"Only found one Excel file: {recent_files[0].name}")
    else:
        print(f"Processing the 2 most recent files: {[f.name for f in recent_files]}")

    # Combine data from both files
    data = []
    headers = None

    for file in recent_files:
        wb = load_workbook(file, data_only=True)
        ws = wb.active

        file_headers = [cell.value for cell in ws[1]]
        if headers is None:
            headers = file_headers
        elif file_headers != headers:
            print(f"⚠️ Warning: headers differ in {file.name}")

        for row in ws.iter_rows(min_row=2, values_only=True):
            row_dict = {file_headers[i]: row[i] for i in range(len(file_headers))}
            data.append(row_dict)

    # Define desired columns
    column_list = [
        "policynum",
        "ccode",
        "name",
        "pcode",
        "csrcode",
        "insurer",
        "buscode",
        "renewal",
        "Pulled",
        "D/L",
    ]

    # Remove duplicates based on policynum
    policynum_counts = {}
    for row in data:
        pnum = row.get("policynum")
        policynum_counts[pnum] = policynum_counts.get(pnum, 0) + 1

    data = [row for row in data if policynum_counts.get(row.get("policynum"), 0) == 1]

    # Get sort keys for renewal dates
    for row in data:
        renewal_value = row.get("renewal")
        row["_sort_date"] = get_sort_key(renewal_value)

    # Sort by insurer, renewal date, and name
    data.sort(
        key=lambda x: (
            str(x.get("insurer", "")).lower(),
            x.get("_sort_date", "9999"),
            str(x.get("name", "")).lower(),
        )
    )

    # Add blank rows between insurers
    data_with_spacing = []
    current_insurer = None

    for row in data:
        if current_insurer and row.get("insurer") != current_insurer:
            data_with_spacing.append({col: None for col in column_list})

        data_with_spacing.append(row)
        current_insurer = row.get("insurer")

    # Write to new Excel file
    output_path = unique_file_name(output_dir / "renewal_list.xlsx")
    new_wb = Workbook()
    new_ws = new_wb.active
    new_ws.title = "Sheet1"

    # Write headers
    for col_idx, col_name in enumerate(column_list, 1):
        cell = new_ws.cell(row=1, column=col_idx)
        cell.value = col_name
        cell.font = Font(size=12)
        cell.alignment = Alignment(horizontal="left")

    # Write data rows
    for row_idx, row_data in enumerate(data_with_spacing, 2):
        for col_idx, col_name in enumerate(column_list, 1):
            cell = new_ws.cell(row=row_idx, column=col_idx)
            cell.value = row_data.get(col_name)
            cell.font = Font(size=12)
            cell.alignment = Alignment(horizontal="left")

    # Create Excel table
    total_rows = len(data_with_spacing) + 1
    ref = f"A1:{chr(64 + len(column_list))}{total_rows}"
    table = Table(displayName="Table1", ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1", showRowStripes=True, showColumnStripes=False
    )
    new_ws.add_table(table)

    # Adjust column widths
    for col_idx, col_name in enumerate(column_list, 1):
        max_len = len(col_name)
        for row_data in data_with_spacing:
            value = row_data.get(col_name)
            if value:
                max_len = max(max_len, len(str(value)))

        if col_name in ["pcode", "csrcode", "Pulled", "D/L"]:
            width = 5.0
        elif col_name == "ccode":
            width = max_len + 4
        elif col_name == "policynum":
            width = max_len + 2.5
        else:
            width = max_len + 1

        new_ws.column_dimensions[chr(64 + col_idx)].width = width

    # Add borders for specific columns
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col_name in ["Pulled", "D/L"]:
        col_idx = column_list.index(col_name) + 1
        for row_idx in range(1, total_rows + 1):
            new_ws.cell(row=row_idx, column=col_idx).border = border

    # Page setup
    new_ws.print_title_rows = "1:1"
    new_ws.page_setup.fitToWidth = 1
    new_ws.page_setup.fitToHeight = False
    new_ws.page_setup.fitToPage = True
    new_ws.page_margins = PageMargins(
        top=1.91 / 2.54, bottom=1.91 / 2.54, left=1.78 / 2.54, right=0.64 / 2.54
    )

    new_wb.save(output_path)
    print("******** Sort Renewal List ran successfully ********")
    print(f"Output file: {output_path}")
