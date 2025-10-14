import os
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.page import PageMargins
from UTILITIES import unique_file_name

input_dir = Path.home() / "Desktop"
output_dir = Path.home() / "Desktop"

def sort_renewal_list():
    """Process the most recent Excel file, clean, and format it as a renewal list."""
    
    # --- Step 0: Find Excel files ---
    xlsx_files = list(Path(input_dir).glob("*.xlsx"))
    xls_files = list(Path(input_dir).glob("*.xls"))
    files = xlsx_files + xls_files

    if not files:
        print(f"No Excel files found in {input_dir}. Please place your renewal lists on the Desktop.")
        return

    # --- Step 1: Pick the most recently modified Excel file ---
    latest_file = sorted(files, key=lambda f: f.stat().st_mtime, reverse=True)[0]
    print(f"Processing latest file: {latest_file.name}")

    engine = "xlrd" if latest_file.suffix.lower() == ".xls" else "openpyxl"
    df = pd.read_excel(latest_file, engine=engine)

    # --- Step 2: Clean and transform data ---
    column_list = [
        "policynum", "ccode", "name", "pcode", "csrcode", "insurer",
        "buscode", "renewal", "Pulled", "D/L"
    ]

    # Reorder columns if they exist
    df = df.reindex(columns=column_list)
    df = df.drop_duplicates(subset=["policynum"], keep=False)

    # --- Step 2a: Safe date parsing ---
    # Convert to datetime, coerce errors, fill missing with today's date
    df["renewal_parsed"] = pd.to_datetime(df["renewal"], dayfirst=True, errors="coerce")
    df["renewal_parsed"].fillna(pd.Timestamp.today(), inplace=True)

    # Format dates
    df["renewal_1"] = df["renewal_parsed"].dt.strftime("%d-%b")
    df["renewal"] = df["renewal_parsed"].dt.strftime("%m%d")

    # Replace renewal column and remove temporary one
    df["renewal"] = df["renewal_1"]
    df.drop(["renewal_1", "renewal_parsed"], axis=1, inplace=True)

    # Sort by insurer, renewal, and name
    df.sort_values(["insurer", "renewal", "name"], inplace=True)

    # Add blank rows between insurers
    dfs_with_spacing = []
    for _, group in df.groupby("insurer", sort=False):
        dfs_with_spacing.append(group)
        dfs_with_spacing.append(pd.DataFrame([[float("NaN")] * len(group.columns)], columns=group.columns))
    df = pd.concat(dfs_with_spacing, ignore_index=True).iloc[:-1]

    # --- Step 3: Write to Excel ---
    output_path = unique_file_name(output_dir / "renewal_list.xlsx")
    mode = "w" if not os.path.isfile(output_path) else "a"
    writer_kwargs = {"engine": "openpyxl", "mode": mode, "if_sheet_exists": "replace"} if mode == "a" else {"engine": "openpyxl", "mode": mode}

    with pd.ExcelWriter(output_path, **writer_kwargs) as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)

    # --- Step 4: Format Excel workbook ---
    wb = load_workbook(output_path)
    ws = wb.active

    # Font and alignment
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(size=12)
            cell.alignment = Alignment(horizontal="left")

    # Create Excel table
    ref = f"A1:{chr(64 + df.shape[1])}{df.shape[0] + 1}"
    table = Table(displayName="Table1", ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1",
        showRowStripes=True,
        showColumnStripes=False
    )
    ws.add_table(table)

    # --- Step 5: Adjust column widths ---
    for i, col in enumerate(column_list, 1):
        if col not in df.columns:
            continue
        max_len = max(df[col].astype(str).map(len).max(), len(col))
        if col in ["pcode", "csrcode", "Pulled", "D/L"]:
            width = 5.0
        elif col == "ccode":
            width = max_len + 4
        elif col == "policynum":
            width = max_len + 2.5
        else:
            width = max_len + 1
        ws.column_dimensions[chr(64 + i)].width = width

    # --- Step 6: Add borders for specific columns ---
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))

    for col in ["Pulled", "D/L"]:
        if col in df.columns:
            col_idx = column_list.index(col) + 1
            for row_idx in range(1, df.shape[0] + 2):
                ws.cell(row=row_idx, column=col_idx).border = border

    # --- Step 7: Page setup ---
    ws.print_title_rows = "1:1"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = False
    ws.page_setup.fitToPage = True
    ws.page_margins = PageMargins(
        top=1.91 / 2.54,
        bottom=1.91 / 2.54,
        left=1.78 / 2.54,
        right=0.64 / 2.54
    )

    wb.save(output_path)
    print("******** Sort Renewal List ran successfully ********")
    print(f"Output file: {output_path}")

# Call the function
if __name__ == "__main__":
    sort_renewal_list()
