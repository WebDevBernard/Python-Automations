import pandas as pd
import fitz
import re
from pathlib import Path
from datetime import datetime

# -------------------- Excel Configuration -------------------- #
def get_excel_data():
    defaults = {
        "number_of_pdfs": 5,
        "agency_name": "",
        "agency_number": "",
        "toggle_timestamp": "Timestamp",
        "toggle_customer_copy": "No",
        "output_dir": str(Path.home() / "Desktop" / "ICBC E-Stamp Copies"),
        "input_dir": str(Path.home() / "Downloads")
    }
    excel_path = Path(__file__).parent.parent / "BM3KXR.xlsx"
    if not excel_path.exists():
        return defaults
    try:
        df = pd.read_excel(excel_path, sheet_name=0, header=None)
        rows = [2, 4, 6, 8, 10, 12, 14]
        keys = list(defaults.keys())
        data = {k: (df.at[r, 1] if r in df.index and not pd.isna(df.at[r, 1]) else defaults[k])
                for k, r in zip(keys, rows)}
        if isinstance(data["number_of_pdfs"], str):
            data["number_of_pdfs"] = defaults["number_of_pdfs"]
        return data
    except:
        return defaults

# -------------------- Patterns and Rectangles -------------------- #
timestamp_rect = (409.979, 63.8488, 576.0, 83.7455)
payment_plan_rect = (425.402, 35.9664, 557.916, 48.3001)
customer_copy_rect = (498.438, 751.953, 578.181, 769.977)

timestamp_pattern = re.compile(r"Transaction Timestamp\s*(\d+)")
payment_plan_pattern = re.compile(r"Payment Plan Agreement", re.IGNORECASE)
license_plate_pattern = re.compile(r"Licence Plate Number\s*([A-Z0-9\- ]+)", re.IGNORECASE)
insured_pattern = re.compile(
    r"(?:Owner|Applicant|Name of Insured\s*\(surname followed by given name\(s\)\))\s*[:\-]?\s*([A-Z][A-Z\s,.'\-]+)",
    re.IGNORECASE
)
customer_copy_pattern = re.compile(r"customer copy", re.IGNORECASE)
not_valid_pattern = re.compile(r"NOT VALID UNLESS STAMPED BY", re.IGNORECASE)
time_of_validation_pattern = re.compile(r"TIME OF VALIDATION", re.IGNORECASE)

# -------------------- Utility Functions -------------------- #
def reverse_name(name: str) -> str:
    parts = [p for p in name.replace(",", " ").split() if p]
    return " ".join(parts[1:] + [parts[0]]).title() if len(parts) > 1 else name.title()

def find_existing_timestamps(root_folder: Path):
    timestamps = set()
    if not root_folder.exists():
        return timestamps
    for pdf_file in root_folder.rglob("*.pdf"):
        try:
            with fitz.open(pdf_file) as doc:
                if doc.page_count > 0:
                    ts_match = timestamp_pattern.search(doc[0].get_text(clip=timestamp_rect))
                    if ts_match:
                        timestamps.add(ts_match.group(1))
        except:
            continue
    return timestamps

def format_transaction_timestamp(timestamp_str):
    """
    Converts a full transaction timestamp string in the format YYYYMMDDHHMMSS
    into a datetime object.
    """
    year = int(timestamp_str[0:4])
    month = int(timestamp_str[4:6])
    day = int(timestamp_str[6:8])
    hour = int(timestamp_str[8:10])
    minute = int(timestamp_str[10:12])
    second = int(timestamp_str[12:14])
    datetime_obj = datetime(year, month, day, hour, minute, second)
    return datetime_obj

def format_timestamp_mmmddyyyy_from_dt(dt):
    """
    Convert a datetime object into MMMDDYYYY format, e.g., Oct132025
    """
    return dt.strftime("%b%d%Y")

def write_text_to_pdf(pdf_path, agency_number, timestamp_dt, rect, output_dir):
    """
    Write agency number and timestamp onto the first page at the given rectangle
    and save directly to the specified output directory.
    """
    text_to_write = f"Agency: {agency_number}   Timestamp: {format_timestamp_mmmddyyyy_from_dt(timestamp_dt)}"
 
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]  # first page
        x0, y0, x1, y1 = rect
        font_size = 10
        page.insert_text(
            (x0, y0),
            text_to_write,
            fontname="helv",
            fontsize=font_size,
            color=(0, 0, 0),
        )
        # Save directly to the output folder without "_stamped"
        output_file = Path(output_dir) / Path(pdf_path).name
        doc.save(output_file)
        doc.close()
        print(f"Saved stamped PDF: {output_file}")
    except Exception as e:
        print(f"Failed to write to PDF {pdf_path}: {e}")

# -------------------- ICBC PDF Processing -------------------- #
def scan_icbc_pdfs(input_dir, output_dir, max_docs=5):
    input_dir, output_dir = Path(input_dir), Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    existing_timestamps = find_existing_timestamps(output_dir)
    icbc_data = {}

    # Patterns for first page checks
    temporary_permit_pattern = re.compile(
        r"Temporary Operation Permit and Owner’s Certificate of Insurance", re.IGNORECASE
    )
    agency_number_pattern = re.compile(
        r"Agency Number\s*[:#]?\s*([A-Z0-9]+)", re.IGNORECASE
    )

    pdf_files = sorted(input_dir.glob("*.pdf"), key=lambda f: f.stat().st_mtime, reverse=True)[:max_docs]

    for pdf_path in pdf_files:
        try:
            with fitz.open(pdf_path) as doc:
                if doc.page_count == 0:
                    continue

                # First page
                first_page = doc[0]
                first_text = first_page.get_text("text")
                ts_text = first_page.get_text(clip=timestamp_rect)
                payment_text = first_page.get_text(clip=payment_plan_rect)

                # Transaction timestamp
                ts_match = timestamp_pattern.search(ts_text)
                if not ts_match or ts_match.group(1) in existing_timestamps:
                    continue
                timestamp = ts_match.group(1)

                # Skip if payment plan exists
                if payment_plan_pattern.search(payment_text):
                    continue

                # License plate
                license_plate_match = license_plate_pattern.search(first_text)
                license_plate = license_plate_match.group(1).strip().upper() if license_plate_match else None

                # Insured name
                insured_match = insured_pattern.search(first_text)
                insured_name = reverse_name(insured_match.group(1).strip()) if insured_match else None

                # Temporary Operation Permit check (exact line)
                temp_permit_found = bool(temporary_permit_pattern.search(first_text))

                # Agency number
                agency_number_match = agency_number_pattern.search(first_text)
                agency_number = agency_number_match.group(1).strip() if agency_number_match else None

                # All-page patterns
                customer_copy_pages, not_valid_coords, time_of_validation_coords = [], [], []
                for page_num, p in enumerate(doc):
                    clipped_customer_copy = p.get_text(clip=customer_copy_rect)
                    if customer_copy_pattern.search(clipped_customer_copy):
                        customer_copy_pages.append(page_num)
                    for w in p.get_text("blocks"):
                        word_text, coords = w[4], w[:4]
                        if not_valid_pattern.search(word_text):
                            not_valid_coords.append((page_num, coords))
                        if time_of_validation_pattern.search(word_text):
                            time_of_validation_coords.append((page_num, coords))

                # Save data for this PDF
                icbc_data[pdf_path] = {
                    "transaction_timestamp": timestamp,
                    "license_plate": license_plate,
                    "insured_name": insured_name,
                    "temporary_operation_permit": temp_permit_found,
                    "agency_number": agency_number,
                    "customer_copy_pages": customer_copy_pages,
                    "not_valid_coords": not_valid_coords,
                    "time_of_validation_coords": time_of_validation_coords
                }

        except Exception as e:
            print(f"Error processing {pdf_path}: {e}")
            continue

    return icbc_data, len(pdf_files)

# -------------------- Main Function -------------------- #
def main():
    config = get_excel_data()
    data, count = scan_icbc_pdfs(
        config["input_dir"],
        config["output_dir"],
        config["number_of_pdfs"]
    )

    print(f"Scanned {count} PDFs. Found {len(data)} new ICBC documents.\n")

    for path, info in data.items():
        print(f"{path}:")
        for key, value in info.items():
            print(f"  {key}: {value}")
        print("-" * 60)

        # Write agency number and timestamp onto PDF
        if info["agency_number"] and info["transaction_timestamp"]:
            ts_dt = format_transaction_timestamp(info["transaction_timestamp"])
            write_text_to_pdf(path, info["agency_number"], ts_dt, timestamp_rect, config["output_dir"])

if __name__ == "__main__":
    main()
