import fitz
import re
from pathlib import Path
from datetime import datetime
import timeit
import time


# -------------------- Output Directory -------------------- #
def get_output_dir(subfolder_name="ICBC E-Stamp Copies"):
    desktop_path = Path.home() / "Desktop"
    if desktop_path.exists():
        return desktop_path / subfolder_name
    else:
        return Path(__file__).parent.parent / subfolder_name


# -------------------- Defaults -------------------- #
DEFAULTS = {
    "number_of_pdfs": 10,
    "output_dir": str(get_output_dir()),
    "input_dir": str(Path.home() / "Downloads"),
}

# -------------------- Patterns and Rectangles -------------------- #
timestamp_rect = (409.979, 63.8488, 576.0, 83.7455)
payment_plan_rect = (425.402, 35.9664, 557.916, 48.3001)
customer_copy_rect = (498.438, 751.953, 578.181, 769.977)

timestamp_pattern = re.compile(r"Transaction Timestamp\s*(\d+)")
payment_plan_pattern = re.compile(r"Payment Plan Agreement", re.IGNORECASE)
license_plate_pattern = re.compile(
    r"Licence Plate Number\s*([A-Z0-9\- ]+)", re.IGNORECASE
)
insured_pattern = re.compile(
    r"(?:Owner|Applicant|Name of Insured\s*\(surname followed by given name\(s\)\))\s*[:\-]?\s*([A-Z][A-Z\s,.'\-]+?)(?:\sAddress|$)",
    re.IGNORECASE,
)
customer_copy_pattern = re.compile(r"customer copy", re.IGNORECASE)
validation_stamp_pattern = re.compile(r"NOT VALID UNLESS STAMPED BY", re.IGNORECASE)
time_of_validation_pattern = re.compile(r"TIME OF VALIDATION", re.IGNORECASE)
temporary_permit_pattern = re.compile(
    r"Temporary Operation Permit and Owner’s Certificate of Insurance",
    re.IGNORECASE,
)
agency_number_pattern = re.compile(
    r"Agency Number\s*[:#]?\s*([A-Z0-9]+)", re.IGNORECASE
)

not_valid_coords_box_offset = (-4.25, 23.77, 1.58, 58.95)
time_of_validation_offset = (0.0, 10.35, 0.0, 40)
time_stamp_offset = (0, 13, 0, 0)
time_of_validation_am_offset = (0, 0.7, 0, 0)
time_of_validation_pm_offset = (0, 21.9, 0, 0)


# -------------------- Utility Functions -------------------- #
def reverse_name(name):
    parts = [p for p in name.replace(",", " ").split() if p]
    return " ".join(parts[1:] + [parts[0]]).title() if len(parts) > 1 else name.title()


def format_transaction_timestamp(timestamp_str):
    return datetime.strptime(timestamp_str, "%Y%m%d%H%M%S")


def format_timestamp_mmmddyyyy_from_dt(dt):
    return dt.strftime("%b%d%Y")


def find_existing_timestamps(folder):
    timestamps = set()
    for pdf_file in Path(folder).rglob("*.pdf"):
        try:
            with fitz.open(pdf_file) as doc:
                if doc.page_count > 0:
                    ts_match = timestamp_pattern.search(
                        doc[0].get_text(clip=timestamp_rect)
                    )
                    if ts_match:
                        timestamps.add(ts_match.group(1))
        except:
            continue
    return timestamps


# -------------------- PDF Scanning -------------------- #
def scan_icbc_pdfs(input_dir, output_dir, max_docs):
    input_dir, output_dir = Path(input_dir), Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    existing_timestamps = find_existing_timestamps(output_dir)
    icbc_data = {}

    pdf_files = sorted(
        input_dir.glob("*.pdf"), key=lambda f: f.stat().st_mtime, reverse=True
    )[:max_docs]

    for pdf_path in pdf_files:
        try:
            with fitz.open(pdf_path) as doc:
                if doc.page_count == 0:
                    continue

                first_page = doc[0]
                text = first_page.get_text("text")
                ts_text = first_page.get_text(clip=timestamp_rect)
                payment_text = first_page.get_text(clip=payment_plan_rect)

                ts_match = timestamp_pattern.search(ts_text)
                timestamp = (
                    ts_match.group(1)
                    if ts_match and ts_match.group(1) not in existing_timestamps
                    else None
                )

                if payment_plan_pattern.search(payment_text):
                    continue

                license_plate_match = license_plate_pattern.search(text)
                license_plate = (
                    license_plate_match.group(1).strip().upper()
                    if license_plate_match
                    else None
                )

                insured_match = insured_pattern.search(text)
                insured_name = (
                    reverse_name(insured_match.group(1).strip())
                    if insured_match
                    else None
                )

                temp_permit_found = bool(temporary_permit_pattern.search(text))

                agency_match = agency_number_pattern.search(text)
                agency_number = (
                    agency_match.group(1).strip() if agency_match else "UNKNOWN"
                )

                customer_copy_pages, not_valid_coords, time_of_validation_coords = (
                    [],
                    [],
                    [],
                )
                for page_num, page in enumerate(doc):
                    clipped_customer_copy = page.get_text(clip=customer_copy_rect)
                    if customer_copy_pattern.search(clipped_customer_copy):
                        customer_copy_pages.append(page_num)
                    for w in page.get_text("blocks"):
                        word_text, coords = w[4], w[:4]
                        if validation_stamp_pattern.search(word_text):
                            not_valid_coords.append((page_num, coords))
                        if time_of_validation_pattern.search(word_text):
                            time_of_validation_coords.append((page_num, coords))

                icbc_data[pdf_path] = {
                    "transaction_timestamp": timestamp,
                    "license_plate": license_plate,
                    "insured_name": insured_name,
                    "temporary_operation_permit": temp_permit_found,
                    "agency_number": agency_number,
                    "customer_copy_pages": customer_copy_pages,
                    "not_valid_coords": not_valid_coords,
                    "time_of_validation_coords": time_of_validation_coords,
                }

        except Exception as e:
            print(f"❌ Error processing {pdf_path}: {e}")

    return icbc_data, len(pdf_files)


# -------------------- PDF Processing -------------------- #
def validation_annot(doc, info, ts_dt):
    for page_num, coords in info.get("not_valid_coords", []):
        page = doc[page_num]
        x0, y0, x1, y1 = coords
        dx0, dy0, dx1, dy1 = not_valid_coords_box_offset
        agency_rect = fitz.Rect(x0 + dx0, y0 + dy0, x1 + dx1, y1 + dy1)
        date_rect = fitz.Rect(
            agency_rect.x0 + time_stamp_offset[0],
            agency_rect.y0 + time_stamp_offset[1],
            agency_rect.x1 + time_stamp_offset[2],
            agency_rect.y1 + time_stamp_offset[3],
        )
        page.insert_textbox(
            agency_rect, info["agency_number"], fontname="spacembo", fontsize=9, align=1
        )
        page.insert_textbox(
            date_rect,
            ts_dt.strftime("%b %d, %Y"),
            fontname="spacemo",
            fontsize=9,
            align=1,
        )
    return doc


def stamp_time_of_validation(doc, info, ts_dt):
    for page_num, coords in info.get("time_of_validation_coords", []):
        page = doc[page_num]
        x0, y0, x1, y1 = coords
        dx0, dy0, dx1, dy1 = time_of_validation_offset
        if ts_dt.hour < 12:
            dx0 += time_of_validation_am_offset[0]
            dy0 += time_of_validation_am_offset[1]
        else:
            dx0 += time_of_validation_pm_offset[0]
            dy0 += time_of_validation_pm_offset[1]
        time_rect = fitz.Rect(x0 + dx0, y0 + dy0, x1 + dx1, y1 + dy1)
        page.insert_textbox(
            time_rect, ts_dt.strftime("%I:%M %p"), fontname="helv", fontsize=6, align=2
        )
    return doc


def save_batch_copy(doc, info, output_dir):
    batch_dir = Path(output_dir) / "ICBC Batch Copies"
    batch_dir.mkdir(parents=True, exist_ok=True)
    base_name = info.get("license_plate") or info.get("insured_name") or "UNKNOWN"
    batch_copy_path = batch_dir / f"{base_name}.pdf"
    doc.save(batch_copy_path)
    return batch_copy_path


def create_customer_copy(doc, info, output_dir):
    total_pages = doc.page_count
    customer_pages = info.get("customer_copy_pages", [])
    if info.get("temporary_operation_permit") and total_pages - 1 not in customer_pages:
        customer_pages.append(total_pages - 1)
    pages_to_delete = [i for i in range(total_pages) if i not in customer_pages]
    for page_num in reversed(pages_to_delete):
        doc.delete_page(page_num)
    base_name = info.get("license_plate") or info.get("insured_name") or "UNKNOWN"
    customer_copy_name = f"{base_name} (Customer Copies).pdf"
    customer_copy_path = Path(output_dir) / customer_copy_name
    doc.save(customer_copy_path)
    print(f"✅ Saved customer copy PDF: {customer_copy_path}")
    doc.close()
    return customer_copy_path


# -------------------- Main -------------------- #
def main():
    ascii_title = r"""
███████╗██╗   ██╗ ██████╗ ██████╗███████╗███████╗███████╗
██╔════╝██║   ██║██╔════╝██╔════╝██╔════╝██╔════╝██╔════╝
███████╗██║   ██║██║     ██║     █████╗  ███████╗███████╗
╚════██║██║   ██║██║     ██║     ██╔══╝  ╚════██║╚════██║
███████║╚██████╔╝╚██████╗╚██████╗███████╗███████║███████║                      
    """
    print(ascii_title)
    print("📄 ICBC PDF Stamping Script")
    print("📧 https://github.com/WebDevBernard/ICBC_E-Stamp_Tool")

    start_total = timeit.default_timer()

    data, total_scanned = scan_icbc_pdfs(
        DEFAULTS["input_dir"], DEFAULTS["output_dir"], DEFAULTS["number_of_pdfs"]
    )
    stamped_counter = 0

    for path, info in data.items():
        if not info["transaction_timestamp"]:
            continue
        ts = info["transaction_timestamp"]
        if ts in find_existing_timestamps(DEFAULTS["output_dir"]):
            continue  # Already stamped, skip

        ts_dt = format_transaction_timestamp(ts)
        try:
            doc = fitz.open(path)

            # Apply only validation and time-of-validation stamps
            doc = validation_annot(doc, info, ts_dt)
            doc = stamp_time_of_validation(doc, info, ts_dt)

            # Save batch copy WITHOUT top timestamp
            save_batch_copy(doc, info, DEFAULTS["output_dir"])

            # Create customer copy WITHOUT top timestamp
            create_customer_copy(doc, info, DEFAULTS["output_dir"])

            stamped_counter += 1
        except Exception as e:
            print(f"❌ Error processing {path}: {e}")

    end_total = timeit.default_timer()
    print(f"\nTotal PDFs scanned: {total_scanned}")
    print(f"Total PDFs stamped: {stamped_counter}")
    print(f"✅ Total script execution time: {end_total - start_total:.2f} seconds")
    print("\nExiting in ", end="")
    for i in range(3, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)
    print("\n")


if __name__ == "__main__":
    main()
