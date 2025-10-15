import fitz
import re
from pathlib import Path
from datetime import datetime
import timeit
import os
import sys
import time

# -------------------- Defaults -------------------- #
DEFAULTS = {
    "number_of_pdfs": 10,
    "output_dir": str(
        (Path.home() / "Desktop" / "ICBC E-Stamp Copies")
        if (Path.home() / "Desktop").exists()
        else (Path(__file__).parent.parent / "ICBC E-Stamp Copies")
    ),
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
temporary_permit_pattern = re.compile(
    r"Temporary Operation Permit and Owner’s Certificate of Insurance", re.IGNORECASE
)
agency_number_pattern = re.compile(
    r"Agency Number\s*[:#]?\s*([A-Z0-9]+)", re.IGNORECASE
)
customer_copy_pattern = re.compile(r"customer copy", re.IGNORECASE)
validation_stamp_pattern = re.compile(r"NOT VALID UNLESS STAMPED BY", re.IGNORECASE)
time_of_validation_pattern = re.compile(r"TIME OF VALIDATION", re.IGNORECASE)

validation_stamp_coords_box_offset = (-4.25, 23.77, 1.58, 58.95)
time_of_validation_offset = (0.0, 10.35, 0.0, 40)
time_stamp_offset = (0, 13, 0, 0)
time_of_validation_am_offset = (0, -0.6, 0, 0)
time_of_validation_pm_offset = (0, 21.2, 0, 0)


# -------------------- Utility Functions -------------------- #
def progressbar(it, prefix="", size=60, out=sys.stdout):
    count = len(it)
    start = time.time()

    def show(j):
        x = int(size * j / count)
        remaining = ((time.time() - start) / j) * (count - j) if j else 0
        mins, sec = divmod(remaining, 60)
        time_str = f"{int(mins):02}:{sec:03.1f}"
        print(
            f"{prefix}[{'█' * x}{'.' * (size - x)}] {j}/{count} Est wait {time_str}",
            end="\r",
            file=out,
            flush=True,
        )

    if len(it) > 0:
        show(0.1)
        for i, item in enumerate(it):
            yield item
            show(i + 1)
        print(flush=True, file=out)


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


def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path


def reverse_insured_name(name):
    if not name:
        return ""

    name = re.sub(r"\s+", " ", name.strip())
    if name.endswith(("Ltd", "Inc")):
        return name
    name = name.replace(",", "")
    parts = name.split(" ")
    if len(parts) == 1:
        return name
    return " ".join(parts[1:] + [parts[0]])


def search_insured_name(full_text_first_page):
    match = re.search(
        r"(?:Owner\s|Applicant|Name of Insured \(surname followed by given name\(s\)\))\s*\n([^\n]+)",
        full_text_first_page,
        re.IGNORECASE,
    )
    if match:
        name = match.group(1)
        name = re.sub(r"[.:/\\*?\"<>|]", "", name)
        name = re.sub(r"\s+", " ", name).strip().title()
        return name
    return None


# -------------------- PDF Stamping Functions -------------------- #
def validation_stamp(doc, info, ts_dt):
    for page_num, coords in info.get("validation_stamp_coords", []):
        page = doc[page_num]
        x0, y0, x1, y1 = coords
        dx0, dy0, dx1, dy1 = validation_stamp_coords_box_offset
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
            time_rect, ts_dt.strftime("%I:%M"), fontname="helv", fontsize=6, align=2
        )
    return doc


def get_base_name(info):
    transaction_timestamp = info.get("transaction_timestamp") or ""
    license_plate = (info.get("license_plate") or "").strip().upper()
    insured_name = (info.get("insured_name") or "").strip()
    insured_name = re.sub(r"[.:/\\*?\"<>|]", "", insured_name)
    insured_name = re.sub(r"\s+", " ", insured_name).strip()
    insured_name = insured_name.title() if insured_name else ""
    if license_plate:
        base_name = license_plate
    elif insured_name:
        base_name = insured_name
    elif transaction_timestamp:
        base_name = transaction_timestamp
    else:
        base_name = "UNKNOWN"
    return base_name


def save_batch_copy(doc, info, output_dir):
    batch_dir = Path(output_dir) / "ICBC Batch Copies"
    batch_dir.mkdir(parents=True, exist_ok=True)
    base_name = get_base_name(info)
    batch_copy_path = batch_dir / f"{base_name}.pdf"
    batch_copy_path = Path(unique_file_name(batch_copy_path))
    doc.save(batch_copy_path, garbage=4, deflate=True)
    return batch_copy_path


def save_customer_copy(doc, info, output_dir):
    total_pages = doc.page_count
    customer_pages = info.get("customer_copy_pages", [])
    if info.get("temporary_operation_permit") and total_pages - 1 not in customer_pages:
        customer_pages.append(total_pages - 1)
    pages_to_delete = [i for i in range(total_pages) if i not in customer_pages]
    for page_num in reversed(pages_to_delete):
        doc.delete_page(page_num)
    base_name = get_base_name(info)
    customer_copy_name = f"{base_name} (Customer Copies).pdf"
    customer_copy_path = Path(output_dir) / customer_copy_name
    customer_copy_path = Path(unique_file_name(customer_copy_path))
    doc.save(customer_copy_path, garbage=4, deflate=True)
    return customer_copy_path


# -------------------- Main Function -------------------- #
def icbc_e_stamp_tool():
    print("📄 ICBC E-Stamp Tool")
    start_total = timeit.default_timer()

    input_dir = DEFAULTS["input_dir"]
    output_dir = DEFAULTS["output_dir"]
    max_docs = DEFAULTS["number_of_pdfs"]

    pdf_files = sorted(
        Path(input_dir).glob("*.pdf"), key=lambda f: f.stat().st_mtime, reverse=True
    )[:max_docs]

    icbc_data = {}
    existing_timestamps = find_existing_timestamps(output_dir)

    # -------------------- Stage 1: Scan PDFs -------------------- #
    print("🔍 Scanning PDFs...")
    for pdf_path in progressbar(pdf_files, prefix="Scanning PDFs: ", size=40):
        try:
            with fitz.open(pdf_path) as doc:
                if doc.page_count == 0:
                    continue

                first_page = doc[0]
                full_text_first_page = "\n".join([first_page.get_text("text")])
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

                license_plate_match = license_plate_pattern.search(full_text_first_page)
                license_plate = (
                    license_plate_match.group(1).strip().upper()
                    if license_plate_match
                    else None
                )

                insured_name = reverse_insured_name(
                    search_insured_name(full_text_first_page)
                )

                temp_permit_found = bool(
                    temporary_permit_pattern.search(full_text_first_page)
                )

                agency_match = agency_number_pattern.search(full_text_first_page)
                agency_number = (
                    agency_match.group(1).strip() if agency_match else "UNKNOWN"
                )

                customer_copy_pages = []
                validation_stamp_coords = []
                time_of_validation_coords = []

                for page_num, page in enumerate(doc):
                    clipped_customer_copy = page.get_text(clip=customer_copy_rect)
                    if customer_copy_pattern.search(clipped_customer_copy):
                        customer_copy_pages.append(page_num)

                    for block in page.get_text("blocks"):
                        word_text, coords = block[4], block[:4]
                        if validation_stamp_pattern.search(word_text):
                            validation_stamp_coords.append((page_num, coords))
                        if time_of_validation_pattern.search(word_text):
                            time_of_validation_coords.append((page_num, coords))

                icbc_data[pdf_path] = {
                    "transaction_timestamp": timestamp,
                    "license_plate": license_plate,
                    "insured_name": insured_name,
                    "temporary_operation_permit": temp_permit_found,
                    "agency_number": agency_number,
                    "customer_copy_pages": customer_copy_pages,
                    "validation_stamp_coords": validation_stamp_coords,
                    "time_of_validation_coords": time_of_validation_coords,
                }

        except Exception as e:
            print(f"❌ Error scanning {pdf_path}: {e}")

    total_scanned = len(pdf_files)

    # -------------------- Stage 2: Process PDFs -------------------- #
    print("\n✍️ Processing PDFs...")
    stamped_counter = 0
    for path, info in progressbar(
        list(icbc_data.items()), prefix="Processing PDFs: ", size=40
    ):
        if not info["transaction_timestamp"]:
            continue
        ts = info["transaction_timestamp"]
        if ts in find_existing_timestamps(output_dir):
            continue

        ts_dt = format_transaction_timestamp(ts)

        try:
            doc_batch = fitz.open(path)
            doc_customer = fitz.open(path)

            doc_batch = validation_stamp(doc_batch, info, ts_dt)
            doc_batch = stamp_time_of_validation(doc_batch, info, ts_dt)
            doc_customer = validation_stamp(doc_customer, info, ts_dt)
            doc_customer = stamp_time_of_validation(doc_customer, info, ts_dt)

            save_batch_copy(doc_batch, info, output_dir)
            save_customer_copy(doc_customer, info, output_dir)

            stamped_counter += 1

        except Exception as e:
            print(f"❌ Error processing {path}: {e}")

    end_total = timeit.default_timer()
    print(f"\nTotal PDFs scanned: {total_scanned}")
    print(f"Total PDFs stamped: {stamped_counter}")
    print(f"✅ Total script execution time: {end_total - start_total:.2f} seconds")


if __name__ == "__main__":
    icbc_e_stamp_tool()
