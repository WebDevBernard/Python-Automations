from pathlib import Path
import fitz
import re
import shutil

# --- Clip regions & regex patterns ---
timestamp_location = fitz.Rect(409.979, 63.8488, 576.0, 83.7455)
payment_plan_location = fitz.Rect(425.402, 35.9664, 557.916, 48.3001)
producer_location = fitz.Rect(198.0, 761.04, 255.011, 769.977)

timestamp_pattern = re.compile(r"Transaction Timestamp\s*(\d+)")
payment_plan_pattern = re.compile(r"Payment Plan Agreement", re.IGNORECASE)
license_plate_pattern = re.compile(r"Licence Plate Number\s*([A-Z0-9\- ]+)", re.IGNORECASE)
insured_pattern = re.compile(
    r"(?:Owner|Applicant|Name of Insured\s*\(surname followed by given name\(s\)\))\s*[:\-]?\s*([A-Z][A-Z\s,.'\-]+)",
    re.IGNORECASE
)
producer_pattern = re.compile(r"[A-Z]{1,3} - [A-Z]{1,3} - [A-Z0-9]+", re.IGNORECASE)


def reverse_name(name: str) -> str:
    parts = [p for p in name.replace(",", " ").strip().split() if p]
    return " ".join(parts[1:] + [parts[0]]).title() if len(parts) > 1 else name.title()


def scan_icbc_pdfs(input_dir, max_docs=None):
    input_dir = Path(input_dir)
    icbc_data = {}
    docs = sorted(input_dir.glob("*.pdf"), key=lambda f: f.stat().st_mtime, reverse=True)

    if max_docs:
        docs = docs[:max_docs]

    for doc_path in docs:
        with fitz.open(doc_path) as doc:
            page = doc[0]

            ts_match = timestamp_pattern.search(page.get_text(clip=timestamp_location))
            if not ts_match:
                continue

            if payment_plan_pattern.search(page.get_text(clip=payment_plan_location)):
                continue

            full_text = page.get_text("text")
            license_plate = (license_plate_pattern.search(full_text).group(1).strip().upper()
                             if license_plate_pattern.search(full_text) else None)
            insured_name = (reverse_name(insured_pattern.search(full_text).group(1).strip())
                            if insured_pattern.search(full_text) else None)

            producer_text = page.get_text(clip=producer_location).strip()
            producer_name = producer_pattern.search(producer_text).group(0) if producer_pattern.search(producer_text) else None

            icbc_data[doc_path] = {
                "transaction_timestamp": ts_match.group(1),
                "license_plate": license_plate,
                "insured_name": insured_name,
                "producer_name": producer_name,
            }

    return icbc_data, len(docs)


def find_existing_timestamps(root_folder: Path, base_name: str):
    existing_timestamps = set()
    for pdf_file in root_folder.rglob(f"{base_name}*.pdf"):
        try:
            with fitz.open(pdf_file) as doc:
                ts_match = timestamp_pattern.search(doc[0].get_text(clip=timestamp_location))
                if ts_match:
                    existing_timestamps.add(ts_match.group(1))
        except Exception as e:
            print(f"Warning: Failed to read {pdf_file}: {e}")
    return existing_timestamps


def copy_and_rename_icbc(input_dir, output_dir, producer_mapping=None, max_docs=None):
    input_dir, output_dir = Path(input_dir), Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    icbc_documents, scan_count = scan_icbc_pdfs(input_dir, max_docs=max_docs)
    copy_count = 0

    for pdf_path, data in icbc_documents.items():
        timestamp = data["transaction_timestamp"]
        license_plate = data["license_plate"]
        insured_name = data["insured_name"] or "Unknown"
        producer_name = data["producer_name"]

        base_name = license_plate if license_plate and license_plate != "NONLIC" else insured_name

        if timestamp in find_existing_timestamps(output_dir, base_name):
            continue

        folder_path = output_dir
        if producer_mapping and producer_name in producer_mapping:
            candidate_folder = output_dir / producer_mapping[producer_name]
            candidate_folder.mkdir(parents=True, exist_ok=True)
            folder_path = candidate_folder

        suffix = 0
        new_base = base_name
        while (folder_path / f"{new_base}.pdf").exists():
            suffix += 1
            new_base = f"{base_name}({suffix})"

        shutil.copy2(pdf_path, folder_path / f"{new_base}.pdf")
        copy_count += 1

    print(f"Scanned: {scan_count} documents; Copied: {copy_count} documents")


if __name__ == "__main__":
    # --- Defaults ---
    input_folder = Path.home() / "Downloads"          # Default input = Downloads
    folder_name = "ICBC Copies"                       # Default output folder
    max_pdfs = 10                                     # Default max copies
    producer_mapping = None                            # No mapping by default

    output_folder = Path.home() / folder_name
    output_folder.mkdir(parents=True, exist_ok=True)

    copy_and_rename_icbc(input_folder, output_folder, producer_mapping, max_docs=max_pdfs)
