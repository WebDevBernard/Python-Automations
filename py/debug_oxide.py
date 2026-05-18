import re
from pdf_oxide import PdfDocument
from pathlib import Path
from tabulate import tabulate

# === Configuration ===
CONFIG = {
    "input_dir": Path.home() / "Downloads",
    "output_dir": Path(r"C:\Users\berna\OneDrive\Desktop"),
}

LICENSE_PLATE_RE = re.compile(
    r"Licence Plate Number\s*([A-Z0-9\- ]+)", re.IGNORECASE
)


def search_license_plate(doc):
    """Search all pages for licence plate number and return the match group."""
    for page_num in range(doc.page_count()):
        text = doc.extract_text(page_num)
        match = LICENSE_PLATE_RE.search(text)
        if match:
            return match.group(1).strip()
    return None


def extract_and_write(doc, pdf_file, output_dir):
    """Extract text lines with coordinates from all pages and save as table."""
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / f"{pdf_file.stem}.txt"

    with open(output_file, "w", encoding="utf-8") as f:
        for page_num in range(doc.page_count()):
            print(f"  Page {page_num + 1}...")
            f.write(f"Page: {page_num + 1}\n")

            try:
                lines = doc.extract_text_lines(page_num)
            except Exception as e:
                f.write(f"(Error extracting text: {e})\n\n")
                print(f"    Error: {e}")
                continue

            if not lines:
                f.write("(No text found on this page)\n\n")
                continue

            # Sort top-to-bottom (Y descending), left-to-right (X ascending).
            # PDF coords: origin bottom-left, so higher Y = higher on page.
            # Lines within 5pt vertically are treated as same row.
            def sort_key(line):
                x0, y0, _, _ = line.bbox
                row = round(y0 / 5) * 5  # snap to 5pt grid for same-line grouping
                return (-row, x0)

            lines = sorted(lines, key=sort_key)

            table_data = [[line.text, str(line.bbox)] for line in lines]
            f.write(
                tabulate(
                    table_data,
                    headers=["Word", "BBox"],
                    tablefmt="grid",
                    maxcolwidths=[None, None],
                )
            )
            f.write("\n\n")

    print(f"Saved extracted text to: {output_file}")


def main(config):
    input_dir = config["input_dir"]
    output_dir = config["output_dir"]
    pdf_files = list(input_dir.glob("*.pdf"))

    if not pdf_files:
        print("No PDF files found in input directory.")
        return

    for pdf_file in pdf_files:
        print(f"Processing: {pdf_file.name}")
        try:
            with PdfDocument(str(pdf_file)) as doc:
                extract_and_write(doc, pdf_file, output_dir)
                plate = search_license_plate(doc)
                if plate:
                    print(f"  Licence plate: {plate}")
                else:
                    print("  Licence plate: not found")
        except Exception as e:
            print(f"ERROR: {e}")


if __name__ == "__main__":
    main(CONFIG)
