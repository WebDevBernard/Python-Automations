import fitz
import re
from pathlib import Path
from tabulate import tabulate

# Directories
input_dir = Path.home() / "Downloads"
output_dir = Path.home() / "Desktop"
output_dir.mkdir(parents=True, exist_ok=True)

# Find PDF files
pdf_files = list(input_dir.glob("*.pdf"))


def get_text(doc, mode="blocks"):
    pages_dict = {}
    for page_num, page in enumerate(doc, start=1):
        blocks = []
        for block in page.get_text(mode):
            coords = tuple(block[:4])
            text_lines = block[4].split("\n")
            blocks.append({"words": text_lines, "coords": coords})
        pages_dict[page_num] = blocks
    return pages_dict


def search_text(pdf_path):

    doc = fitz.open(pdf_path)
    match = re.search(
        r"(?:Owner\s|Applicant|Name of Insured \(surname followed by given name\(s\)\))\s*\n([^\n]+)",
        "\n".join([doc[0].get_text("text")]),
        re.IGNORECASE,
    )
    print("\n".join([doc[0].get_text("text")]))
    if match:
        name = match.group(1)
        name = re.sub(r"[.:/\\*?\"<>|]", "", name)
        name = re.sub(r"\s+", " ", name).strip().title()
        return name

    return None


def write_txt_to_file(file_path, field_dict):
    with open(file_path, "w", encoding="utf-8") as file:
        for page, data in field_dict.items():
            file.write(f"Page: {page}\n")
            table_data = []
            for item in data:
                table_data.append(["\n".join(item["words"]), str(item["coords"])])
            try:
                file.write(
                    tabulate(
                        table_data,
                        headers=["Keywords", "Coordinates"],
                        tablefmt="grid",
                        maxcolwidths=[50, 50],
                    )
                )
                file.write("\n\n")
            except IndexError:
                continue


def save_region_as_png(doc, page_num, coords, prefix="Region"):

    page = doc[page_num - 1]

    # Highlight the region
    highlight_rect = fitz.Rect(*coords)
    annot = page.add_highlight_annot(highlight_rect)
    annot.set_colors(stroke=fitz.pdfcolor["pink"])  # yellow highlight
    annot.update()

    # Render full page to pixmap with highlight visible
    pix = page.get_pixmap(
        matrix=fitz.Matrix(3, 3)
    )  # highlights are rendered by default

    # Save the image
    output_file = output_dir / f"{pdf_file.stem}.png"
    pix.save(output_file)
    print(f"Saved highlighted page to: {output_file}")


# === Toggle options ===
# Choose which feature to run:
#   "extract_text"  → extract and save all text
#   "search_text"   → search a pattern in the PDFs
#   "extract_image" → highlight region and save image
MODE = "extract_text"

SEARCH_PATTERN = r"(Owner\s|Applicant|Name of Insured \(surname followed by given name\(s\)\))"  # Used when MODE == "search_text"

for pdf_file in pdf_files:
    with fitz.open(pdf_file) as doc:
        if MODE == "extract_text":
            text_data = get_text(doc)
            output_file = output_dir / f"{pdf_file.stem}.txt"
            write_txt_to_file(output_file, text_data)
            print(f"✅ Saved text to: {output_file}")

        elif MODE == "extract_image":
            # Example coordinates — replace with your own
            coords = (198.0, 761.04, 255.011, 769.977)
            save_region_as_png(doc, page_num=1, coords=coords)

        else:
            print(f"⚠️ Unknown mode: {MODE}")
