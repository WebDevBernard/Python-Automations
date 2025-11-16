import fitz
import re
from pathlib import Path
from tabulate import tabulate

# === Configuration ===
CONFIG = {
    # Feature toggles
    "extract_text": True,
    "search_text": False,
    "extract_image": False,
    # Directories
    "input_dir": Path.home() / "Downloads",
    "output_dir": Path.home() / "Desktop",
    # Search settings
    "search_pattern": r"(Owner\s|Applicant|Name of Insured \(surname followed by given name\(s\)\))",
    # Image extraction defaults
    "page_num": 1,
    "coords": (198.0, 752.729736328125, 255.011, 769.977),
    "image_prefix": "img",
}


# === Functions ===
def get_text(doc, structured=True):
    """Extracts text; structured=True uses block mode."""
    pages_dict = {}
    mode = "blocks" if structured else "text"
    for page_num, page in enumerate(doc, start=1):
        blocks = []
        for block in page.get_text(mode):
            coords = tuple(block[:4])
            text_lines = block[4].split("\n")
            blocks.append({"words": text_lines, "coords": coords})
        pages_dict[page_num] = blocks
    return pages_dict


def search_text(doc, pattern):
    """Search for a specific text pattern on the first page."""
    text = doc[0].get_text("text")
    match = re.search(rf"(?:{pattern})\s*\n([^\n]+)", text, re.IGNORECASE)
    print(text)
    if match:
        name = match.group(1)
        name = re.sub(r"[.:/\\*?\"<>|]", "", name)
        name = re.sub(r"\s+", " ", name).strip().title()
        return name
    return None


def write_txt_to_file(file_path, field_dict):
    """Save extracted text data as a formatted table."""
    with open(file_path, "w", encoding="utf-8") as file:
        for page, data in field_dict.items():
            file.write(f"Page: {page}\n")
            table_data = [
                ["\n".join(item["words"]), str(item["coords"])] for item in data
            ]
            if not table_data:
                file.write("(No text found on this page)\n\n")
                continue  # Skip tabulate call for empty pages

            file.write(
                tabulate(
                    table_data,
                    headers=["Keywords", "Coordinates"],
                    tablefmt="grid",
                    maxcolwidths=[50, None],
                )
            )
            file.write("\n\n")


def save_region_as_png(
    doc, pdf_file, page_num, coords, prefix="Region", output_dir=Path.home()
):
    """Highlight a region and save the page as PNG."""
    page = doc[page_num - 1]
    highlight_rect = fitz.Rect(*coords)
    annot = page.add_highlight_annot(highlight_rect)
    annot.set_colors(stroke=fitz.pdfcolor["pink"])
    annot.update()
    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
    output_file = output_dir / f"{pdf_file.stem}_{prefix}.png"
    pix.save(output_file)
    print(f"🖼️  Saved highlighted image to: {output_file}")


# === Main Execution ===
def main(config):
    input_dir = config["input_dir"]
    output_dir = config["output_dir"]
    pdf_files = list(input_dir.glob("*.pdf"))

    if not pdf_files:
        print("⚠️ No PDF files found in input directory.")
        return

    for pdf_file in pdf_files:
        print(f"\n📄 Processing: {pdf_file.name}")
        with fitz.open(pdf_file) as doc:

            # 1️⃣ Extract text
            if config["extract_text"]:
                text_data = get_text(doc, structured=True)
                output_file = output_dir / f"{pdf_file.stem}.txt"
                write_txt_to_file(output_file, text_data)
                print(f"✅ Saved extracted text to: {output_file}")

            # 2️⃣ Search text
            if config["search_text"]:
                name = search_text(doc, config["search_pattern"])
                if name:
                    print(f"🔍 Found name: {name}")
                else:
                    print("❌ No match found.")

            # 3️⃣ Extract image
            if config["extract_image"]:
                save_region_as_png(
                    doc,
                    pdf_file,
                    page_num=config["page_num"],
                    coords=config["coords"],
                    prefix=config["image_prefix"],
                    output_dir=output_dir,
                )


if __name__ == "__main__":
    main(CONFIG)
