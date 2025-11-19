import fitz
import re
from pathlib import Path
from tabulate import tabulate

# === Configuration ===
CONFIG = {
    # Feature toggles
    "extract_text": True,
    "extract_tables": False,
    "search_text": False,
    "extract_image": True,
    "calculate_offsets": True,  # NEW: Toggle for offset calculations
    # Directories
    "input_dir": Path.home() / "Downloads",
    "output_dir": Path.home() / "Desktop",
    # Search settings
    "search_pattern": r"(Owner\s|Applicant|Name of Insured \(surname followed by given name\(s\)\))",
    # Image extraction defaults
    "page_num": 13,
    "coords": (
        86.1500015258789,
        526.6749877929688,
        90,
        536.2930297851562,
    ),
    "image_prefix": "img",
    # Offset calculation settings
    "pattern_rect": (
        86.1500015258789,
        507.9750061035156,
        118.82599639892578,
        525.7430419921875,
    ),
    "target_rect": (
        86.1500015258789,
        526.6749877929688,
        90,
        536.2930297851562,
    ),
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


def get_tables(doc):
    """Extracts tables from all pages using PyMuPDF's table detection."""
    tables_dict = {}
    for page_num, page in enumerate(doc, start=1):
        tables = []
        page_tables = page.find_tables()

        for table_index, table in enumerate(page_tables.tables):
            rows = table.extract()
            # Filter out empty rows
            rows = [row for row in rows if row and any(cell for cell in row)]

            # Get cell rectangles safely
            cells_rects = []
            try:
                for row_idx in range(table.row_count):
                    row_rects = []
                    for col_idx in range(table.col_count):
                        try:
                            cell_rect = table.cells[row_idx * table.col_count + col_idx]
                            row_rects.append(cell_rect)
                        except (IndexError, TypeError):
                            row_rects.append(None)
                    cells_rects.append(row_rects)
            except Exception as e:
                print(
                    f"Warning: Could not extract cell rectangles for table {table_index}: {e}"
                )
                cells_rects = []

            table_data = {
                "table_index": table_index,
                "bbox": table.bbox,
                "rows": rows,
                "cells_rects": cells_rects,
                "row_count": len(rows),
                "col_count": len(rows[0]) if rows else 0,
            }
            tables.append(table_data)

        tables_dict[page_num] = tables

    return tables_dict


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
                continue

            file.write(
                tabulate(
                    table_data,
                    headers=["Keywords", "Coordinates"],
                    tablefmt="grid",
                    maxcolwidths=[None, None],
                )
            )
            file.write("\n\n")


def write_tables_to_file(file_path, tables_dict):
    """Save extracted tables as formatted output."""
    with open(file_path, "w", encoding="utf-8") as file:
        for page, tables in tables_dict.items():
            file.write(f"Page: {page}\n")

            if not tables:
                file.write("(No tables found on this page)\n\n")
                continue

            for table_data in tables:
                file.write(f"\n--- Table {table_data['table_index'] + 1} ---\n")
                file.write(
                    f"Rows: {table_data['row_count']}, Columns: {table_data['col_count']}\n"
                )
                file.write(f"Bounding Box: {table_data['bbox']}\n\n")

                # Write table content
                if table_data["rows"] and len(table_data["rows"]) > 0:
                    try:
                        file.write(
                            tabulate(
                                table_data["rows"],
                                headers="firstrow",
                                tablefmt="grid",
                                maxcolwidths=30,
                            )
                        )
                        file.write("\n\n")
                    except Exception as e:
                        file.write(f"(Error formatting table: {e})\n")
                        file.write(f"Raw data: {table_data['rows']}\n\n")
                else:
                    file.write("(Empty table)\n\n")

                # Write cell rectangles
                if table_data["cells_rects"]:
                    file.write("Cell Rectangles:\n")
                    for row_idx, row_rects in enumerate(table_data["cells_rects"]):
                        file.write(f"  Row {row_idx}:\n")
                        for col_idx, rect in enumerate(row_rects):
                            if rect:
                                file.write(f"    Col {col_idx}: {rect}\n")
                            else:
                                file.write(f"    Col {col_idx}: (unavailable)\n")
                    file.write("\n")


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


def offset_below(pattern_rect, target_rect):
    """
    Calculate offset for target rect below pattern rect.
    Prints offset values to copy into constants.
    Args:
        pattern_rect: tuple (x0, y0, x1, y1) - where pattern is found
        target_rect: tuple (x0, y0, x1, y1) - where you want to extract
    """
    p_x0, p_y0, p_x1, p_y1 = pattern_rect
    t_x0, t_y0, t_x1, t_y1 = target_rect
    dx0 = t_x0 - p_x0
    dy0 = t_y0 - p_y0  # distance from pattern's BOTTOM edge
    dx1 = t_x1 - p_x1
    dy1 = t_y1 - p_y1  # distance from pattern's BOTTOM edge
    print(f"  dx0={dx0:.2f}, dy0={dy0:.2f}, dx1={dx1:.2f}, dy1={dy1:.2f}")


def offset_right(pattern_rect, target_rect):
    """
    Calculate offset for target rect to the right of pattern rect.
    Prints offset values to copy into constants.
    Args:
        pattern_rect: tuple (x0, y0, x1, y1) - where pattern is found
        target_rect: tuple (x0, y0, x1, y1) - where you want to extract
    """
    p_x0, p_y0, p_x1, p_y1 = pattern_rect
    t_x0, t_y0, t_x1, t_y1 = target_rect
    # Calculate from pattern's LEFT edge (x0) to work with extract_with_pattern_and_offset
    dx0 = t_x0 - p_x0  # distance from pattern's LEFT edge to target's left
    dy0 = t_y0 - p_y0
    dx1 = t_x1 - p_x1  # keeps the width calculation correct
    dy1 = t_y1 - p_y1
    print(f"  dx0={dx0:.2f}, dy0={dy0:.2f}, dx1={dx1:.2f}, dy1={dy1:.2f}")


def offset_above(pattern_rect, target_rect):
    """
    Calculate offset for target rect above pattern rect.
    Prints offset values to copy into constants.
    Args:
        pattern_rect: tuple (x0, y0, x1, y1) - where pattern is found
        target_rect: tuple (x0, y0, x1, y1) - where you want to extract
    """
    p_x0, p_y0, p_x1, p_y1 = pattern_rect
    t_x0, t_y0, t_x1, t_y1 = target_rect
    dx0 = t_x0 - p_x0
    dy0 = t_y0 - p_y0  # distance from pattern's TOP edge
    dx1 = t_x1 - p_x1
    dy1 = t_y1 - p_y0  # distance from pattern's TOP edge
    print(f"  dx0={dx0:.2f}, dy0={dy0:.2f}, dx1={dx1:.2f}, dy1={dy1:.2f}")


def calculate_all_offsets(pattern_rect, target_rect):
    """Calculate and display all offset directions."""
    print("\n📐 Offset Calculations:")
    print(f"Pattern rect: {pattern_rect}")
    print(f"Target rect:  {target_rect}\n")

    print("Above:")
    offset_above(pattern_rect, target_rect)

    print("\nBelow:")
    offset_below(pattern_rect, target_rect)

    print("\nRight:")
    offset_right(pattern_rect, target_rect)
    print()


# === Main Execution ===
def main(config):
    input_dir = config["input_dir"]
    output_dir = config["output_dir"]
    pdf_files = list(input_dir.glob("*.pdf"))

    # 5️⃣ Calculate offsets (runs once, no PDF needed)
    if config["calculate_offsets"]:
        calculate_all_offsets(config["pattern_rect"], config["target_rect"])
        return  # Exit after calculation if no other features enabled

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

            # 2️⃣ Extract tables
            if config["extract_tables"]:
                table_data = get_tables(doc)
                output_file = output_dir / f"{pdf_file.stem}_tables.txt"
                write_tables_to_file(output_file, table_data)
                print(f"📊 Saved extracted tables to: {output_file}")

            # 3️⃣ Search text
            if config["search_text"]:
                name = search_text(doc, config["search_pattern"])
                if name:
                    print(f"🔍 Found name: {name}")
                else:
                    print("❌ No match found.")

            # 4️⃣ Extract image
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
