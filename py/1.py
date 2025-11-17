import fitz  # PyMuPDF
import re
from pathlib import Path
from constants import RECTS, REGEX_PATTERNS
from utils import write_to_new_docx, unique_file_name


# ----------------- INSURER DETECTION ----------------- #
def detect_insurer(doc, rects):
    """Detect insurer by keyword in policy_type rects."""
    first_page = doc[0]
    for insurer, cfg in rects["policy_type"].items():
        keyword = cfg.get("keyword")
        rect = cfg.get("rect")
        if rect:
            clip_text = first_page.get_text("text", clip=rect)
            if clip_text and keyword.lower() in clip_text.lower():
                return insurer
    return None


# ----------------- TEXT EXTRACTION ----------------- #
def get_text(doc, structured=True):
    """Extract text from PDF as pages -> blocks."""
    pages_dict = {}
    mode = "blocks" if structured else "text"
    for page_num, page in enumerate(doc, start=1):
        blocks = []
        for block in page.get_text(mode):
            coords = tuple(block[:4])
            text_lines = block[4].split("\n")
            blocks.append({"words": text_lines, "coords": coords, "page": page})
        pages_dict[page_num] = blocks
    return pages_dict


def search_text(pages_dict, regex, return_all=False):
    """Search regex across all blocks in all pages.

    Args:
        pages_dict: Dictionary of page data
        regex: Compiled regex pattern
        return_all: If True, return list of all matches; if False, return first match

    Returns:
        Single match string or list of matches (depending on return_all)
    """
    matches = []
    for blocks in pages_dict.values():
        for block in blocks:
            text = " ".join(block["words"])
            if return_all:
                # Find all matches in this block
                for match in regex.finditer(text):
                    matched_text = match.group(1) if match.lastindex else match.group()
                    matches.append(matched_text)
            else:
                # Return first match found
                match = regex.search(text)
                if match:
                    return match.group(1) if match.lastindex else match.group()

    if return_all:
        return matches if matches else None
    return None


# ----------------- FIELD EXTRACTION ----------------- #
def deduplicate_field(value):
    """Remove duplicates from field values while preserving order."""
    if isinstance(value, list):
        seen = set()
        result = []
        for item in value:
            if item and item not in seen:
                seen.add(item)
                result.append(item)
        return result if len(result) > 1 else (result[0] if result else None)
    return value


def extract_fields(doc, field_mapping, insurer=None):
    """Extract multiple fields based on field_mapping."""
    pages_dict = get_text(doc)
    results = {}
    for field, cfg in field_mapping.items():
        results[field] = extract_single_field(pages_dict, cfg, doc)
        # Deduplicate list fields only for Intact
        if insurer == "Intact":
            results[field] = deduplicate_field(results[field])
    return results


def extract_single_field(pages_dict, cfg, doc):
    """Extract a single field based on pattern and/or rect offset.

    If cfg has 'return_all': True, will return list of all matches.
    """
    pattern = cfg.get("pattern")
    rect_offset = cfg.get("rect")
    target_rect = cfg.get("target_rect")
    return_all = cfg.get(
        "return_all", False
    )  # NEW: Check if we should return all matches

    if not pattern and not rect_offset:
        return None

    # If only rect (no pattern), use it as absolute rect to extract text
    if not pattern and rect_offset:
        return extract_text_from_absolute_rect(doc, rect_offset)

    # Automatically compute rect offset if target_rect is provided
    if pattern and target_rect and not rect_offset:
        rect_offset = compute_offset_from_target(doc, pattern, target_rect)

    if pattern and rect_offset:
        if return_all:
            return extract_all_with_pattern_and_offset(pages_dict, pattern, rect_offset)
        else:
            return extract_with_pattern_and_offset(pages_dict, pattern, rect_offset)

    if pattern:
        return search_text(pages_dict, pattern, return_all=return_all)

    return None


def compute_offset_from_target(doc, pattern, target_rect):
    """Find the matched word rect in the PDF and compute relative offset."""
    matched_rect, page = find_match_rect(doc, pattern)
    if matched_rect:
        return fitz.Rect(
            target_rect.x0 - matched_rect.x0,
            target_rect.y0 - matched_rect.y0,
            target_rect.x1 - matched_rect.x1,
            target_rect.y1 - matched_rect.y1,
        )
    return None


def find_match_rect(doc, pattern):
    """Return the rect and page of matched text for a pattern."""
    for page in doc:
        blocks = page.get_text("blocks")
        for b in blocks:
            x0, y0, x1, y1, text = b[:5]
            if pattern.search(text):
                return fitz.Rect(x0, y0, x1, y1), page
    return None, None


def extract_with_pattern_and_offset(pages_dict, pattern, rect_offset):
    """Find pattern in blocks, then extract text from rectangle offset relative to match."""
    found_text = search_text(pages_dict, pattern)
    if not found_text:
        return None

    word_rect, page = find_word_rect(pages_dict, found_text)
    if not word_rect:
        return found_text

    target_rect = fitz.Rect(
        word_rect.x0 + rect_offset.x0,
        word_rect.y0 + rect_offset.y0,
        word_rect.x1 + rect_offset.x1,
        word_rect.y1 + rect_offset.y1,
    )
    return extract_text_from_rect_on_page(page, target_rect)


def extract_all_with_pattern_and_offset(pages_dict, pattern, rect_offset):
    """Find ALL instances of pattern and extract text from offset rects for each."""
    matches = []

    for blocks in pages_dict.values():
        for block in blocks:
            text = " ".join(block["words"])
            if pattern.search(text):
                # Found a match - extract from offset rect
                x0, y0, x1, y1 = block["coords"]
                word_rect = fitz.Rect(x0, y0, x1, y1)
                page = block["page"]

                target_rect = fitz.Rect(
                    word_rect.x0 + rect_offset.x0,
                    word_rect.y0 + rect_offset.y0,
                    word_rect.x1 + rect_offset.x1,
                    word_rect.y1 + rect_offset.y1,
                )

                extracted = extract_text_from_rect_on_page(page, target_rect)
                if extracted:
                    matches.append(extracted)

    return matches if matches else None


def find_word_rect(pages_dict, search_text):
    """Return the rect and page of matched text."""
    for blocks in pages_dict.values():
        for block in blocks:
            text = " ".join(block["words"])
            if search_text in text:
                x0, y0, x1, y1 = block["coords"]
                return fitz.Rect(x0, y0, x1, y1), block["page"]
    return None, None


def extract_text_from_rect(pages_dict, rect):
    """Extract text from blocks intersecting rect across all pages."""
    extracted = []
    for blocks in pages_dict.values():
        for block in blocks:
            block_rect = fitz.Rect(*block["coords"])
            if block_rect.intersects(rect):
                extracted.append(" ".join(block["words"]).strip())
    return " ".join(extracted) if extracted else None


def extract_text_from_rect_on_page(page, rect):
    """Extract text from a single page using fitz.Rect."""
    return page.get_textbox(rect).strip() if page else None


def extract_text_from_absolute_rect(doc, rect):
    """Extract text from an absolute rectangle on the first page."""
    first_page = doc[0]
    text = first_page.get_text("text", clip=rect).strip()
    return text if text else None


# ----------------- UTILITY ----------------- #
def print_fields(fields):
    """Pretty print extracted fields."""
    for key, value in fields.items():
        if isinstance(value, list):
            print(f"{key:20s}: {len(value)} items")
            for i, item in enumerate(value, 1):
                print(f"  [{i}] {item}")
        else:
            print(f"{key:20s}: {value}")


# ----------------- MAIN ----------------- #
def main():
    downloads_dir = Path.home() / "Downloads"
    pdf_files = list(downloads_dir.glob("*.pdf"))
    sorted_files = sorted(pdf_files, key=lambda f: f.stat().st_mtime, reverse=True)
    recent_files = sorted_files[:2]

    if not pdf_files:
        print("No PDF files found in Downloads folder.")
        return

    for pdf_file in recent_files:
        print(f"\nProcessing: {pdf_file.name}")
        with fitz.open(pdf_file) as doc:
            try:
                insurer = detect_insurer(doc, RECTS)
                if insurer:
                    print(f"  Insurer detected: {insurer}")
                    fields = extract_fields(
                        doc, RECTS[insurer], insurer=insurer
                    )  # ← Add insurer parameter
                    print_fields(fields)
                else:
                    print(f"  ✗ No insurer detected")
            except Exception as e:
                print(f"  ✗ Error processing {pdf_file.name}: {e}")


if __name__ == "__main__":
    main()
