import fitz  # PyMuPDF
import re
from pathlib import Path
from datetime import datetime
from utils import write_to_new_docx
from constants import RECTS, REGEX_PATTERNS

# -------------------- Config -------------------- #
DATE_FORMAT = "%B %d, %Y"
TEMPLATE_FILE = Path.cwd() / "assets" / "Renewal Letter.docx"

# Unpack regex patterns
postal_code_regex = REGEX_PATTERNS["postal_code"]
dollar_regex = REGEX_PATTERNS["dollar"]
date_regex = REGEX_PATTERNS["date"]
address_regex = REGEX_PATTERNS["address"]
and_regex = REGEX_PATTERNS["and"]


# -------------------- PDF Extraction -------------------- #
def detect_insurer(doc, rects):
    first_page = doc[0]
    for insurer, cfg in rects["policy_type"].items():
        keyword = cfg.get("keyword")
        rect = cfg.get("rect")
        if rect:
            clip_text = first_page.get_text("text", clip=rect)
            if clip_text and keyword.lower() in clip_text.lower():
                return insurer
    return None


def get_text(doc, structured=True):
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
    matches = []
    for blocks in pages_dict.values():
        for block in blocks:
            text = " ".join(block["words"])
            if return_all:
                for match in regex.finditer(text):
                    matched_text = match.group(1) if match.lastindex else match.group()
                    matches.append(matched_text)
            else:
                match = regex.search(text)
                if match:
                    return match.group(1) if match.lastindex else match.group()
    return matches if return_all else None


# -------------------- Field Extraction -------------------- #
def deduplicate_field(value):
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
    pages_dict = get_text(doc)
    results = {}
    for field, cfg in field_mapping.items():
        results[field] = extract_single_field(pages_dict, cfg, doc)
        if insurer == "Intact":
            results[field] = deduplicate_field(results[field])
    return results


def extract_single_field(pages_dict, cfg, doc):
    pattern = cfg.get("pattern")
    rect_offset = cfg.get("rect")
    target_rect = cfg.get("target_rect")
    return_all = cfg.get("return_all", False)

    if not pattern and not rect_offset:
        return None
    if not pattern and rect_offset:
        first_page = doc[0]
        return first_page.get_text("text", clip=rect_offset).strip()
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
    for page in doc:
        blocks = page.get_text("blocks")
        for b in blocks:
            x0, y0, x1, y1, text = b[:5]
            if pattern.search(text):
                return fitz.Rect(x0, y0, x1, y1), page
    return None, None


def extract_with_pattern_and_offset(pages_dict, pattern, rect_offset):
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
    return page.get_textbox(target_rect).strip() if page else None


def extract_all_with_pattern_and_offset(pages_dict, pattern, rect_offset):
    matches = []
    for blocks in pages_dict.values():
        for block in blocks:
            text = " ".join(block["words"])
            if pattern.search(text):
                x0, y0, x1, y1 = block["coords"]
                word_rect = fitz.Rect(x0, y0, x1, y1)
                page = block["page"]
                target_rect = fitz.Rect(
                    word_rect.x0 + rect_offset.x0,
                    word_rect.y0 + rect_offset.y0,
                    word_rect.x1 + rect_offset.x1,
                    word_rect.y1 + rect_offset.y1,
                )
                extracted = page.get_textbox(target_rect).strip()
                if extracted:
                    matches.append(extracted)
    return matches if matches else None


def find_word_rect(pages_dict, search_text):
    for blocks in pages_dict.values():
        for block in blocks:
            text = " ".join(block["words"])
            if search_text in text:
                x0, y0, x1, y1 = block["coords"]
                return fitz.Rect(x0, y0, x1, y1), block["page"]
    return None, None


# -------------------- Data Mapping -------------------- #
def safe_get(data: dict, key: str):
    value = data.get(key)
    if value is None:
        return ""
    if isinstance(value, list):
        return value
    return value.strip() if isinstance(value, str) else value


def sum_dollar_amounts(amounts_text):
    if not amounts_text:
        return 0
    matches = dollar_regex.findall(amounts_text)
    total = sum(float(amount.replace(",", "")) for amount in matches)
    return int(total)


def map_extracted_data_for_renewal(extracted_data: dict, insurer: str) -> dict:
    """
    Minimal version: maps PDF data to template fields.
    """
    mapped = {
        "task": "Auto Renewal Letter",
        "insurer": insurer,
        "today": datetime.today().strftime(DATE_FORMAT),
        "named_insured": safe_get(extracted_data, "named_insured"),
        "policy_number": safe_get(extracted_data, "policy_number"),
        "effective_date": safe_get(extracted_data, "effective_date"),
        "premium_amount": safe_get(extracted_data, "premium_amount"),
        "risk_address_1": safe_get(extracted_data, "risk_address"),
    }

    # Example: convert premium amount to total
    premium_text = mapped.get("premium_amount")
    if premium_text:
        total = sum_dollar_amounts(premium_text)
        mapped["premium_amount"] = "${:,.2f}".format(total)

    return mapped


# -------------------- PDF Processing -------------------- #
def process_pdf_for_renewal(pdf_path, rects):
    doc = fitz.open(pdf_path)
    try:
        insurer = detect_insurer(doc, rects)
        if not insurer:
            return None, {}
        field_mapping = rects[insurer]
        data = extract_fields(doc, field_mapping, insurer=insurer)
        return insurer, data
    finally:
        doc.close()


# -------------------- Main -------------------- #
def main():
    downloads_dir = Path.home() / "Downloads"
    desktop_dir = Path.home() / "Desktop"
    pdf_files = list(downloads_dir.glob("*.pdf"))
    sorted_files = sorted(pdf_files, key=lambda f: f.stat().st_mtime, reverse=True)
    recent_files = sorted_files[:2]

    if not pdf_files:
        print("No PDF files found in Downloads folder.")
        return

    for pdf_file in recent_files:
        print(f"Processing: {pdf_file.name}")
        try:
            insurer, extracted_data = process_pdf_for_renewal(pdf_file, RECTS)
            if not insurer:
                print(f"  ⚠ Could not detect insurer. Skipping.\n")
                continue

            print(f"  ✓ Detected insurer: {insurer}")
            mapped_data = map_extracted_data_for_renewal(extracted_data, insurer)
            write_to_new_docx(TEMPLATE_FILE, mapped_data, output_dir=desktop_dir)
        except Exception as e:
            print(f"  ✗ Error processing {pdf_file.name}: {e}\n")

    print("Done!")


if __name__ == "__main__":
    main()
