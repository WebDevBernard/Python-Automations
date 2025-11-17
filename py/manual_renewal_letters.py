import fitz  # PyMuPDF
import re
from pathlib import Path
from datetime import datetime
from utils import write_to_new_docx
from constants import RECTS, REGEX_PATTERNS

DATE_FORMAT = "%B %d, %Y"
TEMPLATE_FILE = Path.cwd() / "assets" / "Renewal Letter.docx"

# Unpack regex patterns for convenience
postal_code_regex = REGEX_PATTERNS["postal_code"]
dollar_regex = REGEX_PATTERNS["dollar"]
date_regex = REGEX_PATTERNS["date"]
address_regex = REGEX_PATTERNS["address"]
and_regex = REGEX_PATTERNS["and"]


# --------------- INSURER DETECTION -----------------
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


# --------------- TEXT EXTRACTION -----------------
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
    """
    Search regex across all blocks in all pages.

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


# --------------- FIELD EXTRACTION -----------------
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
    """
    Extract a single field based on pattern and/or rect offset.

    If cfg has 'return_all': True, will return list of all matches.
    """
    pattern = cfg.get("pattern")
    rect_offset = cfg.get("rect")
    target_rect = cfg.get("target_rect")
    return_all = cfg.get("return_all", False)

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


# --------------- HELPER FUNCTIONS -----------------
def find_index(regex, text):
    """Find index of first line matching regex in text."""
    if not text:
        return -1
    lines = text.split("\n") if isinstance(text, str) else text
    for index, line in enumerate(lines):
        if re.search(regex, line):
            return index
    return -1


def join_and_format_names(names):
    """Join multiple names with proper formatting."""
    if len(names) == 1:
        return names[0].title()
    else:
        return (
            ", ".join(names[:-1]).strip().title() + " and " + names[-1].strip().title()
        )


def address_one_title_case(sentence):
    """Title case for address line 1 (handles ordinals like 1st, 2nd)."""
    ordinal_pattern = re.compile(r"\b\d+(st|nd|rd|th)\b")
    return " ".join(
        word.lower() if ordinal_pattern.match(word) else word.capitalize()
        for word in sentence.split()
    )


def address_two_title_case(strings_list):
    """Title case for address line 2 (city, province)."""
    words = strings_list.split()
    capitalized_words = [
        word.strip().capitalize() if len(word) > 2 else word for word in words
    ]
    return " ".join(capitalized_words)


def risk_address_title_case(address):
    """Title case for risk address (preserve 2-letter province codes)."""
    parts = address.split()
    if not parts:
        return address

    last_part = parts[-1]
    if len(last_part) == 2:
        last_part = last_part.upper()

    titlecased_address_parts = []
    for part in parts[:-1]:
        if (
            len(part) > 2
            and part[:-2].isdigit()
            and part[-2:].lower() in ["th", "rd", "nd", "st"]
        ):
            titlecased_address_parts.append(part.lower())
        else:
            titlecased_address_parts.append(part.title())

    return (
        " ".join(titlecased_address_parts) + " " + last_part
        if titlecased_address_parts
        else last_part
    )


def sum_dollar_amounts(amounts_text):
    """Sum all dollar amounts found in text."""
    if not amounts_text:
        return 0

    matches = dollar_regex.findall(amounts_text)
    total = sum(float(amount.replace(",", "")) for amount in matches)
    return int(total)


def safe_get(data: dict, key: str):
    """Safely get a value from dict, handling None values and lists."""
    value = data.get(key)
    if value is None:
        return ""
    if isinstance(value, list):
        return value  # Return list as-is for fields with return_all=True
    return value.strip() if isinstance(value, str) else value


# --------------- DATA MAPPING -----------------
def map_extracted_data_for_renewal(extracted_data: dict, insurer: str) -> dict:
    """
    Maps extracted PDF data to Word template format with comprehensive formatting.
    Applies insurer-specific normalization rules.
    """
    mapped = {
        "task": "Auto Renewal Letter",
        "broker_name": "",
        "on_behalf": "",
        "insurer": insurer,
        "today": datetime.today().strftime(DATE_FORMAT),
    }

    # --------------- NAME AND ADDRESS PARSING -----------------
    name_and_address = safe_get(extracted_data, "name_and_address")
    if name_and_address:
        address_lines = [
            line.strip() for line in name_and_address.split("\n") if line.strip()
        ]

        # Find where address starts (first line with number, PO Box, or Unit)
        address_index = find_index(address_regex, address_lines)
        if address_index == -1:
            address_index = len(address_lines)  # No address found

        # Find postal code position
        pc_index = find_index(postal_code_regex, address_lines)

        # Extract and format named insured (special handling for Intact)
        if insurer == "Intact":
            # Join names before address
            name_string = " ".join(address_lines[:address_index])
            # Split on '&' to separate individuals
            individual_names = name_string.split("&")
            processed_names = []

            for name in individual_names:
                # Clean up and split by commas
                name_parts = [part.strip().replace(":", "") for part in name.split(",")]
                if len(name_parts) == 2:
                    # Reverse "Last, First" to "First Last"
                    processed_names.append(f"{name_parts[1]} {name_parts[0]}".title())
                else:
                    # Already in "First Last" format
                    processed_names.append(name_parts[0].title())

            mapped["named_insured"] = " and ".join(processed_names)
        else:
            # For other insurers, remove "&" and "and", split by comma
            names = re.sub(and_regex, "", ", ".join(address_lines[:address_index]))
            mapped["named_insured"] = (
                join_and_format_names(names.split(", "))
                .replace("  ", " ")
                .replace(":", "")
            )

        # Extract mailing address
        if pc_index != -1:
            # Address line 1 (street)
            if address_index < pc_index - 1:
                mapped["address_line_one"] = address_one_title_case(
                    " ".join(address_lines[address_index : pc_index - 1])
                )
            else:
                mapped["address_line_one"] = address_one_title_case(
                    " ".join(address_lines[address_index:pc_index])
                )

            # Address line 2 (city, province)
            city_province_p_code = " ".join(
                address_lines[address_index + 1 : pc_index + 1]
            )
            mapped["address_line_two"] = address_two_title_case(
                re.sub(
                    re.compile(r"Canada,"),
                    "",
                    re.sub(postal_code_regex, "", city_province_p_code),
                )
            ).strip()

            # Address line 3 (postal code)
            pc_match = re.search(postal_code_regex, city_province_p_code)
            if pc_match:
                mapped["address_line_three"] = pc_match.group().title()

    # --------------- POLICY NUMBER -----------------
    policy_number = safe_get(extracted_data, "policy_number")
    if policy_number:
        mapped["policy_number"] = policy_number.split("\n")[0].strip()

    # --------------- EFFECTIVE DATE -----------------
    effective_date = safe_get(extracted_data, "effective_date")
    if effective_date:
        date_match = re.search(date_regex, effective_date)
        if date_match:
            date_str = date_match.group()
            try:
                # Try parsing various formats
                for fmt in ["%B %d, %Y", "%d %b %Y", "%b %d, %Y"]:
                    try:
                        date_obj = datetime.strptime(date_str, fmt)
                        mapped["effective_date"] = date_obj.strftime(DATE_FORMAT)
                        # Calculate expiry date (1 year later)
                        expiry_date = date_obj.replace(year=date_obj.year + 1)
                        mapped["expiry_date"] = expiry_date.strftime(DATE_FORMAT)
                        break
                    except ValueError:
                        continue
            except Exception as e:
                print(f"Could not parse effective_date: {effective_date}")

    # --------------- PREMIUM AMOUNT -----------------
    premium_amount = safe_get(extracted_data, "premium_amount")
    if premium_amount:
        total = sum_dollar_amounts(premium_amount)
        if total > 0:
            mapped["premium_amount"] = "${:,.2f}".format(total)

    # --------------- RISK ADDRESS -----------------
    risk_address = safe_get(extracted_data, "risk_address")
    if risk_address:
        risk_addresses = []

        # Handle list (from insurers with return_all=True)
        if isinstance(risk_address, list):
            for addr in risk_address:
                formatted_addr = risk_address_title_case(
                    re.sub(postal_code_regex, "", addr).rstrip(", ")
                )
                if formatted_addr and formatted_addr not in risk_addresses:
                    risk_addresses.append(formatted_addr)
        else:
            # Handle string (single address or newline-separated)
            for line in risk_address.split("\n"):
                if re.search(postal_code_regex, line):
                    formatted_addr = risk_address_title_case(
                        re.sub(postal_code_regex, "", line).rstrip(", ")
                    )
                    if formatted_addr and formatted_addr not in risk_addresses:
                        risk_addresses.append(formatted_addr)

        # Assign numbered risk addresses
        for index, addr in enumerate(risk_addresses):
            mapped[f"risk_address_{index + 1}"] = addr

        # Default to risk_address_1 if extracted but no postal code found
        if not risk_addresses and isinstance(risk_address, str):
            mapped["risk_address_1"] = risk_address_title_case(
                re.sub(postal_code_regex, "", risk_address).rstrip(", ")
            )

    # Use mailing address as fallback for risk_address_1
    if not mapped.get("risk_address_1"):
        address_parts = [
            mapped.get(f"address_line_{i}", "").strip()
            for i in ["one", "two", "three"]
            if mapped.get(f"address_line_{i}", "").strip()
        ]
        if address_parts:
            mapped["risk_address_1"] = ", ".join(address_parts)

    # --------------- FORM TYPE -----------------
    form_type = safe_get(extracted_data, "form_type")
    if form_type:
        # Handle list or string
        form_type_str = form_type[0] if isinstance(form_type, list) else form_type
        form_type_lower = form_type_str.lower()

        if "comprehensive" in form_type_lower or "dolce vita" in form_type_lower:
            mapped["form_type_1"] = "Comprehensive Form"
        elif "broad" in form_type_lower:
            mapped["form_type_1"] = "Broad Form"
        elif "basic" in form_type_lower:
            mapped["form_type_1"] = "Basic Form"
        elif (
            "fire & extended" in form_type_lower
            or "fire and extended" in form_type_lower
        ):
            mapped["form_type_1"] = "Fire + EC"
        elif insurer == "Family":
            if "included" in form_type_lower:
                mapped["form_type_1"] = "Comprehensive Form"
            else:
                mapped["form_type_1"] = "Specified Perils"

    # --------------- RISK TYPE -----------------
    risk_type = safe_get(extracted_data, "risk_type")
    if risk_type:
        # Handle list or string
        risk_type_str = risk_type[0] if isinstance(risk_type, list) else risk_type
        risk_type_lower = risk_type_str.lower()

        if "seasonal" in risk_type_lower:
            mapped["seasonal"] = True

        if "home" in risk_type_lower or "dwelling" in risk_type_lower:
            mapped["risk_type_1"] = "home"
        elif (
            "condominium" in risk_type_lower
            and "rented" not in risk_type_lower
            and "rental" not in risk_type_lower
        ):
            mapped["risk_type_1"] = "condo"
            if insurer == "Intact":
                mapped["condo_deductible_1"] = "$100,000"
        elif (
            "rented condominium" in risk_type_lower
            or "rental condominium" in risk_type_lower
        ):
            mapped["risk_type_1"] = "rented_condo"
            if insurer == "Intact":
                mapped["condo_deductible_1"] = "$100,000"
        elif "rented dwelling" in risk_type_lower or "revenue" in risk_type_lower:
            mapped["risk_type_1"] = "rented_dwelling"
        elif "tenant" in risk_type_lower:
            mapped["risk_type_1"] = "tenant"

    # --------------- NUMBER OF FAMILIES -----------------
    number_of_families = safe_get(extracted_data, "number_of_families")
    number_of_units = safe_get(extracted_data, "number_of_units")

    family_keywords = {
        "one": 1,
        "two": 2,
        "three": 3,
        "1": 1,
        "2": 2,
        "3": 3,
        "additional family": 2,
        "002 additional family": 3,
    }

    if number_of_families:
        # Handle list or string
        families_str = (
            number_of_families[0]
            if isinstance(number_of_families, list)
            else number_of_families
        )
        number_str = families_str.lower().strip()
        # Extract number from text
        match = re.search(r"\b(\d+)\b", families_str)
        if insurer == "Family" and match:
            # Family counts rental suites, so add 1
            mapped["number_of_families_1"] = int(match.group(1)) + 1
        else:
            mapped["number_of_families_1"] = family_keywords.get(number_str, 1)
    elif number_of_units:
        # Fallback to number of units for Wawanesa
        units_str = (
            number_of_units[0] if isinstance(number_of_units, list) else number_of_units
        )
        number_str = units_str.lower().strip()
        mapped["number_of_families_1"] = family_keywords.get(number_str, 1)
    else:
        # Default to 1 family
        mapped["number_of_families_1"] = 1

    # --------------- COVERAGES -----------------
    # Earthquake
    earthquake = safe_get(extracted_data, "earthquake_coverage")
    if earthquake:
        if insurer == "Family":
            # Family: check for dollar amount
            if re.search(dollar_regex, earthquake):
                mapped["earthquake_coverage"] = True
        else:
            mapped["earthquake_coverage"] = True

    # Overland Water
    if safe_get(extracted_data, "overland_water"):
        mapped["overland_water"] = True

    # Service Line
    service_line = safe_get(extracted_data, "service_line")
    if service_line:
        # Handle list or string
        if isinstance(service_line, list) or isinstance(service_line, str):
            mapped["service_line"] = True

    # Family: overland water implies service line
    if insurer == "Family" and mapped.get("overland_water"):
        mapped["service_line"] = True

    # Intact: Ground Water
    if insurer == "Intact" and safe_get(extracted_data, "ground_water"):
        mapped["ground_water"] = True

    # Wawanesa: Tenant Vandalism
    tenant_vandalism = safe_get(extracted_data, "tenant_vandalism")
    if insurer == "Wawanesa" and tenant_vandalism:
        # Handle list or string
        if isinstance(tenant_vandalism, list) or isinstance(tenant_vandalism, str):
            mapped["tenant_vandalism"] = True

    # --------------- CONDO DEDUCTIBLES -----------------
    condo_deductible = safe_get(extracted_data, "condo_deductible")
    if condo_deductible:
        deduct_match = re.search(dollar_regex, condo_deductible)
        if deduct_match:
            if insurer == "Family":
                # Family has both regular and earthquake deductible
                mapped["condo_deductible_1"] = deduct_match.group()
                # Check for second amount (earthquake)
                all_amounts = dollar_regex.findall(condo_deductible)
                if len(all_amounts) > 1:
                    mapped["condo_earthquake_deductible_1"] = f"${all_amounts[1]}"
                else:
                    mapped["condo_earthquake_deductible_1"] = deduct_match.group()
            else:
                mapped["condo_deductible_1"] = deduct_match.group()
                if insurer == "Aviva":
                    mapped["condo_earthquake_deductible_1"] = deduct_match.group()

    # Condo Earthquake Deductible (separate field)
    condo_eq_deductible = safe_get(extracted_data, "condo_earthquake_deductible")
    if condo_eq_deductible:
        if insurer == "Intact":
            # Intact: if Additional Loss Assessment exists, use $25,000
            mapped["condo_earthquake_deductible_1"] = "$25,000"
        else:
            deduct_match = re.search(dollar_regex, condo_eq_deductible)
            if deduct_match:
                mapped["condo_earthquake_deductible_1"] = deduct_match.group()
    elif (
        insurer == "Intact"
        and mapped.get("earthquake_coverage")
        and not mapped.get("condo_earthquake_deductible_1")
    ):
        # Intact default earthquake deductible
        mapped["condo_earthquake_deductible_1"] = "$2,500"

    return mapped


# --------------- MAIN PDF PIPELINE -----------------
def process_pdf_for_renewal(pdf_path, rects):
    """
    Detects insurer and extracts fields automatically.
    Returns (insurer_type, extracted_data)
    """
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


# --------------- MAIN EXECUTION -----------------
def main():
    """
    Process all PDFs in Downloads folder, extract data, and generate renewal letters.
    """
    downloads_dir = Path.home() / "Downloads"
    desktop_dir = Path.home() / "Desktop"

    # Find all PDF files in Downloads
    pdf_files = list(downloads_dir.glob("*.pdf"))
    sorted_files = sorted(pdf_files, key=lambda f: f.stat().st_mtime, reverse=True)
    recent_files = sorted_files[:2]

    if not pdf_files:
        print("No PDF files found in Downloads folder.")
        return

    print(f"Found {len(recent_files)} PDF file(s) in Downloads folder.\n")

    for pdf_file in recent_files:
        print(f"Processing: {pdf_file.name}")

        try:
            # Extract data from PDF
            insurer, extracted_data = process_pdf_for_renewal(pdf_file, RECTS)

            if not insurer:
                print(f"  ⚠ Could not detect insurer type. Skipping.\n")
                continue

            print(f"  ✓ Detected insurer: {insurer}")

            # Map data for Word template
            mapped_data = map_extracted_data_for_renewal(extracted_data, insurer)

            # Generate renewal letter
            write_to_new_docx(TEMPLATE_FILE, mapped_data, output_dir=desktop_dir)
            print(f"  ✓ Generated renewal letter for {mapped_data['named_insured']}\n")

        except Exception as e:
            print(f"  ✗ Error processing {pdf_file.name}: {e}\n")

    print("Done!")


if __name__ == "__main__":
    main()
