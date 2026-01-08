import fitz  # PyMuPDF
import re
import time
import openpyxl
import xlrd
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from constants import RECTS, REGEX_PATTERNS
from utils import write_to_new_docx


# Regex patterns
address_regex = REGEX_PATTERNS["address"]
postal_code_regex = REGEX_PATTERNS.get("postal_code")
date_regex = REGEX_PATTERNS.get("date")
dollar_regex = REGEX_PATTERNS.get("dollar")

DATE_FORMAT = "%B %d, %Y"


def map_config_for_renewal(config_data: dict) -> dict:
    mapped = {
        "broker_name": config_data.get("broker_name", "").strip(),
        "on_behalf": config_data.get("on_behalf", "").strip(),
    }
    return mapped


# ----------------- INSURER DETECTION ----------------- #
def detect_insurer(doc, rects):
    """Detect insurer by keyword in policy_type rects."""
    first_page = doc[0]
    for insurer, cfg in rects["policy_type"].items():
        keyword = cfg.get("keyword")
        rect = cfg.get("rect")

        if rect:
            clip_text = first_page.get_text("text", clip=rect)
        else:
            # Scan whole page if no rect specified
            clip_text = first_page.get_text("text")

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
    """Search regex across all blocks in all pages."""
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

    return matches if return_all and matches else (None if not return_all else None)


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
        if insurer == "Intact":
            results[field] = deduplicate_field(results[field])
    return results


def extract_single_field(pages_dict, cfg, doc):
    """Extract a single field based on pattern and/or rect offset."""
    pattern = cfg.get("pattern")
    rect_offset = cfg.get("rect")
    target_rect = cfg.get("target_rect")
    return_all = cfg.get("return_all", False)

    if not pattern and not rect_offset:
        return None

    if not pattern and rect_offset:
        return extract_text_from_absolute_rect(doc, rect_offset)

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


# ----------------- TITLE CASE HELPERS ----------------- #
def address_one_title_case(sentence):
    """Title case with ordinal numbers (1st, 2nd) in lowercase."""
    ordinal_pattern = re.compile(r"\b\d+(st|nd|rd|th)\b")
    return " ".join(
        word.lower() if ordinal_pattern.match(word) else word.capitalize()
        for word in sentence.split()
    )


def address_two_title_case(strings_list):
    """Title case with words longer than 2 characters capitalized, and province codes uppercased."""
    words = strings_list.split()

    # Canadian province codes that should be uppercase
    province_codes = {
        "bc",
        "ab",
        "sk",
        "mb",
        "on",
        "qc",
        "nb",
        "ns",
        "pe",
        "nl",
        "yt",
        "nt",
        "nu",
    }

    capitalized_words = []
    for word in words:
        word_stripped = word.strip()
        # If it's a 2-letter province code, uppercase it
        if len(word_stripped) == 2 and word_stripped.lower() in province_codes:
            capitalized_words.append(word_stripped.upper())
        # Otherwise, capitalize if longer than 2 characters
        elif len(word_stripped) > 2:
            capitalized_words.append(word_stripped.capitalize())
        else:
            capitalized_words.append(word_stripped)

    return " ".join(capitalized_words)


def risk_address_title_case(address):
    """Title case with special handling for state codes and ordinals."""
    parts = address.split()
    if not parts:
        return address

    last_part = parts[-1]
    if len(last_part) == 2:
        last_part = last_part.upper()

    titlecased_parts = []
    for part in parts[:-1]:
        if (
            len(part) > 2
            and part[:-2].isdigit()
            and part[-2:].lower() in ["th", "rd", "nd", "st"]
        ):
            titlecased_parts.append(part.lower())
        else:
            titlecased_parts.append(part.title())

    return " ".join(titlecased_parts) + (" " + last_part if parts else "")


# ----------------- UTILITY HELPERS ----------------- #
def find_index(regex, items):
    """Find first index in list where regex matches."""
    if items and isinstance(items, list):
        for index, string in enumerate(items):
            if re.search(regex, string):
                return index
    return -1


# ----------------- FORMAT FIELDS ----------------- #
def format_fields(raw_data, insurer):
    """Apply insurer-specific formatting to extracted fields."""
    fields = {}

    # Named insured
    if raw_data.get("name_and_address"):
        fields["named_insured"] = format_named_insured(
            raw_data["name_and_address"], insurer
        )

    # Mailing address
    if raw_data.get("name_and_address"):
        address_parts = format_mailing_address(raw_data["name_and_address"])
        fields.update(address_parts)

    # Policy number (direct copy)
    if raw_data.get("policy_number"):
        if insurer == "Intact":
            formatted_policy = format_policy_number(raw_data["policy_number"])
            if formatted_policy:
                fields["policy_number"] = formatted_policy
        else:
            # Keep original for all other insurers
            fields["policy_number"] = raw_data["policy_number"]
    # Effective date (clean up time portion)
    if raw_data.get("effective_date"):
        fields["effective_date"] = format_effective_date(raw_data["effective_date"])

    # Premium amount
    if raw_data.get("premium_amount"):
        fields["premium_amount"] = format_premium_amount(raw_data["premium_amount"])

    # Additional coverage options (boolean flags)
    fields.update(format_additional_coverage(raw_data, insurer))

    # Risk address
    if raw_data.get("risk_address"):
        fields.update(format_risk_addresses(raw_data["risk_address"]))

    # Form type
    if raw_data.get("form_type"):
        fields.update(format_form_types(raw_data["form_type"], insurer))

    # Risk type
    if raw_data.get("risk_type"):
        fields.update(format_risk_types(raw_data, insurer))

    # Number of families
    if raw_data.get("number_of_families") or raw_data.get("number_of_units"):
        fields.update(format_number_of_families(raw_data, insurer))

    # Condo deductible
    if raw_data.get("condo_deductible"):
        fields.update(format_condo_deductibles(raw_data, insurer))

    # Condo earthquake deductible
    if raw_data.get("condo_earthquake_deductible") or (
        insurer == "Intact" and raw_data.get("earthquake_coverage")
    ):
        fields.update(format_condo_earthquake_deductibles(raw_data, insurer))

    return fields


def format_policy_number(policy_number):
    """Format policy number to remove spaces and ensure it contains both letters and numbers."""
    if not policy_number:
        return None

    # Remove all spaces
    cleaned = policy_number.replace(" ", "")

    # Ensure it contains at least one number
    if not any(char.isdigit() for char in cleaned):
        return None

    # Ensure it contains at least one letter
    if not any(char.isalpha() for char in cleaned):
        return None

    return cleaned.upper()  # Optional: uppercase letters for consistency


def format_named_insured(name_and_address_text, insurer):
    """Format named insured from extracted text block, handling multiple lines and cleaning whitespace."""
    if not name_and_address_text:
        return None

    lines = [line.strip() for line in name_and_address_text.split("\n") if line.strip()]
    name_lines = []

    for i, line in enumerate(lines):
        if re.search(address_regex, line):
            break
        name_lines.append(line)

    # Special case: second line might be address
    if len(name_lines) > 1 and re.search(address_regex, name_lines[1]):
        name_lines = name_lines[:1]

    if insurer == "Intact":
        return _format_intact_names(" ".join(name_lines))
    else:
        return _join_names(name_lines)


def _format_intact_names(name_string):
    """Handle Intact-specific name formatting logic."""
    company_suffixes = (
        "Ltd",
        "Ltd.",
        "Inc",
        "Inc.",
        "LLC",
        "Corporation",
        "Corp",
        "Corp.",
    )
    if any(name_string.endswith(suffix) for suffix in company_suffixes):
        return name_string

    name_groups = re.split(r"\s+&\s+|\s+and\s+", name_string, flags=re.IGNORECASE)

    parsed_names = []
    for group in name_groups:
        parts = [p.strip().replace(":", "") for p in group.split(",")]

        if len(parts) == 2:
            parsed_names.append({"first": parts[1], "last": parts[0]})
        elif len(parts) == 1:
            parsed_names.append({"first": parts[0], "last": None})

    # Shared last name pattern
    if (
        len(parsed_names) > 1
        and parsed_names[0]["last"]
        and not parsed_names[1]["last"]
    ):
        shared_last = parsed_names[0]["last"]
        first_names = [name["first"] for name in parsed_names]
        return f"{' and '.join(first_names)} {shared_last}".title()

    # Different last names or single name
    formatted = []
    for name in parsed_names:
        if name["last"]:
            formatted.append(f"{name['first']} {name['last']}")
        else:
            formatted.append(name["first"])

    return _join_names(formatted)


def _join_names(names):
    """
    Join a list of names with proper formatting and lowercase 'and'.
    Cleans extra spaces and strips colons.
    """
    # Strip spaces, remove empty strings, and replace multiple spaces with single
    clean_names = [
        re.sub(r"\s+", " ", n).strip().replace(":", "") for n in names if n.strip()
    ]

    if not clean_names:
        return ""
    elif len(clean_names) == 1:
        return clean_names[0].title()
    elif len(clean_names) == 2:
        return f"{clean_names[0].title()} and {clean_names[1].title()}"
    else:
        return f"{', '.join(n.title() for n in clean_names[:-1])} and {clean_names[-1].title()}"


def format_effective_date(effective_date_text):
    """Extract and normalize the first date in the text to 'Month DD, YYYY'."""

    if not effective_date_text:
        return None

    # Pattern for formats like '26 Jan 2026'
    dmy_pattern = r"(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})"

    # Pattern for 'November 12, 2025'
    long_month_pattern = r"([A-Z][a-z]+\s+\d{1,2},\s+\d{4})"

    # Try long month first
    match = re.search(long_month_pattern, effective_date_text)
    if match:
        return match.group(1)

    # Try DMY (26 Jan 2026)
    match = re.search(dmy_pattern, effective_date_text)
    if match:
        raw = match.group(1)
        # Parse and convert to "Month DD, YYYY"
        try:
            dt = datetime.strptime(raw, "%d %b %Y")
            return dt.strftime("%B %d, %Y")
        except ValueError:
            pass

    return effective_date_text


def format_mailing_address(name_and_address_text):
    """Format mailing address into three lines from extracted text block."""
    if not name_and_address_text:
        return {}

    lines = [line.strip() for line in name_and_address_text.split("\n") if line.strip()]

    # Find where address starts
    address_index = -1
    for i, line in enumerate(lines):
        if re.search(address_regex, line):
            address_index = i
            break

    if address_index == -1:
        return {}

    # Get address lines (everything from address_index onward)
    address_lines = lines[address_index:]

    # Find postal code line
    pc_index = -1
    for i, line in enumerate(address_lines):
        if re.search(postal_code_regex, line):
            pc_index = i
            break

    if pc_index == -1:
        return {}

    result = {}

    # Address line one: ONLY the first line (street address)
    result["address_line_one"] = address_one_title_case(address_lines[0])

    # Address line two: city/province (from line AFTER street address to postal code line)
    city_province_pcode = " ".join(address_lines[1 : pc_index + 1])

    # Remove postal code and "Canada"
    cleaned = re.sub(postal_code_regex, "", city_province_pcode)
    cleaned = re.sub(r"Canada,?", "", cleaned, flags=re.IGNORECASE).strip()

    result["address_line_two"] = address_two_title_case(cleaned)

    # Address line three: postal code
    postal_match = re.search(postal_code_regex, address_lines[pc_index])
    if postal_match:
        result["address_line_three"] = postal_match.group().upper()

    return result


def format_premium_amount(premium_text):
    """Ensure premium amount has $ and two decimal places."""
    if not premium_text:
        return None

    # Remove any existing $ and commas
    cleaned = re.sub(r"[^\d.]", "", premium_text)

    try:
        amount = float(cleaned)
        # Always format as currency
        return f"${amount:,.2f}"
    except ValueError:
        # If conversion fails, return original text
        return premium_text


def format_additional_coverage(raw_data, insurer):
    """Set boolean flags for additional coverage options."""
    fields = {}

    # Earthquake coverage
    if raw_data.get("earthquake_coverage"):
        fields["earthquake_coverage"] = True

    # Ground water (Intact only)
    if insurer == "Intact" and raw_data.get("ground_water"):
        fields["ground_water"] = True

    # Tenant vandalism (Wawanesa only)
    if insurer == "Wawanesa" and raw_data.get("tenant_vandalism"):
        fields["tenant_vandalism"] = True

    # Overland water
    if raw_data.get("overland_water"):
        fields["overland_water"] = True

    # Service line
    if raw_data.get("service_line"):
        fields["service_line"] = True

    if insurer == "Wawanesa" and raw_data.get("sewer_back_up_increased_deductible"):
        fields["increased_sbu"] = raw_data.get("sewer_back_up_increased_deductible")

    if insurer == "Wawanesa" and raw_data.get("overland_water_increased_deductible"):
        fields["increased_ow"] = raw_data.get("overland_water_increased_deductible")

    return fields


def format_risk_addresses(risk_address_data):
    """Format risk addresses with indexing."""
    fields = {}

    # Handle both string and list
    if isinstance(risk_address_data, str):
        risk_addresses = [risk_address_data]
    elif isinstance(risk_address_data, list):
        risk_addresses = risk_address_data
    else:
        return fields

    # Only process addresses with postal codes
    valid_addresses = []
    for addr in risk_addresses:
        addr_stripped = addr.strip()  # ← STRIP FIRST
        if re.search(postal_code_regex, addr_stripped):
            valid_addresses.append(addr_stripped)

    # Format each address
    for index, risk_address in enumerate(valid_addresses, start=1):
        cleaned = re.sub(postal_code_regex, "", risk_address).rstrip(", ")
        fields[f"risk_address_{index}"] = risk_address_title_case(cleaned)

    return fields


def format_form_types(form_type_data, insurer):
    """Classify and format insurance form types."""
    fields = {}

    # Handle both string and list
    if isinstance(form_type_data, str):
        form_types = [form_type_data]
    elif isinstance(form_type_data, list):
        form_types = form_type_data
    else:
        return fields

    form_type_mapping = {
        "comprehensive": "Comprehensive Form",
        "broad": "Broad Form",
        "basic": "Basic Form",
        "fire & extended": "Fire + EC",
        "dolce vita": "Comprehensive Form",
    }

    for index, form_type in enumerate(form_types, start=1):
        form_lower = form_type.lower()

        # Check against mapping
        for key, value in form_type_mapping.items():
            if key in form_lower:
                fields[f"form_type_{index}"] = value
                break

        # Family-specific logic
        if insurer == "Family":
            if "included" in form_lower:
                fields[f"form_type_{index}"] = "Comprehensive Form"
            elif f"form_type_{index}" not in fields:
                fields[f"form_type_{index}"] = "Specified Perils"

    return fields


def format_risk_types(raw_data, insurer):
    """Format risk_type and create indexed fields."""
    fields = {}

    risk_type_data = raw_data.get("risk_type")
    form_type_data = raw_data.get("form_type")

    # Handle both string and list
    if isinstance(risk_type_data, str):
        risk_types = [risk_type_data]
    elif isinstance(risk_type_data, list):
        risk_types = risk_type_data
    else:
        risk_types = []

    if isinstance(form_type_data, str):
        form_types = [form_type_data]
    elif isinstance(form_type_data, list):
        form_types = form_type_data
    else:
        form_types = []

    # Ensure both lists are same length
    max_len = max(len(risk_types), len(form_types))
    risk_types.extend([None] * (max_len - len(risk_types)))
    form_types.extend([None] * (max_len - len(form_types)))

    for index, (risk_type, form_type) in enumerate(
        zip(risk_types, form_types), start=1
    ):
        combined = " ".join(filter(None, [risk_type or "", form_type or ""]))
        combined_lower = combined.lower()

        if not combined.strip():
            continue

        if "seasonal" in combined_lower:
            fields["seasonal"] = True

        if "home" in combined_lower:
            fields[f"risk_type_{index}"] = "home"
        elif insurer == "Aviva" and "condominium" in combined_lower:
            fields[f"risk_type_{index}"] = "condo"
        elif insurer == "Family" and "condominium" in combined_lower:
            fields[f"risk_type_{index}"] = "condo"
        elif insurer == "Intact" and "rented condominium" in combined_lower:
            fields[f"risk_type_{index}"] = "rented_condo"
            fields[f"condo_deductible_{index}"] = "$100,000"
        elif insurer == "Intact" and "condominium" in combined_lower:
            fields[f"risk_type_{index}"] = "condo"
            fields[f"condo_deductible_{index}"] = "$100,000"
        elif insurer == "Wawanesa" and "Rental Condominium" in combined:
            fields[f"risk_type_{index}"] = "rented_condo"
        elif insurer == "Wawanesa" and "Condominium" in combined:
            fields[f"risk_type_{index}"] = "condo"
        elif "rented dwelling" in combined_lower:
            fields[f"risk_type_{index}"] = "rented_dwelling"
        elif "revenue" in combined_lower:
            fields[f"risk_type_{index}"] = "rented_dwelling"
        elif "rental" in combined_lower:
            fields[f"risk_type_{index}"] = "rented_condo"
        elif "tenant" in combined_lower:
            fields[f"risk_type_{index}"] = "tenant"

    return fields


def format_number_of_families(raw_data, insurer):
    """Format number of families field, only allowing 1, 2, or 3."""
    fields = {}

    families_data = raw_data.get("number_of_families")
    units_data = raw_data.get("number_of_units")

    # Function to clamp values to 1, 2, or 3
    def sanitize_number(num):
        try:
            num = int(num)
        except (ValueError, TypeError):
            return 1  # Default to 1 if invalid
        return min(max(num, 1), 3)  # Clamp to 1, 2, or 3

    # Default to 1 family for certain insurers if not specified
    if insurer in ["Wawanesa", "Intact", "Family", "Aviva"]:
        if not families_data and not units_data:
            fields["number_of_families_1"] = 1
            return fields

    # Handle families data
    if families_data:
        if isinstance(families_data, str):
            families = [families_data]
        elif isinstance(families_data, list):
            families = families_data
        else:
            families = []

        for index, count in enumerate(families, start=1):
            num = sanitize_number(count)
            # For Family and Aviva insurers, add 1 but still clamp to 3 max
            if insurer in ["Family", "Aviva"]:
                num = min(num + 1, 3)
            fields[f"number_of_families_{index}"] = num

    # Fallback to units if families not available
    elif units_data:
        if isinstance(units_data, str):
            units = [units_data]
        elif isinstance(units_data, list):
            units = units_data
        else:
            units = []

        for index, unit_count in enumerate(units, start=1):
            fields[f"number_of_families_{index}"] = sanitize_number(unit_count)
    return fields


def format_condo_deductibles(raw_data, insurer):
    """Format condo deductible amounts."""
    fields = {}

    deductible_data = raw_data.get("condo_deductible")
    if not deductible_data:
        return fields

    # Handle both string and list
    if isinstance(deductible_data, str):
        deductibles = [deductible_data]
    elif isinstance(deductible_data, list):
        deductibles = deductible_data
    else:
        return fields

    for index, deductible in enumerate(deductibles, start=1):
        # Intact always uses $100,000 for condo deductibles
        if insurer == "Intact":
            fields[f"condo_deductible_{index}"] = "$100,000"
        else:
            # Extract dollar amount for other insurers
            match = re.search(dollar_regex, deductible)
            if match:
                fields[f"condo_deductible_{index}"] = match.group()

                # Aviva-specific: also set earthquake deductible
                if insurer == "Aviva":
                    fields["condo_earthquake_deductible_1"] = match.group()

    return fields


def format_condo_earthquake_deductibles(raw_data, insurer):
    """Format condo earthquake deductible amounts."""
    fields = {}

    # Intact-specific: default earthquake deductible
    if insurer == "Intact":
        if raw_data.get("earthquake_coverage") and not raw_data.get(
            "condo_earthquake_deductible"
        ):
            fields["condo_earthquake_deductible_1"] = "$2,500"

    deductible_data = raw_data.get("condo_earthquake_deductible")
    if not deductible_data:
        return fields

    # Handle both string and list
    if isinstance(deductible_data, str):
        deductibles = [deductible_data]
    elif isinstance(deductible_data, list):
        deductibles = deductible_data
    else:
        return fields

    for index, deductible in enumerate(deductibles, start=1):
        if insurer == "Intact" and raw_data.get("condo_earthquake_deductible"):
            fields["condo_earthquake_deductible_1"] = "$25,000"
        else:
            # Extract dollar amount
            match = re.search(dollar_regex, deductible)
            if match:
                fields[f"condo_earthquake_deductible_{index}"] = match.group()

    return fields


# Add this helper function
def format_postal_code(postal):
    """Format postal code to standard format."""
    if postal is None or postal == "":
        return None
    postal_str = str(postal).replace(" ", "").upper()
    if len(postal_str) == 6:
        return f"{postal_str[:3]} {postal_str[3:]}"
    return postal_str


def get_month_day(date_str):
    """Extract month and day from date string."""
    if date_str is None or not date_str:
        return None
    try:
        # Handle different date formats
        if isinstance(date_str, str):
            # Try parsing "November 12, 2025" format
            parsed = datetime.strptime(date_str, "%B %d, %Y")
            return parsed.strftime("%m-%d")
    except:
        pass
    return None


def currency_to_float(currency_str):
    """Convert currency string to float."""
    if not currency_str:
        return 0.0
    return float(currency_str.replace("$", "").replace(",", ""))


# Add this function to check for matching glass policy
def get_glass_policies():
    """Load glass policy data from Excel files in assets folder."""
    assets_folder = Path.cwd() / "assets"

    # Check if assets folder exists
    if not assets_folder.exists():
        return None

    try:
        xlsx_files = list(assets_folder.rglob("*.xlsx"))
        xls_files = list(assets_folder.rglob("*.xls"))
    except Exception as e:
        print(f"Warning: Error accessing assets folder: {e}")
        return None

    all_records = []
    files = xlsx_files + xls_files

    # Check for old files
    old_files = []
    one_year_ago = datetime.now() - timedelta(days=365)

    for file in files:
        try:
            # Check file modification time
            try:
                file_mtime = datetime.fromtimestamp(file.stat().st_mtime)
                if file_mtime < one_year_ago:
                    old_files.append((file.name, file_mtime))
            except (OSError, FileNotFoundError):
                # File might have been deleted or is inaccessible
                continue

            # Read Excel file using appropriate library
            if file.suffix.lower() == ".xls":
                # Read .xls files using xlrd
                workbook = xlrd.open_workbook(file)
                sheet = workbook.sheet_by_index(0)

                # Get headers from first row
                headers = [sheet.cell_value(0, col) for col in range(sheet.ncols)]

                # Read data rows
                for row_idx in range(1, sheet.nrows):
                    row_data = {}
                    for col_idx in range(sheet.ncols):
                        row_data[headers[col_idx]] = sheet.cell_value(row_idx, col_idx)
                    all_records.append(row_data)
            else:
                # Read .xlsx files using openpyxl
                workbook = openpyxl.load_workbook(file, data_only=True)
                sheet = workbook.active

                # Get headers from first row
                headers = [cell.value for cell in sheet[1]]

                # Read data rows
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    row_data = {}
                    for col_idx, value in enumerate(row):
                        if col_idx < len(headers):
                            row_data[headers[col_idx]] = value
                    all_records.append(row_data)

        except Exception as e:
            print(f"Warning: Error reading {file.name}: {e}")
            continue

    if not all_records:
        return None

    # Format postal codes
    for record in all_records:
        if "h_postzip" in record:
            record["postal_code"] = format_postal_code(record.get("h_postzip"))

    # Remove duplicates by policynum (keep=False means remove ALL duplicates)
    if all_records and "policynum" in all_records[0]:
        # Count occurrences of each policy number
        policy_counts = {}
        for record in all_records:
            policynum = record.get("policynum")
            if policynum:
                policy_counts[policynum] = policy_counts.get(policynum, 0) + 1

        # Keep only records with unique policy numbers
        all_records = [
            record
            for record in all_records
            if policy_counts.get(record.get("policynum"), 0) == 1
        ]

    # Print warning for old files
    if old_files:
        print("\n" + "!" * 80)
        print(
            "⚠️  WARNING: Glass policy Excel file in /assets older than 1 year detected:"
        )
        for filename, mod_time in old_files:
            age_days = (datetime.now() - mod_time).days
            print(
                f"   - {filename} (Last modified: {mod_time.strftime('%B %d, %Y')} - {age_days} days old)"
            )
        print(
            "   Please replace Excel file with new Reliance Glass renewal list to ensure accuracy."
        )
        print("!" * 80 + "\n")

    return all_records


# Add this function to check for matching glass policy
def check_glass_policy(fields, glass_policies):
    """Check if there's a matching glass policy and update fields accordingly."""
    if glass_policies is None or not glass_policies:
        return fields

    # Only check for home policies
    if fields.get("risk_type_1") != "home":
        return fields

    mailing_postal = fields.get("address_line_three")
    expiry_date = fields.get("effective_date")  # Assuming this is the renewal date

    if not mailing_postal or not expiry_date:
        return fields

    # Check for matches
    for glass_row in glass_policies:
        # Check if insurer is REL
        if glass_row.get("insurer") != "REL":
            continue

        # Check postal code match
        if glass_row.get("postal_code") != mailing_postal:
            continue

        # Check renewal date match (month and day)
        glass_renewal = glass_row.get("renewal")
        if get_month_day(glass_renewal) == get_month_day(expiry_date):
            # Match found!
            fields["glass_policynum"] = glass_row.get("policynum")

            # Combine premium amounts
            current_premium = currency_to_float(fields.get("premium_amount", "$0.00"))
            glass_premium = float(glass_row.get("prem_amt", 0) or 0)
            total_premium = current_premium + glass_premium
            fields["premium_amount"] = f"${total_premium:,.2f}"

            break  # Stop after first match

    return fields


# ----------------- UTILITY ----------------- #
def print_fields(fields):
    """Pretty print extracted fields."""
    for key, value in fields.items():
        if isinstance(value, list):
            print(f"{key:30s}: {len(value)} items")
            for i, item in enumerate(value, 1):
                print(f"  [{i}] {item}")
        else:
            print(f"{key:30s}: {value}")


# ----------------- MAIN ----------------- #
from datetime import datetime
from pathlib import Path


# ----------------- MAIN ----------------- #
def auto_renewal_letter(config=None):
    downloads_dir = Path.home() / "Downloads"
    pdf_files = list(downloads_dir.glob("*.pdf"))
    sorted_files = sorted(pdf_files, key=lambda f: f.stat().st_mtime, reverse=True)
    recent_files = sorted_files[:2]

    if not pdf_files:
        print("No PDF files found in Downloads folder.")
        print("\nExiting in ", end="")
        for i in range(3, 0, -1):
            print(f"{i} ", end="", flush=True)
            time.sleep(1)
        print()
        return

    # Get broker info from config if provided
    broker_config = (
        map_config_for_renewal(config)
        if config
        else {"broker_name": "", "on_behalf": ""}
    )

    # Load glass policies once at the start
    df_glass = get_glass_policies()

    for pdf_file in recent_files:
        try:
            doc = fitz.open(pdf_file)
        except Exception as e:
            print(f"  ✗ Error opening {pdf_file.name}: {e}")
            continue

        try:
            with doc:
                insurer = detect_insurer(doc, RECTS)
                if insurer:
                    # 1. Extract raw data
                    raw_data = extract_fields(doc, RECTS[insurer], insurer=insurer)

                    # 2. Check if Wawanesa statement (skip if so)
                    if insurer == "Wawanesa" and raw_data.get("wawanesa_statement"):
                        print(f"  ⊘ Skipping: Wawanesa Statement")
                        continue

                    # 3. Format into final fields
                    fields = format_fields(raw_data, insurer)

                    # 4. Add broker info from config
                    fields["broker_name"] = broker_config["broker_name"]
                    fields["on_behalf"] = broker_config["on_behalf"]

                    # 5. Check for glass policy match
                    fields = check_glass_policy(fields, df_glass)

                    # 6. Add other required fields for docx template
                    fields["today"] = datetime.today().strftime(DATE_FORMAT)
                    fields["insurer"] = insurer

                    # 7. Print fields for debugging
                    # print_fields(fields)

                    # 8. Write to docx
                    write_to_new_docx(data=fields)

                else:
                    print(f"  ✗ No insurer detected")
        except Exception as e:
            print(f"  ✗ Error processing {pdf_file.name}: {e}")
            traceback.print_exc()

    print("******** Auto Renewal Letter ran successfully ********")


if __name__ == "__main__":
    auto_renewal_letter()
