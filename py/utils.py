import math
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

import openpyxl
from docxtpl import DocxTemplate


DATE_FORMAT = "%B %d, %Y"

POSTAL_CODE_RE = re.compile(
    r"([ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ ]?\d[ABCEGHJ-NPRSTV-Z]\d)$"
)


def safe_strip(value):
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def smart_title(text):
    """
    Title-cases words, but leaves a word untouched if:
      - it contains a digit, or
      - the NEXT word starts with a digit (so labels like "BCS 3746" stay intact)
    Certain small words (e.g. "of") are kept lowercase unless they're the first word.
    """
    if not text:
        return text

    lowercase_words = {"of", "the", "and", "a", "an", "in", "on", "for"}

    words = text.split(" ")
    result = []
    for i, word in enumerate(words):
        has_digit = any(ch.isdigit() for ch in word)
        next_word = words[i + 1] if i + 1 < len(words) else ""
        next_starts_digit = bool(next_word) and next_word[0].isdigit()

        if has_digit or next_starts_digit:
            result.append(word)
        elif word.lower() in lowercase_words and i != 0:
            result.append(word.lower())
        else:
            result.append(word[:1].upper() + word[1:].lower() if word else word)
    return " ".join(result)


def to_float(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace("$", "").replace(",", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def format_amount(value):
    """Strips any existing $ / commas and reformats as a plain number string."""
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return f"{value:,.2f}"
    s = str(value).strip().replace("$", "").replace(",", "")
    if not s:
        return ""
    try:
        return f"{float(s):,.2f}"
    except ValueError:
        return s


def parse_date(value):
    if not value:
        return ""
    if isinstance(value, datetime):
        return value.strftime(DATE_FORMAT)
    value = str(value).strip()
    date_formats = [
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%B %d, %Y",
        "%b %d, %Y",
        "%d-%b-%y",
    ]
    for fmt in date_formats:
        try:
            return datetime.strptime(value, fmt).strftime(DATE_FORMAT)
        except ValueError:
            continue
    print(f"\u26a0\ufe0f Could not parse date: {value}")
    return value


def parse_date_to_dt(value):
    """Same parsing as parse_date, but returns a datetime object (or None)."""
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    value = str(value).strip()
    date_formats = [
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%B %d, %Y",
        "%b %d, %Y",
        "%d-%b-%y",
    ]
    for fmt in date_formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    print(f"\u26a0\ufe0f Could not parse transaction date: {value}")
    return None


def load_producer_mapping(mapping_path):
    wb = openpyxl.load_workbook(mapping_path, data_only=True)
    if "File Completion Tool" not in wb.sheetnames:
        return {}
    ws = wb["File Completion Tool"]
    mapping = {}
    row = 27
    while True:
        code = safe_strip(ws.cell(row=row, column=1).value)
        name = safe_strip(ws.cell(row=row, column=2).value)
        if not code and not name:
            break
        if code:
            mapping[code.lower()] = name
        row += 1
    return mapping


def _split_mailing_address(value):
    """Convert a comma-joined mailing address into newline-separated
    street / city-province / postal lines so the template renders a new
    line after the street address and after the city/province (matching
    the renewal letter). Addresses that are already newline-separated or
    that cannot be reliably split are returned unchanged."""
    text = safe_strip(value)
    if not text or "\n" in text:
        return text

    cleaned = re.sub(r",?\s*Canada\s*$", "", text, flags=re.IGNORECASE).strip()
    match = POSTAL_CODE_RE.search(cleaned)
    if not match:
        return text

    postal = match.group(1).upper()
    before = cleaned[: match.start()].rstrip(" ,")
    if not before:
        return text

    parts = [p.strip() for p in before.split(",") if p.strip()]
    if len(parts) < 2:
        return f"{before}\n{postal}"

    street = parts[0]
    city_province = ", ".join(parts[1:])
    return f"{street}\n{city_province}\n{postal}"


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


# -------------------- Progress Bar -------------------- #
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


# -------------------- File Utilities -------------------- #
def safe_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def unique_file_name(path: str) -> str:
    directory = os.path.dirname(path)
    filename, extension = os.path.splitext(os.path.basename(path))
    filename = safe_filename(filename)

    # Remove existing trailing (n)
    base_name = re.sub(r"\s*\((\d{1,3})\)$", "", filename)

    counter = 1
    new_path = os.path.join(directory, f"{base_name}{extension}")

    while Path(new_path).is_file():
        new_path = os.path.join(directory, f"{base_name} ({counter}){extension}")
        counter += 1

    return new_path


def load_excel_mapping(
    mapping_path, default_mappings, excel_mappings, sheet_name="File Completion Tool"
):

    mapping_path = Path(mapping_path)
    if not mapping_path.exists():
        print(f"Config file not found: {mapping_path.absolute()}")
        print("Please create 'config.xlsx' in the current directory or visit")
        print(
            "https://github.com/WebDevBernard/Python-Automations to download the template."
        )
        print("\nExiting in ", end="")
        for i in range(3, 0, -1):
            print(f"{i} ", end="", flush=True)
            time.sleep(1)
        print()
        raise FileNotFoundError(f"Config file not found: {mapping_path}")

    wb = openpyxl.load_workbook(mapping_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook")

    ws = wb[sheet_name]
    return {key: ws[cell].value for key, cell in excel_mappings.items()}


def _escape_xml_values(data: dict) -> dict:
    """Escape XML special characters (&, <, >, etc.) in all string values."""
    if not data:
        return data
    escaped = {}
    for key, value in data.items():
        if isinstance(value, str):
            escaped[key] = xml_escape(value)
        else:
            escaped[key] = value
    return escaped


def write_to_new_docx(
    template_path: Path | None = None,
    data: dict = None,
    output_dir: Path | None = None,
    output_suffix: str = "Renewal Letter",
    output_year: int | None = None,
) -> bool:
    try:
        # Capture filename-safe value before XML escaping
        named_insured = safe_filename(
            str(data.get("named_insured", "Unnamed Client")).strip()
        )

        data = _escape_xml_values(data)

        # Auto-detect template if not provided
        if template_path is None:
            assets_dir = Path.cwd() / "assets"

            if not assets_dir.exists():
                print(
                    "Assets folder not found. Please create an 'assets' folder with the template."
                )
                print("\nExiting in ", end="")
                for i in range(10, 0, -1):
                    print(f"{i} ", end="", flush=True)
                    time.sleep(1)
                print()
                return False

            docx_files = list(assets_dir.glob("*.docx"))

            if not docx_files:
                print(
                    "Template not found. Visit https://github.com/WebDevBernard/Python-Automations to download the template."
                )
                print("\nExiting in ", end="")
                for i in range(10, 0, -1):
                    print(f"{i} ", end="", flush=True)
                    time.sleep(1)
                print()
                return False

            template_path = docx_files[0]
        else:
            template_path = Path(template_path)
            if not template_path.exists():
                print(
                    "Template not found. Visit https://github.com/WebDevBernard/Python-Automations to download the template."
                )
                print("\nExiting in ", end="")
                for i in range(10, 0, -1):
                    print(f"{i} ", end="", flush=True)
                    time.sleep(1)
                print()
                return False

        doc = DocxTemplate(template_path)
        doc.render(data)

        output_dir = output_dir or (Path.home() / "Desktop")
        if output_year:
            output_filename = (
                output_dir / f"{named_insured} {output_suffix} {output_year}.docx"
            )
        else:
            output_filename = output_dir / f"{named_insured} {output_suffix}.docx"
        doc.save(unique_file_name(output_filename))
        return True

    except Exception as e:
        print(f"Error creating document: {e}")
        print("\nExiting in ", end="")
        for i in range(10, 0, -1):
            print(f"{i} ", end="", flush=True)
            time.sleep(1)
        print()
        return False


def build_index(doc):

    page_index = {}
    text_to_location = []

    for page_num, page in enumerate(doc):
        blocks = []
        for block_idx, block in enumerate(page.get_text("blocks")):
            coords = tuple(block[:4])
            text_lines = block[4].split("\n")
            blocks.append({"words": text_lines, "coords": coords})

            for line_idx, line in enumerate(text_lines):
                text_to_location.append(
                    {
                        "normalized": line.lower(),
                        "page": page_num,
                        "block": block_idx,
                        "line": line_idx,
                        "text": line,
                        "coords": coords,
                    }
                )

        page_index[page_num] = blocks

    return page_index, text_to_location


CONFIG_PATH = Path("../config.xlsx")  # Default path, can be overridden

CONFIG_FIELDS = {
    "task": (2, 1, None),
    "broker_name": (6, 1, None),
    "on_behalf": (8, 1, None),
    "risk_type_1": (12, 1, None),
    "named_insured": (14, 1, None),
    "insurer": (15, 1, None),
    "policy_number": (16, 1, None),
    "effective_date": (17, 1, None),
    "address_line_one": (19, 1, None),
    "address_line_two": (20, 1, None),
    "address_line_three": (21, 1, None),
    "risk_address_1": (23, 1, None),
    "number_of_pdfs": (27, 1, 0),
    "drive_letter": (29, 1, None),
}

PRODUCER_RANGE = (33, 49)
