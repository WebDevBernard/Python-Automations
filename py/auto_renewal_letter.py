import fitz  # PyMuPDF
import re
from pathlib import Path
from datetime import datetime
from utils import write_to_new_docx
from constants import RECTS

DATE_FORMAT = "%B %d, %Y"
TEMPLATE_FILE = Path.cwd() / "assets" / "Renewal Letter.docx"


class PDFExtractor:
    def __init__(self, pdf_path, rects):
        """
        pdf_path : path to PDF
        rects    : full RECTS dict (policy_type + insurer mappings)
        """
        self.pdf_path = pdf_path
        self.rects = rects
        self.doc = fitz.open(pdf_path)
        self.policy_type = None
        self.mapping = None

    # --------------- AUTO-DETECT INSURER -----------------
    def detect_policy_type(self):
        """
        Detects insurer type using RECTS['policy_type'] and fitz.Rect clipping.
        """
        first_page = self.doc[0]

        for insurer, cfg in self.rects["policy_type"].items():
            keyword = cfg.get("keyword")
            rect = cfg.get("rect")  # Already fitz.Rect

            # 1️⃣ Check inside the rectangle
            if rect:
                clip_text = first_page.get_text("text", clip=rect)
                if clip_text and keyword.lower() in clip_text.lower():
                    self.policy_type = insurer
                    self.mapping = self.rects.get(insurer)
                    return insurer

            # 2️⃣ Fallback: full-page search
            if keyword and first_page.search_for(keyword):
                self.policy_type = insurer
                self.mapping = self.rects.get(insurer)
                return insurer

        return None

    # --------------- TEXT EXTRACTION -----------------
    def extract_text_from_rect(self, page, rect):
        """Extract text within a fitz.Rect rectangle."""
        text = page.get_text("text", clip=rect)
        return text.strip() if text else None

    def find_by_keyword(self, page, pattern, rect=None):
        """
        Finds a regex pattern, optionally extracts text from a clipped rectangle.
        """
        if not pattern:
            return None

        matches = page.search_for(pattern.pattern)
        if not matches:
            return None

        keyword_rect = matches[0]
        if rect:  # Rect offset relative to found keyword
            offset_rect = fitz.Rect(
                keyword_rect.x0 + rect.x0,
                keyword_rect.y0 + rect.y0,
                keyword_rect.x1 + rect.x1,
                keyword_rect.y1 + rect.y1,
            )
            return self.extract_text_from_rect(page, offset_rect)

        # Otherwise return text in the keyword rectangle
        return page.get_textbox(keyword_rect)

    # --------------- MAIN EXTRACTION -----------------
    def extract_fields(self):
        """
        Extract all fields for the detected insurer type.
        Returns dict: { field_name: extracted_text }
        """
        if not self.mapping:
            raise ValueError(
                "Policy type not detected. Run detect_policy_type() first."
            )

        results = {}
        for field, cfg in self.mapping.items():
            results[field] = None
            for page in self.doc:
                pattern = cfg.get("pattern")
                rect = cfg.get("rect")

                if pattern and rect:
                    text = self.find_by_keyword(page, pattern, rect)
                elif pattern:
                    text = self.find_by_keyword(page, pattern)
                elif rect:
                    text = self.extract_text_from_rect(page, rect)
                else:
                    text = None

                if text:
                    results[field] = text
                    break  # Stop at first occurrence

        return results

    # --------------- CLEANUP -----------------
    def close(self):
        self.doc.close()

    # --------------- CONVENIENCE ENTRY -----------------
    @classmethod
    def auto_from_pdf(cls, pdf_path, rects):
        """
        Detects policy type and extracts fields automatically.
        Returns (insurer_type, extracted_data)
        """
        extractor = cls(pdf_path, rects)
        insurer = extractor.detect_policy_type()
        if not insurer:
            extractor.close()
            return None, {}

        data = extractor.extract_fields()
        extractor.close()
        return insurer, data


# --------------- DATA MAPPING -----------------
def safe_get(data: dict, key: str) -> str:
    """Safely get a value from dict, handling None values."""
    value = data.get(key)
    return value.strip() if value else ""


def map_extracted_data_for_renewal(extracted_data: dict, insurer: str) -> dict:
    """
    Maps extracted PDF data to Word template format.
    """
    # Parse name and address
    name_and_address = safe_get(extracted_data, "name_and_address")
    address_lines = [line.strip() for line in name_and_address.split("\n") if line.strip()]

    insured_name = address_lines[0] if address_lines else ""
    mailing_street = address_lines[1] if len(address_lines) > 1 else ""
    city_province = address_lines[2] if len(address_lines) > 2 else ""
    mailing_postal = address_lines[3] if len(address_lines) > 3 else ""

    mapped = {
        "task": "Auto Renewal Letter",
        "broker_name": "",  # Not extracted from PDF
        "on_behalf": "",  # Not extracted from PDF
        "risk_type_1": safe_get(extracted_data, "risk_type"),
        "named_insured": insured_name,
        "insurer": insurer,
        "policy_number": safe_get(extracted_data, "policy_number"),
        "effective_date": safe_get(extracted_data, "effective_date"),
        "address_line_one": mailing_street,
        "address_line_two": city_province,
        "address_line_three": mailing_postal,
        "risk_address_1": safe_get(extracted_data, "risk_address"),
        "today": datetime.today().strftime(DATE_FORMAT),
    }

    # Build mailing address
    address_fields = ["address_line_one", "address_line_two", "address_line_three"]
    address_parts = [
        mapped.get(field, "").strip()
        for field in address_fields
        if mapped.get(field, "").strip()
    ]
    mapped["mailing_address"] = ", ".join(address_parts)

    # Use mailing address if risk_address_1 is empty
    if not mapped.get("risk_address_1", "").strip():
        mapped["risk_address_1"] = mapped["mailing_address"]

    # Format effective_date
    effective_date = mapped.get("effective_date")
    if effective_date:
        try:
            if isinstance(effective_date, datetime):
                mapped["effective_date"] = effective_date.strftime(DATE_FORMAT)
            else:
                # Try parsing common string formats
                date_obj = datetime.strptime(str(effective_date), "%Y-%m-%d")
                mapped["effective_date"] = date_obj.strftime(DATE_FORMAT)
        except ValueError:
            try:
                date_obj = datetime.strptime(str(effective_date), "%m/%d/%Y")
                mapped["effective_date"] = date_obj.strftime(DATE_FORMAT)
            except ValueError:
                # Keep original if parsing fails
                print(f"Could not parse effective_date: {effective_date}")

    return mapped


# --------------- MAIN EXECUTION -----------------
def main():
    """
    Process all PDFs in Downloads folder, extract data, and generate renewal letters.
    """
    downloads_dir = Path.home() / "Downloads"
    desktop_dir = Path.home() / "Desktop"

    # Find all PDF files in Downloads
    pdf_files = list(downloads_dir.glob("*.pdf"))

    if not pdf_files:
        print("No PDF files found in Downloads folder.")
        return

    print(f"Found {len(pdf_files)} PDF file(s) in Downloads folder.\n")

    for pdf_file in pdf_files:
        print(f"Processing: {pdf_file.name}")

        try:
            # Extract data from PDF
            insurer, extracted_data = PDFExtractor.auto_from_pdf(pdf_file, RECTS)

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
