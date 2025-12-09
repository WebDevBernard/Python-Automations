import re
import fitz

# --------------- REGEX PATTERNS -----------------
REGEX_PATTERNS = {
    "postal_code": re.compile(
        r"([ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ ]?\d[ABCEGHJ-NPRSTV-Z]\d)$"
    ),
    "dollar": re.compile(r"\$([\d,]+)"),
    "date": re.compile(
        r"\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}|(Jan(uary)?|Feb(ruary)?"
        r"|Mar(ch)?|Apr(il)?|May|Jun(e)?|Jul(y)?|Aug(ust)?|Sep(tember)?|Oct(ober)?|Nov(ember)?"
        r"|Dec(ember)?)\s+\d{1,2},\s+\d{4}"
    ),
    "address": re.compile(
        r"(?!.*\bltd\.)((po box)|(unit)|\d+\s+)", flags=re.IGNORECASE
    ),
    "and": re.compile(r"&|\b(and)\b", flags=re.IGNORECASE),
}

DEFAULT_MAPPING = {
    "event": None,
    "broker_name": None,
    "on_behalf": None,
    "risk_type": None,
    "insured_name": None,
    "insurer": None,
    "policy_number": None,
    "effective_date": None,
    "mailing_street": None,
    "city_province": None,
    "mailing_postal": None,
    "risk_address": None,
}

EXCEL_CELL_MAPPING = {
    "event": "B3",
    "broker_name": "B7",
    "on_behalf": "B9",
    "risk_type": "B13",
    "insured_name": "B15",
    "insurer": "B16",
    "policy_number": "B17",
    "effective_date": "B18",
    "mailing_street": "B20",
    "city_province": "B21",
    "mailing_postal": "B22",
    "risk_address": "B24",
}


# --------------- HELPER FUNCTIONS -----------------
def rect(x0, y0, x1, y1):
    """Create a fitz.Rect with more readable syntax."""
    return fitz.Rect(x0, y0, x1, y1)


def offset(dx0, dy0, dx1, dy1):
    """Create an offset rect (for pattern-based extraction)."""
    return fitz.Rect(dx0, dy0, dx1, dy1)


def pattern(regex_str, flags=re.IGNORECASE):
    """Create a compiled regex pattern."""
    return re.compile(regex_str, flags)


# Field configuration helpers
def absolute_rect_field(x0, y0, x1, y1):
    """Field extracted from absolute coordinates (no pattern)."""
    return {
        "pattern": None,
        "rect": rect(x0, y0, x1, y1),
    }


def pattern_only_field(regex_str, flags=re.IGNORECASE, return_all=False):
    """Field extracted by pattern matching only."""
    return {
        "pattern": pattern(regex_str, flags),
        "rect": None,
        "return_all": return_all,
    }


def pattern_with_offset_field(
    regex_str, dx0, dy0, dx1, dy1, flags=re.IGNORECASE, return_all=False
):
    """Field extracted by pattern + relative offset."""
    return {
        "pattern": pattern(regex_str, flags),
        "rect": offset(dx0, dy0, dx1, dy1),
        "return_all": return_all,
    }


# --------------- INSURER DETECTION -----------------
POLICY_TYPE_DETECTION = {
    "Aviva": {
        "keyword": "Aviva",
        "rect": rect(
            183.83999633789062,
            712.8900146484375,
            197.9759979248047,
            734.4000244140625,
        ),
    },
    "Family": {
        "keyword": "Agent",
        "rect": rect(25.70, 36.37, 51.04, 45.45),
    },
    "Intact": {
        "keyword": "Intact Insurance",
        "rect": None,
    },
    "Wawanesa": {
        "keyword": "BROKER OFFICE",
        "rect": rect(36.0, 102.43, 353.27, 111.37),
    },
}


# --------------- FIELD MAPPINGS BY INSURER -----------------

AVIVA_FIELDS = {
    "name_and_address": absolute_rect_field(80.4, 202.24, 250, 280),
    "policy_number": pattern_with_offset_field(
        r"Policy Number", dx0=267.12, dy0=10.16, dx1=-202.82, dy1=9.16
    ),
    "effective_date": pattern_only_field(
        r"Policy Effective From:\s*([A-Z][a-z]+\s+\d{1,2},\s+\d{4})"
    ),
    "risk_address": pattern_only_field(
        r"Location [123]\s+(?!deductible|discounts)(.*)", return_all=True
    ),  # Can have multiple
    "form_type": pattern_with_offset_field(
        r"Location [123]\s+(?!deductible|discounts)(.*)",
        dx0=245.52,
        dy0=0.80,
        dx1=350.89,
        dy1=-10.00,
        return_all=True,
    ),  # Can have multiple
    "risk_type": pattern_with_offset_field(
        r"Location [123]\s+(?!deductible|discounts)(.*)",
        dx0=245.52,
        dy0=0.80,
        dx1=350.89,
        dy1=-10.00,
        return_all=True,
    ),  # Can have multiple
    "number_of_families": pattern_only_field(r"00([12])\s+Additional Family"),
    "earthquake_coverage": pattern_only_field(
        r"Earthquake (?:- \d+(?:\.\d+)?% Of Personal Property - |Endorsement )(\d+(?:\.\d+)?%)"
    ),
    "overland_water": pattern_only_field(
        r"Overland Water - Deductible (\$[\d,]+(?:\.\d{2})?)"
    ),
    "condo_deductible": pattern_only_field(
        r"Condominium Corporation Deductible - (\$[\d,]+(?:\.\d{2})?)"
    ),
    "service_line": pattern_only_field(
        r"Service Line Coverage Endorsement - (\$[\d,]+(?:\.\d{2})?) Limit"
    ),
    "premium_amount": pattern_only_field(
        r"Total Policy Premium.*?(\$[\d,]+(?:\.\d{2})?)"
    ),
}


FAMILY_FIELDS = {
    "name_and_address": absolute_rect_field(25.34, 153.38, 150, 228.67),
    "policy_number": pattern_with_offset_field(
        r"POLICY NUMBER", dx0=-0.94, dy0=11.03, dx1=-5.61, dy1=10.80
    ),
    "effective_date": pattern_with_offset_field(
        r"EFFECTIVE DATE", dx0=-1.01, dy0=20.17, dx1=24.57, dy1=11.45
    ),
    "risk_address": pattern_only_field(
        r"(?i)LOCATION\s+OF\s+INSURED\s+PROPERTY:\s*(.+)"
    ),
    "form_type": pattern_only_field(r"(?i)All\s+Perils:\s*(Included)"),
    "risk_type": pattern_with_offset_field(
        r"POLICY TYPE", dx0=-0.94, dy0=11.10, dx1=7.76, dy1=11.45
    ),
    "number_of_families": pattern_only_field(
        r"(?i)OPERATION\W+OF\W+([12])\W*RENTAL\W*SUITES?"
    ),
    "earthquake_coverage": pattern_with_offset_field(
        r"EARTHQUAKE PROPERTY LIMITS", dx0=113.5, dy0=12.74, dx1=42, dy1=12.37
    ),
    "overland_water": pattern_only_field(r"Overland Water"),
    "condo_deductible": pattern_only_field(
        r"(?i)Deductible\W+Coverage\W*:\W*(\$[\d,]+)\*?"
    ),
    "service_line": pattern_only_field(r"Service Lines"),
    "premium_amount": pattern_with_offset_field(
        r"RETURN THIS PORTION WITH PAYMENT",
        dx0=5.59,
        dy0=-22.79,
        dx1=-116.08,
        dy1=-22.20,
    ),
}


INTACT_FIELDS = {
    "name_and_address": absolute_rect_field(49.65, 152.65, 250, 212.49),
    "policy_number": pattern_only_field(r"Policy Number:?\s+([A-Z0-9]+)"),
    "effective_date": pattern_with_offset_field(
        r"Policy Period At 12:01 A.M. local time at the postal address of the Named Insured",
        dx0=134.80,
        dy0=12.71,
        dx1=-81.36,
        dy1=15.44,
    ),
    "risk_address": pattern_only_field(
        r"Property Coverage \([^)]+\)\s+(.*)", return_all=True
    ),  # Can have multiple
    "form_type": pattern_only_field(
        r"Property Coverage \(([^)]+)\)", return_all=True
    ),  # Can have multiple
    "risk_type": pattern_only_field(
        r"Property Coverage \(([^)]+)\)", return_all=True
    ),  # Can have multiple
    "number_of_families": pattern_with_offset_field(
        r"Number of Families", dx0=0, dy0=18.7, dx1=0, dy1=18.75, return_all=True
    ),
    "earthquake_coverage": pattern_only_field(
        r"Earthquake\s+Damage\s+Assumption\s+End't:\s*(\d+%)\s*Ded"
    ),
    "overland_water": pattern_only_field(r"Overland Water\s+([\d,]+(?:\.\d{2})?)"),
    "condo_deductible": pattern_only_field(r"(\$[\d,]+)\s+Condo\s+Protection"),
    "condo_earthquake_deductible": pattern_only_field(r"Additional Loss Assessment"),
    "service_line": pattern_only_field(r"Water and Sewer Lines\s+([\d,]+(?:\.\d{2})?)"),
    "premium_amount": pattern_only_field(r"Total\s+for\s+Policy\s+([\d,]+)"),
}


WAWANESA_FIELDS = {
    "name_and_address": absolute_rect_field(36.0, 122.43, 200, 180),
    "wawanesa_statement": pattern_only_field(
        "PERSONAL PROPERTY POLICY STATEMENT OF ACCOUNT"
    ),
    "policy_number": pattern_only_field(r"^Policy\s+Number\s+(\d{8})\s*$"),
    "effective_date": pattern_only_field(r"Policy Period From (.+?) to"),
    "risk_address": pattern_with_offset_field(
        r"Location Description Risk Type Residence Type",
        dx0=-119.74,
        dy0=13.78,
        dx1=-165.01,
        dy1=31.85,
        return_all=True,  # Can have multiple locations
    ),
    "form_type": pattern_with_offset_field(
        r"Section I  -  Property Coverage",
        dx0=0.00,
        dy0=-16.80,
        dx1=414.45,
        dy1=-5.35,
        return_all=True,  # Can have multiple locations
    ),
    "risk_type": pattern_with_offset_field(
        r"Location Description Risk Type Residence Type",
        dx0=199.22,
        dy0=13.78,
        dx1=-75.01,
        dy1=31.85,
        return_all=True,  # Can have multiple locations
    ),
    "number_of_families": pattern_only_field(
        r"Number of Families\s+(\d+)", return_all=True
    ),  # Can have multiple
    "number_of_units": pattern_only_field(
        r"Number of Units\s+(\d+)", return_all=True
    ),  # Can have multiple
    "earthquake_coverage": pattern_only_field(r"Earthquake Coverage"),
    "overland_water": pattern_only_field(r"Water Defence - Overland Water Coverage -"),
    "condo_deductible": pattern_with_offset_field(
        r"Condominium Deductible Coverage-",
        dx0=357.90,
        dy0=0.13,
        dx1=107.95,
        dy1=-9.60,
    ),
    "condo_earthquake_deductible": pattern_with_offset_field(
        r"Condominium Deductible Coverage Earthquake-",
        dx0=357.90,
        dy0=0.13,
        dx1=107.95,
        dy1=-9.60,
    ),
    "tenant_vandalism": pattern_only_field(
        r"Vandalism by Tenant Coverage -", return_all=True
    ),
    "service_line": pattern_only_field(r"Service Line Coverage -", return_all=True),
    "premium_amount": pattern_only_field(
        r"Total Policy Premium\s*(\$\s*[\d,]+\.\d{2})"
    ),
    "sewer_back_up_increased_deductible": pattern_only_field(
        r"Limited Sewer Backup coverage deductible has been increased to\s*(\$\s*[\d,]+)"
    ),
    "overland_water_increased_deductible": pattern_only_field(
        r"Overland Water Coverage deductible has increased\s*to\s*(\$\s*[\d,]+)"
    ),
}


# --------------- MAIN CONFIGURATION -----------------
RECTS = {
    "policy_type": POLICY_TYPE_DETECTION,
    "Aviva": AVIVA_FIELDS,
    "Family": FAMILY_FIELDS,
    "Intact": INTACT_FIELDS,
    "Wawanesa": WAWANESA_FIELDS,
}
