import math
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
    "number_of_families": pattern_only_field(r"(?:00([12])\s+)?Additional Family"),
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


FAMILY_TENANT_OCCUPIED = "Occupancy: Tenant Occupied"


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
    "occupancy": pattern_only_field(r"(?i)Occupancy:\s*Tenant Occupied"),
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
    "policy_number": pattern_with_offset_field(
        r"Policy Number Policy Period", dx0=0.00, dy0=12.71, dx1=-248.93, dy1=15.44
    ),
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
        dx0=350,
        dy0=0.13,
        dx1=107.95,
        dy1=-9.60,
    ),
    "condo_earthquake_deductible": pattern_with_offset_field(
        r"Condominium Deductible Coverage Earthquake-",
        dx0=350,
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


# --------------- INSURER NAMES -----------------
# Policy number -> insurer, tried in order (first match wins). Patterns were
# derived from the policy_number/insurer columns of config.xlsx
# DN_Transactions; examples are in the comments. Prefixes shared by more
# than one insurer (SEL, QA, bare P + digits) are left out on purpose.
INSURER_RULES = [
    (r"^\d-\d{3}-\d{6,7}$", "Family Insurance"),  # 4-984-1234567
    (r"^RG.*BC$", "Reliance Glass"),  # RG0130997BC
    (r"^K.*H$", "Intact Insurance Company"),  # KR54LL928H
    (r"^4[A-Z]\d+H$", "Intact Insurance Company"),  # 4M1234567H
    (r"^G[RC]\d", "Gore Mutual Insurance Company"),  # GR8695276181
    (r"^VLO", "Vailo Insurance Services Ltd."),  # VLO-HAB-12345678
    (r"^(LTRD|BIND)", "Cansure Insurance Company"),
    (r"^CS\d{6}$", "Cansure Insurance Company"),  # CS604841
    (r"^(WDD|EWL|ADH)\d{7}$", "Cansure Insurance Company"),  # WDD3714768
    (r"^10\d{8}$", "Cansure Insurance Company"),  # 1000008907
    (r"^[59]\d{6}$", "Cansure Insurance Company"),  # 9023222, 5518210
    (r"^0[12]\d{6}$", "Cansure Insurance Company"),  # 01xxxxxx
    (r"^01\d{5}$", "Drivesure Insurance Services Canada Ltd."),  # 0143009
    (r"^P\d{8}[A-Z]{3}$", "Aviva Insurance"),  # P12759130HAB
    (r"^S\d{7}$", "Aviva Insurance"),  # S1598734
    (r"^(CPH|COM)\d{9}$", "Royal & Sun Alliance Insurance Company"),
    (r"^(MOT|SNO)\d", "Beacon Underwriting Ltd."),
    (r"^SWG\d", "South Western Insurance Group"),
    (r"^SGC\d", "Signature Risk Partners Inc."),
    (r"^GUARD", "Guardian Risk Managers"),
    (r"^AUS\d", "Agile Underwriting Solutions"),
    (r"^(INSL|IBC|WAT|HV)\d", "InsureBC Underwriting Services Inc"),
    (r"^E\d{7}$", "InsureBC Underwriting Services Inc"),  # E0001520
    (r"^A\d{6}$", "InsureBC Underwriting Services Inc"),  # A103296
    (r"^(WML|ESM)\d", "Western Underwriting Managers Ltd."),
    (r"^WGL\d", "PAL Insurance Brokers Canada Ltd."),
    (r"^SPG\d", "SPG Canada"),
    (r"^SOP\d", "Totten Group Insurance"),
    (r"^CSD-?\d", "Chutter Underwriting Services"),  # CSD-084516
    (r"^C[VT]\d{6,7}$", "Forward Insurance Managers Ltd."),  # CV1234567
    (
        r"^(RRB|CND|MERC|CBO|VRB|STR|SRA|IFT|MPP)\d",
        "Forward Insurance Managers Ltd.",
    ),
    (r"^S[RHPS]\d{6}$", "Special Risk Insurance Managers Ltd."),  # SR068845
    (r"^GLL\d", "Special Risk Insurance Managers Ltd."),
    (r"^W\d{8}[A-Z]$", "Beazley Canada Limited"),  # W15306121A
    (r"^B\d{9}[A-Z]\d{2}[A-Z]$", "Burns & Wilcox Canada, ULC"),
    (r"^B\d{4}[A-Z]{2}\d{7}$", "Burns & Wilcox Canada, ULC"),  # B0142BL2606236
    (r"^(DN|IL|RA)\d{5}-\d$", "Premier Canada Assurance Managers Ltd."),
    (r"^(ST|EL)\d{5}$", "Premier Canada Assurance Managers Ltd."),  # ST03767
    # Intact numbers sometimes carry a trailing H (e.g. KR54LL928H)
    (r"^5[01]\d{7}H?$", "Intact Insurance Company"),  # 501234567 (9 digits)
    (r"^5[01](?=[0-9A-Z]{6,7}$)\d*[A-Z]", "Intact Insurance Company"),  # 50123RLNS
    (r"^[45][A-Z][0-9A-Z]{7}H?$", "Intact Insurance Company"),  # 4M1234567
    (r"^K[A-Z]\d{2}[A-Z]{2}\d{3}H?$", "Intact Insurance Company"),  # KR54LL928H
    (r"^(AA|AT|CO|CN)\d{7}H?$", "Intact Insurance Company"),  # AA2331869
    (r"^[49]\d{8}H?$", "Intact Insurance Company"),  # 917045751
    (r"^00\d", "Economical Mutual Insurance Company"),  # 004983689
    (r"^04\d{7}$", "Economical Mutual Insurance Company"),  # 040287175
    (r"^48\d{5}$", "Economical Mutual Insurance Company"),  # 4860489
    (r"^[2-5]\d{7}$", "Wawanesa Mutual Insurance Company"),  # 38123456
]
INSURER_RULES = [(re.compile(p), name) for p, name in INSURER_RULES]

# Best guess when no INSURER_RULES entry matches. Wawanesa, Intact and Aviva
# are ~60% of DN_Transactions, so an unknown number goes to whichever of the
# three it looks most like.
LIKELY_INSURER_RULES = [
    (r"^\d{8}$", "Wawanesa Mutual Insurance Company"),  # 494 of 507 8-digit rows
    (r"^[A-Z]\d+[A-Z]{3}$", "Aviva Insurance"),  # P12759130HAB
    (r"^P\d", "Aviva Insurance"),
    (r"^\d+$", "Wawanesa Mutual Insurance Company"),  # other all-digit numbers
    (r".", "Intact Insurance Company"),  # anything else (mostly 9-char alphanumeric)
]
LIKELY_INSURER_RULES = [(re.compile(p), name) for p, name in LIKELY_INSURER_RULES]


def get_insurer(policy_number, guess=False):
    """Determines insurer from the policy number using INSURER_RULES.
    Returns "" when no rule matches, unless guess=True, in which case the
    most likely of Wawanesa / Intact / Aviva is returned instead."""
    if not policy_number:
        return ""

    p = str(policy_number).strip().upper()
    rules = INSURER_RULES + LIKELY_INSURER_RULES if guess else INSURER_RULES
    for rx, name in rules:
        if rx.search(p):
            return name
    return ""


INSURER_MAP = {
    "999999": "Cansure Insurance Company",
    "ACTURI": "Acturis Rating",
    "AGI": "Agile Underwriting Solutions",
    "ALL": "Allianz",
    "AVIV": "Aviva Insurance",
    "BEA": "Beacon Underwriting Ltd.",
    "BEAZ": "Beazley Canada Limited",
    "BECK": "Beck Glass (2012) Ltd.",
    "BUR": "Burns & Wilcox Canada, ULC",
    "CAN": "Cansure Insurance Company",
    "CHU": "Chutter Underwriting Services",
    "CNS": "Royal & Sun Alliance Insurance Company",
    "DRI": "Drivesure Insurance Services Canada Ltd.",
    "ECON": "Economical Mutual Insurance Company",
    "ELT": "Aviva Elite",
    "FIC": "Family Insurance",
    "FOR": "Forward Insurance Managers Ltd.",
    "GUA": "Guardian Risk Managers",
    "HWI": "Horizon West Insurance Services Ltd.",
    "I3U": "I3 Underwriting Services",
    "INS": "Insurebc Underwriting Services Inc",
    "OPT": "Optiom Inc.",
    "PAL": "PAL Insurance Brokers Canada Ltd.",
    "PBC": "Pacific Blue Cross",
    "PRE": "Premier Canada Assurance Managers Ltd.",
    "PREM": "Premier Marine Insurance Managers Group",
    "REL": "Reliance Glass",
    "SIG": "Signature Risk Partners Inc.",
    "SOU": "South Western Insurance Group",
    "SPG": "SPG Canada",
    "SRI": "Special Risk Insurance Managers Ltd.",
    "SUM": "Strategic Underwriting Managers Inc.",
    "TOT": "Totten Group Insurance",
    "TSW": "TSW Management Services Inc.",
    "VAI": "Vailo Insurance Services Ltd.",
    "WAWA": "Wawanesa Mutual Insurance Company",
    "WELL": "Intact Specialty Solutions",
    "WES": "Western Underwriting Managers Ltd.",
    "WESU": "Intact Insurance Company",
}


def _capitalize_word(word: str) -> str:
    return word[:1].upper() + word[1:].lower() if word else word


def title_case_generic(text) -> str:
    """Plain title-case: capitalizes each letter-run, leaves everything
    else (digits, punctuation, spacing) untouched."""
    if text is None:
        return ""
    if isinstance(text, float) and math.isnan(text):
        return ""
    return re.sub(r"[A-Za-z]+", lambda m: _capitalize_word(m.group(0)), str(text))


SHORT_INSURER_FULL = {
    "Gore Mutual": "Gore Mutual Insurance Company",
    "Wawanesa": "Wawanesa Mutual Insurance Company",
    "Intact": "Intact Insurance Company",
    "Economical": "Economical Mutual Insurance Company",
    "Cansure": "Cansure Insurance Company",
    "Vailo": "Vailo Insurance Services Ltd.",
}


def resolve_insurer(insurer_code, policy_number=""):
    """Maps an insurer code (e.g. WAWA, ECON, WESU) to the same full
    company name used by disclosure_notice.py. HWI rows are resolved by
    policy-number prefix; empty/unmapped codes fall back to prefix
    detection. Short names returned by get_insurer (e.g. 'Wawanesa',
    'Gore Mutual') are expanded to the full company name, and bare
    'Cansure' is always expanded to the full name."""
    raw = "" if insurer_code is None else str(insurer_code).strip()
    code = raw.upper()

    if not code:
        name = get_insurer(policy_number)
    elif code == "HWI":
        name = get_insurer(policy_number) or "Cansure Insurance Company"
    else:
        name = INSURER_MAP.get(code, title_case_generic(raw))

    name = SHORT_INSURER_FULL.get(str(name).strip(), name)
    if str(name).strip().upper() == "CANSURE":
        name = "Cansure Insurance Company"
    return name
