import re
import fitz

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

RECTS = {
    "policy_type": {
        "Aviva": {"keyword": "Aviva", "rect": fitz.Rect(183.84, 728.4, 197.98, 734.4)},
        "Family": {"keyword": "Agent", "rect": fitz.Rect(25.70, 36.37, 51.04, 45.45)},
        "Intact": {"keyword": "Intact", "rect": fitz.Rect(36, 633.03, 576.26, 767.95)},
        "Wawanesa": {
            "keyword": "BROKER OFFICE",
            "rect": fitz.Rect(36.0, 102.43, 353.27, 111.37),
        },
    },
    "Aviva": {
        "name_and_address": {
            "pattern": None,
            "rect": fitz.Rect(80.4, 202.24, 250, 280),
        },
        "policy_number": {
            "pattern": re.compile(r"Policy Number", re.IGNORECASE),
            "rect": fitz.Rect(267.12, 10.16, -202.82, 9.16),
        },
        "effective_date": {
            "pattern": re.compile(r"Policy Effective From:", re.IGNORECASE),
            "rect": None,
        },
        "risk_address": {
            "pattern": re.compile(r"Location\s", re.IGNORECASE),
            "rect": None,
        },
        "form_type": {
            "pattern": re.compile(r"Residence Locations:\s", re.IGNORECASE),
            "rect": None,
        },
        "risk_type": {
            "pattern": re.compile(r"Residence Locations:\s", re.IGNORECASE),
            "rect": None,
        },
        "number_of_families": {
            "pattern": re.compile(r"Extended Liability", re.IGNORECASE),
            "rect": None,
        },
        "earthquake_coverage": {
            "pattern": re.compile(r"Earthquake", re.IGNORECASE),
            "rect": None,
        },
        "overland_water": {
            "pattern": re.compile(r"Overland Water", re.IGNORECASE),
            "rect": None,
        },
        "condo_deductible": {
            "pattern": re.compile(r"Condominium Corporation Deductible", re.IGNORECASE),
            "rect": None,
        },
        "service_line": {
            "pattern": re.compile(r"Service Line Coverage", re.IGNORECASE),
            "rect": None,
        },
        "premium_amount": {
            "pattern": re.compile(r"TOTAL", re.IGNORECASE),
            "rect": None,
        },
    },
    "Family": {
        "name_and_address": {
            "pattern": None,
            "rect": fitz.Rect(25.34, 153.38, 150, 228.67),
        },
        "policy_number": {
            "pattern": re.compile(r"POLICY NUMBER", re.IGNORECASE),
            "rect": fitz.Rect(-1.08, 11.03, 8.86, 25.99),
        },
        "effective_date": {
            "pattern": re.compile(r"EFFECTIVE DATE", re.IGNORECASE),
            "rect": fitz.Rect(-1.01, 20.17, 24.57, 11.45),
        },
        "risk_address": {
            "pattern": re.compile(r"LOCATION OF INSURED PROPERTY:", re.IGNORECASE),
            "rect": None,
        },
        "form_type": {
            "pattern": re.compile(r"All Perils:", re.IGNORECASE),
            "rect": None,
        },
        "risk_type": {
            "pattern": re.compile(r"POLICY TYPE", re.IGNORECASE),
            "rect": fitz.Rect(-0.94, 11.10, 7.76, 11.45),
        },
        "number_of_families": {
            "pattern": re.compile(r"RENTAL SUITES", re.IGNORECASE),
            "rect": None,
        },
        "earthquake_coverage": {
            "pattern": re.compile(r"EARTHQUAKE PROPERTY LIMITS", re.IGNORECASE),
            "rect": fitz.Rect(112.40, 12.74, 42, 12.37),
        },
        "overland_water": {
            "pattern": re.compile(r"Overland Water", re.IGNORECASE),
            "rect": None,
        },
        "condo_deductible": {
            "pattern": re.compile(r"Deductible Coverage:", re.IGNORECASE),
            "rect": None,
        },
        "service_line": {
            "pattern": re.compile(r"Service Lines", re.IGNORECASE),
            "rect": None,
        },
        "premium_amount": {
            "pattern": re.compile(r"RETURN THIS PORTION WITH PAYMENT", re.IGNORECASE),
            "rect": fitz.Rect(5.59, -22.79, -116.08, -22.20),
        },
    },
    "Intact": {
        "name_and_address": {
            "pattern": None,
            "rect": fitz.Rect(49.65, 152.65, 250, 212.49),
        },
        "policy_number": {
            "pattern": re.compile(r"Policy Number", re.IGNORECASE),
            "rect": fitz.Rect(267.12, 10.16, -202.82, 9.16),
        },
        "effective_date": {
            "pattern": re.compile(r"Policy Number", re.IGNORECASE),
            "rect": None,
        },
        "risk_address": {
            "pattern": re.compile(r"Property Coverage", re.IGNORECASE),
            "rect": fitz.Rect(0.0, 16.9, 0.0, 0.0),
        },
        "form_type": {
            "pattern": re.compile(r"Property Coverage", re.IGNORECASE),
            "rect": fitz.Rect(110.65, 0.0, 0, -10.03),
        },
        "risk_type": {
            "pattern": re.compile(r"Property Coverage", re.IGNORECASE),
            "rect": fitz.Rect(110.65, 0.0, 0, -10.03),
        },
        "number_of_families": {
            "pattern": re.compile(r"Families", re.IGNORECASE),
            "rect": fitz.Rect(0, 18.7, 0, 18.75),
        },
        "earthquake_coverage": {
            "pattern": re.compile(r"Earthquake Damage Assumption", re.IGNORECASE),
            "rect": fitz.Rect(46.94, 11.17, 35, 12.27),
        },
        "overland_water": {
            "pattern": re.compile(r"Enhanced Water Damage", re.IGNORECASE),
            "rect": None,
        },
        "condo_deductible": {
            "pattern": re.compile(r"Deductible Coverage:", re.IGNORECASE),
            "rect": None,
        },
        "condo_earthquake_deductible": {
            "pattern": re.compile(r"Additional Loss Assessment", re.IGNORECASE),
            "rect": None,
        },
        "service_line": {
            "pattern": re.compile(r"Water and Sewer Lines", re.IGNORECASE),
            "rect": None,
        },
        "premium_amount": {
            "pattern": re.compile(r"Total for Policy", re.IGNORECASE),
            "rect": None,
        },
    },
    "Wawanesa": {
        "name_and_address": {
            "pattern": None,
            "rect": fitz.Rect(36.0, 122.43, 200, 180),
        },
        "policy_number": {
            "pattern": re.compile(r"NAMED INSURED AND ADDRESS", re.IGNORECASE),
            "rect": None,
        },
        "effective_date": {
            "pattern": re.compile(r"NAMED INSURED AND ADDRESS", re.IGNORECASE),
            "rect": None,
        },
        "risk_address": {
            "pattern": re.compile(r"Location Description", re.IGNORECASE),
            "rect": None,
        },
        "form_type": {
            "pattern": re.compile(
                r"subject to all conditions of the policy.", re.IGNORECASE
            ),
            "rect": None,
        },
        "risk_type": {"pattern": re.compile(r"Risk Type", re.IGNORECASE), "rect": None},
        "number_of_families": {
            "pattern": re.compile(r"Number of Families", re.IGNORECASE),
            "rect": None,
        },
        "number_of_units": {
            "pattern": re.compile(r"Number of Units", re.IGNORECASE),
            "rect": None,
        },
        "earthquake_coverage": {
            "pattern": re.compile(r"Earthquake Coverage", re.IGNORECASE),
            "rect": None,
        },
        "overland_water": {
            "pattern": re.compile(r"Overland Water", re.IGNORECASE),
            "rect": None,
        },
        "condo_deductible": {
            "pattern": re.compile(r"Condominium Deductible Coverage-", re.IGNORECASE),
            "rect": None,
        },
        "condo_earthquake_deductible": {
            "pattern": re.compile(
                r"Condominium Deductible Coverage Earthquake", re.IGNORECASE
            ),
            "rect": fitz.Rect(350.90, 0.0, 95.43, -9.60),
        },
        "tenant_vandalism": {
            "pattern": re.compile(r"Vandalism by Tenant Coverage -", re.IGNORECASE),
            "rect": None,
        },
        "service_line": {
            "pattern": re.compile(r"Service Line Coverage", re.IGNORECASE),
            "rect": None,
        },
        "premium_amount": {
            "pattern": re.compile(r"Total Policy Premium", re.IGNORECASE),
            "rect": None,
        },
    },
}
