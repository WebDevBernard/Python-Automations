from datetime import datetime
from pathlib import Path
from utils import write_to_new_docx

DATE_FORMAT = "%B %d, %Y"


def safe_strip(value):
    """Safely convert value to stripped string."""
    if value is None:
        return ""
    return str(value).strip()


def parse_date(value):
    """Parse dates from multiple formats including Excel datetimes."""
    if not value:
        return ""

    if isinstance(value, datetime):
        return value.strftime(DATE_FORMAT)

    value = str(value).strip()

    date_formats = [
        "%Y-%m-%d",  # 2026-02-01
        "%m/%d/%Y",  # 02/01/2026
        "%B %d, %Y",  # February 01, 2026
        "%b %d, %Y",  # Feb 01, 2026
    ]

    for fmt in date_formats:
        try:
            return datetime.strptime(value, fmt).strftime(DATE_FORMAT)
        except ValueError:
            continue

    print(f"⚠️ Could not parse effective_date: {value}")
    return value


def map_config_for_renewal(config_data: dict) -> dict:
    mapped = {
        "task": config_data.get("event", "").strip(),
        "broker_name": config_data.get("broker_name", "").strip(),
        "on_behalf": config_data.get("on_behalf", "").strip(),
        "risk_type_1": config_data.get("risk_type", "").strip(),
        "named_insured": config_data.get("insured_name", "").strip(),
        "insurer": config_data.get("insurer", "").strip(),
        "policy_number": config_data.get("policy_number", "").strip(),
        "effective_date": config_data.get("effective_date"),
        "address_line_one": config_data.get("mailing_street", "").strip(),
        "address_line_two": config_data.get("city_province", "").strip(),
        "address_line_three": config_data.get("mailing_postal", "").strip(),
        "risk_address_1": config_data.get("risk_address", "").strip(),
    }
    return mapped


def manual_renewal_letter(config: dict) -> None:
    try:
        config = map_config_for_renewal(config)
        config["today"] = datetime.today().strftime(DATE_FORMAT)

        # Build mailing address
        address_fields = ["address_line_one", "address_line_two", "address_line_three"]
        address_parts = [config[f] for f in address_fields if config.get(f)]
        config["mailing_address"] = ", ".join(address_parts)

        # Use mailing address if risk address missing
        if not config.get("risk_address_1"):
            config["risk_address_1"] = config["mailing_address"]

        # Parse effective_date
        config["effective_date"] = parse_date(config.get("effective_date"))

        if write_to_new_docx(data=config):
            print("******** Manual Renewal Letter ran successfully ********")

    except Exception as e:
        import traceback

        print("❌ Manual Renewal Letter failed")
        traceback.print_exc()
