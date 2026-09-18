import math
from datetime import datetime
from pathlib import Path
from utils import write_to_new_docx

DATE_FORMAT = "%B %d, %Y"


def safe_strip(value):
    """Safely convert value to stripped string, handling None and NaN."""
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
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
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%B %d, %Y",
        "%b %d, %Y",
    ]
    for fmt in date_formats:
        try:
            return datetime.strptime(value, fmt).strftime(DATE_FORMAT)
        except ValueError:
            continue
    print(f"⚠️ Could not parse effective_date: {value}")
    return value


def map_config_for_renewal(config_data: dict) -> dict:
    return {
        "task": safe_strip(config_data.get("event")),
        "broker_name": safe_strip(config_data.get("broker_name")),
        "on_behalf": safe_strip(config_data.get("on_behalf")),
        "risk_type_1": safe_strip(config_data.get("risk_type")),
        "named_insured": safe_strip(config_data.get("insured_name")),
        "insurer": safe_strip(config_data.get("insurer")),
        "policy_number": safe_strip(config_data.get("policy_number")),
        "effective_date": config_data.get("effective_date"),
        "address_line_one": safe_strip(config_data.get("mailing_street")),
        "address_line_two": safe_strip(config_data.get("city_province")),
        "address_line_three": safe_strip(config_data.get("mailing_postal")),
        "risk_address_1": safe_strip(config_data.get("risk_address")),
    }


def manual_renewal_letter(config: dict) -> None:
    try:
        config = map_config_for_renewal(config)
        config["today"] = datetime.today().strftime(DATE_FORMAT)

        # Build mailing address (all three lines)
        address_fields = ["address_line_one", "address_line_two", "address_line_three"]
        config["mailing_address"] = ", ".join(
            config[f] for f in address_fields if config.get(f)
        )

        # Fall back to address_line_one + address_line_two if risk address missing
        if not config.get("risk_address_1"):
            risk_parts = [
                config.get("address_line_one", ""),
                config.get("address_line_two", ""),
            ]
            config["risk_address_1"] = ", ".join(p for p in risk_parts if p)

        config["effective_date"] = parse_date(config.get("effective_date"))

        template_path = Path.cwd() / "assets" / "Renewal Letter.docx"
        if write_to_new_docx(template_path=template_path, data=config):
            print("******** Manual Renewal Letter ran successfully ********")
    except Exception as e:
        import traceback

        print("❌ Manual Renewal Letter failed")
        traceback.print_exc()
