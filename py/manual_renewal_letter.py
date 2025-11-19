from datetime import datetime
from pathlib import Path
from utils import write_to_new_docx

DATE_FORMAT = "%B %d, %Y"


def map_config_for_renewal(config_data: dict) -> dict:
    mapped = {
        "task": config_data.get("event", "").strip(),
        "broker_name": config_data.get("broker_name", "").strip(),
        "on_behalf": config_data.get("on_behalf", "").strip(),
        "risk_type_1": config_data.get("risk_type", "").strip(),
        "named_insured": config_data.get("insured_name", "").strip(),
        "insurer": config_data.get("insurer", "").strip(),
        "policy_number": config_data.get("policy_number", "").strip(),
        "effective_date": config_data.get("effective_date", "").strip(),
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
        address_parts = [
            config.get(field, "").strip()
            for field in address_fields
            if config.get(field, "").strip()
        ]
        config["mailing_address"] = ", ".join(address_parts)

        # Use mailing address if risk_address_1 is empty
        if not config.get("risk_address_1", "").strip():
            config["risk_address_1"] = config["mailing_address"]

        # Parse and format effective_date
        effective_date = config.get("effective_date")
        if effective_date:
            try:
                if isinstance(effective_date, datetime):
                    config["effective_date"] = effective_date.strftime(DATE_FORMAT)
                else:
                    date_obj = datetime.strptime(str(effective_date), "%Y-%m-%d")
                    config["effective_date"] = date_obj.strftime(DATE_FORMAT)
            except ValueError:
                try:
                    date_obj = datetime.strptime(str(effective_date), "%m/%d/%Y")
                    config["effective_date"] = date_obj.strftime(DATE_FORMAT)
                except ValueError:
                    print(f"Could not parse effective_date: {effective_date}")

        if write_to_new_docx(data=config):
            print("******** Manual Renewal Letter ran successfully ********")
    except Exception as e:
        print(f"Manual Renewal Letter failed: {e}", exc_info=True)
