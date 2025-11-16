from datetime import datetime
from pathlib import Path

from utils import write_to_new_docx

DATE_FORMAT = "%B %d, %Y"
TEMPLATE_FILE = Path.cwd() / "assets" / "Renewal Letter.docx"


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
        # Mailing address parts
        "address_line_one": config_data.get("mailing_street", "").strip(),
        "address_line_two": config_data.get("city_province", "").strip(),
        "address_line_three": config_data.get("mailing_postal", "").strip(),
        # Risk address (fallback logic happens inside manual_renewal_letter)
        "risk_address_1": config_data.get("risk_address", "").strip(),
    }

    return mapped


def manual_renewal_letter(config: dict) -> None:
    try:
        config = map_config_for_renewal(config)
        # Add today's date
        config["today"] = datetime.today().strftime(DATE_FORMAT)

        # Build mailing address from address fields
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
                # Try parsing as datetime object first
                if isinstance(effective_date, datetime):
                    config["effective_date"] = effective_date.strftime(DATE_FORMAT)
                else:
                    # Try parsing common string formats
                    date_obj = datetime.strptime(str(effective_date), "%Y-%m-%d")
                    config["effective_date"] = date_obj.strftime(DATE_FORMAT)
            except ValueError:
                try:
                    # Try alternative format
                    date_obj = datetime.strptime(str(effective_date), "%m/%d/%Y")
                    config["effective_date"] = date_obj.strftime(DATE_FORMAT)
                except ValueError:
                    # Keep original if parsing fails
                    print(f"Could not parse effective_date: {effective_date}")

        write_to_new_docx(TEMPLATE_FILE, config)
        print("******** Manual Renewal Letter ran successfully ********")

    except Exception as e:
        print(f"Manual Renewal Letter failed: {e}", exc_info=True)


if __name__ == "__main__":
    test_config = {
        "task": "Renewal",
        "broker_name": "ABC Insurance Brokers",
        "on_behalf": "Client Name",
        "risk_type_1": "Commercial Property",
        "named_insured": "John Doe",
        "insurer": "XYZ Insurance Ltd.",
        "policy_number": "POL123456",
        "effective_date": "2025-01-01",
        "address_line_one": "123 Main Street",
        "address_line_two": "Suite 400",
        "address_line_three": "",
        "risk_address_1": "",
        "number_of_pdfs": 1,
        "drive_letter": "C",
        "producer_names": {"Producer A": "producerA@email.com"},
    }
    manual_renewal_letter(test_config)
