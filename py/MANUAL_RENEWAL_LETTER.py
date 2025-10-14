from datetime import datetime
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
import logging

from UTILITIES import unique_file_name

DATE_FORMAT = "%B %d, %Y"
TEMPLATE_FILE = Path(__file__).resolve().parent.parent / "assets" / "Renewal Letter New.docx"

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


def write_to_new_docx(template_path: Path, data: dict, output_dir: Path | None = None) -> None:
    template_path = Path(template_path)
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")
    
    doc = DocxTemplate(template_path)
    doc.render(data)

    named_insured = str(data.get("named_insured", "Unnamed Client")).rstrip(".:").strip()
    output_dir = output_dir or (Path.home() / "Desktop")
    output_filename = output_dir / f"{named_insured} Renewal Letter.docx"

    doc.save(unique_file_name(output_filename))
    logging.info(f"Saved renewal letter for {named_insured} -> {output_filename}")


def manual_renewal_letter(config: dict) -> None:
    try:
        df = pd.DataFrame([config])
        df["today"] = datetime.today().strftime(DATE_FORMAT)

        address_fields = ["address_line_one", "address_line_two", "address_line_three"]
        df["mailing_address"] = (
            df[address_fields]
            .fillna("")
            .agg(lambda x: ", ".join(filter(None, map(str.strip, x))), axis=1)
        )

        df["risk_address_1"] = df["risk_address_1"].fillna(df["mailing_address"])
        df["effective_date"] = pd.to_datetime(df["effective_date"], errors="coerce").dt.strftime(DATE_FORMAT)

        for row in df.to_dict(orient="records"):
            write_to_new_docx(TEMPLATE_FILE, row)

        logging.info("******** Manual Renewal Letter ran successfully ********")

    except Exception as e:
        logging.error(f"Manual Renewal Letter failed: {e}", exc_info=True)


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
