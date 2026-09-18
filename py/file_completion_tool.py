import time
from utils import load_excel_mapping
from constants import DEFAULT_MAPPING, EXCEL_CELL_MAPPING
from manual_renewal_letter import manual_renewal_letter
from sort_renewal_list import sort_renewal_list
from auto_renewal_letter import auto_renewal_letter
from disclosure_notice_strata_from_statement import strata_disclosure_from_statement


def main():
    mapping_path = "config.xlsx"
    config_data = {}
    try:
        config_data = load_excel_mapping(
            mapping_path, DEFAULT_MAPPING, EXCEL_CELL_MAPPING
        )
    except (ValueError, FileNotFoundError):
        config_data = {}

    event = config_data.get("event", "").strip().lower() if config_data else ""

    if event == "manual renewal letter":
        manual_renewal_letter(config_data)
    elif event == "auto renewal letter":
        auto_renewal_letter(config_data)
    elif event == "sort renewal list":
        sort_renewal_list()
    elif event == "strata disclosure letter from statment":
        strata_disclosure_from_statement(config_data)
    else:
        print(f"Unknown event: {event}")

    print("\nExiting in ", end="")
    for i in range(3, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)


if __name__ == "__main__":
    main()
