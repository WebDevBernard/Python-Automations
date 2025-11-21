import time
from utils import load_excel_mapping
from constants import DEFAULT_MAPPING, EXCEL_CELL_MAPPING
from pathlib import Path
from manual_renewal_letter import manual_renewal_letter
from sort_renewal_list import sort_renewal_list
from auto_renewal_letter import auto_renewal_letter
from reconciller import reconciller


def main():
    mapping_path = "config.xlsx"
    try:
        config_data = load_excel_mapping(
            mapping_path, DEFAULT_MAPPING, EXCEL_CELL_MAPPING
        )
    except Exception as e:
        print(f"Failed to load Excel mapping: {e}")
        return

    event = config_data.get("event", "").strip().lower()

    if event == "manual renewal letter":
        manual_renewal_letter(config_data)
    elif event == "auto renewal letter":
        auto_renewal_letter(config_data)
    elif event == "sort renewal list":
        sort_renewal_list()
    elif event == "reconciller":
        reconciller(config_data)
    else:
        print(f"Unknown event: {event}")

    print("\nExiting in ", end="")
    for i in range(3, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)


if __name__ == "__main__":
    main()
