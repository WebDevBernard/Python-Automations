from utils import load_excel_mapping
from constants import DEFAULT_MAPPING, EXCEL_CELL_MAPPING
from pathlib import Path
from manual_renewal_letter import (
    manual_renewal_letter,
)
from sort_renewal_list import sort_renewal_list


def main():
    mapping_path = "config.xlsx"

    try:
        config_data = load_excel_mapping(
            mapping_path, DEFAULT_MAPPING, EXCEL_CELL_MAPPING
        )
    except Exception as e:
        print(f"Failed to load Excel mapping: {e}")
        return

    if config_data.get("event", "").strip().lower() == "manual renewal letter":
        manual_renewal_letter(config_data)
    if config_data.get("event", "").strip().lower() == "sort renewal list":
        sort_renewal_list()


if __name__ == "__main__":
    main()
