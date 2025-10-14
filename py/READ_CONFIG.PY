import pandas as pd
from pathlib import Path

CONFIG_PATH = Path("../config.xlsx")  # Default path, can be overridden

CONFIG_FIELDS = {
    "task": (2, 1, None),
    "broker_name": (6, 1, None),
    "on_behalf": (8, 1, None),
    "risk_type_1": (12, 1, None),
    "named_insured": (14, 1, None),
    "insurer": (15, 1, None),
    "policy_number": (16, 1, None),
    "effective_date": (17, 1, None),
    "address_line_one": (19, 1, None),
    "address_line_two": (20, 1, None),
    "address_line_three": (21, 1, None),
    "risk_address_1": (23, 1, None),
    "number_of_pdfs": (27, 1, 0),
    "drive_letter": (29, 1, None)
}

PRODUCER_RANGE = (33, 49)

def safe_get(df, row, col, default=None, cast=None):
    try:
        value = df.iat[row, col]
        if pd.isna(value) or value == "":
            return default
        if cast:
            return cast(value)
        return value
    except (IndexError, KeyError):
        return default

def read_config(config_path=CONFIG_PATH):
    """
    Reads the Excel config file and returns a dictionary of config values.
    """
    try:
        df = pd.read_excel(config_path, sheet_name=0, header=None)
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel file '{config_path}': {e}")

    # Single-value config items
    data = {key: safe_get(df, row, col, default) for key, (row, col, default) in CONFIG_FIELDS.items()}

    # Multiple producer names
    start, end = PRODUCER_RANGE
    producer_df = df.iloc[start:end, 1:3].dropna(subset=[1])
    data["producer_names"] = dict(zip(producer_df[1], producer_df[2]))

    return data

# Auto-load the config at import
config = read_config()
