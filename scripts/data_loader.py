# ---------------------------------------------------------------
# data_loader.py
#
# Purpose:
#   Loads raw work order data from a specified file, normalizes column names,
#   converts date columns, and prepares a clean DataFrame for analysis.
#
# Requirements:
#   - Input: Excel (.xlsx) or CSV file containing raw work order data.
#   - Config: COLUMN_MAP for column normalization, RAW_DATA_DIR for default paths.
#   - Libraries: pandas, pathlib.
#
# Output:
#   - Returns a pandas DataFrame with normalized columns and converted date fields.
#   - Adds a 'report_month' column as a pandas Period for monthly grouping.
#
# Notes:
#   - Used by analysis and classification modules as the first step in the workflow.
#   - Handles both Excel and CSV formats.
# ---------------------------------------------------------------

# scripts/data_loader.py

import pandas as pd
from pathlib import Path
from config.settings import RAW_DATA_DIR, COLUMN_MAP

def load_work_order_files(file_path):
    """Loads a single file specified by file_path, normalizes columns, converts dates, and returns a clean DataFrame."""

    # Load the file based on its extension
    if file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path, engine="openpyxl")
    else:
        df = pd.read_csv(file_path)

    # Normalize column names
    df = df.rename(columns=COLUMN_MAP)

    # Convert date columns
    for col in ["target_date", "actual_finish", "finish_no_later", "report_date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Add report month column
    df["report_month"] = df["target_date"].dt.to_period("M")

    return df


