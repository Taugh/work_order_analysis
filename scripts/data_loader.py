# scripts/data_loader.py

import pandas as pd
from pathlib import Path
from config.settings import RAW_DATA_DIR, COLUMN_MAP

def load_work_order_files(file_path):
    """A clean and extensible data loader that reads all Excel files in
    your '/data/raw' directory, normalizes the column names, converts
    dates, and returns a clean DataFrame ready for classification and
    analysis."""

    all_files = list(Path(RAW_DATA_DIR).glob("*.xlsx"))
    if not all_files:
        raise FileNotFoundError(f"No Excel files found in {RAW_DATA_DIR}")

    dfs = []

    for file in all_files:
        df = pd.read_excel(file, engine="openpyxl")

        # Normalize column names
        df = df.rename(columns=COLUMN_MAP)

        # Convert date columns
        for col in ["target_date", "actual_finish", "finish_no_later", "report_date"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")

        # Add report month column
        df["report_month"] = df["target_date"].dt.to_period("M")

        dfs.append(df)
            # After loading
    # Combine all files
    combined_df = pd.concat(dfs, ignore_index=True)
    print(combined_df.columns)
    return combined_df
    

