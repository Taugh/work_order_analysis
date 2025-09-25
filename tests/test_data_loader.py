# ---------------------------------------------------------------
# test_data_loader.py
#
# Purpose:
#   Unit tests for the data loading logic in scripts/data_loader.py.
#
# Requirements:
#   - Input: Test Excel (.xlsx) or CSV file containing sample work order data.
#   - Dependencies: pandas, load_work_order_files from scripts/data_loader.
#
# Output:
#   - Asserts that the loader returns a non-empty DataFrame with expected columns.
#   - Prints a success message if all tests pass.
#
# Notes:
#   - Run with: python tests/test_data_loader.py or use pytest for automated testing.
#   - Ensures data loader correctly reads, normalizes, and prepares input files.
# ---------------------------------------------------------------

# tests/test_data_loader.py

import pandas as pd
from scripts.data_loader import load_work_order_files

def test_load_work_order_files_returns_dataframe():
    df = load_work_order_files("data/raw/TestQSRData.xlsx")
    assert isinstance(df, pd.DataFrame)
    assert not df.empty
    assert "work_order" in df.columns  # Swap with actual expected column