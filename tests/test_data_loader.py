# tests/test_data_loader.py

import pandas as pd
from scripts.data_loader import load_work_order_files

def test_load_work_order_files_returns_dataframe():
    df = load_work_order_files("data/raw/TestQSRData.xlsx")
    assert isinstance(df, pd.DataFrame)
    assert not df.empty
    assert "work_order" in df.columns  # Swap with actual expected column