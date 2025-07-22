# tests/test_classifier.py

import pandas as pd
from scripts.classifier import apply_classification

def test_classify_canceled():
    df = pd.DataFrame([{
        "status": "CAN",
        "actual_finish": "2023-06-01",
        "grace_date": "2023-06-10"
    }])
    result = apply_classification(df)
    assert result["wo_class"].iloc[0] == "canceled"

def test_classify_open_missing_finish():
    df = pd.DataFrame([{
        "status": "COMP",
        "actual_finish": pd.NaT,
        "grace_date": "2023-06-10"
    }])
    result = apply_classification(df)
    assert result["wo_class"].iloc[0] == "open"

def test_classify_open_unexpected_status():
    df = pd.DataFrame([{
        "status": "CREATED",
        "actual_finish": "2023-06-01",
        "grace_date": "2023-06-10"
    }])
    result = apply_classification(df)
    assert result["wo_class"].iloc[0] == "open"

def test_classify_on_time():
    df = pd.DataFrame([{
        "status": "COMP",
        "actual_finish": pd.Timestamp("2023-06-01"),
        "grace_date": pd.Timestamp("2023-06-10")
    }])
    result = apply_classification(df)
    assert result["wo_class"].iloc[0] == "on_time"

def test_classify_missed():
    df = pd.DataFrame([{
        "status": "REVWD",
        "actual_finish": pd.Timestamp("2023-06-15"),
        "grace_date": pd.Timestamp("2023-06-10")
    }])
    result = apply_classification(df)
    assert result["wo_class"].iloc[0] == "missed"

if __name__ == "__main__":
    print("✔️ test_classifier.py ran successfully.")
