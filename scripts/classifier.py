# ---------------------------------------------------------------
# classifier.py
#
# Purpose:
#   Provides functions to classify work orders based on status, dates, and type.
#
# Requirements:
#   - Input: pandas DataFrame with columns: 'status', 'target_date', 'actual_finish', 'grace_date'.
#   - Libraries: pandas.
#
# Output:
#   - Adds a 'wo_class' column to the DataFrame with values: 'canceled', 'open', 'on_time', or 'missed'.
#   - Optionally, can classify work order type (e.g., PM, CA, RQL, Other).
#
# Notes:
#   - Used by analysis and reporting modules to segment work orders for summary and charts.
#   - Main functions: classify_work_order(row), apply_classification(df).
# ---------------------------------------------------------------

# scripts/classifier.py

import pandas as pd

def classify_work_order(row):
    status = str(row.get("status", "")).upper()
    target_date = row.get("target_date")
    finish_date = row.get("actual_finish")
    grace_date = row.get("grace_date")

    if status == "CAN":
        return "canceled"
    
    if pd.isna(finish_date) or status not in ["COMP", "CORRECTED",
                                              "CORRTD", "PENDQA", "PENRVW", "REVWD",
                                              "CLOSE"]:
        return "open"
    elif finish_date <= grace_date:
        return "on_time"
    else:
        return "missed"

def apply_classification(df):
    df["wo_class"] = df.apply(classify_work_order, axis=1)    
    return df

##def classify_work_type(row):
##    desc = str(row.get("work_type", "")).upper()
##    if "PM" in desc:
##        return "PM"
##    elif "CA" in desc:
##        return "CA"
##    elif "RQL" in desc:
##        return "RQL"
##    else:
##        return "Other"
