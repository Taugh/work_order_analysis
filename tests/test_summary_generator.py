# tests/test_summary_generator.py

import pandas as pd
from scripts.summary_generator import (
    generate_monthly_summary,
    get_extreme_late_work_orders
)

def sample_df():
    df = pd.DataFrame([
        {
            "WorkOrderID": 1,
            "work_order": "WO-001",
            "wo_assigned_group": "Facilities",
            "description": "Fix door hinge",
            "Date": "2023-06-01",
            "status": "REVWD",
            "actual_finish": pd.Timestamp("2023-06-15"),
            "target_date": pd.Timestamp("2023-06-10"),  # adjust per row for variety
            "grace_date": pd.Timestamp("2023-06-10"),
            "wo_class": "missed"
        },
        {
            "WorkOrderID": 2,
            "work_order": "WO-001",
            "wo_assigned_group": "Facilities",
            "description": "Fix door hinge",
            "Date": "2023-07-10",
            "status": "COMP",
            "actual_finish": pd.Timestamp("2023-07-12"),
            "target_date": pd.Timestamp("2023-06-10"),  # adjust per row for variety
            "grace_date": pd.Timestamp("2023-07-15"),
            "wo_class": "on_time"
        },
        {
            "WorkOrderID": 3,
            "work_order": "WO-001",
            "wo_assigned_group": "Facilities",
            "description": "Fix door hinge",
            "Date": "2023-07-25",
            "status": "INPRG",
            "actual_finish": pd.NaT,
            "target_date": pd.Timestamp("2023-06-10"),  # adjust per row for variety
            "grace_date": pd.Timestamp("2023-08-01"),
            "wo_class": "open"
        },
        {
            "WorkOrderID": 4,
            "work_order": "WO-001",
            "wo_assigned_group": "Facilities",
            "description": "Fix door hinge",
            "Date": "2023-07-30", 
            "status": "CAN",
            "actual_finish": pd.Timestamp("2023-07-30"),
            "target_date": pd.Timestamp("2023-06-10"),  # adjust per row for variety
            "grace_date": pd.Timestamp("2023-08-05"),
            "wo_class": "canceled"
        }

    ])
    df["report_month"] = pd.to_datetime(df["Date"]).dt.to_period("M").astype(str)
    return df

def test_generate_monthly_summary_counts():
    df = sample_df()
    summary = generate_monthly_summary(df)

    assert summary.loc[0, "Missed"] == 1
    assert summary.loc[1, "Canceled"] == 1
    assert summary.loc["Grand Total", "Completed"] == 1
    assert summary.loc["Grand Total", "Still Open"] == 1
    assert summary.loc["Grand Total", "Completion %"] == 25.0


    # print(summary)
    

def test_get_extreme_late_work_orders():
    df = sample_df()
    late_df = get_extreme_late_work_orders(df)

    assert late_df["late_days"].iloc[0] > 90
    assert late_df["wo_class"].iloc[0] == "open"
    assert late_df["status"].iloc[0] in ["INPRG", "APPR", "WAPPR"]


