# ---------------------------------------------------------------
# main.py
#
# Purpose:
#   Entry point for the work order analysis application.
#   Supports both CLI and GUI modes for loading, classifying, analyzing,
#   and reporting work order data.
#
# Requirements:
#   - Input: Excel (.xlsx) or CSV file containing raw work order data.
#   - Dependencies: wxPython for GUI, pandas, and all scripts in /scripts and /gui.
#   - Functions: Uses data_loader, classifier, summary_generator, slide_generator, wx_app.
#
# Output:
#   - CLI mode: Prints summary and group/monthly data to console, exports Excel and PowerPoint reports to /outputs.
#   - GUI mode: Launches dashboard for interactive analysis and export.
#
# Workflow:
#   - Loads and cleans input data.
#   - Applies classification logic.
#   - Generates summaries and breakdowns.
#   - Exports results to Excel and PowerPoint.
#   - In GUI mode, provides user feedback and export options.
# ---------------------------------------------------------------

# main.py

import sys
import wx
import pandas as pd

from scripts.data_loader import load_work_order_files
from scripts.classifier import apply_classification
from scripts.summary_generator import (
    generate_monthly_summary,
    get_extreme_late_work_orders,
    export_summary_to_excel,
)
from scripts.slide_generator import create_full_governance_deck
from gui.wx_app import WorkOrderDashboard

def main():
    app = wx.App(False)
    dashboard = WorkOrderDashboard(None, title="Work Order Analysis Dashboard")
    app.MainLoop()

def prepare_data(file_path):
    df_cleaned = load_work_order_files(file_path)
    print("Loaded raw data shape:", df_cleaned.shape)
    print(df_cleaned.head())
    df_classified = apply_classification(df_cleaned)
    cleaned_path = "data/processed/cleaned_work_orders.csv"
    df_classified.to_csv(cleaned_path, index=False)

    df_classified["report_month"] = pd.to_datetime(df_classified["target_date"], errors="coerce").dt.strftime("%b-%y")

    # --- Get the last 12 months ---
    last_12_months = (
        pd.to_datetime(df_classified["report_month"], format="%b-%y")
        .sort_values()
        .drop_duplicates()
        .iloc[-12:]
        .dt.strftime("%b-%y")
        .tolist()
    )
    df_last_12 = df_classified[df_classified["report_month"].isin(last_12_months)]

    # 📊 Generate and export summaries for last 12 months
    summary = generate_monthly_summary(df_last_12)

    by_month_df = (
        df_last_12
        .groupby("report_month")
        .agg(
            missed=("wo_class", lambda x: (x == "missed").sum()),
            completed=("wo_class", lambda x: (x == "on_time").sum()),
            generated=("wo_class", "count")
        )
        .reset_index()
    )

    # --- Ensure by_month_df is sorted and only last 12 months ---
    last_12_months_sorted = (
        pd.to_datetime(by_month_df["report_month"], format="%b-%y")  # Convert strings to datetime
        .sort_values()
        .drop_duplicates()
        .iloc[-12:]
        .dt.strftime("%b-%y")  # Format back to string
        .tolist()
    )
    by_month_df_12 = by_month_df[by_month_df["report_month"].isin(last_12_months_sorted)]

    # --- Use only the newest month for group charts ---
    newest_month = df_classified["report_month"].max()
    df_newest = df_classified[df_classified["report_month"] == newest_month]

    by_group_df = (
        df_newest
        .groupby("group")
        .agg(
            missed=("wo_class", lambda x: (x == "missed").sum()),
            completed=("wo_class", lambda x: (x == "on_time").sum()),
            generated=("wo_class", "count"),
            missed_percent=("wo_class", lambda x: 100 * (x == "missed").sum() / len(x)),
            still_open=("wo_class", lambda x: (x == "open").sum())
        )
        .reset_index()
    )

    late_df = get_extreme_late_work_orders(df_last_12)
    return summary, by_group_df, by_month_df_12, late_df

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # CLI mode
        file_path = sys.argv[1]

        summary, by_group_df, by_month_df_12, late_df = prepare_data(file_path)
        export_summary_to_excel(summary, late_df)

        # Rename columns for summary_df to match slide update expectations
        summary = summary.rename(columns={
            "due": "Due",
            "completed": "Completed",
            "missed": "Missed",
            "completion_pct": "Completion %",
            "canceled": "Canceled"
        })

        

        print("by_month_df_12:\n", by_month_df_12)
        print("by_group_df:\n", by_group_df)
        create_full_governance_deck(summary, by_group_df, by_month_df_12)

    else:
        # GUI mode
        main()

