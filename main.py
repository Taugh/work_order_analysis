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
    cleaned_path = "data/processed/cleaned_work_orders.csv"
    df_cleaned.to_csv(cleaned_path, index=False)

    df_classified = apply_classification(df_cleaned)

    # --- Get the last 12 months ---
    last_12_months = (
        df_classified["report_month"]
        .sort_values()
        .drop_duplicates()
        .iloc[-12:]
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
        by_month_df["report_month"]
        .sort_values(key=lambda x: pd.to_datetime(x, format="%b-%y"))
        .drop_duplicates()
        .iloc[-12:]
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
       
