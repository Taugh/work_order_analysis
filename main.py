# ---------------------------------------------------------------
# main.py
#
# Purpose:
#   Entry point for the work order analysis application.
#   Supports both CLI and GUI modes for loading, classifying, analyzing,
#   and reporting work order data.
# ---------------------------------------------------------------

import sys
import pandas as pd

from scripts.data_loader import load_work_order_files
from scripts.classifier import apply_classification
from scripts.summary_generator import (
    generate_monthly_summary,
    get_extreme_late_work_orders,
    export_summary_to_excel,
)
from scripts.slide_generator import create_full_governance_deck
from scripts.summary_generator import generate_12_month_trend

print("main.py started")

def prepare_data(file_path):
    # Load and classify data
    print("prepare_data called with", file_path)
    print("Starting prepare_data")  # Debug print
    df_cleaned = load_work_order_files(file_path)
    print("Loaded raw data shape:", df_cleaned.shape)
    df_classified = apply_classification(df_cleaned)
    cleaned_path = "data/processed/cleaned_work_orders.csv"
    df_classified.to_csv(cleaned_path, index=False)
    

    # Ensure target_date is datetime
    df_classified["target_date"] = pd.to_datetime(df_classified["target_date"], errors="coerce")

    # --- Build last 12 complete months using true date boundaries ---
    trend_df = generate_12_month_trend(df_classified)
    print("trend_df created")
    print(trend_df)
    print(trend_df["report_month"])

    # --- Use only the previous month for group charts ---
    today = pd.Timestamp.today()
    first_of_current = today.replace(day=1)
    first_of_previous = (first_of_current - pd.DateOffset(months=1)).replace(day=1)
    mask = (df_classified["target_date"] > first_of_previous) & (df_classified["target_date"] <= first_of_current)
    df_prev_month = df_classified[mask]

    by_group_df = (
        df_prev_month
        .groupby("group")
        .agg(
            missed=("wo_class", lambda x: (x == "missed").sum()),
            completed=("wo_class", lambda x: (x == "on_time").sum()),
            generated=("wo_class", "count"),
            missed_percent=("wo_class", lambda x: 100 * (x == "missed").sum() / len(x) if len(x) else 0),
            still_open=("wo_class", lambda x: (x == "open").sum())
        )
        .reset_index()
    )

    # For summary and late_df, use only the last 12 months' data (from trend_df boundaries)
    # Use the same month boundaries as trend_df for summary and late_df
    today = pd.Timestamp.today()
    first_of_current = today.replace(day=1)
    month_starts = [first_of_current - pd.DateOffset(months=i) for i in range(12, 0, -1)]
    month_starts.append(first_of_current)

    month_dfs = []
    for i in range(12):
        start = month_starts[i]
        end = month_starts[i+1]
        mask = (df_classified["target_date"] > start) & (df_classified["target_date"] <= end)
        month_df = df_classified[mask].copy()
        month_df["report_month"] = start.strftime("%b-%y")
        month_dfs.append(month_df)

    df_last_12 = pd.concat(month_dfs, ignore_index=True)

    summary = generate_monthly_summary(df_last_12)
    late_df = get_extreme_late_work_orders(df_last_12)
    return summary, by_group_df, trend_df, late_df

def main():
    import wx
    from gui.wx_app import WorkOrderDashboard

    def on_file_selected(file_path):
        # This should use the same logic as CLI
        summary, by_group_df, trend_df, late_df = prepare_data(file_path)
        # Pass these to your dashboard for display

    app = wx.App(False)
    dashboard = WorkOrderDashboard(None, title="Work Order Analysis Dashboard",  on_file_selected=on_file_selected)
    app.MainLoop()

if __name__ == "__main__":
    print("Running as __main__")
    if len(sys.argv) > 1:
        print("CLI mode detected")
        # CLI mode
        file_path = sys.argv[1]
        print("File path argument:", file_path)
        try:
            summary, by_group_df, trend_df, late_df = prepare_data(file_path)
            export_summary_to_excel(summary, late_df)

            # Rename columns for summary_df to match slide update expectations
            summary = summary.rename(columns={
                "due": "Due",
                "completed": "Completed",
                "missed": "Missed",
                "completion_pct": "Completion %",
                "canceled": "Canceled"
            })

            print("by_month_df_12:\n", trend_df)
            print("by_group_df:\n", by_group_df)
            by_group_df = by_group_df[by_group_df["missed"] > 0]
            create_full_governance_deck(summary, by_group_df, trend_df)
        except Exception as e:
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()
        input("Press Enter to exit...")
    else:
        # GUI mode
        main()


