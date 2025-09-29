# ---------------------------------------------------------------
# analysis_runner.py
#
# Purpose:
#   Runs work order analysis from cleaned data and provides summary metrics.
#
# Requirements:
#   - Input: 'cleaned_work_orders.csv' in 'data/processed' directory.
#   - Columns: Must include 'target_date' (date), 'OrderType', and other relevant fields.
#
# Output:
#   - Prints summary metrics to console.
#   - (Future) Can export summary to file via export_summary().
#   - Used in CLI mode with options for summary or governance analysis.
# ---------------------------------------------------------------

import argparse
import logging
import os
import pandas as pd
from datetime import datetime, timedelta

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)


def load_cleaned_data(filepath="data/processed/cleaned_work_orders.csv"):
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"🚨 Missing file: {filepath}")
    return pd.read_csv(filepath)

def generate_summary(df):
    # Ensure target_date is datetime
    df["target_date"] = pd.to_datetime(df["target_date"], errors="coerce")

    # Use the current calendar month, not the latest date in your data
    today = pd.Timestamp.today()
    first_of_current = today.replace(day=1)
    first_of_previous = (first_of_current - pd.DateOffset(months=1)).replace(day=1)

    # Only include work orders due between first_of_previous (exclusive) and first_of_current (inclusive)
    due_for_month = df[(df["target_date"] > first_of_previous) & (df["target_date"] <= first_of_current)].shape[0]

    return {
        "total_orders": len(df),
        "by_type": df["OrderType"].value_counts().to_dict() if "OrderType" in df.columns else {},
        "monthly_trend": df.groupby(df["target_date"].dt.to_period("M")).size().to_dict(),
        "due_for_month": due_for_month
    }

def run_analysis(filepath="data/processed/cleaned_work_orders.csv", dry_run=False):
    df = load_cleaned_data(filepath)
    summary = generate_summary(df)
    if dry_run:
        logging.info("🧪 Dry run mode: Summary generated, no files exported")

    return summary

def export_summary(summary, output_path):
    pass  # placeholder for future logic

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run work order analysis")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Run analysis without exporting files"
    )
    parser.add_argument("--mode", choices=["summary", "governance"], default="summary")
    args = parser.parse_args()  # ← needs to come before you use args

    if args.mode == "summary":
        metrics = run_analysis(filepath="data/processed/cleaned_work_orders.csv", dry_run=args.dry_run)
        # Later: export_summary(metrics, output_path)
        for k, v in metrics.items():
            print(f"{k}: {v}")
    elif args.mode == "governance":
        logging.info("🧪 Governance mode not implemented yet.")

