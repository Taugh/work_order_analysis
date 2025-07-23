# scripts/analysis_runner.py

import argparse
import logging
import os
import pandas as pd

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
    # Generates a summary of work orders
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")  # Ensure it's datetime

    return {
        "total_orders": len(df),
        "by_type": df["OrderType"].value_counts().to_dict(),
        "monthly_trend": df.groupby(df["Date"].dt.to_period("M")).size().to_dict()
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
    if args.mode == "summary":
        metrics = run_analysis(filepath="data/processed/cleaned_work_orders.csv", dry_run=args.dry_run)
        # Later: export_summary(metrics, output_path)
    elif args.mode == "governance":
        logging.info("🧪 Governance mode not implemented yet.")

    args = parser.parse_args()

    metrics = run_analysis(filepath="data/processed/cleaned_work_orders.csv", dry_run=args.dry_run)

    for k, v in metrics.items():
        print(f"{k}: {v}")