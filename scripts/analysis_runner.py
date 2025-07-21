# scripts/analysis_runner.py

import pandas as pd

def load_cleaned_data(filepath="data/processed/cleaned_work_orders.csv"):
    """Load pre-cleaned work order data."""
    return pd.read_csv(filepath)

def generate_summary(df):
    """Return key metrics for reporting."""
    return {
        "total_orders": len(df),
        "by_type": df["OrderType"].value_counts().to_dict(),
        "monthly_trend": df.groupby(df["Date"].str[:7]).size().to_dict()  # Assuming 'Date' is yyyy-mm-dd
    }

def run_analysis(filepath="data/processed/cleaned_work_orders.csv"):
    df = load_cleaned_data(filepath)
    summary = generate_summary(df)
    return summary

if __name__ == "__main__":
    metrics = run_analysis()
    for k, v in metrics.items():
        print(f"{k}: {v}")