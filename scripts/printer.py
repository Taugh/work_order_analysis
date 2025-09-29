# scripts/printer.py

import pandas as pd
import os
from config import REPORT_DIR

def print_centered_summary(df):
    columns = ["   Month", "Due", "Completed", "Missed", "Open", "Canceled", "Completion %"]
    widths = [20, 15, 15, 12, 15, 20, 20]

    header = " | ".join(f"{col:^{w}}" for col, w in zip(columns, widths))
    print("\n📊 Monthly Summary\n")
    print(header)
    print("-" * len(header))

    for _, row in df.iterrows():
        row_values = [
            str(row["Month"]),
            str(row["Due"]),
            str(row["Completed"]),
            str(row["Missed"]),
            str(row["Still Open"]),
            str(row["Canceled"]),
            f"{row['Completion %']:.2f}%"
        ]
        print(" | ".join(f"{val:^{w}}" for val, w in zip(row_values, widths)))

def export_summary_to_excel(summary_df, late_df, filename="monthly_summary.xlsx"):
        filepath = os.path.join(REPORT_DIR, filename)
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="Monthly Summary", index=False)
        print(f"✅ Exported summary and late WOs to Excel: {filepath}")
