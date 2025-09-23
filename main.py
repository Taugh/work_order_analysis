# main.py

import sys
import wx

from scripts.data_loader import load_work_order_files
from scripts.classifier import apply_classification
from scripts.summary_generator import (
    generate_monthly_summary,
    get_extreme_late_work_orders,
    export_summary_to_excel,
)
from scripts.printer import print_centered_summary
from scripts.slide_generator import create_full_governance_deck
from gui.wx_app import WorkOrderDashboard

def main():
    app = wx.App(False)
    dashboard = WorkOrderDashboard(None, title="Work Order Analysis Dashboard")
    app.MainLoop()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # CLI mode
        file_path = sys.argv[1]

        df_cleaned = load_work_order_files(file_path)

        # 💾 Save cleaned version before classification
        cleaned_path = "data/processed/cleaned_work_orders.csv"
        df_cleaned.to_csv(cleaned_path, index=False)

        # 🏷️ Apply classification
        df_classified = apply_classification(df_cleaned)
        print(df_classified.columns)     # After classification
        
        # 📊 Generate and export summaries
        summary = generate_monthly_summary(df_classified)
        by_group_df = (
            df_classified
            .groupby("group")
            .agg(
                missed=("wo_class", lambda x: (x == "missed").sum()),
                completed=("wo_class", lambda x: (x == "on_time").sum()),
                generated=("wo_class", "count")
            )
            .reset_index()
        )
        by_month_df = (
            df_classified
            .groupby("report_month")
            .agg(
                missed=("wo_class", lambda x: (x == "missed").sum()),
                completed=("wo_class", lambda x: (x == "on_time").sum()),
                generated=("wo_class", "count")
            )
            .reset_index()
        )
        late_df = get_extreme_late_work_orders(df_classified)
        export_summary_to_excel(summary, late_df)
        create_full_governance_deck(summary, by_group_df, by_month_df)

        # print_centered_summary(summary)
    else:
        # GUI mode
        main()
