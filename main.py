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
import wx

# Import prepare_data from the new module
from scripts.data_processor import prepare_data
from scripts.summary_generator import export_summary_to_excel
from scripts.slide_generator import create_full_governance_deck

print("main.py started")

def main():
    import wx
    from gui.wx_app import WorkOrderDashboard

    def on_file_selected(file_path):
        # FIX: Unpack all 8 return values including disposition_df
        summary, by_group_df, trend_df, late_df, pm_month_df, ytd_df, df_classified, disposition_df = prepare_data(file_path)
        # Pass these to your dashboard for display

    app = wx.App(False)
    dashboard = WorkOrderDashboard(None, title="Work Order Analysis Dashboard",  on_file_selected=on_file_selected)
    app.MainLoop()

if __name__ == "__main__":
    print("Running as __main__")
    if len(sys.argv) > 1:
        print("CLI mode detected")
        file_path = sys.argv[1]
        print("File path argument:", file_path)
        try:
            # FIX: Unpack all 8 return values including disposition_df
            summary, by_group_df, trend_df, late_df, pm_month_df, ytd_df, df_classified, disposition_df = prepare_data(file_path)
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
            print("disposition_df:\n", disposition_df)
            
            # Use the grouped by_group_df that was created in prepare_data
            filtered_by_group_df = by_group_df[by_group_df["missed"] > 0]
            
            # FIX: Pass the correct parameters including disposition_df
            create_full_governance_deck(trend_df, late_df, disposition_df, filtered_by_group_df, filename=None)
        except Exception as e:
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()
        input("Press Enter to exit...")
    else:
        # GUI mode
        main()


