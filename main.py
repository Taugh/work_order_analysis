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
from gui.wx_app import WorkOrderDashboard

def main():
    app = wx.App(False)
    dashboard = WorkOrderDashboard(None, title="Work Order Analysis Dashboard")
    app.MainLoop()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # CLI mode
        file_path = sys.argv[1]
        df = load_work_order_files(file_path)
        df = apply_classification(df)
        summary = generate_monthly_summary(df)
        late_df = get_extreme_late_work_orders(df)
        export_summary_to_excel(summary, late_df)
        # print_centered_summary(summary)
    else:
        # GUI mode
        main()
