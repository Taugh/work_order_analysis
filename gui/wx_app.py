# ---------------------------------------------------------------
# wx_app.py
#
# Purpose:
#   Provides a graphical user interface (GUI) for work order analysis and reporting.
#
# Requirements:
#   - wxPython library for GUI components.
#   - Input: Excel (.xlsx) work order file selected by the user.
#   - Functions from scripts: data_loader, classifier, summary_generator, slide_generator.
#
# Output:
#   - Loads, classifies, and analyzes work order data.
#   - Exports summary and governance reports to Excel and PowerPoint in the outputs folder.
#   - Displays status and feedback to the user.
#
# Notes:
#   - Users can select a file, choose report type, and generate/export reports.
#   - Output files are saved in the 'outputs' directory.
#   - GUI provides options for monthly summary and governance overview.
# ---------------------------------------------------------------

# gui/wx_app.py

import os
import wx
import pandas as pd
import threading
from datetime import datetime
from scripts.data_loader import  load_work_order_files
from scripts.summary_generator import (
    generate_monthly_summary,
    get_extreme_late_work_orders,
    export_summary_to_excel,
    generate_governance_overview,
    export_governance_report,
    generate_pm_breakdowns
)
from scripts.classifier import apply_classification
from scripts.slide_generator import (
    create_governance_slide,
    create_full_governance_deck
)
from main import prepare_data



class WorkOrderDashboard(wx.Frame):
    def __init__(self, parent, title="Work Order Analysis Dashboard", on_file_selected=None):
        super().__init__(parent, title=title, size=wx.Size(800, 500))
        self.panel = wx.Panel(self)
        self.df = None  # will hold the loaded DataFrame

        self.init_ui()
        self.Center()
        self.Show()

    def  init_ui(self):
        # Layout components will go here


        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # File Picker Section

        file_label = wx.StaticText(self.panel, label="📁 Select Work Order File:")
        self.file_picker = wx.FilePickerCtrl(self.panel, message="Choose Exel file ... ",
                                             wildcard="*.xlsx")

        # Report Options Section
        report_label = wx.StaticText(self.panel, label="🛠️ Report Options:")
        self.report_type = wx.ComboBox(self.panel, choices=["Monthly Summary",
                                                            "Governance Overview"], style=wx.CB_READONLY)
        self.include_late = wx.CheckBox(self.panel, label="Include Open > 90 Days")

        # Action Buttons


        self.generate_btn = wx.Button(self.panel, label="Generate Report")
        self.open_folder_btn =  wx.Button(self.panel, label="Open Output Folder")
        
        # Feedback Text

        self.status_text = wx.StaticText(self.panel, label="Ready to generate your summary.")
        
        # Assemble layout

        main_sizer.Add(file_label, 0, wx.ALL, 5)
        main_sizer.Add(self.file_picker, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(report_label, 0, wx.ALL, 5)
        main_sizer.Add(self.report_type, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(self.include_late, 0, wx.ALL, 5)
        main_sizer.Add(self.generate_btn, 0, wx.ALL | wx.ALIGN_LEFT, 10)
        main_sizer.Add(self.open_folder_btn, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        main_sizer.Add(self.status_text, 0, wx.ALL, 10)

        self.file_picker.Bind(wx.EVT_FILEPICKER_CHANGED, self.on_file_selected)
        self.generate_btn.Bind(wx.EVT_BUTTON, self.on_generate_report)
        self.open_folder_btn.Bind(wx.EVT_BUTTON, self.on_open_folder)

        self.panel.SetSizer(main_sizer)

    def on_file_selected(self, event):
        file_path = self.file_picker.GetPath()
        self.status_text.SetLabel("⏳ Loading file, please wait...")

        def load_file():
            try:
                raw_df = load_work_order_files(file_path)
                df = apply_classification(raw_df)
                df["report_month"] = pd.to_datetime(df["target_date"], errors="coerce").dt.strftime("%b-%y")
                processed_path = os.path.join("data", "processed", "cleaned_work_orders.csv")
                df.to_csv(processed_path, index=False)
                # Update self.df and status text on the GUI thread
                def update_ui():
                    self.df = df
                    self.status_text.SetLabel(f"✅ Loaded: {os.path.basename(file_path)}")
                wx.CallAfter(update_ui)
            except Exception as err:
                wx.CallAfter(self.status_text.SetLabel, f"❌ Load error: {err}")
                wx.CallAfter(setattr, self, "df", None)

        threading.Thread(target=load_file).start()
        
    def on_generate_report(self, event):
        if self.df is None:
            self.status_text.SetLabel("⚠️ Please load a file first.")
            return

        def run_report():
            try:
                file_path = self.file_picker.GetPath()
                # FIX: Add df_classified to the return values
                summary, by_group_df, trend_df, late_df, df1, df2, df_classified = prepare_data(file_path)

                include_late_orders = self.include_late.GetValue()
                
                report_choice = self.report_type.GetValue()
                if report_choice == "Monthly Summary":
                    # Include late orders in Excel if checkbox is checked
                    late_data = late_df if include_late_orders else None
                    export_summary_to_excel(summary, late_data, filename="monthly_summary.xlsx")
                elif report_choice == "Governance Overview":
                    by_group_df = by_group_df[by_group_df["missed"] > 0]
                    
                    # Pass late_df when checkbox is checked
                    late_data = late_df if include_late_orders else None
                    # FIX: Add df_classified parameter to the function call
                    create_full_governance_deck(summary, by_group_df, trend_df, df_classified, late_data)
                                    
                    # If checkbox is checked, could add late orders to a separate slide
                    if include_late_orders and not late_df.empty:
                        # Could extend governance deck with late orders slide
                        pass
                else:
                    raise ValueError(f"Unknown report: {report_choice}")

                wx.CallAfter(self.status_text.SetLabel, f"📊 {report_choice} exported successfully.")
            except Exception as err:
                print(f"Full error: {err}")
                wx.CallAfter(self.status_text.SetLabel, f"❌ Error: {err}")

        threading.Thread(target=run_report).start()

    def on_open_folder(self, event):
        target = os.path.abspath("outputs")
        if not os.path.isdir(target):
            self.status_text.SetLabel("⚠️ Output folder not found.")

            return
        os.startfile(target)





















