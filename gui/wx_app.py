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



class WorkOrderDashboard(wx.Frame):
    def __init__(self, parent, title="Work Order Analysis Dashboard"):
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
        try:
            raw_df = load_work_order_files(file_path)
            self.df = apply_classification(raw_df)
            self.df["report_month"] = pd.to_datetime(self.df["target_date"], errors="coerce").dt.strftime("%b-%y")
            processed_path = os.path.join("data", "processed", "cleaned_work_orders.csv")
            self.df.to_csv(processed_path, index=False)
            self.status_text.SetLabel(f"✅ Loaded: {os.path.basename(file_path)}")
        except Exception as err:
            self.status_text.SetLabel(f"❌ Load error: {err}")
            self.df = None
        
    def on_generate_report(self, event):
        if self.df is None:
            self.status_text.SetLabel("⚠️ Please load a file first.")
            return

        try:
            # Generate main summary
            report_choice = self.report_type.GetValue()

            
            if report_choice == "Monthly Summary":
                summary_df = generate_monthly_summary(self.df)
                export_summary_to_excel(summary_df, None, filename="monthly_summary.xlsx")
            elif report_choice == "Governance Overview":
                gov_data = generate_governance_overview(self.df)
                print("📄 Final columns before breakdowns:", self.df.columns.tolist())

                print("gov_data['summary'] columns before rename:", gov_data["summary"].columns)

                # Rename columns for summary_df to match slide update expectations
                gov_data["summary"] = gov_data["summary"].rename(columns={
                    "report_month": "Month",
                    "due": "Due",
                    "completed": "Completed",
                    "missed": "Missed",
                    "completion_pct": "Completion %",
                    "canceled": "Canceled"
                })

                pm_breakdowns = generate_pm_breakdowns(self.df)

                # --- Filter last 12 months for PM Missed Chart ---
                by_month_df = pm_breakdowns["by_month"].copy()
                # Convert to datetime for sorting
                by_month_df["report_month_dt"] = pd.to_datetime(by_month_df["report_month"], format="%b-%y")
                by_month_df = by_month_df.sort_values("report_month_dt")

                # Drop duplicate months (if any)
                by_month_df = by_month_df.drop_duplicates(subset=["report_month"])

                # Select the last 12 unique months
                by_month_df_12 = by_month_df.iloc[-12:].drop(columns=["report_month_dt"])

                # --- Filter summary for current year for YTD ---
                summary_df = gov_data["summary"].copy()
                summary_df["year"] = summary_df["Month"].str[-2:].astype(int)
                current_year = datetime.now().year % 100
                ytd_df = summary_df[summary_df["year"] == current_year].drop(columns=["year"])

                # Assume your original DataFrame is called df_classified or similar
                newest_month = self.df["report_month"].max()
                by_group_df = self.df[self.df["report_month"] == newest_month].groupby("group").agg(
                    missed=("wo_class", lambda x: (x == "missed").sum()),
                    completed=("wo_class", lambda x: (x == "on_time").sum()),
                    generated=("wo_class", "count"),
                    missed_percent=("wo_class", lambda x: 100 * (x == "missed").sum() / len(x)),
                    still_open=("wo_class", lambda x: (x == "open").sum())
                ).reset_index()

                # Ignore groups where missed == 0
                by_group_df = by_group_df[by_group_df["missed"] > 0]

                # Pass filtered DataFrames to the deck function
                create_full_governance_deck(
                    summary_df=summary_df,
                    by_group_df=by_group_df,  # <-- now only the latest month
                    by_month_df=by_month_df_12
                )

                export_governance_report(gov_data, filename="governance_overview.xlsx")
                ##create_governance_slide(gov_data["summary"])

                # ✅ Generate slide from the summary sheet
                #create_governance_slide(gov_data["summary"])

            else:
                raise ValueError(f"Unknown report: {report_choice}")

            self.status_text.SetLabel(f"📊 {report_choice} exported successfully.")

        except Exception as err:
            self.status_text.SetLabel(f"❌ Error: {err}")

    def on_open_folder(self, event):
        target = os.path.abspath("outputs")
        if not os.path.isdir(target):
            self.status_text.SetLabel("⚠️ Output folder not found.")

            return
        os.startfile(target)





















