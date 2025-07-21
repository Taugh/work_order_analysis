# gui/wx_app.py

import os
import wx
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
        super().__init__(parent, title=title, size=(800, 500))
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
            # 💾 Save cleaned data before classification
            processed_path = os.path.join("data", "processed", "cleaned_work_orders.csv")
            raw_df.to_csv(processed_path, index=False)
            
            self.df = apply_classification(raw_df)
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

                pm_breakdowns = generate_pm_breakdowns(self.df)

                create_full_governance_deck(
                    summary_df=gov_data["summary"],
                    by_group_df=pm_breakdowns["by_group"],
                    by_month_df=pm_breakdowns["by_month"]
                )


                export_governance_report(gov_data, filename="governance_overview.xlsx")
                create_governance_slide(gov_data["summary"])

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

                                      


















        
