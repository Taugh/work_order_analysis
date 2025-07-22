# scripts/summary_generator.py

import os
import pandas as pd
from datetime import date
from config.settings import REPORT_DIR

def generate_monthly_summary(df):
     # Ensure wo_class and report_month columns exist
     # wo_class is the classification each work order is placed into (canceled, completed, missed, or open)

    if "report_month" not in df.columns or "wo_class" not in df.columns:  # wo_class is the
        # classification each work order is placed in
        raise ValueError("Missing 'report_month' or 'wo_class'. Run classfier first.")

    summary = df.groupby("report_month")["wo_class"].value_counts().unstack(fill_value=0)
    summary["total_due"] = summary.sum(axis=1)
    summary["completion_pct"] = (summary.get("on_time", 0) / summary["total_due"]) * 100

    # Clean and reorder columns

    summary = summary.reset_index().rename(columns={
            "report_month": "Month",
            "on_time": "Completed",
            "missed": "Missed",
            "open": "Still Open",
            "canceled": "Canceled",
            "completion_pct": "Completion %",
            "total_due": "Due",
            "status": "Status"
         })

    columns_to_sum = ["Completed", "Missed", "Still Open", "Due", "Canceled"]
    summary.loc["Grand Total", columns_to_sum] = summary[columns_to_sum].dropna(axis=1, how="all").sum()
    summary.loc["Grand Total", "Completion %"] = (
        summary.loc["Grand Total", "Completed"] / summary.loc["Grand Total", "Due"]
    ) * 100

    # Round percentage
    summary.loc["Grand Total", "Completion %"] = round(summary.loc["Grand Total", "Completion %"], 2)

    # Add a grand total row
    summary["Month"] = summary["Month"].astype(str)
    summary.loc["Grand Total", "Month"] = "Grand Total"

    return summary

def export_summary_to_excel(summary_df, late_df, filename="monthly_summary.xlsx"):
    filepath = os.path.join(REPORT_DIR, filename)

    # 1) Try to delete an existing file first

    if os.path.exists(filepath):
        try:
            os.remove(filepath)
        except PermissionError:
            raise PermissionError(
                    f"Cannot overwrite '{filename}'. Close it in Excel and try again."
            )

    # 2) Write the new report
    
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Monthly Summary", index=False)
        if late_df is not None:
            late_df.to_excel(writer, sheet_name="Late >90 Days", index=False)

    return filepath
            
    # Use for debugging
    #print(f"✅ Exported summary and late WOs to Excel: {filepath}")

def get_extreme_late_work_orders(df, days_late=90):
    today = pd.Timestamp.today()
    late_df = df[df["status"].isin(["APPR", "INPRG", "WAPPR"])].copy()
    late_df["late_days"] = (today - late_df["target_date"]).dt.days
    late_df = late_df.query("late_days > @days_late").copy()

    late_df["report_month"] = late_df["target_date"].dt.to_period("M").astype(str)
    late_df = late_df[
        [
            "report_month",
            "work_order",
            "wo_assigned_group",
            "target_date",
            "late_days",
            "description",
            "wo_class",
            "status"
        ]
    ].sort_values(["report_month", "wo_assigned_group", "late_days"], ascending=[True, True, False])

    return late_df

def generate_governance_overview(df):
    # Normalize wo_class just in case
    df["wo_class"] = df["wo_class"].str.strip().str.lower()

    total_due = len(df)
    total_completed = len(df[df["wo_class"] == "on_time"])
    total_missed = len(df[df["wo_class"] == "missed"])
    canceled = len(df[df["wo_class"] ==  "canceled"])
    completion_pct = round((total_completed / total_due) * 100, 1) if total_due > 0 else 0

    print("🔍 Total work orders:", total_due)
    print("🔍 wo_class breakdown:\n", df["wo_class"].value_counts())

    summary_df = pd.DataFrame([{
        "due": total_due,
        "completed": total_completed,
        "missed": total_missed,
        "canceled": canceled,
        "completion_pct": completion_pct
    }])

    return {
        "summary": summary_df
    }

def export_governance_report(data_dict, filename="governance_overview.xlsx"):
    output_dir = "outputs/reports"
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, filename)

    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        data_dict["summary"].to_excel(writer, sheet_name="PM Totals", index=False)
        # You can add more sheets later like:
        # data_dict["by_group"].to_excel(writer, sheet_name="PM by Group", index=False)

    print(f"✅ Governance report saved to: {filepath}")

def generate_pm_breakdowns(df):
    df["wo_class"] = df["wo_class"].str.strip().str.lower()
    df["group"] = df["group"].fillna("Unassigned")
    df["target_date"] = pd.to_datetime(df["target_date"], errors="coerce")

    # Format for display and sorting
    df["report_month"] = df["target_date"].dt.strftime("%b-%y")
    df["month_sort"] = df["target_date"].dt.to_period("M").dt.to_timestamp()

    # Count missed, completed, and total (generated) per month
    monthly_counts = (
        df.groupby(["report_month", "month_sort", "wo_class"])
        .size()
        .unstack(fill_value=0)
        .reset_index()
        .sort_values("month_sort")
    )

    # Add total generated (sum of all classes)
    monthly_counts["generated"] = monthly_counts.sum(axis=1, numeric_only=True)

    # Rename columns for consistency
    monthly_counts = monthly_counts.rename(columns={
        "missed": "missed",
        "on_time": "completed"
    })

    # Ensure all expected columns exist
    for col in ["missed", "completed"]:
        if col not in monthly_counts:
            monthly_counts[col] = 0

    # Final output
    by_month = monthly_counts[["report_month", "missed", "completed", "generated"]]

    # Missed by group
    by_group = df[df["wo_class"] == "missed"].groupby("group").size().reset_index(name="missed")

    return {
        "by_group": by_group,
        "by_month": by_month
    }


def generate_monthly_governance_overview(df):
    df["wo_class"] = df["wo_class"].str.strip().str.lower()
    df["target_date"] = pd.to_datetime(df["target_date"], errors="coerce")
    df["report_month"] = df["target_date"].dt.strftime("%b-%y")

    def summarize(group):
        total_due = len(group)
        total_completed = len(group[group["wo_class"] == "on_time"])
        total_missed = len(group[group["wo_class"] == "missed"])
        completion_pct = round((total_completed / total_due) * 100, 1) if total_due > 0 else 0
        return pd.Series({
            "due": total_due,
            "completed": total_completed,
            "missed": total_missed,
            "completion_pct": completion_pct
        })

    monthly_summary = df.groupby("report_month").apply(summarize).reset_index()
    return monthly_summary.sort_values("report_month")

