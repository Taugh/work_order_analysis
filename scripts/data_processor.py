import pandas as pd
from scripts.data_loader import load_work_order_files
from scripts.classifier import apply_classification
from scripts.summary_generator import (
    generate_monthly_summary,
    get_extreme_late_work_orders,
    export_summary_to_excel,
)
from scripts.summary_generator import generate_12_month_trend

def prepare_data(file_path):
    # Load and classify data
    print("prepare_data called with", file_path)
    print("Starting prepare_data")  # Debug print
    df_cleaned = load_work_order_files(file_path)
    print("Loaded raw data shape:", df_cleaned.shape)
    df_classified = apply_classification(df_cleaned)
    cleaned_path = "data/processed/cleaned_work_orders.csv"
    df_classified.to_csv(cleaned_path, index=False)

    # Ensure target_date is datetime
    df_classified["target_date"] = pd.to_datetime(df_classified["target_date"], errors="coerce")

    # --- Build last 12 complete months using true date boundaries ---
    trend_df = generate_12_month_trend(df_classified)
    print("trend_df created")
    print(trend_df)
    print(trend_df["report_month"])

    # --- Use only the previous month for group charts ---
    today = pd.Timestamp.today()
    first_of_current = today.replace(day=1)
    first_of_previous = (first_of_current - pd.DateOffset(months=1)).replace(day=1)
    mask = (df_classified["target_date"] > first_of_previous) & (df_classified["target_date"] <= first_of_current)
    df_prev_month = df_classified[mask]

    by_group_df = (
        df_prev_month
        .groupby("group")
        .agg(
            missed=("wo_class", lambda x: (x == "missed").sum()),
            completed=("wo_class", lambda x: (x == "on_time").sum()),
            generated=("wo_class", "count"),
            # FIX: Use consistent column name
            missed_percentage=("wo_class", lambda x: 100 * (x == "missed").sum() / len(x) if len(x) else 0),
            still_open=("wo_class", lambda x: (x == "open").sum())
        )
        .reset_index()
    )

    # Build month boundaries
    month_starts = [first_of_current - pd.DateOffset(months=i) for i in range(12, 0, -1)]
    month_starts.append(first_of_current)

    # Build last 12 months DataFrame
    month_dfs = []
    for i in range(12):
        start = month_starts[i]
        end = month_starts[i+1]
        mask = (df_classified["target_date"] > start) & (df_classified["target_date"] <= end)
        month_df = df_classified[mask].copy()
        month_df["report_month"] = start.strftime("%b-%y")
        month_dfs.append(month_df)
    df_last_12 = pd.concat(month_dfs, ignore_index=True)

    summary = generate_monthly_summary(df_last_12)

    # FIX: Remove Grand Total from summary
    summary = summary[summary["Month"] != "Grand Total"].copy()
    
    # FIX: Generate disposition data
    disposition_df = generate_disposition_data(df_classified)
    
    late_df = get_extreme_late_work_orders(df_classified)
    # Add debugging:
    print(f"Late work orders found: {len(late_df)}")
    print("Late work orders preview:")
    print(late_df[['work_order', 'target_date', 'status', 'group']].head(15))

    # Now build PM Month and YTD after summary exists
    pm_month_label = first_of_previous.strftime("%b-%y")
    pm_month_df = summary[summary["Month"] == pm_month_label]

    current_year = today.year % 100
    summary_with_year = summary.copy()
    def safe_year_extract(month_str):
        try:
            if pd.isna(month_str) or len(str(month_str)) < 2:
                return None
            return int(str(month_str)[-2:])
        except (ValueError, TypeError):
            print(f"Warning: Could not extract year from '{month_str}'")
            return None

    summary_with_year["year"] = summary_with_year["Month"].apply(safe_year_extract)
    summary_with_year = summary_with_year.dropna(subset=["year"])
    ytd_df = summary_with_year[summary_with_year["year"] == current_year].drop(columns=["year"])

    return summary, by_group_df, trend_df, late_df, pm_month_df, ytd_df, df_classified, disposition_df 

def generate_disposition_data(df_classified):
    """
    Generate disposition data for missed work orders by month and current status
    """
    print("🔍 DEBUG: Starting disposition data generation...")
    
    # Filter to only missed work orders
    missed_df = df_classified[df_classified["wo_class"] == "missed"].copy()
    print(f"🔍 DEBUG: Found {len(missed_df)} missed work orders")
    
    if missed_df.empty:
        print("⚠️ No missed work orders found for disposition data")
        return pd.DataFrame()
    
    # FIX: Use the same month logic as the trend_df generation
    today = pd.Timestamp.today()
    first_of_current = today.replace(day=1)
    
    # Build the same 12 months as trend_df
    month_starts = []
    for i in range(12, 0, -1):
        month_start = first_of_current - pd.DateOffset(months=i)
        month_starts.append(month_start)
    
    # Filter missed_df to the same 12-month period
    start_date = month_starts[0]
    end_date = first_of_current
    missed_df = missed_df[
        (missed_df['target_date'] >= start_date) & 
        (missed_df['target_date'] < end_date)
    ].copy()
    
    print(f"🔍 DEBUG: After date filter ({start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}): {len(missed_df)} missed work orders")
    
    if missed_df.empty:
        print("⚠️ No missed work orders in the 12-month period")
        return pd.DataFrame()
    
    # Group current status into disposition categories
    def categorize_status(status):
        closed_statuses = ['CLOSE', 'REVWD', 'PENRVW', 'COMP', 'CORRTD']
        qa_statuses = ['PENDQA']
        dept_statuses = ['FLAGGED', 'MISSED', 'WAPPR', 'APPR', 'INPRG']
        
        if status in closed_statuses:
            return 'Closed'
        elif status in qa_statuses:
            return 'Awaiting QA'
        elif status in dept_statuses:
            return 'Awaiting Dept'
        else:
            return 'Awaiting Dept'  # Default
    
    missed_df['disposition'] = missed_df['status'].apply(categorize_status)
    
    # Debug: Show status distribution
    print("🔍 DEBUG: Status distribution:")
    print(missed_df['status'].value_counts())
    print("🔍 DEBUG: Disposition distribution:")
    print(missed_df['disposition'].value_counts())
    
    # FIX: Create report_month column using the same logic as trend generation
    def assign_report_month(target_date):
        for i, month_start in enumerate(month_starts):
            month_end = month_start + pd.DateOffset(months=1)
            if month_start <= target_date < month_end:
                return month_start.strftime("%b-%y")
        return None
    
    missed_df['report_month'] = missed_df['target_date'].apply(assign_report_month)
    
    # Remove rows where report_month is None
    missed_df = missed_df.dropna(subset=['report_month'])
    
    print(f"🔍 DEBUG: After report_month assignment: {len(missed_df)} missed work orders")
    print(f"🔍 DEBUG: Report months found: {sorted(missed_df['report_month'].unique())}")
    
    # Group by month and disposition
    disposition_summary = (
        missed_df.groupby(['report_month', 'disposition'])
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )
    
    print(f"🔍 DEBUG: disposition_summary shape: {disposition_summary.shape}")
    print(f"🔍 DEBUG: disposition_summary columns: {disposition_summary.columns.tolist()}")
    
    # Ensure all columns exist
    for col in ['Closed', 'Awaiting QA', 'Awaiting Dept']:
        if col not in disposition_summary.columns:
            disposition_summary[col] = 0
    
    # Rename columns to match slide_generator expectations
    disposition_summary = disposition_summary.rename(columns={
        'Closed': 'closed',
        'Awaiting QA': 'awaiting_qa', 
        'Awaiting Dept': 'awaiting_dept'
    })
    
    print(f"📊 Disposition data generated with {len(disposition_summary)} months")
    print(f"🔍 DEBUG: Final disposition_summary:")
    print(disposition_summary)
    
    return disposition_summary