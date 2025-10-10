# ---------------------------------------------------------------
# slide_generator.py
#
# Purpose:
#   Generates and updates PowerPoint slides for work order governance reporting.
#
# Requirements:
#   - Input: pandas DataFrames containing summary, group, and monthly work order data.
#   - PowerPoint template: 'data/templates/governance_slide_template.pptx' with named charts/shapes.
#   - DataFrames must include columns such as 'Month', 'Due', 'Completed', 'Missed', 'group', etc.
#
# Output:
#   - Creates and updates slides with summary metrics, charts, and tables.
#   - Saves the final presentation to 'outputs/presentations/governance_slide.pptx' (or specified path).
#
# Notes:
#   - Functions are provided for both creating new slides and updating existing charts/tables.
#   - All chart and table updates require matching shape/chart names in the PowerPoint template.
#   - Designed for use in both CLI and GUI workflows.
# ---------------------------------------------------------------

# scripts/slide_generator.py

import os
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from datetime import datetime
import pandas as pd


def create_governance_slide(summary_df, prs):
    print("create_governance_slide called")
    print("summary_df columns:", summary_df.columns)

    # Get most current month (last row after sorting by Month)
    summary_df_sorted = summary_df.sort_values("Month", key=lambda x: pd.to_datetime(x, format="%b-%y"), ignore_index=True)
    current_month_row = summary_df_sorted.iloc[-1]
    # Year-to-date summary (sum all months)
    ytd_row = summary_df_sorted.drop(columns=["Month"]).sum(numeric_only=True)
    completion_pct = 100 * ytd_row["Completed"] / ytd_row["Due"] if ytd_row["Due"] else 0

    # Add a new slide
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only layout
    slide.shapes.title.text = "Preventive Maintenance Summary"

    # Create a rounded rectangle for the summary text
    left = Inches(1)
    top = Inches(1.5)
    width = Inches(8)
    height = Inches(2.5)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)

    # Format the shape
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(0, 0, 0)

    # Add formatted summary text
    text_frame = shape.text_frame
    text_frame.clear()
    p = text_frame.paragraphs[0]
    p.text = (
        f"Current Month ({current_month_row['Month']}):\n"
        f"  Due: {current_month_row['Due']:.0f}\n"
        f"  Completed: {current_month_row['Completed']:.0f}\n"
        f"  Missed: {current_month_row['Missed']:.0f}\n"
        f"  Completion %: {current_month_row['Completion %']:.1f}\n\n"
        f"Year to Date:\n"
        f"  Due: {ytd_row['Due']:.0f}\n"
        f"  Completed: {ytd_row['Completed']:.0f}\n"
        f"  Missed: {ytd_row['Missed']:.0f}\n"
        f"  Completion %: {completion_pct:.1f}"
    )
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)

    return prs


def create_missed_by_month_slide(prs, by_month_df):
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Preventive Maintenance Missed"

    chart_data = CategoryChartData()
    chart_data.categories = by_month_df["report_month"].astype(str).tolist()
    chart_data.add_series("Missed", by_month_df["missed"].tolist())
    chart_data.add_series("Completed", by_month_df["completed"].tolist())
    chart_data.add_series("Generated", by_month_df["generated"].tolist())

    x, y, cx, cy = Inches(1), Inches(1.5), Inches(8), Inches(4.5)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    ).chart

    # Try to force series 1 and 2 to be lines
    chart.series[1].chart_type = XL_CHART_TYPE.LINE
    chart.series[2].chart_type = XL_CHART_TYPE.LINE

    chart.has_legend = True
    chart.legend.position = 2  # Top
    chart.category_axis.has_major_gridlines = False
    chart.value_axis.minimum_scale = 0

def create_missed_by_group_slide(prs, by_group_df):
    slide_layout = prs.slide_layouts[5]  # Title only
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Preventive Maintenance Missed by Group"

    rows, cols = by_group_df.shape
    left, top, width, height = Inches(1), Inches(1.5), Inches(8), Inches(5)
    table = slide.shapes.add_table(rows + 1, 2, left, top, width, height).table

    # Set headers
    table.cell(0, 0).text = "Group"
    table.cell(0, 1).text = "Missed"

    # Fill table rows
    for i, row in by_group_df.iterrows():
        table.cell(i + 1, 0).text = str(row["group"])
        table.cell(i + 1, 1).text = str(row["missed"])

    # Optional: format font size
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(12)
    
def update_governance_slide(summary_df, prs, slide_index=0):
    slide = prs.slides[slide_index]
    slide_width = prs.slide_width.inches if hasattr(prs.slide_width, 'inches') else 10
    slide_height = prs.slide_height.inches if hasattr(prs.slide_height, 'inches') else 7.5

    # Set a complementary background color (light blue)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(220, 235, 250)

    # --- FIX: Filter out "Grand Total" before sorting ---
    summary_df_no_total = summary_df[summary_df["Month"] != "Grand Total"].copy()  # Add .copy() to fix warning
    summary_df_sorted = summary_df_no_total.sort_values(
        "Month", key=lambda x: pd.to_datetime(x, format="%b-%y"), ignore_index=True
    )
    # Optionally, add the Grand Total row back if needed
    grand_total_row = summary_df[summary_df["Month"] == "Grand Total"]
    if not grand_total_row.empty:
        summary_df_sorted = pd.concat([summary_df_sorted, grand_total_row], ignore_index=True)

    today = pd.Timestamp.today()
    first_of_current = today.replace(day=1)
    first_of_previous = (first_of_current - pd.DateOffset(months=1)).replace(day=1)
    current_month_label = first_of_previous.strftime("%b-%y")
    current_month_row = summary_df_sorted[summary_df_sorted["Month"] == current_month_label].iloc[0]

    # YTD: filter for current year only
    summary_df_no_total = summary_df_sorted[summary_df_sorted["Month"] != "Grand Total"].copy()  # Add .copy() to fix warning

    def safe_year_extract(month_str):
        try:
            if pd.isna(month_str) or len(str(month_str)) < 2:
                return None
            return int(str(month_str)[-2:])
        except (ValueError, TypeError):
            print(f"Warning: Could not extract year from '{month_str}'")
            return None

    summary_df_no_total["year"] = summary_df_no_total["Month"].apply(safe_year_extract)
    summary_df_no_total = summary_df_no_total.dropna(subset=["year"])
    
    current_year = datetime.now().year % 100
    ytd_df = summary_df_no_total[summary_df_no_total["year"] == current_year]
    ytd_row = ytd_df.drop(columns=["Month", "year"]).sum(numeric_only=True)
    completion_pct = 100 * ytd_row["Completed"] / ytd_row["Due"] if ytd_row["Due"] else 0

    # Shape sizes (50% larger)
    shape_width = 4.5  # 3 * 1.5
    shape_height = 3.75  # 2.5 * 1.5

    # Font size (20% larger)
    font_size = int(18 * 1.5)  # 22

    # Vertically center the shapes
    top = (slide_height - shape_height) / 2

    # Horizontally: center in left and right halves
    left_left = (slide_width / 4) - (shape_width / 2)
    left_right = (3 * slide_width / 4) - (shape_width / 2)

    # Left shape: Current Month
    shape_left = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_left), Inches(top), Inches(shape_width), Inches(shape_height))
    shape_left.fill.solid()
    shape_left.fill.fore_color.rgb = RGBColor(230, 240, 255)
    shape_left.line.color.rgb = RGBColor(0, 0, 0)
    text_frame_left = shape_left.text_frame
    text_frame_left.clear()
    p_left = text_frame_left.paragraphs[0]
    p_left.text = (
        f"Current Month ({current_month_row['Month']}):\n"
        f"  Due: {current_month_row['Due']:.0f}\n"
        f"  Completed: {current_month_row['Completed']:.0f}\n"
        f"  Missed: {current_month_row['Missed']:.0f}\n"
        f"  Completion %: {current_month_row['Completion %']:.1f}"
    )
    p_left.font.size = Pt(font_size)
    p_left.font.bold = True
    p_left.font.color.rgb = RGBColor(0, 0, 80)

    # Right shape: Year to Date
    shape_right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_right), Inches(top), Inches(shape_width), Inches(shape_height))
    shape_right.fill.solid()
    shape_right.fill.fore_color.rgb = RGBColor(255, 245, 230)
    shape_right.line.color.rgb = RGBColor(0, 0, 0)
    text_frame_right = shape_right.text_frame
    text_frame_right.clear()
    p_right = text_frame_right.paragraphs[0]
    p_right.text = (
        f"Year to Date:\n"
        f"  Due: {ytd_row['Due']:.0f}\n"
        f"  Completed: {ytd_row['Completed']:.0f}\n"
        f"  Missed: {ytd_row['Missed']:.0f}\n"
        f"  Completion %: {completion_pct:.1f}"
    )
    p_right.font.size = Pt(font_size)
    p_right.font.bold = True
    p_right.font.color.rgb = RGBColor(80, 40, 0)

def update_missed_by_month_chart(prs, by_month_df, slide_index=1):
    slide = prs.slides[slide_index]
    
    # Update existing chart
    chart_updated = False
    for shape in slide.shapes:
        if hasattr(shape, "chart"):
            chart_data = CategoryChartData()
            chart_data.categories = by_month_df["report_month"].astype(str).tolist()
            chart_data.add_series("Missed", by_month_df["missed"].tolist())
            chart_data.add_series("Completed", by_month_df["completed"].tolist())
            chart_data.add_series("Generated", by_month_df["generated"].tolist())
            shape.chart.replace_data(chart_data)
            chart_updated = True
            print("✅ Missed by Month chart updated")
            break
    
    if not chart_updated:
        print("⚠️ No chart found on slide to update")
    
    # Calculate totals
    total_due = by_month_df['generated'].sum()
    total_completed = by_month_df['completed'].sum()  
    total_missed = by_month_df['missed'].sum()
    
    print(f"📊 Rolling totals calculated:")
    print(f"  Total Due: {total_due}")
    print(f"  Total Completed: {total_completed}")
    print(f"  Total Missed: {total_missed}")
    
    # Find and update the existing text box with "Rolling 12-Month" text
    text_box_found = False
    for shape in slide.shapes:
        try:
            if hasattr(shape, 'text_frame') and shape.text_frame and hasattr(shape.text_frame, 'text'):
                if 'Rolling 12-Month' in shape.text_frame.text:
                    print("🎯 Found existing Rolling 12-Month text box - updating it")
                    print(f"📍 Template text box position: left={shape.left}, top={shape.top}, width={shape.width}, height={shape.height}")
                    
                    text_frame = shape.text_frame
                    text_frame.clear()
                    
                    # Add title
                    p = text_frame.paragraphs[0]
                    p.text = "12-Month Totals:"
                    p.font.size = Pt(20)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0, 0, 0)
                    
                    # Add spacing paragraph
                    # p = text_frame.add_paragraph()
                    # p.text = ""
                    
                    # Add Due
                    p = text_frame.add_paragraph()
                    p.text = f"Due: {total_due:,}"
                    p.font.size = Pt(18)
                    p.font.color.rgb = RGBColor(79, 98, 40)  # Dark Green
                    
                    # Add Completed  
                    p = text_frame.add_paragraph()
                    p.text = f"Completed: {total_completed:,}"
                    p.font.size = Pt(18)
                    p.font.color.rgb = RGBColor(149, 55, 53)  # Dark Red
                    
                    # Add Missed
                    p = text_frame.add_paragraph()
                    p.text = f"Missed: {total_missed:,}"
                    p.font.size = Pt(16)
                    p.font.color.rgb = RGBColor(55, 96, 146)  # Dark Blue
                    
                    text_box_found = True
                    print("✅ Rolling 12-month totals updated in existing text box")
                    print(f"✅ Final text frame content: '{text_frame.text}'")
                    break
        except Exception as e:
            print(f"⚠️ Error checking shape: {e}")
            continue
    
    if not text_box_found:
        print("❌ Could not find existing 'Rolling 12-Month' text box in template")
        print("💡 Make sure your template has a text box containing 'Rolling 12-Month' text")

def update_missed_disposition_chart(prs, df_classified, slide_index=2):
    """
    Creates a stacked bar chart showing the disposition of missed work orders over 12 months.
    
    Logic:
    - Use target_date to assign missed work orders to months (when they were due)
    - Use status to categorize work orders 
    - Use completion_status = 'MISSED' to identify missed work orders
    """
    slide = prs.slides[slide_index]
    
    print("📊 Building missed work order disposition data...")
    
    # Add safety check for df_classified
    if df_classified is None:
        print("❌ Error: df_classified is None - cannot build disposition chart")
        return None
    
    if df_classified.empty:
        print("❌ Error: df_classified is empty - cannot build disposition chart")
        return None
    
    # Check required columns exist
    required_columns = ['target_date', 'status', 'completion_status']
    missing_columns = [col for col in required_columns if col not in df_classified.columns]
    if missing_columns:
        print(f"❌ Error: Missing required columns in df_classified: {missing_columns}")
        print(f"Available columns: {df_classified.columns.tolist()}")
        return None
    
    print(f"📅 Using 'target_date' to assign work orders to months (when they were due)")
    print(f"🏷️ Using 'status' to categorize work orders")
    print(f"🔍 Using 'completion_status' = 'MISSED' to identify missed work orders")
    
    # Build last 12 months boundaries
    today = pd.Timestamp.today()
    first_of_current = today.replace(day=1)
    month_starts = [first_of_current - pd.DateOffset(months=i) for i in range(12, 0, -1)]
    month_starts.append(first_of_current)
    
    # Ensure target_date is datetime
    df_classified['target_date'] = pd.to_datetime(df_classified['target_date'], errors="coerce")
    
    # UPDATED STATUS GROUPINGS - Need to match original chart definition
    # Based on your expected numbers, let's see what statuses need to be regrouped
    closed_statuses = ["CLOSE", "REVWD", "PENRVW"]  # Keep these
    awaiting_qa_statuses = ["PENDQA"]  # Keep this
    awaiting_dept_statuses = ["FLAGGED", "MISSED"]  # Explicitly define these
    
    print("🔧 UPDATED status groupings:")
    print(f"   Closed: {closed_statuses}")
    print(f"   Awaiting QA: {awaiting_qa_statuses}")
    print(f"   Awaiting Dept: {awaiting_dept_statuses}")
    
    # Show overall completion_status breakdown for verification
    print("📊 Overall completion_status breakdown:")
    completion_status_counts = df_classified['completion_status'].value_counts(dropna=False)
    for status, count in completion_status_counts.items():
        print(f"   '{status}': {count}")
    
    # Find work orders that were flagged as missed using completion_status = 'MISSED'
    all_missed = df_classified[df_classified['completion_status'] == 'MISSED']
    print(f"📊 Total work orders with completion_status = 'MISSED': {len(all_missed)}")
    
    # DEBUGGING: Show detailed breakdown for July to understand the discrepancy
    print("\n🔍 DETAILED JULY ANALYSIS:")
    july_start = pd.Timestamp('2025-07-01')
    july_end = pd.Timestamp('2025-08-01')
    july_mask = (df_classified['target_date'] > july_start) & (df_classified['target_date'] <= july_end)
    july_all = df_classified[july_mask]
    july_missed = july_all[july_all['completion_status'] == 'MISSED']
    
    print(f"📅 July work orders (target_date 2025-07-01 to 2025-08-01):")
    print(f"   Total work orders in July: {len(july_all)}")
    print(f"   Missed work orders in July: {len(july_missed)}")
    
    if not july_missed.empty:
        print(f"\n📊 July missed work orders by status:")
        july_status_breakdown = july_missed["status"].value_counts()
        for status, count in july_status_breakdown.items():
            if status in closed_statuses:
                category = "Closed"
            elif status in awaiting_qa_statuses:
                category = "Awaiting QA"
            elif status in awaiting_dept_statuses:
                category = "Awaiting Dept"
            else:
                category = "UNASSIGNED"
            print(f"   {status}: {count} ({category})")
        
        # Check if we have unassigned statuses
        all_assigned_statuses = set(closed_statuses + awaiting_qa_statuses + awaiting_dept_statuses)
        unassigned = july_missed[~july_missed["status"].isin(all_assigned_statuses)]
        if not unassigned.empty:
            print(f"\n⚠️ UNASSIGNED statuses in July: {unassigned['status'].unique()}")
            print("💡 These statuses need to be added to one of the categories above")
    
    disposition_data = []
    
    for i in range(12):
        start = month_starts[i]
        end = month_starts[i+1]
        month_label = start.strftime("%b-%y")
        
        # Get work orders assigned to this month based on TARGET_DATE (when they were due)
        mask = (df_classified['target_date'] > start) & (df_classified['target_date'] <= end)
        month_df = df_classified[mask]
        
        # Find work orders that were flagged as missed using completion_status = 'MISSED'
        missed_wo_df = month_df[month_df['completion_status'] == 'MISSED']
        
        # Print unique statuses for debugging
        if not missed_wo_df.empty:
            unique_current_statuses = missed_wo_df["status"].unique()
            print(f"  {month_label}: Current statuses: {unique_current_statuses}")
        else:
            print(f"  {month_label}: No missed work orders found")
        
        # Categorize by CURRENT STATUS (using explicit lists)
        closed_count = missed_wo_df[missed_wo_df["status"].isin(closed_statuses)].shape[0]
        awaiting_qa_count = missed_wo_df[missed_wo_df["status"].isin(awaiting_qa_statuses)].shape[0]
        awaiting_dept_count = missed_wo_df[missed_wo_df["status"].isin(awaiting_dept_statuses)].shape[0]
        
        # Check for unassigned statuses
        all_assigned_statuses = set(closed_statuses + awaiting_qa_statuses + awaiting_dept_statuses)
        unassigned_count = missed_wo_df[~missed_wo_df["status"].isin(all_assigned_statuses)].shape[0]
        
        total_missed = len(missed_wo_df)
        
        disposition_data.append({
            "month": month_label,
            "closed": closed_count,
            "awaiting_qa": awaiting_qa_count, 
            "awaiting_dept": awaiting_dept_count,
            "unassigned": unassigned_count,  # Track unassigned for debugging
            "total": total_missed
        })
        
        print(f"  {month_label}: Total missed={total_missed}, Closed={closed_count}, QA={awaiting_qa_count}, Dept={awaiting_dept_count}")
        if unassigned_count > 0:
            unassigned_statuses = missed_wo_df[~missed_wo_df["status"].isin(all_assigned_statuses)]["status"].value_counts()
            print(f"    ⚠️ UNASSIGNED: {unassigned_count} - {unassigned_statuses.to_dict()}")
    
    # Convert to DataFrame
    disposition_df = pd.DataFrame(disposition_data)
    
    # Verify totals
    print(f"📊 Verification - Total missed from disposition: {disposition_df['total'].sum()}")
    print(f"📊 Total unassigned statuses: {disposition_df['unassigned'].sum()}")
    
    # Show overall status breakdown for all missed work orders
    if not all_missed.empty:
        print(f"📈 Overall CURRENT status breakdown for ALL missed work orders:")
        status_breakdown = all_missed["status"].value_counts()
        for status, count in status_breakdown.items():
            if status in closed_statuses:
                category = "Closed"
            elif status in awaiting_qa_statuses:
                category = "Awaiting QA"
            elif status in awaiting_dept_statuses:
                category = "Awaiting Dept"
            else:
                category = "UNASSIGNED"
            print(f"   {status}: {count} ({category})")
    
    # Chart creation code (rest remains the same)
    chart_shape = None
    chart_position = None
    
    # Find the chart and store its position
    for shape in slide.shapes:
        if hasattr(shape, "chart"):
            chart_position = {
                'left': shape.left,
                'top': shape.top, 
                'width': shape.width,
                'height': shape.height
            }
            print(f"🎯 Found chart at position: {chart_position}")
            chart_shape = shape
            break
    
    if chart_shape and chart_position:
        # Delete the old chart
        sp = chart_shape._element
        sp.getparent().remove(sp)
        print("🗑️ Removed problematic chart")
        
        # Create new chart at same position
        try:
            chart_data = CategoryChartData()
            chart_data.categories = disposition_df["month"].tolist()
            
            # Add series in order: "Closed", "Awaiting QA", "Awaiting Dept"
            chart_data.add_series("Closed", disposition_df["closed"].tolist())
            chart_data.add_series("Awaiting QA", disposition_df["awaiting_qa"].tolist()) 
            chart_data.add_series("Awaiting Dept", disposition_df["awaiting_dept"].tolist())
            
            # Create new chart
            left = chart_position['left']
            top = chart_position['top']
            width = chart_position['width'] 
            height = chart_position['height']
            
            new_chart = slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_STACKED, left, top, width, height, chart_data
            ).chart
            
            # Configure the new chart
            new_chart.has_legend = True
            new_chart.legend.position = XL_LEGEND_POSITION.TOP
            
            # Apply Style 8
            try:
                new_chart.chart_style = 8
                print("✅ Applied Style 8 to chart")
            except:
                print("⚠️ Could not apply Style 8 - using default style")
            
            # Format series with proper colors
            series_colors = [
                RGBColor(40, 167, 69),    # Green for Closed (bottom)
                RGBColor(255, 193, 7),    # Yellow for Awaiting QA (middle)  
                RGBColor(220, 53, 69)     # Red for Awaiting Dept (top)
            ]
            
            for i, series in enumerate(new_chart.series):
                if i < len(series_colors):
                    series.format.fill.solid()
                    series.format.fill.fore_color.rgb = series_colors[i]
                    print(f"  ✅ Formatted series '{series.name}' with color {series_colors[i]}")
            
            print("✅ Successfully created new disposition chart")
            
        except Exception as e:
            print(f"❌ Error creating new chart: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("❌ Could not find chart to replace")
    
    print(f"📈 Disposition summary: Total missed work orders = {disposition_df['total'].sum()}")
    
    return disposition_df

def update_missed_by_group_charts(prs, by_group_df, slide_index=3):  # Updated index: was 2, now 3
    slide = prs.slides[slide_index]
    for shape in slide.shapes:
        if hasattr(shape, "chart"):
            if shape.name == "Qty Missed by Group":
                chart_data = CategoryChartData()
                chart_data.categories = by_group_df["group"].tolist()
                chart_data.add_series("Missed", by_group_df["missed"].tolist())
                shape.chart.replace_data(chart_data)
            elif shape.name == "% Missed by Group":
                chart_data = CategoryChartData()
                chart_data.categories = by_group_df["group"].tolist()
                chart_data.add_series("% Missed", by_group_df["missed_percent"].tolist())
                shape.chart.replace_data(chart_data)
            
            print(f"Updating chart: {shape.name}")

def update_missed_still_open_chart(prs, by_group_df, slide_index=4):  # Updated index: was 3, now 4
    slide = prs.slides[slide_index]
    for shape in slide.shapes:
        if hasattr(shape, "chart") and shape.name == "Missed Still Open by Group":
            # Filter out groups where still_open == 0
            filtered_df = by_group_df[by_group_df["still_open"] > 0]
            chart_data = CategoryChartData()
            chart_data.categories = filtered_df["group"].tolist()
            chart_data.add_series("Still Open", filtered_df["still_open"].tolist())
            shape.chart.replace_data(chart_data)
            print(f"Updating chart: {shape.name}")

def create_full_governance_deck(summary_df, by_group_df, by_month_df, df_classified, late_df=None):
    print("create_full_governance_deck called")
    print("summary_df columns:", summary_df.columns)
    
    # Define output_path
    output_path = "outputs/presentations/governance_slide.pptx"
    
    prs = Presentation(r"data\templates\governance_slide_template.pptx")
    update_governance_slide(summary_df, prs, slide_index=0)
    update_missed_by_month_chart(prs, by_month_df, slide_index=1)
    
    # NEW: Add missed disposition chart at index 2
    update_missed_disposition_chart(prs, df_classified, slide_index=2)
    
    # UPDATED: Shift existing slides to new indices
    update_missed_by_group_charts(prs, by_group_df, slide_index=3)  # Was 2, now 3
    update_missed_still_open_chart(prs, by_group_df, slide_index=4)  # Was 3, now 4
    
    # If late_df is provided, also export it to Excel for presenter reference
    if late_df is not None and not late_df.empty:
        excel_path = "outputs/presentations/governance_with_late_orders.xlsx"
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            late_df.to_excel(writer, sheet_name='Late > 90 Days', index=False)
        print(f"✅ Excel with late orders saved to: {excel_path}")
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    print(f"✅ Full governance deck saved to: {output_path}")

