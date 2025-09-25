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
from pptx.enum.chart import XL_CHART_TYPE
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
    summary_df_no_total = summary_df[summary_df["Month"] != "Grand Total"]
    summary_df_sorted = summary_df_no_total.sort_values(
        "Month", key=lambda x: pd.to_datetime(x, format="%b-%y"), ignore_index=True
    )
    # Optionally, add the Grand Total row back if needed
    grand_total_row = summary_df[summary_df["Month"] == "Grand Total"]
    if not grand_total_row.empty:
        summary_df_sorted = pd.concat([summary_df_sorted, grand_total_row], ignore_index=True)

    current_month_row = summary_df_sorted.iloc[-1]

    # YTD: filter for current year only
    summary_df_no_total = summary_df_sorted[summary_df_sorted["Month"] != "Grand Total"]
    summary_df_no_total["year"] = summary_df_no_total["Month"].str[-2:].astype(int)
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

def update_missed_by_month_chart(prs, by_month_df, slide_index=1, chart_name="PM Missed Chart"):
    slide = prs.slides[slide_index]
    for shape in slide.shapes:
        if hasattr(shape, "chart") and shape.name == chart_name:
            chart_data = CategoryChartData()
            chart_data.categories = by_month_df["report_month"].astype(str).tolist()
            chart_data.add_series("Missed", by_month_df["missed"].tolist())
            chart_data.add_series("Completed", by_month_df["completed"].tolist())
            chart_data.add_series("Generated", by_month_df["generated"].tolist())
            shape.chart.replace_data(chart_data)
            break

def update_missed_by_group_charts(prs, by_group_df, slide_index=2):
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

def update_missed_still_open_chart(prs, by_group_df, slide_index=3):
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

def create_full_governance_deck(summary_df, by_group_df, by_month_df, output_path="outputs/presentations/governance_slide.pptx"):
    print("create_full_governance_deck called")
    print("summary_df columns:", summary_df.columns)
    prs = Presentation(r"data\templates\governance_slide_template.pptx")
    update_governance_slide(summary_df, prs, slide_index=0)  # Summary slide
    update_missed_by_month_chart(prs, by_month_df, slide_index=1)  # PM Missed Chart
    update_missed_by_group_charts(prs, by_group_df, slide_index=2)  # Group charts
    update_missed_still_open_chart(prs, by_group_df, slide_index=3)  # Still Open chart
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    print(f"✅ Full governance deck saved to: {output_path}")
    for slide_idx, slide in enumerate(prs.slides):
        for i, shape in enumerate(slide.shapes):
            print(f"Slide {slide_idx} - Shape {i}: type={shape.shape_type}, name={getattr(shape, 'name', None)}")



