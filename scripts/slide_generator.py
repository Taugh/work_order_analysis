# scripts/slide_generator.py

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor


def create_governance_slide(summary_df, prs):
    print("create_governance_slide called")
    print("summary_df columns:", summary_df.columns)
    
    """
        Creates a slide titled 'Preventive Maintenance Summary' using the grand total row
        from the summary DataFrame. The layout and formatting match the SLT governance deck.
        """

    grand_total = summary_df.iloc[0]

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
        f"Due: {grand_total['Due']:.0f}\n"
        f"Completed: {grand_total['Completed']:.0f}\n"
        f"Missed: {grand_total['Missed']:.0f}\n"
        f"Completion %: {grand_total['Completion %']:.1f}"
    )
    p.font.size = Pt(24)
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
    
def update_governance_slide(summary_df, prs, slide_index=1):
    slide = prs.slides[slide_index]  # Use the existing slide
    # Find the shape (e.g., a text box or table) you want to update
    # Example: update a text box with summary info
    for shape in slide.shapes:
        if shape.has_text_frame and "Preventive Maintenance Summary" in shape.text:
            text_frame = shape.text_frame
            text_frame.clear()
            grand_total = summary_df.iloc[0]
            p = text_frame.paragraphs[0]
            p.text = (
                f"Due: {grand_total['Due']:.0f}\n"
                f"Completed: {grand_total['Completed']:.0f}\n"
                f"Missed: {grand_total['Missed']:.0f}\n"
                f"Completion %: {grand_total['Completion %']:.1f}"
            )
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)
            break

def update_missed_by_month_chart(prs, by_month_df, slide_index=2, chart_name="PM Missed Chart"):
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

def update_missed_by_group_charts(prs, by_group_df, slide_index=3):
    slide = prs.slides[slide_index]
    # Example: update two charts by name
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

def create_full_governance_deck(summary_df, by_group_df, by_month_df, output_path="outputs/presentations/governance_slide.pptx"):
    print("create_full_governance_deck called")
    prs = Presentation(r"data\templates\governance_slide_template.pptx")
    update_governance_slide(summary_df, prs, slide_index=1)
    update_missed_by_month_chart(prs, by_month_df, slide_index=2)
    update_missed_by_group_charts(prs, by_group_df, slide_index=3)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    print(f"✅ Full governance deck saved to: {output_path}")

