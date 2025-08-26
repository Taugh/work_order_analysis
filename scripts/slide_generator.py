# scripts/slide_generator.py

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor


def create_governance_slide(summary_df, prs=None):
    
    """
        Creates a slide titled 'Preventive Maintenance Summary' using the grand total row
        from the summary DataFrame. The layout and formatting match the SLT governance deck.
        """

    
    # Extract the grand total row
    grand_total = summary_df[summary_df["Month"] == "Grand Total"].iloc[0]

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
        f"Completion %: {grand_total['Completion %']:.1f}%"
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
    chart_data.categories = by_month_df["report_month"].tolist()

    chart_data.add_series("Missed", by_month_df["missed"].tolist())
    chart_data.add_series("Completed", by_month_df["completed"].tolist())
    chart_data.add_series("Generated", by_month_df["generated"].tolist())

    x, y, cx, cy = Inches(1), Inches(1.5), Inches(8), Inches(4.5)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    ).chart

    # Convert second and third series to line charts
    chart.series[1].chart_type = XL_CHART_TYPE.LINE_MARKERS
    chart.series[2].chart_type = XL_CHART_TYPE.LINE_MARKERS

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
    
def create_full_governance_deck(summary_df, by_group_df, by_month_df, output_path=
                                "outputs/presentations/governance_slide.pptx"):

    prs = Presentation()

    # Add slides
    create_governance_slide(summary_df, prs)
    create_missed_by_month_slide(prs, by_month_df)
    create_missed_by_group_slide(prs, by_group_df)

    # Save deck
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    print(f"✅ Full governance deck saved to: {output_path}")


