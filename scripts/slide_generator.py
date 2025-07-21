# scripts/slide_generator.py

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

def create_governance_slide(summary_df, prs=None):
    if prs is None:
        prs = Presentation()

    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Preventive Maintenance Summary"

    pm_row = summary_df.iloc[0]
    left, top, width, height = Inches(1), Inches(2), Inches(6), Inches(2)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = True
    tf.text = (
        f"Due: {pm_row['due']:,}\n"
        f"Completed: {pm_row['completed']:,}\n"
        f"Missed: {pm_row['missed']:,}\n"
        f"Completion %: {pm_row['completion_pct']}%"
    )

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


