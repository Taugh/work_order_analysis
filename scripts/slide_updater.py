# Scripts/slide_updater.py

from pptx import Presentation
from pptx.util import Inches
import pandas as pd
import logging
import os
from scripts.charts.pm_missed_chart import build_pm_missed_chart
from scripts.charts.group_missed_chart import build_group_missed_chart, build_group_missed_percent_chart


logging.basicConfig(level=logging.INFO)

def update_pm_missed_slide(prs_path, slide_index, chart_data, output_image_path):
    # Build chart image
    chart_path = build_pm_missed_chart(chart_data, output_image_path)

    # Extract last month from chart data
    last_month_label = pd.to_datetime(chart_data["months"][-1]).strftime("%b-%Y")

    # Define output presentation path
    output_ppt_path = f"outputs/presentations/pm_slide_deck_{last_month_label}.pptx"

    # Load and update slide
    prs = Presentation(prs_path)
    slide = prs.slides[slide_index]

    left = Inches(1)
    top = Inches(2)
    height = Inches(4)

    slide.shapes.add_picture(chart_path, left, top, height=height)
    os.makedirs(os.path.dirname(output_ppt_path), exist_ok=True)
    prs.save(output_ppt_path)

    logging.info(f"✅ Saved updated presentation to {output_ppt_path}")

    return output_ppt_path

summary_data = {
    "months": ["Jan", "Feb", "Mar", "Apr"],
    "due": [50, 55, 60, 58],
    "complete": [45, 52, 55, 53],
    "missed": [5, 3, 5, 5]
}

update_pm_missed_slide(
    prs_path="template.pptx",
    slide_index=2,
    chart_data=summary_data,
)

def update_group_slides(prs_path, group_data, slide_index_chart, slide_index_percent):
    """
    Updates two slides: missed work orders by group and percent missed by group.

    Args:
        prs_path: Path to PowerPoint template.
        group_data: Dictionary with keys 'groups', 'missed', 'missed_percent'.
        slide_index_chart: Index of slide for missed WOs by group.
        slide_index_percent: Index of slide for missed % WOs by group.
    """

    # Extract reporting label from latest date (optional)
    last_month_label = pd.to_datetime(group_data["reporting_month"], format="%Y-%m").strftime("%b-%Y")

    # Output paths
    chart_path_1 = f"outputs/group_missed_chart_{last_month_label}.png"
    chart_path_2 = f"outputs/group_missed_percent_chart_{last_month_label}.png"
    output_ppt_path = f"outputs/presentations/group_slide_deck_{last_month_label}.pptx"

    # Generate charts
    build_group_missed_chart({"groups": group_data["groups"], "missed": group_data["missed"]}, chart_path_1)
    build_group_missed_percent_chart({"groups": group_data["groups"], "missed_percent": group_data["missed_percent"]}, chart_path_2)

    # Load presentation
    prs = Presentation(prs_path)

    # Insert chart 1
    slide_chart = prs.slides[slide_index_chart]
    slide_chart.shapes.add_picture(chart_path_1, Inches(1), Inches(2), height=Inches(4))

    # Insert chart 2
    slide_percent = prs.slides[slide_index_percent]
    slide_percent.shapes.add_picture(chart_path_2, Inches(1), Inches(2), height=Inches(4))

    # Save presentation
    prs.save(output_ppt_path)
    logging.info(f"✅ Saved group report slides to {output_ppt_path}")

    return output_ppt_path

group_data = {
    "groups": ["Ops", "Tech", "Admin"],
    "missed": [6, 8, 3],
    "missed_percent": [12.5, 18.0, 7.1],
    "reporting_month": "2025-06"  # or dynamically fetched
}
update_group_slides("template.pptx", group_data, slide_index_chart=3, slide_index_percent=4)