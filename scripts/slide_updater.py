# Scripts/slide_updater.py

from pptx import Presentation
from pptx.chart.data import CategoryChartData
import pandas as pd
import logging
import os

logging.basicConfig(level=logging.INFO)

def update_pm_missed_chart(prs, slide_index, chart_name, chart_data):
    slide = prs.slides[slide_index]
    chart_shape = None
    for shape in slide.shapes:
        if hasattr(shape, "chart") and shape.name == chart_name:
            chart_shape = shape
            break
    if not chart_shape:
        logging.warning(f"Chart '{chart_name}' not found on slide {slide_index+1}.")
        return
    cat_data = CategoryChartData()
    cat_data.categories = chart_data["months"]
    cat_data.add_series("Due", chart_data["due"])
    cat_data.add_series("Completed", chart_data["complete"])
    cat_data.add_series("Missed", chart_data["missed"])
    chart_shape.chart.replace_data(cat_data)

def update_group_charts(prs, group_data, slide_index):
    slide = prs.slides[slide_index]
    qty_chart_shape = None
    percent_chart_shape = None
    for shape in slide.shapes:
        if hasattr(shape, "chart"):
            if shape.name == "Qty Missed by Group":
                qty_chart_shape = shape
            elif shape.name == "% Missed by Group":
                percent_chart_shape = shape
    filtered = [
        (g, m, p)
        for g, m, p in zip(group_data["groups"], group_data["missed"], group_data["missed_percent"])
        if m > 0
    ]
    if not filtered:
        logging.warning("No groups with missed work orders.")
        return
    groups, missed, missed_percent = zip(*filtered)
    if qty_chart_shape:
        qty_data = CategoryChartData()
        qty_data.categories = groups
        qty_data.add_series("Missed", missed)
        qty_chart_shape.chart.replace_data(qty_data)
    else:
        logging.warning("Qty Missed by Group chart not found.")
    if percent_chart_shape:
        percent_data = CategoryChartData()
        percent_data.categories = groups
        percent_data.add_series("% Missed", missed_percent)
        percent_chart_shape.chart.replace_data(percent_data)
    else:
        logging.warning("% Missed by Group chart not found.")

# Example usage:
summary_data = {
    "months": ["Jan", "Feb", "Mar", "Apr"],
    "due": [50, 55, 60, 58],
    "complete": [45, 52, 55, 53],
    "missed": [5, 3, 5, 5]
}
group_data = {
    "groups": ["Ops", "Tech", "Admin"],
    "missed": [6, 8, 3],
    "missed_percent": [12.5, 18.0, 7.1],
    "reporting_month": "2025-06"
}

prs = Presentation("template.pptx")
update_pm_missed_chart(prs, slide_index=1, chart_name="PM Missed Chart", chart_data=summary_data)
update_group_charts(prs, group_data, slide_index=2)

for i, shape in enumerate(prs.slides[1].shapes):
    print(f"Slide 2 Shape {i}: type={shape.shape_type}, name={getattr(shape, 'name', None)}")
for i, shape in enumerate(prs.slides[2].shapes):
    print(f"Slide 3 Shape {i}: type={shape.shape_type}, name={getattr(shape, 'name', None)}")

last_month_label = pd.to_datetime(group_data["reporting_month"], format="%Y-%m").strftime("%b-%Y")
output_ppt_path = f"outputs/presentations/group_slide_deck_{last_month_label}.pptx"
os.makedirs(os.path.dirname(output_ppt_path), exist_ok=True)
prs.save(output_ppt_path)
logging.info(f"✅ Saved updated charts to {output_ppt_path}")