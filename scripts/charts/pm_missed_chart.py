# ---------------------------------------------------------------
# pm_missed_chart.py
#
# Purpose:
#   Generates a chart visualizing Preventive Maintenance (PM) work orders by month,
#   including due, completed, and missed counts, with stoplight annotations.
#
# Requirements:
#   - Input: Dictionary with keys 'months', 'due', 'complete', 'missed'.
#   - Libraries: matplotlib for plotting.
#   - Output path: Filepath to save the generated chart image (PNG, JPG, etc.).
#
# Output:
#   - Saves the PM missed chart image to the specified output path.
#   - Returns the filepath to the saved chart image for use in reports or presentations.
#
# Notes:
#   - Used by reporting modules to visualize monthly PM work order metrics.
#   - Function: build_pm_missed_chart(data, output_path)
# ---------------------------------------------------------------

# scripts/charts/pm_missed_chart.py

import matplotlib.pyplot as plt

def build_pm_missed_chart(data: dict, output_path: str) -> str:
    """
    Generates a Preventive Maintenance chart with line and bar plots.
    
    Args:
        data: Dictionary with keys 'months', 'due', 'complete', 'missed'.
        output_path: Filepath where the chart image will be saved.
    
    Returns:
        The filepath to the saved chart image.
    """

    months = data["months"]
    due = data["due"]
    complete = data["complete"]
    missed = data["missed"]

    fig, ax1 = plt.subplots(figsize=(10, 6))

    # Line plots
    ax1.plot(months, due, label="Due", marker="o", color="steelblue")
    ax1.plot(months, complete, label="Completed", marker="o", color="darkgreen")

    # Bar plot for missed
    ax1.bar(months, missed, label="Missed", color="firebrick", alpha=0.7)

    # Annotate stop light thresholds
    for i, val in enumerate(missed):
        if val <= 4:
            label = "✅ Acceptable"
        elif val <= 7:
            label = "⚠️ Caution"
        else:
            label = "❌ Critical"
        ax1.text(months[i], val + 1, label, ha="center", fontsize=10)

    # Labels, grid, legend
    ax1.set_title("Preventive Maintenance Work Orders by Month")
    ax1.set_ylabel("Work Order Count")
    ax1.grid(True, linestyle="--", alpha=0.5)
    ax1.legend()

    plt.tight_layout()
    plt.savefig(output_path)
    plt.close(fig)  # Close for memory efficiency

    return output_path