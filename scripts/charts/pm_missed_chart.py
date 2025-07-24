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