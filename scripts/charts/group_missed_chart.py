# scripts/charts/group_missed_chart.py

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm  # Only if you’re using custom font files

def build_group_missed_chart(data: dict, output_path: str) -> str:
    """
    Generates a Group Missed chart with line and bar plots.
    
    Args:
        data: Dictionary with keys 'groups', 'missed'.
        output_path: Filepath where the chart image will be saved.
    
    Returns:
        The filepath to the saved chart image.
    """

    groups = data["groups"]
    missed = data["missed"]

    fig, ax1 = plt.subplots(figsize=(10, 6))

    # Bar plot for missed
    ax1.bar(groups, missed, label="Missed", color="firebrick", alpha=0.7)

    for i, val in enumerate(missed):
        ax1.text(i, val + 1, str(val), ha="center", fontsize=10)

    # Annotate stop light thresholds
    for i, val in enumerate(missed):
        if val <= 4:
            label = "✅ Acceptable"
        elif val <= 7:
            label = "⚠️ Caution"
        else:
            label = "❌ Critical"
        ax1.text(groups[i], val + 1, label, ha="center", fontsize=10)

    # Labels, grid, legend
    ax1.set_title("Missed Work Orders by Group")
    ax1.set_ylabel("Work Order Count")
    ax1.grid(True, linestyle="--", alpha=0.5)
    ax1.legend()

    plt.tight_layout()
    plt.savefig(output_path)
    plt.close(fig)  # Close for memory efficiency

    return output_path

def  build_group_missed_percent_chart(data: dict, output_path: str) -> str:
    """
    Generates a Group Missed Percent chart with bar plots.
    
    Args:
        data: Dictionary with keys 'groups', 'missed_percent'.
        output_path: Filepath where the chart image will be saved.
    
    Returns:
        The filepath to the saved chart image.
    """

    groups = data["groups"]
    missed_percent = data["missed_percent"]

    fig, ax1 = plt.subplots(figsize=(10, 6))

    # Bar plot for missed percent
    ax1.bar(groups, missed_percent, label="Missed Percent", color="firebrick", alpha=0.7)

    for i, val in enumerate(missed_percent):
        if val > 0:
            ax1.text(i, val + (val * 0.05) + 1, f"{round(val, 1)}%", ha="center", fontsize=10)

    # Labels, grid, legend
    ax1.set_title("Missed Work Orders Percentage by Group")
    ax1.set_ylabel("Missed Percentage (%)")
    ax1.set_ylim(0, 100)
    ax1.grid(True, linestyle="--", alpha=0.5)
    ax1.legend()

    plt.tight_layout()
    plt.savefig(output_path)
    plt.close(fig)  # Close for memory efficiency

    return output_path


