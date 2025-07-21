# config/settings.py

from pathlib import Path

# Base directory

BASE_DIR = Path(__file__).resolve().parent.parent

# Data paths

RAW_DATA_DIR = BASE_DIR / "data" / "raw"
PROCESSED_DATA_DIR = BASE_DIR / "data" / "processed"
LOG_DIR = BASE_DIR / "data" / "logs"
REPORT_DIR = BASE_DIR / "outputs" / "reports"
PPTX_DIR = BASE_DIR / "outputs" / "presentations"

# Expected column mappings

COLUMN_MAP ={
    "work_orders": "work_orders",
    "current_status": "status",
    "targ_comp_date": "target_date",
    "act_finish": "actual_finish",
    "finish_no_later": "grace_date",
    "work_type": "type",
    "wo_description": "description",
    "report_date": "report_date",
    "wo_assigned_group": "group"
    }

# Reporting window

MONTH_LOOKBACK = 12  # number of months to include

# PowerPoint slide targeting keywords (if automating updates)

PPTX_TEXT_MAP = {
    "total_gen": "Gen",
    "completed": "Comp",
    "missed": "Missed",
    "open": "Open"
    }

# Logging

LOGGING_ENABLED = True








