# 🛠️ work_order_analysis

## 📋 Overview

A Python-based tool that analyzes work order data to support governance reporting.  
This system handles data classification, chart and slide generation, and structured output for stakeholder visibility.

---

## 🌟 Features

- Modular workflow design for scalability and maintainability
- Excel chart creation via `xlsxwriter`
- Automated PowerPoint slide generation
- Customizable classification logic with edge case handling
- wxPython GUI for user interaction
- Unit-tested components for high reliability
- Markdown tree output to document directory structure

---

## 🗂️ Project Structure

work_order_analysis/
├── config/              # YAML and config files
├── data/                # Input datasets
├── gui/                 # wxPython GUI components
├── outputs/             # Generated files (Excel, PowerPoint, reports)
├── scripts/             # Core logic modules
│   ├── classifier.py        # Classifies work orders by status and type
│   ├── data_loader.py       # Loads and normalizes raw work order data
│   ├── summary_generator.py # Generates summary tables and breakdowns
│   ├── analysis_runner.py   # Runs analysis and prints/exports summaries
│   ├── slide_generator.py   # Creates and updates PowerPoint slides
│   ├── slide_updater.py     # Updates charts in PowerPoint presentations
│   ├── chart_builder.py     # Builds and exports charts for reporting
│   └── charts/
│       ├── pm_missed_chart.py      # Generates PM missed charts
│       └── group_missed_chart.py   # Generates group missed charts
├── tests/                # Unit tests validating functionality
│   ├── test_classifier.py         # Tests classification logic
│   ├── test_data_loader.py        # Tests data loading logic
│   └── test_summary_generator.py  # Tests summary generation logic
├── main.py               # Entry point for launching GUI and CLI processing
├── directory_tree.md     # Auto-generated directory documentation
└── README.md             # Project documentation

---

## 🧩 Module Descriptions

### scripts/classifier.py

- **Purpose:** Classifies work orders by status, dates, and type.
- **Requirements:** DataFrame with columns like 'status', 'target_date', 'actual_finish'.
- **Output:** Adds 'wo_class' column for reporting and analysis.

### scripts/data_loader.py

- **Purpose:** Loads raw work order data, normalizes columns, converts dates.
- **Requirements:** Excel or CSV input file, COLUMN_MAP config.
- **Output:** Cleaned DataFrame with normalized columns and 'report_month'.

### scripts/summary_generator.py

- **Purpose:** Generates summary tables and breakdowns for monthly, group, and governance metrics.
- **Requirements:** Classified DataFrame.
- **Output:** Summary DataFrames, Excel exports, late work order lists.

### scripts/analysis_runner.py

- **Purpose:** Runs analysis from cleaned data, prints and exports summary metrics.
- **Requirements:** 'cleaned_work_orders.csv' with required columns.
- **Output:** Console output, optional file export.

### scripts/slide_generator.py

- **Purpose:** Generates and updates PowerPoint slides for governance reporting.
- **Requirements:** DataFrames with summary/group/monthly data, PowerPoint template.
- **Output:** Updated presentation saved to outputs.

### scripts/slide_updater.py

- **Purpose:** Updates PowerPoint charts with new data.
- **Requirements:** Presentation object, chart data, chart names.
- **Output:** Updated charts in PowerPoint file.

### scripts/chart_builder.py

- **Purpose:** Builds and exports charts for reporting (Excel and PowerPoint).
- **Requirements:** DataFrame with monthly metrics.
- **Output:** Excel chart, PowerPoint slide.

### scripts/charts/pm_missed_chart.py

- **Purpose:** Generates PM missed charts by month.
- **Requirements:** Dict with 'months', 'due', 'complete', 'missed'.
- **Output:** Chart image file.

### scripts/charts/group_missed_chart.py

- **Purpose:** Generates group missed charts (count and percent).
- **Requirements:** Dict with 'groups', 'missed', 'missed_percent'.
- **Output:** Chart image file.

### gui/wx_app.py

- **Purpose:** wxPython GUI for user interaction, file selection, and report generation.
- **Requirements:** wxPython, input file, core scripts.
- **Output:** Interactive analysis, exports to Excel/PowerPoint.

### main.py

- **Purpose:** Entry point for CLI and GUI workflows.
- **Requirements:** Input file, all core modules.
- **Output:** CLI: console and file output; GUI: dashboard and export options.

### tests/

- **Purpose:** Unit tests for classifier, data loader, and summary generator.
- **Requirements:** Sample data, relevant modules.
- **Output:** Pass/fail results for each test case.

---

## 🚀 Getting Started

### Setup Virtual Environment

```bash
python -m venv .venv
.\.venv\Scripts\activate  # On Windows
pip install -r requirements.txt
```

### Run the Application

```bash
python main.py
```

- GUI will launch for user input, file selection, and report generation.

### Run in CLI Mode

```bash
python main.py path\to\your\input_file.csv
```

### 🧪 Running Tests

```bash
pytest tests/
```

- Includes unit tests for classifier logic, data flows, and output integrity.

---

## 📈 Governance Integration

- Classifications link to governance tags
- Slide decks generated with key metrics and summaries
- Clean handoff for stakeholder review using PowerPoint templates

---

## 📌 Future Enhancements

- CI/CD for automated testing and deployment
- HTML dashboard output option
- Tag releases for milestone tracking

---

## 🧠 Author Notes

This project was built with a focus on modularity, maintainability, and usability—empowering technical and non-technical stakeholders to rely on meaningful insights from work order data.

Future Enhancements

- Add CI/CD pipeline for automated validation
- Integrate dashboard summary in HTML
- Tag versions for traceable governance milestones
