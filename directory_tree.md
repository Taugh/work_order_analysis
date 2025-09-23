```
work_order_analysis/
├── .git
│   ├── COMMIT_EDITMSG
│   ├── HEAD
│   ├── config
│   ├── description
│   ├── hooks
│   │   ├── applypatch-msg.sample
│   │   ├── commit-msg.sample
│   │   ├── fsmonitor-watchman.sample
│   │   ├── post-update.sample
│   │   ├── pre-applypatch.sample
│   │   ├── pre-commit.sample
│   │   ├── pre-push.sample
│   │   ├── pre-rebase.sample
│   │   ├── pre-receive.sample
│   │   ├── prepare-commit-msg.sample
│   │   └── update.sample
│   └── logs
│       └── HEAD
├── .gitignore
├── .vscode
│   └── launch.json
├── README.md
├── config
│   ├── __init__.py
│   └── settings.py
├── data
│   ├── logs
│   ├── processed
│   │   ├── cleaned_work_orders.csv
│   │   └── governance_slide_template.pptx
│   └── raw
│       └── TestQSRData.xlsx
├── directory_tree.md
├── gui
│   ├── __init__.py
│   └── wx_app.py
├── main.py
├── outputs
│   ├── presentations
│   │   └── governance_slide.pptx
│   └── reports
│       ├── governance_overview.xlsx
│       └── monthly_summary.xlsx
├── requirements.txt
├── scripts
│   ├── __init__.py
│   ├── analysis_runner.py
│   ├── chart_builder.py
│   ├── charts
│   │   ├── __init__.py
│   │   ├── group_missed_chart.py
│   │   └── pm_missed_chart.py
│   ├── classifier.py
│   ├── data_loader.py
│   ├── printer.py
│   ├── slide_generator.py
│   ├── slide_updater.py
│   └── summary_generator.py
└── tests
    ├── __init__.py
    ├── test_classifier.py
    ├── test_data_loader.py
    └── test_summary_generator.py
```