# 🛠️ work_order_analysis

## 📋 Overview
A Python-based tool that analyzes work order data to support governance reporting. This system handles data classification, chart and slide generation, and structured output for stakeholder visibility.

## 🌟 Features
- Modular workflow design for scalability and maintainability
- Excel chart creation via `xlsxwriter`
- Automated PowerPoint slide generation
- Customizable classification logic with edge case handling
- wxPython GUI for user interaction
- Unit-tested components for high reliability
- Markdown tree output to document directory structure

## 🗂️ Project Structure
work_order_analysis/ 
├── config/              # YAML and config files 
├── data/                # Input datasets 
├── gui/                 # wxPython GUI components 
├── outputs/             # Generated files (Excel, PowerPoint, reports) 
├── scripts/             # Core logic modules (e.g. classifier.py) 
├── tests/               # Unit tests validating functionality 
├── main.py              # Entry point for launching GUI and processing 
├── directory_tree.md    # Auto-generated directory documentation 
└── README.md            # Project documentation


## 🚀 Getting Started

### Setup Virtual Environment
```bash
python -m venv .venv
.\.venv\Scripts\activate  # On Windows
pip install -r requirements.txt

Run the Application
python main.py


GUI will launch for user input, file selection, and report generation.

🧪 Running Test
pytest tests/

Includes unit tests for classifier logic, data flows, and output integrity.

📈 Governance Integration
- Classifications link to governance tags
- Slide decks generated with key metrics and summaries
- Clean handoff for stakeholder review using PowerPoint templates

📌 Future Enhancements
- CI/CD for automated testing and deployment
- HTML dashboard output option
- Tag releases for milestone tracking

🧠 Author Notes
This project was built with a focus on modularity, maintainability, and usability—empowering technical and non-technical stakeholders to rely on meaningful insights from work order data

Future Enhancements
- Add CI/CD pipeline for automated validation
- Integrate dashboard summary in HTML
- Tag versions for traceable governance milestones




