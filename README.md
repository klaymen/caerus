# Caerus Risk Manager

Caerus converts an Excel risk register into a single, shareable HTML dashboard.

The generated dashboard includes modern filtering, sortable columns, KPI cards, tooltips, CSV export, and responsive layout.

## What This Tool Generates

- sample input workbook: `sample_input/risk_register.xlsx`
- dashboard output: `output/risk_dashboard.html`

The dashboard is self-contained (one HTML file).

## Prerequisites

- Python 3.10+
- Terminal access

## Quick Start (Anyone Can Run)

1. Open a terminal in this project folder.
2. Create a virtual environment.
3. Install dependencies.
4. Generate sample data and dashboard.
5. Open the HTML output in a browser.

macOS/Linux:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python risk_manager_generator.py --create-sample
```

Windows (PowerShell):

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python risk_manager_generator.py --create-sample
```

Then open:

- `output/risk_dashboard.html`

## Use Your Own Excel File

Run with your file path:

macOS/Linux:

```bash
.venv/bin/python risk_manager_generator.py --input your_file.xlsx --output-dir output
```

Windows (PowerShell):

```powershell
.venv\Scripts\python risk_manager_generator.py --input your_file.xlsx --output-dir output
```

## Excel Input Format

The generator accepts common aliases, but these canonical fields are supported:

- Risk ID
- Title
- Description
- Project
- Severity
- Impact
- Probability
- Status
- Owner
- Category
- Identified Date
- Due Date
- Last Review
- Mitigation
- Trend

Required (must contain at least one populated value):

- Risk ID
- Title
- Project
- Severity
- Impact
- Identified Date

Date format recommendation:

- `YYYY-MM-DD`

## Sample Workbook Details

`sample_input/risk_register.xlsx` includes 2 sheets:

- `Risk Register`: ready-to-run example records
- `How To Fill`: formatted guide with required fields and allowed options/examples

## Dashboard Features

- Filters for date, severity, project, impact, probability, status, owner, category, trend
- Quick filters for open-only and overdue-only
- Search across ID/title/description/owner/project/mitigation
- Sort by clicking table headers
- KPI cards including Average Risk Score
- Mitigation Health gauge showing coverage of High+ severity risks
- Tooltip explanations for key fields and metrics
- Row detail panel for full risk context with Opportunity pill for opportunity items
- Configurable column visibility (Risk ID, Owner, and Identified Date hidden by default)
- Dark mode support
- Download filtered results as CSV
- Download full register as Excel (`.xlsx`) with Risk Register and Highlights sheets

## Command Reference

Create sample workbook and dashboard:

```bash
risk_manager_generator.py --create-sample
```

Generate dashboard from specific file:

```bash
risk_manager_generator.py --input <excel_file> --output-dir <folder>
```

Arguments:

- `--create-sample`: writes sample workbook before generating dashboard
- `--input`: path to Excel input (default: `sample_input/risk_register.xlsx`)
- `--output-dir`: output directory (default: `output`)

## Troubleshooting

`ModuleNotFoundError: No module named pandas`

- Use the virtual environment Python (`.venv/bin/python` or `.venv\Scripts\python`).
- Reinstall dependencies from `requirements.txt`.

Dashboard appears blank

- Regenerate output and hard-refresh browser.
- Confirm `output/risk_dashboard.html` was updated recently.

Excel file is open/locked

- Close the workbook in Excel before re-running generation.
