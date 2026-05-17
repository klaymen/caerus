# Lightweight Risk Manager (Excel -> HTML + JS)

This tool reads a risk register from an Excel file and generates:

- `risk_dashboard.html`

The dashboard is modern, lightweight, fully client-side, and delivered as one self-contained file for quick sharing.

## Features

- Fast filtering for:
  - Identified date range
  - Severity
  - Project
  - Impact
  - Probability
  - Status
  - Owner
  - Category
  - Trend
  - Open only / Overdue only
- Search across risk ID, title, description, owner, project, mitigation
- Sortable columns
- KPI cards (total, high+critical, open, overdue, average risk score)
- Risk details drawer
- Download currently filtered view as CSV

## Setup

1. Create a local virtual environment and install dependencies:

```bash
/opt/homebrew/bin/python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

2. Generate sample input and dashboard:

```bash
.venv/bin/python risk_manager_generator.py --create-sample
```

This creates:

- `sample_input/risk_register.xlsx`
- `output/risk_dashboard.html`

The sample Excel includes two sheets:

- `Risk Register` with example risks
- `How To Fill` with field-by-field instructions, required flags, and allowed option values

3. Open `output/risk_dashboard.html` in your browser.

## Use your own Excel file

Your sheet can use these columns (aliases are accepted):

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

Run:

```bash
.venv/bin/python risk_manager_generator.py --input your_file.xlsx --output-dir output
```

## Suggested additional filters

Beyond your requested filters (date, severity, project, impact), the tool includes:

- Probability (likelihood)
- Status (open/in progress/closed)
- Owner (accountability)
- Category (technology/security/legal/people/etc.)
- Trend (improving/stable/worsening)
- Open-only and overdue-only quick views
