```markdown
# Accumulated Leave Balance Report Automation

Python-based automation workflow for transforming employee leave data into standardized multi-branch reports and management summaries.

This project demonstrates data cleaning, Excel report automation, data consolidation, and final report generation using Python and Excel Automation.

---

## Project Overview

This project automates the monthly accumulated leave balance reporting process.

The workflow covers:

- Data preparation and cleaning
- Branch-level report generation
- Summary report consolidation
- Final report formatting and publishing

The goal is to reduce repetitive manual Excel operations, improve reporting consistency, and create a reusable reporting workflow.

---

## Business Problem

The original reporting process required multiple manual steps:

- Cleaning raw employee leave data
- Copying data into Excel templates
- Updating formulas and formatting for each branch
- Combining summary information from multiple reports
- Preparing final reports for management review

These processes were time-consuming and could introduce errors from manual operations.

---

## Solution

Developed a Python automation pipeline divided into 4 phases:

```

Raw Data
↓
Data Cleaning
↓
Branch Report Generation
↓
Consolidated Summary Report
↓
Final Excel Report + Screenshot

```

---

# Workflow

## Phase 1: Data Preparation & Cleaning

Process raw employee leave data and prepare structured datasets.

Features:

- Import raw data files
- Clean and standardize employee information
- Validate data format
- Prepare branch-level datasets

Example:

| Employee ID | Branch | Employee Name | Annual Leave | Sick Leave | Leave Balance |
|-------------|--------|---------------|--------------|------------|---------------|
| 10001 | STY | Employee A | 10 | 5 | 15 |
| 10002 | SYY | Employee B | 8 | 3 | 11 |

---

## Phase 2: Automated Branch Report Generation

Generate standardized reports for multiple branches.

Features:

- Load Excel report templates
- Copy template structure and formatting
- Maintain merged cells and styles
- Generate summary formulas automatically
- Apply filters based on business conditions
- Highlight important report columns

Output example:

| Employee ID | Employee Name | Leave Balance |
|-------------|---------------|---------------|
| 10001 | Employee A | 15 |
| 10002 | Employee B | 11 |

---

## Phase 3: Consolidated Management Report

Combine branch summaries into a centralized report.

Features:

- Extract summary values from branch reports
- Consolidate multiple branch results
- Update reporting dates automatically
- Compare leave balance changes between periods

Example:

| Branch | Employee Count | Leave > 5 Days | Total Leave Balance |
|--------|----------------|-----------------|---------------------|
| STY | 120 | 35 | 450 |
| SYY | 250 | 80 | 920 |
| SPN | 100 | 20 | 300 |

---

## Phase 4: Final Report Publishing

Prepare final output files for reporting usage.

Features:

- Standardize worksheet names
- Remove temporary worksheets
- Adjust Excel formatting
- Auto-fit report columns
- Export report screenshots

Final output:

```

report_STY_YYYY-MM-DD.xlsx
report_SYY_YYYY-MM-DD.xlsx
report_SPN_YYYY-MM-DD.xlsx
report_SKC_YYYY-MM-DD.xlsx
report_SNT_YYYY-MM-DD.xlsx

```

---

# Technologies

## Programming

- Python
- OpenPyXL
- PyWin32 (Excel COM Automation)
- Subprocess

## Excel Automation

- Excel Formula Automation
- SUMIFS / COUNTIFS
- AutoFilter
- Template-based Reporting
- Report Consolidation
- Excel Formatting Automation

---

# Project Structure

```

Accumulated-Leave-Report-Automation/

├── src/
│   ├── data_cleaning.py
│   ├── generate_report.py
│   ├── consolidate_report.py
│   └── run_all_phases.py
│
├── sample_data/
│   ├── raw_data_sample.xlsx
│   └── cleaned_data_sample.xlsx
│
├── templates/
│   └── report_template.xlsx
│
├── screenshots/
│
├── requirements.txt
│
└── README.md

```

---

# Key Skills Demonstrated

- Data Cleaning & Transformation
- Python Automation
- Excel Automation
- Report Generation
- Data Consolidation
- Business Reporting Workflow

---

# Business Impact

- Reduced repetitive manual Excel processes
- Improved consistency across multiple branch reports
- Reduced risk of human errors
- Created a reusable monthly reporting workflow
- Improved efficiency in management report preparation

---

# Note

This repository contains a portfolio version using sample data.

Actual business data and confidential information have been removed.
```


