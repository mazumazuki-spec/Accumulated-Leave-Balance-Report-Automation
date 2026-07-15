# Excel Report Automation Portfolio

A sanitized portfolio version of a real-world Python Excel reporting workflow.

This project demonstrates how to automate a multi-phase Excel reporting process with Python.  
It reads branch-level Excel files, cleans raw data, applies configurable demo rules, formats reports, builds branch summaries, consolidates a final summary, and optionally exports report screenshots on Windows machines with Microsoft Excel installed.

This repository contains **demo data rules only**. It does not include real company names, real branch names, internal URLs, machine paths, employee IDs, employee data, or confidential business logic.

## Project Overview

The original workflow was used to reduce repetitive Excel reporting tasks. This GitHub version keeps the technical workflow while replacing all confidential details with generic demo configuration.

The workflow is split into 4 phases:

1. **Phase 1 — Download or read raw Excel files and clean data**
2. **Phase 2 — Build branch-level reports from templates**
3. **Phase 3 — Consolidate all branch summaries into one summary workbook**
4. **Phase 4 — Finalize workbooks and optionally export screenshots**

A runner script is included to execute all phases in order.

## Features

- Read Excel workbooks
- Generate demo sample input files
- Clean raw Excel data
- Delete rows by keyword
- Delete rows by dummy employee ID
- Delete configured columns
- Add remark column for configured demo employees
- Merge cells
- Format Excel reports
- Apply borders and fills
- Auto fit column widths
- Create branch summary rows
- Consolidate branch summaries into one workbook
- Export finalized reports
- Optional screenshot export using Microsoft Excel COM on Windows
- Config-driven setup with JSON

## Technology Used

- Python 3.10+
- openpyxl
- requests
- Pillow
- pywin32 optional, only required for screenshot export on Windows

## Folder Structure

```text
excel-report-automation-portfolio/
├── config/
│   └── config.example.json
├── data/
│   ├── input/
│   ├── templates/
│   ├── output/
│   └── screenshots/
├── src/
│   ├── config_loader.py
│   ├── create_sample_data.py
│   ├── excel_utils.py
│   ├── phase1_download_clean.py
│   ├── phase2_build_branch_reports.py
│   ├── phase3_consolidate_summary.py
│   └── phase4_finalize_export.py
├── run_all.py
├── README.md
├── requirements.txt
└── .gitignore
```

## Installation

```bash
git clone https://github.com/your-username/excel-report-automation-portfolio.git
cd excel-report-automation-portfolio

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
```

For macOS or Linux:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## How to Run

### 1. Create demo input and template files

```bash
python src/create_sample_data.py --config config/config.example.json
```

### 2. Run all phases

```bash
python run_all.py --config config/config.example.json
```

### 3. Run a single phase

```bash
python src/phase1_download_clean.py --config config/config.example.json
python src/phase2_build_branch_reports.py --config config/config.example.json
python src/phase3_consolidate_summary.py --config config/config.example.json
python src/phase4_finalize_export.py --config config/config.example.json
```

### 4. Optional API download mode

The demo is designed to run with local Excel files by default.

To use an API source, set `api_base_url` in `config/config.example.json`, then run Phase 1 with:

```bash
python src/phase1_download_clean.py --config config/config.example.json --download
```

Use a dummy or public demo endpoint only. Do not commit real internal URLs.

## Example Workflow

1. Create sample raw workbooks and branch templates.
2. Phase 1 reads raw files from `data/input`.
3. Phase 1 cleans the files and saves them into `data/output/<run_date>/cleaned`.
4. Phase 2 creates branch reports from templates.
5. Phase 3 consolidates all branch summaries into one summary workbook.
6. Phase 4 finalizes each report workbook and exports screenshots when supported.

## Config

All frequently changed values are stored in `config/config.example.json`.

Examples:

- Branch list
- Dummy branch IDs
- Dummy employee IDs
- Row deletion keywords
- Columns to delete
- Threshold days
- Report template ranges
- Output folders

## Privacy and Sanitization Checklist

This portfolio version removes or replaces:

- Real company names
- Real branch names
- Internal URLs
- Local company machine paths
- Real employee IDs
- Employee personal data
- Company-specific business rules
- Internal comments and labels that identify the original organization

All names, IDs, URLs, and rules are generic demo examples.
