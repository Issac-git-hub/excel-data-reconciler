# Excel & CSV Data Reconciliation Checker

A lightweight Python automation tool that compares two Excel (`.xlsx`) or CSV (`.csv`) datasets—such as an internal invoice ledger vs. a vendor statement, or an HR roster vs. a benefits/ORSO schedule—and generates a formatted, color-coded Excel exception report.

## Features
- Supports `.csv` and `.xlsx` input files interchangeably
- Detects duplicate IDs in either file before matching
- Identifies missing records (`Missing in Target` / `Missing in Source`)
- Flags amount mismatches (with configurable tolerance) and status discrepancies
- Exports a multi-sheet Excel report (`Summary`, `Discrepancies`, `Missing_Records`, `Duplicates`, `Matched_Records`)

## Quick Start
```bash
pip install -r requirements.txt
python sample_data/generate_sample_data.py
python reconcile.py
```
