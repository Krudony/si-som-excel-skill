# Si-Som EXCEL Skill (si-som-excel)

## Overview
A specialized skill for "Don" (Krudony) to act as an Expert Excel Data Analyst and Automator. This skill empowers the AI to read, analyze, create, and modify Excel files (`.xlsx`) using Python engines.

## Core Mandates
1. **Analyze First:** Before modifying an existing Excel file, ALWAYS use `scripts/read_excel.py` to understand its structure (Sheets, Headers, Data Types).
2. **Precision:** When updating formulas or values, ensure exact cell targeting (e.g., 'Sheet1!A1').
3. **Data Integrity:** Never overwrite source data unless explicitly requested. Prefer creating new columns or summary sheets.

## Toolchain
The skill provides the following Python scripts located in the `scripts/` directory:

### 1. `read_excel.py` (The Eyes)
**Purpose:** Inspect the structure of an Excel file.
**Usage:** `python scripts/read_excel.py <path_to_excel_file>`
**Output:** Lists Sheet names, Column headers, and the first 5 rows of data for context.

### 2. `write_excel.py` (The Hands)
**Purpose:** Modify specific cells, add formulas, or write new data.
**Usage:** `python scripts/write_excel.py <path_to_file> <sheet_name> <cell_ref> <value_or_formula>`
**Example:** `python scripts/write_excel.py "data.xlsx" "Sheet1" "C2" "=A2+B2"`
