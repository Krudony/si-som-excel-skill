# Si-Som EXCEL Skill (si-som-excel)

## Overview
A specialized skill for "Don" (Krudony) to act as an Expert Excel Data Analyst and Automator. This skill empowers the AI to read, analyze, create, and modify Excel files (.xlsx) using Python engines.

## Core Mandates
1. **Analyze First:** Before modifying an existing Excel file, ALWAYS use scripts/read_excel.py to understand its structure.
2. **Precision Over Aesthetics:** Functionality comes first. AVOID unnecessary cell merging in data areas to maintain sorting/filtering capabilities.
3. **Thai Standard Font (CRITICAL):** ALWAYS apply **TH SarabunPSK 16pt** to all cells modified or created by this skill.
4. **Excel Table Objects:** Prefer using Excel Table Objects (ListObjects) for data ranges to enable automatic filters and professional styling.
5. **Data Types:** Ensure numeric values are stored as numbers (with commas/decimals as needed) and dates as proper Excel date objects.

## Toolchain
- ead_excel.py: Inspect structure and data sample.
- write_excel.py: Expert-grade cell modification with **TH SarabunPSK 16pt** styling.
