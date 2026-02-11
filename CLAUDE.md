# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Python CLI tool for bidirectional conversion between Excel (.xlsx) and JSON files. Uses pandas for data handling.

## Commands

```bash
# Activate virtual environment (Windows)
venv\Scripts\activate

# Excel to JSON
python test.py <source.xlsx> <output_dir>

# JSON to Excel
python test.py <source.json> <output_dir>
```

Dependencies: `pandas`, `openpyxl` (installed in local venv).

## Architecture

- **convert.py** — Core conversion functions:
  - `convertXLtoJSON(xlsx_path, output_dir)` — Reads Excel, writes JSON (records-oriented, non-ASCII preserved)
  - `convertJSONtoXL(json_path, output_dir)` — Reads JSON, expands array columns into numbered columns (`col_0`, `col_1`, ...), and merges with existing Excel output (deduplicates, keeps last)
- **test.py** — CLI entry point. Dispatches to the appropriate converter based on source file extension.

## Key Behaviors

- `convertJSONtoXL` appends to an existing Excel file if one already exists at the output path, deduplicating rows.
- Array-type JSON fields are flattened into separate numbered columns before writing to Excel.
- Output directories are created automatically.
