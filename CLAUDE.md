# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Python CLI tool for bidirectional conversion between Excel (.xlsx) and JSON files. Supports simple columns, flat arrays, and structured (nested object) columns with merged headers.

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

### Core Modules

- **test.py** — CLI entry point. Dispatches to converter based on file extension.
- **convert.py** — Core conversion logic with support for three column types:

### Column Type System

The converter handles three distinct column patterns:

1. **Simple columns** — Scalar values (strings, numbers)
2. **Flat arrays** — Numbered columns like `tags_0`, `tags_1` → JSON array `["a", "b"]`
3. **Structured columns** — Merged headers with sub-fields:
   - **Single objects**: `address` (merged) with sub-columns `street`, `city` → JSON object `{"street": "...", "city": "..."}`
   - **Array of objects**: `items_0`, `items_1` (merged) with sub-columns `name`, `price` → JSON array `[{"name": "...", "price": ...}, ...]`

### Excel → JSON (`convert_xl_to_json`)

1. `_detect_xl_columns(ws)` — Inspects row-1 merged cells and `_N` suffixes to classify columns
2. Reads data starting from row 3 (if structured) or row 2 (if flat-only)
3. Collapses numbered/structured columns back into arrays or objects
4. Writes JSON with `ensure_ascii=False` to preserve non-ASCII characters

### JSON → Excel (`convert_json_to_xl`)

1. `_analyze_json_columns(df)` — Inspects DataFrame to detect lists and dicts
2. **Merge behavior**: If output file exists, reads existing records and merges with new data:
   - Uses simple (scalar) columns as deduplication key
   - Overwrites existing rows with matching keys (keeps latest)
3. Expands arrays/objects into numbered/merged-header columns
4. Writes Excel with merged header cells for structured columns

### Helper Functions

- `_detect_xl_columns(ws)` — Returns classification of Excel columns (simple, flat_groups, struct_groups)
- `_analyze_json_columns(df)` — Returns classification of JSON columns (simple, flat_arrays, struct_arrays, single_structs)
- `_read_xl_as_records(xlsx_path)` — Reads existing Excel file back into record format for merging

## Key Behaviors

- **Deduplication**: `convert_json_to_xl` appends to existing Excel files, deduplicating by scalar column values
- **Structured detection**: Merged cells in row 1 indicate structured columns; data starts at row 3
- **Array flattening**: JSON lists are expanded into `name_0`, `name_1`, ... columns
- **Object nesting**: JSON objects (single or in arrays) use merged headers with sub-field names in row 2
- **Non-ASCII preservation**: All JSON output uses `ensure_ascii=False`
