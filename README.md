# XLtoJSON

A Python CLI tool for bidirectional conversion between Excel (`.xlsx`) and JSON files.

## Setup

```bash
python -m venv venv
venv\Scripts\activate
pip install pandas openpyxl
```

## Usage

```bash
# Excel to JSON
python test.py source.xlsx output/

# JSON to Excel
python test.py source.json output/
```

- **Excel → JSON**: Converts rows to a JSON array of records with non-ASCII characters preserved.
- **JSON → Excel**: Flattens array fields into numbered columns and merges with existing output files (deduplicates, keeps latest).
