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

## List Conversion

List (array) fields are automatically converted between JSON and Excel formats in both directions.

- **JSON → Excel**: A list column like `tags: ["a", "b", "c"]` is expanded into separate numbered columns: `tags_0`, `tags_1`, `tags_2`.
- **Excel → JSON**: Numbered columns like `tags_0`, `tags_1`, `tags_2` are collapsed back into a single list field: `tags: ["a", "b", "c"]`. Empty (`NaN`) values are excluded from the list.

### Example

**Excel** (`data.xlsx`)

| name  | tags_0 | tags_1 | tags_2 |
|-------|--------|--------|--------|
| Alice | python | pandas |        |
| Bob   | java   | spring | docker |

**JSON** (`data.json`)

```json
[
  { "name": "Alice", "tags": ["python", "pandas"] },
  { "name": "Bob", "tags": ["java", "spring", "docker"] }
]
```
