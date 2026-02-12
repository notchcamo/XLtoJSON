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

## Structured Cell Conversion

Structured (nested object) fields are automatically converted using merged header cells in Excel.

- **JSON → Excel**: Object fields use merged cells in row 1 for the field name, with sub-field names in row 2. Data starts at row 3.
- **Excel → JSON**: Merged header cells indicate structured fields. Sub-columns are collapsed back into nested objects.

### Single Object Example

**Excel** (`users.xlsx`)

Row 1 has merged cells. For example, `address` is merged across columns B-C.

| name  | address (merged) | |
|-------|----------|------|
|       | street   | city |
| Alice | 123 Main St | NYC |
| Bob   | 456 Oak Ave | LA  |

**JSON** (`users.json`)

```json
[
  {
    "name": "Alice",
    "address": {
      "street": "123 Main St",
      "city": "NYC"
    }
  },
  {
    "name": "Bob",
    "address": {
      "street": "456 Oak Ave",
      "city": "LA"
    }
  }
]
```

### Array of Objects Example

**Excel** (`orders.xlsx`)

Row 1 has merged cells with `_N` suffixes. For example, `items_0` is merged across columns B-C, `items_1` is merged across columns D-E.

| customer | items_0 (merged) | | items_1 (merged) | |
|----------|------|-------|------|-------|
|          | name | price | name | price |
| Alice    | Book | 12.99 | Pen  | 2.50  |
| Bob      | Mug  | 8.00  |      |       |

**JSON** (`orders.json`)

```json
[
  {
    "customer": "Alice",
    "items": [
      { "name": "Book", "price": 12.99 },
      { "name": "Pen", "price": 2.5 }
    ]
  },
  {
    "customer": "Bob",
    "items": [
      { "name": "Mug", "price": 8.0 }
    ]
  }
]
```
