import pandas as pd
import os
import re


def convert_xl_to_json(xlsx_path: str, output_dir: str) -> None:
    """Convert Excel file to JSON."""
    df = pd.read_excel(xlsx_path)

    # Collapse numbered columns (e.g. col_0, col_1, ...) back into list columns
    groups: dict[str, list[str]] = {}
    for col in df.columns:
        m = re.match(r"^(.+)_(\d+)$", col)
        if m:
            base = m.group(1)
            groups.setdefault(base, []).append(col)
    for base, cols in groups.items():
        # Skip if a non-numbered column with the same base name exists
        if base not in df.columns:
            cols.sort(key=lambda c: int(c[len(base) + 1:]))
            # Merge numbered columns into a single list column, dropping NaN values
            df[base] = df[cols].apply(
                lambda row: [v for v in row if pd.notna(v)], axis=1
            )
            df.drop(columns=cols, inplace=True)

    # Write JSON output
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(xlsx_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.json")
    df.to_json(output_path, orient="records", force_ascii=False, indent=2)

def convert_json_to_xl(json_path: str, output_dir: str) -> None:
    """Convert JSON file to Excel."""
    df = pd.read_json(json_path)

    # Expand list columns into numbered columns (e.g. col -> col_0, col_1, ...)
    for col in df.columns:
        if df[col].apply(lambda x: isinstance(x, list)).any():
            expanded = df[col].apply(lambda x: x if isinstance(x, list) else [])
            max_len = expanded.apply(len).max()
            for i in range(max_len):
                df[f"{col}_{i}"] = expanded.apply(lambda x: x[i] if i < len(x) else None)
            df.drop(columns=[col], inplace=True)

    # Write Excel output
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(json_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.xlsx")

    # Append to existing file if present, deduplicating rows
    if os.path.exists(output_path):
        existing_df = pd.read_excel(output_path)
        df = pd.concat([existing_df, df], ignore_index=True).drop_duplicates(keep="last")
    df.to_excel(output_path, index=False)
