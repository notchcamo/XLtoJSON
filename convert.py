import pandas as pd
import os


def convertXLtoJSON(xlsx_path: str, output_dir: str) -> None:
    df = pd.read_excel(xlsx_path)
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(xlsx_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.json")
    df.to_json(output_path, orient="records", force_ascii=False, indent=2)

def convertJSONtoXL(json_path: str, output_dir: str) -> None:
    df = pd.read_json(json_path)
    for col in df.columns:
        if df[col].apply(lambda x: isinstance(x, list)).any():
            expanded = df[col].apply(lambda x: x if isinstance(x, list) else [])
            max_len = expanded.apply(len).max()
            for i in range(max_len):
                df[f"{col}_{i}"] = expanded.apply(lambda x: x[i] if i < len(x) else None)
            df.drop(columns=[col], inplace=True)
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(json_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.xlsx")
    if os.path.exists(output_path):
        existing_df = pd.read_excel(output_path)
        df = pd.concat([existing_df, df], ignore_index=True).drop_duplicates(keep="last")
    df.to_excel(output_path, index=False)
