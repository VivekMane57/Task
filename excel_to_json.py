import json
from pathlib import Path
from typing import List, Tuple, Dict, Any

import pandas as pd
import math

BASE_DIR = Path(r"D:\Aadiswan Task")
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
FILE_PATTERN = "*.xlsx"


def is_empty_value(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def is_empty_row(row: pd.Series) -> bool:
    for v in row:
        if not is_empty_value(v):
            return False
    return True


def split_into_tables(
    df_raw: pd.DataFrame
) -> List[Tuple[int, int, int, int, pd.DataFrame]]:
    """
    Split a sheet into blocks separated by fully empty rows.
    """
    df = df_raw.astype(object)
    df = df.where(pd.notnull(df), None)

    tables: List[Tuple[int, int, int, int, pd.DataFrame]] = []
    current_start: int | None = None

    for idx, row in df.iterrows():
        if is_empty_row(row):
            if current_start is not None:
                start_r = current_start
                end_r = idx - 1
                block = df.loc[start_r:end_r]
                block = block.dropna(axis=1, how="all")
                if not block.empty:
                    start_c = block.columns[0]
                    end_c = block.columns[-1]
                    tables.append((start_r, start_c, end_r, end_c, block))
                current_start = None
        else:
            if current_start is None:
                current_start = idx

    # last block
    if current_start is not None:
        start_r = current_start
        end_r = df.index[-1]
        block = df.loc[start_r:end_r]
        block = block.dropna(axis=1, how="all")
        if not block.empty:
            start_c = block.columns[0]
            end_c = block.columns[-1]
            tables.append((start_r, start_c, end_r, end_c, block))

    return tables


def table_to_json_entry(
    table_idx: int,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    block: pd.DataFrame,
) -> Dict[str, Any]:
    matrix = block.values.tolist()

    for i in range(len(matrix)):
        for j in range(len(matrix[i])):
            v = matrix[i][j]
            if isinstance(v, float) and math.isnan(v):
                matrix[i][j] = None

    return {
        "table_index": table_idx,
        "start_row": int(start_row) + 1,
        "start_col": int(start_col) + 1,
        "end_row": int(end_row) + 1,
        "end_col": int(end_col) + 1,
        "row_count": int(block.shape[0]),
        "column_count": int(block.shape[1]),
        "data": matrix,
    }


def workbook_to_json(path: Path) -> Dict[str, Any]:
    print(f"[INFO] Processing: {path.name}")
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheets_json: Dict[str, Any] = {}

    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name=sheet_name, header=None)
        tables = split_into_tables(df)

        sheet_entry = {
            "table_count": len(tables),
            "tables": []
        }

        for idx, (sr, sc, er, ec, block) in enumerate(tables, start=1):
            sheet_entry["tables"].append(
                table_to_json_entry(idx, sr, sc, er, ec, block)
            )

        sheets_json[sheet_name] = sheet_entry

    return {
        "file_name": path.name,
        "sheets": sheets_json,
    }


def save_workbook_json(wb_json: Dict[str, Any], excel_file: Path):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_file = OUTPUT_DIR / excel_file.with_suffix(".json").name
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(wb_json, f, indent=2, ensure_ascii=False)
    print(f"[OK] JSON created: {out_file}")


def main():
    excel_files = sorted(DATA_DIR.glob(FILE_PATTERN))
    if not excel_files:
        print("[ERROR] No Excel files found in 'data' folder.")
        return

    for excel in excel_files:
        wb_json = workbook_to_json(excel)
        save_workbook_json(wb_json, excel)

    print("[DONE] All Excel files converted.")


if __name__ == "__main__":
    main()
