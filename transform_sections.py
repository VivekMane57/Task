import json
from pathlib import Path
from typing import Any, Dict, List, Tuple, Optional, Set

BASE_DIR = Path(r"D:\Aadiswan Task")
OUTPUT_DIR = BASE_DIR / "output"


def clean_number(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s == "":
        return None
    if s in ("-", "--", "NA", "N/A"):
        return None
    s = s.replace(",", "")
    if s.endswith("%"):
        s = s[:-1]
    try:
        return float(s)
    except Exception:
        return value


def is_number_token(tok: str) -> bool:
    s = tok.strip()
    if s == "":
        return False
    s = s.replace(",", "")
    if s.endswith("%"):
        s = s[:-1]
    try:
        float(s)
        return True
    except Exception:
        return False


def detect_header_row(matrix: List[List[Any]]) -> Optional[int]:
    for i, row in enumerate(matrix):
        if not row:
            continue
        upper = [str(c).upper() if c else "" for c in row]
        joined = " ".join(upper)
        if (
            "FY 2023-24" in joined
            or "FY 2024-25" in joined
            or "FY 2025-26" in joined
            or "TTM" in joined
        ):
            return i
    return None


def spread_header_labels(header_row: List[Any]) -> List[str]:
    labels: List[str] = []
    current = ""
    for cell in header_row:
        text = str(cell).strip() if cell not in (None, "", " ") else ""
        if text:
            current = text
        labels.append(current)
    return labels


def extract_fy_columns(header_row: List[Any]) -> Dict[str, List[int]]:
    spread = spread_header_labels(header_row)
    fy_cols: Dict[str, List[int]] = {}
    for idx, text in enumerate(spread):
        if not text:
            continue
        t = text.upper()
        if "FY 2023-24" in t:
            fy_cols.setdefault("fy_2023_24", []).append(idx)
        if "FY 2024-25" in t:
            fy_cols.setdefault("fy_2024_25", []).append(idx)
        if "FY 2025-26" in t:
            fy_cols.setdefault("fy_2025_26", []).append(idx)
        if "TTM" in t:
            fy_cols.setdefault("ttm", []).append(idx)
    return fy_cols


def detect_section_title(
    matrix: List[List[Any]], header_idx: int, header_row: List[Any]
) -> Optional[str]:
    for i in range(header_idx - 1, -1, -1):
        row = matrix[i]
        if not row:
            continue
        cells = [str(c).strip() for c in row if c not in (None, "", " ")]
        if not cells:
            continue
        text = " ".join(cells)
        up = text.upper()
        if up in ("PARTICULARS", "MONTH"):
            continue
        return text
    for cell in header_row:
        if cell not in (None, "", " "):
            return str(cell).strip()
    return None


def slug(text: str) -> str:
    text = text.strip().lower()
    out: List[str] = []
    for ch in text:
        if ch.isalnum():
            out.append(ch)
        elif ch in (" ", "-", "/", "\\"):
            out.append("_")
    s = "".join(out)
    while "__" in s:
        s = s.replace("__", "_")
    s = s.strip("_")
    return s or "section"


def parse_fy_table(
    matrix: List[List[Any]],
    prev_context: Optional[Dict[str, Any]],
) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]], bool]:
    header_idx = detect_header_row(matrix)
    header_found = header_idx is not None

    if header_found:
        header_row = matrix[header_idx]
        fy_cols = extract_fy_columns(header_row)
        if not fy_cols:
            return None, prev_context, False
        title = detect_section_title(matrix, header_idx, header_row)
        start_data_row = header_idx + 1
        context = {"fy_cols": fy_cols, "title": title}
    else:
        if not prev_context:
            return None, prev_context, False
        fy_cols = prev_context["fy_cols"]
        title = prev_context["title"]
        start_data_row = 0
        context = prev_context

    records: List[Dict[str, Any]] = []

    for row in matrix[start_data_row:]:
        if not row:
            continue
        metric_raw = row[0] if len(row) > 0 else None
        if metric_raw is None or str(metric_raw).strip() == "":
            continue
        metric = str(metric_raw).strip()
        if metric.upper() == "PARTICULARS":
            continue
        item: Dict[str, Any] = {"metric": metric}
        for key, col_indexes in fy_cols.items():
            values: List[Any] = []
            for col_idx in col_indexes:
                v = row[col_idx] if col_idx < len(row) else None
                if v is None or str(v).strip() == "":
                    values.append(None)
                else:
                    values.append(clean_number(v))
            if not values:
                item[key] = None
            elif len(values) == 1:
                item[key] = values[0]
            else:
                item[key] = values
        records.append(item)

    if not records:
        return None, prev_context, header_found

    parsed_block = {
        "section_title": title,
        "metrics": records,
    }
    return parsed_block, context, header_found


def parse_state_wise_fy_table(
    matrix: List[List[Any]],
    prev_context: Optional[Dict[str, Any]],
) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]], bool]:
    header_idx = detect_header_row(matrix)
    header_found = header_idx is not None

    if header_found:
        header_row = matrix[header_idx]
        fy_cols = extract_fy_columns(header_row)
        if not fy_cols:
            return None, prev_context, False
        title = detect_section_title(matrix, header_idx, header_row)
        start_data_row = header_idx + 1
        context = {"fy_cols": fy_cols, "title": title}
    else:
        if not prev_context:
            return None, prev_context, False
        fy_cols = prev_context["fy_cols"]
        title = prev_context["title"]
        start_data_row = 0
        context = prev_context

    records: List[Dict[str, Any]] = []

    for row in matrix[start_data_row:]:
        if not row:
            continue
        metric_raw = row[0] if len(row) > 0 else None
        if metric_raw is None or str(metric_raw).strip() == "":
            continue
        metric = str(metric_raw).strip()
        if metric.upper() == "PARTICULARS":
            continue
        item: Dict[str, Any] = {"metric": metric}
        if len(row) > 1:
            scell = row[1]
            if scell not in (None, "", " "):
                item["state_code"] = str(scell).strip()
        for key, col_indexes in fy_cols.items():
            values: List[Any] = []
            for col_idx in col_indexes:
                v = row[col_idx] if col_idx < len(row) else None
                if v is None or str(v).strip() == "":
                    values.append(None)
                else:
                    values.append(clean_number(v))
            if not values:
                item[key] = None
            elif len(values) == 1:
                item[key] = values[0]
            else:
                item[key] = values
        records.append(item)

    if not records:
        return None, prev_context, header_found

    return {"section_title": title, "metrics": records}, context, header_found


def parse_product_wise_fy_table(
    matrix: List[List[Any]],
    prev_context: Optional[Dict[str, Any]],
) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]], bool]:
    header_idx = detect_header_row(matrix)
    header_found = header_idx is not None

    if header_found:
        header_row = matrix[header_idx]
        fy_cols = extract_fy_columns(header_row)
        if not fy_cols:
            return None, prev_context, False
        title = detect_section_title(matrix, header_idx, header_row)
        start_data_row = header_idx + 1
        context = {"fy_cols": fy_cols, "title": title}
    else:
        if not prev_context:
            return None, prev_context, False
        fy_cols = prev_context["fy_cols"]
        title = prev_context["title"]
        start_data_row = 0
        context = prev_context

    records: List[Dict[str, Any]] = []

    for row in matrix[start_data_row:]:
        if not row:
            continue
        code_cell = row[0] if len(row) > 0 else None
        if code_cell is None or str(code_cell).strip() == "":
            continue
        metric = str(code_cell).strip()
        if metric.upper() in ("PRODUCT (HSN)", "PRODUCT HSN"):
            continue
        item: Dict[str, Any] = {"product hsn ": metric}
        if len(row) > 1:
            hsn_cell = row[1]
            if hsn_cell not in (None, "", " "):
                item["hsn_name"] = str(hsn_cell).strip()
        for key, col_indexes in fy_cols.items():
            values: List[Any] = []
            for col_idx in col_indexes:
                v = row[col_idx] if col_idx < len(row) else None
                if v is None or str(v).strip() == "":
                    values.append(None)
                else:
                    values.append(clean_number(v))
            if not values:
                item[key] = None
            elif len(values) == 1:
                item[key] = values[0]
            else:
                item[key] = values
        records.append(item)

    if not records:
        return None, prev_context, header_found

    return {"section_title": title, "metrics": records}, context, header_found


def parse_monthly_particulars_table(
    matrix: List[List[Any]],
    months_context: Optional[List[str]],
) -> Tuple[Optional[Dict[str, Any]], Optional[List[str]], bool]:
    header_idx: Optional[int] = None
    header_row: Optional[List[Any]] = None
    particulars_col_index: Optional[int] = None

    for i, row in enumerate(matrix):
        if not row:
            continue
        for j, cell in enumerate(row):
            if cell is None:
                continue
            if "PARTICULARS" in str(cell).upper():
                header_idx = i
                header_row = row
                particulars_col_index = j
                break
        if header_idx is not None:
            break

    if (
        header_idx is not None
        and header_row is not None
        and particulars_col_index is not None
    ):
        months: List[str] = []
        for cell in header_row[particulars_col_index + 1 :]:
            if cell in (None, "", " "):
                continue
            months.append(str(cell).strip())
        if not months:
            return None, months_context, False
        title = detect_section_title(matrix, header_idx, header_row) or ""
        records: List[Dict[str, Any]] = []
        for row in matrix[header_idx + 1 :]:
            if not row:
                continue
            if all((c is None or str(c).strip() == "") for c in row):
                break
            label_cell = row[particulars_col_index]
            if label_cell is None or str(label_cell).strip() == "":
                continue
            metric = str(label_cell).strip()
            if metric.upper().startswith("PARTICULARS"):
                continue
            monthly_map: Dict[str, Any] = {}
            for idx, month in enumerate(months):
                col_idx = particulars_col_index + 1 + idx
                v = row[col_idx] if col_idx < len(row) else None
                if v is None or str(v).strip() == "":
                    val = None
                else:
                    val = clean_number(v)
                monthly_map[month] = val
            records.append({"metric": metric, "monthly_values": monthly_map})
        if not records:
            return None, months_context, False
        return {"section_title": title, "metrics": records}, months, True

    if months_context is not None:
        months = months_context
        first_label: Optional[str] = None
        records: List[Dict[str, Any]] = []
        for row in matrix:
            if not row:
                continue
            if all((c is None or str(c).strip() == "") for c in row):
                continue
            label_cell = row[0]
            if label_cell is None or str(label_cell).strip() == "":
                continue
            metric = str(label_cell).strip()
            if first_label is None:
                first_label = metric
            monthly_map: Dict[str, Any] = {}
            for idx, month in enumerate(months):
                col_idx = 1 + idx
                v = row[col_idx] if col_idx < len(row) else None
                if v is None or str(v).strip() == "":
                    val = None
                else:
                    val = clean_number(v)
                monthly_map[month] = val
            records.append({"metric": metric, "monthly_values": monthly_map})
        if not records:
            return None, months_context, False
        title = first_label or ""
        return {"section_title": title, "metrics": records}, months_context, True

    return None, months_context, False


def parse_simple_text_table(matrix: List[List[Any]]) -> Optional[Dict[str, Any]]:
    non_empty_rows: List[List[Any]] = []
    for row in matrix:
        if not row:
            continue
        cells = [str(c).strip() for c in row if c not in (None, "", " ")]
        if cells:
            non_empty_rows.append(cells)
    if not non_empty_rows:
        return None
    title = " ".join(non_empty_rows[0])
    metrics: List[Dict[str, Any]] = []
    for row in non_empty_rows[1:]:
        text = " ".join(row)
        if text:
            metrics.append({"metric": text})
    if not metrics:
        metrics.append({"metric": title})
    return {"section_title": title, "metrics": metrics}


def first_non_empty_text(matrix: List[List[Any]]) -> Optional[str]:
    for row in matrix:
        if not row:
            continue
        cells = [str(c).strip() for c in row if c not in (None, "", " ")]
        if cells:
            return " ".join(cells)
    return None


def parse_profile_block(matrix: List[List[Any]]) -> Optional[Dict[str, Any]]:
    profile_row = None
    for i, row in enumerate(matrix):
        if not row:
            continue
        first = row[0]
        if first is None or str(first).strip() == "":
            continue
        if str(first).strip().lower().startswith("profile"):
            profile_row = i
            break
    if profile_row is None:
        return None
    metrics: List[Dict[str, Any]] = []
    for j in range(profile_row + 1, len(matrix)):
        row = matrix[j]
        if not row:
            break
        if all((c is None or str(c).strip() == "") for c in row):
            break
        key = row[0] if len(row) > 0 else None
        val = row[1] if len(row) > 1 else None
        if (key is None or str(key).strip() == "") and (
            val is None or str(val).strip() == ""
        ):
            continue
        metrics.append(
            {
                "metric": "" if key is None else str(key).strip(),
                "value": None
                if val is None or str(val).strip() == ""
                else str(val).strip(),
            }
        )
    if not metrics:
        return None
    return {"section_title": "Profile", "metrics": metrics}


def parse_filing_block(matrix: List[List[Any]]) -> Optional[Dict[str, Any]]:
    if not matrix:
        return None
    title_idx = None
    title_text = ""
    for i, row in enumerate(matrix):
        if not row:
            continue
        cells = [str(c).strip() for c in row if c not in (None, "", " ")]
        if cells:
            title_idx = i
            title_text = " ".join(cells)
            break
    if title_idx is None:
        return None
    header_idx = None
    for j in range(title_idx + 1, len(matrix)):
        row = matrix[j]
        if not row:
            continue
        cells = [str(c).strip() for c in row if c not in (None, "", " ")]
        if cells:
            header_idx = j
            break
    if header_idx is None:
        return None
    header_row = matrix[header_idx]
    col_keys: List[str] = []
    for cell in header_row:
        if cell in (None, "", " "):
            col_keys.append("")
        else:
            col_keys.append(slug(str(cell)))
    records: List[Dict[str, Any]] = []
    for i in range(header_idx + 1, len(matrix)):
        row = matrix[i]
        if not row:
            break
        if all((c is None or str(c).strip() == "") for c in row):
            break
        rec: Dict[str, Any] = {}
        for j, key in enumerate(col_keys):
            if not key:
                continue
            value = row[j] if j < len(row) else None
            if value is None or str(value).strip() == "":
                rec[key] = None
            else:
                rec[key] = value
        if rec:
            records.append(rec)
    if not records:
        return None
    return {"section_title": title_text, "metrics": records}


def parse_adjusted_amounts_sheet(
    sheet_name: str, sheet_data: Dict[str, Any]
) -> Dict[str, Dict[str, Any]]:
    tables = sheet_data.get("tables", [])
    if not isinstance(tables, list) or not tables:
        return {}
    revenue_idx = None
    purchase_idx = None
    for idx, t in enumerate(tables):
        matrix = t.get("data")
        if not isinstance(matrix, list) or not matrix:
            continue
        title = first_non_empty_text(matrix) or ""
        up = title.upper()
        if "BIFURCATION OF REVENUE" in up and revenue_idx is None:
            revenue_idx = idx
        if "BIFURCATION OF PURCHASE AND EXPENSES" in up and purchase_idx is None:
            purchase_idx = idx
    if revenue_idx is None and purchase_idx is None:
        return {}
    revenue_tables: List[Dict[str, Any]] = []
    purchase_tables: List[Dict[str, Any]] = []
    for idx, t in enumerate(tables):
        if revenue_idx is not None and idx >= revenue_idx and (
            purchase_idx is None or idx < purchase_idx
        ):
            revenue_tables.append(t)
        if purchase_idx is not None and idx >= purchase_idx:
            purchase_tables.append(t)

    def build_block_with_fy(
        block_tables: List[Dict[str, Any]],
        default_title: str,
    ) -> Optional[Dict[str, Any]]:
        if not block_tables:
            return None
        combined: List[List[Any]] = []
        first_table = block_tables[0]
        for t in block_tables:
            matrix = t.get("data")
            if isinstance(matrix, list) and matrix:
                combined.extend(matrix)
        if not combined:
            return None
        parsed_block, _, header_found = parse_fy_table(combined, None)
        if not parsed_block or not header_found:
            return None
        metrics = parsed_block["metrics"]
        return {
            "section_title": default_title,
            "sheet": sheet_name,
            "start_row": first_table.get("start_row"),
            "start_col": first_table.get("start_col"),
            "row_count": len(metrics),
            "column_count": first_table.get("column_count"),
            "metrics": metrics,
        }

    result: Dict[str, Dict[str, Any]] = {}
    rev_block = build_block_with_fy(
        revenue_tables,
        "Bifurcation of Revenue (in INR)",
    )
    if rev_block:
        key = f"{sheet_name}_{slug(rev_block['section_title'])}"
        result[key] = rev_block
    pur_block = build_block_with_fy(
        purchase_tables,
        "Bifurcation of Purchase and Expenses (in INR)",
    )
    if pur_block:
        key = f"{sheet_name}_{slug(pur_block['section_title'])}"
        result[key] = pur_block
    return result


def parse_partywise_with_gstin(
    matrix: List[List[Any]],
    prev_context: Optional[Dict[str, Any]],
    role: str,
) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]], bool]:
    header_idx = detect_header_row(matrix)
    header_found = header_idx is not None
    if header_found:
        header_row = matrix[header_idx]
        fy_cols = extract_fy_columns(header_row)
        if not fy_cols:
            return None, prev_context, False
        title = detect_section_title(matrix, header_idx, header_row)
        start_data_row = header_idx + 1
        context = {"fy_cols": fy_cols, "title": title}
    else:
        if not prev_context:
            return None, prev_context, False
        fy_cols = prev_context["fy_cols"]
        title = prev_context["title"]
        start_data_row = 0
        context = prev_context
    records: List[Dict[str, Any]] = []
    name_key = f"{role} name " if role == "customer" else f"{role} name"
    gstin_key = f"{role}_gstin"
    for row in matrix[start_data_row:]:
        if not row:
            continue
        name_cell = row[0] if len(row) > 0 else None
        if name_cell is None or str(name_cell).strip() == "":
            continue
        name = str(name_cell).strip()
        if role == "customer" and "CUSTOMER" in name.upper():
            continue
        if role == "supplier" and "SUPPLIER" in name.upper():
            continue
        item: Dict[str, Any] = {name_key: name}
        gstin_cell = row[1] if len(row) > 1 else None
        if gstin_cell not in (None, "", " "):
            item[gstin_key] = str(gstin_cell).strip()
        for key, col_indexes in fy_cols.items():
            values: List[Any] = []
            for col_idx in col_indexes:
                v = row[col_idx] if col_idx < len(row) else None
                if v is None or str(v).strip() == "":
                    values.append(None)
                else:
                    values.append(clean_number(v))
            if not values:
                item[key] = None
            elif len(values) == 1:
                item[key] = values[0]
            else:
                item[key] = values
        records.append(item)
    if not records:
        return None, prev_context, header_found
    return {"section_title": title, "metrics": records}, context, header_found


def parse_customer_supplier_details_table(
    matrix: List[List[Any]],
) -> Optional[Dict[str, Any]]:
    if not matrix:
        return None
    title_idx = None
    for i, row in enumerate(matrix):
        if not row:
            continue
        cells = [str(c).strip() for c in row if c not in (None, "", " ")]
        if cells:
            title_idx = i
            break
    if title_idx is None:
        return None
    header_idx = None
    for i in range(title_idx + 1, len(matrix)):
        row = matrix[i]
        if not row:
            continue
        text0 = str(row[0]).upper() if row[0] not in (None, "", " ") else ""
        if "GSTN" in text0 or "GSTIN" in text0:
            header_idx = i
            break
    if header_idx is None:
        return None
    header_row = matrix[header_idx]
    col_keys: List[str] = []
    for cell in header_row:
        if cell in (None, "", " "):
            col_keys.append("")
            continue
        txt = str(cell).strip()
        up = txt.upper()
        if "CUSTOMER" in up and "GST" in up:
            col_keys.append("Customer GSTN")
        elif "SUPPLIER" in up and "GST" in up:
            col_keys.append("Supplier GSTN")
        else:
            col_keys.append(txt)
    records: List[Dict[str, Any]] = []
    for i in range(header_idx + 1, len(matrix)):
        row = matrix[i]
        if not row:
            break
        if all((c is None or str(c).strip() == "") for c in row):
            break
        rec: Dict[str, Any] = {}
        for j, key in enumerate(col_keys):
            if not key:
                continue
            value = row[j] if j < len(row) else None
            if value is None or str(value).strip() == "":
                rec[key] = None
            else:
                rec[key] = str(value).strip()
        if rec:
            records.append(rec)
    if not records:
        return None
    title_row = matrix[title_idx]
    title = " ".join(str(c).strip() for c in title_row if c not in (None, "", " "))
    return {"section_title": title, "metrics": records}


def parse_index_table(matrix: List[List[Any]]) -> Optional[Dict[str, Any]]:
    if not matrix:
        return None
    title = first_non_empty_text(matrix) or "Index"
    records: List[Dict[str, Any]] = []
    for row in matrix:
        if not row:
            continue
        code_cell = row[0] if len(row) > 0 else None
        title_cell = row[1] if len(row) > 1 else None
        desc_cell = row[2] if len(row) > 2 else None
        if (
            code_cell in (None, "", " ")
            and title_cell in (None, "", " ")
            and desc_cell in (None, "", " ")
        ):
            continue
        if code_cell is None or str(code_cell).strip() == "":
            continue
        code_text = str(code_cell).strip()
        up = code_text.upper()
        if "INDEX" in up or "GST ANALYTICS" in up or "GST DATA TABLES" in up:
            continue
        if not any(ch.isdigit() for ch in code_text):
            continue
        if isinstance(code_cell, (int, float)):
            code_text = f"{code_cell:.2f}".rstrip("0").rstrip(".")
        table_title = (
            str(title_cell).strip()
            if title_cell not in (None, "", " ")
            else None
        )
        description = (
            str(desc_cell).strip()
            if desc_cell not in (None, "", " ")
            else None
        )
        records.append(
            {
                "table_code": code_text,
                "table_title": table_title,
                "description": description,
            }
        )
    if not records:
        return None
    return {"section_title": title, "metrics": records}


def process_workbook_json(path: Path) -> Optional[Dict[str, Any]]:
    with open(path, "r", encoding="utf-8") as f:
        wb = json.load(f)
    sheets = wb.get("sheets", {})
    output: Dict[str, Any] = {
        "file_name": wb.get("file_name", path.name),
        "tables": {},
    }
    for sheet_name, sheet_data in sheets.items():
        tables = sheet_data.get("tables", [])
        if not isinstance(tables, list):
            continue
        normal_sheet = sheet_name.strip().lower()
        if normal_sheet == "adjusted amounts":
            special_tables = parse_adjusted_amounts_sheet(sheet_name, sheet_data)
            if special_tables:
                output["tables"].update(special_tables)
            continue
        used_keys: Set[str] = set()
        prev_context: Optional[Dict[str, Any]] = None
        last_key: Optional[str] = None
        months_context: Optional[List[str]] = None
        for t in tables:
            matrix = t.get("data")
            if not isinstance(matrix, list) or not matrix:
                continue
            parsed: Optional[Dict[str, Any]] = None
            header_found = False
            if normal_sheet.startswith("profile & filing"):
                title = first_non_empty_text(matrix) or ""
                low_title = title.lower()
                if low_title.startswith("profile"):
                    parsed = parse_profile_block(matrix)
                elif "filing details - gstr3b" in low_title:
                    parsed = parse_filing_block(matrix)
                elif "filing details - gstr1" in low_title:
                    parsed = parse_filing_block(matrix)
                else:
                    parsed = parse_simple_text_table(matrix)
                header_found = parsed is not None
                prev_context = None
            elif normal_sheet == "details of customers and supp.":
                parsed = parse_customer_supplier_details_table(matrix)
                if not parsed:
                    parsed = parse_simple_text_table(matrix)
                header_found = parsed is not None
                prev_context = None
            elif normal_sheet == "index":
                parsed = parse_index_table(matrix)
                if not parsed:
                    parsed = parse_simple_text_table(matrix)
                header_found = parsed is not None
                prev_context = None
            else:
                if normal_sheet in ("gstr 3b", "tax", "summary"):
                    parsed, months_context, header_found = parse_monthly_particulars_table(
                        matrix, months_context
                    )
                    if not parsed:
                        parsed, prev_context, header_found = parse_fy_table(
                            matrix, prev_context
                        )
                        if not parsed:
                            parsed = parse_simple_text_table(matrix)
                            header_found = parsed is not None
                            prev_context = None
                elif normal_sheet == "state wise":
                    parsed, prev_context, header_found = parse_state_wise_fy_table(
                        matrix, prev_context
                    )
                    if not parsed:
                        parsed = parse_simple_text_table(matrix)
                        header_found = parsed is not None
                        prev_context = None
                elif normal_sheet == "product wise":
                    parsed, prev_context, header_found = parse_product_wise_fy_table(
                        matrix, prev_context
                    )
                    if not parsed:
                        parsed = parse_simple_text_table(matrix)
                        header_found = parsed is not None
                        prev_context = None
                elif normal_sheet == "customer wise":
                    parsed, prev_context, header_found = parse_partywise_with_gstin(
                        matrix, prev_context, role="customer"
                    )
                    if not parsed:
                        parsed, prev_context, header_found = parse_fy_table(
                            matrix, prev_context
                        )
                        if not parsed:
                            parsed = parse_simple_text_table(matrix)
                            header_found = parsed is not None
                            prev_context = None
                elif normal_sheet == "supplier wise":
                    parsed, prev_context, header_found = parse_partywise_with_gstin(
                        matrix, prev_context, role="supplier"
                    )
                    if not parsed:
                        parsed, prev_context, header_found = parse_fy_table(
                            matrix, prev_context
                        )
                        if not parsed:
                            parsed = parse_simple_text_table(matrix)
                            header_found = parsed is not None
                            prev_context = None
                else:
                    parsed, prev_context, header_found = parse_fy_table(
                        matrix, prev_context
                    )
                    if not parsed:
                        parsed = parse_simple_text_table(matrix)
                        header_found = parsed is not None
                        prev_context = None
            if not parsed:
                continue
            if header_found or last_key is None:
                section_title = (
                    parsed.get("section_title") or f"table_{t.get('table_index')}"
                )
                section_slug = slug(section_title)
                base_key = f"{sheet_name}_{section_slug}"
                key = base_key
                idx = 2
                while key in used_keys or key in output["tables"]:
                    key = f"{base_key}_{idx}"
                    idx += 1
                used_keys.add(key)
                last_key = key
                output["tables"][key] = {
                    "sheet": sheet_name,
                    "section_title": section_title,
                    "start_row": t.get("start_row"),
                    "start_col": t.get("start_col"),
                    "row_count": t.get("row_count"),
                    "column_count": t.get("column_count"),
                    "metrics": parsed["metrics"],
                }
            else:
                if last_key is not None:
                    output["tables"][last_key]["metrics"].extend(parsed["metrics"])
    if not output["tables"]:
        return None
    return output


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    workbook_jsons = sorted(
        p for p in OUTPUT_DIR.glob("*.json") if not p.name.startswith("structured_")
    )
    if not workbook_jsons:
        print("[ERROR] No workbook JSON files found.")
        return
    for path in workbook_jsons:
        print(f"[INFO] Processing: {path.name}")
        structured = process_workbook_json(path)
        if structured is None:
            print(f"[WARN] No tables parsed in: {path.name}")
            continue
        out_path = OUTPUT_DIR / f"structured_{path.name}"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(structured, f, indent=2, ensure_ascii=False)
        print(f"[OK] {out_path}")
    print("[DONE]")


if __name__ == "__main__":
    main()
