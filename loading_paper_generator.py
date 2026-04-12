import math
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import xlwings as xw


# =============================
# SETTINGS (EDIT THIS SECTION)
# =============================
ORDERS_FILE = r"E:\Marbah Products\Marbah Invoices script\orders_data.xlsx"
LOADING_TEMPLATE_FILE = r"E:\Marbah Products\Marbah Invoices script\invoice_template.xlsx"
OUTPUT_ROOT = r"E:\Marbah Products\Marbah Invoices script"


@dataclass
class LoadingPaperConfig:
    orders_file: str = ORDERS_FILE
    template_file: str = LOADING_TEMPLATE_FILE
    output_root: str = OUTPUT_ROOT

    loading_sheet_candidates: Tuple[str, ...] = ("Loading paper", "loading paper", "loading")

    # Source columns
    route_agent_candidates: Tuple[str, ...] = ("route_agent", "driver")
    run_candidates: Tuple[str, ...] = ("run", "trip", "run_name")
    delivery_date_candidates: Tuple[str, ...] = ("estimated_delivery_date", "delivery_date")
    brand_candidates: Tuple[str, ...] = ("brand_name",)
    size_candidates: Tuple[str, ...] = (
        "size",
        "pack_size",
        "sku_size",
        "item_size",
        "variant_size",
    )
    product_name_candidates: Tuple[str, ...] = (
        "sku_name",
        "item_name",
        "product_name",
        "product_name_ar",
        "item",
        "description",
    )
    qty_candidates: Tuple[str, ...] = ("purchased_item_count", "qty", "quantity", "item_qty")

    # Template layout
    start_row: int = 3
    end_row: int = 140
    summary_anchor_row: int = 141
    left_name_col: int = 1
    left_qty_col: int = 3
    right_name_col: int = 4
    right_qty_col: int = 6

    # Optional header cells in loading sheet
    route_agent_cell: str = "B2"
    run_cell: str = "E2"
    date_cell: str = "A2"
    total_qty_cell: str = "C143"

    include_brand_subtotal_row: bool = False

    brand_order: Dict[str, int] = None
    size_order: Dict[str, int] = None

    def __post_init__(self) -> None:
        if self.brand_order is None:
            self.brand_order = {
                "lays": 1,
                "lays max": 2,
                "crunchy": 3,
                "cheetos": 4,
                "doritos": 5,
                "alyoum": 6,
                "pasta": 7,
                "nodules": 8,
                "pepsi": 9,
                "youmy juice": 10,
                "aquafina": 11,
            }
        if self.size_order is None:
            self.size_order = {
                "صغير": 1,
                "وسط": 2,
                "كبير": 3,
                "ميكا": 4,
                "70 غم": 5,
                "50غم": 6,
                "54 غم": 7,
                "45 غم": 8,
                "400 غم": 9,
                "200 غم": 10,
                "1.6 كغم": 11,
                "185 مل": 12,
                "750 مل": 13,
                "330 مل": 14,
                "1.25 لتر": 15,
                "1.75 لتر": 16,
                "250 مل": 17,
                "300 مل": 18,
                "250مل": 19,
                "200 مل": 20,
                "1 لتر": 21,
                "180 مل": 22,
                "1.5 لتر": 23,
                "500 مل": 24,
            }


COLUMN_ALIASES: Dict[str, List[str]] = {
    "route_agent": ["driver"],
    "run": ["trip", "run_name"],
    "estimated_delivery_date": ["delivery_date"],
    "brand_name": ["brand"],
    "size": ["pack_size", "sku_size", "item_size", "variant_size"],
    "sku_name": ["item_name", "product_name", "product_name_ar", "item", "description"],
    "purchased_item_count": ["qty", "quantity", "item_qty"],
}


def log_info(msg: str) -> None:
    print(f"INFO: {msg}")


def log_warning(msg: str) -> None:
    print(f"WARNING: {msg}")


def safe_str(value: object, default: str = "") -> str:
    if pd.isna(value):
        return default
    s = str(value).strip()
    if s.lower() in {"nan", "none", "null", "nat"}:
        return default
    return s


def safe_float(value: object, default: float = 0.0) -> float:
    try:
        if pd.isna(value):
            return default
        x = str(value).strip()
        arabic_digits = "٠١٢٣٤٥٦٧٨٩"
        for i, d in enumerate(arabic_digits):
            x = x.replace(d, str(i))
        x = x.replace(",", "")
        x = re.sub(r"[^0-9.\-]", "", x)
        if x in {"", ".", "-", "-.", ".-"}:
            return default
        return float(x)
    except Exception:
        return default


def normalize_identifier(value: object, default: str = "") -> str:
    text = safe_str(value, default)
    if not text:
        return default
    text = text.replace(",", "").strip()
    if re.fullmatch(r"[+-]?\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def normalize_brand_for_order(brand_value: object) -> str:
    # Keep brand_name-driven grouping, but normalize formatting for reliable ordering.
    text = safe_str(brand_value, "").lower().strip()
    text = text.replace("’", "'")
    text = re.sub(r"\s+", " ", text)
    text = text.replace("'", "")
    return text


def normalize_size_for_order(size_value: object) -> str:
    return re.sub(r"\s+", " ", safe_str(size_value, "")).strip()


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [safe_str(c, "").lower() for c in df.columns]
    return df


def apply_column_aliases(df: pd.DataFrame, aliases: Dict[str, List[str]]) -> pd.DataFrame:
    rename_map: Dict[str, str] = {}
    existing = set(df.columns)
    for target, candidates in aliases.items():
        if target in existing:
            continue
        for candidate in candidates:
            candidate_norm = safe_str(candidate, "").lower()
            if candidate_norm in existing:
                rename_map[candidate_norm] = target
                break
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def find_existing_column(df: pd.DataFrame, candidates: Tuple[str, ...]) -> Optional[str]:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def get_sheet_if_exists(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def get_loading_sheet(workbook: xw.Book, candidates: Tuple[str, ...]) -> Optional[xw.Sheet]:
    for name in candidates:
        ws = get_sheet_if_exists(workbook, name)
        if ws is not None:
            return ws
    if workbook.sheets.count > 0:
        log_warning(f"Loading sheet not found by name. Using first sheet '{workbook.sheets[0].name}'.")
        return workbook.sheets[0]
    return None


def copy_sheet(template_ws: xw.Sheet, output_wb: xw.Book, temp_sheet_name: str) -> Optional[xw.Sheet]:
    before = [s.name for s in output_wb.sheets]
    template_ws.api.Copy(Before=output_wb.sheets[temp_sheet_name].api)
    for sheet in output_wb.sheets:
        if sheet.name not in before:
            return sheet
    return None


def make_unique_sheet_name(base_name: str, existing_names: set) -> str:
    clean = re.sub(r"[\[\]\*\/\\\?\:]", "", safe_str(base_name, "LOAD"))[:25]
    candidate = f"LOAD_{clean}" if clean else "LOAD"
    i = 1
    while candidate in existing_names:
        candidate = f"LOAD_{clean}_{i}"[:31]
        i += 1
    existing_names.add(candidate)
    return candidate[:31]


def adjust_loading_capacity(ws: xw.Sheet, required_rows: int, config: LoadingPaperConfig) -> int:
    """
    Resize loading lines area to required row count.
    - Insert rows before summary anchor if required rows exceed template capacity.
    - Delete unused rows before summary anchor if required rows are fewer.

    Returns signed row shift applied to rows below anchor.
    """
    base_capacity = config.end_row - config.start_row + 1
    target = max(0, required_rows)
    requested_shift = target - base_capacity

    if requested_shift > 0:
        inserted = 0
        for _ in range(requested_shift):
            try:
                ws.api.Rows(config.summary_anchor_row).Insert()
                inserted += 1
            except Exception as exc:
                log_warning(f"Unable to insert extra loading rows before summary: {exc}")
                break
        return inserted

    if requested_shift < 0:
        delete_count = abs(requested_shift)
        delete_start_row = config.start_row + target
        deleted = 0
        for _ in range(delete_count):
            try:
                ws.api.Rows(delete_start_row).Delete()
                deleted += 1
            except Exception as exc:
                log_warning(f"Unable to delete unused loading row before summary: {exc}")
                break
        return -deleted

    return 0


def build_brand_rows(group_df: pd.DataFrame, config: LoadingPaperConfig) -> Tuple[List[List[Dict[str, object]]], float]:
    brand_blocks: List[List[Dict[str, object]]] = []
    if group_df.empty:
        return brand_blocks, 0.0

    grouped = (
        group_df.groupby(
            ["_brand_display", "_size_display", "_brand_order", "_size_order", "_product"],
            dropna=False,
        )["_qty"]
        .sum()
        .reset_index()
        .sort_values(["_brand_order", "_size_order", "_product"], kind="stable")
    )
    grand_total = float(grouped["_qty"].sum())

    for (brand_name, _brand_order), brand_df in grouped.groupby(
        ["_brand_display", "_brand_order"], sort=False
    ):
        brand_rows: List[Dict[str, object]] = []
        brand_clean = safe_str(brand_name, "OTHER")
        brand_total = float(brand_df["_qty"].sum())
        # Brand separator row.
        brand_rows.append({"kind": "brand_separator", "name": brand_clean, "qty": brand_total})

        for (size_name, _size_order), section_df in brand_df.groupby(
            ["_size_display", "_size_order"], sort=False
        ):
            size_clean = safe_str(size_name, "").strip()
            section_label = f"{brand_clean} - {size_clean}" if size_clean else brand_clean
            section_total = float(section_df["_qty"].sum())
            brand_rows.append({"kind": "size_separator", "name": section_label, "qty": section_total})

            for _, rec in section_df.iterrows():
                prod = safe_str(rec["_product"], "Unnamed Item")
                qty = safe_float(rec["_qty"], 0.0)
                brand_rows.append({"kind": "item", "name": prod, "qty": qty})

            if config.include_brand_subtotal_row:
                brand_rows.append({"kind": "subtotal", "name": f"اجمالي {section_label}", "qty": section_total})

        brand_blocks.append(brand_rows)

    return brand_blocks, grand_total


def style_loading_row(ws: xw.Sheet, row_no: int, name_col: int, qty_col: int, row_kind: str) -> None:
    name_cell = ws.range((row_no, name_col))
    qty_cell = ws.range((row_no, qty_col))
    row_range = ws.range((row_no, name_col), (row_no, qty_col))
    row_range.api.Font.Name = "Calibri"
    row_range.api.Font.Size = 10

    if row_kind == "brand_separator":
        row_range.api.Font.Bold = True
        row_range.color = (47, 117, 181)
        row_range.api.Font.Color = 0xFFFFFF
    elif row_kind == "size_separator":
        row_range.api.Font.Bold = True
        row_range.color = (198, 224, 180)
    elif row_kind == "brand_gap":
        row_range.api.Font.Bold = False
        row_range.color = None
    elif row_kind == "subtotal":
        row_range.api.Font.Bold = True
        row_range.color = (217, 217, 217)
    else:
        row_range.api.Font.Bold = False
        row_range.color = None

    name_cell.api.HorizontalAlignment = -4152  # xlRight
    qty_cell.api.HorizontalAlignment = -4108  # xlCenter
    qty_cell.number_format = "#,##0"


def clear_loading_column_block(ws: xw.Sheet, start_row: int, end_row: int, name_col: int, qty_col: int) -> None:
    ws.range((start_row, name_col), (end_row, qty_col)).value = None
    ws.range((start_row, name_col), (end_row, qty_col)).color = None
    ws.range((start_row, name_col), (end_row, qty_col)).api.Font.Bold = False
    ws.range((start_row, name_col), (end_row, qty_col)).api.Font.Color = 0


def clear_loading_data_area(ws: xw.Sheet, start_row: int, end_row: int) -> None:
    """
    Remove all existing data in template lines area (A:F).
    Keeps workbook structure and formulas outside the lines area.
    """
    ws.range((start_row, 1), (end_row, 6)).value = None
    ws.range((start_row, 1), (end_row, 6)).color = None
    ws.range((start_row, 1), (end_row, 6)).api.Font.Bold = False
    ws.range((start_row, 1), (end_row, 6)).api.Font.Color = 0


def write_loading_column(
    ws: xw.Sheet,
    rows: List[Dict[str, object]],
    start_row: int,
    end_row: int,
    name_col: int,
    qty_col: int,
) -> None:
    clear_loading_column_block(ws, start_row, end_row, name_col, qty_col)
    for idx, row in enumerate(rows):
        row_no = start_row + idx
        if row_no > end_row:
            break
        row_kind = safe_str(row.get("kind", "item"), "item")
        ws.range((row_no, name_col)).value = row["name"]
        if row_kind == "brand_gap":
            ws.range((row_no, qty_col)).value = ""
        else:
            ws.range((row_no, qty_col)).value = safe_float(row["qty"], 0.0)
        style_loading_row(ws, row_no, name_col, qty_col, row_kind)


def split_brand_blocks_two_columns(
    brand_blocks: List[List[Dict[str, object]]],
) -> Tuple[List[Dict[str, object]], List[Dict[str, object]]]:
    """
    Split brand groups into left/right columns without breaking a brand block.
    Each brand block (brand separator + all size sections/items) stays intact
    in one column, and groups are split by count.
    """
    left_rows: List[Dict[str, object]] = []
    right_rows: List[Dict[str, object]] = []

    split_idx = math.ceil(len(brand_blocks) / 2)
    left_blocks = brand_blocks[:split_idx]
    right_blocks = brand_blocks[split_idx:]

    for idx, block in enumerate(left_blocks):
        if idx > 0:
            left_rows.append({"kind": "brand_gap", "name": "", "qty": None})
        left_rows.extend(block)

    for idx, block in enumerate(right_blocks):
        if idx > 0:
            right_rows.append({"kind": "brand_gap", "name": "", "qty": None})
        right_rows.extend(block)

    return left_rows, right_rows


def shift_cell_ref_if_below_anchor(cell_ref: str, row_shift: int, anchor_row: int) -> str:
    """
    Shift a cell reference row by row_shift if its row is >= anchor_row.
    Example: C143 with row_shift=-10, anchor 141 => C133
    """
    match = re.fullmatch(r"([A-Za-z]+)(\d+)", safe_str(cell_ref, ""))
    if not match:
        return cell_ref
    col_letters = match.group(1)
    row_no = int(match.group(2))
    if row_no >= anchor_row:
        row_no = max(1, row_no + row_shift)
    return f"{col_letters}{row_no}"


def write_loading_sheet(
    ws: xw.Sheet,
    route_agent: str,
    run_name: str,
    group_df: pd.DataFrame,
    config: LoadingPaperConfig,
) -> None:
    brand_blocks, grand_total = build_brand_rows(group_df, config)
    left_rows, right_rows = split_brand_blocks_two_columns(brand_blocks)
    required_rows = max(len(left_rows), len(right_rows))

    row_shift = adjust_loading_capacity(ws, required_rows, config)
    end_row = config.end_row + row_shift

    # Remove all existing template data in A:F lines area before writing new rows.
    clear_loading_data_area(ws, config.start_row, end_row)

    write_loading_column(
        ws,
        left_rows,
        config.start_row,
        end_row,
        config.left_name_col,
        config.left_qty_col,
    )
    write_loading_column(
        ws,
        right_rows,
        config.start_row,
        end_row,
        config.right_name_col,
        config.right_qty_col,
    )

    # Optional header metadata
    max_date = pd.to_datetime(group_df["_delivery_date"], errors="coerce").max()
    ws[config.route_agent_cell].value = route_agent
    ws[config.run_cell].value = run_name
    ws[config.date_cell].value = max_date.to_pydatetime() if pd.notna(max_date) else ""
    ws[config.date_cell].number_format = "yyyy-mm-dd"
    total_qty_cell = shift_cell_ref_if_below_anchor(config.total_qty_cell, row_shift, config.summary_anchor_row)
    ws[total_qty_cell].value = grand_total
    ws[total_qty_cell].number_format = "#,##0"


def load_orders_dataframe(config: LoadingPaperConfig) -> pd.DataFrame:
    if not os.path.exists(config.orders_file):
        raise FileNotFoundError(f"Orders file not found: {config.orders_file}")

    df = pd.read_excel(config.orders_file)
    df = normalize_columns(df)
    df = apply_column_aliases(df, COLUMN_ALIASES)

    route_col = find_existing_column(df, config.route_agent_candidates)
    run_col = find_existing_column(df, config.run_candidates)
    delivery_col = find_existing_column(df, config.delivery_date_candidates)
    brand_col = find_existing_column(df, config.brand_candidates)
    size_col = find_existing_column(df, config.size_candidates)
    product_col = find_existing_column(df, config.product_name_candidates)
    qty_col = find_existing_column(df, config.qty_candidates)

    if route_col is None or run_col is None or product_col is None or qty_col is None:
        raise ValueError(
            "Missing required columns for loading paper generation. "
            "Need route_agent/driver, run/trip, product name, and quantity."
        )

    brand_rank_map = {
        normalize_brand_for_order(key): rank
        for key, rank in (config.brand_order or {}).items()
    }
    size_rank_map = {
        normalize_size_for_order(key): rank
        for key, rank in (config.size_order or {}).items()
    }

    out = df.copy()
    out["_route_agent"] = out[route_col].apply(lambda x: safe_str(x, "UNKNOWN-AGENT"))
    out["_run"] = out[run_col].apply(lambda x: normalize_identifier(x, "UNKNOWN-RUN"))
    out["_product"] = out[product_col].apply(lambda x: safe_str(x, "Unnamed Item"))
    out["_qty"] = out[qty_col].apply(lambda x: safe_float(x, 0.0))
    out["_brand_raw"] = (
        out[brand_col].apply(lambda x: safe_str(x, "OTHER"))
        if brand_col is not None
        else out["_product"].apply(lambda x: safe_str(x.split(" ")[0], "OTHER"))
    )
    out["_size_raw"] = out[size_col].apply(lambda x: safe_str(x, "")) if size_col is not None else ""
    out["_brand_key"] = out["_brand_raw"].apply(normalize_brand_for_order)
    out["_size_key"] = out["_size_raw"].apply(normalize_size_for_order)
    out["_brand_display"] = out["_brand_raw"].apply(lambda x: safe_str(x, "OTHER"))
    out["_size_display"] = out["_size_raw"].apply(lambda x: safe_str(x, "").replace("  ", " ").strip())
    out["_brand_order"] = out["_brand_key"].apply(lambda x: brand_rank_map.get(safe_str(x, ""), 999))
    out["_size_order"] = out["_size_key"].apply(lambda x: size_rank_map.get(safe_str(x, ""), 999))
    if delivery_col is not None:
        out["_delivery_date"] = pd.to_datetime(out[delivery_col], errors="coerce")
    else:
        out["_delivery_date"] = pd.NaT

    out = out[(out["_route_agent"] != "") & (out["_run"] != "")]
    out = out[out["_qty"] != 0]
    return out


def generate_loading_papers() -> None:
    config = LoadingPaperConfig()
    df = load_orders_dataframe(config)
    if df.empty:
        log_info("No rows available to generate loading papers.")
        return

    if not os.path.exists(config.template_file):
        raise FileNotFoundError(f"Template file not found: {config.template_file}")

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    template_wb = None
    output_wb = None
    try:
        template_wb = app.books.open(config.template_file)
        template_ws = get_loading_sheet(template_wb, config.loading_sheet_candidates)
        if template_ws is None:
            raise ValueError("Loading sheet not found in template workbook.")

        output_wb = app.books.add()
        output_wb.sheets[0].name = "TempSheet"
        existing_sheet_names = set()

        grouped = df.groupby(["_route_agent", "_run"], sort=True)
        for (route_agent, run_name), group_df in grouped:
            new_ws = copy_sheet(template_ws, output_wb, "TempSheet")
            if new_ws is None:
                log_warning(f"Could not create loading sheet for route '{route_agent}' run '{run_name}'.")
                continue

            sheet_base = f"{route_agent}_{run_name}"
            new_ws.name = make_unique_sheet_name(sheet_base, existing_sheet_names)
            write_loading_sheet(new_ws, route_agent, run_name, group_df, config)

        if "TempSheet" in [s.name for s in output_wb.sheets]:
            output_wb.sheets["TempSheet"].delete()

        latest_date = pd.to_datetime(df["_delivery_date"], errors="coerce").max()
        if pd.isna(latest_date):
            latest_date = datetime.today()
        month_folder = latest_date.strftime("%Y-%m")
        output_dir = Path(config.output_root) / month_folder
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / latest_date.strftime("%d-%m-%Y_LOADING_PAPER.xlsx")

        output_wb.save(str(output_path))
        log_info(f"Loading paper workbook generated: {output_path}")
    finally:
        if output_wb is not None:
            try:
                output_wb.close()
            except Exception:
                pass
        if template_wb is not None:
            try:
                template_wb.close()
            except Exception:
                pass
        app.quit()


if __name__ == "__main__":
    generate_loading_papers()
