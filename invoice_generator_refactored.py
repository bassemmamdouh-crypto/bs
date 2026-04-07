import os
import re
from dataclasses import dataclass, field
from datetime import datetime, time
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import xlwings as xw


# =============================
# SETTINGS (edit here only)
# =============================
ORDERS_FILE = r"E:\Marbah Products\Marbah Invoices script\orders_data.xlsx"
TEMPLATE_FILE = r"E:\Marbah Products\Marbah Invoices script\invoice_template.xlsx"
DORA_TEMPLATE_FILE = r"E:\Marbah Products\Marbah Invoices script\eldora_invoice_template.xlsx"
PEPSI_TEMPLATE_FILE = r"E:\Marbah Products\Marbah Invoices script\Pepsi_invoice_template.xlsx"
OUTPUT_ROOT = r"E:\Marbah Products\Marbah Invoices script"


@dataclass
class RouteConfig:
    route_name: str
    template_file: str
    output_suffix: str


@dataclass
class ScriptConfig:
    orders_file: str = ORDERS_FILE
    output_root: str = OUTPUT_ROOT
    free_products: List[int] = field(default_factory=lambda: [180, 181, 182])
    invoice_sheet_candidates: List[str] = field(
        default_factory=lambda: ["invoice templet", "invoice template", "invoice"]
    )
    daily_summary_sheet_name: str = "Daily Summary"
    mapping_column_candidates: List[str] = field(
        default_factory=lambda: ["new_inovice_mapping", "new_invoice_mapping"]
    )
    sku_name_column_candidates: List[str] = field(
        default_factory=lambda: [
            "sku_name",
            "item_name",
            "product_name",
            "product_name_ar",
            "new_inovice_mapping",
            "new_invoice_mapping",
        ]
    )
    sku_code_column_candidates: List[str] = field(
        default_factory=lambda: ["sku", "sku_code", "product_id", "item_id"]
    )
    line_start_row: int = 16
    line_end_row: int = 153
    totals_row: int = 154
    subtotal_row: int = 155
    discount_row: int = 156
    adjustment_row: int = 157
    net_total_row: int = 158
    free_row: int = 159
    routes: List[RouteConfig] = field(
        default_factory=lambda: [
            RouteConfig("EL MANSOUR", TEMPLATE_FILE, "El-Mansour"),
            RouteConfig("AMRIA", DORA_TEMPLATE_FILE, "Amria"),
            RouteConfig("EL GAMAA", PEPSI_TEMPLATE_FILE, "El-Gamaa"),
            RouteConfig("OTAYFIA & ALAWY", TEMPLATE_FILE, "Otayfia-Alawy"),
        ]
    )


REQUIRED_COLUMNS: Dict[str, object] = {
    "route": "",
    "order_id": "",
    "estimated_delivery_date": pd.NaT,
    "purchased_item_count": 0,
    "product_id": 0,
    "total_price": 0,
    "order_date": pd.NaT,
    "order_time": pd.NaT,
    "retailer_name": "",
    "market_name": "",
    "mobile": "",
    "polygon_name": "",
    "sales_agent": "",
    "delivery_status": "",
    "route_agent": "",
    "run": "",
    "new_inovice_mapping": "",
    "new_invoice_mapping": "",
    "sku": "",
    "sku_code": "",
    "item_id": "",
    "sku_name": "",
    "item_name": "",
    "product_name": "",
    "product_name_ar": "",
    "second_run_gift": 0,
    "doritos_gift": 0,
}


def log_warning(message: str) -> None:
    print(f"WARNING: {message}")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def ensure_columns(df: pd.DataFrame, defaults: Dict[str, object]) -> pd.DataFrame:
    for col, default_value in defaults.items():
        if col not in df.columns:
            df[col] = default_value
            log_warning(f"Missing column '{col}' in orders file. Default value was added.")
    return df


def parse_datetime_columns(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def clean_order_id(value: object) -> str:
    x = str(value or "").strip()
    x = x.replace("\u200f", "").replace("\u200e", "").replace("\u200b", "")
    x = re.sub(r"\s+", "", x)
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    for i, digit in enumerate(arabic_digits):
        x = x.replace(digit, str(i))
    return x


def safe_str(value: object, default: str = "") -> str:
    if pd.isna(value):
        return default
    return str(value).strip()


def safe_float(value: object, default: float = 0.0) -> float:
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def safe_date_str(value: object, default: str = "") -> str:
    if pd.isna(value):
        return default
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return default
    return parsed.strftime("%Y-%m-%d")


def safe_time_str(value: object, default: str = "") -> str:
    if pd.isna(value):
        return default
    if isinstance(value, datetime):
        return value.strftime("%H:%M")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return default
    return parsed.strftime("%H:%M")


def first_value(df: pd.DataFrame, column: str, default: object = "") -> object:
    if df.empty or column not in df.columns:
        return default
    value = df.iloc[0][column]
    if pd.isna(value):
        return default
    return value


def first_existing_value(row: pd.Series, candidates: List[str], default: str = "") -> str:
    for col in candidates:
        if col in row.index:
            value = safe_str(row.get(col, ""))
            if value:
                return value
    return default


def safe_write(ws: xw.Sheet, cell: str, value: object) -> None:
    try:
        ws[cell].value = value
    except Exception as exc:
        log_warning(f"Could not write to cell {cell} in sheet '{ws.name}': {exc}")


def get_sheet_if_exists(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def get_invoice_template_sheet(workbook: xw.Book, candidates: List[str]) -> Optional[xw.Sheet]:
    for name in candidates:
        sheet = get_sheet_if_exists(workbook, name)
        if sheet is not None:
            return sheet
    if workbook.sheets.count > 0:
        log_warning(
            f"Invoice template sheet not found. Falling back to first sheet: '{workbook.sheets[0].name}'."
        )
        return workbook.sheets[0]
    return None


def get_mapping_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def infer_item_label(
    row: pd.Series,
    order_id: str,
    mapping_column: Optional[str],
    config: ScriptConfig,
) -> str:
    item_name = first_existing_value(row, config.sku_name_column_candidates, "")

    if mapping_column and (not item_name):
        mapping_value = safe_str(row.get(mapping_column, ""))
        if mapping_value:
            item_name = (
                safe_str(mapping_value[len(order_id):], "")
                if mapping_value.startswith(order_id)
                else mapping_value
            )

    sku_code = first_existing_value(row, config.sku_code_column_candidates, "")

    if item_name and sku_code and sku_code not in item_name:
        return f"{sku_code} - {item_name}"
    if item_name:
        return item_name
    if sku_code:
        return f"SKU {sku_code}"
    return "Unknown SKU"


def build_invoice_items(
    order_data: pd.DataFrame,
    order_id: str,
    mapping_column: Optional[str],
    config: ScriptConfig,
) -> List[Dict[str, float]]:
    aggregated: Dict[str, Dict[str, float]] = {}

    for _, row in order_data.iterrows():
        item_label = infer_item_label(row, order_id, mapping_column, config)
        if item_label not in aggregated:
            aggregated[item_label] = {"item_name": item_label, "qty": 0.0, "price": 0.0}

        aggregated[item_label]["qty"] += safe_float(row.get("purchased_item_count", 0))
        aggregated[item_label]["price"] += safe_float(row.get("total_price", 0))

    items = list(aggregated.values())
    items.sort(key=lambda x: x["item_name"])
    return items


def ensure_line_capacity(ws: xw.Sheet, required_lines: int, config: ScriptConfig) -> int:
    base_capacity = config.line_end_row - config.line_start_row + 1
    if required_lines <= base_capacity:
        return 0

    extra_lines = required_lines - base_capacity
    inserted_lines = 0
    for _ in range(extra_lines):
        try:
            # Insert before totals/footer to preserve template bottom section.
            ws.api.Rows(config.totals_row).Insert()
            inserted_lines += 1
        except Exception as exc:
            log_warning(f"Could not insert extra SKU rows before totals: {exc}")
            break
    return inserted_lines


def clear_line_area(ws: xw.Sheet, start_row: int, end_row: int) -> None:
    if end_row < start_row:
        return
    # Clear only SKU name/qty/price columns to preserve other template formulas.
    ws.range((start_row, 2), (end_row, 2)).value = None
    ws.range((start_row, 3), (end_row, 3)).value = None
    ws.range((start_row, 6), (end_row, 6)).value = None


def make_unique_sheet_name(order_id: str, existing_names: set) -> str:
    base = re.sub(r'[\[\]\*\/\\\?\:]', "", order_id)[:25]
    candidate = f"INV_{base}" if base else "INV"
    counter = 1
    while candidate in existing_names:
        candidate = f"INV_{base}_{counter}"[:31]
        counter += 1
    existing_names.add(candidate)
    return candidate[:31]


def update_daily_summary(summary_ws: xw.Sheet, route_df: pd.DataFrame, free_products: List[int]) -> datetime:
    latest_date = route_df["estimated_delivery_date"].dropna().max()
    if pd.isna(latest_date):
        latest_date = datetime.today()

    total_orders = route_df["order_id"].nunique() if "order_id" in route_df.columns else 0
    total_items = safe_float(route_df["purchased_item_count"].sum())

    if "product_id" in route_df.columns:
        total_smily_gifts = safe_float(
            route_df[route_df["product_id"].isin(free_products)]["purchased_item_count"].sum()
        )
    else:
        total_smily_gifts = 0.0

    total_second_run_gifts = safe_float(route_df.get("second_run_gift", pd.Series(dtype=float)).sum())
    total_doritos_gifts = safe_float(route_df.get("doritos_gift", pd.Series(dtype=float)).sum())

    safe_write(summary_ws, "C2", latest_date.strftime("%Y-%m-%d"))
    safe_write(summary_ws, "C6", total_orders)
    safe_write(summary_ws, "C7", total_items)
    safe_write(summary_ws, "C8", total_second_run_gifts)
    safe_write(summary_ws, "C9", total_doritos_gifts)
    safe_write(summary_ws, "C10", total_smily_gifts)
    safe_write(summary_ws, "C11", total_items + total_second_run_gifts + total_doritos_gifts)

    return latest_date


def write_header_data(ws: xw.Sheet, order_data: pd.DataFrame, order_id: str) -> None:
    retailer = safe_str(first_value(order_data, "retailer_name", ""))
    market = safe_str(first_value(order_data, "market_name", ""))

    safe_write(ws, "B7", f"{retailer} \\ {market}".strip())
    safe_write(ws, "B8", first_value(order_data, "mobile", ""))
    safe_write(ws, "B9", first_value(order_data, "polygon_name", ""))
    safe_write(ws, "B10", safe_date_str(first_value(order_data, "estimated_delivery_date", pd.NaT), ""))
    safe_write(ws, "B11", first_value(order_data, "sales_agent", ""))
    safe_write(ws, "B12", market)

    safe_write(ws, "E7", order_id)
    safe_write(ws, "D8", safe_date_str(first_value(order_data, "order_date", pd.NaT), ""))
    safe_write(ws, "E8", safe_time_str(first_value(order_data, "order_time", pd.NaT), ""))

    safe_write(ws, "E10", first_value(order_data, "delivery_status", ""))
    safe_write(ws, "E11", first_value(order_data, "route_agent", ""))
    safe_write(ws, "E12", first_value(order_data, "run", ""))


def fill_invoice_lines(
    ws: xw.Sheet,
    line_items: List[Dict[str, float]],
    config: ScriptConfig,
    row_shift: int,
) -> None:
    dynamic_line_end = config.line_end_row + row_shift
    clear_line_area(ws, config.line_start_row, dynamic_line_end)

    for idx, item in enumerate(line_items):
        excel_row = config.line_start_row + idx
        ws.range((excel_row, 2)).value = item["item_name"]
        ws.range((excel_row, 3)).value = item["qty"]
        ws.range((excel_row, 6)).value = item["price"]


def add_free_item_and_totals(
    ws: xw.Sheet,
    order_data: pd.DataFrame,
    free_products: List[int],
    config: ScriptConfig,
    row_shift: int,
) -> None:
    if "product_id" in order_data.columns:
        free_qty = safe_float(order_data[order_data["product_id"].isin(free_products)]["purchased_item_count"].sum())
    else:
        free_qty = 0.0

    shifted_free_row = config.free_row + row_shift
    shifted_totals_row = config.totals_row + row_shift
    shifted_subtotal_row = config.subtotal_row + row_shift
    shifted_discount_row = config.discount_row + row_shift
    shifted_adjustment_row = config.adjustment_row + row_shift
    shifted_net_total_row = config.net_total_row + row_shift

    if free_qty > 0:
        ws.range((shifted_free_row, 2)).value = "هدايه"
        ws.range((shifted_free_row, 3)).value = free_qty
        ws.range((shifted_free_row, 6)).value = 0

    total_qty = safe_float(order_data.get("purchased_item_count", pd.Series(dtype=float)).sum()) + free_qty
    total_price = safe_float(order_data.get("total_price", pd.Series(dtype=float)).sum())

    ws.range((shifted_totals_row, 3)).value = total_qty
    ws.range((shifted_totals_row, 6)).value = total_price
    ws.range((shifted_subtotal_row, 5)).value = total_price
    ws.range((shifted_discount_row, 5)).value = 0
    ws.range((shifted_adjustment_row, 5)).value = 0
    ws.range((shifted_net_total_row, 5)).value = total_price


def process_route(route_cfg: RouteConfig, orders_df: pd.DataFrame, app: xw.App, config: ScriptConfig) -> Optional[str]:
    route_df = orders_df[
        orders_df["route"].astype(str).str.strip().str.upper() == route_cfg.route_name.strip().upper()
    ].copy()

    if route_df.empty:
        print(f"INFO: No orders found for route '{route_cfg.route_name}'.")
        return None

    route_df["order_id"] = route_df["order_id"].apply(clean_order_id)
    route_df["estimated_delivery_date"] = pd.to_datetime(route_df["estimated_delivery_date"], errors="coerce")

    if not os.path.exists(route_cfg.template_file):
        log_warning(
            f"Template file not found for route '{route_cfg.route_name}': {route_cfg.template_file}. "
            "Skipping this route."
        )
        return None

    template_wb = app.books.open(route_cfg.template_file)
    output_wb = app.books.add()
    output_wb.sheets[0].name = "TempSheet"
    existing_names = set()

    invoice_template_ws = get_invoice_template_sheet(template_wb, config.invoice_sheet_candidates)
    summary_ws = get_sheet_if_exists(template_wb, config.daily_summary_sheet_name)

    if invoice_template_ws is None:
        log_warning(f"No usable invoice sheet found in '{route_cfg.template_file}'. Skipping route.")
        output_wb.close()
        template_wb.close()
        return None

    if summary_ws is not None:
        latest_date = update_daily_summary(summary_ws, route_df, config.free_products)
    else:
        log_warning(
            f"Sheet '{config.daily_summary_sheet_name}' not found in template. "
            "Daily Summary will be skipped."
        )
        latest_date = route_df["estimated_delivery_date"].dropna().max()
        if pd.isna(latest_date):
            latest_date = datetime.today()

    mapping_column = get_mapping_column(route_df, config.mapping_column_candidates)

    for order_id in sorted(route_df["order_id"].dropna().unique()):
        order_data = route_df[route_df["order_id"] == order_id]
        normalized_order_id = safe_str(order_id)

        invoice_template_ws.api.Copy(Before=output_wb.sheets["TempSheet"].api)
        new_ws = output_wb.sheets[0]
        new_ws.name = make_unique_sheet_name(normalized_order_id, existing_names)

        write_header_data(new_ws, order_data, normalized_order_id)
        line_items = build_invoice_items(order_data, normalized_order_id, mapping_column, config)
        row_shift = ensure_line_capacity(new_ws, len(line_items), config)
        fill_invoice_lines(
            new_ws,
            line_items,
            config,
            row_shift,
        )
        add_free_item_and_totals(new_ws, order_data, config.free_products, config, row_shift)

    if "TempSheet" in [sheet.name for sheet in output_wb.sheets]:
        output_wb.sheets["TempSheet"].delete()

    if summary_ws is not None:
        summary_ws.api.Copy(Before=output_wb.sheets[0].api)
        output_wb.sheets[0].name = "Daily Summary"

    month_folder = latest_date.strftime("%Y-%m")
    day_file = latest_date.strftime(f"%d-%m-%Y_{route_cfg.output_suffix}")
    output_dir = Path(config.output_root) / month_folder
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / f"{day_file}.xlsx"
    output_wb.save(str(output_file))

    output_wb.close()
    template_wb.close()
    return str(output_file)


def main() -> None:
    config = ScriptConfig()

    if not os.path.exists(config.orders_file):
        raise FileNotFoundError(f"Orders file not found: {config.orders_file}")

    orders_df = pd.read_excel(config.orders_file)
    orders_df = normalize_columns(orders_df)
    orders_df = ensure_columns(orders_df, REQUIRED_COLUMNS)
    orders_df = parse_datetime_columns(
        orders_df, ["estimated_delivery_date", "order_date", "order_time"]
    )

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    saved_files: List[str] = []
    try:
        for route_cfg in config.routes:
            saved_path = process_route(route_cfg, orders_df, app, config)
            if saved_path:
                saved_files.append(saved_path)
    finally:
        app.quit()

    if saved_files:
        print("\nDONE: Invoices generated successfully:")
        for file_path in saved_files:
            print(f" - {file_path}")
    else:
        print("DONE: No files were generated (no matching routes or missing templates).")


if __name__ == "__main__":
    main()
