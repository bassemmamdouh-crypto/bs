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
    line_start_row: int = 16
    line_end_row: int = 154
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
    order_data: pd.DataFrame,
    order_id: str,
    line_start_row: int,
    line_end_row: int,
    mapping_column: Optional[str],
) -> None:
    qty_lookup: Dict[str, float] = {}
    price_lookup: Dict[str, float] = {}

    if mapping_column is None:
        log_warning(
            "Mapping column not found ('new_inovice_mapping'/'new_invoice_mapping'). "
            "Invoice lines will stay with quantity/price = 0."
        )
    else:
        for _, row in order_data.iterrows():
            key = safe_str(row.get(mapping_column, ""))
            if not key:
                continue
            qty_lookup[key] = qty_lookup.get(key, 0.0) + safe_float(row.get("purchased_item_count", 0))
            price_lookup[key] = price_lookup.get(key, 0.0) + safe_float(row.get("total_price", 0))

    for excel_row in range(line_start_row, line_end_row + 1):
        item_name = ws.range((excel_row, 2)).value
        if item_name in ("", None):
            continue
        key = f"{order_id}{safe_str(item_name)}"
        ws.range((excel_row, 3)).value = qty_lookup.get(key, 0.0)
        ws.range((excel_row, 6)).value = price_lookup.get(key, 0.0)


def add_free_item_and_totals(
    ws: xw.Sheet,
    order_data: pd.DataFrame,
    free_products: List[int],
    free_row: int,
) -> None:
    if "product_id" in order_data.columns:
        free_qty = safe_float(order_data[order_data["product_id"].isin(free_products)]["purchased_item_count"].sum())
    else:
        free_qty = 0.0

    if free_qty > 0:
        ws.range((free_row, 2)).value = "هدايه"
        ws.range((free_row, 3)).value = free_qty
        ws.range((free_row, 6)).value = 0

    total_qty = safe_float(order_data.get("purchased_item_count", pd.Series(dtype=float)).sum()) + free_qty
    total_price = safe_float(order_data.get("total_price", pd.Series(dtype=float)).sum())

    ws.range("C154").value = total_qty
    ws.range("F154").value = total_price
    ws.range("E155").value = total_price
    ws.range("E156").value = 0
    ws.range("E157").value = 0
    ws.range("E158").value = total_price


def remove_zero_qty_rows(ws: xw.Sheet, line_start_row: int, line_end_row: int) -> None:
    for excel_row in range(line_end_row, line_start_row - 1, -1):
        value = ws.range((excel_row, 3)).value
        if value in ("", None):
            try:
                ws.api.Rows(excel_row).Delete()
            except Exception as exc:
                log_warning(f"Could not delete empty qty row {excel_row}: {exc}")
            continue
        if safe_float(value, default=0.0) == 0.0:
            try:
                ws.api.Rows(excel_row).Delete()
            except Exception as exc:
                log_warning(f"Could not delete zero qty row {excel_row}: {exc}")


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

        invoice_template_ws.api.Copy(Before=output_wb.sheets["TempSheet"].api)
        new_ws = output_wb.sheets[0]
        new_ws.name = make_unique_sheet_name(safe_str(order_id), existing_names)

        write_header_data(new_ws, order_data, safe_str(order_id))
        fill_invoice_lines(
            new_ws,
            order_data,
            safe_str(order_id),
            config.line_start_row,
            config.line_end_row,
            mapping_column,
        )
        add_free_item_and_totals(new_ws, order_data, config.free_products, config.free_row)
        remove_zero_qty_rows(new_ws, config.line_start_row, config.line_end_row)

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
