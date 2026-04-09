import math
import os
import re
from dataclasses import dataclass, field
from datetime import datetime, time
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import xlwings as xw


# =============================
# SETTINGS (EDIT THIS SECTION)
# =============================
ORDERS_FILE = r"E:\Marbah Products\Marbah Invoices script\orders_data.xlsx"
DEFAULT_TEMPLATE_FILE = r"E:\Marbah Products\Marbah Invoices script\invoice_templet.xlsx"
OUTPUT_ROOT = r"E:\Marbah Products\Marbah Invoices script"

# One invoice template for all areas (mapped explicitly).
AREA_TEMPLATE_MAP = {
    "EL MANSOUR": r"E:\Marbah Products\Marbah Invoices script\invoice_template.xlsx",
    "AMRIA": r"E:\Marbah Products\Marbah Invoices script\invoice_template.xlsx",
    "EL GAMAA": r"E:\Marbah Products\Marbah Invoices script\invoice_template.xlsx",
    "OTAYFIA & ALAWY": r"E:\Marbah Products\Marbah Invoices script\invoice_template.xlsx",
}
TEMPLATE_FILENAME_CANDIDATES = [
    "invoice_templet.xlsx",
    "invoice_template.xlsx",
    "invoice templet.xlsx",
    "invoice template.xlsx",
]


@dataclass
class OfferRule:
    name: str
    rule_type: str
    active: bool = True
    buy_sku: str = ""
    buy_qty: float = 0.0
    gift_sku: str = ""
    gift_name: str = ""
    gift_qty: float = 0.0
    min_subtotal: float = 0.0
    discount_amount: float = 0.0
    discount_percent: float = 0.0


@dataclass
class ScriptConfig:
    orders_file: str = ORDERS_FILE
    default_template_file: str = DEFAULT_TEMPLATE_FILE
    output_root: str = OUTPUT_ROOT
    area_template_map: Dict[str, str] = field(default_factory=lambda: AREA_TEMPLATE_MAP.copy())

    # Sheet names
    invoice_sheet_candidates: List[str] = field(
        default_factory=lambda: ["invoice templet", "invoice template", "invoice"]
    )

    # Grouping
    area_column: str = "route"
    order_id_column: str = "order_id"
    delivery_date_column: str = "estimated_delivery_date"
    split_by_section: bool = True
    section_column_candidates: List[str] = field(
        default_factory=lambda: ["section", "section_name", "business_unit", "category", "brand_section", "brand"]
    )
    target_sections: List[str] = field(default_factory=lambda: ["PEPSI", "LAYS"])
    other_section_name: str = "OTHER"
    generation_mode: str = "per_order_section_workbooks"
    # summary_by_section: one workbook per area with two summary sheets (PEPSI/LAYS)
    # per_order_section_workbooks: up to two workbooks per area (PEPSI and LAYS),
    # each workbook contains one invoice sheet per order.
    create_empty_summary_sections: bool = True

    # Header fields (from first row of each order)
    retailer_column_candidates: List[str] = field(default_factory=lambda: ["retailer_name", "customer_name", "customer"])
    market_column_candidates: List[str] = field(default_factory=lambda: ["market_name", "market", "store_name"])
    mobile_column_candidates: List[str] = field(default_factory=lambda: ["mobile", "phone", "mobile_number"])
    area_name_column_candidates: List[str] = field(default_factory=lambda: ["polygon_name", "zone", "area"])
    order_date_column_candidates: List[str] = field(default_factory=lambda: ["order_date", "created_date"])
    order_time_column_candidates: List[str] = field(default_factory=lambda: ["order_time", "created_time"])
    sales_agent_column_candidates: List[str] = field(default_factory=lambda: ["sales_agent", "salesman", "sales_rep"])
    delivery_status_column_candidates: List[str] = field(default_factory=lambda: ["delivery_status", "status"])
    route_agent_column_candidates: List[str] = field(default_factory=lambda: ["route_agent", "driver"])
    run_column_candidates: List[str] = field(default_factory=lambda: ["run", "trip", "run_name"])

    # Invoice line candidates
    sku_code_column_candidates: List[str] = field(
        default_factory=lambda: ["sku", "sku_code", "product_id", "item_id", "sku_id", "product_code"]
    )
    item_name_column_candidates: List[str] = field(
        default_factory=lambda: ["sku_name", "item_name", "product_name", "product_name_ar", "item", "description"]
    )
    brand_column_candidates: List[str] = field(
        default_factory=lambda: ["brand", "brand_name", "product_brand", "manufacturer"]
    )
    size_column_candidates: List[str] = field(
        default_factory=lambda: ["size", "pack_size", "sku_size", "item_size", "variant_size"]
    )
    qty_column_candidates: List[str] = field(default_factory=lambda: ["purchased_item_count", "qty", "quantity", "item_qty"])
    unit_column_candidates: List[str] = field(default_factory=lambda: ["unit", "uom", "unit_name"])
    amount_column_candidates: List[str] = field(default_factory=lambda: ["total_price", "line_total", "item_total", "amount", "price"])

    # Offer columns from dataframe (optional, safe if missing)
    offer_item_name_candidates: List[str] = field(
        default_factory=lambda: ["offer_item_name", "offer_name", "promo_name"]
    )
    offer_item_qty_candidates: List[str] = field(default_factory=lambda: ["offer_item_qty", "offer_qty", "promo_qty"])
    offer_item_price_candidates: List[str] = field(
        default_factory=lambda: ["offer_item_price", "offer_item_amount", "promo_amount"]
    )
    gift_qty_columns: List[str] = field(
        default_factory=lambda: [
            "second_run_gift",
            "doritos_gift",
            "pepsi_gift",
            "pepsi_gift_qty",
            "pepsi_offer_gift",
            "pepsi_offer_qty",
            "gift_qty",
            "free_qty",
        ]
    )
    discount_amount_columns: List[str] = field(
        default_factory=lambda: ["discount_amount", "offer_discount_amount", "promo_discount", "order_discount"]
    )
    discount_percent_columns: List[str] = field(
        default_factory=lambda: ["discount_percent", "offer_discount_percent", "promo_discount_percent"]
    )

    # Invoice table + footer layout
    line_start_row: int = 16
    line_end_row: int = 153
    totals_row: int = 154
    subtotal_row: int = 155
    discount_row: int = 156
    adjustment_row: int = 157
    net_total_row: int = 158

    # Enable/disable offer lines (gifts and explicit offer items)
    include_offer_lines: bool = True

    # Rule-based offers (edit based on business rules)
    offer_rules: List[OfferRule] = field(
        default_factory=lambda: [
            # Example rule: buy 10 of SKU 180 -> get 1 gift of same SKU
            OfferRule(
                name="Buy10Get1-SKU180",
                rule_type="buy_qty_get_free_same_sku",
                active=False,
                buy_sku="180",
                buy_qty=10,
                gift_qty=1,
                gift_name="Offer Gift SKU 180",
            ),
            # Example rule: subtotal >= 1000 -> 5% discount
            OfferRule(
                name="Subtotal-5pct-over-1000",
                rule_type="order_subtotal_discount_pct",
                active=False,
                min_subtotal=1000,
                discount_percent=5,
            ),
        ]
    )


# Optional input aliases if source columns have different names
COLUMN_ALIASES: Dict[str, List[str]] = {
    "route": ["route_name", "delivery_route"],
    "order_id": ["orderid", "order_number", "order no", "order_no"],
    "estimated_delivery_date": ["delivery_date", "estimated_date"],
    "purchased_item_count": ["qty", "quantity", "item_qty"],
    "total_price": ["line_total", "item_total", "amount", "price"],
    "retailer_name": ["retailer", "customer_name", "customer"],
    "market_name": ["market", "store_name", "store"],
    "mobile": ["phone", "mobile_number", "customer_phone"],
    "polygon_name": ["area", "zone", "region"],
    "sales_agent": ["salesman", "sales_rep"],
    "route_agent": ["driver"],
    "section": ["section_name", "business_unit", "category", "division"],
    "sku_name": ["item_name", "product_name", "item", "description"],
    "sku_code": ["sku", "product_code", "sku_id"],
}


REQUIRED_COLUMNS_DEFAULTS: Dict[str, object] = {
    "route": "",
    "order_id": "",
    "estimated_delivery_date": pd.NaT,
    "order_date": pd.NaT,
    "order_time": pd.NaT,
    "purchased_item_count": 0,
    "total_price": 0,
    "retailer_name": "",
    "market_name": "",
    "mobile": "",
    "polygon_name": "",
    "sales_agent": "",
    "delivery_status": "",
    "route_agent": "",
    "run": "",
    "sku_name": "",
}

OPTIONAL_COLUMNS_DEFAULTS: Dict[str, object] = {
    "sku_code": "",
    "section": "",
}


def log_info(msg: str) -> None:
    print(f"INFO: {msg}")


def log_warning(msg: str) -> None:
    print(f"WARNING: {msg}")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def apply_column_aliases(df: pd.DataFrame, aliases: Dict[str, List[str]]) -> pd.DataFrame:
    rename_map: Dict[str, str] = {}
    existing = set(df.columns)
    for target, candidates in aliases.items():
        if target in existing:
            continue
        for candidate in candidates:
            candidate_norm = candidate.strip().lower()
            if candidate_norm in existing:
                rename_map[candidate_norm] = target
                break
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def ensure_columns(df: pd.DataFrame, defaults: Dict[str, object], warn_missing: bool = True) -> pd.DataFrame:
    for col, default in defaults.items():
        if col not in df.columns:
            df[col] = default
            if warn_missing:
                log_warning(f"Column '{col}' missing in orders file. Default value applied.")
    return df


def safe_str(value: object, default: str = "") -> str:
    if pd.isna(value):
        return default
    s = str(value).strip()
    if s.lower() in {"nan", "none", "null", "nat"}:
        return default
    return s


def normalize_column_key(name: object) -> str:
    return re.sub(r"[^a-z0-9]", "", safe_str(name, "").lower())


def resolve_candidate_columns(df: pd.DataFrame, candidates: List[str]) -> List[str]:
    existing_cols = [str(c) for c in df.columns]
    normalized_to_col: Dict[str, str] = {}
    for col in existing_cols:
        key = normalize_column_key(col)
        if key and key not in normalized_to_col:
            normalized_to_col[key] = col

    resolved: List[str] = []
    seen = set()

    # Exact normalized matching.
    for cand in candidates:
        key = normalize_column_key(cand)
        matched = normalized_to_col.get(key)
        if matched and matched not in seen:
            seen.add(matched)
            resolved.append(matched)

    # Fuzzy fallback by containment.
    for cand in candidates:
        cand_key = normalize_column_key(cand)
        if not cand_key:
            continue
        for col in existing_cols:
            col_key = normalize_column_key(col)
            if cand_key in col_key and col not in seen:
                seen.add(col)
                resolved.append(col)

    return resolved


def normalize_section_name(value: object, config: ScriptConfig) -> str:
    text = safe_str(value, "").strip().upper()
    if not text:
        return config.other_section_name
    if "PEPSICO" in text:
        return "PEPSI"
    if "PEPSI" in text:
        return "PEPSI"
    if "LAYS" in text or "LAY'S" in text or "LAYS" in text:
        return "LAYS"
    if "CHIPSY" in text:
        return "LAYS"
    return config.other_section_name


def get_target_sections(config: ScriptConfig) -> List[str]:
    target_sections: List[str] = []
    seen_sections = set()
    for sec in config.target_sections:
        normalized = normalize_section_name(sec, config)
        if normalized == config.other_section_name:
            continue
        if normalized not in seen_sections:
            seen_sections.add(normalized)
            target_sections.append(normalized)
    if not target_sections:
        return ["PEPSI", "LAYS"]
    return target_sections


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


def first_row_value(order_df: pd.DataFrame, candidates: List[str], default: object = "") -> object:
    if order_df.empty:
        return default
    row = order_df.iloc[0]
    normalized_index = {normalize_column_key(c): c for c in order_df.columns}
    for col in candidates:
        if col in order_df.columns:
            val = row.get(col, default)
            if safe_str(val, "") != "":
                return val
        normalized_col = normalized_index.get(normalize_column_key(col))
        if normalized_col:
            val = row.get(normalized_col, default)
            if safe_str(val, "") != "":
                return val
    return default


def first_row_float(order_df: pd.DataFrame, candidates: List[str], default: float = 0.0) -> float:
    return safe_float(first_row_value(order_df, candidates, default), default)


def aggregate_order_level_numeric(order_df: pd.DataFrame, column_name: str) -> float:
    if column_name not in order_df.columns:
        return 0.0
    series = pd.to_numeric(order_df[column_name], errors="coerce").fillna(0)
    non_zero = series[series != 0]
    if non_zero.empty:
        return 0.0
    # If value is repeated in each row as order-level attribute, keep one.
    unique = non_zero.unique()
    if len(unique) == 1 and len(non_zero) > 1:
        return float(unique[0])
    return float(non_zero.sum())


def resolve_template_path(template_path: str) -> str:
    """
    Return first existing template path among common filename variants.
    """
    if os.path.exists(template_path):
        return template_path

    base_path = Path(template_path)
    parent = base_path.parent
    for candidate_name in TEMPLATE_FILENAME_CANDIDATES:
        candidate_path = parent / candidate_name
        if candidate_path.exists():
            return str(candidate_path)
    return template_path


def choose_template_file(area_value: str, config: ScriptConfig) -> str:
    key = safe_str(area_value, "").upper()
    template = config.area_template_map.get(key, config.default_template_file)
    return resolve_template_path(template)


def find_existing_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for col in candidates:
        if col in df.columns:
            return col
    normalized_index = {normalize_column_key(c): c for c in df.columns}
    for col in candidates:
        normalized_col = normalized_index.get(normalize_column_key(col))
        if normalized_col:
            return normalized_col
    return None


def get_sheet_if_exists(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def get_template_invoice_sheet(workbook: xw.Book, candidates: List[str]) -> Optional[xw.Sheet]:
    for name in candidates:
        sheet = get_sheet_if_exists(workbook, name)
        if sheet is not None:
            return sheet
    if workbook.sheets.count > 0:
        log_warning(f"Invoice sheet not found by name. Using first sheet '{workbook.sheets[0].name}'.")
        return workbook.sheets[0]
    return None


def copy_invoice_sheet(template_ws: xw.Sheet, output_wb: xw.Book, temp_sheet_name: str) -> Optional[xw.Sheet]:
    before = [s.name for s in output_wb.sheets]
    template_ws.api.Copy(Before=output_wb.sheets[temp_sheet_name].api)
    for sheet in output_wb.sheets:
        if sheet.name not in before:
            return sheet
    names = [s.name for s in output_wb.sheets]
    if temp_sheet_name in names:
        idx = names.index(temp_sheet_name)
        if idx > 0:
            return output_wb.sheets[idx - 1]
    return None


def make_unique_sheet_name(base_name: str, existing_names: set) -> str:
    clean = re.sub(r"[\[\]\*\/\\\?\:]", "", safe_str(base_name, "INV"))[:25]
    candidate = f"INV_{clean}" if clean else "INV"
    i = 1
    while candidate in existing_names:
        candidate = f"INV_{clean}_{i}"[:31]
        i += 1
    existing_names.add(candidate)
    return candidate[:31]


def adjust_line_capacity(ws: xw.Sheet, required_count: int, config: ScriptConfig) -> int:
    """
    Adjust line section size to exactly match required_count:
    - insert rows before footer if items exceed template capacity
    - delete empty template rows before footer if items are fewer

    Returns signed row shift applied to footer rows.
    """
    base_capacity = config.line_end_row - config.line_start_row + 1
    target_count = max(0, required_count)
    requested_shift = target_count - base_capacity

    if requested_shift > 0:
        inserted = 0
        for _ in range(requested_shift):
            try:
                ws.api.Rows(config.totals_row).Insert()
                inserted += 1
            except Exception as exc:
                log_warning(f"Unable to insert invoice rows before footer: {exc}")
                break
        return inserted

    if requested_shift < 0:
        delete_count = abs(requested_shift)
        delete_start_row = config.line_start_row + target_count
        deleted = 0
        for _ in range(delete_count):
            try:
                ws.api.Rows(delete_start_row).Delete()
                deleted += 1
            except Exception as exc:
                log_warning(f"Unable to delete empty invoice row before footer: {exc}")
                break
        return -deleted

    return 0


def clear_invoice_lines(ws: xw.Sheet, start_row: int, end_row: int) -> None:
    if end_row < start_row:
        return
    ws.range((start_row, 1), (end_row, 6)).value = None


def write_header(ws: xw.Sheet, order_df: pd.DataFrame, order_id: str, config: ScriptConfig) -> None:
    retailer = safe_str(first_row_value(order_df, config.retailer_column_candidates, ""))
    market = safe_str(first_row_value(order_df, config.market_column_candidates, ""))
    mobile = safe_str(first_row_value(order_df, config.mobile_column_candidates, ""))
    zone = safe_str(first_row_value(order_df, config.area_name_column_candidates, ""))
    delivery_date = safe_date_str(first_row_value(order_df, [config.delivery_date_column], pd.NaT), "")
    order_date = safe_date_str(first_row_value(order_df, config.order_date_column_candidates, pd.NaT), "")
    order_time = safe_time_str(first_row_value(order_df, config.order_time_column_candidates, pd.NaT), "")
    sales_agent = safe_str(first_row_value(order_df, config.sales_agent_column_candidates, ""))
    delivery_status = safe_str(first_row_value(order_df, config.delivery_status_column_candidates, ""))
    route_agent = safe_str(first_row_value(order_df, config.route_agent_column_candidates, ""))
    run_name = safe_str(first_row_value(order_df, config.run_column_candidates, ""))

    # Keep template labels, write only values.
    ws["B7"].value = f"{retailer} \\ {market}".strip(" \\")
    ws["B8"].value = mobile
    ws["B9"].value = zone
    ws["B10"].value = delivery_date
    ws["B11"].value = sales_agent
    ws["B12"].value = market

    ws["E7"].value = order_id
    ws["D8"].value = order_date
    ws["E8"].value = order_time
    ws["E10"].value = delivery_status
    ws["E11"].value = route_agent
    ws["E12"].value = run_name


def infer_line_item_name(row: pd.Series, config: ScriptConfig) -> str:
    item_name = ""
    for col in config.item_name_column_candidates:
        if col in row.index:
            item_name = safe_str(row.get(col, ""), "")
            if item_name:
                return item_name
    sku = ""
    for col in config.sku_code_column_candidates:
        if col in row.index:
            sku = safe_str(row.get(col, ""), "")
            if sku:
                return f"SKU {sku}"
    return "Unnamed Item"


def extract_size_text(item_name: str) -> str:
    # Extract simple textual size patterns (e.g., 330ml, 1.5L, 90 g).
    match = re.search(r"(\d+(?:[.,]\d+)?)\s*(ml|l|g|kg|gm|cl|oz|pcs|pc)?", safe_str(item_name, ""), re.IGNORECASE)
    if not match:
        return ""
    num = match.group(1).replace(",", ".")
    unit = safe_str(match.group(2), "").lower()
    return f"{num}{unit}"


def parse_size_sort(size_text: str) -> tuple:
    text = safe_str(size_text, "").lower()
    if not text:
        return (9, float("inf"), "")

    unit_order = {
        "ml": 1,
        "l": 2,
        "g": 3,
        "gm": 3,
        "kg": 4,
        "cl": 5,
        "oz": 6,
        "pcs": 7,
        "pc": 7,
    }
    match = re.search(r"(\d+(?:[.,]\d+)?)\s*([a-z]+)?", text)
    if not match:
        return (8, float("inf"), text)

    size_num = safe_float(match.group(1).replace(",", "."), float("inf"))
    unit = safe_str(match.group(2), "").lower()
    return (unit_order.get(unit, 8), size_num, unit)


def sku_sort_key(item: Dict[str, object]) -> tuple:
    brand = safe_str(item.get("brand", ""), "").lower()
    size = safe_str(item.get("size", ""), "")
    unit_rank, size_num, unit = parse_size_sort(size)
    item_name = safe_str(item.get("item_name", ""), "").lower()
    return (brand, unit_rank, size_num, unit, item_name)


def row_first_value(row: pd.Series, candidates: List[str], default: object = "") -> object:
    normalized_index = {normalize_column_key(c): c for c in row.index}
    for col in candidates:
        if col in row.index:
            value = row.get(col, default)
            if safe_str(value, "") != "":
                return value
        normalized_col = normalized_index.get(normalize_column_key(col))
        if normalized_col:
            value = row.get(normalized_col, default)
            if safe_str(value, "") != "":
                return value
    return default


def build_base_items(order_df: pd.DataFrame, config: ScriptConfig) -> List[Dict[str, object]]:
    aggregated: Dict[str, Dict[str, object]] = {}
    for _, row in order_df.iterrows():
        name = infer_line_item_name(row, config)
        unit = safe_str(row_first_value(row, config.unit_column_candidates, ""), "")
        brand = safe_str(row_first_value(row, config.brand_column_candidates, ""), "")
        if not brand:
            brand = safe_str(name.split(" ")[0] if safe_str(name, "") else "", "")
        size = safe_str(row_first_value(row, config.size_column_candidates, ""), "")
        if not size:
            size = extract_size_text(name)
        qty = safe_float(row_first_value(row, config.qty_column_candidates, 0), 0.0)
        amount = safe_float(row_first_value(row, config.amount_column_candidates, 0), 0.0)
        key = f"{brand}||{size}||{name}||{unit}"
        if key not in aggregated:
            aggregated[key] = {
                "item_name": name,
                "brand": brand,
                "size": size,
                "qty": 0.0,
                "unit": unit,
                "gift_qty": 0.0,
                "net_amount": 0.0,
            }
        aggregated[key]["qty"] = safe_float(aggregated[key]["qty"], 0.0) + qty
        aggregated[key]["net_amount"] = safe_float(aggregated[key]["net_amount"], 0.0) + amount

    items = [
        item
        for item in aggregated.values()
        if safe_float(item["qty"], 0.0) != 0.0 or safe_float(item["net_amount"], 0.0) != 0.0
    ]
    # Required ordering: brand first, then size.
    items.sort(key=sku_sort_key)
    return items


def quantity_by_sku(order_df: pd.DataFrame, config: ScriptConfig) -> Dict[str, float]:
    result: Dict[str, float] = {}
    for _, row in order_df.iterrows():
        sku = safe_str(row_first_value(row, config.sku_code_column_candidates, ""), "")
        if not sku:
            continue
        qty = safe_float(row_first_value(row, config.qty_column_candidates, 0), 0.0)
        result[sku] = result.get(sku, 0.0) + qty
    return result


def apply_offer_rules(order_df: pd.DataFrame, subtotal: float, config: ScriptConfig) -> Dict[str, object]:
    offer_lines: List[Dict[str, object]] = []
    discount_total = 0.0

    # A) Offer/gift columns in dataframe (safe if missing)
    for gift_col in config.gift_qty_columns:
        gift_qty = aggregate_order_level_numeric(order_df, gift_col)
        if gift_qty > 0:
            offer_lines.append(
                {
                    "item_name": f"{gift_col.replace('_', ' ').title()}",
                    "qty": 0.0,
                    "unit": "",
                    "gift_qty": gift_qty,
                    "net_amount": 0.0,
                }
            )

    for _, row in order_df.iterrows():
        offer_name = safe_str(row_first_value(row, config.offer_item_name_candidates, ""), "")
        if not offer_name:
            continue
        offer_qty = safe_float(row_first_value(row, config.offer_item_qty_candidates, 0), 0.0)
        offer_price = safe_float(row_first_value(row, config.offer_item_price_candidates, 0), 0.0)
        offer_lines.append(
            {
                "item_name": offer_name,
                "qty": 0.0,
                "unit": "",
                "gift_qty": offer_qty,
                "net_amount": offer_price,
            }
        )

    for col in config.discount_amount_columns:
        discount_total += aggregate_order_level_numeric(order_df, col)

    discount_percent = 0.0
    for col in config.discount_percent_columns:
        discount_percent += aggregate_order_level_numeric(order_df, col)
    if discount_percent:
        discount_total += subtotal * (discount_percent / 100.0)

    # B) Configured rule-engine offers
    sku_qty_map = quantity_by_sku(order_df, config)
    for rule in config.offer_rules:
        if not rule.active:
            continue

        if rule.rule_type == "buy_qty_get_free_same_sku":
            buy_sku = safe_str(rule.buy_sku, "")
            if not buy_sku or rule.buy_qty <= 0 or rule.gift_qty <= 0:
                continue
            bought = sku_qty_map.get(buy_sku, 0.0)
            multiplier = math.floor(bought / rule.buy_qty)
            if multiplier > 0:
                gift_qty = multiplier * rule.gift_qty
                gift_name = safe_str(rule.gift_name, "") or f"Offer Gift SKU {buy_sku}"
                offer_lines.append(
                    {
                        "item_name": gift_name,
                        "qty": 0.0,
                        "unit": "",
                        "gift_qty": gift_qty,
                        "net_amount": 0.0,
                    }
                )

        elif rule.rule_type == "buy_sku_get_other_sku_free":
            buy_sku = safe_str(rule.buy_sku, "")
            gift_sku = safe_str(rule.gift_sku, "")
            if not buy_sku or not gift_sku or rule.buy_qty <= 0 or rule.gift_qty <= 0:
                continue
            bought = sku_qty_map.get(buy_sku, 0.0)
            multiplier = math.floor(bought / rule.buy_qty)
            if multiplier > 0:
                gift_qty = multiplier * rule.gift_qty
                gift_name = safe_str(rule.gift_name, "") or f"Offer Gift SKU {gift_sku}"
                offer_lines.append(
                    {
                        "item_name": gift_name,
                        "qty": 0.0,
                        "unit": "",
                        "gift_qty": gift_qty,
                        "net_amount": 0.0,
                    }
                )

        elif rule.rule_type == "order_subtotal_discount_amount":
            if subtotal >= rule.min_subtotal and rule.discount_amount > 0:
                discount_total += rule.discount_amount

        elif rule.rule_type == "order_subtotal_discount_pct":
            if subtotal >= rule.min_subtotal and rule.discount_percent > 0:
                discount_total += subtotal * (rule.discount_percent / 100.0)

    # Merge duplicate offer lines by name
    merged: Dict[str, Dict[str, object]] = {}
    for line in offer_lines:
        name = safe_str(line.get("item_name", "Offer"), "Offer")
        if name not in merged:
            merged[name] = {
                "item_name": name,
                "qty": 0.0,
                "unit": "",
                "gift_qty": 0.0,
                "net_amount": 0.0,
            }
        merged[name]["gift_qty"] = safe_float(merged[name]["gift_qty"], 0.0) + safe_float(line.get("gift_qty", 0), 0.0)
        merged[name]["net_amount"] = safe_float(merged[name]["net_amount"], 0.0) + safe_float(line.get("net_amount", 0), 0.0)

    return {
        "offer_lines": sorted(list(merged.values()), key=lambda x: safe_str(x["item_name"], "")),
        "discount_total": max(0.0, safe_float(discount_total, 0.0)),
    }


def fill_invoice_lines(ws: xw.Sheet, lines: List[Dict[str, object]], config: ScriptConfig, row_shift: int) -> None:
    end_row = config.line_end_row + row_shift
    clear_invoice_lines(ws, config.line_start_row, end_row)
    for idx, line in enumerate(lines):
        row_no = config.line_start_row + idx
        ws.range((row_no, 1)).value = idx + 1
        ws.range((row_no, 2)).value = line.get("item_name", "")
        ws.range((row_no, 3)).value = safe_float(line.get("qty", 0), 0.0)
        ws.range((row_no, 4)).value = line.get("unit", "")
        ws.range((row_no, 5)).value = safe_float(line.get("gift_qty", 0), 0.0)
        ws.range((row_no, 6)).value = safe_float(line.get("net_amount", 0), 0.0)


def write_totals(
    ws: xw.Sheet,
    lines: List[Dict[str, object]],
    discount_amount: float,
    config: ScriptConfig,
    row_shift: int,
) -> None:
    totals_row = config.totals_row + row_shift
    subtotal_row = config.subtotal_row + row_shift
    discount_row = config.discount_row + row_shift
    adjustment_row = config.adjustment_row + row_shift
    net_total_row = config.net_total_row + row_shift

    total_qty = sum(safe_float(line.get("qty", 0), 0.0) + safe_float(line.get("gift_qty", 0), 0.0) for line in lines)
    subtotal = sum(safe_float(line.get("net_amount", 0), 0.0) for line in lines)
    discount = max(0.0, safe_float(discount_amount, 0.0))
    net_total = subtotal - discount

    ws.range((totals_row, 3)).value = total_qty
    ws.range((totals_row, 6)).value = subtotal
    ws.range((subtotal_row, 5)).value = subtotal
    ws.range((discount_row, 5)).value = discount
    ws.range((adjustment_row, 5)).value = 0
    ws.range((net_total_row, 5)).value = net_total


def prepare_area_sections(area_df: pd.DataFrame, config: ScriptConfig) -> pd.DataFrame:
    prepared = area_df.copy()
    section_col = find_existing_column(prepared, config.section_column_candidates)
    if config.split_by_section and section_col:
        prepared["_section_group"] = prepared[section_col].apply(lambda x: normalize_section_name(x, config))
    else:
        prepared["_section_group"] = config.other_section_name
    return prepared


def area_output_context(area_df: pd.DataFrame, area_value: str, config: ScriptConfig) -> Dict[str, object]:
    latest_delivery_date = pd.to_datetime(area_df[config.delivery_date_column], errors="coerce").max()
    if pd.isna(latest_delivery_date):
        latest_delivery_date = datetime.today()
    month_folder = latest_delivery_date.strftime("%Y-%m")
    safe_area = re.sub(r'[<>:"/\\|?*]+', "_", safe_str(area_value, "Unknown-Area"))
    output_dir = Path(config.output_root) / month_folder
    output_dir.mkdir(parents=True, exist_ok=True)
    return {
        "latest_delivery_date": latest_delivery_date,
        "output_dir": output_dir,
        "safe_area": safe_area,
    }


def render_invoice_sheet(
    ws: xw.Sheet,
    source_df: pd.DataFrame,
    header_order_id: str,
    config: ScriptConfig,
) -> None:
    write_header(ws, source_df, header_order_id, config)
    base_lines = build_base_items(source_df, config)
    subtotal = sum(safe_float(line.get("net_amount", 0), 0.0) for line in base_lines)
    offer_result = apply_offer_rules(source_df, subtotal, config)

    final_lines = list(base_lines)
    if config.include_offer_lines:
        final_lines.extend(offer_result["offer_lines"])

    row_shift = adjust_line_capacity(ws, len(final_lines), config)
    fill_invoice_lines(ws, final_lines, config, row_shift)
    write_totals(ws, final_lines, safe_float(offer_result["discount_total"], 0.0), config, row_shift)


def process_area_summary(area_value: str, area_df: pd.DataFrame, app: xw.App, config: ScriptConfig) -> Optional[str]:
    """
    Keep the summary behavior:
    one workbook per area with one summary sheet per target section.
    """
    template_file = choose_template_file(area_value, config)
    if not os.path.exists(template_file):
        log_warning(f"Template file missing for area '{area_value}': {template_file}")
        return None

    template_wb = app.books.open(template_file)
    output_wb = app.books.add()
    output_wb.sheets[0].name = "TempSheet"
    existing_sheet_names = set()

    template_ws = get_template_invoice_sheet(template_wb, config.invoice_sheet_candidates)
    if template_ws is None:
        log_warning(f"No invoice sheet found in template: {template_file}")
        output_wb.close()
        template_wb.close()
        return None

    if area_df.empty:
        log_info(f"No rows found in area '{area_value}'.")
        output_wb.close()
        template_wb.close()
        return None

    area_df = prepare_area_sections(area_df, config)
    target_sections = get_target_sections(config)

    for section_name in target_sections:
        section_df = area_df[area_df["_section_group"] == section_name].copy()
        if section_df.empty and not config.create_empty_summary_sections:
            continue

        new_ws = copy_invoice_sheet(template_ws, output_wb, "TempSheet")
        if new_ws is None:
            log_warning(f"Could not copy invoice sheet for area '{area_value}' section '{section_name}'.")
            continue

        sheet_base = f"{safe_str(area_value, 'AREA')}_{section_name}"
        new_ws.name = make_unique_sheet_name(sheet_base, existing_sheet_names)

        header_source_df = section_df if not section_df.empty else area_df
        if section_df.empty:
            write_header(new_ws, header_source_df, f"{section_name} SUMMARY", config)
            row_shift = adjust_line_capacity(new_ws, 0, config)
            fill_invoice_lines(new_ws, [], config, row_shift)
            write_totals(new_ws, [], 0.0, config, row_shift)
            log_info(f"No rows found for section '{section_name}' in area '{area_value}'. Added empty sheet.")
        else:
            render_invoice_sheet(new_ws, section_df, f"{section_name} SUMMARY", config)

    if "TempSheet" in [sheet.name for sheet in output_wb.sheets]:
        output_wb.sheets["TempSheet"].delete()

    area_ctx = area_output_context(area_df, area_value, config)
    latest_delivery_date = area_ctx["latest_delivery_date"]
    output_dir = area_ctx["output_dir"]
    safe_area = area_ctx["safe_area"]
    day_file = latest_delivery_date.strftime(f"%d-%m-%Y_{safe_area}_SUMMARY")
    output_path = output_dir / f"{day_file}.xlsx"
    output_wb.save(str(output_path))

    output_wb.close()
    template_wb.close()
    return str(output_path)


def process_area_order_workbooks(
    area_value: str,
    area_df: pd.DataFrame,
    app: xw.App,
    config: ScriptConfig,
) -> List[str]:
    """
    Create up to two workbooks per area (PEPSI + LAYS, if each section exists).
    Each workbook contains one invoice sheet per order for its section.
    """
    template_file = choose_template_file(area_value, config)
    if not os.path.exists(template_file):
        log_warning(f"Template file missing for area '{area_value}': {template_file}")
        return []

    if area_df.empty:
        log_info(f"No rows found in area '{area_value}'.")
        return []

    area_df = prepare_area_sections(area_df, config)
    target_sections = get_target_sections(config)
    area_ctx = area_output_context(area_df, area_value, config)
    latest_delivery_date = area_ctx["latest_delivery_date"]
    output_dir = area_ctx["output_dir"]
    safe_area = area_ctx["safe_area"]
    saved_paths: List[str] = []

    template_wb = app.books.open(template_file)
    try:
        template_ws = get_template_invoice_sheet(template_wb, config.invoice_sheet_candidates)
        if template_ws is None:
            log_warning(f"No invoice sheet found in template: {template_file}")
            return []

        for section_name in target_sections:
            section_df = area_df[area_df["_section_group"] == section_name].copy()
            if section_df.empty:
                log_info(f"Skipping workbook for section '{section_name}' in area '{area_value}' (no rows).")
                continue

            output_wb = app.books.add()
            output_wb.sheets[0].name = "TempSheet"
            existing_sheet_names = set()
            created_count = 0

            try:
                raw_order_ids = section_df[config.order_id_column].dropna().unique().tolist()
                order_ids = [safe_str(oid, "") for oid in raw_order_ids]
                valid_order_ids = [oid for oid in order_ids if oid and oid.lower() not in {"nan", "none", "null"}]

                for order_str in sorted(valid_order_ids):
                    order_df = section_df[section_df[config.order_id_column] == order_str]
                    if order_df.empty:
                        continue

                    new_ws = copy_invoice_sheet(template_ws, output_wb, "TempSheet")
                    if new_ws is None:
                        log_warning(
                            f"Could not copy invoice sheet for area '{area_value}' section '{section_name}' order '{order_str}'."
                        )
                        continue

                    new_ws.name = make_unique_sheet_name(order_str, existing_sheet_names)
                    render_invoice_sheet(new_ws, order_df, order_str, config)
                    created_count += 1

                if "TempSheet" in [sheet.name for sheet in output_wb.sheets]:
                    output_wb.sheets["TempSheet"].delete()

                if created_count == 0:
                    output_wb.close()
                    continue

                day_file = latest_delivery_date.strftime(f"%d-%m-%Y_{safe_area}_{section_name}")
                output_path = output_dir / f"{day_file}.xlsx"
                output_wb.save(str(output_path))
                saved_paths.append(str(output_path))
                output_wb.close()
            except Exception:
                try:
                    output_wb.close()
                except Exception:
                    pass
                raise
    finally:
        template_wb.close()

    return saved_paths


def process_area(area_value: str, area_df: pd.DataFrame, app: xw.App, config: ScriptConfig) -> List[str]:
    mode = safe_str(config.generation_mode, "per_order_section_workbooks").lower()
    if mode == "summary_by_section":
        saved = process_area_summary(area_value, area_df, app, config)
        return [saved] if saved else []
    if mode == "per_order_section_workbooks":
        return process_area_order_workbooks(area_value, area_df, app, config)

    log_warning(f"Unknown generation_mode '{config.generation_mode}'. Falling back to per_order_section_workbooks.")
    return process_area_order_workbooks(area_value, area_df, app, config)


def main() -> None:
    config = ScriptConfig()

    if not os.path.exists(config.orders_file):
        raise FileNotFoundError(f"Orders file not found: {config.orders_file}")

    orders_df = pd.read_excel(config.orders_file)
    orders_df = normalize_columns(orders_df)
    orders_df = apply_column_aliases(orders_df, COLUMN_ALIASES)
    orders_df = ensure_columns(orders_df, REQUIRED_COLUMNS_DEFAULTS, warn_missing=True)
    orders_df = ensure_columns(orders_df, OPTIONAL_COLUMNS_DEFAULTS, warn_missing=False)
    for dt_col in [config.delivery_date_column, "order_date", "order_time"]:
        if dt_col in orders_df.columns:
            orders_df[dt_col] = pd.to_datetime(orders_df[dt_col], errors="coerce")
    orders_df[config.order_id_column] = orders_df[config.order_id_column].apply(safe_str)
    orders_df[config.area_column] = orders_df[config.area_column].apply(safe_str)

    non_empty_areas = [a for a in orders_df[config.area_column].dropna().unique() if safe_str(a, "")]
    if not non_empty_areas:
        log_info("No areas found in orders data.")
        return

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    generated_files: List[str] = []
    try:
        for area in sorted(non_empty_areas):
            area_df = orders_df[orders_df[config.area_column] == area].copy()
            if area_df.empty:
                continue
            saved = process_area(area, area_df, app, config)
            if saved:
                generated_files.extend(saved)
    finally:
        app.quit()

    if generated_files:
        print("\nDONE: Generated invoice workbooks:")
        for path in generated_files:
            print(f" - {path}")
    else:
        print("DONE: No invoice files generated.")


if __name__ == "__main__":
    main()
