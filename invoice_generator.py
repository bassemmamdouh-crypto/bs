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
    target_by_supply_chain: bool = True
    section_column_candidates: List[str] = field(
        default_factory=lambda: ["section", "section_name", "business_unit", "category", "brand_section", "brand"]
    )
    supply_chain_column_candidates: List[str] = field(
        default_factory=lambda: ["supply_chain", "supply chain", "supply_chain_name", "supplychain"]
    )
    target_sections: List[str] = field(default_factory=lambda: ["PEPSI", "LAYS", "MARAII"])
    target_supply_chains: List[str] = field(default_factory=lambda: ["PEPSI", "LAYS", "MARAII"])
    supply_chain_section_map: Dict[str, List[str]] = field(
        default_factory=lambda: {
            "PEPSI": ["PEPSI"],
            "LAYS": ["LAYS"],
            "MARAII": ["MARAII"],
        }
    )
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
    price_before_discount_column_candidates: List[str] = field(
        default_factory=lambda: ["price_befor_discount", "price_before_discount"]
    )
    line_discount_column_candidates: List[str] = field(
        default_factory=lambda: ["total_discount", "line_discount", "discount_value"]
    )

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
    # Preferred: define one or many bundle offers here.
    # Each item:
    # {
    #   "source_skus": [...],
    #   "divisor": 6.0,
    #   "gift_sku": "",
    #   "gift_name": "Gift Name",
    #   # Optional extra condition:
    #   # apply only if total qty of these SKUs reaches minimum
    #   "condition_source_skus": [...],
    #   "condition_min_qty": 5,
    # }
    bundle_offers: List[Dict[str, object]] = field(
        default_factory=lambda: [
            {
                "source_skus": ["200", "201", "202", "203", "204", "205", "206", "207", "221"],
                "divisor": 6.0,
                "gift_sku": "",
                "gift_name": "عصير يومي برتقال 200 مل * 36",
            }
        ]
    )
    # Legacy single-bundle fields (kept for backward compatibility).
    # If bundle_offers has entries, these legacy fields are ignored.
    bundle_offer_active: bool = True
    bundle_offer_source_skus: List[str] = field(
        default_factory=lambda: ["200", "201", "202", "203", "204", "205", "206", "207", "221"]
    )
    bundle_offer_divisor: float = 6.0
    bundle_offer_gift_sku: str = ""
    bundle_offer_gift_name: str = "عصير يومي برتقال 200 مل * 36"

    # Invoice table + footer layout
    line_start_row: int = 16
    line_end_row: int = 291
    totals_row: int = 292
    subtotal_row: int = 293
    discount_row: int = 294
    adjustment_row: int = 295
    net_total_row: int = 296

    # Invoice line columns in template
    line_serial_col: int = 1
    line_item_name_col: int = 2
    line_qty_col: int = 3
    line_unit_col: int = 4
    line_gift_qty_col: int = 5
    line_price_before_discount_col: int = 6
    line_discount_value_col: int = 7
    line_total_price_col: int = 8

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
    "supply_chain": ["supply chain", "supply_chain_name", "supplychain"],
    "sku_name": ["item_name", "product_name", "item", "description"],
    "sku_code": ["sku", "product_code", "sku_id"],
    "price_befor_discount": ["price_before_discount"],
    "total_discount": ["discount_value", "line_discount"],
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
    "supply_chain": "",
    "price_befor_discount": 0,
    "total_discount": 0,
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


def normalize_identifier(value: object, default: str = "") -> str:
    """
    Normalize ID-like values coming from Excel so numeric IDs do not keep trailing .0.
    Examples: 123.0 -> "123", "00123" -> "00123", "A-12" -> "A-12"
    """
    text = safe_str(value, default)
    if not text:
        return default

    # Normalize Arabic digits commonly found in mixed datasets.
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    for i, d in enumerate(arabic_digits):
        text = text.replace(d, str(i))

    text = text.replace(",", "").strip()
    if re.fullmatch(r"[+-]?\d+\.0+", text):
        text = text.split(".", 1)[0]

    return text


def normalize_sku(value: object) -> str:
    return normalize_identifier(value, "")


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
    if (
        "MARAII" in text
        or "MARAII" in text
        or "MARAI" in text
        or "ALMARAI" in text
        or "AL MARAI" in text
        or "AL MARAII" in text
        or "المراعي" in text
    ):
        return "MARAII"
    return config.other_section_name


def normalize_supply_chain_name(value: object, config: ScriptConfig) -> str:
    text = safe_str(value, "").strip().upper()
    if not text:
        return config.other_section_name
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_group_key(value: object) -> str:
    return re.sub(r"\s+", " ", safe_str(value, "").strip().upper())


def build_supply_chain_lookup(config: ScriptConfig) -> Dict[str, str]:
    lookup: Dict[str, str] = {}
    for chain_name, section_names in (config.supply_chain_section_map or {}).items():
        normalized_chain = normalize_supply_chain_name(chain_name, config)
        if normalized_chain == config.other_section_name:
            continue
        for section_name in section_names:
            raw_key = normalize_group_key(section_name)
            if raw_key:
                lookup[raw_key] = normalized_chain
            normalized_section = normalize_section_name(section_name, config)
            normalized_section_key = normalize_group_key(normalized_section)
            if normalized_section_key:
                lookup[normalized_section_key] = normalized_chain
    return lookup


def get_target_sections(config: ScriptConfig) -> List[str]:
    target_sections: List[str] = []
    seen_sections = set()

    if config.target_by_supply_chain:
        target_source = config.target_supply_chains or []
        if not target_source and config.supply_chain_section_map:
            target_source = list(config.supply_chain_section_map.keys())
        if not target_source:
            target_source = config.target_sections

        for chain_name in target_source:
            normalized = normalize_supply_chain_name(chain_name, config)
            if normalized == config.other_section_name:
                continue
            if normalized not in seen_sections:
                seen_sections.add(normalized)
                target_sections.append(normalized)
    else:
        for sec in config.target_sections:
            normalized = normalize_section_name(sec, config)
            if normalized == config.other_section_name:
                continue
            if normalized not in seen_sections:
                seen_sections.add(normalized)
                target_sections.append(normalized)

    if not target_sections:
        return ["PEPSI", "LAYS", "MARAII"]
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
    Adjust line section size to match required_count:
    - insert rows before footer if items exceed template capacity
    - delete unused line rows before footer if items are fewer

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
    ws.range((start_row, 1), (end_row, 8)).value = None


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
        price_before_discount = safe_float(
            row_first_value(row, config.price_before_discount_column_candidates, 0),
            0.0,
        )
        line_discount_value = safe_float(
            row_first_value(row, config.line_discount_column_candidates, 0),
            0.0,
        )
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
                "price_before_discount": 0.0,
                "line_discount": 0.0,
                "net_amount": 0.0,
            }
        aggregated[key]["qty"] = safe_float(aggregated[key]["qty"], 0.0) + qty
        aggregated[key]["price_before_discount"] = (
            safe_float(aggregated[key]["price_before_discount"], 0.0) + price_before_discount
        )
        aggregated[key]["line_discount"] = (
            safe_float(aggregated[key]["line_discount"], 0.0) + line_discount_value
        )
        aggregated[key]["net_amount"] = safe_float(aggregated[key]["net_amount"], 0.0) + amount

    items = [
        item
        for item in aggregated.values()
        if (
            safe_float(item["qty"], 0.0) != 0.0
            or safe_float(item["net_amount"], 0.0) != 0.0
            or safe_float(item["price_before_discount"], 0.0) != 0.0
            or safe_float(item["line_discount"], 0.0) != 0.0
        )
    ]
    # Required ordering: brand first, then size.
    items.sort(key=sku_sort_key)
    return items


def quantity_by_sku(order_df: pd.DataFrame, config: ScriptConfig) -> Dict[str, float]:
    result: Dict[str, float] = {}
    for _, row in order_df.iterrows():
        sku = normalize_sku(row_first_value(row, config.sku_code_column_candidates, ""))
        if not sku:
            continue
        qty = safe_float(row_first_value(row, config.qty_column_candidates, 0), 0.0)
        result[sku] = result.get(sku, 0.0) + qty
    return result


def apply_offer_rules(
    order_df: pd.DataFrame,
    subtotal: float,
    config: ScriptConfig,
    bundle_offer_df: Optional[pd.DataFrame] = None,
) -> Dict[str, object]:
    offer_lines: List[Dict[str, object]] = []
    discount_total = 0.0
    sku_qty_map = quantity_by_sku(order_df, config)

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

    # B) Bundle offers:
    # Supports the new `bundle_offers` list and keeps legacy single-offer fields working.
    bundle_source_df = bundle_offer_df if bundle_offer_df is not None else order_df
    bundle_sku_qty_map = quantity_by_sku(bundle_source_df, config)

    bundle_offers: List[Dict[str, object]] = []
    configured_bundle_offers = getattr(config, "bundle_offers", None)
    if isinstance(configured_bundle_offers, list) and configured_bundle_offers:
        bundle_offers = configured_bundle_offers
    else:
        legacy_active = bool(getattr(config, "bundle_offer_active", False))
        legacy_divisor = safe_float(getattr(config, "bundle_offer_divisor", 0), 0.0)
        legacy_source_skus = getattr(config, "bundle_offer_source_skus", []) or []
        if legacy_active and legacy_divisor > 0 and legacy_source_skus:
            bundle_offers = [
                {
                    "source_skus": list(legacy_source_skus),
                    "divisor": legacy_divisor,
                    "gift_sku": safe_str(getattr(config, "bundle_offer_gift_sku", ""), ""),
                    "gift_name": safe_str(getattr(config, "bundle_offer_gift_name", "Gift Item"), "Gift Item"),
                }
            ]

    for bundle_offer in bundle_offers:
        source_skus = {
            normalize_sku(sku)
            for sku in (bundle_offer.get("source_skus", []) if isinstance(bundle_offer, dict) else [])
            if normalize_sku(sku)
        }
        divisor = safe_float(bundle_offer.get("divisor", 0) if isinstance(bundle_offer, dict) else 0, 0.0)
        if not source_skus or divisor <= 0:
            continue

        # Optional condition: require minimum qty from another SKU set before applying this offer.
        condition_skus_raw = []
        condition_min_qty = 0.0
        if isinstance(bundle_offer, dict):
            condition_skus_raw = (
                bundle_offer.get("condition_source_skus")
                or bundle_offer.get("required_source_skus")
                or bundle_offer.get("condition_skus")
                or []
            )
            condition_min_qty = safe_float(
                bundle_offer.get("condition_min_qty", bundle_offer.get("required_min_qty", 0)),
                0.0,
            )
        condition_source_skus = {
            normalize_sku(sku)
            for sku in condition_skus_raw
            if normalize_sku(sku)
        }
        if condition_min_qty > 0:
            if not condition_source_skus:
                # Misconfigured conditional offer: minimum exists without condition SKU set.
                continue
            condition_qty_total = sum(
                qty
                for sku, qty in bundle_sku_qty_map.items()
                if normalize_sku(sku) in condition_source_skus
            )
            if condition_qty_total < condition_min_qty:
                continue

        combo_qty_total = sum(
            qty
            for sku, qty in bundle_sku_qty_map.items()
            if normalize_sku(sku) in source_skus
        )
        combo_gift_qty = math.floor(combo_qty_total / divisor)
        if combo_gift_qty <= 0:
            continue

        gift_line_name = safe_str(
            bundle_offer.get("gift_name", "Gift Item") if isinstance(bundle_offer, dict) else "Gift Item",
            "Gift Item",
        )
        gift_sku = normalize_sku(bundle_offer.get("gift_sku", "") if isinstance(bundle_offer, dict) else "")
        if gift_sku:
            gift_line_name = f"SKU {gift_sku} - {gift_line_name}"

        offer_lines.append(
            {
                "item_name": gift_line_name,
                "qty": 0.0,
                "unit": "",
                "gift_qty": float(combo_gift_qty),
                "net_amount": 0.0,
            }
        )

    # C) BUY 1 GET 1 SAME SKU (180,181,182)
    same_sku_offer = {"180", "181", "182"}
    sku_gift_map: Dict[str, float] = {}

    for _, row in order_df.iterrows():
        sku = normalize_sku(row_first_value(row, config.sku_code_column_candidates, ""))
        if sku in same_sku_offer:
            qty = safe_float(row_first_value(row, config.qty_column_candidates, 0), 0.0)
            sku_gift_map[sku] = sku_gift_map.get(sku, 0.0) + qty

    for sku, total_qty in sku_gift_map.items():
        if total_qty <= 0:
            continue
        sku_rows = order_df[
            order_df.apply(
                lambda r: normalize_sku(row_first_value(r, config.sku_code_column_candidates, "")) == sku,
                axis=1,
            )
        ]
        gift_name = infer_line_item_name(sku_rows.iloc[0], config) if not sku_rows.empty else f"SKU {sku}"
        offer_lines.append(
            {
                "item_name": gift_name,
                "qty": 0.0,
                "unit": "",
                "gift_qty": total_qty,
                "net_amount": 0.0,
            }
        )

    # D) Configured rule-engine offers
    for rule in config.offer_rules:
        if not rule.active:
            continue

        if rule.rule_type == "buy_qty_get_free_same_sku":
            buy_sku = normalize_sku(rule.buy_sku)
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
            buy_sku = normalize_sku(rule.buy_sku)
            gift_sku = normalize_sku(rule.gift_sku)
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
        ws.range((row_no, config.line_serial_col)).value = idx + 1
        ws.range((row_no, config.line_item_name_col)).value = line.get("item_name", "")
        ws.range((row_no, config.line_qty_col)).value = safe_float(line.get("qty", 0), 0.0)
        ws.range((row_no, config.line_unit_col)).value = line.get("unit", "")
        ws.range((row_no, config.line_gift_qty_col)).value = safe_float(line.get("gift_qty", 0), 0.0)
        ws.range((row_no, config.line_price_before_discount_col)).value = safe_float(
            line.get("price_before_discount", 0),
            0.0,
        )
        ws.range((row_no, config.line_discount_value_col)).value = safe_float(
            line.get("line_discount", 0),
            0.0,
        )
        ws.range((row_no, config.line_total_price_col)).value = safe_float(line.get("net_amount", 0), 0.0)


def compute_invoice_totals(lines: List[Dict[str, object]], discount_amount: float) -> Dict[str, float]:
    total_qty = sum(safe_float(line.get("qty", 0), 0.0) + safe_float(line.get("gift_qty", 0), 0.0) for line in lines)
    total_price_before_discount = sum(safe_float(line.get("price_before_discount", 0), 0.0) for line in lines)
    total_line_discount = sum(safe_float(line.get("line_discount", 0), 0.0) for line in lines)
    subtotal = sum(safe_float(line.get("net_amount", 0), 0.0) for line in lines)
    discount = max(0.0, safe_float(discount_amount, 0.0))
    net_total = subtotal - discount
    return {
        "total_qty": total_qty,
        "total_price_before_discount": total_price_before_discount,
        "total_line_discount": total_line_discount,
        "subtotal": subtotal,
        "discount": discount,
        "net_total": net_total,
    }


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

    totals = compute_invoice_totals(lines, discount_amount)
    total_qty = totals["total_qty"]
    total_price_before_discount = totals["total_price_before_discount"]
    total_line_discount = totals["total_line_discount"]
    subtotal = totals["subtotal"]
    discount = totals["discount"]
    net_total = totals["net_total"]

    ws.range((totals_row, 3)).value = total_qty
    ws.range((totals_row, config.line_price_before_discount_col)).value = total_price_before_discount
    ws.range((totals_row, config.line_discount_value_col)).value = total_line_discount
    ws.range((totals_row, config.line_total_price_col)).value = subtotal
    ws.range((subtotal_row, 5)).value = subtotal
    ws.range((discount_row, 5)).value = discount
    ws.range((adjustment_row, 5)).value = 0
    ws.range((net_total_row, 5)).value = net_total


def prepare_area_sections(area_df: pd.DataFrame, config: ScriptConfig) -> pd.DataFrame:
    prepared = area_df.copy()
    section_col = find_existing_column(prepared, config.section_column_candidates)

    if not config.split_by_section:
        prepared["_section_group"] = config.other_section_name
        return prepared

    if config.target_by_supply_chain:
        supply_chain_col = find_existing_column(prepared, config.supply_chain_column_candidates)
        supply_chain_lookup = build_supply_chain_lookup(config)

        def resolve_supply_chain_group(row: pd.Series) -> str:
            if supply_chain_col:
                direct_chain = normalize_supply_chain_name(row.get(supply_chain_col, ""), config)
                if direct_chain != config.other_section_name:
                    return direct_chain

            if section_col:
                raw_section_key = normalize_group_key(row.get(section_col, ""))
                if raw_section_key in supply_chain_lookup:
                    return supply_chain_lookup[raw_section_key]

                raw_section_chain = normalize_supply_chain_name(row.get(section_col, ""), config)
                if raw_section_chain in supply_chain_lookup.values():
                    return raw_section_chain

                normalized_section = normalize_section_name(row.get(section_col, ""), config)
                normalized_section_key = normalize_group_key(normalized_section)
                if normalized_section_key in supply_chain_lookup:
                    return supply_chain_lookup[normalized_section_key]
                if normalized_section != config.other_section_name:
                    return normalize_supply_chain_name(normalized_section, config)

            return config.other_section_name

        prepared["_section_group"] = prepared.apply(resolve_supply_chain_group, axis=1)
    else:
        if section_col:
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


def add_workbook_summary_sheet(
    workbook: xw.Book,
    invoice_count: int,
    total_qty: float,
    total_prices: float,
    max_estimated_delivery_date: object,
    offer_breakdown: Dict[str, float],
) -> None:
    existing_names = {sheet.name for sheet in workbook.sheets}
    base_name = "SUMMARY"
    sheet_name = base_name
    suffix = 1
    while sheet_name in existing_names:
        sheet_name = f"{base_name}_{suffix}"[:31]
        suffix += 1

    summary_ws = workbook.sheets.add(name=sheet_name, after=workbook.sheets[-1])
    summary_ws["A1"].value = "Metric"
    summary_ws["B1"].value = "Value"
    summary_ws["A2"].value = "# of invoices in the sheet"
    summary_ws["B2"].value = int(invoice_count)
    summary_ws["A3"].value = "# of total quantities in all invoices"
    summary_ws["B3"].value = float(total_qty)
    summary_ws["A4"].value = "total prices of the invoices in the sheet"
    summary_ws["B4"].value = float(total_prices)
    summary_ws["A5"].value = "max estimated delivery date"
    summary_ws["B5"].value = safe_date_str(max_estimated_delivery_date, "")
    summary_ws["A6"].value = "total amount of gifts in the invoices (offers)"
    summary_ws["B6"].value = float(sum(safe_float(v, 0.0) for v in offer_breakdown.values()))

    summary_ws["A8"].value = "offer name"
    summary_ws["B8"].value = "total gift qty"
    offer_row = 9
    if offer_breakdown:
        for offer_name in sorted(offer_breakdown.keys(), key=lambda x: safe_str(x, "")):
            summary_ws.range((offer_row, 1)).value = offer_name
            summary_ws.range((offer_row, 2)).value = safe_float(offer_breakdown.get(offer_name, 0.0), 0.0)
            offer_row += 1
    else:
        summary_ws["A9"].value = "No offer gifts"
        summary_ws["B9"].value = 0.0

    # Visual formatting to make the sheet readable as a table.
    summary_ws.range("A:B").autofit()
    if safe_float(summary_ws.range("A:A").column_width, 0.0) < 34:
        summary_ws.range("A:A").column_width = 34
    if safe_float(summary_ws.range("B:B").column_width, 0.0) < 18:
        summary_ws.range("B:B").column_width = 18

    # Global font defaults for summary table.
    summary_ws.range("A:B").api.Font.Name = "Calibri"
    summary_ws.range("A:B").api.Font.Size = 11

    # Main header style.
    summary_ws.range("A1:B1").color = (47, 84, 150)
    summary_ws.range("A1:B1").api.Font.Color = 16777215
    summary_ws.range("A1:B1").api.Font.Bold = True
    summary_ws.range("A1:B1").api.HorizontalAlignment = -4108  # Center

    # Metric labels and values styles.
    summary_ws.range("A2:A6").api.Font.Bold = True
    summary_ws.range("A2:A6").api.HorizontalAlignment = -4152  # Right
    summary_ws.range("B2:B6").api.HorizontalAlignment = -4108  # Center
    summary_ws["B2"].number_format = "0"
    summary_ws["B3"].number_format = "#,##0.00"
    summary_ws["B4"].number_format = "#,##0.00"
    summary_ws["B5"].number_format = "yyyy-mm-dd"
    summary_ws["B6"].number_format = "#,##0.00"

    # Offers table header style.
    summary_ws.range("A8:B8").color = (217, 217, 217)
    summary_ws.range("A8:B8").api.Font.Bold = True
    summary_ws.range("A8:B8").api.HorizontalAlignment = -4108  # Center

    # Offer rows formatting.
    last_offer_row = max(9, offer_row - 1)
    summary_ws.range((9, 1), (last_offer_row, 1)).api.HorizontalAlignment = -4152  # Right
    summary_ws.range((9, 2), (last_offer_row, 2)).api.HorizontalAlignment = -4108  # Center
    summary_ws.range((9, 2), (last_offer_row, 2)).number_format = "#,##0.00"

    # Apply borders to both tables.
    summary_ws.range("A1:B6").api.Borders.LineStyle = 1
    summary_ws.range("A1:B6").api.Borders.Weight = 2
    summary_ws.range((8, 1), (last_offer_row, 2)).api.Borders.LineStyle = 1
    summary_ws.range((8, 1), (last_offer_row, 2)).api.Borders.Weight = 2


def render_invoice_sheet(
    ws: xw.Sheet,
    source_df: pd.DataFrame,
    header_order_id: str,
    config: ScriptConfig,
    bundle_offer_df: Optional[pd.DataFrame] = None,
) -> Dict[str, object]:
    write_header(ws, source_df, header_order_id, config)
    base_lines = build_base_items(source_df, config)
    subtotal = sum(safe_float(line.get("net_amount", 0), 0.0) for line in base_lines)
    offer_result = apply_offer_rules(source_df, subtotal, config, bundle_offer_df=bundle_offer_df)

    final_lines = list(base_lines)
    if config.include_offer_lines:
        final_lines.extend(offer_result["offer_lines"])

    discount_total = safe_float(offer_result["discount_total"], 0.0)
    totals = compute_invoice_totals(final_lines, discount_total)
    row_shift = adjust_line_capacity(ws, len(final_lines), config)
    fill_invoice_lines(ws, final_lines, config, row_shift)
    write_totals(ws, final_lines, discount_total, config, row_shift)
    offer_breakdown: Dict[str, float] = {}
    for offer_line in offer_result["offer_lines"]:
        offer_name = safe_str(offer_line.get("item_name", "Offer"), "Offer")
        gift_qty = safe_float(offer_line.get("gift_qty", 0.0), 0.0)
        if gift_qty <= 0:
            continue
        offer_breakdown[offer_name] = offer_breakdown.get(offer_name, 0.0) + gift_qty

    return {
        "invoice_count": 1.0,
        "total_qty": totals["total_qty"],
        "total_prices": totals["subtotal"],
        "offer_breakdown": offer_breakdown,
    }


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
    summary_invoice_count = 0
    summary_total_qty = 0.0
    summary_total_prices = 0.0
    summary_offer_breakdown: Dict[str, float] = {}

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
            metrics = render_invoice_sheet(new_ws, section_df, f"{section_name} SUMMARY", config)
            summary_invoice_count += int(metrics["invoice_count"])
            summary_total_qty += metrics["total_qty"]
            summary_total_prices += metrics["total_prices"]
            for offer_name, gift_qty in metrics["offer_breakdown"].items():
                summary_offer_breakdown[offer_name] = (
                    summary_offer_breakdown.get(offer_name, 0.0) + safe_float(gift_qty, 0.0)
                )

    if "TempSheet" in [sheet.name for sheet in output_wb.sheets]:
        output_wb.sheets["TempSheet"].delete()
    max_delivery_date = pd.to_datetime(area_df[config.delivery_date_column], errors="coerce").max()
    add_workbook_summary_sheet(
        output_wb,
        summary_invoice_count,
        summary_total_qty,
        summary_total_prices,
        max_delivery_date,
        summary_offer_breakdown,
    )

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
    Create one workbook per target section per area (if section rows exist).
    Each workbook contains one invoice sheet per order_id for its section.
    """
    template_file = choose_template_file(area_value, config)
    if not os.path.exists(template_file):
        log_warning(f"Template file missing for area '{area_value}': {template_file}")
        return []

    if area_df.empty:
        log_info(f"No rows found in area '{area_value}'.")
        return []

    area_df = prepare_area_sections(area_df, config)
    area_df["_order_key"] = area_df[config.order_id_column].apply(lambda x: normalize_identifier(x, ""))
    area_df = area_df[area_df["_order_key"] != ""].copy()
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
            section_total_qty = 0.0
            section_total_prices = 0.0
            section_offer_breakdown: Dict[str, float] = {}

            try:
                raw_order_ids = section_df["_order_key"].dropna().unique().tolist()
                valid_order_ids = [
                    oid
                    for oid in [normalize_identifier(oid, "") for oid in raw_order_ids]
                    if oid and oid.lower() not in {"nan", "none", "null"}
                ]

                log_info(
                    f"Area '{area_value}' section '{section_name}': "
                    f"{len(valid_order_ids)} order invoice sheet(s) will be created."
                )

                for order_str in sorted(valid_order_ids):
                    order_df = section_df[section_df["_order_key"] == order_str]
                    if order_df.empty:
                        continue
                    full_order_df = area_df[area_df["_order_key"] == order_str]

                    new_ws = copy_invoice_sheet(template_ws, output_wb, "TempSheet")
                    if new_ws is None:
                        log_warning(
                            f"Could not copy invoice sheet for area '{area_value}' section '{section_name}' order '{order_str}'."
                        )
                        continue

                    new_ws.name = make_unique_sheet_name(order_str, existing_sheet_names)
                    metrics = render_invoice_sheet(new_ws, order_df, order_str, config, bundle_offer_df=full_order_df)
                    created_count += 1
                    section_total_qty += metrics["total_qty"]
                    section_total_prices += metrics["total_prices"]
                    for offer_name, gift_qty in metrics["offer_breakdown"].items():
                        section_offer_breakdown[offer_name] = (
                            section_offer_breakdown.get(offer_name, 0.0) + safe_float(gift_qty, 0.0)
                        )

                if "TempSheet" in [sheet.name for sheet in output_wb.sheets]:
                    output_wb.sheets["TempSheet"].delete()

                if created_count == 0:
                    output_wb.close()
                    continue

                max_delivery_date = pd.to_datetime(section_df[config.delivery_date_column], errors="coerce").max()
                add_workbook_summary_sheet(
                    output_wb,
                    created_count,
                    section_total_qty,
                    section_total_prices,
                    max_delivery_date,
                    section_offer_breakdown,
                )
                day_file = latest_delivery_date.strftime(f"%d-%m-%Y_{safe_area}_{section_name}_ORDERS")
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
    orders_df[config.order_id_column] = orders_df[config.order_id_column].apply(lambda x: normalize_identifier(x, ""))
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
