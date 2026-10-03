import math
import os
import re
from dataclasses import dataclass, field
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
class LoadingPaperConfig:
    orders_file: str = ORDERS_FILE
    template_file: str = LOADING_TEMPLATE_FILE
    output_root: str = OUTPUT_ROOT

    loading_sheet_candidates: Tuple[str, ...] = ("Loading paper", "loading paper", "loading")

    # Source columns
    route_agent_candidates: Tuple[str, ...] = ("route_agent", "driver")
    run_candidates: Tuple[str, ...] = ("run", "trip", "run_name")
    delivery_date_candidates: Tuple[str, ...] = ("estimated_delivery_date", "delivery_date")
    brand_candidates: Tuple[str, ...] = ("brand_name_en", "brand_name", "brand")
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
    order_id_candidates: Tuple[str, ...] = ("order_id", "order_number", "orderid", "order_no")
    sku_code_candidates: Tuple[str, ...] = ("sku_code", "sku", "product_id", "item_id", "sku_id", "product_code")

    # Optional offer columns (same offer quantities logic as invoice script)
    offer_item_name_candidates: Tuple[str, ...] = ("offer_item_name", "offer_name", "promo_name")
    offer_item_qty_candidates: Tuple[str, ...] = ("offer_item_qty", "offer_qty", "promo_qty")
    gift_qty_columns: Tuple[str, ...] = (
        "second_run_gift",
        "doritos_gift",
        "pepsi_gift",
        "pepsi_gift_qty",
        "pepsi_offer_gift",
        "pepsi_offer_qty",
        "gift_qty",
        "free_qty",
    )

    # Preferred: define one or many bundle offers here.
    # Each item:
    # {
    #   "active": True,
    #   "source_skus": [...],                # optional
    #   "source_brand_names": [...],         # optional (OR with source_skus)
    #   "divisor": 6.0,
    #   "gift_sku": "",
    #   "gift_name": "Gift Name",
    #   "gift_size": "",
    #   # Optional condition:
    #   "condition_source_skus": [...],      # optional
    #   "condition_brand_names": [...],      # optional (OR with condition_source_skus)
    #   "condition_min_qty": 5,
    # }
    bundle_offers: List[Dict[str, object]] = field(
        default_factory=lambda: [
            {
                "active": True,
                "source_skus": ["200", "201", "202", "203", "204", "205", "206", "207", "221"],
                "divisor": 6.0,
                "gift_sku": "",
                "gift_name": "عصير يومي برتقال 200 مل * 36",
                "gift_size": "200 مل",
            }
        ]
    )
    # Legacy single-bundle fields (kept for backward compatibility).
    # If bundle_offers has entries, these legacy fields are ignored.
    bundle_offer_active: bool = True
    bundle_offer_source_skus: Tuple[str, ...] = ("200", "201", "202", "203", "204", "205", "206", "207", "221")
    bundle_offer_divisor: float = 6.0
    bundle_offer_gift_sku: str = ""
    bundle_offer_gift_name: str = "عصير يومي برتقال 200 مل * 36"
    bundle_offer_gift_brand: str = "youmy juice"
    bundle_offer_gift_size: str = "200 مل"
    same_sku_offer_skus: Tuple[str, ...] = ("180", "181", "182")
    offer_fallback_brand: str = "عروض"
    offer_group_order: int = 10000
    offer_rules: List[OfferRule] = field(
        default_factory=lambda: [
            OfferRule(
                name="Buy10Get1-SKU180",
                rule_type="buy_qty_get_free_same_sku",
                active=False,
                buy_sku="180",
                buy_qty=10,
                gift_qty=1,
                gift_name="Offer Gift SKU 180",
            ),
            OfferRule(
                name="Subtotal-5pct-over-1000",
                rule_type="order_subtotal_discount_pct",
                active=False,
                min_subtotal=1000,
                discount_percent=5,
            ),
        ]
    )

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
    "order_id": ["order_number", "orderid", "order_no"],
    "sku_code": ["sku", "product_id", "item_id", "sku_id", "product_code"],
    "brand_name_en": ["brand_name", "brand"],
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


def normalize_sku(value: object) -> str:
    return normalize_identifier(value, "")


def normalize_column_key(name: object) -> str:
    return re.sub(r"[^a-z0-9]", "", safe_str(name, "").lower())


def extract_row_sku(row: pd.Series) -> str:
    """
    Resolve SKU robustly from normalized `_sku` first, then from any SKU-like raw source column.
    This prevents bundle-offer misses when input headers vary unexpectedly.
    """
    direct = normalize_sku(row.get("_sku", ""))
    if direct:
        return direct

    explicit_keys = {
        "sku",
        "skucode",
        "skuid",
        "productid",
        "itemid",
        "productcode",
        "productsku",
        "itemsku",
    }
    for key, value in row.items():
        key_norm = normalize_column_key(key)
        if not key_norm:
            continue
        looks_like_sku = (
            key_norm in explicit_keys
            or (key_norm.endswith("sku") and "name" not in key_norm)
            or key_norm.endswith("skuid")
            or key_norm.endswith("productid")
            or key_norm.endswith("itemid")
            or key_norm.endswith("productcode")
        )
        if not looks_like_sku:
            continue
        candidate = normalize_sku(value)
        if candidate:
            return candidate
    return ""


def to_selector_list(value: object) -> List[object]:
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        return list(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return []
        return [part.strip() for part in re.split(r"[,;\n\r\t|]+", text) if part.strip()]
    return [value]


def parse_sku_selector_set(value: object) -> set:
    parsed: set = set()
    for item in to_selector_list(value):
        text = safe_str(item, "")
        if not text:
            continue
        # Support whitespace-separated numeric SKU lists in one cell/string.
        parts = [text]
        if re.search(r"\s", text) and not re.search(r"[A-Za-z\u0600-\u06FF]", text):
            parts = [p for p in re.split(r"\s+", text) if p]
        for part in parts:
            sku = normalize_sku(part)
            if sku:
                parsed.add(sku)
    return parsed


def parse_brand_selector_set(value: object) -> set:
    parsed: set = set()
    for item in to_selector_list(value):
        brand = normalize_brand_key(item)
        if brand:
            parsed.add(brand)
    return parsed


def normalize_brand_for_order(brand_value: object) -> str:
    # Keep brand_name-driven grouping, but normalize formatting for reliable ordering.
    text = safe_str(brand_value, "").lower().strip()
    text = text.replace("’", "'")
    text = text.replace("_", " ")
    text = text.replace("-", " ")
    text = re.sub(r"\s+", " ", text)
    text = text.replace("'", "")
    return text


def normalize_brand_key(value: object) -> str:
    return normalize_brand_for_order(value)


def normalize_size_for_order(size_value: object) -> str:
    return re.sub(r"\s+", " ", safe_str(size_value, "")).strip()


def aggregate_order_level_numeric(order_df: pd.DataFrame, column_name: str) -> float:
    if column_name not in order_df.columns:
        return 0.0
    series = pd.to_numeric(order_df[column_name], errors="coerce").fillna(0)
    non_zero = series[series != 0]
    if non_zero.empty:
        return 0.0
    unique = non_zero.unique()
    if len(unique) == 1 and len(non_zero) > 1:
        return float(unique[0])
    return float(non_zero.sum())


def build_offer_rows_for_order(order_df: pd.DataFrame, config: LoadingPaperConfig) -> List[Dict[str, object]]:
    offer_rows: List[Dict[str, object]] = []
    effective_sku = order_df.apply(extract_row_sku, axis=1)
    sku_qty_map: Dict[str, float] = {}
    for idx, row in order_df.iterrows():
        sku = safe_str(effective_sku.get(idx, ""), "")
        if not sku:
            continue
        qty = safe_float(row.get("_qty", 0), 0.0)
        sku_qty_map[sku] = sku_qty_map.get(sku, 0.0) + qty

    # A) Gift quantity columns from source sheet.
    for gift_col in config.gift_qty_columns:
        gift_qty = aggregate_order_level_numeric(order_df, gift_col)
        if gift_qty > 0:
            offer_rows.append(
                {
                    "_product": gift_col.replace("_", " ").title(),
                    "_qty": gift_qty,
                    "_brand_raw": config.offer_fallback_brand,
                    "_size_raw": "",
                    "_sku": "",
                }
            )

    # B) Explicit offer item rows from source sheet.
    for _, row in order_df.iterrows():
        offer_name = safe_str(row.get("_offer_item_name", ""), "")
        offer_qty = safe_float(row.get("_offer_item_qty", 0), 0.0)
        if offer_name and offer_qty > 0:
            offer_rows.append(
                {
                    "_product": offer_name,
                    "_qty": offer_qty,
                    "_brand_raw": config.offer_fallback_brand,
                    "_size_raw": "",
                    "_sku": "",
                }
            )

    # C) Bundle offers:
    # supports list-based bundle_offers plus legacy single-bundle fallback.
    bundle_rows: List[Dict[str, object]] = []
    for idx, row in order_df.iterrows():
        sku_value = safe_str(effective_sku.get(idx, ""), "")
        qty_value = safe_float(row.get("_qty", 0), 0.0)
        if qty_value == 0:
            continue
        bundle_rows.append(
            {
                "sku": sku_value,
                "qty": qty_value,
                "brand_key": normalize_brand_key(row.get("_brand_raw", "")),
            }
        )
    has_any_bundle_sku = any(safe_str(bundle_row.get("sku", ""), "") for bundle_row in bundle_rows)

    bundle_offers: List[Dict[str, object]] = []
    if isinstance(config.bundle_offers, dict):
        bundle_offers = [config.bundle_offers]
    elif isinstance(config.bundle_offers, list) and config.bundle_offers:
        bundle_offers = config.bundle_offers
    elif config.bundle_offer_active and config.bundle_offer_divisor > 0:
        bundle_offers = [
            {
                "active": True,
                "source_skus": list(config.bundle_offer_source_skus),
                "divisor": config.bundle_offer_divisor,
                "gift_sku": config.bundle_offer_gift_sku,
                "gift_name": config.bundle_offer_gift_name,
                "gift_size": config.bundle_offer_gift_size,
            }
        ]

    for bundle_offer in bundle_offers:
        if not isinstance(bundle_offer, dict):
            continue
        if not bool(bundle_offer.get("active", True)):
            continue

        source_skus_raw = (
            bundle_offer.get("source_skus")
            or bundle_offer.get("source_sku")
            or bundle_offer.get("source_sku_ids")
            or bundle_offer.get("source_ids")
            or bundle_offer.get("skus")
            or []
        )
        source_skus = parse_sku_selector_set(source_skus_raw)
        source_brands_raw = (
            bundle_offer.get("source_brand_names")
            or bundle_offer.get("source_brand_name")
            or bundle_offer.get("source_brands")
            or bundle_offer.get("source_brand_name_en")
            or bundle_offer.get("source_brands_en")
            or bundle_offer.get("brands")
            or []
        )
        source_brand_keys = parse_brand_selector_set(source_brands_raw)
        if source_skus and not source_brand_keys and not has_any_bundle_sku:
            log_warning("Bundle offer source SKUs configured, but no SKU values were resolved for current group.")
            continue
        divisor = safe_float(bundle_offer.get("divisor", 0), 0.0)
        if divisor <= 0 or (not source_skus and not source_brand_keys):
            log_warning(
                "Bundle offer skipped due to invalid selectors or divisor. "
                f"divisor={divisor}, source_skus={len(source_skus)}, source_brands={len(source_brand_keys)}"
            )
            continue

        condition_source_skus_raw = (
            bundle_offer.get("condition_source_skus")
            or bundle_offer.get("condition_source_sku")
            or bundle_offer.get("required_source_skus")
            or bundle_offer.get("condition_skus")
            or bundle_offer.get("condition_sku_ids")
            or []
        )
        condition_source_skus = parse_sku_selector_set(condition_source_skus_raw)
        condition_brand_raw = (
            bundle_offer.get("condition_brand_names")
            or bundle_offer.get("condition_brand_name")
            or bundle_offer.get("condition_brands")
            or bundle_offer.get("condition_brand_name_en")
            or bundle_offer.get("condition_brands_en")
            or []
        )
        condition_brand_keys = parse_brand_selector_set(condition_brand_raw)
        condition_min_qty = safe_float(
            bundle_offer.get("condition_min_qty", bundle_offer.get("required_min_qty", 0)),
            0.0,
        )
        if condition_min_qty > 0:
            if not condition_source_skus and not condition_brand_keys:
                continue
            condition_qty_total = sum(
                safe_float(bundle_row.get("qty", 0), 0.0)
                for bundle_row in bundle_rows
                if (
                    (
                        bool(condition_source_skus)
                        and safe_str(bundle_row.get("sku", ""), "") in condition_source_skus
                    )
                    or (
                        bool(condition_brand_keys)
                        and safe_str(bundle_row.get("brand_key", ""), "") in condition_brand_keys
                    )
                )
            )
            if condition_qty_total < condition_min_qty:
                continue

        combo_qty_total = sum(
            safe_float(bundle_row.get("qty", 0), 0.0)
            for bundle_row in bundle_rows
            if (
                (
                    bool(source_skus)
                    and safe_str(bundle_row.get("sku", ""), "") in source_skus
                )
                or (
                    bool(source_brand_keys)
                    and safe_str(bundle_row.get("brand_key", ""), "") in source_brand_keys
                )
            )
        )
        combo_gift_qty = math.floor(combo_qty_total / divisor)
        if combo_gift_qty <= 0:
            continue

        gift_name = safe_str(bundle_offer.get("gift_name", "Gift Item"), "Gift Item")
        gift_sku = normalize_sku(bundle_offer.get("gift_sku", ""))
        gift_size = safe_str(bundle_offer.get("gift_size", ""), "")
        if gift_sku:
            sku_rows = order_df[effective_sku == gift_sku]
            if not sku_rows.empty:
                gift_name = safe_str(sku_rows.iloc[0].get("_product", gift_name), gift_name)
                gift_size = safe_str(sku_rows.iloc[0].get("_size_raw", gift_size), gift_size)
            else:
                gift_name = f"SKU {gift_sku} - {gift_name}"

        offer_rows.append(
            {
                "_product": gift_name,
                "_qty": float(combo_gift_qty),
                "_brand_raw": config.offer_fallback_brand,
                "_size_raw": gift_size,
                "_sku": gift_sku,
            }
        )

    # D) BUY 1 GET 1 SAME SKU (same as invoice script logic).
    same_sku_offer = parse_sku_selector_set(config.same_sku_offer_skus)
    for sku in sorted(same_sku_offer):
        sku_rows = order_df[effective_sku == sku]
        if sku_rows.empty:
            continue
        total_qty = safe_float(sku_rows["_qty"].sum(), 0.0)
        if total_qty <= 0:
            continue
        first = sku_rows.iloc[0]
        offer_rows.append(
            {
                "_product": safe_str(first.get("_product", f"SKU {sku}"), f"SKU {sku}"),
                "_qty": total_qty,
                "_brand_raw": config.offer_fallback_brand,
                "_size_raw": safe_str(first.get("_size_raw", ""), ""),
                "_sku": sku,
            }
        )

    # E) Configured rule-engine offers (gift quantity only for loading paper).
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
                buy_rows = order_df[effective_sku == buy_sku]
                fallback_name = safe_str(rule.gift_name, "") or f"Offer Gift SKU {buy_sku}"
                inferred_name = (
                    safe_str(buy_rows.iloc[0].get("_product", fallback_name), fallback_name)
                    if not buy_rows.empty
                    else fallback_name
                )
                inferred_brand = (
                    safe_str(buy_rows.iloc[0].get("_brand_raw", config.offer_fallback_brand), config.offer_fallback_brand)
                    if not buy_rows.empty
                    else config.offer_fallback_brand
                )
                inferred_size = safe_str(buy_rows.iloc[0].get("_size_raw", ""), "") if not buy_rows.empty else ""
                offer_rows.append(
                    {
                        "_product": inferred_name,
                        "_qty": gift_qty,
                        "_brand_raw": config.offer_fallback_brand,
                        "_size_raw": inferred_size,
                        "_sku": buy_sku,
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
                gift_rows = order_df[effective_sku == gift_sku]
                fallback_name = safe_str(rule.gift_name, "") or f"Offer Gift SKU {gift_sku}"
                inferred_name = (
                    safe_str(gift_rows.iloc[0].get("_product", fallback_name), fallback_name)
                    if not gift_rows.empty
                    else fallback_name
                )
                inferred_brand = (
                    safe_str(gift_rows.iloc[0].get("_brand_raw", config.offer_fallback_brand), config.offer_fallback_brand)
                    if not gift_rows.empty
                    else config.offer_fallback_brand
                )
                inferred_size = safe_str(gift_rows.iloc[0].get("_size_raw", ""), "") if not gift_rows.empty else ""
                offer_rows.append(
                    {
                        "_product": inferred_name,
                        "_qty": gift_qty,
                        "_brand_raw": config.offer_fallback_brand,
                        "_size_raw": inferred_size,
                        "_sku": gift_sku,
                    }
                )

    # Merge duplicate offer lines by product name (same behavior as invoices).
    merged: Dict[str, Dict[str, object]] = {}
    for line in offer_rows:
        name = safe_str(line.get("_product", "Offer"), "Offer")
        if name not in merged:
            merged[name] = {
                "_product": name,
                "_qty": 0.0,
                "_brand_raw": safe_str(line.get("_brand_raw", config.offer_fallback_brand), config.offer_fallback_brand),
                "_size_raw": safe_str(line.get("_size_raw", ""), ""),
                "_sku": normalize_sku(line.get("_sku", "")),
            }
        merged[name]["_qty"] = safe_float(merged[name]["_qty"], 0.0) + safe_float(line.get("_qty", 0), 0.0)

    return sorted(list(merged.values()), key=lambda x: safe_str(x["_product"], ""))


def append_offer_quantities(df: pd.DataFrame, config: LoadingPaperConfig) -> pd.DataFrame:
    if df.empty:
        return df
    if "_sku" not in df.columns:
        return df

    offer_records: List[Dict[str, object]] = []
    has_order_col = "_order_id" in df.columns
    has_non_empty_orders = has_order_col and not df[df["_order_id"] != ""].empty
    grouped_batches: List[Tuple[str, object]] = []
    if has_non_empty_orders:
        grouped_batches.append(
            (
                "order",
                df[df["_order_id"] != ""].groupby(["_route_agent", "_run", "_order_id"], sort=False),
            )
        )
        if has_order_col and not df[df["_order_id"] == ""].empty:
            log_warning(
                "Some rows have empty order_id. Those rows are excluded from offer qualification "
                "because offers are computed per order, then summed in loading output."
            )
    else:
        # Fallback when order_id is missing: compute at route/run level so gifts are still visible.
        grouped_batches.append(("route-run", df.groupby(["_route_agent", "_run"], sort=False)))

    total_offer_qty_added = 0.0
    used_fallback_group = False
    for group_mode, grouped_iter in grouped_batches:
        for group_key, order_df in grouped_iter:
            rows = build_offer_rows_for_order(order_df, config)
            if not rows:
                continue
            route_agent = safe_str(order_df.iloc[0].get("_route_agent", ""), "")
            run_name = safe_str(order_df.iloc[0].get("_run", ""), "")
            order_id = safe_str(order_df.iloc[0].get("_order_id", ""), "") if group_mode == "order" else ""
            delivery_date = pd.to_datetime(order_df["_delivery_date"], errors="coerce").max()
            if group_mode != "order":
                used_fallback_group = True
            for row in rows:
                offer_qty = safe_float(row.get("_qty", 0), 0.0)
                total_offer_qty_added += offer_qty
                offer_records.append(
                    {
                        "_route_agent": route_agent,
                        "_run": run_name,
                        "_order_id": order_id,
                        "_product": safe_str(row.get("_product", "Offer"), "Offer"),
                        "_qty": offer_qty,
                        "_brand_raw": safe_str(row.get("_brand_raw", config.offer_fallback_brand), config.offer_fallback_brand),
                        "_size_raw": safe_str(row.get("_size_raw", ""), ""),
                        "_sku": normalize_sku(row.get("_sku", "")),
                        "_delivery_date": delivery_date,
                        "_offer_item_name": "",
                        "_offer_item_qty": 0.0,
                    }
                )

    if not offer_records:
        if used_fallback_group or not has_non_empty_orders:
            log_warning("Offer rows were computed by route/run because order_id was missing.")
        if isinstance(config.bundle_offers, list) and config.bundle_offers:
            log_warning(
                "No bundle offer rows were added from qualified orders. "
                "Verify selectors/conditions against per-order data."
            )
        return df

    offers_df = pd.DataFrame(offer_records)
    if used_fallback_group or not has_non_empty_orders:
        log_warning("Added offer rows using route/run fallback because order_id was missing.")
    log_info(
        f"Added {len(offers_df)} offer row(s) with total gift qty {total_offer_qty_added:.2f}."
    )
    return pd.concat([df, offers_df], ignore_index=True, sort=False)


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
    normalized_index = {normalize_column_key(c): c for c in df.columns}
    for col in candidates:
        normalized_col = normalized_index.get(normalize_column_key(col))
        if normalized_col:
            return normalized_col
    for col in candidates:
        col_key = normalize_column_key(col)
        if not col_key:
            continue
        for existing_col in df.columns:
            existing_key = normalize_column_key(existing_col)
            if col_key in existing_key:
                return existing_col
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
    base = f"LOAD_{clean}" if clean else "LOAD"
    candidate = base[:31]
    i = 1
    while candidate in existing_names:
        suffix = f"_{i}"
        trim_len = max(0, 31 - len(suffix))
        candidate = f"{base[:trim_len]}{suffix}"[-31:]
        i += 1
    existing_names.add(candidate)
    return candidate


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
        try:
            insert_end = config.summary_anchor_row + requested_shift - 1
            ws.api.Rows(f"{config.summary_anchor_row}:{insert_end}").Insert()
            return requested_shift
        except Exception as exc:
            log_warning(f"Bulk insert failed; falling back to row insert loop: {exc}")
            inserted = 0
            for _ in range(requested_shift):
                try:
                    ws.api.Rows(config.summary_anchor_row).Insert()
                    inserted += 1
                except Exception as row_exc:
                    log_warning(f"Unable to insert extra loading rows before summary: {row_exc}")
                    break
            return inserted

    if requested_shift < 0:
        delete_count = abs(requested_shift)
        delete_start_row = config.start_row + target
        delete_end_row = delete_start_row + delete_count - 1
        try:
            ws.api.Rows(f"{delete_start_row}:{delete_end_row}").Delete()
            return -delete_count
        except Exception as exc:
            log_warning(f"Bulk delete failed; falling back to row delete loop: {exc}")
            deleted = 0
            for _ in range(delete_count):
                try:
                    ws.api.Rows(delete_start_row).Delete()
                    deleted += 1
                except Exception as row_exc:
                    log_warning(f"Unable to delete unused loading row before summary: {row_exc}")
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
    order_id_col = find_existing_column(df, config.order_id_candidates)
    sku_col = find_existing_column(df, config.sku_code_candidates)
    offer_item_name_col = find_existing_column(df, config.offer_item_name_candidates)
    offer_item_qty_col = find_existing_column(df, config.offer_item_qty_candidates)

    if route_col is None or run_col is None or product_col is None or qty_col is None:
        raise ValueError(
            "Missing required columns for loading paper generation. "
            "Need route_agent/driver, run/trip, product name, and quantity."
        )
    log_info(
        "Resolved columns | "
        f"route={route_col}, run={run_col}, product={product_col}, qty={qty_col}, "
        f"order_id={order_id_col}, sku={sku_col}, brand={brand_col}, size={size_col}"
    )
    if sku_col is None:
        log_warning(
            "SKU column was not explicitly matched. Offers will attempt SKU fallback from any SKU-like columns."
        )

    brand_rank_map = {
        normalize_brand_for_order(key): rank
        for key, rank in (config.brand_order or {}).items()
    }
    # Keep offers in one dedicated group sorted after all product brands.
    brand_rank_map[normalize_brand_for_order(config.offer_fallback_brand)] = config.offer_group_order
    size_rank_map = {
        normalize_size_for_order(key): rank
        for key, rank in (config.size_order or {}).items()
    }

    out = df.copy()
    out["_route_agent"] = out[route_col].apply(lambda x: safe_str(x, "UNKNOWN-AGENT"))
    out["_run"] = out[run_col].apply(lambda x: normalize_identifier(x, "UNKNOWN-RUN"))
    out["_order_id"] = out[order_id_col].apply(lambda x: normalize_identifier(x, "")) if order_id_col is not None else ""
    out["_sku"] = out[sku_col].apply(normalize_sku) if sku_col is not None else ""
    out["_product"] = out[product_col].apply(lambda x: safe_str(x, "Unnamed Item"))
    out["_qty"] = out[qty_col].apply(lambda x: safe_float(x, 0.0))
    out["_offer_item_name"] = (
        out[offer_item_name_col].apply(lambda x: safe_str(x, ""))
        if offer_item_name_col is not None
        else ""
    )
    out["_offer_item_qty"] = (
        out[offer_item_qty_col].apply(lambda x: safe_float(x, 0.0))
        if offer_item_qty_col is not None
        else 0.0
    )
    out["_brand_raw"] = (
        out[brand_col].apply(lambda x: safe_str(x, "OTHER"))
        if brand_col is not None
        else out["_product"].apply(lambda x: safe_str(x.split(" ")[0], "OTHER"))
    )
    out["_size_raw"] = out[size_col].apply(lambda x: safe_str(x, "")) if size_col is not None else ""
    if delivery_col is not None:
        out["_delivery_date"] = pd.to_datetime(out[delivery_col], errors="coerce")
    else:
        out["_delivery_date"] = pd.NaT

    out = append_offer_quantities(out, config)

    # Compute grouping/sorting keys after offers are appended so offer rows
    # follow the same brand/size grouping and ordering pipeline.
    out["_brand_key"] = out["_brand_raw"].apply(normalize_brand_for_order)
    out["_size_key"] = out["_size_raw"].apply(normalize_size_for_order)
    out["_brand_display"] = out["_brand_raw"].apply(lambda x: safe_str(x, "OTHER"))
    out["_size_display"] = out["_size_raw"].apply(lambda x: safe_str(x, "").replace("  ", " ").strip())
    out["_brand_order"] = out["_brand_key"].apply(lambda x: brand_rank_map.get(safe_str(x, ""), 999))
    out["_size_order"] = out["_size_key"].apply(lambda x: size_rank_map.get(safe_str(x, ""), 999))

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
        total_groups = grouped.ngroups
        for idx, ((route_agent, run_name), group_df) in enumerate(grouped, start=1):
            log_info(f"Generating loading sheet {idx}/{total_groups} for route '{route_agent}' run '{run_name}'.")
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
