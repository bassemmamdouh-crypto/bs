import math
import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd


# =============================
# SETTINGS (EDIT THIS SECTION)
# =============================
ORDERS_FILE = r"C:\Users\Adel El orashy\Downloads\Marbah Project\scripts\orders_data.xlsx"
VEHICLES_FILE = r"C:\Users\Adel El orashy\Downloads\Marbah Project\scripts\vehicles_data.xlsx"
OUTPUT_ROOT = r"C:\Users\Adel El orashy\Downloads\Marbah Project\scripts"
KML_FILE = r"C:\Users\Adel El orashy\Downloads\Marbah Project\scripts\Marbah IRAQ Map.kml"


@dataclass
class RunSheetConfig:
    orders_file: str = ORDERS_FILE
    vehicles_file: str = VEHICLES_FILE
    output_root: str = OUTPUT_ROOT
    kml_file: str = KML_FILE

    # Order columns
    order_id_candidates: Tuple[str, ...] = ("sales_order_id", "order_id", "order_number", "order_no", "orderid")
    supply_chain_candidates: Tuple[str, ...] = ("supply_chain", "supply chain", "supply_chain_name", "supplychain")
    warehouse_candidates: Tuple[str, ...] = ("warehouse_id", "warehouse", "warehouse_name", "wh_id", "depot")
    segment_candidates: Tuple[str, ...] = ("segment", "retailer_segment", "customer_segment", "channel", "trade_channel")
    route_candidates: Tuple[str, ...] = ("polygon_name", "route", "route_name", "delivery_route", "route_text", "district_name")
    latitude_candidates: Tuple[str, ...] = ("retailer_lat", "latitude", "lat", "customer_lat", "store_lat")
    longitude_candidates: Tuple[str, ...] = ("retailer_long", "longitude", "long", "lng", "customer_long", "store_long")
    # CBM is prioritized for planning load/capacity.
    order_cbm_candidates: Tuple[str, ...] = ("order_cbm", "cbm", "total_cbm", "volume_cbm", "volume")
    carton_cbm_candidates: Tuple[str, ...] = ("carton_cbm", "cbm_per_carton", "carton_volume", "item_cbm")
    order_weight_candidates: Tuple[str, ...] = ("order_weight_kg", "order_weight", "weight", "total_weight", "weight_kg", "kg")
    order_load_candidates: Tuple[str, ...] = ("order_load", "order_items", "items_count", "item_count")
    carton_weight_candidates: Tuple[str, ...] = ("carton_weight", "weight_per_carton", "item_weight", "carton_kg")
    # In this project, order cbm/weight columns are often per-carton metrics.
    order_cbm_is_per_carton: bool = False
    order_weight_is_per_carton: bool = False
    purchased_items_candidates: Tuple[str, ...] = ("purchased_item_count", "qty", "quantity", "item_qty")
    # Fallback only when CBM column is missing/empty.
    order_volume_fallback_candidates: Tuple[str, ...] = (
        "order_volume",
        "vehicle_load",
        "load",
        "total_qty",
        "purchased_item_count",
        "qty",
        "quantity",
        "item_qty",
    )
    delivery_date_candidates: Tuple[str, ...] = ("estimated_delivery_date", "delivery_date")

    # Vehicle columns
    vehicle_id_candidates: Tuple[str, ...] = ("vehicle_id", "truck_id", "car_id", "plate_no", "vehicle")
    vehicle_supply_chain_candidates: Tuple[str, ...] = ("supply_chain", "supply chain", "supply_chain_name", "supplychain")
    vehicle_warehouse_candidates: Tuple[str, ...] = ("warehouse_id", "warehouse", "warehouse_name", "wh_id", "depot")
    vehicle_segment_candidates: Tuple[str, ...] = ("segment", "assigned_segment", "vehicle_segment", "channel")
    vehicle_agent_candidates: Tuple[str, ...] = ("assigned_agent", "agent", "driver", "route_agent", "agent_name")
    vehicle_type_candidates: Tuple[str, ...] = ("vehicle_type", "type", "truck_type")
    vehicle_capacity_candidates: Tuple[str, ...] = (
        "cbm_capcity",
        "cbm_capacity",
        "cbm",
        "cbm_m3",
        "cbm_capacity_m3",
        "vehicle_capacity",
        "capacity_cbm",
        "load_capacity",
        "max_load",
    )
    vehicle_weight_capacity_candidates: Tuple[str, ...] = (
        "weight_kg_capacity",
        "weight_kg_capcity",
        "weight_capacity_kg",
        "weight_capcity_kg",
        "weight_capacitykg",
        "weightkgcapcity",
        "weightkgcapacity",
        "vehicle_weight_capacity",
        "weight_capacity",
        "capacity_weight",
        "max_weight",
        "max_weight_kg",
    )
    vehicle_stop_candidates: Tuple[str, ...] = (
        "number_of_stops",
        "number_of_stops_per_vehicle",
        "stops_per_vehicle",
        "max_stops",
        "stops_capacity",
        "stops_limit",
    )
    vehicle_count_candidates: Tuple[str, ...] = ("vehicle_count", "count", "qty")
    vehicle_active_candidates: Tuple[str, ...] = ("active", "is_active", "enabled")

    # Assignment behavior
    high_volume_min_load_ratio: float = 0.60
    max_stops_per_run: int = 23
    nearest_routes_per_seed: int = 8
    min_utilization_target_pct: float = 98.0
    # Max dispatch waves per vehicle in the same planning cycle.
    # <=0 means unlimited waves until all orders are assigned.
    max_runs_per_vehicle: int = 0
    # Keep run-wave labeling neutral so planners can decide first/second later.
    label_run_wave_as_first_second: bool = False
    # Supply chain capacity-priority mode during packing:
    # - CBM-first means prioritize filling CBM before weight tie-break.
    # - Weight-first means prioritize weight before CBM tie-break.
    cbm_first_supply_chains: Tuple[str, ...] = ("LAYS",)
    weight_first_supply_chains: Tuple[str, ...] = ("PEPSI",)
    # Geographic compactness controls to avoid runsheet overlap.
    allow_topup_distance_relaxation: bool = False
    compactness_iterations: int = 6
    compactness_min_improvement_km: float = 0.05
    # Consolidate sparse runsheets by moving orders into nearby feasible fuller runsheets.
    consolidation_min_utilization_pct: float = 65.0
    consolidation_iterations: int = 4
    # Retailer proximity guardrails (set <=0 to disable a guard).
    max_retailer_distance_to_centroid_km: float = 7.0
    max_retailer_pair_distance_km: float = 10.0
    runsheet_prefix: str = "RS"
    include_inactive_vehicles: bool = False


def log_info(msg: str) -> None:
    print(f"INFO: {msg}")


def log_warning(msg: str) -> None:
    print(f"WARNING: {msg}")


def safe_str(value: object, default: str = "") -> str:
    if pd.isna(value):
        return default
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null", "nat"}:
        return default
    return text


def safe_float(value: object, default: float = 0.0) -> float:
    try:
        if pd.isna(value):
            return default
        text = safe_str(value, "")
        if not text:
            return default
        arabic_digits = "٠١٢٣٤٥٦٧٨٩"
        for i, digit in enumerate(arabic_digits):
            text = text.replace(digit, str(i))
        text = text.replace(",", "")
        text = re.sub(r"[^0-9.\-]", "", text)
        if text in {"", ".", "-", "-.", ".-"}:
            return default
        return float(text)
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


def normalize_key(value: object) -> str:
    return re.sub(r"\s+", " ", safe_str(value, "").strip().upper())


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [safe_str(c, "").lower() for c in df.columns]
    return df


def normalize_column_key(name: object) -> str:
    return re.sub(r"[^a-z0-9]", "", safe_str(name, "").lower())


def find_existing_column(df: pd.DataFrame, candidates: Tuple[str, ...]) -> Optional[str]:
    for candidate in candidates:
        if candidate in df.columns:
            return candidate
    normalized = {normalize_column_key(c): c for c in df.columns}
    for candidate in candidates:
        resolved = normalized.get(normalize_column_key(candidate))
        if resolved:
            return resolved
    for candidate in candidates:
        candidate_key = normalize_column_key(candidate)
        if not candidate_key:
            continue
        for col in df.columns:
            if candidate_key in normalize_column_key(col):
                return col
    return None


def read_tabular_file(path: str) -> pd.DataFrame:
    suffix = Path(path).suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path)
    if suffix in {".csv", ".txt"}:
        return pd.read_csv(path)
    raise ValueError(f"Unsupported file format: {path}")


def bool_from_value(value: object, default: bool = True) -> bool:
    text = safe_str(value, "").lower()
    if text in {"1", "true", "yes", "y"}:
        return True
    if text in {"0", "false", "no", "n"}:
        return False
    return default


def aggregate_order_measure(series: pd.Series) -> float:
    numeric = pd.to_numeric(series, errors="coerce").fillna(0)
    non_zero = numeric[numeric > 0]
    if non_zero.empty:
        return 0.0
    unique = non_zero.unique()
    # If same non-zero value repeats in each line, treat it as order-level value.
    if len(unique) == 1 and len(non_zero) > 1:
        return float(unique[0])
    return float(non_zero.sum())


def get_order_cbm(order: Dict[str, object]) -> float:
    return max(0.0, safe_float(order.get("order_cbm", order.get("order_load", 0)), 0.0))


def get_order_weight(order: Dict[str, object]) -> float:
    return max(0.0, safe_float(order.get("order_weight", 0), 0.0))


def is_finite_positive(value: float) -> bool:
    return math.isfinite(value) and value > 0


def safe_sum_capacity(values: List[float]) -> float:
    finite_values = [v for v in values if math.isfinite(v)]
    if len(finite_values) != len(values):
        return float("inf")
    return float(sum(finite_values))


def normalize_chain_set(chains: Tuple[str, ...]) -> set:
    return {normalize_key(c) for c in chains if safe_str(c, "")}


def get_supply_chain_priority_mode(supply_chain: str, config: Optional[RunSheetConfig]) -> str:
    if config is None:
        return "BALANCED"
    sc_key = normalize_key(supply_chain)
    if sc_key in normalize_chain_set(config.weight_first_supply_chains):
        return "WEIGHT_FIRST"
    if sc_key in normalize_chain_set(config.cbm_first_supply_chains):
        return "CBM_FIRST"
    return "BALANCED"


def order_capacity_pressure(
    order: Dict[str, object],
    cbm_capacity: float,
    weight_capacity: float,
    supply_chain: str = "",
    config: Optional[RunSheetConfig] = None,
) -> float:
    cbm = get_order_cbm(order)
    weight = get_order_weight(order)
    cbm_ratio = (cbm / cbm_capacity) if cbm_capacity > 0 else 0.0
    weight_ratio = (weight / weight_capacity) if is_finite_positive(weight_capacity) else 0.0
    mode = get_supply_chain_priority_mode(supply_chain, config)
    if mode == "WEIGHT_FIRST":
        return (weight_ratio * 2.0) + cbm_ratio
    if mode == "CBM_FIRST":
        return (cbm_ratio * 2.0) + weight_ratio
    return max(cbm_ratio, weight_ratio)


def format_capacity_value(value: float) -> object:
    return value if math.isfinite(value) else "INF"


def build_run_wave_fields(run_number: int, config: RunSheetConfig) -> Tuple[object, str]:
    if bool(config.label_run_wave_as_first_second):
        is_second = run_number > 1
        return is_second, ("SECOND_RUN" if is_second else "FIRST_RUN")
    return "", "TO_BE_DECIDED"


def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    r = 6371.0
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlon / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def order_within_bin_distance_limits(order: Dict[str, object], bin_obj: "VehicleBin", config: RunSheetConfig) -> bool:
    if not bin_obj.orders:
        return True

    lat = safe_float(order.get("lat", 0), 0.0)
    lon = safe_float(order.get("lon", 0), 0.0)
    if bin_obj.centroid_lat is None or bin_obj.centroid_lon is None:
        return True

    max_centroid = safe_float(config.max_retailer_distance_to_centroid_km, 0.0)
    if max_centroid > 0:
        centroid_dist = haversine_km(lat, lon, bin_obj.centroid_lat, bin_obj.centroid_lon)
        if centroid_dist > max_centroid:
            return False

    max_pair = safe_float(config.max_retailer_pair_distance_km, 0.0)
    if max_pair > 0:
        for existing in bin_obj.orders:
            ex_lat = safe_float(existing.get("lat", 0), 0.0)
            ex_lon = safe_float(existing.get("lon", 0), 0.0)
            d = haversine_km(lat, lon, ex_lat, ex_lon)
            if d > max_pair:
                return False

    return True


def distance_order_to_bin_centroid_km(order: Dict[str, object], bin_obj: "VehicleBin") -> float:
    if bin_obj.centroid_lat is None or bin_obj.centroid_lon is None:
        return 0.0
    lat = safe_float(order.get("lat", 0), 0.0)
    lon = safe_float(order.get("lon", 0), 0.0)
    return haversine_km(lat, lon, bin_obj.centroid_lat, bin_obj.centroid_lon)


def projected_fill_quality(
    bin_obj: "VehicleBin",
    order: Dict[str, object],
    config: RunSheetConfig,
) -> Tuple[float, float]:
    projected_cbm = bin_obj.assigned_load + get_order_cbm(order)
    projected_weight = bin_obj.assigned_weight + get_order_weight(order)
    projected_util = bin_obj.projected_utilization_pct(projected_cbm, projected_weight)
    target = min(100.0, max(0.0, safe_float(config.min_utilization_target_pct, 0.0)))
    util_gap = abs(target - projected_util)

    cbm_remaining_ratio = (
        max(0.0, bin_obj.capacity - projected_cbm) / bin_obj.capacity if bin_obj.capacity > 0 else 1.0
    )
    if is_finite_positive(bin_obj.weight_capacity):
        weight_remaining_ratio = max(0.0, bin_obj.weight_capacity - projected_weight) / bin_obj.weight_capacity
        bottleneck_remaining_ratio = max(cbm_remaining_ratio, weight_remaining_ratio)
    else:
        bottleneck_remaining_ratio = cbm_remaining_ratio
    return util_gap, bottleneck_remaining_ratio


def compact_bins_for_min_distance(bins: List["VehicleBin"], config: RunSheetConfig) -> None:
    active_bins = [b for b in bins if b.orders]
    if len(active_bins) <= 1:
        return
    stop_cap = max(0, int(config.max_stops_per_run))
    iterations = max(0, int(config.compactness_iterations))
    if iterations <= 0:
        return
    min_gain = max(0.0, safe_float(config.compactness_min_improvement_km, 0.0))
    max_cbm_capacity = max((b.capacity for b in active_bins), default=0.0)
    finite_weight_capacities = [b.weight_capacity for b in active_bins if math.isfinite(b.weight_capacity)]
    max_weight_capacity = max(finite_weight_capacities) if finite_weight_capacities else float("inf")
    chain_for_priority = safe_str(active_bins[0].supply_chain, "")

    def snapshot_state() -> List[List[Dict[str, object]]]:
        return [list(b.orders) for b in active_bins]

    def restore_state(state: List[List[Dict[str, object]]]) -> None:
        for bin_obj, orders in zip(active_bins, state):
            bin_obj.orders = list(orders)
            bin_obj.recompute_state()

    def total_compactness_distance() -> float:
        total = 0.0
        for bin_obj in active_bins:
            for order in bin_obj.orders:
                total += distance_order_to_bin_centroid_km(order, bin_obj)
        return total

    all_orders: List[Dict[str, object]] = []
    seen_ids: set = set()
    for bin_obj in active_bins:
        for order in bin_obj.orders:
            order_id = safe_str(order.get("_order_id", ""), "")
            if order_id and order_id in seen_ids:
                continue
            if order_id:
                seen_ids.add(order_id)
            all_orders.append(order)

    best_state = snapshot_state()
    best_distance = total_compactness_distance()

    for _ in range(iterations):
        seed_centers: List[Tuple[float, float]] = []
        for bin_obj in active_bins:
            if bin_obj.centroid_lat is not None and bin_obj.centroid_lon is not None:
                seed_centers.append((bin_obj.centroid_lat, bin_obj.centroid_lon))
                continue
            if bin_obj.orders:
                first = bin_obj.orders[0]
                seed_centers.append((safe_float(first.get("lat", 0), 0.0), safe_float(first.get("lon", 0), 0.0)))
            else:
                seed_centers.append((0.0, 0.0))

        for bin_obj in active_bins:
            bin_obj.orders = []
            bin_obj.recompute_state()

        unassigned: List[Dict[str, object]] = []
        orders_sorted = sorted(
            all_orders,
            key=lambda o: order_capacity_pressure(
                o,
                max_cbm_capacity,
                max_weight_capacity,
                chain_for_priority,
                config,
            ),
            reverse=True,
        )

        for order in orders_sorted:
            order_cbm = get_order_cbm(order)
            order_weight = get_order_weight(order)
            lat = safe_float(order.get("lat", 0), 0.0)
            lon = safe_float(order.get("lon", 0), 0.0)
            candidates: List[Tuple[int, float, float, int]] = []
            for idx, bin_obj in enumerate(active_bins):
                if not bin_obj.can_fit(order_cbm, order_weight, stop_cap):
                    continue
                if bin_obj.orders:
                    center_lat = safe_float(bin_obj.centroid_lat, 0.0)
                    center_lon = safe_float(bin_obj.centroid_lon, 0.0)
                    within_limits = order_within_bin_distance_limits(order, bin_obj, config)
                else:
                    center_lat, center_lon = seed_centers[idx]
                    within_limits = True
                distance = haversine_km(lat, lon, center_lat, center_lon)
                candidates.append(
                    (
                        0 if within_limits else 1,
                        distance,
                        -bin_obj.utilization_pct,
                        idx,
                    )
                )

            if not candidates:
                unassigned.append(order)
                continue

            candidates.sort(key=lambda t: (t[0], t[1], t[2]))
            best_idx = candidates[0][3]
            active_bins[best_idx].add_order(order)

        if unassigned:
            # Retry remaining with relaxed distance guard but strict capacity/stops.
            retry_unassigned: List[Dict[str, object]] = []
            for order in unassigned:
                order_cbm = get_order_cbm(order)
                order_weight = get_order_weight(order)
                lat = safe_float(order.get("lat", 0), 0.0)
                lon = safe_float(order.get("lon", 0), 0.0)
                candidates: List[Tuple[float, float, int]] = []
                for idx, bin_obj in enumerate(active_bins):
                    if not bin_obj.can_fit(order_cbm, order_weight, stop_cap):
                        continue
                    if bin_obj.centroid_lat is None or bin_obj.centroid_lon is None:
                        center_lat, center_lon = seed_centers[idx]
                    else:
                        center_lat, center_lon = bin_obj.centroid_lat, bin_obj.centroid_lon
                    distance = haversine_km(lat, lon, safe_float(center_lat, 0.0), safe_float(center_lon, 0.0))
                    candidates.append((distance, -bin_obj.utilization_pct, idx))
                if not candidates:
                    retry_unassigned.append(order)
                    continue
                candidates.sort(key=lambda t: (t[0], t[1]))
                active_bins[candidates[0][2]].add_order(order)

            if retry_unassigned:
                # Keep previous best if reassignment could not place all orders.
                restore_state(best_state)
                break

        current_distance = total_compactness_distance()
        if current_distance + min_gain < best_distance:
            best_distance = current_distance
            best_state = snapshot_state()
            continue

        # No meaningful gain this iteration.
        restore_state(best_state)
        break

    restore_state(best_state)


def effective_stop_cap_for_bin(bin_obj: "VehicleBin", fallback_stop_cap: int) -> int:
    return bin_obj.stop_capacity if bin_obj.stop_capacity > 0 else max(0, int(fallback_stop_cap))


def consolidate_sparse_bins(
    bins: List["VehicleBin"],
    config: RunSheetConfig,
) -> None:
    active_bins = [b for b in bins if b.orders]
    if len(active_bins) <= 1:
        return

    min_util = max(0.0, min(100.0, safe_float(config.consolidation_min_utilization_pct, 65.0)))
    max_iterations = max(0, int(config.consolidation_iterations))
    if max_iterations <= 0:
        return

    fallback_stop_cap = max(0, int(config.max_stops_per_run))

    for _ in range(max_iterations):
        moved_any = False
        donors = [
            b
            for b in active_bins
            if b.orders and (len(b.orders) <= 1 or b.utilization_pct < min_util)
        ]
        donors.sort(key=lambda b: (len(b.orders), b.utilization_pct))
        if not donors:
            break

        for donor in donors:
            if not donor.orders:
                continue
            donor_orders = sorted(
                list(donor.orders),
                key=lambda o: order_capacity_pressure(
                    o,
                    max(donor.capacity, 1e-9),
                    donor.weight_capacity,
                    donor.supply_chain,
                    config,
                ),
            )
            for order in donor_orders:
                order_id = safe_str(order.get("_order_id", ""), "")
                if not order_id:
                    continue
                recipients = [b for b in active_bins if b is not donor and b.orders]
                best_recipient: Optional[Tuple[float, "VehicleBin"]] = None
                for recipient in recipients:
                    local_stop_cap = effective_stop_cap_for_bin(recipient, fallback_stop_cap)
                    order_cbm = get_order_cbm(order)
                    order_weight = get_order_weight(order)
                    if not recipient.can_fit(order_cbm, order_weight, local_stop_cap):
                        continue
                    if not order_within_bin_distance_limits(order, recipient, config):
                        continue
                    util_gap, bottleneck_remaining = projected_fill_quality(recipient, order, config)
                    distance = distance_order_to_bin_centroid_km(order, recipient)
                    score = (util_gap * 1.0) + (bottleneck_remaining * 0.6) + (distance * 0.4)
                    rank = (score, recipient)
                    if best_recipient is None or rank[0] < best_recipient[0]:
                        best_recipient = rank
                if best_recipient is None:
                    continue
                removed = donor.remove_order_by_id(order_id)
                if removed is None:
                    continue
                best_recipient[1].add_order(removed)
                moved_any = True

        if not moved_any:
            break
        active_bins = [b for b in bins if b.orders]
        if len(active_bins) <= 1:
            break


def enforce_hard_limits_on_bins(
    bins: List["VehicleBin"],
    config: RunSheetConfig,
) -> List[Dict[str, object]]:
    overflow_orders: List[Dict[str, object]] = []
    tolerance = 1e-9

    for bin_obj in bins:
        if not bin_obj.orders:
            continue
        stop_cap = bin_obj.stop_capacity if bin_obj.stop_capacity > 0 else max(0, int(config.max_stops_per_run))

        while True:
            over_cbm = bin_obj.remaining < -tolerance
            over_weight = is_finite_positive(bin_obj.weight_capacity) and (bin_obj.weight_remaining < -tolerance)
            over_stops = stop_cap > 0 and len(bin_obj.orders) > stop_cap
            if not (over_cbm or over_weight or over_stops):
                break

            # Remove the order with the weakest geographical fit / highest capacity burden.
            scored: List[Tuple[float, int]] = []
            for idx, order in enumerate(bin_obj.orders):
                distance = distance_order_to_bin_centroid_km(order, bin_obj)
                pressure = order_capacity_pressure(
                    order,
                    max(bin_obj.capacity, 1e-9),
                    bin_obj.weight_capacity,
                    bin_obj.supply_chain,
                    config,
                )
                score = (distance * 3.0) + pressure
                scored.append((score, idx))
            if not scored:
                break
            scored.sort(key=lambda t: t[0], reverse=True)
            drop_idx = scored[0][1]
            removed = bin_obj.orders.pop(drop_idx)
            bin_obj.recompute_state()
            overflow_orders.append(removed)

    return overflow_orders


def pick_topup_candidate(
    bin_obj: "VehicleBin",
    orders_pool: List[Dict[str, object]],
    assigned_ids: set,
    target_utilization_pct: float,
    seed_route: str,
    route_neighbors: Dict[str, List[str]],
    stop_cap: int,
    config: RunSheetConfig,
    enforce_distance: bool,
) -> Optional[Dict[str, object]]:
    best: Optional[Tuple[float, float, Dict[str, object]]] = None
    for order in orders_pool:
        order_id = safe_str(order.get("_order_id", ""), "")
        if not order_id or order_id in assigned_ids:
            continue
        order_cbm = get_order_cbm(order)
        order_weight = get_order_weight(order)
        if not bin_obj.can_fit(order_cbm, order_weight, stop_cap):
            continue
        if enforce_distance and (not order_within_bin_distance_limits(order, bin_obj, config)):
            continue
        projected_cbm = bin_obj.assigned_load + order_cbm
        projected_weight = bin_obj.assigned_weight + order_weight
        projected_util = bin_obj.projected_utilization_pct(projected_cbm, projected_weight)
        gap_to_target = abs(max(0.0, target_utilization_pct - projected_util))

        remaining_cbm_ratio = 0.0
        if bin_obj.capacity > 0:
            remaining_cbm_ratio = max(0.0, bin_obj.capacity - projected_cbm) / bin_obj.capacity
        remaining_weight_ratio = 0.0
        if is_finite_positive(bin_obj.weight_capacity):
            remaining_weight_ratio = max(0.0, bin_obj.weight_capacity - projected_weight) / bin_obj.weight_capacity
        remaining_after = remaining_cbm_ratio + remaining_weight_ratio

        lat = safe_float(order.get("lat", 0), 0.0)
        lon = safe_float(order.get("lon", 0), 0.0)
        if bin_obj.centroid_lat is None or bin_obj.centroid_lon is None:
            distance = 0.0
        else:
            distance = haversine_km(lat, lon, bin_obj.centroid_lat, bin_obj.centroid_lon)
        route_penalty = route_proximity_penalty(safe_str(order.get("route", ""), ""), seed_route, route_neighbors)
        score = gap_to_target + (remaining_after * 0.2) + (distance * 0.05) + (route_penalty * 0.1)
        rank = (
            score,
            -order_capacity_pressure(
                order,
                bin_obj.capacity,
                bin_obj.weight_capacity,
                bin_obj.supply_chain,
                config,
            ),
            order,
        )
        if best is None or rank[0] < best[0] or (rank[0] == best[0] and rank[1] < best[1]):
            best = rank
    return best[2] if best is not None else None


def top_up_bin_to_target_utilization(
    bin_obj: "VehicleBin",
    orders_pool: List[Dict[str, object]],
    assigned_ids: set,
    seed_route: str,
    route_neighbors: Dict[str, List[str]],
    stop_cap: int,
    config: RunSheetConfig,
) -> None:
    target_pct = safe_float(config.min_utilization_target_pct, 0.0)
    if target_pct <= 0 or bin_obj.capacity <= 0:
        return
    target_pct = min(target_pct, 100.0)
    if bin_obj.utilization_pct >= target_pct:
        return
    local_stop_cap = effective_stop_cap_for_bin(bin_obj, stop_cap)

    # First pass: keep strict distance guards.
    while bin_obj.utilization_pct < target_pct:
        if local_stop_cap > 0 and len(bin_obj.orders) >= local_stop_cap:
            break
        candidate = pick_topup_candidate(
            bin_obj,
            orders_pool,
            assigned_ids,
            target_pct,
            seed_route,
            route_neighbors,
            local_stop_cap,
            config,
            enforce_distance=True,
        )
        if candidate is None:
            break
        candidate_id = safe_str(candidate.get("_order_id", ""), "")
        if not candidate_id:
            break
        bin_obj.add_order(candidate)
        assigned_ids.add(candidate_id)

    if not bool(config.allow_topup_distance_relaxation):
        return

    # Second pass: relax distance only if still below target.
    while bin_obj.utilization_pct < target_pct:
        if local_stop_cap > 0 and len(bin_obj.orders) >= local_stop_cap:
            break
        candidate = pick_topup_candidate(
            bin_obj,
            orders_pool,
            assigned_ids,
            target_pct,
            seed_route,
            route_neighbors,
            local_stop_cap,
            config,
            enforce_distance=False,
        )
        if candidate is None:
            break
        candidate_id = safe_str(candidate.get("_order_id", ""), "")
        if not candidate_id:
            break
        bin_obj.add_order(candidate)
        assigned_ids.add(candidate_id)


def route_proximity_penalty(route_name: str, seed_route: str, route_neighbors: Dict[str, List[str]]) -> float:
    if not seed_route:
        return 0.0
    if route_name == seed_route:
        return 0.0
    neighbors = route_neighbors.get(seed_route, [])
    if route_name in neighbors:
        return (neighbors.index(route_name) + 1) * 0.25
    return 3.0


def parse_kml_route_centroids(kml_path: str) -> Dict[str, Tuple[float, float]]:
    if not kml_path:
        return {}
    if not os.path.exists(kml_path):
        log_warning(f"KML file not found, falling back to order coordinates: {kml_path}")
        return {}
    try:
        tree = ET.parse(kml_path)
        root = tree.getroot()
    except Exception as exc:
        log_warning(f"Unable to parse KML file, fallback to order coordinates: {exc}")
        return {}

    centroids: Dict[str, Tuple[float, float]] = {}
    placemarks = root.findall(".//{*}Placemark")
    for placemark in placemarks:
        name_node = placemark.find(".//{*}name")
        raw_name = name_node.text if name_node is not None else ""
        route_key = normalize_key(raw_name)
        if not route_key:
            continue
        coords_node = placemark.find(".//{*}coordinates")
        if coords_node is None or not safe_str(coords_node.text, ""):
            continue
        coords_text = safe_str(coords_node.text, "")
        lat_lon_points: List[Tuple[float, float]] = []
        for token in re.split(r"\s+", coords_text.strip()):
            if not token:
                continue
            parts = token.split(",")
            if len(parts) < 2:
                continue
            lon = safe_float(parts[0], 0.0)
            lat = safe_float(parts[1], 0.0)
            lat_lon_points.append((lat, lon))
        if not lat_lon_points:
            continue
        centroid_lat = sum(p[0] for p in lat_lon_points) / len(lat_lon_points)
        centroid_lon = sum(p[1] for p in lat_lon_points) / len(lat_lon_points)
        centroids[route_key] = (centroid_lat, centroid_lon)
    if centroids:
        log_info(f"Loaded {len(centroids)} route centroids from KML.")
    return centroids


def build_route_centers(
    scoped_orders: pd.DataFrame,
    kml_centroids: Dict[str, Tuple[float, float]],
) -> Dict[str, Tuple[float, float]]:
    centers: Dict[str, Tuple[float, float]] = {}
    for route, route_df in scoped_orders.groupby("route", sort=False):
        route_key = normalize_key(route)
        if route_key in kml_centroids:
            centers[route_key] = kml_centroids[route_key]
            continue
        lat = safe_float(route_df["lat"].mean(), 0.0)
        lon = safe_float(route_df["lon"].mean(), 0.0)
        centers[route_key] = (lat, lon)
    return centers


def build_route_neighbors(
    route_centers: Dict[str, Tuple[float, float]],
    nearest_routes_per_seed: int,
) -> Dict[str, List[str]]:
    neighbors: Dict[str, List[str]] = {}
    routes = list(route_centers.keys())
    for route in routes:
        lat1, lon1 = route_centers[route]
        scored: List[Tuple[float, str]] = []
        for other in routes:
            if other == route:
                continue
            lat2, lon2 = route_centers[other]
            scored.append((haversine_km(lat1, lon1, lat2, lon2), other))
        scored.sort(key=lambda t: t[0])
        limit = max(0, nearest_routes_per_seed)
        neighbors[route] = [x[1] for x in (scored[:limit] if limit else scored)]
    return neighbors


class VehicleBin:
    def __init__(
        self,
        vehicle_id: str,
        supply_chain: str,
        warehouse_id: str,
        segment: str,
        assigned_agent: str,
        vehicle_type: str,
        route: str,
        run_number: int,
        capacity: float,
        weight_capacity: float,
        stop_capacity: int,
        runsheet_id: str,
    ):
        self.vehicle_id = vehicle_id
        self.supply_chain = supply_chain
        self.warehouse_id = warehouse_id
        self.segment = segment
        self.assigned_agent = assigned_agent
        self.vehicle_type = vehicle_type
        self.route = route
        self.covered_routes: List[str] = []
        self.run_number = run_number
        self.capacity = max(0.0, capacity)
        self.remaining = max(0.0, capacity)
        self.weight_capacity = weight_capacity if is_finite_positive(weight_capacity) else float("inf")
        self.weight_remaining = self.weight_capacity
        self.stop_capacity = max(0, int(stop_capacity))
        self.runsheet_id = runsheet_id
        self.orders: List[Dict[str, object]] = []
        self.centroid_lat: Optional[float] = None
        self.centroid_lon: Optional[float] = None

    def can_fit(self, cbm_load: float, weight_load: float, max_stops_per_run: int = 0) -> bool:
        stop_limit = self.stop_capacity if self.stop_capacity > 0 else max(0, int(max_stops_per_run))
        if stop_limit > 0 and len(self.orders) >= stop_limit:
            return False
        if self.remaining < cbm_load:
            return False
        if is_finite_positive(self.weight_capacity) and self.weight_remaining < weight_load:
            return False
        return True

    def add_order(self, order: Dict[str, object]) -> None:
        load = get_order_cbm(order)
        load_weight = get_order_weight(order)
        self.orders.append(order)
        self.remaining -= load
        if is_finite_positive(self.weight_capacity):
            self.weight_remaining -= load_weight
        order_route = safe_str(order.get("route", ""), "")
        if order_route and order_route not in self.covered_routes:
            self.covered_routes.append(order_route)
        if not self.route and order_route:
            self.route = order_route
        lat = safe_float(order.get("lat", 0), 0.0)
        lon = safe_float(order.get("lon", 0), 0.0)
        if self.centroid_lat is None or self.centroid_lon is None:
            self.centroid_lat = lat
            self.centroid_lon = lon
            return
        count = len(self.orders)
        self.centroid_lat = ((self.centroid_lat * (count - 1)) + lat) / count
        self.centroid_lon = ((self.centroid_lon * (count - 1)) + lon) / count

    def recompute_state(self) -> None:
        self.remaining = self.capacity
        self.weight_remaining = self.weight_capacity
        self.covered_routes = []
        self.route = ""
        self.centroid_lat = None
        self.centroid_lon = None
        for idx, order in enumerate(self.orders, start=1):
            load = get_order_cbm(order)
            load_weight = get_order_weight(order)
            self.remaining -= load
            if is_finite_positive(self.weight_capacity):
                self.weight_remaining -= load_weight

            order_route = safe_str(order.get("route", ""), "")
            if order_route and order_route not in self.covered_routes:
                self.covered_routes.append(order_route)
            if not self.route and order_route:
                self.route = order_route

            lat = safe_float(order.get("lat", 0), 0.0)
            lon = safe_float(order.get("lon", 0), 0.0)
            if self.centroid_lat is None or self.centroid_lon is None:
                self.centroid_lat = lat
                self.centroid_lon = lon
            else:
                self.centroid_lat = ((self.centroid_lat * (idx - 1)) + lat) / idx
                self.centroid_lon = ((self.centroid_lon * (idx - 1)) + lon) / idx

    def remove_order_by_id(self, order_id: str) -> Optional[Dict[str, object]]:
        target = safe_str(order_id, "")
        if not target:
            return None
        for idx, order in enumerate(self.orders):
            if safe_str(order.get("_order_id", ""), "") == target:
                removed = self.orders.pop(idx)
                self.recompute_state()
                return removed
        return None

    @property
    def assigned_load(self) -> float:
        return self.capacity - self.remaining

    @property
    def assigned_weight(self) -> float:
        if not is_finite_positive(self.weight_capacity):
            return sum(get_order_weight(o) for o in self.orders)
        return self.weight_capacity - self.weight_remaining

    @property
    def cbm_utilization_pct(self) -> float:
        if self.capacity <= 0:
            return 0.0
        return (self.assigned_load / self.capacity) * 100.0

    @property
    def weight_utilization_pct(self) -> float:
        if not is_finite_positive(self.weight_capacity):
            return 0.0
        return (self.assigned_weight / self.weight_capacity) * 100.0

    @property
    def utilization_pct(self) -> float:
        # Vehicle is effectively full if either governing constraint gets full.
        return max(self.cbm_utilization_pct, self.weight_utilization_pct)

    def projected_utilization_pct(self, projected_cbm: float, projected_weight: float) -> float:
        cbm_util = (projected_cbm / self.capacity) * 100.0 if self.capacity > 0 else 0.0
        if is_finite_positive(self.weight_capacity):
            weight_util = (projected_weight / self.weight_capacity) * 100.0
        else:
            weight_util = 0.0
        return max(cbm_util, weight_util)


def choose_best_bin(order: Dict[str, object], bins: List[VehicleBin], max_stops_per_run: int = 0) -> Optional[VehicleBin]:
    load = get_order_cbm(order)
    load_weight = get_order_weight(order)
    fit_bins = [b for b in bins if b.can_fit(load, load_weight, max_stops_per_run)]
    if not fit_bins:
        return None

    lat = safe_float(order.get("lat", 0), 0.0)
    lon = safe_float(order.get("lon", 0), 0.0)
    scored: List[Tuple[float, VehicleBin]] = []
    target_util = 98.0
    for b in fit_bins:
        projected_cbm = b.assigned_load + load
        projected_weight = b.assigned_weight + load_weight
        projected_util = b.projected_utilization_pct(projected_cbm, projected_weight)
        util_gap = abs(target_util - projected_util)
        remaining_after_cbm = b.remaining - load
        remaining_after_weight = (b.weight_remaining - load_weight) if is_finite_positive(b.weight_capacity) else float("inf")
        if b.centroid_lat is None or b.centroid_lon is None:
            distance = 0.0
        else:
            distance = haversine_km(lat, lon, b.centroid_lat, b.centroid_lon)
        pressure = order_capacity_pressure(order, b.capacity, b.weight_capacity, b.supply_chain, None)
        score = (
            (util_gap * 0.8)
            + (distance * 0.6)
            + max(0.0, remaining_after_cbm) * 0.05
            + (0.0 if not math.isfinite(remaining_after_weight) else max(0.0, remaining_after_weight) * 0.001)
            - (pressure * 2.0)
        )
        scored.append((score, b))
    scored.sort(key=lambda t: t[0])
    return scored[0][1]


def load_orders(config: RunSheetConfig) -> pd.DataFrame:
    if not os.path.exists(config.orders_file):
        raise FileNotFoundError(f"Orders file not found: {config.orders_file}")
    raw = read_tabular_file(config.orders_file)
    raw = normalize_columns(raw)

    order_id_col = find_existing_column(raw, config.order_id_candidates)
    chain_col = find_existing_column(raw, config.supply_chain_candidates)
    warehouse_col = find_existing_column(raw, config.warehouse_candidates)
    segment_col = find_existing_column(raw, config.segment_candidates)
    route_col = find_existing_column(raw, config.route_candidates)
    lat_col = find_existing_column(raw, config.latitude_candidates)
    lon_col = find_existing_column(raw, config.longitude_candidates)
    cbm_col = find_existing_column(raw, config.order_cbm_candidates)
    carton_cbm_col = find_existing_column(raw, config.carton_cbm_candidates)
    weight_col = find_existing_column(raw, config.order_weight_candidates)
    order_load_col = find_existing_column(raw, config.order_load_candidates)
    carton_weight_col = find_existing_column(raw, config.carton_weight_candidates)
    items_col = find_existing_column(raw, config.purchased_items_candidates)
    load_fallback_col = find_existing_column(raw, config.order_volume_fallback_candidates)
    date_col = find_existing_column(raw, config.delivery_date_candidates)

    required = {
        "order_id": order_id_col,
        "supply_chain": chain_col,
        "warehouse_id": warehouse_col,
        "route": route_col,
        "latitude": lat_col,
        "longitude": lon_col,
    }
    missing = [k for k, v in required.items() if v is None]
    if missing:
        raise ValueError(f"Orders data missing required columns: {', '.join(missing)}")

    orders = raw.copy()
    orders["_order_id"] = orders[order_id_col].apply(lambda x: normalize_identifier(x, ""))
    orders["_supply_chain"] = orders[chain_col].apply(lambda x: normalize_key(x))
    orders["_warehouse_id"] = orders[warehouse_col].apply(lambda x: normalize_identifier(x, ""))
    if segment_col is not None:
        orders["_segment"] = orders[segment_col].apply(lambda x: normalize_key(x))
        orders["_segment"] = orders["_segment"].apply(lambda s: s if s else "GENERAL")
    else:
        orders["_segment"] = "GENERAL"
        log_warning("Order segment column not found. All orders will be treated as segment GENERAL.")
    orders["_route"] = orders[route_col].apply(lambda x: normalize_key(x))
    orders["_lat"] = orders[lat_col].apply(lambda x: safe_float(x, 0.0))
    orders["_lon"] = orders[lon_col].apply(lambda x: safe_float(x, 0.0))
    if order_load_col is not None:
        orders["_line_order_load"] = orders[order_load_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
        log_info(f"Using order load (items count) column: {order_load_col}")
    else:
        orders["_line_order_load"] = 0.0
    if items_col is not None:
        orders["_line_items"] = orders[items_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
    elif order_load_col is not None:
        orders["_line_items"] = orders["_line_order_load"]
        log_info("Purchased-items column not found. Using order_load as purchased_items source.")
    else:
        orders["_line_items"] = 0.0
        log_warning("Purchased-items and order_load columns were not found. Summary purchased items will be zero.")

    if order_load_col is None:
        orders["_line_order_load"] = orders["_line_items"]
        log_warning("Order load column not found. Falling back to purchased-items for order_load.")
    orders["_qty_for_carton"] = orders["_line_items"].where(orders["_line_items"] > 0, orders["_line_order_load"])
    orders["_line_load_count"] = orders["_line_order_load"].where(orders["_line_order_load"] > 0, orders["_line_items"])

    direct_cbm_series = (
        orders[cbm_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
        if cbm_col is not None
        else pd.Series(0.0, index=orders.index)
    )
    carton_cbm_series = (
        orders[carton_cbm_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
        if carton_cbm_col is not None
        else pd.Series(0.0, index=orders.index)
    )
    fallback_cbm_series = (
        orders[load_fallback_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
        if load_fallback_col is not None
        else pd.Series(0.0, index=orders.index)
    )
    direct_cbm_total = (
        (direct_cbm_series * orders["_qty_for_carton"]).where(orders["_qty_for_carton"] > 0, direct_cbm_series)
        if config.order_cbm_is_per_carton
        else direct_cbm_series
    )
    carton_cbm_total = (carton_cbm_series * orders["_qty_for_carton"]).where(
        orders["_qty_for_carton"] > 0,
        carton_cbm_series,
    )
    orders["_line_cbm"] = direct_cbm_total.where(
        direct_cbm_total > 0,
        carton_cbm_total.where(
            (carton_cbm_series > 0) & (orders["_qty_for_carton"] > 0),
            fallback_cbm_series,
        ),
    )
    if cbm_col is not None:
        if config.order_cbm_is_per_carton:
            log_info(f"Using per-carton CBM column and multiplying by purchased items: {cbm_col}")
        else:
            log_info(f"Using order-level CBM column when available: {cbm_col}")
    elif carton_cbm_col is not None:
        log_info(f"Order CBM will be derived from carton CBM x purchased items: {carton_cbm_col}")
    elif load_fallback_col is not None:
        log_warning(f"CBM column not found. Using fallback load column: {load_fallback_col}")
    else:
        orders["_line_cbm"] = 1.0
        log_warning("No CBM/carton/fallback load columns found. Using default load=1 per order line.")

    direct_weight_series = (
        orders[weight_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
        if weight_col is not None
        else pd.Series(0.0, index=orders.index)
    )
    carton_weight_series = (
        orders[carton_weight_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
        if carton_weight_col is not None
        else pd.Series(0.0, index=orders.index)
    )
    direct_weight_total = (
        (direct_weight_series * orders["_qty_for_carton"]).where(orders["_qty_for_carton"] > 0, direct_weight_series)
        if config.order_weight_is_per_carton
        else direct_weight_series
    )
    carton_weight_total = (carton_weight_series * orders["_qty_for_carton"]).where(
        orders["_qty_for_carton"] > 0,
        carton_weight_series,
    )
    orders["_line_weight"] = direct_weight_total.where(
        direct_weight_total > 0,
        carton_weight_total.where(
            (carton_weight_series > 0) & (orders["_qty_for_carton"] > 0),
            0.0,
        ),
    )
    if weight_col is not None:
        if config.order_weight_is_per_carton:
            log_info(f"Using per-carton weight column and multiplying by purchased items: {weight_col}")
        else:
            log_info(f"Using order-level weight column when available: {weight_col}")
    elif carton_weight_col is not None:
        log_info(f"Order weight will be derived from carton weight x purchased items: {carton_weight_col}")
    else:
        log_warning("No order/carton weight columns found. Weight constraints will be non-restrictive on orders.")

    if date_col is not None:
        orders["_delivery_date"] = pd.to_datetime(orders[date_col], errors="coerce")
    else:
        orders["_delivery_date"] = pd.NaT

    orders = orders[
        (orders["_order_id"] != "")
        & (orders["_supply_chain"] != "")
        & (orders["_warehouse_id"] != "")
        & (orders["_route"] != "")
    ]
    grouped = (
        orders.groupby("_order_id", as_index=False)
        .agg(
            supply_chain=("_supply_chain", "first"),
            warehouse_id=("_warehouse_id", "first"),
            segment=("_segment", "first"),
            route=("_route", "first"),
            lat=("_lat", "first"),
            lon=("_lon", "first"),
            # _line_cbm/_line_weight are line totals, so order totals must be summed.
            order_cbm=("_line_cbm", "sum"),
            order_weight=("_line_weight", "sum"),
            order_load=("_line_load_count", "sum"),
            purchased_items=("_line_load_count", "sum"),
            delivery_date=("_delivery_date", "max"),
        )
    )
    grouped["order_cbm"] = grouped["order_cbm"].apply(lambda x: max(0.0, safe_float(x, 0.0)))
    grouped["order_weight"] = grouped["order_weight"].apply(lambda x: max(0.0, safe_float(x, 0.0)))
    grouped["order_load"] = grouped["order_load"].apply(lambda x: max(0.0, safe_float(x, 0.0)))
    grouped["order_load"] = grouped["order_load"].where(grouped["order_load"] > 0, grouped["purchased_items"])
    grouped["order_load"] = grouped["order_load"].where(grouped["order_load"] > 0, grouped["order_cbm"])
    grouped["purchased_items"] = grouped["purchased_items"].apply(lambda x: max(0.0, safe_float(x, 0.0)))
    grouped = grouped[grouped["order_cbm"] > 0]
    if grouped.empty:
        raise ValueError("No valid orders after preprocessing.")
    return grouped


def load_vehicles(config: RunSheetConfig) -> pd.DataFrame:
    if not os.path.exists(config.vehicles_file):
        raise FileNotFoundError(f"Vehicles file not found: {config.vehicles_file}")
    raw = read_tabular_file(config.vehicles_file)
    raw = normalize_columns(raw)

    vehicle_id_col = find_existing_column(raw, config.vehicle_id_candidates)
    chain_col = find_existing_column(raw, config.vehicle_supply_chain_candidates)
    warehouse_col = find_existing_column(raw, config.vehicle_warehouse_candidates)
    segment_col = find_existing_column(raw, config.vehicle_segment_candidates)
    agent_col = find_existing_column(raw, config.vehicle_agent_candidates)
    type_col = find_existing_column(raw, config.vehicle_type_candidates)
    capacity_col = find_existing_column(raw, config.vehicle_capacity_candidates)
    weight_capacity_col = find_existing_column(raw, config.vehicle_weight_capacity_candidates)
    stops_col = find_existing_column(raw, config.vehicle_stop_candidates)
    count_col = find_existing_column(raw, config.vehicle_count_candidates)
    active_col = find_existing_column(raw, config.vehicle_active_candidates)

    if chain_col is None:
        raise ValueError("Vehicles data missing supply_chain column.")
    if warehouse_col is None:
        raise ValueError("Vehicles data missing warehouse column.")
    if capacity_col is None:
        raise ValueError("Vehicles data missing vehicle capacity column.")
    if vehicle_id_col is None and count_col is None:
        raise ValueError("Vehicles data needs vehicle_id or vehicle_count.")

    vehicles = raw.copy()
    vehicles["_supply_chain"] = vehicles[chain_col].apply(lambda x: normalize_key(x))
    vehicles["_warehouse_id"] = vehicles[warehouse_col].apply(lambda x: normalize_identifier(x, ""))
    vehicles["_segment"] = (
        vehicles[segment_col].apply(lambda x: normalize_key(x)).apply(lambda s: s if s else "UNSPECIFIED")
        if segment_col is not None
        else "UNSPECIFIED"
    )
    vehicles["_assigned_agent"] = (
        vehicles[agent_col].apply(lambda x: normalize_identifier(x, ""))
        if agent_col is not None
        else ""
    )
    vehicles["_vehicle_type"] = (
        vehicles[type_col].apply(lambda x: normalize_identifier(x, "GENERIC"))
        if type_col is not None
        else "GENERIC"
    )
    vehicles["_capacity"] = vehicles[capacity_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
    if weight_capacity_col is not None:
        vehicles["_weight_capacity"] = vehicles[weight_capacity_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
        vehicles["_weight_capacity"] = vehicles["_weight_capacity"].apply(
            lambda x: x if x > 0 else float("inf")
        )
    else:
        vehicles["_weight_capacity"] = float("inf")
        log_warning("Vehicle weight capacity column not found. Weight capacity will be treated as unlimited.")
    vehicles["_max_stops"] = (
        vehicles[stops_col].apply(lambda x: int(max(0.0, safe_float(x, 0.0))))
        if stops_col is not None
        else int(max(0, config.max_stops_per_run))
    )
    vehicles = vehicles[
        (vehicles["_supply_chain"] != "")
        & (vehicles["_warehouse_id"] != "")
        & (vehicles["_capacity"] > 0)
    ]

    if active_col is not None and not config.include_inactive_vehicles:
        vehicles = vehicles[vehicles[active_col].apply(lambda x: bool_from_value(x, default=True))]

    rows: List[Dict[str, object]] = []
    used_ids: set = set()

    def unique_vehicle_id(base_id: str) -> str:
        candidate = re.sub(r"[^A-Za-z0-9_]+", "", base_id)[:50] or "VEHICLE"
        if candidate not in used_ids:
            used_ids.add(candidate)
            return candidate
        suffix = 2
        while True:
            with_suffix = f"{candidate[:44]}_{suffix}"
            if with_suffix not in used_ids:
                used_ids.add(with_suffix)
                return with_suffix
            suffix += 1

    for _, row in vehicles.iterrows():
        chain = safe_str(row["_supply_chain"], "")
        warehouse_id = safe_str(row["_warehouse_id"], "")
        assigned_agent = safe_str(row["_assigned_agent"], "")
        segment = safe_str(row["_segment"], "UNSPECIFIED")
        vehicle_type = safe_str(row["_vehicle_type"], "GENERIC") or "GENERIC"
        base_vehicle_id = normalize_identifier(row.get(vehicle_id_col, ""), "") if vehicle_id_col is not None else ""

        if count_col is not None:
            count = int(max(0.0, safe_float(row.get(count_col, 0), 0.0)))
            if count <= 0:
                # If count is missing/zero but a concrete vehicle_id exists, treat it as 1 physical vehicle.
                count = 1 if base_vehicle_id else 0
        else:
            count = 1 if base_vehicle_id else 0

        for idx in range(1, count + 1):
            if base_vehicle_id:
                vehicle_id = base_vehicle_id if count == 1 else f"{base_vehicle_id}_{idx}"
            else:
                vehicle_id = f"{chain}_{warehouse_id}_{vehicle_type}_V{idx}"
            rows.append(
                {
                    "vehicle_id": unique_vehicle_id(vehicle_id),
                    "supply_chain": chain,
                    "warehouse_id": warehouse_id,
                    "segment": segment,
                    "assigned_agent": assigned_agent,
                    "vehicle_type": vehicle_type,
                    "capacity": row["_capacity"],
                    "weight_capacity": row["_weight_capacity"],
                    "max_stops": int(row["_max_stops"]) if safe_float(row["_max_stops"], 0.0) > 0 else int(config.max_stops_per_run),
                }
            )
    out = pd.DataFrame(rows)
    if out.empty:
        raise ValueError("No active vehicles available after preprocessing.")
    total_cbm_capacity = float(out["capacity"].sum()) if "capacity" in out.columns else 0.0
    finite_weight_caps = [safe_float(x, float("inf")) for x in out.get("weight_capacity", pd.Series(dtype=float)).tolist()]
    total_weight_capacity = safe_sum_capacity(finite_weight_caps) if finite_weight_caps else float("inf")
    total_stops_capacity = int(out["max_stops"].sum()) if "max_stops" in out.columns else 0
    log_info(
        "Vehicle capacity on hand (count-adjusted): "
        f"vehicles={len(out)}, cbm={round(total_cbm_capacity, 3)}, "
        f"weight={format_capacity_value(total_weight_capacity)}, stops={total_stops_capacity}"
    )
    return out


def build_runsheet_id(
    prefix: str,
    supply_chain: str,
    warehouse_id: str,
    run_number: int,
    vehicle_id: str,
) -> str:
    chain_clean = re.sub(r"[^A-Z0-9]+", "", normalize_key(supply_chain)) or "SC"
    warehouse_clean = re.sub(r"[^A-Za-z0-9]+", "", safe_str(warehouse_id, "W"))
    vehicle_clean = re.sub(r"[^A-Za-z0-9]+", "", safe_str(vehicle_id, "V"))
    return f"{prefix}-{chain_clean}-{warehouse_clean}-R{run_number:02d}-{vehicle_clean}"[:60]


def create_vehicle_bins(
    vehicle_list: List[Dict[str, object]],
    supply_chain: str,
    warehouse_id: str,
    segment: str,
    run_number: int,
    config: RunSheetConfig,
) -> List[VehicleBin]:
    bins: List[VehicleBin] = []
    for v in vehicle_list:
        bins.append(
            VehicleBin(
                vehicle_id=safe_str(v["vehicle_id"], ""),
                supply_chain=supply_chain,
                warehouse_id=safe_str(warehouse_id, ""),
                segment=segment,
                assigned_agent=safe_str(v.get("assigned_agent", ""), ""),
                vehicle_type=safe_str(v.get("vehicle_type", "GENERIC"), "GENERIC"),
                route="",
                run_number=run_number,
                capacity=safe_float(v["capacity"], 0.0),
                weight_capacity=safe_float(v.get("weight_capacity", float("inf")), float("inf")),
                stop_capacity=int(max(0.0, safe_float(v.get("max_stops", config.max_stops_per_run), config.max_stops_per_run))),
                runsheet_id=build_runsheet_id(
                    config.runsheet_prefix,
                    supply_chain,
                    safe_str(warehouse_id, ""),
                    run_number,
                    safe_str(v["vehicle_id"], ""),
                ),
            )
        )
    return bins


def allocate_vehicles_to_segments(
    scoped_vehicles: pd.DataFrame,
    scoped_orders: pd.DataFrame,
) -> Dict[str, pd.DataFrame]:
    segments = [s for s in scoped_orders["segment"].dropna().unique().tolist() if safe_str(s, "")]
    if not segments:
        return {}

    segment_vehicle_rows: Dict[str, List[Dict[str, object]]] = {safe_str(s, ""): [] for s in segments}
    if scoped_vehicles.empty:
        return {seg: pd.DataFrame(columns=scoped_vehicles.columns) for seg in segment_vehicle_rows.keys()}

    working = scoped_vehicles.copy()
    if "segment" not in working.columns:
        working["segment"] = "UNSPECIFIED"

    # Respect explicit vehicle-to-segment assignment when available.
    explicit_mask = working["segment"].apply(lambda s: safe_str(s, "") not in {"", "UNSPECIFIED"})
    explicit = working[explicit_mask]
    unassigned = working[~explicit_mask].sort_values("vehicle_id")

    for segment in segments:
        segment_key = safe_str(segment, "")
        if not segment_key:
            continue
        seg_rows = explicit[explicit["segment"] == segment_key].to_dict("records")
        segment_vehicle_rows[segment_key].extend(seg_rows)

    if unassigned.empty:
        return {seg: pd.DataFrame(rows) for seg, rows in segment_vehicle_rows.items()}

    # Auto-split remaining vehicles by demand share so segments are dispatched separately.
    demand_by_segment = (
        scoped_orders.groupby("segment", sort=False)["order_cbm"].sum().to_dict()
        if "order_cbm" in scoped_orders.columns
        else {seg: 1.0 for seg in segments}
    )
    total_demand = sum(max(0.0, safe_float(demand_by_segment.get(seg, 0.0), 0.0)) for seg in segments)
    total_vehicles = len(unassigned)
    alloc_count: Dict[str, int] = {seg: 0 for seg in segments}
    if total_vehicles <= 0:
        return {seg: pd.DataFrame(rows) for seg, rows in segment_vehicle_rows.items()}

    if total_vehicles >= len(segments):
        for seg in segments:
            alloc_count[seg] = 1
        remaining_slots = total_vehicles - len(segments)
    else:
        # Fewer vehicles than segments: prioritize highest-demand segments.
        ranked = sorted(
            segments,
            key=lambda seg: safe_float(demand_by_segment.get(seg, 0.0), 0.0),
            reverse=True,
        )
        for seg in ranked[:total_vehicles]:
            alloc_count[seg] = 1
        remaining_slots = 0

    if remaining_slots > 0:
        if total_demand <= 0:
            ranked = sorted(segments)
            for i in range(remaining_slots):
                alloc_count[ranked[i % len(ranked)]] += 1
        else:
            quotas = {}
            base_sum = 0
            for seg in segments:
                share = max(0.0, safe_float(demand_by_segment.get(seg, 0.0), 0.0)) / total_demand
                raw = share * remaining_slots
                base = int(math.floor(raw))
                quotas[seg] = (raw - base)
                alloc_count[seg] += base
                base_sum += base
            leftovers = remaining_slots - base_sum
            ranked_frac = sorted(segments, key=lambda seg: quotas.get(seg, 0.0), reverse=True)
            for seg in ranked_frac[:leftovers]:
                alloc_count[seg] += 1

    unassigned_rows = unassigned.to_dict("records")
    cursor = 0
    for seg in sorted(segments, key=lambda s: safe_float(demand_by_segment.get(s, 0.0), 0.0), reverse=True):
        take = alloc_count.get(seg, 0)
        for _ in range(take):
            if cursor >= len(unassigned_rows):
                break
            row = dict(unassigned_rows[cursor])
            row["segment"] = seg
            segment_vehicle_rows[seg].append(row)
            cursor += 1

    return {seg: pd.DataFrame(rows) for seg, rows in segment_vehicle_rows.items()}


def assign_single_first_run_route_first(
    first_run_orders: List[Dict[str, object]],
    bins: List[VehicleBin],
    route_neighbors: Dict[str, List[str]],
    config: RunSheetConfig,
) -> set:
    assigned_ids: set = set()
    stop_cap = max(0, int(config.max_stops_per_run))

    for bin_obj in bins:
        local_stop_cap = effective_stop_cap_for_bin(bin_obj, stop_cap)
        remaining = [
            o
            for o in first_run_orders
            if safe_str(o.get("_order_id", ""), "") not in assigned_ids
        ]
        if not remaining:
            break

        route_loads: Dict[str, float] = {}
        for order in remaining:
            route_name = safe_str(order.get("route", ""), "")
            route_loads[route_name] = route_loads.get(route_name, 0.0) + get_order_cbm(order)
        if not route_loads:
            continue
        seed_route = max(route_loads.items(), key=lambda t: t[1])[0]
        route_sequence = [seed_route] + [r for r in route_neighbors.get(seed_route, []) if r != seed_route]
        for route_name in sorted(route_loads.keys(), key=lambda r: route_loads[r], reverse=True):
            if route_name not in route_sequence:
                route_sequence.append(route_name)

        for route_name in route_sequence:
            route_orders = [
                o
                for o in first_run_orders
                if safe_str(o.get("route", ""), "") == route_name
                and safe_str(o.get("_order_id", ""), "") not in assigned_ids
            ]
            route_orders.sort(
                key=lambda o: order_capacity_pressure(
                    o,
                    bin_obj.capacity,
                    bin_obj.weight_capacity,
                    bin_obj.supply_chain,
                    config,
                ),
                reverse=True,
            )
            for order in route_orders:
                if local_stop_cap > 0 and len(bin_obj.orders) >= local_stop_cap:
                    break
                load_cbm = get_order_cbm(order)
                load_weight = get_order_weight(order)
                if not bin_obj.can_fit(load_cbm, load_weight, local_stop_cap):
                    continue
                if not order_within_bin_distance_limits(order, bin_obj, config):
                    continue
                bin_obj.add_order(order)
                assigned_ids.add(safe_str(order.get("_order_id", ""), ""))
            if local_stop_cap > 0 and len(bin_obj.orders) >= local_stop_cap:
                break

        # Final fill for this first-run vehicle: nearest feasible orders if capacity still exists.
        while True:
            if local_stop_cap > 0 and len(bin_obj.orders) >= local_stop_cap:
                break
            candidates: List[Tuple[float, float, float, float, Dict[str, object]]] = []
            for order in first_run_orders:
                order_id = safe_str(order.get("_order_id", ""), "")
                if not order_id or order_id in assigned_ids:
                    continue
                load_cbm = get_order_cbm(order)
                load_weight = get_order_weight(order)
                if not bin_obj.can_fit(load_cbm, load_weight, local_stop_cap):
                    continue
                if not order_within_bin_distance_limits(order, bin_obj, config):
                    continue
                lat = safe_float(order.get("lat", 0), 0.0)
                lon = safe_float(order.get("lon", 0), 0.0)
                dist = (
                    0.0
                    if bin_obj.centroid_lat is None or bin_obj.centroid_lon is None
                    else haversine_km(lat, lon, bin_obj.centroid_lat, bin_obj.centroid_lon)
                )
                util_gap, bottleneck_remaining = projected_fill_quality(bin_obj, order, config)
                candidates.append(
                    (
                        util_gap,
                        bottleneck_remaining,
                        dist,
                        -order_capacity_pressure(
                            order,
                            bin_obj.capacity,
                            bin_obj.weight_capacity,
                            bin_obj.supply_chain,
                            config,
                        ),
                        order,
                    )
                )
            if not candidates:
                break
            candidates.sort(key=lambda t: (t[0], t[1], t[2], t[3]))
            chosen = candidates[0][4]
            chosen_id = safe_str(chosen.get("_order_id", ""), "")
            if not chosen_id:
                break
            bin_obj.add_order(chosen)
            assigned_ids.add(chosen_id)

        # Capacity-utilization top-up toward target (98% by default).
        top_up_bin_to_target_utilization(
            bin_obj,
            first_run_orders,
            assigned_ids,
            seed_route,
            route_neighbors,
            local_stop_cap,
            config,
        )

    compact_bins_for_min_distance(bins, config)
    return assigned_ids


def assign_multi_runs_nearest(
    orders_pool: List[Dict[str, object]],
    vehicle_list: List[Dict[str, object]],
    supply_chain: str,
    warehouse_id: str,
    segment: str,
    start_run_number: int,
    route_neighbors: Dict[str, List[str]],
    config: RunSheetConfig,
) -> Tuple[List[VehicleBin], List[Dict[str, object]], List[Dict[str, object]]]:
    if not orders_pool:
        return [], [], []

    max_vehicle_capacity = max(safe_float(v.get("capacity", 0), 0.0) for v in vehicle_list) if vehicle_list else 0.0
    max_vehicle_weight_capacity = max(
        [
            safe_float(v.get("weight_capacity", float("inf")), float("inf"))
            for v in vehicle_list
            if math.isfinite(safe_float(v.get("weight_capacity", float("inf")), float("inf")))
        ]
        or [float("inf")]
    )
    remaining = list(orders_pool)
    created_bins: List[VehicleBin] = []
    capacity_unassigned: List[Dict[str, object]] = []
    run_number = start_run_number
    stop_cap = max(0, int(config.max_stops_per_run))
    safety_counter = 0
    configured_max_runs = int(config.max_runs_per_vehicle)
    max_run_number = configured_max_runs if configured_max_runs > 0 else 999

    while remaining and run_number <= max_run_number:
        safety_counter += 1
        if safety_counter > 1000:
            log_warning(f"Safety break triggered for extra runs in {supply_chain}/{warehouse_id}.")
            break

        bins = create_vehicle_bins(vehicle_list, supply_chain, warehouse_id, segment, run_number, config)
        assigned_ids: set = set()
        for bin_obj in bins:
            local_stop_cap = effective_stop_cap_for_bin(bin_obj, stop_cap)
            if not remaining:
                break
            seed = max(
                remaining,
                key=lambda o: order_capacity_pressure(
                    o,
                    bin_obj.capacity,
                    bin_obj.weight_capacity,
                    bin_obj.supply_chain,
                    config,
                ),
            )
            seed_id = safe_str(seed.get("_order_id", ""), "")
            seed_load_cbm = get_order_cbm(seed)
            seed_load_weight = get_order_weight(seed)
            if seed_id and bin_obj.can_fit(seed_load_cbm, seed_load_weight, local_stop_cap):
                bin_obj.add_order(seed)
                assigned_ids.add(seed_id)
            seed_route = safe_str(seed.get("route", ""), "")

            while True:
                if local_stop_cap > 0 and len(bin_obj.orders) >= local_stop_cap:
                    break
                if not bin_obj.orders:
                    break
                candidates: List[Tuple[float, float, float, float, Dict[str, object]]] = []
                for order in remaining:
                    order_id = safe_str(order.get("_order_id", ""), "")
                    if not order_id or order_id in assigned_ids:
                        continue
                    load_cbm = get_order_cbm(order)
                    load_weight = get_order_weight(order)
                    if not bin_obj.can_fit(load_cbm, load_weight, local_stop_cap):
                        continue
                    if not order_within_bin_distance_limits(order, bin_obj, config):
                        continue
                    lat = safe_float(order.get("lat", 0), 0.0)
                    lon = safe_float(order.get("lon", 0), 0.0)
                    distance = haversine_km(lat, lon, safe_float(bin_obj.centroid_lat, 0.0), safe_float(bin_obj.centroid_lon, 0.0))
                    route_penalty = route_proximity_penalty(safe_str(order.get("route", ""), ""), seed_route, route_neighbors)
                    util_gap, bottleneck_remaining = projected_fill_quality(bin_obj, order, config)
                    score = distance + route_penalty
                    candidates.append(
                        (
                            util_gap,
                            bottleneck_remaining,
                            score,
                            -order_capacity_pressure(
                                order,
                                bin_obj.capacity,
                                bin_obj.weight_capacity,
                                bin_obj.supply_chain,
                                config,
                            ),
                            order,
                        )
                    )
                if not candidates:
                    break
                candidates.sort(key=lambda t: (t[0], t[1], t[2], t[3]))
                chosen = candidates[0][4]
                chosen_id = safe_str(chosen.get("_order_id", ""), "")
                if not chosen_id:
                    break
                bin_obj.add_order(chosen)
                assigned_ids.add(chosen_id)

            top_up_bin_to_target_utilization(
                bin_obj,
                remaining,
                assigned_ids,
                seed_route,
                route_neighbors,
                local_stop_cap,
                config,
            )

        if not assigned_ids:
            # Capacity-first fallback (still strict capacity).
            capacity_sorted = sorted(
                remaining,
                key=lambda o: order_capacity_pressure(
                    o,
                    max_vehicle_capacity,
                    max_vehicle_weight_capacity,
                    supply_chain,
                    config,
                ),
                reverse=True,
            )
            for order in capacity_sorted:
                order_id = safe_str(order.get("_order_id", ""), "")
                if not order_id or order_id in assigned_ids:
                    continue
                load_cbm = get_order_cbm(order)
                load_weight = get_order_weight(order)
                for bin_obj in bins:
                    local_stop_cap = effective_stop_cap_for_bin(bin_obj, stop_cap)
                    if bin_obj.can_fit(load_cbm, load_weight, local_stop_cap):
                        bin_obj.add_order(order)
                        assigned_ids.add(order_id)
                        break

        if not assigned_ids:
            oversized_ids = set()
            for order in remaining:
                load_cbm = get_order_cbm(order)
                load_weight = get_order_weight(order)
                order_id = safe_str(order.get("_order_id", ""), "")
                too_heavy = is_finite_positive(max_vehicle_weight_capacity) and load_weight > max_vehicle_weight_capacity
                if (load_cbm > max_vehicle_capacity or too_heavy) and order_id:
                    oversized_ids.add(order_id)
                    tagged = dict(order)
                    tagged["unassigned_reason"] = "order_cbm_or_weight_exceeds_max_vehicle_capacity"
                    capacity_unassigned.append(tagged)
            if oversized_ids:
                remaining = [o for o in remaining if safe_str(o.get("_order_id", ""), "") not in oversized_ids]
                continue
            run_number += 1
            continue

        compact_bins_for_min_distance(bins, config)
        overflow_orders = enforce_hard_limits_on_bins(bins, config)
        overflow_ids = {safe_str(o.get("_order_id", ""), "") for o in overflow_orders}
        created_bins.extend([b for b in bins if b.orders])
        remaining = [
            o
            for o in remaining
            if safe_str(o.get("_order_id", ""), "") not in assigned_ids or safe_str(o.get("_order_id", ""), "") in overflow_ids
        ]
        run_number += 1

    if remaining and run_number > max_run_number:
        for order in remaining:
            tagged = dict(order)
            tagged["unassigned_reason"] = "max_runs_per_vehicle_reached"
            capacity_unassigned.append(tagged)
        return created_bins, [], capacity_unassigned

    return created_bins, remaining, capacity_unassigned


def assign_scoped_orders(
    group_orders: pd.DataFrame,
    scoped_vehicles: pd.DataFrame,
    supply_chain: str,
    warehouse_id: str,
    segment: str,
    config: RunSheetConfig,
    route_neighbors: Dict[str, List[str]],
) -> Tuple[List[VehicleBin], List[Dict[str, object]], List[Dict[str, object]]]:
    if scoped_vehicles.empty:
        unassigned = group_orders.to_dict("records")
        return [], unassigned, []

    vehicle_list = scoped_vehicles.sort_values("vehicle_id").to_dict("records")
    all_orders = group_orders.to_dict("records")

    created_bins: List[VehicleBin] = []
    capacity_unassigned: List[Dict[str, object]] = []

    # Assign all orders without pre-labeling second-run pool.
    first_run_orders = list(all_orders)
    first_run_bins = create_vehicle_bins(vehicle_list, supply_chain, warehouse_id, segment, 1, config)
    first_run_assigned_ids = assign_single_first_run_route_first(first_run_orders, first_run_bins, route_neighbors, config)
    first_run_overflow = enforce_hard_limits_on_bins(first_run_bins, config)
    if first_run_overflow:
        overflow_ids = {safe_str(o.get("_order_id", ""), "") for o in first_run_overflow}
        first_run_assigned_ids = {oid for oid in first_run_assigned_ids if oid and oid not in overflow_ids}
    created_bins.extend([b for b in first_run_bins if b.orders])

    remaining_after_first_wave = [
        o for o in first_run_orders if safe_str(o.get("_order_id", ""), "") not in first_run_assigned_ids
    ]
    remaining_pool_raw = remaining_after_first_wave + first_run_overflow
    seen_remaining: set = set()
    remaining_pool: List[Dict[str, object]] = []
    for order in remaining_pool_raw:
        oid = safe_str(order.get("_order_id", ""), "")
        if not oid or oid in seen_remaining:
            continue
        seen_remaining.add(oid)
        remaining_pool.append(order)

    # Assign remaining orders to additional runsheets with strict capacity controls.
    extra_bins, extra_leftover, extra_unassigned = assign_multi_runs_nearest(
        remaining_pool,
        vehicle_list,
        supply_chain,
        warehouse_id,
        segment,
        2,
        route_neighbors,
        config,
    )
    created_bins.extend(extra_bins)
    capacity_unassigned.extend(extra_unassigned)
    consolidate_sparse_bins(created_bins, config)

    return created_bins, extra_leftover, capacity_unassigned


def build_runsheets(
    orders: pd.DataFrame,
    vehicles: pd.DataFrame,
    config: RunSheetConfig,
    kml_centroids: Dict[str, Tuple[float, float]],
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    assignment_rows: List[Dict[str, object]] = []
    summary_rows: List[Dict[str, object]] = []
    unassigned_rows: List[Dict[str, object]] = []
    capacity_plan_rows: List[Dict[str, object]] = []

    vehicles_by_chain_warehouse = {
        (chain, warehouse): group.copy()
        for (chain, warehouse), group in vehicles.groupby(["supply_chain", "warehouse_id"], sort=False)
    }

    for (supply_chain, warehouse_id), scoped_orders in orders.groupby(
        ["supply_chain", "warehouse_id"],
        sort=True,
    ):
        scoped_vehicles = vehicles_by_chain_warehouse.get((supply_chain, warehouse_id), pd.DataFrame())
        warehouse_total_vehicles_on_hand = int(len(scoped_vehicles))
        warehouse_total_cbm_capacity_on_hand = (
            float(scoped_vehicles["capacity"].sum()) if not scoped_vehicles.empty else 0.0
        )
        warehouse_total_weight_capacity_on_hand = (
            safe_sum_capacity(scoped_vehicles["weight_capacity"].astype(float).tolist())
            if (not scoped_vehicles.empty and "weight_capacity" in scoped_vehicles.columns)
            else float("inf")
        )
        warehouse_total_stops_capacity_on_hand = (
            int(scoped_vehicles["max_stops"].sum())
            if (not scoped_vehicles.empty and "max_stops" in scoped_vehicles.columns)
            else 0
        )
        segment_vehicle_pools = allocate_vehicles_to_segments(scoped_vehicles, scoped_orders)
        for segment, segment_orders in scoped_orders.groupby("segment", sort=True):
            segment_key = safe_str(segment, "") or "GENERAL"
            segment_vehicles = segment_vehicle_pools.get(segment_key, pd.DataFrame(columns=scoped_vehicles.columns))
            route_centers = build_route_centers(segment_orders, kml_centroids)
            route_neighbors = build_route_neighbors(route_centers, config.nearest_routes_per_seed)

            demand = float(segment_orders["order_cbm"].sum())
            weight_demand = float(segment_orders["order_weight"].sum())
            first_run_capacity = float(segment_vehicles["capacity"].sum()) if not segment_vehicles.empty else 0.0
            first_run_weight_capacity = (
                safe_sum_capacity(segment_vehicles["weight_capacity"].astype(float).tolist())
                if (not segment_vehicles.empty and "weight_capacity" in segment_vehicles.columns)
                else float("inf")
            )
            overload_after_first_run = max(0.0, demand - first_run_capacity)
            overload_weight_after_first_run = (
                max(0.0, weight_demand - first_run_weight_capacity)
                if is_finite_positive(first_run_weight_capacity)
                else 0.0
            )
            runs_needed_by_cbm = int(math.ceil(demand / first_run_capacity)) if first_run_capacity > 0 else 0
            runs_needed_by_weight = (
                int(math.ceil(weight_demand / first_run_weight_capacity))
                if is_finite_positive(first_run_weight_capacity)
                else 0
            )
            estimated_runs_needed = max(runs_needed_by_cbm, runs_needed_by_weight)
            capacity_plan_rows.append(
                {
                    "supply_chain": supply_chain,
                    "warehouse_id": warehouse_id,
                    "segment": segment_key,
                    "warehouse_total_vehicles_on_hand": warehouse_total_vehicles_on_hand,
                    "warehouse_total_capacity_cbm_on_hand": warehouse_total_cbm_capacity_on_hand,
                    "warehouse_total_capacity_weight_on_hand": format_capacity_value(warehouse_total_weight_capacity_on_hand),
                    "warehouse_total_stops_on_hand": warehouse_total_stops_capacity_on_hand,
                    "orders_count": int(len(segment_orders)),
                    "routes_count": int(segment_orders["route"].nunique()),
                    "first_run_vehicles_count": int(len(segment_vehicles)),
                    "max_runs_per_vehicle": (
                        int(config.max_runs_per_vehicle) if int(config.max_runs_per_vehicle) > 0 else "UNLIMITED"
                    ),
                    "total_demand_cbm": demand,
                    "first_run_total_capacity_cbm": first_run_capacity,
                    "overload_after_first_run_cbm": overload_after_first_run,
                    "total_demand_weight": weight_demand,
                    "first_run_total_capacity_weight": format_capacity_value(first_run_weight_capacity),
                    "overload_after_first_run_weight": overload_weight_after_first_run,
                    "estimated_runs_by_cbm": runs_needed_by_cbm,
                    "estimated_runs_by_weight": runs_needed_by_weight,
                    "estimated_runs_needed": estimated_runs_needed,
                }
            )

            bins, leftover, capacity_unassigned = assign_scoped_orders(
                segment_orders,
                segment_vehicles,
                supply_chain,
                warehouse_id,
                segment_key,
                config,
                route_neighbors,
            )
            post_validation_overflow = enforce_hard_limits_on_bins(bins, config)
            if post_validation_overflow:
                log_warning(
                    f"Post-validation trimmed {len(post_validation_overflow)} overflow orders "
                    f"in {supply_chain}/{warehouse_id}/{segment_key}."
                )
                for overflow_order in post_validation_overflow:
                    tagged = dict(overflow_order)
                    tagged["unassigned_reason"] = "post_validation_vehicle_capacity_exceeded"
                    capacity_unassigned.append(tagged)

            for b in bins:
                if not b.orders:
                    continue
                is_second_run_value, run_wave_value = build_run_wave_fields(b.run_number, config)
                high_count = 0
                for order in b.orders:
                    is_high = get_order_cbm(order) >= (
                        b.capacity * max(0.0, config.high_volume_min_load_ratio)
                    )
                    if is_high:
                        high_count += 1
                    assignment_rows.append(
                        {
                            "order_id": order.get("_order_id", ""),
                            "supply_chain": order.get("supply_chain", ""),
                            "warehouse_id": order.get("warehouse_id", ""),
                            "segment": order.get("segment", segment_key),
                            "route": order.get("route", ""),
                            "retailer_lat": order.get("lat", 0.0),
                            "retailer_long": order.get("lon", 0.0),
                            "order_load": order.get("order_load", 0.0),
                            "order_cbm": order.get("order_cbm", order.get("order_load", 0.0)),
                            "order_weight": order.get("order_weight", 0.0),
                            "purchased_items": order.get("purchased_items", 0.0),
                            "runsheet_id": b.runsheet_id,
                            "run_number": b.run_number,
                            "is_second_run": is_second_run_value,
                            "run_wave": run_wave_value,
                            "dispatch_priority": "",
                            "vehicle_id": b.vehicle_id,
                            "assigned_agent": b.assigned_agent,
                            "vehicle_type": b.vehicle_type,
                        }
                    )
                summary_rows.append(
                    {
                        "runsheet_id": b.runsheet_id,
                        "supply_chain": b.supply_chain,
                        "warehouse_id": b.warehouse_id,
                        "segment": b.segment,
                        "route_seed": b.route,
                        "covered_routes": " | ".join(b.covered_routes),
                        "run_number": b.run_number,
                        "is_second_run": is_second_run_value,
                        "run_wave": run_wave_value,
                        "dispatch_priority": "",
                        "vehicle_id": b.vehicle_id,
                        "assigned_agent": b.assigned_agent,
                        "vehicle_type": b.vehicle_type,
                        "vehicle_capacity": b.capacity,
                        "assigned_load": b.assigned_load,
                        "remaining_capacity": b.remaining,
                        "vehicle_weight_capacity": format_capacity_value(b.weight_capacity),
                        "assigned_weight": b.assigned_weight,
                        "remaining_weight_capacity": format_capacity_value(b.weight_remaining),
                        "cbm_utilization_pct": b.cbm_utilization_pct,
                        "weight_utilization_pct": b.weight_utilization_pct,
                        "utilization_pct": b.utilization_pct,
                        "orders_count": len(b.orders),
                        "purchased_items_total": sum(
                            safe_float(order.get("purchased_items", 0), 0.0) for order in b.orders
                        ),
                        "min_utilization_target_pct": config.min_utilization_target_pct,
                        "meets_utilization_target": b.utilization_pct >= safe_float(config.min_utilization_target_pct, 0.0),
                        "max_stops_per_run": b.stop_capacity if b.stop_capacity > 0 else config.max_stops_per_run,
                        "high_volume_orders_count": high_count,
                    }
                )

            for order in leftover:
                unassigned_rows.append(
                    {
                        "order_id": order.get("_order_id", ""),
                        "supply_chain": order.get("supply_chain", ""),
                        "warehouse_id": order.get("warehouse_id", ""),
                        "segment": order.get("segment", segment_key),
                        "route": order.get("route", ""),
                        "retailer_lat": order.get("lat", 0.0),
                        "retailer_long": order.get("lon", 0.0),
                        "order_load": order.get("order_load", 0.0),
                        "order_cbm": order.get("order_cbm", order.get("order_load", 0.0)),
                        "order_weight": order.get("order_weight", 0.0),
                        "purchased_items": order.get("purchased_items", 0.0),
                        "reason": "No compatible vehicle/scheduling slot",
                    }
                )
            for order in capacity_unassigned:
                unassigned_rows.append(
                    {
                        "order_id": order.get("_order_id", ""),
                        "supply_chain": order.get("supply_chain", ""),
                        "warehouse_id": order.get("warehouse_id", ""),
                        "segment": order.get("segment", segment_key),
                        "route": order.get("route", ""),
                        "retailer_lat": order.get("lat", 0.0),
                        "retailer_long": order.get("lon", 0.0),
                        "order_load": order.get("order_load", 0.0),
                        "order_cbm": order.get("order_cbm", order.get("order_load", 0.0)),
                        "order_weight": order.get("order_weight", 0.0),
                        "purchased_items": order.get("purchased_items", 0.0),
                        "reason": safe_str(order.get("unassigned_reason", ""), "No compatible vehicle/scheduling slot"),
                    }
                )

    assignment_df = pd.DataFrame(assignment_rows)
    summary_df = pd.DataFrame(summary_rows)
    unassigned_df = pd.DataFrame(unassigned_rows)
    capacity_df = pd.DataFrame(capacity_plan_rows)
    return assignment_df, summary_df, unassigned_df, capacity_df


def export_output(
    assignment_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    unassigned_df: pd.DataFrame,
    capacity_df: pd.DataFrame,
    config: RunSheetConfig,
) -> str:
    now = datetime.now()
    month_folder = now.strftime("%Y-%m")
    out_dir = Path(config.output_root) / month_folder
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / now.strftime("%d-%m-%Y_RUNSHEET_PLAN.xlsx")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        assignment_df.to_excel(writer, sheet_name="order_assignment", index=False)
        summary_df.to_excel(writer, sheet_name="runsheet_summary", index=False)
        capacity_df.to_excel(writer, sheet_name="capacity_plan", index=False)
        if not unassigned_df.empty:
            unassigned_df.to_excel(writer, sheet_name="unassigned_orders", index=False)

    return str(out_path)


def main() -> None:
    config = RunSheetConfig()
    orders_df = load_orders(config)
    vehicles_df = load_vehicles(config)
    kml_centroids = parse_kml_route_centroids(config.kml_file)

    log_info(
        f"Loaded {len(orders_df)} orders across {orders_df['supply_chain'].nunique()} supply chains "
        f"and {orders_df['warehouse_id'].nunique()} warehouses and {orders_df['route'].nunique()} routes."
    )
    log_info(f"Loaded {len(vehicles_df)} active vehicles.")

    assignment_df, summary_df, unassigned_df, capacity_df = build_runsheets(
        orders_df,
        vehicles_df,
        config,
        kml_centroids,
    )
    if assignment_df.empty:
        log_warning("No runsheets were generated.")
        return

    if not summary_df.empty:
        max_run = int(summary_df["run_number"].max())
        log_info(f"Max run number generated in one warehouse/supply-chain scope: {max_run}")

    output_path = export_output(assignment_df, summary_df, unassigned_df, capacity_df, config)
    log_info(f"Runsheet plan generated: {output_path}")
    log_info(f"Assigned orders: {len(assignment_df)}")
    log_info(f"Runsheets created: {summary_df['runsheet_id'].nunique() if not summary_df.empty else 0}")
    if not unassigned_df.empty:
        log_warning(f"Unassigned orders: {len(unassigned_df)} (see 'unassigned_orders' sheet).")


if __name__ == "__main__":
    main()

