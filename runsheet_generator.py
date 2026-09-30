import math
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd


# =============================
# SETTINGS (EDIT THIS SECTION)
# =============================
ORDERS_FILE = r"E:\Marbah Products\Marbah Scripts\Marbah Invoices script\orders_data.xlsx"
VEHICLES_FILE = r"E:\Marbah Products\Marbah Scripts\Marbah Invoices script\vehicles_data.xlsx"
OUTPUT_ROOT = r"E:\Marbah Products\Marbah Scripts\Marbah Invoices script"


@dataclass
class RunSheetConfig:
    orders_file: str = ORDERS_FILE
    vehicles_file: str = VEHICLES_FILE
    output_root: str = OUTPUT_ROOT

    # Order columns
    order_id_candidates: Tuple[str, ...] = ("order_id", "order_number", "order_no", "orderid")
    supply_chain_candidates: Tuple[str, ...] = ("supply_chain", "supply chain", "supply_chain_name", "supplychain")
    warehouse_candidates: Tuple[str, ...] = ("warehouse_id", "warehouse", "warehouse_name", "wh_id", "depot")
    route_candidates: Tuple[str, ...] = ("route", "route_name", "delivery_route", "route_text")
    latitude_candidates: Tuple[str, ...] = ("retailer_lat", "latitude", "lat", "customer_lat", "store_lat")
    longitude_candidates: Tuple[str, ...] = ("retailer_long", "longitude", "long", "lng", "customer_long", "store_long")
    # CBM is prioritized for planning load/capacity.
    order_cbm_candidates: Tuple[str, ...] = ("order_cbm", "cbm", "total_cbm", "volume_cbm", "volume")
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
    vehicle_agent_candidates: Tuple[str, ...] = ("assigned_agent", "agent", "driver", "route_agent", "agent_name")
    vehicle_type_candidates: Tuple[str, ...] = ("vehicle_type", "type", "truck_type")
    vehicle_capacity_candidates: Tuple[str, ...] = ("vehicle_capacity", "capacity", "load_capacity", "max_load")
    vehicle_count_candidates: Tuple[str, ...] = ("vehicle_count", "count", "qty")
    vehicle_active_candidates: Tuple[str, ...] = ("active", "is_active", "enabled")

    # Assignment behavior
    high_volume_min_load_ratio: float = 0.60
    max_stops_per_run: int = 23
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


def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    r = 6371.0
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlon / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


class VehicleBin:
    def __init__(
        self,
        vehicle_id: str,
        supply_chain: str,
        warehouse_id: str,
        assigned_agent: str,
        vehicle_type: str,
        route: str,
        run_number: int,
        capacity: float,
        runsheet_id: str,
    ):
        self.vehicle_id = vehicle_id
        self.supply_chain = supply_chain
        self.warehouse_id = warehouse_id
        self.assigned_agent = assigned_agent
        self.vehicle_type = vehicle_type
        self.route = route
        self.run_number = run_number
        self.capacity = max(0.0, capacity)
        self.remaining = max(0.0, capacity)
        self.runsheet_id = runsheet_id
        self.orders: List[Dict[str, object]] = []
        self.centroid_lat: Optional[float] = None
        self.centroid_lon: Optional[float] = None

    def can_fit(self, load: float, max_stops_per_run: int = 0) -> bool:
        if max_stops_per_run > 0 and len(self.orders) >= max_stops_per_run:
            return False
        return self.remaining >= load

    def add_order(self, order: Dict[str, object]) -> None:
        load = safe_float(order.get("order_load", 0), 0.0)
        self.orders.append(order)
        self.remaining -= load
        lat = safe_float(order.get("lat", 0), 0.0)
        lon = safe_float(order.get("lon", 0), 0.0)
        if self.centroid_lat is None or self.centroid_lon is None:
            self.centroid_lat = lat
            self.centroid_lon = lon
            return
        count = len(self.orders)
        self.centroid_lat = ((self.centroid_lat * (count - 1)) + lat) / count
        self.centroid_lon = ((self.centroid_lon * (count - 1)) + lon) / count

    @property
    def assigned_load(self) -> float:
        return self.capacity - self.remaining

    @property
    def utilization_pct(self) -> float:
        if self.capacity <= 0:
            return 0.0
        return (self.assigned_load / self.capacity) * 100.0


def choose_best_bin(order: Dict[str, object], bins: List[VehicleBin], max_stops_per_run: int = 0) -> Optional[VehicleBin]:
    load = safe_float(order.get("order_load", 0), 0.0)
    fit_bins = [b for b in bins if b.can_fit(load, max_stops_per_run)]
    if not fit_bins:
        return None

    lat = safe_float(order.get("lat", 0), 0.0)
    lon = safe_float(order.get("lon", 0), 0.0)
    scored: List[Tuple[float, VehicleBin]] = []
    for b in fit_bins:
        remaining_after = b.remaining - load
        if b.centroid_lat is None or b.centroid_lon is None:
            distance = 0.0
        else:
            distance = haversine_km(lat, lon, b.centroid_lat, b.centroid_lon)
        score = (distance * 1.5) + remaining_after
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
    route_col = find_existing_column(raw, config.route_candidates)
    lat_col = find_existing_column(raw, config.latitude_candidates)
    lon_col = find_existing_column(raw, config.longitude_candidates)
    cbm_col = find_existing_column(raw, config.order_cbm_candidates)
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
    orders["_route"] = orders[route_col].apply(lambda x: normalize_key(x))
    orders["_lat"] = orders[lat_col].apply(lambda x: safe_float(x, 0.0))
    orders["_lon"] = orders[lon_col].apply(lambda x: safe_float(x, 0.0))
    if cbm_col is not None:
        orders["_line_load"] = orders[cbm_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
        log_info(f"Using CBM column for load planning: {cbm_col}")
    elif load_fallback_col is not None:
        orders["_line_load"] = orders[load_fallback_col].apply(lambda x: max(0.0, safe_float(x, 0.0)))
        log_warning(f"CBM column not found. Using fallback load column: {load_fallback_col}")
    else:
        orders["_line_load"] = 1.0
        log_warning("No CBM/load column found. Using default load=1 per order line.")
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
            route=("_route", "first"),
            lat=("_lat", "first"),
            lon=("_lon", "first"),
            order_load=("_line_load", aggregate_order_measure),
            delivery_date=("_delivery_date", "max"),
        )
    )
    grouped["order_load"] = grouped["order_load"].apply(lambda x: max(0.0, safe_float(x, 0.0)))
    grouped = grouped[grouped["order_load"] > 0]
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
    agent_col = find_existing_column(raw, config.vehicle_agent_candidates)
    type_col = find_existing_column(raw, config.vehicle_type_candidates)
    capacity_col = find_existing_column(raw, config.vehicle_capacity_candidates)
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
    vehicles = vehicles[
        (vehicles["_supply_chain"] != "")
        & (vehicles["_warehouse_id"] != "")
        & (vehicles["_capacity"] > 0)
    ]

    if active_col is not None and not config.include_inactive_vehicles:
        vehicles = vehicles[vehicles[active_col].apply(lambda x: bool_from_value(x, default=True))]

    rows: List[Dict[str, object]] = []
    if vehicle_id_col is not None:
        for _, row in vehicles.iterrows():
            vehicle_id = normalize_identifier(row.get(vehicle_id_col, ""), "")
            if not vehicle_id:
                continue
            rows.append(
                {
                    "vehicle_id": vehicle_id,
                    "supply_chain": row["_supply_chain"],
                    "warehouse_id": row["_warehouse_id"],
                    "assigned_agent": row["_assigned_agent"],
                    "vehicle_type": row["_vehicle_type"],
                    "capacity": row["_capacity"],
                }
            )
    else:
        for _, row in vehicles.iterrows():
            count = int(max(0.0, safe_float(row.get(count_col, 0), 0.0)))
            chain = safe_str(row["_supply_chain"], "")
            warehouse_id = safe_str(row["_warehouse_id"], "")
            assigned_agent = safe_str(row["_assigned_agent"], "")
            vehicle_type = safe_str(row["_vehicle_type"], "GENERIC") or "GENERIC"
            for idx in range(1, count + 1):
                generated_id = f"{chain}_{warehouse_id}_{vehicle_type}_V{idx}"
                rows.append(
                    {
                        "vehicle_id": re.sub(r"[^A-Za-z0-9_]+", "", generated_id)[:50],
                        "supply_chain": chain,
                        "warehouse_id": warehouse_id,
                        "assigned_agent": assigned_agent,
                        "vehicle_type": vehicle_type,
                        "capacity": row["_capacity"],
                    }
                )
    out = pd.DataFrame(rows)
    if out.empty:
        raise ValueError("No active vehicles available after preprocessing.")
    return out


def build_runsheet_id(
    prefix: str,
    supply_chain: str,
    warehouse_id: str,
    route: str,
    run_number: int,
    vehicle_id: str,
) -> str:
    chain_clean = re.sub(r"[^A-Z0-9]+", "", normalize_key(supply_chain)) or "SC"
    warehouse_clean = re.sub(r"[^A-Za-z0-9]+", "", safe_str(warehouse_id, "W"))
    route_clean = re.sub(r"[^A-Za-z0-9]+", "", safe_str(route, "R"))
    vehicle_clean = re.sub(r"[^A-Za-z0-9]+", "", safe_str(vehicle_id, "V"))
    return f"{prefix}-{chain_clean}-{warehouse_clean}-{route_clean}-R{run_number:02d}-{vehicle_clean}"[:60]


def assign_group_orders(
    group_orders: pd.DataFrame,
    scoped_vehicles: pd.DataFrame,
    supply_chain: str,
    warehouse_id: str,
    route: str,
    config: RunSheetConfig,
) -> Tuple[List[VehicleBin], List[Dict[str, object]]]:
    if scoped_vehicles.empty:
        unassigned = group_orders.to_dict("records")
        return [], unassigned

    vehicle_list = scoped_vehicles.sort_values("vehicle_id").to_dict("records")
    capacity_reference = float(pd.Series([safe_float(v["capacity"], 0.0) for v in vehicle_list]).median())
    if capacity_reference <= 0:
        capacity_reference = float(pd.Series([safe_float(v["capacity"], 0.0) for v in vehicle_list]).mean())
    if capacity_reference <= 0:
        capacity_reference = 1.0
    high_volume_threshold = capacity_reference * max(0.0, config.high_volume_min_load_ratio)

    unassigned = group_orders.to_dict("records")
    created_bins: List[VehicleBin] = []
    run_number = 1
    safety_counter = 0

    while unassigned:
        safety_counter += 1
        if safety_counter > 500:
            log_warning(f"Safety break triggered for {supply_chain}/{warehouse_id}/{route}.")
            break

        bins: List[VehicleBin] = []
        for v in vehicle_list:
            bin_obj = VehicleBin(
                vehicle_id=safe_str(v["vehicle_id"], ""),
                supply_chain=supply_chain,
                warehouse_id=safe_str(warehouse_id, ""),
                assigned_agent=safe_str(v.get("assigned_agent", ""), ""),
                vehicle_type=safe_str(v.get("vehicle_type", "GENERIC"), "GENERIC"),
                route=safe_str(route, ""),
                run_number=run_number,
                capacity=safe_float(v["capacity"], 0.0),
                runsheet_id=build_runsheet_id(
                    config.runsheet_prefix,
                    supply_chain,
                    safe_str(warehouse_id, ""),
                    safe_str(route, ""),
                    run_number,
                    safe_str(v["vehicle_id"], ""),
                ),
            )
            bins.append(bin_obj)

        high_orders = [o for o in unassigned if safe_float(o.get("order_load", 0), 0.0) >= high_volume_threshold]
        normal_orders = [o for o in unassigned if safe_float(o.get("order_load", 0), 0.0) < high_volume_threshold]
        high_orders.sort(key=lambda o: safe_float(o.get("order_load", 0), 0.0), reverse=True)
        normal_orders.sort(key=lambda o: safe_float(o.get("order_load", 0), 0.0), reverse=True)
        ordered_candidates = high_orders + normal_orders

        assigned_ids = set()
        for order in ordered_candidates:
            chosen_bin = choose_best_bin(order, bins, config.max_stops_per_run)
            if chosen_bin is None:
                continue
            chosen_bin.add_order(order)
            assigned_ids.add(safe_str(order.get("_order_id", ""), ""))

        if not assigned_ids:
            # If all bins reached stop cap, open next run and retry.
            stop_cap = max(0, int(config.max_stops_per_run))
            has_stop_slot = any(stop_cap <= 0 or len(b.orders) < stop_cap for b in bins)
            if not has_stop_slot:
                run_number += 1
                continue

            # Oversized order fallback: assign largest order to vehicle with max capacity.
            unassigned.sort(key=lambda o: safe_float(o.get("order_load", 0), 0.0), reverse=True)
            forced = unassigned[0]
            candidate_bins = [b for b in bins if stop_cap <= 0 or len(b.orders) < stop_cap]
            biggest_bin = max(candidate_bins, key=lambda b: b.capacity)
            biggest_bin.add_order(forced)
            assigned_ids.add(safe_str(forced.get("_order_id", ""), ""))
            log_warning(
                f"Oversized order '{forced.get('_order_id')}' exceeded vehicle capacity in "
                f"{supply_chain}/{warehouse_id}/{route}; forced assignment applied."
            )

        non_empty_bins = [b for b in bins if b.orders]
        created_bins.extend(non_empty_bins)
        unassigned = [o for o in unassigned if safe_str(o.get("_order_id", ""), "") not in assigned_ids]
        run_number += 1

    return created_bins, unassigned


def build_runsheets(orders: pd.DataFrame, vehicles: pd.DataFrame, config: RunSheetConfig) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    assignment_rows: List[Dict[str, object]] = []
    summary_rows: List[Dict[str, object]] = []
    unassigned_rows: List[Dict[str, object]] = []

    vehicles_by_chain_warehouse = {
        (chain, warehouse): group.copy()
        for (chain, warehouse), group in vehicles.groupby(["supply_chain", "warehouse_id"], sort=False)
    }

    for (supply_chain, warehouse_id, route), group_df in orders.groupby(
        ["supply_chain", "warehouse_id", "route"],
        sort=True,
    ):
        scoped_vehicles = vehicles_by_chain_warehouse.get((supply_chain, warehouse_id), pd.DataFrame())
        bins, leftover = assign_group_orders(
            group_df,
            scoped_vehicles,
            supply_chain,
            warehouse_id,
            route,
            config,
        )

        for b in bins:
            high_count = 0
            for order in b.orders:
                is_high = safe_float(order.get("order_load", 0), 0.0) >= (
                    b.capacity * max(0.0, config.high_volume_min_load_ratio)
                )
                if is_high:
                    high_count += 1
                assignment_rows.append(
                    {
                        "order_id": order.get("_order_id", ""),
                        "supply_chain": order.get("supply_chain", ""),
                        "warehouse_id": order.get("warehouse_id", ""),
                        "route": order.get("route", ""),
                        "retailer_lat": order.get("lat", 0.0),
                        "retailer_long": order.get("lon", 0.0),
                        "order_load": order.get("order_load", 0.0),
                        "runsheet_id": b.runsheet_id,
                        "run_number": b.run_number,
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
                    "route": b.route,
                    "run_number": b.run_number,
                    "vehicle_id": b.vehicle_id,
                    "assigned_agent": b.assigned_agent,
                    "vehicle_type": b.vehicle_type,
                    "vehicle_capacity": b.capacity,
                    "assigned_load": b.assigned_load,
                    "remaining_capacity": b.remaining,
                    "utilization_pct": b.utilization_pct,
                    "orders_count": len(b.orders),
                    "max_stops_per_run": config.max_stops_per_run,
                    "high_volume_orders_count": high_count,
                }
            )

        for order in leftover:
            unassigned_rows.append(
                {
                    "order_id": order.get("_order_id", ""),
                    "supply_chain": order.get("supply_chain", ""),
                    "warehouse_id": order.get("warehouse_id", ""),
                    "route": order.get("route", ""),
                    "retailer_lat": order.get("lat", 0.0),
                    "retailer_long": order.get("lon", 0.0),
                    "order_load": order.get("order_load", 0.0),
                    "reason": "No compatible vehicle/scheduling slot",
                }
            )

    assignment_df = pd.DataFrame(assignment_rows)
    summary_df = pd.DataFrame(summary_rows)
    unassigned_df = pd.DataFrame(unassigned_rows)
    return assignment_df, summary_df, unassigned_df


def export_output(
    assignment_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    unassigned_df: pd.DataFrame,
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
        if not unassigned_df.empty:
            unassigned_df.to_excel(writer, sheet_name="unassigned_orders", index=False)

    return str(out_path)


def main() -> None:
    config = RunSheetConfig()
    orders_df = load_orders(config)
    vehicles_df = load_vehicles(config)

    log_info(
        f"Loaded {len(orders_df)} orders across {orders_df['supply_chain'].nunique()} supply chains "
        f"and {orders_df['warehouse_id'].nunique()} warehouses and {orders_df['route'].nunique()} routes."
    )
    log_info(f"Loaded {len(vehicles_df)} active vehicles.")

    assignment_df, summary_df, unassigned_df = build_runsheets(orders_df, vehicles_df, config)
    if assignment_df.empty:
        log_warning("No runsheets were generated.")
        return

    output_path = export_output(assignment_df, summary_df, unassigned_df, config)
    log_info(f"Runsheet plan generated: {output_path}")
    log_info(f"Assigned orders: {len(assignment_df)}")
    log_info(f"Runsheets created: {summary_df['runsheet_id'].nunique() if not summary_df.empty else 0}")
    if not unassigned_df.empty:
        log_warning(f"Unassigned orders: {len(unassigned_df)} (see 'unassigned_orders' sheet).")


if __name__ == "__main__":
    main()

