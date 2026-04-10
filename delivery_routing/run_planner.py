#!/usr/bin/env python3
"""
Delivery run planner.

Builds delivery runs from an Excel order file using these rules:
1) Keep each run inside one polygon whenever possible.
2) If capacity remains, fill the run with closest retailers/orders.
3) Vehicle capacity can change by supply chain via config.
4) Output both run sheet and load summary sheet.
"""

from __future__ import annotations

import argparse
import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Sequence

import pandas as pd


NO_POLYGON = "NO_POLYGON"


@dataclass(frozen=True)
class ColumnMap:
    order_id: str
    retailer_id: str
    retailer_name: str
    latitude: str
    longitude: str
    product: str
    quantity: str
    supply_chain: str

    @classmethod
    def from_config(cls, raw: Dict[str, Any]) -> "ColumnMap":
        required = {
            "order_id",
            "retailer_id",
            "retailer_name",
            "latitude",
            "longitude",
            "product",
            "quantity",
            "supply_chain",
        }
        missing = [key for key in required if key not in raw]
        if missing:
            raise ValueError(f"Missing column mappings in config: {', '.join(sorted(missing))}")
        return cls(**{k: raw[k] for k in required})

    def as_list(self) -> List[str]:
        return [
            self.order_id,
            self.retailer_id,
            self.retailer_name,
            self.latitude,
            self.longitude,
            self.product,
            self.quantity,
            self.supply_chain,
        ]


def sanitize_label(text: Any) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "-", str(text).strip()).strip("-")
    return cleaned.upper() or "SC"


def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius_km = 6371.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    d_phi = math.radians(lat2 - lat1)
    d_lambda = math.radians(lon2 - lon1)
    a = (
        math.sin(d_phi / 2) ** 2
        + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2) ** 2
    )
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return radius_km * c


def point_in_polygon(latitude: float, longitude: float, polygon: Sequence[Sequence[float]]) -> bool:
    """
    Ray-casting algorithm.
    Polygon points are expected as [lat, lon].
    """
    x, y = longitude, latitude
    inside = False
    points = list(polygon)
    n = len(points)
    if n < 3:
        return False

    for i in range(n):
        lat1, lon1 = points[i]
        lat2, lon2 = points[(i + 1) % n]
        x1, y1 = lon1, lat1
        x2, y2 = lon2, lat2

        intersects = ((y1 > y) != (y2 > y)) and (
            x < (x2 - x1) * (y - y1) / ((y2 - y1) or 1e-12) + x1
        )
        if intersects:
            inside = not inside
    return inside


def find_polygon_id(latitude: float, longitude: float, polygons: Sequence[Dict[str, Any]]) -> str:
    for polygon in polygons:
        if point_in_polygon(latitude, longitude, polygon["points"]):
            return polygon["id"]
    return NO_POLYGON


def _build_order_key(df: pd.DataFrame, key_columns: Sequence[str]) -> pd.Series:
    text_cols = [df[col].astype(str).fillna("") for col in key_columns]
    return pd.concat(text_cols, axis=1).agg("||".join, axis=1)


def _run_centroid(order_ids: Iterable[str], order_map: Dict[str, Dict[str, Any]]) -> tuple[float, float]:
    coords = [(order_map[item]["latitude"], order_map[item]["longitude"]) for item in order_ids]
    lat = sum(p[0] for p in coords) / len(coords)
    lon = sum(p[1] for p in coords) / len(coords)
    return lat, lon


def _nearest_fitting_order(
    run_order_ids: List[str],
    candidate_ids: Iterable[str],
    order_map: Dict[str, Dict[str, Any]],
    max_distance_km: float | None,
) -> str | None:
    center_lat, center_lon = _run_centroid(run_order_ids, order_map)
    best_id = None
    best_distance = None

    for item in candidate_ids:
        lat = order_map[item]["latitude"]
        lon = order_map[item]["longitude"]
        distance = haversine_km(center_lat, center_lon, lat, lon)
        if max_distance_km is not None and distance > max_distance_km:
            continue
        if best_distance is None or distance < best_distance:
            best_distance = distance
            best_id = item
    return best_id


def _build_single_run(
    seed_order_id: str,
    candidate_ids: set[str],
    order_map: Dict[str, Dict[str, Any]],
    capacity: float,
) -> tuple[List[str], float]:
    run_order_ids = [seed_order_id]
    used = float(order_map[seed_order_id]["demand"])
    remaining_ids = set(candidate_ids)
    remaining_ids.discard(seed_order_id)

    while remaining_ids and used < capacity:
        remaining_capacity = capacity - used
        fitting = [item for item in remaining_ids if order_map[item]["demand"] <= remaining_capacity]
        if not fitting:
            break
        next_id = _nearest_fitting_order(run_order_ids, fitting, order_map, max_distance_km=None)
        if not next_id:
            break
        run_order_ids.append(next_id)
        used += float(order_map[next_id]["demand"])
        remaining_ids.discard(next_id)

    return run_order_ids, used


def assign_runs_for_supply_chain(
    orders_df: pd.DataFrame,
    capacity: float,
    allow_cross_polygon_fill: bool,
    max_cross_fill_distance_km: float | None,
) -> List[Dict[str, Any]]:
    if orders_df.empty:
        return []

    order_map = {row["order_key"]: row for row in orders_df.to_dict("records")}
    unassigned: set[str] = set(order_map.keys())
    runs: List[Dict[str, Any]] = []

    # 1) Build polygon-first runs.
    polygon_ids = list(orders_df["polygon_id"].fillna(NO_POLYGON).unique())
    for polygon_id in polygon_ids:
        polygon_orders = {
            item
            for item in unassigned
            if str(order_map[item]["polygon_id"]) == str(polygon_id)
        }
        while polygon_orders:
            seed = max(polygon_orders, key=lambda item: float(order_map[item]["demand"]))
            seed_demand = float(order_map[seed]["demand"])

            if seed_demand > capacity:
                run_order_ids = [seed]
                used = seed_demand
            else:
                run_order_ids, used = _build_single_run(
                    seed_order_id=seed,
                    candidate_ids=polygon_orders,
                    order_map=order_map,
                    capacity=capacity,
                )

            for item in run_order_ids:
                unassigned.discard(item)
                polygon_orders.discard(item)

            runs.append(
                {
                    "primary_polygon": polygon_id,
                    "order_ids": run_order_ids,
                    "used_capacity": used,
                }
            )

    # 2) Fill remaining capacity with nearest unassigned orders.
    if allow_cross_polygon_fill:
        for run in runs:
            while unassigned and run["used_capacity"] < capacity:
                remaining_capacity = capacity - run["used_capacity"]
                fitting = [item for item in unassigned if order_map[item]["demand"] <= remaining_capacity]
                if not fitting:
                    break
                next_id = _nearest_fitting_order(
                    run_order_ids=run["order_ids"],
                    candidate_ids=fitting,
                    order_map=order_map,
                    max_distance_km=max_cross_fill_distance_km,
                )
                if not next_id:
                    break
                run["order_ids"].append(next_id)
                run["used_capacity"] += float(order_map[next_id]["demand"])
                unassigned.discard(next_id)

    # 3) Anything still unassigned gets new nearest-neighbor runs.
    while unassigned:
        seed = max(unassigned, key=lambda item: float(order_map[item]["demand"]))
        seed_demand = float(order_map[seed]["demand"])
        if seed_demand > capacity:
            run_order_ids = [seed]
            used = seed_demand
        else:
            run_order_ids, used = _build_single_run(
                seed_order_id=seed,
                candidate_ids=unassigned,
                order_map=order_map,
                capacity=capacity,
            )
        for item in run_order_ids:
            unassigned.discard(item)
        runs.append(
            {
                "primary_polygon": str(order_map[seed]["polygon_id"]),
                "order_ids": run_order_ids,
                "used_capacity": used,
            }
        )

    return runs


def build_outputs(
    orders_raw: pd.DataFrame,
    columns: ColumnMap,
    polygons: Sequence[Dict[str, Any]],
    default_capacity: float,
    capacity_by_supply_chain: Dict[str, float],
    allow_cross_polygon_fill: bool,
    max_cross_fill_distance_km: float | None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    required = set(columns.as_list())
    missing_input_columns = [column for column in required if column not in orders_raw.columns]
    if missing_input_columns:
        raise ValueError(
            f"Missing required input columns in Excel: {', '.join(sorted(missing_input_columns))}"
        )

    df = orders_raw.copy()
    df[columns.latitude] = pd.to_numeric(df[columns.latitude], errors="coerce")
    df[columns.longitude] = pd.to_numeric(df[columns.longitude], errors="coerce")
    df[columns.quantity] = pd.to_numeric(df[columns.quantity], errors="coerce")

    bad_rows = df[
        df[columns.latitude].isna() | df[columns.longitude].isna() | df[columns.quantity].isna()
    ]
    if not bad_rows.empty:
        raise ValueError(
            "Some rows have invalid latitude/longitude/quantity values. "
            "Please clean the input file before running."
        )

    key_columns = [
        columns.order_id,
        columns.supply_chain,
        columns.retailer_id,
        columns.latitude,
        columns.longitude,
    ]
    df["order_key"] = _build_order_key(df, key_columns=key_columns)
    df["polygon_id"] = df.apply(
        lambda row: find_polygon_id(
            latitude=float(row[columns.latitude]),
            longitude=float(row[columns.longitude]),
            polygons=polygons,
        ),
        axis=1,
    )

    orders = (
        df.groupby("order_key", as_index=False)
        .agg(
            {
                columns.order_id: "first",
                columns.retailer_id: "first",
                columns.retailer_name: "first",
                columns.latitude: "first",
                columns.longitude: "first",
                columns.supply_chain: "first",
                "polygon_id": "first",
                columns.quantity: "sum",
            }
        )
        .rename(
            columns={
                columns.order_id: "order_id",
                columns.retailer_id: "retailer_id",
                columns.retailer_name: "retailer_name",
                columns.latitude: "latitude",
                columns.longitude: "longitude",
                columns.supply_chain: "supply_chain",
                columns.quantity: "demand",
            }
        )
    )

    assignment_rows: List[Dict[str, Any]] = []

    for supply_chain, supply_chain_orders in orders.groupby("supply_chain"):
        capacity = float(capacity_by_supply_chain.get(str(supply_chain), default_capacity))
        if capacity <= 0:
            raise ValueError(
                f"Vehicle capacity must be > 0 for supply chain '{supply_chain}'. Got {capacity}."
            )

        runs = assign_runs_for_supply_chain(
            orders_df=supply_chain_orders,
            capacity=capacity,
            allow_cross_polygon_fill=allow_cross_polygon_fill,
            max_cross_fill_distance_km=max_cross_fill_distance_km,
        )
        run_label = sanitize_label(supply_chain)
        for run_index, run in enumerate(runs, start=1):
            run_id = f"{run_label}-{run_index:03d}"
            run_polygons = {
                str(
                    supply_chain_orders.loc[
                        supply_chain_orders["order_key"] == order_key, "polygon_id"
                    ].iloc[0]
                )
                for order_key in run["order_ids"]
            }
            run_is_mixed = len(run_polygons) > 1
            used_capacity = float(run["used_capacity"])
            remaining_capacity = capacity - used_capacity
            over_capacity = max(0.0, used_capacity - capacity)

            for stop_sequence, order_key in enumerate(run["order_ids"], start=1):
                order_row = supply_chain_orders.loc[supply_chain_orders["order_key"] == order_key].iloc[0]
                assignment_rows.append(
                    {
                        "order_key": order_key,
                        "run_id": run_id,
                        "supply_chain": supply_chain,
                        "stop_sequence": stop_sequence,
                        "vehicle_capacity": capacity,
                        "used_capacity": round(used_capacity, 3),
                        "remaining_capacity": round(remaining_capacity, 3),
                        "over_capacity": round(over_capacity, 3),
                        "primary_polygon": run["primary_polygon"],
                        "mixed_polygons": run_is_mixed,
                        "run_order_count": len(run["order_ids"]),
                        "run_polygon_count": len(run_polygons),
                        "order_demand": float(order_row["demand"]),
                        "order_id": order_row["order_id"],
                        "retailer_id": order_row["retailer_id"],
                        "retailer_name": order_row["retailer_name"],
                        "latitude": float(order_row["latitude"]),
                        "longitude": float(order_row["longitude"]),
                        "polygon_id": str(order_row["polygon_id"]),
                    }
                )

    assignments = pd.DataFrame(assignment_rows)
    if assignments.empty:
        raise ValueError("No assignments were generated. Check input data and config.")

    run_sheet = df.merge(assignments, on="order_key", how="left")
    run_sheet = run_sheet.sort_values(
        by=["run_id", "stop_sequence", columns.order_id, columns.product], na_position="last"
    )

    load_summary = (
        run_sheet.groupby(["run_id", "supply_chain", columns.product], as_index=False)[columns.quantity]
        .sum()
        .rename(
            columns={
                columns.product: "product",
                columns.quantity: "total_quantity",
            }
        )
        .sort_values(by=["run_id", "product"])
    )

    return run_sheet, load_summary


def load_config(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as fp:
        return json.load(fp)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate delivery run sheet and vehicle load summary from Excel orders."
    )
    parser.add_argument("--input", required=True, help="Input Excel file path.")
    parser.add_argument("--config", required=True, help="JSON config file path.")
    parser.add_argument("--output", required=True, help="Output Excel file path.")
    parser.add_argument(
        "--sheet",
        default=None,
        help="Optional input sheet name override. If not provided, uses config sheet_name or first sheet.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = load_config(Path(args.config))

    columns = ColumnMap.from_config(config.get("columns", {}))
    sheet_name = args.sheet if args.sheet is not None else config.get("sheet_name", 0)
    polygons = config.get("polygons", [])
    routing_cfg = config.get("routing", {})

    default_capacity = routing_cfg.get("default_vehicle_capacity")
    if default_capacity is None:
        raise ValueError("routing.default_vehicle_capacity is required in config.")
    default_capacity = float(default_capacity)

    capacity_by_supply_chain = {
        str(k): float(v) for k, v in routing_cfg.get("capacity_by_supply_chain", {}).items()
    }
    allow_cross_polygon_fill = bool(routing_cfg.get("allow_cross_polygon_fill", True))
    max_cross_fill_distance_km = routing_cfg.get("max_cross_fill_distance_km")
    if max_cross_fill_distance_km is not None:
        max_cross_fill_distance_km = float(max_cross_fill_distance_km)

    orders_raw = pd.read_excel(args.input, sheet_name=sheet_name)
    run_sheet, load_summary = build_outputs(
        orders_raw=orders_raw,
        columns=columns,
        polygons=polygons,
        default_capacity=default_capacity,
        capacity_by_supply_chain=capacity_by_supply_chain,
        allow_cross_polygon_fill=allow_cross_polygon_fill,
        max_cross_fill_distance_km=max_cross_fill_distance_km,
    )

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        run_sheet.to_excel(writer, sheet_name="run_sheet", index=False)
        load_summary.to_excel(writer, sheet_name="load_summary", index=False)

    total_runs = run_sheet["run_id"].nunique()
    print(f"Created {output_path} with {total_runs} runs.")


if __name__ == "__main__":
    main()
