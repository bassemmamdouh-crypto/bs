import re
from math import radians, cos, sin, asin, sqrt
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd

# =============================
# SETTINGS
# =============================
FILE_PATH = r"E:\Marbah Products\Other Scripts\Delivery Plan\orders_in_delivery.xlsx"

SUPPLY_CHAINS = {
    "lays": 470,
    "pepsi": 790,
}

TARGET_UTILIZATION = 0.90

# Aggressive compact routing profile.
SEED_NEIGHBOR_RADIUS_KM = 1.0
INITIAL_MAX_NEXT_STOP_KM = 1.0
INITIAL_MAX_FROM_SEED_KM = 2.0
RELAX_DISTANCE_STEP_KM = 0.4
MAX_NEXT_STOP_KM = 2.5
MAX_FROM_SEED_KM = 3.5


def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    return 6371 * 2 * asin(sqrt(a))


def parse_route_parts(route_value: str) -> List[str]:
    text = str(route_value).strip().lower()
    if text == "":
        return ["no_route"]
    parts = [p.strip() for p in re.split(r"\s*(?:->|>|,|;|\||/)\s*", text) if p.strip()]
    return parts if parts else ["no_route"]


def build_polygon_graph(route_parts_list: List[List[str]]) -> Dict[str, Set[str]]:
    graph: Dict[str, Set[str]] = {}
    for parts in route_parts_list:
        for node in parts:
            graph.setdefault(node, set())
        for i in range(len(parts) - 1):
            left, right = parts[i], parts[i + 1]
            graph[left].add(right)
            graph[right].add(left)
    return graph


def build_component_map(graph: Dict[str, Set[str]]) -> Dict[str, str]:
    component_map: Dict[str, str] = {}
    visited: Set[str] = set()
    for node in graph:
        if node in visited:
            continue
        stack = [node]
        members: List[str] = []
        while stack:
            cur = stack.pop()
            if cur in visited:
                continue
            visited.add(cur)
            members.append(cur)
            stack.extend(graph[cur] - visited)
        key = "|".join(sorted(members))
        for member in members:
            component_map[member] = key
    return component_map


def min_distance_to_run(orders: pd.DataFrame, candidate_idx: int, run_indices: List[int]) -> float:
    cand = orders.loc[candidate_idx]
    distances = [
        haversine(
            cand["latitude"],
            cand["longitude"],
            orders.loc[idx, "latitude"],
            orders.loc[idx, "longitude"],
        )
        for idx in run_indices
    ]
    return min(distances) if distances else float("inf")


def compute_run_distance_km(orders: pd.DataFrame, run_indices: List[int]) -> float:
    if len(run_indices) <= 1:
        return 0.0
    distance_km = 0.0
    for i in range(1, len(run_indices)):
        prev_idx = run_indices[i - 1]
        curr_idx = run_indices[i]
        distance_km += haversine(
            orders.loc[prev_idx, "latitude"],
            orders.loc[prev_idx, "longitude"],
            orders.loc[curr_idx, "latitude"],
            orders.loc[curr_idx, "longitude"],
        )
    return distance_km


def choose_seed_component(orders: pd.DataFrame) -> Optional[str]:
    pool = orders[(~orders["assigned"]) & (orders["quantity"] <= orders["capacity_limit"])]
    if pool.empty:
        return None
    stats = (
        pool.groupby("component_key", as_index=False)
        .agg(unassigned_qty=("quantity", "sum"))
        .sort_values("unassigned_qty", ascending=False)
    )
    return str(stats.iloc[0]["component_key"])


def choose_seed_order_idx(orders: pd.DataFrame, component_key: str, capacity_left: float) -> Optional[int]:
    component_unassigned = orders[
        (~orders["assigned"])
        & (orders["component_key"] == component_key)
        & (orders["quantity"] <= capacity_left)
    ].copy()
    if component_unassigned.empty:
        return None

    def local_qty_score(idx: int) -> float:
        center = component_unassigned.loc[idx]
        dists = component_unassigned.apply(
            lambda row: haversine(
                center["latitude"],
                center["longitude"],
                row["latitude"],
                row["longitude"],
            ),
            axis=1,
        )
        return float(component_unassigned.loc[dists <= SEED_NEIGHBOR_RADIUS_KM, "quantity"].sum())

    component_unassigned["local_qty_score"] = component_unassigned.index.map(local_qty_score)
    return int(
        component_unassigned.sort_values(
            ["local_qty_score", "quantity"],
            ascending=[False, False],
        ).index[0]
    )


def add_from_polygon(
    orders: pd.DataFrame,
    component_key: str,
    polygon_key: str,
    capacity_left: float,
    run_indices: List[int],
    capacity: float,
) -> Tuple[List[int], float]:
    added_indices: List[int] = []
    seed_idx = run_indices[0]
    current_idx = run_indices[-1]
    max_next_stop_km = INITIAL_MAX_NEXT_STOP_KM
    max_from_seed_km = INITIAL_MAX_FROM_SEED_KM

    while capacity_left > 0:
        utilization = ((capacity - capacity_left) / capacity) if capacity else 1.0
        if utilization >= TARGET_UTILIZATION:
            break

        candidates = orders[
            (~orders["assigned"])
            & (~orders.index.isin(run_indices))
            & (orders["component_key"] == component_key)
            & (orders["primary_polygon"] == polygon_key)
            & (orders["quantity"] <= capacity_left)
        ].copy()
        if candidates.empty:
            break

        current_lat = float(orders.loc[current_idx, "latitude"])
        current_lon = float(orders.loc[current_idx, "longitude"])
        seed_lat = float(orders.loc[seed_idx, "latitude"])
        seed_lon = float(orders.loc[seed_idx, "longitude"])

        candidates["leg_km"] = candidates.apply(
            lambda row: haversine(current_lat, current_lon, row["latitude"], row["longitude"]),
            axis=1,
        )
        candidates["distance_to_seed"] = candidates.apply(
            lambda row: haversine(seed_lat, seed_lon, row["latitude"], row["longitude"]),
            axis=1,
        )
        candidates["distance_to_run"] = candidates.index.map(
            lambda idx: min_distance_to_run(orders, idx, run_indices)
        )

        feasible = candidates[
            (candidates["leg_km"] <= max_next_stop_km)
            & (candidates["distance_to_seed"] <= max_from_seed_km)
        ].copy()
        if feasible.empty:
            feasible = candidates[
                (candidates["distance_to_run"] <= max_next_stop_km)
                & (candidates["distance_to_seed"] <= max_from_seed_km)
            ].copy()

        if feasible.empty:
            can_relax = (max_next_stop_km < MAX_NEXT_STOP_KM) or (max_from_seed_km < MAX_FROM_SEED_KM)
            if can_relax:
                max_next_stop_km = min(MAX_NEXT_STOP_KM, max_next_stop_km + RELAX_DISTANCE_STEP_KM)
                max_from_seed_km = min(MAX_FROM_SEED_KM, max_from_seed_km + RELAX_DISTANCE_STEP_KM)
                continue
            break

        next_idx = int(
            feasible.sort_values(
                ["leg_km", "distance_to_run", "quantity"],
                ascending=[True, True, False],
            ).index[0]
        )
        added_indices.append(next_idx)
        run_indices.append(next_idx)
        current_idx = next_idx
        capacity_left -= float(orders.loc[next_idx, "quantity"])

    return added_indices, capacity_left


def choose_nearest_polygon_in_component(
    orders: pd.DataFrame,
    component_key: str,
    used_polygons: Set[str],
    run_indices: List[int],
    capacity_left: float,
) -> Optional[str]:
    candidates = orders[
        (~orders["assigned"])
        & (~orders.index.isin(run_indices))
        & (orders["component_key"] == component_key)
        & (~orders["primary_polygon"].isin(used_polygons))
        & (orders["quantity"] <= capacity_left)
    ].copy()
    if candidates.empty:
        return None

    poly_stats = (
        candidates.groupby("primary_polygon", as_index=False)
        .agg(
            unassigned_qty=("quantity", "sum"),
            centroid_lat=("latitude", "mean"),
            centroid_lon=("longitude", "mean"),
        )
    )
    if poly_stats.empty:
        return None

    poly_stats["distance_to_run"] = poly_stats.apply(
        lambda row: min(
            haversine(
                row["centroid_lat"],
                row["centroid_lon"],
                orders.loc[idx, "latitude"],
                orders.loc[idx, "longitude"],
            )
            for idx in run_indices
        ),
        axis=1,
    )

    choice = poly_stats.sort_values(
        ["distance_to_run", "unassigned_qty"],
        ascending=[True, False],
    ).iloc[0]
    return str(choice["primary_polygon"])


def build_runs(df_sc: pd.DataFrame, capacity: float, sc_name: str, start_run_id: int):
    orders = (
        df_sc.groupby(
            ["order_id", "route", "primary_polygon", "component_key", "latitude", "longitude"],
            as_index=False,
        )
        .agg(quantity=("quantity", "sum"))
        .reset_index(drop=True)
    )

    orders["assigned"] = False
    orders["run_id"] = pd.NA
    orders["stop_sequence"] = pd.NA
    orders["capacity_limit"] = capacity

    too_large = orders["quantity"] > capacity
    if too_large.any():
        print(f"⚠️ {sc_name}: {int(too_large.sum())} orders exceed truck capacity and were skipped.")

    runs = []
    run_id = start_run_id

    while True:
        unassigned_fit = orders[(~orders["assigned"]) & (orders["quantity"] <= capacity)]
        if unassigned_fit.empty:
            break

        capacity_left = float(capacity)
        run_indices: List[int] = []

        seed_component = choose_seed_component(orders)
        if seed_component is None:
            break

        seed_idx = choose_seed_order_idx(orders, seed_component, capacity_left)
        if seed_idx is None:
            break

        run_indices.append(seed_idx)
        capacity_left -= float(orders.loc[seed_idx, "quantity"])
        used_polygons: Set[str] = {str(orders.loc[seed_idx, "primary_polygon"])}

        _, capacity_left = add_from_polygon(
            orders=orders,
            component_key=seed_component,
            polygon_key=str(orders.loc[seed_idx, "primary_polygon"]),
            capacity_left=capacity_left,
            run_indices=run_indices,
            capacity=capacity,
        )

        while capacity_left > 0:
            utilization = ((capacity - capacity_left) / capacity) if capacity else 1.0
            if utilization >= TARGET_UTILIZATION:
                break

            next_polygon = choose_nearest_polygon_in_component(
                orders=orders,
                component_key=seed_component,
                used_polygons=used_polygons,
                run_indices=run_indices,
                capacity_left=capacity_left,
            )
            if next_polygon is None:
                break

            used_polygons.add(next_polygon)
            added, capacity_left = add_from_polygon(
                orders=orders,
                component_key=seed_component,
                polygon_key=next_polygon,
                capacity_left=capacity_left,
                run_indices=run_indices,
                capacity=capacity,
            )
            if not added:
                break

        orders.loc[run_indices, "assigned"] = True
        orders.loc[run_indices, "run_id"] = run_id
        for seq, idx in enumerate(run_indices, start=1):
            orders.loc[idx, "stop_sequence"] = seq

        total_load = float(orders.loc[run_indices, "quantity"].sum())
        utilization = (total_load / capacity) if capacity else 0.0
        runs.append(
            {
                "run_id": run_id,
                "supply_chain": sc_name,
                "route_component": seed_component,
                "polygons_covered": ", ".join(sorted(used_polygons)),
                "capacity": capacity,
                "load": total_load,
                "utilization": round(utilization, 2),
                "stops": len(run_indices),
                "run_distance_km": round(compute_run_distance_km(orders, run_indices), 2),
                "under_utilized": utilization < TARGET_UTILIZATION,
            }
        )
        run_id += 1

    return orders, runs, run_id


def validate_columns(df: pd.DataFrame):
    required = {"order_id", "quantity", "latitude", "longitude", "route", "supply_chain"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")


def prepare_route_components(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["route_parts"] = df["route"].map(parse_route_parts)
    df["primary_polygon"] = df["route_parts"].map(lambda parts: parts[0] if parts else "no_route")

    graph = build_polygon_graph(df["route_parts"].tolist())
    component_map = build_component_map(graph)
    df["component_key"] = df["primary_polygon"].map(lambda p: component_map.get(p, p))
    return df


def main():
    df = pd.read_excel(FILE_PATH)
    df.columns = df.columns.str.lower().str.strip()
    validate_columns(df)

    df["supply_chain"] = df["supply_chain"].astype(str).str.lower().str.strip()
    df["route"] = df["route"].astype(str).str.strip().str.lower()
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(0)
    df["latitude"] = pd.to_numeric(df["latitude"], errors="coerce")
    df["longitude"] = pd.to_numeric(df["longitude"], errors="coerce")
    df = df.dropna(subset=["latitude", "longitude"])

    df = prepare_route_components(df)

    all_orders = []
    all_runs = []
    run_id = 1

    for sc_name, capacity in SUPPLY_CHAINS.items():
        df_sc = df[df["supply_chain"] == sc_name].copy()
        print(f"🚚 {sc_name}: {df_sc['order_id'].nunique()} unique orders")
        if df_sc.empty:
            continue

        orders, runs, run_id = build_runs(df_sc, capacity, sc_name, run_id)
        orders["supply_chain"] = sc_name
        all_orders.append(orders)
        all_runs.extend(runs)

    if not all_runs:
        raise ValueError("❌ No runs generated — check supply_chain and quantity values.")

    final_orders = pd.concat(all_orders, ignore_index=True)[
        [
            "order_id",
            "supply_chain",
            "component_key",
            "primary_polygon",
            "latitude",
            "longitude",
            "run_id",
            "stop_sequence",
        ]
    ]

    detailed_df = df.merge(
        final_orders,
        on=[
            "order_id",
            "supply_chain",
            "component_key",
            "primary_polygon",
            "latitude",
            "longitude",
        ],
        how="left",
    )
    runs_df = pd.DataFrame(all_runs)

    output_path = FILE_PATH.replace(".xlsx", "_final_routes.xlsx")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        detailed_df.to_excel(writer, sheet_name="Detailed Runs", index=False)
        runs_df.to_excel(writer, sheet_name="Summary", index=False)

    print("✅ DONE")
    print(f"🚚 Runs: {len(runs_df)}")
    print(f"💾 File: {output_path}")


if __name__ == "__main__":
    main()
