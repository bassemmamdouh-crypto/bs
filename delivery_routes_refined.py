import pandas as pd
from math import radians, cos, sin, asin, sqrt
from typing import List, Tuple, Set, Optional

# =============================
# SETTINGS
# =============================
FILE_PATH = r"E:\Marbah Products\Other Scripts\Delivery Plan\orders_in_delivery.xlsx"

SUPPLY_CHAINS = {
    "lays": 470,
    "pepsi": 790,
}

# Prefer filling runs as much as possible, but keep route logic first.
TARGET_UTILIZATION = 0.90

# Optional hard limit for route (polygon) merge distance.
# Set to None to allow any distance when no closer orders can fit.
MAX_ROUTE_MERGE_DISTANCE_KM = None


# =============================
# DISTANCE
# =============================
def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    return 6371 * 2 * asin(sqrt(a))


def min_distance_to_run(orders: pd.DataFrame, candidate_idx: int, run_indices: List[int]) -> float:
    cand = orders.loc[candidate_idx]
    distances = [
        haversine(
            cand["latitude"],
            cand["longitude"],
            orders.loc[i, "latitude"],
            orders.loc[i, "longitude"],
        )
        for i in run_indices
    ]
    return min(distances) if distances else float("inf")


def compute_route_stats(orders: pd.DataFrame) -> pd.DataFrame:
    route_unassigned = orders[~orders["assigned"]].copy()
    if route_unassigned.empty:
        return pd.DataFrame(columns=["route_key", "unassigned_qty", "centroid_lat", "centroid_lon"])

    stats = (
        route_unassigned.groupby("route_key", dropna=False)
        .agg(
            unassigned_qty=("quantity", "sum"),
            centroid_lat=("latitude", "mean"),
            centroid_lon=("longitude", "mean"),
        )
        .reset_index()
    )
    return stats


def choose_seed_route(orders: pd.DataFrame) -> Optional[str]:
    stats = compute_route_stats(orders)
    if stats.empty:
        return None
    return stats.sort_values(["unassigned_qty"], ascending=False).iloc[0]["route_key"]


def choose_seed_order_idx(orders: pd.DataFrame, route_key: str, capacity_left: float) -> Optional[int]:
    candidates = orders[
        (~orders["assigned"])
        & (orders["route_key"] == route_key)
        & (orders["quantity"] <= capacity_left)
    ].copy()
    if candidates.empty:
        return None
    # Start with biggest order in the route to maximize run utilization.
    return int(candidates.sort_values("quantity", ascending=False).index[0])


def add_from_specific_route(
    orders: pd.DataFrame,
    route_key: str,
    capacity_left: float,
    run_indices: List[int],
    current_lat: float,
    current_lon: float,
) -> Tuple[List[int], float, float, float]:
    added_indices: List[int] = []

    while capacity_left > 0:
        candidates = orders[
            (~orders["assigned"])
            & (orders["route_key"] == route_key)
            & (~orders.index.isin(run_indices))
            & (orders["quantity"] <= capacity_left)
        ].copy()

        if candidates.empty:
            break

        candidates["distance"] = candidates.apply(
            lambda x: haversine(current_lat, current_lon, x["latitude"], x["longitude"]),
            axis=1,
        )
        next_idx = int(candidates.sort_values(["distance", "quantity"], ascending=[True, False]).index[0])
        row = orders.loc[next_idx]
        added_indices.append(next_idx)
        run_indices.append(next_idx)
        capacity_left -= float(row["quantity"])
        current_lat, current_lon = float(row["latitude"]), float(row["longitude"])

    return added_indices, capacity_left, current_lat, current_lon


def choose_next_route_to_merge(
    orders: pd.DataFrame,
    active_routes: Set[str],
    run_indices: List[int],
    capacity_left: float,
) -> Optional[str]:
    stats = compute_route_stats(orders)
    if stats.empty:
        return None

    stats = stats[~stats["route_key"].isin(active_routes)].copy()
    if stats.empty:
        return None

    # Keep only routes with at least one order that can still fit.
    fit_routes = (
        orders[
            (~orders["assigned"])
            & (orders["quantity"] <= capacity_left)
            & (~orders["route_key"].isin(active_routes))
        ]["route_key"]
        .drop_duplicates()
        .tolist()
    )
    if not fit_routes:
        return None
    stats = stats[stats["route_key"].isin(fit_routes)].copy()
    if stats.empty:
        return None

    def route_distance(row: pd.Series) -> float:
        distances = [
            haversine(
                row["centroid_lat"],
                row["centroid_lon"],
                orders.loc[i, "latitude"],
                orders.loc[i, "longitude"],
            )
            for i in run_indices
        ]
        return min(distances) if distances else float("inf")

    stats["distance_to_run"] = stats.apply(route_distance, axis=1)
    stats = stats.sort_values(["distance_to_run", "unassigned_qty"], ascending=[True, False])

    selected = stats.iloc[0]
    if MAX_ROUTE_MERGE_DISTANCE_KM is not None and selected["distance_to_run"] > MAX_ROUTE_MERGE_DISTANCE_KM:
        return None
    return selected["route_key"]


def add_global_nearest_backfill(
    orders: pd.DataFrame,
    capacity_left: float,
    run_indices: List[int],
) -> Tuple[List[int], float]:
    added_indices: List[int] = []

    while capacity_left > 0:
        candidates = orders[
            (~orders["assigned"])
            & (~orders.index.isin(run_indices))
            & (orders["quantity"] <= capacity_left)
        ].copy()
        if candidates.empty:
            break

        candidates["distance"] = candidates.index.map(
            lambda idx: min_distance_to_run(orders, idx, run_indices)
        )
        next_idx = int(candidates.sort_values(["distance", "quantity"], ascending=[True, False]).index[0])

        run_indices.append(next_idx)
        added_indices.append(next_idx)
        capacity_left -= float(orders.loc[next_idx, "quantity"])

    return added_indices, capacity_left


# =============================
# ROUTING ENGINE
# =============================
def build_runs(df_sc: pd.DataFrame, capacity: float, sc_name: str, start_run_id: int):
    orders = (
        df_sc.groupby("order_id", as_index=False)
        .agg(
            quantity=("quantity", "sum"),
            latitude=("latitude", "first"),
            longitude=("longitude", "first"),
            route=("route", "first"),
        )
        .reset_index(drop=True)
    )

    orders["route_key"] = orders["route"].astype(str).str.strip().str.lower().replace({"": "no_route"})
    orders["assigned"] = False
    orders["run_id"] = pd.NA
    orders["stop_sequence"] = pd.NA

    runs = []
    run_id = start_run_id

    too_large = orders["quantity"] > capacity
    if too_large.any():
        skipped_count = int(too_large.sum())
        print(f"⚠️ {sc_name}: {skipped_count} orders exceed truck capacity and were skipped.")

    while True:
        unassigned_fit = orders[(~orders["assigned"]) & (orders["quantity"] <= capacity)]
        if unassigned_fit.empty:
            break

        capacity_left = float(capacity)
        run_indices: List[int] = []
        active_routes: Set[str] = set()

        seed_route = choose_seed_route(orders[(~orders["assigned"]) & (orders["quantity"] <= capacity)])
        if seed_route is None:
            break

        seed_idx = choose_seed_order_idx(orders, seed_route, capacity_left)
        if seed_idx is None:
            # No fittable order in that route (defensive).
            orders.loc[orders["route_key"] == seed_route, "assigned"] = True
            continue

        run_indices.append(seed_idx)
        active_routes.add(seed_route)
        capacity_left -= float(orders.loc[seed_idx, "quantity"])
        current_lat = float(orders.loc[seed_idx, "latitude"])
        current_lon = float(orders.loc[seed_idx, "longitude"])

        # 1) Fill from seed route first (polygon-by-polygon behavior).
        _, capacity_left, current_lat, current_lon = add_from_specific_route(
            orders=orders,
            route_key=seed_route,
            capacity_left=capacity_left,
            run_indices=run_indices,
            current_lat=current_lat,
            current_lon=current_lon,
        )

        # 2) Merge nearest additional routes while capacity remains.
        while capacity_left > 0:
            next_route = choose_next_route_to_merge(
                orders=orders,
                active_routes=active_routes,
                run_indices=run_indices,
                capacity_left=capacity_left,
            )
            if next_route is None:
                break

            active_routes.add(next_route)
            added_from_route, capacity_left, current_lat, current_lon = add_from_specific_route(
                orders=orders,
                route_key=next_route,
                capacity_left=capacity_left,
                run_indices=run_indices,
                current_lat=current_lat,
                current_lon=current_lon,
            )

            if not added_from_route:
                # Prevent endless cycling on a route with no fitting remaining orders.
                break

            current_util = (capacity - capacity_left) / capacity if capacity else 0
            if current_util >= TARGET_UTILIZATION:
                # Utilization is healthy; still do global nearest fill afterwards.
                break

        # 3) Backfill from all remaining unassigned orders by nearest distance to current run.
        _, capacity_left = add_global_nearest_backfill(
            orders=orders,
            capacity_left=capacity_left,
            run_indices=run_indices,
        )

        if not run_indices:
            break

        orders.loc[run_indices, "assigned"] = True
        orders.loc[run_indices, "run_id"] = run_id
        for stop_no, idx in enumerate(run_indices, start=1):
            orders.loc[idx, "stop_sequence"] = stop_no

        total_load = float(orders.loc[run_indices, "quantity"].sum())
        routes_covered = (
            orders.loc[run_indices, "route_key"].dropna().astype(str).drop_duplicates().tolist()
        )
        runs.append(
            {
                "run_id": run_id,
                "supply_chain": sc_name,
                "capacity": capacity,
                "load": total_load,
                "utilization": round(total_load / capacity, 2) if capacity else 0,
                "stops": len(run_indices),
                "routes_covered": ", ".join(routes_covered),
            }
        )

        run_id += 1

    return orders, runs, run_id


def validate_columns(df: pd.DataFrame):
    required = {"order_id", "quantity", "latitude", "longitude", "route", "supply_chain"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")


def main():
    # =============================
    # LOAD
    # =============================
    df = pd.read_excel(FILE_PATH)
    df.columns = df.columns.str.lower().str.strip()
    validate_columns(df)

    # Normalize key columns.
    df["supply_chain"] = df["supply_chain"].astype(str).str.lower().str.strip()
    df["route"] = df["route"].astype(str).str.strip()
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(0)
    df["latitude"] = pd.to_numeric(df["latitude"], errors="coerce")
    df["longitude"] = pd.to_numeric(df["longitude"], errors="coerce")
    df = df.dropna(subset=["latitude", "longitude"])

    # =============================
    # RUN PER SUPPLY CHAIN
    # =============================
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

    # =============================
    # SAFETY CHECK
    # =============================
    if len(all_runs) == 0:
        raise ValueError("❌ No runs generated — check supply_chain and quantity values.")

    final_orders = pd.concat(all_orders, ignore_index=True)
    final_orders = final_orders[["order_id", "run_id", "stop_sequence"]]

    detailed_df = df.merge(final_orders, on="order_id", how="left")
    runs_df = pd.DataFrame(all_runs)

    # =============================
    # OUTPUT
    # =============================
    output_path = FILE_PATH.replace(".xlsx", "_final_routes.xlsx")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        detailed_df.to_excel(writer, sheet_name="Detailed Runs", index=False)
        runs_df.to_excel(writer, sheet_name="Summary", index=False)

    print("✅ DONE")
    print(f"🚚 Runs: {len(runs_df)}")
    print(f"💾 File: {output_path}")


if __name__ == "__main__":
    main()
