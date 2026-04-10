import pandas as pd
from math import radians, cos, sin, asin, sqrt
from typing import List, Tuple, Optional

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
) -> Tuple[List[int], float]:
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

        # Pick the nearest retailer to any already planned stop in this run.
        candidates["distance_to_run"] = candidates.index.map(
            lambda idx: min_distance_to_run(orders, idx, run_indices)
        )
        next_idx = int(
            candidates.sort_values(["distance_to_run", "quantity"], ascending=[True, False]).index[0]
        )
        row = orders.loc[next_idx]
        added_indices.append(next_idx)
        run_indices.append(next_idx)
        capacity_left -= float(row["quantity"])

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

        seed_route = choose_seed_route(orders[(~orders["assigned"]) & (orders["quantity"] <= capacity)])
        if seed_route is None:
            break

        seed_idx = choose_seed_order_idx(orders, seed_route, capacity_left)
        if seed_idx is None:
            # No fittable order in that route (defensive).
            orders.loc[orders["route_key"] == seed_route, "assigned"] = True
            continue

        run_indices.append(seed_idx)
        capacity_left -= float(orders.loc[seed_idx, "quantity"])

        # Strict rule: fill only from the same route polygon as the run seed.
        _, capacity_left = add_from_specific_route(
            orders=orders,
            route_key=seed_route,
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
        run_distance_km = compute_run_distance_km(orders, run_indices)
        utilization = (total_load / capacity) if capacity else 0
        runs.append(
            {
                "run_id": run_id,
                "supply_chain": sc_name,
                "route_polygon": seed_route,
                "capacity": capacity,
                "load": total_load,
                "utilization": round(utilization, 2),
                "stops": len(run_indices),
                "run_distance_km": round(run_distance_km, 2),
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
