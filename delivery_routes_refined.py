import pandas as pd
from math import radians, cos, sin, asin, sqrt
from typing import List, Tuple, Optional
import re

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

# Compact-cluster controls (same route only) to reduce fuel burn.
# Aggressive mode: tighter clusters, fewer long jumps.
SEED_NEIGHBOR_RADIUS_KM = 1.0
INITIAL_MAX_NEXT_STOP_KM = 1.0
INITIAL_MAX_FROM_SEED_KM = 2.0
RELAX_DISTANCE_STEP_KM = 0.4
MAX_NEXT_STOP_KM = 2.5
MAX_FROM_SEED_KM = 3.5


# =============================
# DISTANCE
# =============================
def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    return 6371 * 2 * asin(sqrt(a))


def canonical_route_key(route_value: str) -> str:
    # If a route value encodes connected polygons (e.g. "A > B > C"),
    # use a canonical grouped key so planning stays inside that connected group.
    text = str(route_value).strip().lower()
    if text == "":
        return "no_route"
    parts = [p.strip() for p in re.split(r"\s*(?:->|>|,|;|\||/)\s*", text) if p.strip()]
    if not parts:
        return "no_route"
    return "|".join(sorted(set(parts)))


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
    route_unassigned = orders[
        (~orders["assigned"])
        & (orders["route_key"] == route_key)
    ].copy()
    candidates = route_unassigned[route_unassigned["quantity"] <= capacity_left].copy()

    if candidates.empty:
        return None

    # Seed from a dense local pocket first (better ground-level route compactness).
    def local_qty_score(idx: int) -> float:
        center = route_unassigned.loc[idx]
        dists = route_unassigned.apply(
            lambda row: haversine(
                center["latitude"], center["longitude"], row["latitude"], row["longitude"]
            ),
            axis=1,
        )
        return float(route_unassigned.loc[dists <= SEED_NEIGHBOR_RADIUS_KM, "quantity"].sum())

    candidates["local_qty_score"] = candidates.index.map(local_qty_score)
    return int(
        candidates.sort_values(["local_qty_score", "quantity"], ascending=[False, False]).index[0]
    )


def add_from_specific_route(
    orders: pd.DataFrame,
    route_key: str,
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
        current_utilization = ((capacity - capacity_left) / capacity) if capacity else 1
        if current_utilization >= TARGET_UTILIZATION:
            break

        candidates = orders[
            (~orders["assigned"])
            & (orders["route_key"] == route_key)
            & (~orders.index.isin(run_indices))
            & (orders["quantity"] <= capacity_left)
        ].copy()

        if candidates.empty:
            break

        current_lat = float(orders.loc[current_idx, "latitude"])
        current_lon = float(orders.loc[current_idx, "longitude"])
        seed_lat = float(orders.loc[seed_idx, "latitude"])
        seed_lon = float(orders.loc[seed_idx, "longitude"])

        # Compact run logic:
        # - short next leg from current stop
        # - bounded spread from the seed stop
        # - still aware of nearest distance to whole run
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
            # If current stop has no good neighbor, allow nearest to any stop in the same run.
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
                ["leg_km", "distance_to_run", "quantity"], ascending=[True, True, False]
            ).index[0]
        )
        row = orders.loc[next_idx]
        added_indices.append(next_idx)
        run_indices.append(next_idx)
        current_idx = next_idx
        capacity_left -= float(row["quantity"])

    return added_indices, capacity_left


# =============================
# ROUTING ENGINE
# =============================
def build_runs(df_sc: pd.DataFrame, capacity: float, sc_name: str, start_run_id: int):
    orders = (
        df_sc.groupby(["order_id", "route", "route_key", "latitude", "longitude"], as_index=False)
        .agg(
            quantity=("quantity", "sum"),
        )
        .reset_index(drop=True)
    )

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
            capacity=capacity,
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
    df["route_key"] = df["route"].map(canonical_route_key)
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
    final_orders = final_orders[
        ["order_id", "supply_chain", "route_key", "latitude", "longitude", "run_id", "stop_sequence"]
    ]

    detailed_df = df.merge(
        final_orders,
        on=["order_id", "supply_chain", "route_key", "latitude", "longitude"],
        how="left",
    )
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
