import argparse
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


MAX_PRODUCTS = 300
INPUT_SHEET = "Input_Data"
SCORING_SHEET = "Scoring_Model"
BUNDLE_SHEET = "All_Possible_Bundles"
TOP_BUNDLE_SHEET = "Top_30_Bundles"

# KEEP ONLY TOP N STRONGEST PRODUCTS AS ANCHORS
MAX_ANCHORS = 7

# KEEP ONLY THE TOP N BUNDLES (BY BUNDLE PRIORITY SCORE)
TOP_N_BUNDLES = 30

# Prefer same-category anchors when choosing the single best anchor
SAME_CATEGORY_BONUS = 0.10

# Soft penalty so candidates spread across anchors instead of all
# locking onto the single strongest product
ANCHOR_LOAD_PENALTY = 0.03


def to_number(value, default=0.0):
    if value is None or value == "":
        return default

    try:
        return float(value)

    except (TypeError, ValueError):
        return default


def normalize_percent(value):

    number = to_number(value, 0.0)

    return number / 100.0 if number > 1 else number


def percent_rank_inc(values, target):

    sorted_vals = sorted(values)

    n = len(sorted_vals)

    if n <= 1:
        return 0.0

    if target <= sorted_vals[0]:
        return 0.0

    if target >= sorted_vals[-1]:
        return 1.0

    for idx, val in enumerate(sorted_vals):

        if val == target:

            first = idx
            last = idx

            while first - 1 >= 0 and sorted_vals[first - 1] == target:
                first -= 1

            while last + 1 < n and sorted_vals[last + 1] == target:
                last += 1

            return ((first + last) / 2.0) / (n - 1)

    for idx in range(1, n):

        lower = sorted_vals[idx - 1]
        upper = sorted_vals[idx]

        if lower < target < upper:

            frac = (target - lower) / (upper - lower)

            return ((idx - 1) + frac) / (n - 1)

    return 0.0


def read_products(ws):

    products = []

    for row in range(2, MAX_PRODUCTS + 2):

        name = ws[f"A{row}"].value
        category = ws[f"B{row}"].value
        product_id = ws[f"C{row}"].value

        if not product_id:
            continue

        qty_m1 = to_number(ws[f"E{row}"].value)
        qty_m2 = to_number(ws[f"F{row}"].value)
        qty_m3 = to_number(ws[f"G{row}"].value)
        qty_m4 = to_number(ws[f"H{row}"].value)

        sold_qty_4m = qty_m1 + qty_m2 + qty_m3 + qty_m4

        avg_qty_4m = (
            sold_qty_4m / 4
            if sold_qty_4m > 0
            else 0.0
        )

        stock = to_number(ws[f"J{row}"].value)

        reserved_stock = to_number(ws[f"K{row}"].value)

        available_stock = to_number(
            ws[f"L{row}"].value,
            stock - reserved_stock
        )

        if available_stock == 0 and stock > 0:
            available_stock = stock - reserved_stock

        products.append(
            {
                "product_id": str(product_id),
                "name": str(name) if name else "",
                "category": str(category) if category else "",

                "purchased_item_count": to_number(
                    ws[f"D{row}"].value
                ),

                "sold_qty_last_4m": sold_qty_4m,

                "avg_qty_4m": avg_qty_4m,

                "contribution": normalize_percent(
                    ws[f"I{row}"].value
                ),

                "stock": stock,

                "reserved_stock": reserved_stock,

                "available_stock": max(
                    available_stock,
                    0.0
                ),
            }
        )

    return products


def score_products(products):

    sold_values = [
        p["sold_qty_last_4m"]
        for p in products
    ]

    contribution_values = [
        p["contribution"]
        for p in products
    ]

    coverage_values = []

    for product in products:

        avg = product["avg_qty_4m"]

        coverage = (
            99.0
            if avg == 0
            else product["available_stock"] / avg
        )

        coverage_values.append(coverage)

    for idx, product in enumerate(products):

        movement_percentile = percent_rank_inc(
            sold_values,
            product["sold_qty_last_4m"]
        )

        slow_mover_score = 1.0 - movement_percentile

        contribution_percentile = percent_rank_inc(
            contribution_values,
            product["contribution"]
        )

        stock_coverage = coverage_values[idx]

        stock_pressure_percentile = percent_rank_inc(
            coverage_values,
            stock_coverage
        )

        priority_score = (
            0.45 * slow_mover_score
            + 0.35 * stock_pressure_percentile
            + 0.20 * contribution_percentile
        )

        # MOVEMENT BAND
        if movement_percentile <= 0.33:
            movement_band = "Low Movement"

        elif movement_percentile <= 0.66:
            movement_band = "Medium Movement"

        else:
            movement_band = "High Movement"

        # CLUSTER
        if priority_score >= 0.67:

            cluster = "High Value Bundle"

            base_discount = 0.05

        elif priority_score >= 0.34:

            cluster = "Medium Value Bundle"

            base_discount = 0.07

        else:

            cluster = "Low Value Bundle"

            base_discount = 0.10

        discount = min(
            0.12,
            base_discount + (
                0.02 if stock_coverage >= 8 else 0.0
            )
        )

        # STRONGEST PRODUCTS ONLY
        anchor_eligible = (
            movement_percentile >= 0.85
            and contribution_percentile >= 0.70
            and product["available_stock"] > 0
        )

        anchor_strength = (
            0.6 * movement_percentile
            + 0.4 * contribution_percentile
        )

        product.update(
            {
                "movement_percentile": movement_percentile,

                "slow_mover_score": slow_mover_score,

                "contribution_percentile": contribution_percentile,

                "stock_coverage": stock_coverage,

                "stock_pressure_percentile": stock_pressure_percentile,

                "priority_score": priority_score,

                "movement_band": movement_band,

                "cluster": cluster,

                "discount": discount,

                "anchor_eligible": anchor_eligible,

                "anchor_strength": anchor_strength,
            }
        )

    # KEEP ONLY TOP ANCHORS
    sorted_anchors = sorted(
        [p for p in products if p["anchor_eligible"]],
        key=lambda x: x["anchor_strength"],
        reverse=True
    )

    top_anchor_ids = {
        p["product_id"]
        for p in sorted_anchors[:MAX_ANCHORS]
    }

    # FINALIZE ELIGIBILITY
    for product in products:

        product["anchor_eligible"] = (
            product["product_id"] in top_anchor_ids
        )

        # EVERY NON-ANCHOR PRODUCT CAN BE BUNDLED
        product["candidate_eligible"] = (
            not product["anchor_eligible"]
            and product["available_stock"] > 0
        )


def update_scoring_sheet(ws, products):

    product_map = {
        p["product_id"]: p
        for p in products
    }

    for row in range(2, MAX_PRODUCTS + 2):

        product_id = ws[f"A{row}"].value

        if not product_id:
            continue

        product = product_map.get(str(product_id))

        if not product:
            continue

        ws[f"E{row}"] = product["sold_qty_last_4m"]
        ws[f"F{row}"] = product["avg_qty_4m"]
        ws[f"G{row}"] = product["movement_percentile"]
        ws[f"H{row}"] = product["slow_mover_score"]
        ws[f"I{row}"] = product["contribution"]
        ws[f"J{row}"] = product["contribution_percentile"]
        ws[f"K{row}"] = product["stock"]
        ws[f"L{row}"] = product["reserved_stock"]
        ws[f"M{row}"] = product["available_stock"]

        ws[f"N{row}"] = (
            product["stock"]
            - product["reserved_stock"]
        )

        ws[f"O{row}"] = (
            "OK"
            if abs(
                product["available_stock"]
                - (
                    product["stock"]
                    - product["reserved_stock"]
                )
            ) <= 1
            else "CHECK"
        )

        ws[f"P{row}"] = product["stock_coverage"]
        ws[f"Q{row}"] = product["stock_pressure_percentile"]
        ws[f"R{row}"] = product["priority_score"]
        ws[f"S{row}"] = product["movement_band"]
        ws[f"T{row}"] = product["cluster"]

        ws[f"U{row}"] = (
            "Yes"
            if product["anchor_eligible"]
            else "No"
        )

        ws[f"V{row}"] = (
            "Yes"
            if product["candidate_eligible"]
            else "No"
        )

        ws[f"W{row}"] = product["discount"]

        ws[f"X{row}"] = (
            f'{product["category"]}|anchor'
            if product["anchor_eligible"]
            else ""
        )

        ws[f"Y{row}"] = (
            product["name"]
            if product["anchor_eligible"]
            else ""
        )

        for col in (
            "G",
            "H",
            "I",
            "J",
            "Q",
            "R",
            "W"
        ):
            ws[f"{col}{row}"].number_format = "0.00%"


def style_header(ws):

    header_fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )

    for cell in ws[1]:

        cell.font = Font(
            color="FFFFFF",
            bold=True
        )

        cell.fill = header_fill

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )


def set_widths(ws, widths):

    for col_idx, width in widths.items():

        ws.column_dimensions[
            get_column_letter(col_idx)
        ].width = width


def pair_score(candidate, anchor):
    """Score one candidate against one anchor (higher is better)."""

    score = (
        0.75 * candidate["priority_score"]
        + 0.25 * anchor["anchor_strength"]
    )

    if candidate["category"] and (
        candidate["category"] == anchor["category"]
    ):
        score += SAME_CATEGORY_BONUS

    return score


def build_candidate_bundles(products):
    """Assign each candidate product to exactly one best anchor.

    Creates one 2-product bundle per candidate. With 7 anchors and 30+
    eligible candidates this yields 30+ bundles, and no product is paired
    with more than one anchor.
    """

    anchors = [
        p for p in products
        if p["anchor_eligible"]
    ]

    candidates = [
        p for p in products
        if p["candidate_eligible"]
    ]

    anchors.sort(
        key=lambda x: x["anchor_strength"],
        reverse=True
    )

    candidates.sort(
        key=lambda x: x["priority_score"],
        reverse=True
    )

    if not anchors:
        return []

    bundles = []
    anchor_loads = {
        anchor["product_id"]: 0
        for anchor in anchors
    }

    # ==========================================
    # ONE BUNDLE PER CANDIDATE → ONE BEST ANCHOR
    # ==========================================
    for candidate in candidates:

        best_anchor = None
        best_score = None
        best_adjusted = None

        for anchor in anchors:

            if anchor["product_id"] == candidate["product_id"]:
                continue

            score = pair_score(candidate, anchor)

            # Prefer stronger / same-category matches, but spread load
            # across anchors so 7 anchors can support 30+ bundles.
            adjusted = (
                score
                - ANCHOR_LOAD_PENALTY
                * anchor_loads[anchor["product_id"]]
            )

            if (
                best_adjusted is None
                or adjusted > best_adjusted
            ):
                best_adjusted = adjusted
                best_score = score
                best_anchor = anchor

        if best_anchor is None:
            continue

        anchor_loads[best_anchor["product_id"]] += 1

        category_mix = (
            "Same Category"
            if best_anchor["category"] == candidate["category"]
            else "Cross Category"
        )

        bundles.append(
            {
                "score": best_score,
                "anchor_id": best_anchor["product_id"],
                "candidate_ids": {
                    candidate["product_id"],
                },
                "row": [
                    2,

                    best_anchor["product_id"],
                    best_anchor["name"],

                    candidate["product_id"],
                    candidate["name"],

                    "",
                    "",

                    category_mix,

                    best_score,

                    candidate["discount"],

                    "Moves one medium/slow mover using one strong anchor",
                ],
            }
        )

    return bundles


def write_bundle_sheet(ws, bundles):
    """Write the given (already ranked) bundles to a worksheet.

    ``bundles`` is a list of dicts as produced by ``build_candidate_bundles``.
    Rows are re-numbered sequentially (``B-000001``, ``B-000002``, ...).
    """

    headers = [
        "bundle_id",
        "bundle_size",
        "anchor_product_id",
        "anchor_product_name",
        "item_2_product_id",
        "item_2_product_name",
        "item_3_product_id",
        "item_3_product_name",
        "category_mix",
        "bundle_priority_score",
        "suggested_discount_%",
        "reason",
    ]

    merged_ranges = list(
        ws.merged_cells.ranges
    )

    for cell_range in merged_ranges:
        ws.unmerge_cells(str(cell_range))

    ws.delete_rows(1, ws.max_row)

    ws.append(headers)

    style_header(ws)

    for bundle_id, bundle in enumerate(bundles, start=1):

        ws.append(
            [f"B-{bundle_id:06d}"] + bundle["row"]
        )

    for row in range(2, ws.max_row + 1):

        ws[f"J{row}"].number_format = "0.00%"

        ws[f"K{row}"].number_format = "0.00%"

    ws.freeze_panes = "A2"

    ws.auto_filter.ref = f"A1:L{ws.max_row}"

    set_widths(
        ws,
        {
            1: 14,
            2: 12,
            3: 16,
            4: 30,
            5: 16,
            6: 30,
            7: 16,
            8: 30,
            9: 20,
            10: 20,
            11: 18,
            12: 56,
        },
    )

    return len(bundles)


def ensure_sheet(workbook, name):

    if name in workbook.sheetnames:
        return workbook[name]

    return workbook.create_sheet(name)


def main():

    parser = argparse.ArgumentParser(
        description="Generate the top bundles"
    )

    parser.add_argument(
        "workbook_path",
        nargs="?",
        default="bundle_planning_template.xlsx",
        help="Path to workbook"
    )

    # FIX FOR JUPYTER
    args, unknown = parser.parse_known_args()

    workbook_path = Path(args.workbook_path)

    print(f"Loading workbook: {workbook_path}")

    wb = load_workbook(workbook_path)

    input_ws = wb[INPUT_SHEET]

    products = read_products(input_ws)

    if not products:
        raise ValueError(
            "No products found in Input_Data"
        )

    print(f"Products loaded: {len(products)}")

    score_products(products)

    scoring_ws = ensure_sheet(
        wb,
        SCORING_SHEET
    )

    update_scoring_sheet(
        scoring_ws,
        products
    )

    # ONE BEST ANCHOR PER CANDIDATE → ONE BUNDLE PER PRODUCT
    bundles = build_candidate_bundles(products)

    bundles.sort(
        key=lambda b: b["score"],
        reverse=True
    )

    top_bundles = bundles[:TOP_N_BUNDLES]

    # TAB 1: ALL ASSIGNED BUNDLES (1 ANCHOR PER PRODUCT)
    all_bundles_ws = ensure_sheet(
        wb,
        BUNDLE_SHEET
    )

    total_bundles = write_bundle_sheet(
        all_bundles_ws,
        bundles
    )

    # TAB 2: TOP N BUNDLES
    top_bundles_ws = ensure_sheet(
        wb,
        TOP_BUNDLE_SHEET
    )

    kept_bundles = write_bundle_sheet(
        top_bundles_ws,
        top_bundles
    )

    wb.save(workbook_path)

    print("\nWorkbook updated successfully.")

    print(
        f"Saved to: "
        f"{workbook_path.resolve()}"
    )

    print(f"Products: {len(products)}")

    print(
        f"Anchors: "
        f"{sum(1 for p in products if p['anchor_eligible'])}"
    )

    print(
        f"Candidates: "
        f"{sum(1 for p in products if p['candidate_eligible'])}"
    )

    print(
        f"'{BUNDLE_SHEET}' tab: {total_bundles} bundles "
        f"(one best anchor per product)"
    )

    print(
        f"'{TOP_BUNDLE_SHEET}' tab: {kept_bundles} bundles "
        f"(top {TOP_N_BUNDLES} by priority score)"
    )


if __name__ == "__main__":
    main()
