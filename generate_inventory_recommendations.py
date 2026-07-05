"""
generate_inventory_recommendations.py
Suggests two kinds of inventory moves, using output_company.parquet:

1. WAREHOUSE -> STORE replenishment
   Stores that are selling well but running low on stock, where the
   AKL warehouse (or WH1) has spare units sitting idle.

2. STORE -> STORE transfers
   A store with slow-moving excess stock of a SKU, paired with another
   store that sells that SKU well but is low/out of stock.

Both suggestions respect store limits where available, and intentionally
DO NOT touch dead stock (zero sales everywhere) — that's a separate
clearance decision, not a transfer decision.

Usage:
    python generate_inventory_recommendations.py
    python generate_inventory_recommendations.py --data-dir D:\\PopStop_App\\popstop-dashboard\\data
"""

import os, sys, json, argparse
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np

# ─── CONFIG ────────────────────────────────────────────────────────────────────
STORES = ["DUN", "PAP", "QG", "RICC", "RICH", "SP", "TR"]
STORE_LABELS = {
    "DUN": "Dunedin", "PAP": "Papanui", "QG": "Queensgate",
    "RICC": "Riccarton", "RICH": "Richmond", "SP": "Sylvia Park", "TR": "Te Rapa",
}

# Replenishment thresholds
LOW_STOCK_WOH_THRESHOLD   = 2.0   # store has < 2 weeks of stock on hand for that SKU
MIN_WEEKLY_SALES_TO_FLAG  = 0.5   # must sell at least this many units/week to matter
REPLEN_TARGET_WEEKS       = 4.0   # top up to ~4 weeks of cover

# Transfer thresholds
EXCESS_WOH_THRESHOLD      = 12.0  # store holding > 12 weeks of cover = excess
EXCESS_MIN_UNITS          = 3     # don't bother suggesting transfer of 1-2 units
DEAD_STOCK_WEEKS          = 16.0  # beyond this, it's a clearance candidate, not a transfer candidate


def _safe(v, default=0.0):
    if v is None:
        return default
    try:
        f = float(v)
        if not np.isfinite(f):
            return default
        return f
    except (TypeError, ValueError):
        return default


def load_data(data_dir):
    pq_path = Path(data_dir) / "output_company.parquet"
    if not pq_path.exists():
        sys.exit(f"❌ output_company.parquet not found in {data_dir}")
    df = pd.read_parquet(pq_path)

    numeric_cols = (
        [s for s in STORES] +
        [f"{s}_L1W" for s in STORES] +
        ["AKLWH", "WH1", "STORE_TOT", "L1W", "L3W", "L6W", "avg_weekly", "WOH",
         "retail_price", "supply_price"]
    )
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    return df


def per_store_weekly_rate(row, store):
    """
    Estimate this SKU's weekly sell-through rate AT THIS STORE.
    We don't have a clean per-store L6W, so we approximate using the store's
    share of company-wide L1W sales applied to the company avg_weekly rate.
    Falls back to 0 if there's no signal.
    """
    store_l1w = _safe(row.get(f"{store}_L1W"))
    total_l1w = _safe(row.get("L1W"))
    company_avg = _safe(row.get("avg_weekly"))

    if total_l1w > 0 and company_avg > 0:
        store_share = store_l1w / total_l1w
        return company_avg * store_share
    # Fallback: if this store sold anything last week, use that as a rough rate
    return store_l1w if store_l1w > 0 else 0.0


def generate_replenishment(df):
    """
    WAREHOUSE -> STORE suggestions.

    IMPORTANT: negative SOH at a store means pre-orders/backorders — customers
    who already paid and are waiting. This is NOT data noise. A store sitting
    at SOH = -5 needs those 5 units PLUS its normal restock target; flooring
    it to 0 would silently under-supply customers who are already owed stock.
    """
    suggestions = []

    for _, row in df.iterrows():
        wh_avail = _safe(row.get("AKLWH")) + _safe(row.get("WH1"))
        if wh_avail <= 0:
            continue  # nothing in the warehouse to send

        sku  = str(row.get("supplier_code", "")).strip()
        prod = str(row.get("product", ""))
        if not sku:
            continue

        for store in STORES:
            soh = _safe(row.get(store))
            backorder_units = max(0, -soh)  # e.g. soh=-5 -> 5 units already owed to customers
            rate = per_store_weekly_rate(row, store)

            # A store with backorders needs stock even if its sales-velocity signal is weak/noisy
            has_backorder = backorder_units > 0
            if rate < MIN_WEEKLY_SALES_TO_FLAG and not has_backorder:
                continue  # not selling meaningfully and nothing owed — don't replenish

            effective_soh = max(soh, 0)  # for cover-calculation purposes only
            woh_at_store = (effective_soh / rate) if rate > 0 else (99 if not has_backorder else 0)

            if woh_at_store >= LOW_STOCK_WOH_THRESHOLD and not has_backorder:
                continue  # this store has enough cover already and nothing owed

            target_units = rate * REPLEN_TARGET_WEEKS
            # Need = (normal restock target - what's on the shelf) + whatever is already owed to customers
            need = max(0, round(target_units - effective_soh)) + round(backorder_units)
            if need <= 0:
                continue

            send_qty = int(min(need, wh_avail))
            if send_qty <= 0:
                continue

            reason = f"Selling ~{rate:.1f}/wk, only {woh_at_store:.1f} weeks cover at {STORE_LABELS.get(store, store)}"
            if has_backorder:
                reason = f"⚠️ {int(backorder_units)} units already on backorder/pre-order at {STORE_LABELS.get(store, store)} — " + reason

            suggestions.append({
                "type": "replenish",
                "sku": sku,
                "product": prod[:70],
                "to_store": STORE_LABELS.get(store, store),
                "to_store_code": store,
                "current_soh": int(soh),  # kept as-is (can be negative) so the UI can show the real backorder state
                "backorder_units": int(backorder_units),
                "weekly_rate": round(rate, 2),
                "weeks_cover": round(woh_at_store, 1),
                "suggested_qty": send_qty,
                "warehouse_available": int(wh_avail),
                "reason": reason,
            })
            wh_avail -= send_qty  # don't oversubscribe the same warehouse stock across stores

    # Backorders first (most urgent — customers already waiting), then by thinnest cover
    suggestions.sort(key=lambda x: (-x["backorder_units"], x["weeks_cover"], -x["weekly_rate"]))
    return suggestions


def generate_transfers(df):
    """
    STORE -> STORE suggestions for slow stock at one store, demand at another.

    Negative SOH = pre-orders/backorders (customers already waiting), not noise.
    A store with backorders is NEVER a donor (it has no real spare stock — the
    "negative" units are already owed to someone), but it IS treated as the
    most urgent possible receiver.
    """
    suggestions = []

    for _, row in df.iterrows():
        sku  = str(row.get("supplier_code", "")).strip()
        prod = str(row.get("product", ""))
        if not sku:
            continue

        # Build per-store snapshot for this SKU
        store_info = {}
        for store in STORES:
            soh  = _safe(row.get(store))            # keep real sign — negative means backorder
            backorder = max(0, -soh)
            rate = per_store_weekly_rate(row, store)
            effective_soh = max(soh, 0)
            woh  = (effective_soh / rate) if rate > 0 else (99 if (effective_soh > 0 and backorder == 0) else 0)
            store_info[store] = {
                "soh": soh, "effective_soh": effective_soh,
                "backorder": backorder, "rate": rate, "woh": woh,
            }

        # Donor stores: real positive excess stock only — backorder stores can NEVER donate
        donor_spare = {}
        for s, info in store_info.items():
            if info["backorder"] > 0:
                continue  # has pre-orders owed — not a donor under any circumstance
            if info["effective_soh"] >= EXCESS_MIN_UNITS and EXCESS_WOH_THRESHOLD <= info["woh"] < DEAD_STOCK_WEEKS:
                buffer = max(EXCESS_MIN_UNITS, round(info["rate"] * 2))
                spare  = max(0, info["effective_soh"] - buffer)
                if spare > 0:
                    donor_spare[s] = spare
        if not donor_spare:
            continue

        # Receiver stores: backorder stores are always eligible (most urgent),
        # plus normal low-cover-but-selling stores.
        receivers = []
        for s, info in store_info.items():
            if info["backorder"] > 0:
                receivers.append((s, info, True))   # urgent = True
            elif info["rate"] >= MIN_WEEKLY_SALES_TO_FLAG and info["woh"] < LOW_STOCK_WOH_THRESHOLD:
                receivers.append((s, info, False))
        if not receivers:
            continue

        # Most urgent (backorder) receivers first, then lowest cover first
        receivers.sort(key=lambda x: (not x[2], x[1]["woh"]))

        for donor_store, spare_remaining in list(donor_spare.items()):
            for recv_store, recv_info, is_urgent in receivers:
                if spare_remaining <= 0:
                    break
                if donor_store == recv_store:
                    continue

                if is_urgent:
                    # Cover the backorder first, then top up to the normal target
                    target_units = recv_info["backorder"] + recv_info["rate"] * REPLEN_TARGET_WEEKS
                    need = max(0, round(target_units - recv_info["effective_soh"]))
                else:
                    target_units = recv_info["rate"] * REPLEN_TARGET_WEEKS
                    need = max(0, round(target_units - recv_info["effective_soh"]))

                if need <= 0:
                    continue

                transfer_qty = int(min(need, spare_remaining))
                if transfer_qty <= 0:
                    continue

                donor_info = store_info[donor_store]
                reason = (
                    f"{STORE_LABELS.get(donor_store, donor_store)} has "
                    f"{donor_info['woh']:.1f} wks cover (slow); "
                    f"{STORE_LABELS.get(recv_store, recv_store)} "
                )
                reason += (
                    f"has {int(recv_info['backorder'])} units on backorder"
                    if is_urgent else
                    f"sells ~{recv_info['rate']:.1f}/wk but is low on stock"
                )

                suggestions.append({
                    "type": "transfer",
                    "sku": sku,
                    "product": prod[:70],
                    "from_store": STORE_LABELS.get(donor_store, donor_store),
                    "from_store_code": donor_store,
                    "to_store": STORE_LABELS.get(recv_store, recv_store),
                    "to_store_code": recv_store,
                    "from_soh": int(donor_info["soh"]),
                    "from_weeks_cover": round(donor_info["woh"], 1),
                    "to_soh": int(recv_info["soh"]),
                    "to_backorder_units": int(recv_info["backorder"]),
                    "to_weekly_rate": round(recv_info["rate"], 2),
                    "suggested_qty": transfer_qty,
                    "urgent_backorder": is_urgent,
                    "reason": reason,
                })

                spare_remaining -= transfer_qty
                recv_info["effective_soh"] += transfer_qty
                if is_urgent:
                    recv_info["backorder"] = max(0, recv_info["backorder"] - transfer_qty)

    # Backorder-driven transfers first, then largest quantity
    suggestions.sort(key=lambda x: (not x["urgent_backorder"], -x["suggested_qty"]))
    return suggestions


def generate_clearance_flags(df):
    """SKUs that are excess EVERYWHERE — not a transfer candidate, flag for markdown/clearance instead."""
    flags = []
    for _, row in df.iterrows():
        sku  = str(row.get("supplier_code", "")).strip()
        prod = str(row.get("product", ""))
        if not sku:
            continue

        store_tot   = _safe(row.get("STORE_TOT"))  # keep real sign — negative means net backorder, never a clearance candidate
        company_woh = _safe(row.get("WOH"), default=0)
        company_l6w = _safe(row.get("L6W"))

        if store_tot >= EXCESS_MIN_UNITS and company_woh >= DEAD_STOCK_WEEKS:
            flags.append({
                "type": "clearance_candidate",
                "sku": sku,
                "product": prod[:70],
                "total_units": int(store_tot),
                "weeks_cover": round(company_woh, 1),
                "l6w_units_sold": int(company_l6w),
                "inventory_cost": round(_safe(row.get("inventory_cost")), 2),
                "reason": f"{company_woh:.0f} weeks of cover company-wide with minimal recent sales — not a transfer fix, consider markdown",
            })

    flags.sort(key=lambda x: -x["inventory_cost"])
    return flags


def main():
    parser = argparse.ArgumentParser(description="Generate inventory rebalancing recommendations")
    parser.add_argument("--data-dir", default=r"D:\PopStop_App\popstop-dashboard\data")
    parser.add_argument("--out", default=None)
    args = parser.parse_args()

    out_dir = args.out or args.data_dir
    Path(out_dir).mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("  Inventory Rebalancing Recommendation Engine")
    print("=" * 60)
    print(f"📂 Data dir: {args.data_dir}\n")

    df = load_data(args.data_dir)
    print(f"📊 Loaded {len(df):,} SKUs\n")

    print("🔍 Generating warehouse → store replenishment suggestions...")
    replen = generate_replenishment(df)
    print(f"   Found {len(replen)} replenishment suggestions")

    print("🔍 Generating store → store transfer suggestions...")
    transfers = generate_transfers(df)
    print(f"   Found {len(transfers)} transfer suggestions")

    print("🔍 Flagging clearance candidates (excess everywhere)...")
    clearance = generate_clearance_flags(df)
    print(f"   Found {len(clearance)} clearance candidates")

    output = {
        "generated_at": datetime.now().isoformat(),
        "thresholds": {
            "low_stock_woh": LOW_STOCK_WOH_THRESHOLD,
            "excess_woh": EXCESS_WOH_THRESHOLD,
            "dead_stock_woh": DEAD_STOCK_WEEKS,
            "replen_target_weeks": REPLEN_TARGET_WEEKS,
        },
        "replenishment": replen,
        "transfers": transfers,
        "clearance_candidates": clearance,
    }

    out_path = Path(out_dir) / "inventory_recommendations.json"
    with open(out_path, "w") as f:
        json.dump(output, f, indent=2)

    print(f"\n✅ Saved → {out_path}")
    print(f"\nSummary: {len(replen)} replenish · {len(transfers)} transfer · {len(clearance)} clearance flags")


if __name__ == "__main__":
    main()
