#!/usr/bin/env python3
"""Builds dashboard/jmn_dashboard.html from db/export.json + clients/jmn/config.json.

Run `python3 db/build_db.py` first (to refresh db/export.json with the latest batch),
then `python3 dashboard/build_dashboard.py`. The output is a single self-contained HTML
file — publish it as a Claude Artifact (or open it locally) to view it. It's a static
snapshot: re-run this after every new batch to update it, it does not live-refresh.
"""
import json
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXPORT_PATH = os.path.join(BASE, "db", "export.json")
CONFIG_PATH = os.path.join(BASE, "clients", "jmn", "config.json")
TEMPLATE_PATH = os.path.join(BASE, "dashboard", "template.html")
OUTPUT_PATH = os.path.join(BASE, "dashboard", "jmn_dashboard.html")

# Not yet persisted anywhere per batch (optimize.py computes these at run time but
# doesn't write them to logs/changes_*.csv — see the history caveat in CLAUDE.md).
# Update these two numbers by hand from the terminal output of `python3 optimize.py jmn`
# next time it's run, or wire real persistence into optimize.py.
ACCOUNT_AOV = 118
BASELINE_CVR_PCT = 3.95


def build_payload():
    with open(EXPORT_PATH) as f:
        export = json.load(f)
    with open(CONFIG_PATH) as f:
        cfg = json.load(f)
    jmn = export["clients"]["jmn"]

    series = {}
    for row in jmn["product_sales"]:
        series.setdefault(row["sku"], {}).setdefault(row["market"], {})[
            f"{row['year']:04d}-{row['month']:02d}"
        ] = round(row["sales"], 2)

    roi = {}
    for r in jmn["roi_targets"]:
        roi.setdefault(r["sku"], {})[r["bucket"]] = r["roi_target"]

    all_months = sorted({m for sku in series.values() for mkt in sku.values() for m in mkt})
    if all_months and all_months[0] == "2021-12":
        all_months = all_months[1:]  # stray single-month boundary artifact, not a real trend point

    products = {}
    for sku, mkts in series.items():
        entry = {mkt: [vals.get(m) for m in all_months] for mkt, vals in mkts.items()}
        if sku in roi:
            entry["roi"] = roi[sku]
        products[sku] = entry

    batches = jmn["batches"]
    latest_batch = batches[-1] if batches else None
    decisions = []
    if latest_batch:
        for dec in latest_batch["decisions"]:
            d = {k: v for k, v in dec.items() if k not in ("decision_id", "batch_id")}
            for k in ("old_bid", "new_bid", "clicks", "spend", "sales", "orders"):
                if d.get(k) is not None:
                    d[k] = round(d[k], 2)
            decisions.append(d)

    ha, la, ws, hr = (
        cfg["bid_rules"]["high_acos"], cfg["bid_rules"]["low_acos"],
        cfg["bid_rules"]["wasted_spend"], cfg["history_rules"],
    )
    return {
        "client": "jmn",
        "targetAcos": cfg["target_acos"],
        "accountAov": ACCOUNT_AOV,
        "baselineCvr": BASELINE_CVR_PCT,
        "config": {
            "minClicks": cfg["data_gate"]["min_clicks"],
            "highAcos": {
                "extremeMultiple": ha["extreme_multiple_of_target"],
                "normalMaxDecrease": ha["normal_max_decrease_pct"],
                "extremeMaxDecrease": ha["extreme_max_decrease_pct"],
            },
            "lowAcos": {
                "thresholdPct": la["threshold_pct_of_target"],
                "extremePct": la["extreme_pct_of_target"],
                "normalMinIncrease": la["normal_min_increase_pct"],
                "normalMaxIncrease": la["normal_max_increase_pct"],
                "extremeMaxIncrease": la["extreme_max_increase_pct"],
            },
            "wastedSpend": {
                "maxZeroProb": round(ws["max_zero_order_probability_to_act"] * 100, 2),
                "extremeZeroProb": round(ws["extreme_zero_order_probability"] * 100, 2),
                "extremeMinSpendMultiple": ws["extreme_min_spend_multiple_of_target_cpa"],
                "normalMaxDecrease": ws["normal_max_decrease_pct"],
                "extremeMaxDecrease": ws["extreme_max_decrease_pct"],
                "bidFloor": ws["bid_floor"],
            },
            "cooldownDays": hr["cooldown_days"],
            "reversalDays": hr["reversal_protection_days"],
        },
        "batch": {"date": latest_batch["batch_date"] if latest_batch else None, "decisions": decisions},
        "months": all_months,
        "products": products,
    }


def main():
    payload = build_payload()
    with open(TEMPLATE_PATH) as f:
        template = f.read()
    html = template.replace("__DATA_JSON__", json.dumps(payload, separators=(",", ":")))
    with open(OUTPUT_PATH, "w") as f:
        f.write(html)
    print(f"Built {OUTPUT_PATH} ({len(html)} bytes) from batch {payload['batch']['date']}")


if __name__ == "__main__":
    main()
