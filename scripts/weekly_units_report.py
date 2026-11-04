"""
Weekly units file for the Stock Health Dashboard, built from the sales
ledger (no Zoho calls). Replaces the retired Item Runway weekly report, whose
reports/Sales_by_item(WKx-WKy).xlsx the weekly stock-health run reads every
Monday.

Same layout as before: SKU, Item Name, Brand, then one column per
Monday-Sunday week ("WK <ISO week>"), oldest first, 36 weeks, with the
number of units invoiced per SKU. One row per SKU that sold in the window,
except charge lines (shipping, delivery fees, adjustments). Item Name and
Brand come from the latest daily SOH report where the SKU is in it (combo
items such as AWB-TPTB are not), else from the ledger.

Basis: units invoiced, credit notes not netted off. This is what the old
report actually produced (its credit-note sync never worked) and the
"same basis as the archive" the stock-health method asks for.

The Saturday-night run includes the week that ends the next day, so a
Monday stock-health run finds the week just ended. That week's Sunday
invoices, and any late invoices, land when the file is rebuilt in full the
following Saturday.

    python scripts/weekly_units_report.py [YYYY-MM-DD]   # "today", default the real date (EAT)
"""
import glob
import os
import re
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

import sales_by_item_common as common
import sales_ledger as ledger

NUM_PERIODS = 36
NOT_STOCK = {"SHIPPING CHARGES", "ADJUSTMENT", "DELIVERY ORDER FEE"}


def latest_soh_catalog(reports_dir):
    """{SKU: (Item Name, Brand)} from the newest reports/SOH_DD_MM_YYYY.xlsx."""
    dated = []
    for path in glob.glob(os.path.join(reports_dir, "SOH_*.xlsx")):
        m = re.search(r"SOH_(\d{2})_(\d{2})_(\d{4})\.xlsx$", path)
        if m:
            dated.append((date(int(m[3]), int(m[2]), int(m[1])), path))
    if not dated:
        return {}
    sheet = load_workbook(max(dated)[1], read_only=True).active
    rows = sheet.iter_rows(values_only=True)
    header = [str(h or "").strip() for h in next(rows)]
    sku_i, name_i, brand_i = header.index("SKU"), header.index("Item Name"), header.index("Brand")
    return {str(r[sku_i]).strip(): (r[name_i] or "", r[brand_i] or "") for r in rows if r[sku_i]}


def build_rows(ledger_rows, buckets, catalog):
    start, end = buckets[0][0].isoformat(), buckets[-1][1].isoformat()
    index = {}
    for i, (s, e) in enumerate(buckets):
        for k in range((e - s).days + 1):
            index[(s + timedelta(days=k)).isoformat()] = i
    units = defaultdict(lambda: [0] * len(buckets))
    names = {}
    for r in sorted(ledger_rows, key=lambda r: r["date"]):
        if r["sku"].strip():
            names[r["sku"].strip()] = (r["item_name"], r["brand"])  # newest last
        d = r["date"]
        if not (start <= d <= end) or r["doc_type"] == "creditnote":
            continue
        qty = ledger.to_number(r["quantity"])
        if qty <= 0 or ledger.to_number(r["total"]) < 0:
            continue  # credit notes (old export: negative, no id) are not invoiced units
        units[r["sku"].strip()][index[d]] += qty
    rows = []
    for sku, q in units.items():
        if not any(q) or sku.upper() in NOT_STOCK:
            continue
        name, brand = catalog.get(sku) or names.get(sku, ("", ""))
        rows.append([sku, name, brand, *[int(v) if float(v).is_integer() else v for v in q]])
    rows.sort(key=lambda r: str(r[1]).lower())
    return rows


def main(today=None):
    today = today or datetime.now(ZoneInfo("Africa/Nairobi")).date()
    week_end = today + timedelta(days=6 - today.weekday())  # the Sunday ending this week
    buckets = common.compute_week_buckets_ending_by(week_end, NUM_PERIODS)
    oldest, newest = buckets[0][0].isocalendar()[1], buckets[-1][0].isocalendar()[1]
    out = os.path.join(ledger.REPO_ROOT, "reports", f"Sales_by_item(WK{oldest}-WK{newest}).xlsx")
    catalog = latest_soh_catalog(os.path.join(ledger.REPO_ROOT, "reports"))
    org = os.environ.get("ZOHO_ORGANIZATION_ID") or "780891784"
    rows = build_rows(ledger.load_ledger(ledger.ledger_path(org)), buckets, catalog)
    common.write_excel(rows, [common.week_label(s) for s, _ in buckets], out)
    print(f"Wrote {os.path.relpath(out, ledger.REPO_ROOT)}: {len(rows):,} items, "
          f"{buckets[0][0]} to {buckets[-1][1]} (latest week partial until it ends).")
    return out


if __name__ == "__main__":
    main(date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else None)
