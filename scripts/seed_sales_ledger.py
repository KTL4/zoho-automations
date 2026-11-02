"""
One-off: builds data/sales_ledger_<org>/<year>.csv and its state file from
  1. the team's Sales_13.xlsx export (all history; rows dated before
     SPLIT_DATE are kept),
  2. Zoho's invoice-line CSV export (Invoices > Export, SPLIT_DATE onward),
  3. data/sales_lines_<org>.csv (the Container Sellout invoice cache: full
     invoice ids, dates, last_modified_time and SKU/qty per line), used to
     tie history rows to Zoho invoice ids so later edits can be detected.

Usage:
    python scripts/seed_sales_ledger.py Sales_13.xlsx Invoice1.csv 780891784

Excel stored the old export's invoice ids as rounded numbers, so a history
row is tied to an invoice by (rounded id, date) and then by SKU+quantity
when several invoices share the same rounded id that day. Rows that still
can't be tied, and every old credit note (they have no id), are kept as-is
as doc_type "history"; the invoice ids they might belong to are frozen so the
weekly sync never fetches them again (that would double-count them).

A tied invoice whose lines differ from what Zoho holds now gets a blank
last_modified_time, so the first weekly run re-fetches it.
"""
import csv
import os
import sys
from collections import Counter, defaultdict
from datetime import datetime

import openpyxl

import sales_ledger as ledger

SPLIT_DATE = "2026-09-21"


def qty_key(sku, qty):
    return (str(sku or "").strip(), round(float(qty or 0), 4))


def load_zoho_cache(path):
    cache = defaultdict(lambda: {"date": "", "modified": "", "lines": Counter()})
    with open(path, newline="", encoding="utf-8") as handle:
        for r in csv.DictReader(handle):
            if r["doc_type"] != "invoice":
                continue
            e = cache[r["doc_id"]]
            e["date"], e["modified"] = r["doc_date"], r["last_modified_time"]
            if r["sku"] or r["item_id"]:
                e["lines"][qty_key(r["sku"], r["quantity"])] += 1
    return cache


def history_row(r, doc_type="history", doc_id="", modified=""):
    raw_id = r[0]
    shown = doc_id or ("" if raw_id in (None, "") else repr(raw_id) if isinstance(raw_id, float) else str(raw_id))
    total, qty = r[8] or 0, r[9] or 0
    return {"doc_type": doc_type, "doc_id": doc_id, "last_modified_time": modified, "invoice_id": shown,
            "date": r[1].strftime("%Y-%m-%d"), "sales_person": r[2] or "", "customer": r[3] or "",
            "sku": r[4] or "", "item_name": r[5] or "", "brand": r[6] or "", "category": r[7] or "",
            "total": f"{float(total):.2f}", "quantity": f"{float(qty):g}"}


def assign_group(rows, candidates, cache):
    """Splits rows sharing one (rounded id, date) across the candidate
    invoices by SKU+quantity. Returns {invoice_id: [rows]} or None."""
    remaining = {c: Counter(cache[c]["lines"]) for c in candidates}
    out = defaultdict(list)
    for r in rows:
        k = qty_key(r[4], r[9])
        hit = next((c for c in sorted(candidates) if remaining[c][k] > 0), None)
        if hit is None:
            return None
        remaining[hit][k] -= 1
        out[hit].append(r)
    return out


def main(xlsx_path, export_csv, org):
    cache = load_zoho_cache(os.path.join(ledger.REPO_ROOT, "data", f"sales_lines_{org}.csv"))
    by_key = defaultdict(set)
    for doc_id, e in cache.items():
        by_key[(float(doc_id), e["date"])].add(doc_id)

    sheet = openpyxl.load_workbook(xlsx_path, read_only=True).active
    groups, loose = defaultdict(list), []
    for r in sheet.iter_rows(min_row=2, values_only=True):
        if r[1] is None or r[1].strftime("%Y-%m-%d") >= SPLIT_DATE:
            continue
        if isinstance(r[0], (int, float)):
            groups[(float(r[0]), r[1].strftime("%Y-%m-%d"))].append(r)
        else:
            loose.append(r)

    rows, frozen = [], set()
    stats = Counter()
    for key, g_rows in groups.items():
        candidates = by_key.get(key, set())
        split = assign_group(g_rows, candidates, cache) if candidates else None
        if split is None:
            frozen |= candidates
            rows.extend(history_row(r) for r in g_rows)
            stats["history rows (not tied to an id)"] += len(g_rows)
            continue
        for doc_id, inv_rows in split.items():
            same = Counter(qty_key(r[4], r[9]) for r in inv_rows) == cache[doc_id]["lines"]
            modified = cache[doc_id]["modified"] if same else ""
            stats["tied invoices" if same else "tied invoices to re-fetch"] += 1
            rows.extend(history_row(r, "invoice", doc_id, modified) for r in inv_rows)
    rows.extend(history_row(r) for r in loose)
    stats["history credit-note/blank-id rows"] += len(loose)

    export_lines = defaultdict(list)
    with open(export_csv, newline="", encoding="utf-8-sig") as handle:
        for x in csv.DictReader(handle):
            export_lines[x["Invoice ID"]].append(x)
    for doc_id, lines in export_lines.items():
        same = Counter(qty_key(x["SKU"], x["Quantity"]) for x in lines) == cache.get(doc_id, {"lines": None})["lines"]
        modified = cache[doc_id]["modified"] if same else ""
        stats["export invoices" if same else "export invoices to re-fetch"] += 1
        for x in lines:
            total = float(x["Item Total"]) + float(x["Item Tax Amount"] or 0)
            rows.append({"doc_type": "invoice", "doc_id": doc_id, "last_modified_time": modified,
                         "invoice_id": doc_id, "date": x["Invoice Date"], "sales_person": x["Sales Person"],
                         "customer": x["Customer Name"], "sku": x["SKU"], "item_name": x["Item Name"],
                         "brand": x["Brand"], "category": "", "total": f"{total:.2f}",
                         "quantity": f"{float(x['Quantity']):g}"})

    ledger.sort_ledger(rows)
    lookups = ledger.Lookups(rows, ledger.load_overrides(ledger.overrides_path()), {})
    for r in rows:
        if r["doc_type"] == "invoice" and r["date"] >= SPLIT_DATE:
            r["category"] = lookups.category(r["sku"], "")

    held = {r["doc_id"] for r in rows if r["doc_type"] == "invoice"}
    missing = [d for d, e in cache.items() if d not in held and d not in frozen]
    stats["Zoho invoices in neither file (first run fetches them)"] = len(missing)

    ledger.save_ledger(ledger.ledger_path(org), rows)
    ledger.save_json(ledger.state_path(org), {
        "creditnotes_from": SPLIT_DATE,
        "frozen_invoice_ids": sorted(frozen),
        "seeded": datetime.now().strftime("%Y-%m-%d"),
    })
    print(f"{len(rows):,} ledger rows")
    for k, v in stats.items():
        print(f"  {k}: {v:,}")


if __name__ == "__main__":
    main(*sys.argv[1:4])
