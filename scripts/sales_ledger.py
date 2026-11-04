"""
Weekly Sales Ledger: keeps reports/Sales_13.xlsx (one row per invoice /
credit note line, every sale since Dec 2019) up to date from Zoho Inventory.

Columns (same as the file the team already uses): Invoice ID, Date, Sales
Person Name, Customer Name, SKU, Item Name, Brand, Category, Total, Quantity.
Credit notes are their own rows with negative Total and Quantity, dated on
the credit note's date (the convention the file already used for history).

Every run:
  1. Lists every invoice (all history) and every credit note dated on/after
     the ledger's `creditnotes_from` date. Listing is cheap: ~200 documents
     per call.
  2. Compares each listed document's `last_modified_time` with the one the
     ledger holds. A document that is new, or that was edited since (an
     invoice changed, a line removed, a price corrected), gets its detail
     fetched (one call each) and all of its rows are replaced. A document
     the ledger holds that Zoho no longer lists (voided or deleted) has its
     rows removed.
  3. Rewrites data/sales_ledger_<org>/<year>.csv (the master copy, with the extra
     tracking columns) and reports/Sales_13.xlsx (the file people open).

The ledger was seeded on 2026-11-02 from the team's Sales_13.xlsx export
(history to 20 Sep 2026) and Zoho's invoice-line CSV export (21 Sep - 1 Nov
2026) by scripts/seed_sales_ledger.py. Some history rows could not be tied
to a single Zoho invoice id (credit notes in the old export carry no id, and
Excel had rounded the rest); those rows are kept exactly as exported
(doc_type "history") and never re-checked, and the invoice ids they belong
to are listed in the state file as `frozen_invoice_ids` so the sync never
fetches them a second time.

Line Total = VAT-inclusive value after the line discount (see line_total). Checked against all 188 lines for 21 Sep 2026 in the old export
(Zoho's "Item Total" column is the ex-VAT amount and was NOT what the file
used). Each fetched invoice's line totals are also checked against the
invoice total Zoho lists; mismatches are reported in the log.

Environment variables:
    ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET, ZOHO_REFRESH_TOKEN  (required)
    ZOHO_ORGANIZATION_ID     - required only for multi-org accounts
    MAX_API_CALLS            - stop fetching details before this many calls
                               (blank/0 = no limit). Whatever is left is
                               picked up by the next run.
    RATE_LIMIT_PER_MINUTE    - default 60
"""
import csv
import json
import os
import sys
import time
from datetime import date, datetime

import requests
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font

import sales_by_item_common as common
from container_sellout_report import (ApiCallLimitReached, CallBudget, authorized_get,
                                      install_call_budget, parse_date, to_number)

REPO_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
OUTPUT_PATH = os.path.join(REPO_ROOT, "reports", "Sales_13.xlsx")
LEDGER_COLUMNS = ["doc_type", "doc_id", "last_modified_time", "invoice_id", "date", "sales_person",
                  "customer", "sku", "item_name", "brand", "category", "total", "quantity"]
XLSX_HEADER = ["Invoice ID", "Date", "Sales Person Name", "Customer Name", "SKU", "Item Name",
               "Brand", "Category", "Total", "Quantity"]
XLSX_WIDTHS = [20.29, 11.14, 21.71, 52.29, 15.86, 49.29, 11.29, 20.14, 11.57, 8.86]
ALL_HISTORY = date(2000, 1, 1)
# Refuse to apply a listing that suddenly holds far fewer invoices than the
# ledger: that is a broken listing, not thousands of voided invoices.
MIN_LISTED_SHARE = 0.9
# (doc_type, list path, list key, id field, detail key, sign)
DOC_TYPES = [
    ("invoice", "invoices", "invoices", "invoice_id", "invoice", 1),
    ("creditnote", "creditnotes", "creditnotes", "creditnote_id", "creditnote", -1),
]


def ledger_path(org):
    """A folder with one CSV per year (data/sales_ledger_<org>/2026.csv), so
    no single file outgrows GitHub's limits and a week's changes only touch
    the recent years' files."""
    return os.path.join(REPO_ROOT, "data", f"sales_ledger_{org}")


def state_path(org):
    return os.path.join(REPO_ROOT, "data", f"sales_ledger_state_{org}.json")


def overrides_path():
    return os.path.join(REPO_ROOT, "data", "sales_category_overrides.csv")


# ------------------------------------------------------------------ storage

def load_ledger(path):
    rows = []
    for name in sorted(os.listdir(path), reverse=True):  # newest year first
        if name.endswith(".csv"):
            with open(os.path.join(path, name), newline="", encoding="utf-8") as handle:
                rows.extend(csv.DictReader(handle))
    for r in rows:
        r["invoice_id"] = r["invoice_id"] or r["doc_id"]
    return sort_ledger(rows)


def sort_ledger(rows):
    """Newest date first, keeping the existing order within a day (stable)."""
    rows.sort(key=lambda r: r["date"], reverse=True)
    return rows


def save_ledger(path, rows):
    """Rows must already be sorted. invoice_id is only stored when it differs
    from doc_id (history rows), to keep the files small."""
    os.makedirs(path, exist_ok=True)
    by_year = {}
    for r in rows:
        out = dict(r)
        if out["invoice_id"] == out["doc_id"]:
            out["invoice_id"] = ""
        by_year.setdefault(r["date"][:4], []).append(out)
    for name in os.listdir(path):
        if name.endswith(".csv") and name[:-4] not in by_year:
            os.remove(os.path.join(path, name))
    for year, year_rows in by_year.items():
        target = os.path.join(path, f"{year}.csv")
        tmp = target + ".tmp"
        with open(tmp, "w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=LEDGER_COLUMNS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(year_rows)
        os.replace(tmp, target)


def load_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


def save_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump(data, handle, indent=2, sort_keys=True)
        handle.write("\n")
    os.replace(tmp, path)


def load_overrides(path):
    if not os.path.exists(path):
        return {}
    with open(path, newline="", encoding="utf-8") as handle:
        return {r["SKU"].strip(): r["Category"].strip() for r in csv.DictReader(handle) if r.get("SKU")}


# ------------------------------------------------------------------ xlsx

def xlsx_invoice_id(row):
    """History rows keep the value exactly as the old export had it (a
    rounded number, or blank for credit notes); synced invoices carry
    Zoho's full id as text so no digits are lost. Credit notes are always
    blank, as in the old export: the team's reports treat a blank Invoice ID
    with a negative Total as a credit note."""
    if row["doc_type"] == "creditnote":
        return None
    value = row["invoice_id"]
    if row["doc_type"] == "history" and value:
        try:
            return float(value)
        except ValueError:
            return value
    return value or None


def number(value):
    v = to_number(value)
    return int(v) if v.is_integer() else v


def write_xlsx(rows, output_path):
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("Sheet1")
    for letter, width in zip("ABCDEFGHIJ", XLSX_WIDTHS):
        sheet.column_dimensions[letter].width = width
    bold = Font(bold=True)

    def cell(value, fmt=None, font=None):
        c = WriteOnlyCell(sheet, value)
        if fmt:
            c.number_format = fmt
        if font:
            c.font = font
        return c

    sheet.append([cell(h, font=bold) for h in XLSX_HEADER])
    written = 0
    for r in rows:
        if not r["sku"] and not r["item_name"] and not r["quantity"]:
            continue  # placeholder for a document with no item lines
        invoice_id = xlsx_invoice_id(r)
        sheet.append([
            cell(invoice_id, "0" if isinstance(invoice_id, float) else None),
            cell(datetime.strptime(r["date"], "%Y-%m-%d"), "dd MMM yyyy"),
            cell(r["sales_person"] or None),
            cell(r["customer"] or None),
            cell(r["sku"] or None),
            cell(r["item_name"] or None),
            cell(r["brand"] or None),
            cell(r["category"]),
            cell(number(r["total"]), "#,##0.00"),
            cell(number(r["quantity"]), "0.00"),
        ])
        written += 1
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    tmp = output_path + ".tmp.xlsx"
    workbook.save(tmp)
    os.replace(tmp, output_path)
    return written


# ------------------------------------------------------------------ line building

def discount_percent(line, gross):
    """The line discount as a percentage, or None if it was entered as an
    amount. Zoho returns a percentage either as "10%" or as a bare 10 next to
    a `discount_amount`; on a VAT-inclusive invoice that amount is sometimes
    the share of the ex-VAT value (e.g. INV-017807: 10% of 5,520,000 incl.
    VAT came back as 467,796.61) and sometimes of the incl.-VAT value
    (INV-017002), so a bare number counts as a percentage when the amount
    matches it on either basis."""
    discount = line.get("discount")
    if isinstance(discount, str) and discount.strip().endswith("%"):
        return to_number(discount.strip()[:-1])
    pct = to_number(discount)
    amount = line.get("discount_amount")
    if not (0 < pct <= 100) or amount in (None, ""):
        return None
    amount = to_number(amount)
    tax = to_number(line.get("tax_percentage")) / 100
    for base in (gross, gross / (1 + tax)):
        if abs(base * pct / 100 - amount) <= max(1.0, 0.005 * amount):
            return pct
    return None


def line_total(line, inclusive=True):
    """VAT-inclusive line value after the line discount.

    The rate is VAT-inclusive on an inclusive invoice (all KTL invoices so
    far), so a percentage discount is simply gross x (1 - pct). Checked
    against all 9,906 lines of Zoho's 21 Sep - 1 Nov export, where it equals
    Zoho's Item Total + Item Tax Amount on every line. Subtracting
    `discount_amount` instead overstated 40 invoices in the first live run
    (2026-11-04), because that amount is sometimes ex-VAT."""
    gross = to_number(line.get("rate")) * to_number(line.get("quantity"))
    pct = discount_percent(line, gross)
    if pct is not None:
        net = gross * (1 - pct / 100)
    elif line.get("discount_amount") not in (None, ""):
        net = gross - to_number(line.get("discount_amount"))
    else:
        net = gross - to_number(line.get("discount"))
    if not inclusive:
        net *= 1 + to_number(line.get("tax_percentage")) / 100
    return net


def entity_discount(detail):
    """Invoice-level (not line-level) discount amount, spread over the lines."""
    if (detail.get("discount_type") or "").lower() != "entity_level":
        return 0.0
    return to_number(detail.get("discount_amount") or detail.get("discount_total"))


class Lookups:
    """Brand/Category for new rows. Category: manual override, else the SKU's
    most recent category in the ledger (keeps the file consistent with its
    history), else Zoho's item category. Brand: the line, else Zoho's item,
    else the SKU's most recent brand in the ledger."""

    def __init__(self, rows, overrides, catalog):
        self.overrides = overrides
        self.catalog = catalog  # item_id -> {"brand", "category"}
        self.category_by_sku, self.brand_by_sku = {}, {}
        for r in rows:  # newest first, so the first hit is the latest
            if r["sku"] and r["category"]:
                self.category_by_sku.setdefault(r["sku"], r["category"])
            if r["sku"] and r["brand"]:
                self.brand_by_sku.setdefault(r["sku"], r["brand"])

    def category(self, sku, item_id):
        if sku in self.overrides:
            return self.overrides[sku]
        if sku in self.category_by_sku:
            return self.category_by_sku[sku]
        return (self.catalog.get(item_id, {}).get("category") or "").upper()

    def brand(self, line, sku, item_id):
        return (line.get("brand") or self.catalog.get(item_id, {}).get("brand")
                or self.brand_by_sku.get(sku, ""))


def document_rows(doc_type, doc_id, modified, detail, sign, lookups):
    lines = [l for l in detail.get("line_items") or [] if to_number(l.get("quantity"))]
    inclusive = str(detail.get("is_inclusive_tax", True)).lower() != "false"
    totals = [line_total(l, inclusive) for l in lines]
    discount = entity_discount(detail)
    gross = sum(totals)
    if discount and gross:
        totals = [t * (gross - discount) / gross for t in totals]
    base = {
        "doc_type": doc_type, "doc_id": doc_id, "last_modified_time": modified,
        "invoice_id": doc_id, "date": str(detail.get("date") or "")[:10],
        "sales_person": detail.get("salesperson_name") or "",
        "customer": detail.get("customer_name") or "",
    }
    rows = []
    for line, total in zip(lines, totals):
        sku = str(line.get("sku") or "").strip()
        item_id = str(line.get("item_id") or "")
        rows.append({**base, "sku": sku, "item_name": line.get("name") or "",
                     "brand": lookups.brand(line, sku, item_id),
                     "category": lookups.category(sku, item_id),
                     "total": f"{sign * total:.2f}",
                     "quantity": f"{sign * to_number(line.get('quantity')):g}"})
    if not rows:  # keep a placeholder so the document counts as synced
        rows.append({**base, "sku": "", "item_name": "", "brand": "", "category": "",
                     "total": "", "quantity": ""})
    return rows


def check_document_total(detail, rows):
    """Difference between Zoho's document total and what the rows add up to
    (plus shipping/adjustment/round-off, which aren't item lines)."""
    expected = to_number(detail.get("total"))
    extras = sum(to_number(detail.get(k)) for k in ("shipping_charge", "adjustment", "roundoff_value"))
    got = sum(abs(to_number(r["total"])) for r in rows) + extras
    return expected - got


# ------------------------------------------------------------------ Zoho

def list_docs(session, api_domain, org, token_store, path, list_key, id_field, since):
    """{doc_id: (date, last_modified_time)}, non-draft/non-void, dated on/after `since`."""
    docs, page = {}, 1
    while True:
        response = authorized_get(session, token_store, f"{api_domain}/inventory/v1/{path}",
                                  params={"organization_id": org, "sort_column": "date", "sort_order": "D",
                                          "page": page, "per_page": common.INVOICE_LIST_PAGE_SIZE},
                                  timeout=60)
        if response.status_code == 429:
            time.sleep(common.parse_retry_after(response, 3))
            continue
        response.raise_for_status()
        payload = response.json()
        reached_start = False
        for doc in payload.get(list_key) or []:
            doc_date = parse_date(doc.get("date"))
            if doc_date is None:
                continue
            if doc_date < since:
                reached_start = True
                break
            if (doc.get("status") or "").strip().lower() in common.EXCLUDED_INVOICE_STATUSES:
                continue
            if doc.get(id_field):
                docs[str(doc[id_field])] = (doc_date.isoformat(), str(doc.get("last_modified_time") or ""))
        if reached_start or not (payload.get("page_context") or {}).get("has_more_page"):
            return docs
        page += 1


def fetch_item_catalog(session, api_domain, org, token_store):
    """item_id -> {"brand", "category"} for every item (any status)."""
    catalog, page = {}, 1
    while True:
        response = authorized_get(session, token_store, f"{api_domain}/inventory/v1/items",
                                  params={"organization_id": org, "filter_by": "Status.All",
                                          "page": page, "per_page": 200}, timeout=60)
        if response.status_code == 429:
            time.sleep(common.parse_retry_after(response, 3))
            continue
        response.raise_for_status()
        payload = response.json()
        for item in payload.get("items") or []:
            catalog[str(item.get("item_id"))] = {"brand": common.extract_brand(item),
                                                 "category": item.get("category_name") or ""}
        if not (payload.get("page_context") or {}).get("has_more_page"):
            return catalog
        page += 1


# ------------------------------------------------------------------ sync

def plan_sync(rows, listed_by_type, frozen_ids, creditnotes_from):
    """What to fetch and what to drop.

    `listed_by_type`: {doc_type: {doc_id: (date, modified)}} -- a type is
    missing if it could not be listed this run, and is then left alone.
    Returns (todo, removed): todo is [(doc_type, doc_id)] new or changed
    since the ledger; removed is the set of (doc_type, doc_id) the ledger
    holds that Zoho no longer lists."""
    held = {}
    for r in rows:
        if r["doc_type"] in ("invoice", "creditnote"):
            held[(r["doc_type"], r["doc_id"])] = (r["date"], r["last_modified_time"])

    removed = set()
    for key, (doc_date, _modified) in held.items():
        listed = listed_by_type.get(key[0])
        if listed is None:
            continue
        if key[0] == "creditnote" and doc_date < creditnotes_from:
            continue
        if key[1] not in listed:
            removed.add(key)

    todo = []
    for doc_type, listed in listed_by_type.items():
        for doc_id, (doc_date, modified) in listed.items():
            if doc_type == "invoice" and doc_id in frozen_ids:
                continue
            current = held.get((doc_type, doc_id))
            if current is None or current[1] != modified or not modified:
                todo.append((doc_type, doc_id))
    todo.sort(key=lambda k: (listed_by_type[k[0]][k[1]][0], k))  # oldest first
    return todo, removed


def apply_changes(rows, removed, fetched):
    """Drops removed documents and replaces/adds fetched ones. `fetched`:
    {(doc_type, doc_id): [rows]}. Returns the new, sorted ledger."""
    replace = removed | set(fetched)
    kept = [r for r in rows if (r["doc_type"], r["doc_id"]) not in replace]
    for key in sorted(fetched, key=lambda k: k[1]):
        kept.extend(fetched[key])
    return sort_ledger(kept)


def main():
    budget = CallBudget(int(os.environ.get("MAX_API_CALLS") or 0))
    try:
        run(budget)
    except ApiCallLimitReached:
        print(f"STOPPED: reached the limit of {budget.limit} Zoho API calls before the ledger was saved.",
              file=sys.stderr, flush=True)
        sys.exit(2)
    finally:
        print(f"Zoho API calls this run: {budget.used}" + (f" (limit {budget.limit})" if budget.limit else ""),
              flush=True)


def run(budget):
    access_token, api_domain = common.get_access_token()
    token_store = common.TokenStore(access_token)
    session = requests.Session()
    install_call_budget(session, budget)
    session.headers["Authorization"] = f"Zoho-oauthtoken {access_token}"
    rate_limiter = common.RateLimiter(int(os.environ.get("RATE_LIMIT_PER_MINUTE") or common.DEFAULT_RATE_LIMIT_PER_MINUTE))
    org = common.resolve_organization_id(session, api_domain)

    path = ledger_path(org)
    rows = load_ledger(path)
    state = load_json(state_path(org), {})
    frozen_ids = set(state.get("frozen_invoice_ids") or [])
    creditnotes_from = state.get("creditnotes_from") or "2026-09-21"
    print(f"Ledger: {len(rows):,} rows; credit notes synced from {creditnotes_from}; "
          f"{len(frozen_ids):,} frozen history invoice ids", flush=True)

    listed_by_type = {}
    for doc_type, list_path, list_key, id_field, _detail_key, _sign in DOC_TYPES:
        since = ALL_HISTORY if doc_type == "invoice" else date.fromisoformat(creditnotes_from)
        try:
            listed_by_type[doc_type] = list_docs(session, api_domain, org, token_store,
                                                 list_path, list_key, id_field, since)
            print(f"Listed {len(listed_by_type[doc_type]):,} {list_path} since {since}", flush=True)
        except requests.exceptions.HTTPError as exc:
            if doc_type == "invoice":
                raise
            print(f"Warning: could not list {list_path} ({exc}); credit notes left as they were this run.",
                  file=sys.stderr, flush=True)

    held_invoices = {r["doc_id"] for r in rows if r["doc_type"] == "invoice"}
    if len(listed_by_type["invoice"]) < MIN_LISTED_SHARE * len(held_invoices):
        raise RuntimeError(f"Zoho listed only {len(listed_by_type['invoice']):,} invoices but the ledger holds "
                           f"{len(held_invoices):,}; refusing to delete that many. Nothing was changed.")

    todo, removed = plan_sync(rows, listed_by_type, frozen_ids, creditnotes_from)
    print(f"To fetch: {len(todo):,} new/edited documents; to remove: {len(removed):,} voided/deleted", flush=True)

    remaining = len(todo)
    if budget.limit:
        allowed = max(0, budget.limit - budget.used - 25)  # leave room for the item catalog
        if len(todo) > allowed:
            print(f"Call limit: fetching {allowed:,} of {len(todo):,} now; the rest next run.", flush=True)
            todo = todo[:allowed]

    fetched, mismatches = {}, []
    if todo:
        lookups = Lookups(rows, load_overrides(overrides_path()),
                          fetch_item_catalog(session, api_domain, org, token_store))
        types = {t[0]: t for t in DOC_TYPES}

        def fetch_one(key):
            doc_type, doc_id = key
            _t, list_path, _lk, _id, detail_key, _s = types[doc_type]
            detail = common.fetch_detail(session, api_domain, org, token_store, rate_limiter,
                                         url=f"{api_domain}/inventory/v1/{list_path}/{doc_id}",
                                         result_key=detail_key, label=f"{detail_key} {doc_id}")
            return None if detail is None else (key, detail)

        for key, detail in common.fetch_concurrently(todo, fetch_one, "document detail"):
            doc_type, doc_id = key
            modified = listed_by_type[doc_type][doc_id][1]
            doc_rows = document_rows(doc_type, doc_id, modified, detail, types[doc_type][5], lookups)
            gap = check_document_total(detail, doc_rows)
            if abs(gap) > 1:
                mismatches.append((doc_type, detail.get("invoice_number") or detail.get("creditnote_number") or doc_id, gap))
            fetched[key] = doc_rows

    rows = apply_changes(rows, removed, fetched)
    remaining -= len(fetched)
    state["last_run"] = date.today().isoformat()
    state["pending_documents"] = remaining
    save_ledger(path, rows)
    save_json(state_path(org), state)
    written = write_xlsx(rows, OUTPUT_PATH)

    added = sum(1 for k in fetched if k[0] == "invoice")
    print(f"Updated {added:,} invoices and {len(fetched) - added:,} credit notes; removed {len(removed):,}.")
    print(f"Wrote {written:,} lines to {os.path.relpath(OUTPUT_PATH, REPO_ROOT)}.")
    if remaining:
        print(f"PARTIAL: {remaining:,} documents still to fetch next run.")
    if mismatches:
        print(f"Check: {len(mismatches)} document(s) whose lines don't add up to Zoho's total "
              f"(first few: {mismatches[:5]}). Line totals may need a different formula.", file=sys.stderr)


if __name__ == "__main__":
    main()
