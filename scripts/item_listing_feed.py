"""
ITEM LISTING feed: a compact workbook the ITEM LISTING Claude project links
from GitHub instead of the full reports/Sales_13.xlsx (which is too big for a
project). Written by the weekly Sales Ledger run, after the ledger is updated.

Output: reports/ITEM_LISTING_FEED.xlsx, sheets (all formatted as Excel tables):
  Info       - when it was built, data through, detail window, row counts
  Sales      - daily customer x SKU detail from 1 Jul of the previous year: Date, Cust, SKU, Qty, Total
  Orders     - one row per order (and per customer-day of credit notes) from 1 Jan of the
               previous year, with the value split by brand
  Customers  - one row per customer, full history: assigned BDM, first/last dates
  Items      - SKU, Item Name, Brand, Category
  Listings   - one row per customer-SKU pair, full history back to Dec 2019,
               with the listing date and whether it counts as a new listing

Rules (from the project's instructions):
  - Brands AIWIBI, AIWINA and UBL only; lines with Total = 0 and Quantity = 0 dropped.
  - Credit notes: negative lines; no order number.
  - Order = Invoice ID + Customer + Date (old invoice ids lost precision).
  - BDM = the active BDM on the customer's most recent invoice sold by an
    active BDM; "Unassigned" if none of them ever sold to the customer.
  - Paid purchase = line Total > 0. Listing date = first paid purchase of the
    customer-SKU pair. It counts as a new listing only if it falls on a day
    after the customer's first paid purchase day, and it isn't a Carrefour
    ("MAJID AL FUTTAIM...") X-SKU first bought in January (Christmas items
    listed the year before).

Sales holds detail from 1 July of the previous year (a 1 January start date
keeps its 182-day look-back) and Orders from 1 January (the 365-day "gone
quiet" look-back), which keeps the file small enough for a project; everything
that needs the full history (first purchases, listing dates, BDM) is worked
out here from the whole ledger and stored on Customers / Listings.
"""
import os
import sys
from collections import defaultdict
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo

import sales_ledger as ledger

BRANDS = ("AIWIBI", "AIWINA", "UBL")
ACTIVE_BDMS = ("SAM AGABA", "PROMISE PRISCILLAR", "MOREEN NAHWERA", "RAYMOND MPAMIZO", "GEOFREY OGEN",
               "MARGARET KIRABO NAMALE", "DAVID ARUO OCAN", "EDDY WAMBUZI", "PASORINE AGABA", "JOAN MBABAZI")
CARREFOUR_PREFIX = "MAJID AL FUTTAIM"
OUTPUT_PATH = os.path.join(ledger.REPO_ROOT, "reports", "ITEM_LISTING_FEED.xlsx")


def num(value):
    return ledger.to_number(value)


def prepare(rows):
    """Ledger rows -> cleaned lines (dicts), oldest first."""
    out = []
    for r in rows:
        brand = (r["brand"] or "").strip().upper()
        if brand not in BRANDS:
            continue
        total, qty = num(r["total"]), num(r["quantity"])
        if total == 0 and qty == 0:
            continue
        is_credit = r["doc_type"] == "creditnote" or (total < 0 and not r["invoice_id"])
        out.append({"date": r["date"], "customer": r["customer"].strip(), "sku": r["sku"].strip(),
                    "item": r["item_name"], "brand": brand, "category": r["category"],
                    "salesperson": (r["sales_person"] or "").strip().upper(),
                    "invoice": "" if is_credit else r["invoice_id"], "credit": is_credit,
                    "total": total, "qty": qty})
    out.sort(key=lambda x: x["date"])
    return out


def build(lines, sales_from, orders_from):
    customers = sorted({l["customer"] for l in lines})
    cust_no = {c: i for i, c in enumerate(customers, 1)}

    first_any, first_paid, last_paid, bdm = {}, {}, {}, {}
    first_paid_brand = defaultdict(dict)
    items, pair_first, pair_last = {}, {}, {}
    for l in lines:  # oldest first
        c, d = l["customer"], l["date"]
        first_any.setdefault(c, d)
        items[l["sku"]] = (l["item"], l["brand"], l["category"])  # latest wins
        if not l["credit"] and l["salesperson"] in ACTIVE_BDMS:
            bdm[c] = l["salesperson"]  # latest wins
        if l["total"] > 0:
            first_paid.setdefault(c, d)
            first_paid_brand[c].setdefault(l["brand"], d)
            last_paid[c] = d
            pair_first.setdefault((c, l["sku"]), d)
            pair_last[(c, l["sku"])] = d

    listings = []
    for (c, sku), listed_on in pair_first.items():
        counts = listed_on > first_paid[c]
        if c.upper().startswith(CARREFOUR_PREFIX) and sku.upper().startswith("X") and listed_on[5:7] == "01":
            counts = False
        listings.append((cust_no[c], sku, listed_on, pair_last[(c, sku)], "Y" if counts else "N"))
    listings.sort(key=lambda r: (r[2], r[0], r[1]))

    orders, sales = {}, defaultdict(lambda: [0.0, 0.0])
    for l in lines:
        if l["date"] >= orders_from:
            # Credit notes have no reliable id: one row per customer and day.
            key = ("CN", l["customer"], l["date"]) if l["credit"] else (l["invoice"], l["customer"], l["date"])
            o = orders.setdefault(key, {"lines": 0, "qty": 0.0, "value": 0.0, **{b: 0.0 for b in BRANDS}})
            o["lines"] += 1
            o["qty"] += l["qty"]
            o["value"] += l["total"]
            o[l["brand"]] += l["total"]
        if l["date"] >= sales_from:
            s = sales[(l["date"], cust_no[l["customer"]], l["sku"])]
            s[0] += l["qty"]
            s[1] += l["total"]

    customer_rows = []
    for c in customers:
        fb = first_paid_brand.get(c, {})
        customer_rows.append((cust_no[c], c, bdm.get(c, "Unassigned"), first_any[c], first_paid.get(c),
                              fb.get("AIWIBI"), fb.get("AIWINA"), fb.get("UBL"), last_paid.get(c)))
    item_rows = sorted((sku, *meta) for sku, meta in items.items())
    order_rows, number = [], 0
    for (inv, c, d), o in sorted(orders.items(), key=lambda kv: (kv[0][2], kv[0][1], kv[0][0])):
        credit = inv == "CN"
        if not credit:
            number += 1
        order_rows.append(("" if credit else number, d, cust_no[c], "" if credit else inv,
                           "Credit note" if credit else "Invoice", o["lines"], o["qty"], round(o["value"], 2),
                           *(round(o[b], 2) for b in BRANDS)))
    sales_rows = [(d, cn, sku, q, round(t, 2)) for (d, cn, sku), (q, t) in sorted(sales.items(), key=lambda kv: kv[0][0])]
    return customer_rows, item_rows, listings, order_rows, sales_rows


def as_date(value):
    return datetime.strptime(value, "%Y-%m-%d") if value else None


def clean(value):
    return int(value) if isinstance(value, float) and value.is_integer() else value


def write(path, info, sheets):
    wb = Workbook()
    wb.remove(wb.active)
    font = Font(name="Calibri", size=11)
    for name, header, rows, date_cols in [("Info", ["Item", "Value"], info, ())] + sheets:
        ws = wb.create_sheet(name)
        ws.append(header)
        for row in rows:
            ws.append([as_date(v) if i in date_cols else clean(v) for i, v in enumerate(row)])
        for col in date_cols:
            for cell in ws.iter_cols(min_col=col + 1, max_col=col + 1, min_row=2):
                for c in cell:
                    c.number_format = "dd mmm yyyy"
        ref = f"A1:{ws.cell(row=max(len(rows), 1) + 1, column=len(header)).coordinate}"
        table = Table(displayName=f"tbl{name}", ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
        ws.add_table(table)
        for col_cells in ws.columns:
            width = max(len(str(col_cells[0].value)), *(len(str(c.value or "")) for c in col_cells[1:200])) + 2
            ws.column_dimensions[col_cells[0].column_letter].width = min(width, 50)
            for c in col_cells:
                c.font = font
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"
    wb.save(tmp)
    os.replace(tmp, path)


def main(org=None, output_path=OUTPUT_PATH, today=None):
    org = org or os.environ.get("ZOHO_ORGANIZATION_ID") or "780891784"
    lines = prepare(ledger.load_ledger(ledger.ledger_path(org)))
    data_through = max(l["date"] for l in lines)
    today = today or date.today()
    year = int(data_through[:4])
    sales_from = date(year - 1, 7, 1).isoformat()   # a 1 Jan start date still has its 182-day look-back
    orders_from = date(year - 1, 1, 1).isoformat()  # plus the 365-day "gone quiet" look-back
    customers, items, listings, orders, sales = build(lines, sales_from, orders_from)
    info = [("Built", today.isoformat()), ("Data through", data_through),
            ("Sales detail from", sales_from), ("Orders detail from", orders_from), ("Brands", ", ".join(BRANDS)),
            ("Customers", len(customers)), ("Items", len(items)), ("Listings (customer-SKU pairs)", len(listings)),
            ("Orders", len(orders)), ("Sales lines", len(sales)),
            ("Credit notes", "Negative Sales lines; on Orders, one 'Credit note' row per customer per day with no Order number"),
            ("Order", "Invoice ID + Customer + Date"),
            ("BDM", "Active BDM on the customer's most recent invoice by an active BDM, else Unassigned"),
            ("Listing", "First paid purchase (Total > 0) of a customer-SKU pair; Counts = Y when after the customer's "
                        "first paid purchase day and not a Carrefour X-SKU first bought in January")]
    write(output_path, info, [
        ("Sales", ["Date", "Cust", "SKU", "Qty", "Total"], sales, (0,)),
        ("Orders", ["Order", "Date", "Cust", "Invoice ID", "Type", "Lines", "Qty", "Value",
                    "Value AIWIBI", "Value AIWINA", "Value UBL"], orders, (1,)),
        ("Customers", ["Cust", "Customer", "BDM", "First bought", "First paid", "First paid AIWIBI",
                       "First paid AIWINA", "First paid UBL", "Last paid"], customers, (3, 4, 5, 6, 7, 8)),
        ("Items", ["SKU", "Item Name", "Brand", "Category"], items, ()),
        ("Listings", ["Cust", "SKU", "Listed on", "Last bought", "Counts as listing"], listings, (2, 3)),
    ])
    print(f"Wrote {os.path.relpath(output_path, ledger.REPO_ROOT)}: {len(sales):,} sales lines from {sales_from}, "
          f"{len(orders):,} orders, {len(customers):,} customers, {len(listings):,} listings.")


if __name__ == "__main__":
    main(*sys.argv[1:2])
