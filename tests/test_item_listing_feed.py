import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))

import item_listing_feed as F  # noqa: E402


def line(day, customer, sku, total, qty=1, brand="UBL", sp="SAM AGABA", invoice="INV1", doc_type="invoice"):
    return {"doc_type": doc_type, "doc_id": invoice, "invoice_id": invoice if doc_type != "creditnote" else "",
            "date": day, "customer": customer, "sku": sku, "item_name": f"Item {sku}", "brand": brand,
            "category": "CAT", "sales_person": sp, "total": str(total), "quantity": str(qty)}


def run(rows, sales_from="2000-01-01", orders_from="2000-01-01"):
    return F.build(F.prepare(rows), sales_from, orders_from)


class FeedTests(unittest.TestCase):
    def test_other_brands_and_zero_lines_dropped(self):
        customers, items, _l, _o, sales = run([line("2026-01-02", "A", "S1", 100),
                                               line("2026-01-02", "A", "S2", 100, brand="OTHER"),
                                               line("2026-01-02", "A", "S3", 0, qty=0)])
        self.assertEqual([i[0] for i in items], ["S1"])
        self.assertEqual(len(sales), 1)

    def test_bdm_is_latest_active_bdm_else_unassigned(self):
        customers, *_ = run([line("2026-01-01", "A", "S1", 10, sp="JOAN MBABAZI"),
                             line("2026-02-01", "A", "S1", 10, sp="Somebody Else", invoice="INV2"),
                             line("2026-01-01", "B", "S1", 10, sp="Somebody Else", invoice="INV3")])
        bdm = {c[1]: c[2] for c in customers}
        self.assertEqual(bdm, {"A": "JOAN MBABAZI", "B": "Unassigned"})

    def test_listing_rules(self):
        rows = [line("2025-03-01", "SHOP", "S1", 10),
                line("2025-03-01", "SHOP", "S2", 10),           # first day: new customer, not a listing
                line("2025-04-01", "SHOP", "S3", 10, invoice="I2"),  # later: listing
                line("2025-03-01", "MAJID AL FUTTAIM HYPERMARKETS", "S1", 10, invoice="I3"),
                line("2026-01-05", "MAJID AL FUTTAIM HYPERMARKETS", "X100", 10, invoice="I4"),  # Christmas item
                line("2026-02-05", "MAJID AL FUTTAIM HYPERMARKETS", "X200", 10, invoice="I5")]
        customers, _i, listings, *_ = run(rows)
        names = {c[0]: c[1] for c in customers}
        flags = {(names[l[0]], l[1]): l[4] for l in listings}
        self.assertEqual(flags[("SHOP", "S2")], "N")
        self.assertEqual(flags[("SHOP", "S3")], "Y")
        self.assertEqual(flags[("MAJID AL FUTTAIM HYPERMARKETS", "X100")], "N")
        self.assertEqual(flags[("MAJID AL FUTTAIM HYPERMARKETS", "X200")], "Y")

    def test_listing_date_is_first_paid_purchase(self):
        rows = [line("2025-01-01", "A", "S0", 10),
                line("2025-02-01", "A", "S1", 0, qty=1, invoice="I2"),   # free sample first
                line("2025-03-01", "A", "S1", 50, invoice="I3")]
        _c, _i, listings, *_ = run(rows)
        self.assertEqual([l[2] for l in listings if l[1] == "S1"], ["2025-03-01"])

    def test_orders_split_by_brand_and_credit_notes_have_no_number(self):
        rows = [line("2026-01-02", "A", "S1", 100, brand="UBL"),
                line("2026-01-02", "A", "S2", 40, brand="AIWIBI"),
                line("2026-01-03", "A", "S1", -30, qty=-1, doc_type="creditnote", invoice="cn1")]
        *_, orders, sales = run(rows)
        inv, cn = orders
        self.assertEqual((inv[0], inv[4], inv[7], inv[8], inv[10]), (1, "Invoice", 140, 40, 100))
        self.assertEqual((cn[0], cn[3], cn[4], cn[7]), ("", "", "Credit note", -30))
        self.assertEqual(sum(s[4] for s in sales), 110)

    def test_detail_windows(self):
        rows = [line("2025-01-10", "A", "S1", 10), line("2025-08-01", "A", "S1", 20, invoice="I2")]
        customers, _i, listings, orders, sales = run(rows, sales_from="2025-07-01", orders_from="2025-01-01")
        self.assertEqual(len(sales), 1)
        self.assertEqual(len(orders), 2)
        self.assertEqual(customers[0][4], "2025-01-10")  # first paid still from full history


if __name__ == "__main__":
    unittest.main()
