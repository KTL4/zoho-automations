import os
import sys
import tempfile
import unittest

from openpyxl import load_workbook

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))

import sales_ledger as L  # noqa: E402


def row(doc_type="invoice", doc_id="1", modified="m1", day="2026-10-01", sku="A1", total="100.00",
        qty="2", category="", brand="UBL", invoice_id=None):
    return {"doc_type": doc_type, "doc_id": doc_id, "last_modified_time": modified,
            "invoice_id": doc_id if invoice_id is None else invoice_id, "date": day,
            "sales_person": "SP", "customer": "Cust", "sku": sku, "item_name": f"Item {sku}",
            "brand": brand, "category": category, "total": total, "quantity": qty}


class LineTotalTests(unittest.TestCase):
    def test_rate_times_quantity_less_discount_amount(self):
        # Matches the old export: Item Price 40000 x 12 = 480000 (VAT inclusive).
        self.assertEqual(L.line_total({"rate": 40000, "quantity": 12, "discount_amount": 0}), 480000)
        self.assertEqual(L.line_total({"rate": 100, "quantity": 3, "discount_amount": 30}), 270)

    def test_percentage_discount_when_no_amount(self):
        self.assertAlmostEqual(L.line_total({"rate": 100, "quantity": 2, "discount": "10%"}), 180)


class PlanSyncTests(unittest.TestCase):
    def test_new_changed_removed_and_unchanged(self):
        rows = [row(doc_id="same", modified="m1"), row(doc_id="edited", modified="m1"),
                row(doc_id="voided", modified="m1")]
        listed = {"invoice": {"same": ("2026-10-01", "m1"), "edited": ("2026-10-01", "m2"),
                              "new": ("2026-10-05", "m1")}}
        todo, removed = L.plan_sync(rows, listed, set(), "2026-09-21")
        self.assertEqual(sorted(todo), [("invoice", "edited"), ("invoice", "new")])
        self.assertEqual(removed, {("invoice", "voided")})

    def test_blank_modified_is_always_refetched(self):
        todo, _ = L.plan_sync([row(doc_id="x", modified="")], {"invoice": {"x": ("2026-10-01", "")}},
                              set(), "2026-09-21")
        self.assertEqual(todo, [("invoice", "x")])

    def test_frozen_history_invoices_are_never_fetched(self):
        todo, removed = L.plan_sync([row(doc_type="history", doc_id="")],
                                    {"invoice": {"old": ("2020-01-01", "m")}}, {"old"}, "2026-09-21")
        self.assertEqual((todo, removed), ([], set()))

    def test_credit_notes_untouched_when_listing_failed(self):
        rows = [row(doc_type="creditnote", doc_id="cn1")]
        todo, removed = L.plan_sync(rows, {"invoice": {}}, set(), "2026-09-21")
        self.assertEqual((todo, removed), ([], set()))


class DocumentRowsTests(unittest.TestCase):
    def setUp(self):
        self.lookups = L.Lookups([row(sku="A1", category="PET ITEMS")], {"OVR": "STATIONERY"},
                                 {"i2": {"brand": "AIWIBI", "category": "Baby Care"}})

    def test_credit_note_lines_are_negative(self):
        detail = {"date": "2026-10-03", "customer_name": "C", "line_items": [
            {"sku": "A1", "item_id": "i1", "name": "N", "rate": 50, "quantity": 2, "discount_amount": 0}]}
        out = L.document_rows("creditnote", "cn1", "m", detail, -1, self.lookups)
        self.assertEqual((out[0]["total"], out[0]["quantity"]), ("-100.00", "-2"))
        self.assertEqual(out[0]["category"], "PET ITEMS")

    def test_category_order_override_then_history_then_zoho(self):
        self.assertEqual(self.lookups.category("OVR", "i2"), "STATIONERY")
        self.assertEqual(self.lookups.category("A1", "i2"), "PET ITEMS")
        self.assertEqual(self.lookups.category("NEW", "i2"), "BABY CARE")
        self.assertEqual(self.lookups.category("NEW", "unknown"), "")

    def test_document_total_check(self):
        detail = {"total": 1100, "shipping_charge": 100, "line_items": [
            {"sku": "A1", "rate": 500, "quantity": 2, "discount_amount": 0}]}
        out = L.document_rows("invoice", "1", "m", detail, 1, self.lookups)
        self.assertAlmostEqual(L.check_document_total(detail, out), 0)

    def test_document_without_lines_keeps_placeholder(self):
        out = L.document_rows("invoice", "1", "m", {"date": "2026-10-01", "line_items": []}, 1, self.lookups)
        self.assertEqual(len(out), 1)
        self.assertEqual(out[0]["sku"], "")


class StorageTests(unittest.TestCase):
    def test_apply_changes_replaces_whole_document_and_sorts_newest_first(self):
        rows = [row(doc_id="1", sku="A"), row(doc_id="1", sku="B"), row(doc_id="2", day="2026-09-01")]
        out = L.apply_changes(rows, {("invoice", "2")},
                              {("invoice", "1"): [row(doc_id="1", sku="C", modified="m2")],
                               ("invoice", "3"): [row(doc_id="3", day="2026-11-01")]})
        self.assertEqual([(r["doc_id"], r["sku"]) for r in out], [("3", "A1"), ("1", "C")])

    def test_ledger_round_trip_by_year_and_xlsx(self):
        rows = L.sort_ledger([row(doc_id="9", day="2026-01-02"),
                              row(doc_type="history", doc_id="", day="2019-12-13",
                                  invoice_id="3.307694000008943e+18"),
                              row(doc_type="history", doc_id="", day="2020-05-01", invoice_id="",
                                  total="-10.00", qty="-1"),
                              {**row(doc_id="9", day="2026-01-02", sku="", total="", qty=""), "item_name": ""}])
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "ledger")
            L.save_ledger(path, rows)
            self.assertEqual(sorted(os.listdir(path)), ["2019.csv", "2020.csv", "2026.csv"])
            loaded = L.load_ledger(path)
            self.assertEqual(loaded, rows)
            xlsx = os.path.join(tmp, "out.xlsx")
            self.assertEqual(L.write_xlsx(loaded, xlsx), 3)  # placeholder row skipped
            values = list(load_workbook(xlsx).active.iter_rows(values_only=True))
            self.assertEqual(values[0][0], "Invoice ID")
            self.assertEqual(values[1][0], "9")  # full id kept as text
            self.assertEqual(values[2][0], None)  # history credit note: blank id
            self.assertEqual(values[2][8], -10)
            self.assertIsInstance(values[3][0], float)  # history: rounded number as before


if __name__ == "__main__":
    unittest.main()
