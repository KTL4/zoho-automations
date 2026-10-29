import datetime
import os
import sys
import tempfile
import time
import unittest
from unittest.mock import MagicMock, patch

from openpyxl import load_workbook

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))

import container_sellout_report as report  # noqa: E402

D = datetime.date
PATTERNS = report.supplier_patterns()


def lot(seq, day, qty_in, on_hand, sku="A1", number="C1", contact="Uncle Bill Ltd", ttype="Bill", cost=10.0):
    return {"seq": seq, "date": day, "transaction_type": ttype, "transaction_number": number,
            "contact_name": contact, "lot_number": "", "sku": sku, "item_name": f"Item {sku}",
            "quantity_in": float(qty_in), "quantity_on_hand": float(on_hand), "unit_cost": cost}


class AllocateSalesFifoTests(unittest.TestCase):
    def test_oldest_lot_fills_first_and_room_is_capped_by_zoho_consumed(self):
        lots = [lot(0, D(2026, 1, 1), 100, 40), lot(1, D(2026, 2, 1), 100, 100)]
        allocations, unmatched = report.allocate_sales_fifo(lots, {D(2026, 2, 10): 70})
        # Lot 0 only had 60 consumed per Zoho; lot 1 shows nothing consumed, so 10 is unmatched.
        self.assertEqual(allocations[0], [(D(2026, 2, 10), 60)])
        self.assertNotIn(1, allocations)
        self.assertEqual(unmatched, 10)

    def test_sale_never_goes_to_a_lot_that_had_not_arrived(self):
        lots = [lot(0, D(2026, 3, 1), 50, 0)]
        allocations, unmatched = report.allocate_sales_fifo(lots, {D(2026, 2, 1): 5, D(2026, 3, 2): 5})
        self.assertEqual(allocations[0], [(D(2026, 3, 2), 5)])
        self.assertEqual(unmatched, 5)

    def test_net_return_is_carried_into_next_sale(self):
        lots = [lot(0, D(2026, 1, 1), 50, 0)]
        allocations, _ = report.allocate_sales_fifo(lots, {D(2026, 1, 5): -3, D(2026, 1, 6): 10})
        self.assertEqual(allocations[0], [(D(2026, 1, 6), 7)])


class SummaryTests(unittest.TestCase):
    def test_container_metrics(self):
        lots = [lot(0, D(2026, 1, 1), 100, 10), lot(1, D(2026, 1, 1), 100, 100, sku="B2")]
        daily = {"A1": {D(2026, 1, 11): 50, D(2026, 2, 20): 30}}
        rows, sku_rows, unmatched = report.build_report_data(lots, daily, PATTERNS, D(2026, 4, 1))
        self.assertEqual(len(rows), 1)
        r = rows[0]
        self.assertEqual((r["supplier"], r["container"], r["skus"]), ("UBL", "C1", 2))
        self.assertEqual(r["units_in"], 200)
        self.assertAlmostEqual(r["sellout"], 90 / 200)
        self.assertEqual(r["sold"], 80)
        self.assertEqual(r["other_out"], 10)
        self.assertAlmostEqual(r["sold_by"][30], 50 / 200)
        self.assertAlmostEqual(r["sold_by"][60], 80 / 200)
        self.assertIsNone(r["sold_by"][180])  # container only 90 days old
        self.assertIsNone(r["days_to"][0.5])   # 40% sold, never reached 50%
        a1 = next(s for s in sku_rows if s["sku"] == "A1")
        self.assertEqual(a1["days_to"][0.5], 10)
        self.assertEqual(a1["days_to"][0.8], 50)
        self.assertEqual(unmatched, 0)

    def test_non_container_lots_are_fifo_competitors_but_not_reported(self):
        lots = [lot(0, D(2025, 12, 1), 20, 0, ttype="Opening Stock", contact=""),
                lot(1, D(2026, 1, 1), 100, 70)]
        rows, _, _ = report.build_report_data(lots, {"A1": {D(2026, 1, 5): 50}}, PATTERNS, D(2026, 4, 1))
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["sold"], 30)  # first 20 came out of opening stock

    def test_supplier_grouping(self):
        self.assertEqual(report.supplier_of("UNCLE BILL LIMITED", PATTERNS), "UBL")
        self.assertEqual(report.supplier_of("Aiwibi Trading", PATTERNS), "Aiwibi")
        self.assertEqual(report.supplier_of("EFRIS STOCK", PATTERNS), "EFRIS STOCK")
        self.assertIsNone(report.supplier_of("Local Hardware", PATTERNS))


def fake_response(payload, status=200):
    response = MagicMock()
    response.status_code = status
    response.json.return_value = payload
    response.headers = {}
    response.raise_for_status = MagicMock()
    return response


class EndToEndTests(unittest.TestCase):
    def test_full_run_writes_report_and_cache_then_reuses_cache(self):
        lot_rows = [
            {"date": "2026-01-01", "transaction_type": "Bill", "transaction_number": "C30", "contact_name": "Uncle Bill",
             "sku": "A1", "item_name": "Widget", "quantity_in": "1,000", "quantity_on_hand": "400", "cost_price_per_unit": "5"},
            {"date": "2026-01-01", "transaction_type": "Invoice", "transaction_number": "INV-1", "contact_name": "Shop",
             "sku": "A1", "quantity_in": 0, "quantity_on_hand": 0},
        ] + [{"date": "2025-01-01", "transaction_type": "Opening Stock", "transaction_number": f"OS{i}",
              "contact_name": "", "sku": f"Z{i}", "quantity_in": 1, "quantity_on_hand": 1} for i in range(60)]
        invoices = [{"invoice_id": "i1", "date": "2026-01-10", "status": "paid", "last_modified_time": "t1"},
                    {"invoice_id": "i2", "date": "2026-01-12", "status": "void", "last_modified_time": "t1"}]
        detail_calls = []

        def get(url, params=None, **_):
            if url.endswith("/reports/fifocostlottracking"):
                return fake_response({"fifo_cost_lot_tracking": lot_rows, "page_context": {"has_more_page": False}})
            if url.endswith("/invoices"):
                return fake_response({"invoices": invoices, "page_context": {"has_more_page": False}})
            if url.endswith("/creditnotes"):
                return fake_response({"creditnotes": [], "page_context": {"has_more_page": False}})
            raise AssertionError(url)

        def fetch_detail(session, api_domain, org, token_store, limiter, url, result_key, label):
            detail_calls.append(url)
            return {"line_items": [{"sku": "a1", "quantity": 600}]}

        session = MagicMock()
        session.get.side_effect = get
        session.headers = {}
        with tempfile.TemporaryDirectory() as tmp, \
                patch.object(report.requests, "Session", return_value=session), \
                patch.object(report.common, "get_access_token", return_value=("tok", "https://api")), \
                patch.object(report.common, "fetch_detail", side_effect=fetch_detail), \
                patch.object(report.time, "sleep"), \
                patch.dict(os.environ, {"ZOHO_ORGANIZATION_ID": "42", "REPORT_AS_OF_DATE": "2026-04-01"}):
            cwd = os.getcwd()
            os.chdir(tmp)
            try:
                report.main()
                report.main()  # second run: the invoice is cached, no detail call
                wb = load_workbook(os.path.join("reports", "Container_Sellout_01_04_2026.xlsx"))
                self.assertTrue(os.path.exists(os.path.join("data", "sales_lines_42.csv")))
                self.assertTrue(os.path.exists(os.path.join("data", "fifo_lots_42.csv")))
            finally:
                os.chdir(cwd)

        self.assertEqual(len(detail_calls), 1)
        sheet = wb["Containers"]
        headers = [c.value for c in sheet[1]]
        row = dict(zip(headers, [c.value for c in sheet[2]]))
        self.assertEqual(row["Container"], "C30")
        self.assertEqual(row["Units In"], 1000)
        self.assertAlmostEqual(row["Sellout % (Zoho FIFO)"], 0.6)
        self.assertEqual(row["Units Sold"], 600)
        self.assertAlmostEqual(row["Sold % by Day 30"], 0.6)
        self.assertEqual(row["Days to 50% Sold"], 9)
        self.assertEqual(sheet.max_row, 2)
        self.assertEqual(wb.sheetnames, ["Containers", "Container x SKU", "Read Me"])


if __name__ == "__main__":
    unittest.main()
