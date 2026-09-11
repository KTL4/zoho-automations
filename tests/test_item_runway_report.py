import datetime
import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))

import item_runway_report as report  # noqa: E402


class ComputeWeekBucketsTests(unittest.TestCase):
    def test_35_periods_ending_with_previous_full_week(self):
        # Wednesday 2026-10-01. That week's Monday is 2026-09-28, so the
        # "previous week" is 2026-09-21 (Mon) .. 2026-09-27 (Sun).
        today = datetime.date(2026, 10, 1)
        buckets = report.compute_week_buckets(today)

        self.assertEqual(len(buckets), 35)
        self.assertEqual(buckets[-1], (datetime.date(2026, 9, 21), datetime.date(2026, 9, 27)))
        # Oldest bucket is exactly 34 weeks before the most recent one.
        self.assertEqual(buckets[0], (datetime.date(2026, 1, 26), datetime.date(2026, 2, 1)))
        # Buckets are chronological and contiguous, 7 days apart.
        for (start, end) in buckets:
            self.assertEqual((end - start).days, 6)
        for i in range(1, len(buckets)):
            self.assertEqual(buckets[i][0] - buckets[i - 1][0], datetime.timedelta(days=7))

    def test_previous_week_always_precedes_run_date(self):
        for offset in range(7):
            today = datetime.date(2026, 10, 1) + datetime.timedelta(days=offset)
            buckets = report.compute_week_buckets(today)
            self.assertLess(buckets[-1][1], today)

    def test_month_used_for_filename_is_previous_month(self):
        # Regardless of which weekday the 1st falls on, the previous week
        # (and therefore the report's month) should be the prior month.
        for year_month in [(2026, 10), (2026, 1), (2027, 3)]:
            today = datetime.date(year_month[0], year_month[1], 1)
            buckets = report.compute_week_buckets(today)
            prev_week_end_month = buckets[-1][1].month
            expected_prev_month = 12 if today.month == 1 else today.month - 1
            self.assertEqual(prev_week_end_month, expected_prev_month)


class WeekLabelTests(unittest.TestCase):
    def test_label_uses_iso_week_number(self):
        self.assertEqual(report.week_label(datetime.date(2026, 9, 21)), "WK 39")


class AggregateSalesTests(unittest.TestCase):
    def setUp(self):
        self.buckets = report.compute_week_buckets(datetime.date(2026, 10, 1))

    def test_sums_quantities_into_correct_week_bucket(self):
        invoices = [
            {
                "date": "2026-09-22",  # falls in the most recent bucket
                "line_items": [
                    {"item_id": "item-1", "quantity": 3},
                    {"item_id": "item-2", "quantity": 1},
                ],
            },
            {
                "date": "2026-09-25",  # same week as above
                "line_items": [{"item_id": "item-1", "quantity": 2}],
            },
            {
                "date": "2026-01-28",  # oldest bucket (2026-01-26 .. 2026-02-01)
                "line_items": [{"item_id": "item-1", "quantity": 10}],
            },
        ]
        sales = report.aggregate_sales(invoices, self.buckets)

        self.assertEqual(sales["item-1"][-1], 5)
        self.assertEqual(sales["item-2"][-1], 1)
        self.assertEqual(sales["item-1"][0], 10)
        # Untouched buckets stay zero.
        self.assertEqual(sum(sales["item-2"]) , 1)

    def test_skips_invoices_outside_window_and_missing_dates(self):
        invoices = [
            {"date": "2020-01-01", "line_items": [{"item_id": "item-1", "quantity": 99}]},
            {"line_items": [{"item_id": "item-1", "quantity": 5}]},  # no date
            None,
        ]
        sales = report.aggregate_sales(invoices, self.buckets)
        self.assertEqual(sales, {})

    def test_line_items_without_item_id_are_ignored(self):
        invoices = [{"date": "2026-09-22", "line_items": [{"quantity": 5}]}]
        sales = report.aggregate_sales(invoices, self.buckets)
        self.assertEqual(sales, {})


class FetchInvoiceIdsInWindowTests(unittest.TestCase):
    def _make_response(self, invoices, has_more_page):
        response = MagicMock()
        response.raise_for_status = MagicMock()
        response.json.return_value = {
            "invoices": invoices,
            "page_context": {"has_more_page": has_more_page},
        }
        return response

    def test_paginates_and_stops_past_window_start(self):
        window_start = datetime.date(2026, 9, 1)
        window_end = datetime.date(2026, 9, 30)

        page1 = self._make_response(
            [
                {"invoice_id": "future", "date": "2026-10-05", "status": "paid"},  # newer than window
                {"invoice_id": "inv-1", "date": "2026-09-20", "status": "paid"},
                {"invoice_id": "inv-draft", "date": "2026-09-15", "status": "draft"},
                {"invoice_id": "inv-void", "date": "2026-09-10", "status": "void"},
            ],
            has_more_page=True,
        )
        page2 = self._make_response(
            [
                {"invoice_id": "inv-2", "date": "2026-09-05", "status": "sent"},
                {"invoice_id": "too-old", "date": "2026-08-20", "status": "paid"},  # should stop here
                {"invoice_id": "never-fetched", "date": "2026-08-01", "status": "paid"},
            ],
            has_more_page=True,  # would keep going if not for the window_start cutoff
        )

        session = MagicMock()
        session.get.side_effect = [page1, page2]

        ids = report.fetch_invoice_ids_in_window(session, "https://api", "org1", window_start, window_end)

        self.assertEqual(ids, ["inv-1", "inv-2"])
        self.assertEqual(session.get.call_count, 2)

    def test_stops_when_no_more_pages(self):
        window_start = datetime.date(2026, 9, 1)
        window_end = datetime.date(2026, 9, 30)
        page1 = self._make_response(
            [{"invoice_id": "inv-1", "date": "2026-09-15", "status": "paid"}],
            has_more_page=False,
        )
        session = MagicMock()
        session.get.side_effect = [page1]

        ids = report.fetch_invoice_ids_in_window(session, "https://api", "org1", window_start, window_end)
        self.assertEqual(ids, ["inv-1"])
        self.assertEqual(session.get.call_count, 1)


class ExtractBrandTests(unittest.TestCase):
    def test_prefers_top_level_brand_field(self):
        item = {"brand": "Acme", "custom_fields": [{"label": "Brand", "value": "Other"}]}
        self.assertEqual(report.extract_brand(item), "Acme")

    def test_falls_back_to_custom_field(self):
        item = {"custom_fields": [{"label": "Brand Name", "value": "Acme"}]}
        self.assertEqual(report.extract_brand(item), "Acme")

    def test_missing_brand_returns_empty_string(self):
        self.assertEqual(report.extract_brand({}), "")


class BuildRowsTests(unittest.TestCase):
    def test_includes_zero_sales_items_and_sorts_by_name(self):
        catalog = {
            "item-1": {"SKU": "SKU1", "Item Name": "Zebra Widget", "Brand": "Acme"},
            "item-2": {"SKU": "SKU2", "Item Name": "Apple Widget", "Brand": "Acme"},
        }
        buckets = [(datetime.date(2026, 1, 1), datetime.date(2026, 1, 7))]
        sales = {"item-1": [5]}

        rows = report.build_rows(catalog, sales, buckets)

        self.assertEqual([row[1] for row in rows], ["Apple Widget", "Zebra Widget"])
        apple_row = next(row for row in rows if row[1] == "Apple Widget")
        self.assertEqual(apple_row[3], 0)
        zebra_row = next(row for row in rows if row[1] == "Zebra Widget")
        self.assertEqual(zebra_row[3], 5)

    def test_sales_for_items_outside_catalog_are_dropped_with_warning(self):
        catalog = {"item-1": {"SKU": "SKU1", "Item Name": "Widget", "Brand": "Acme"}}
        buckets = [(datetime.date(2026, 1, 1), datetime.date(2026, 1, 7))]
        sales = {"item-1": [5], "discontinued-item": [10]}

        rows = report.build_rows(catalog, sales, buckets)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][1], "Widget")


class WriteExcelTests(unittest.TestCase):
    def test_writes_header_frozen_panes_and_number_format(self):
        import openpyxl

        rows = [
            ["SKU1", "Widget A", "Acme", 3, 0],
            ["SKU2", "Widget B", "Acme", 0, 7],
        ]
        week_columns = ["WK 1", "WK 2"]

        with patch("os.makedirs"):
            output_path = os.path.join(
                os.path.dirname(__file__), "_tmp_test_output.xlsx"
            )
            report.write_excel(rows, week_columns, output_path)

        try:
            workbook = openpyxl.load_workbook(output_path)
            sheet = workbook.active

            self.assertEqual(
                [cell.value for cell in sheet[1]],
                ["SKU", "Item Name", "Brand", "WK 1", "WK 2"],
            )
            self.assertTrue(sheet[1][0].font.bold)
            self.assertEqual(sheet.freeze_panes, "B2")
            self.assertEqual(sheet.cell(row=2, column=4).number_format, "#,##0")
            self.assertEqual(sheet.cell(row=2, column=1).value, "SKU1")
        finally:
            if os.path.exists(output_path):
                os.remove(output_path)


class MainIntegrationTest(unittest.TestCase):
    """End-to-end wiring check with every HTTP call mocked."""

    @patch.dict(
        os.environ,
        {
            "ZOHO_CLIENT_ID": "id",
            "ZOHO_CLIENT_SECRET": "secret",
            "ZOHO_REFRESH_TOKEN": "refresh",
            "ZOHO_ORGANIZATION_ID": "org1",
            "REPORT_TZ": "UTC",
            "RATE_LIMIT_PER_MINUTE": "6000",
        },
        clear=True,
    )
    @patch("item_runway_report.datetime")
    @patch("item_runway_report.requests.Session")
    @patch("item_runway_report.requests.post")
    def test_main_writes_expected_file(self, mock_post, mock_session_cls, mock_datetime):
        # Fix "today" for a deterministic window/filename.
        fixed_today = datetime.datetime(2026, 10, 1, 6, 0, tzinfo=datetime.timezone.utc)
        mock_datetime.now.return_value = fixed_today
        mock_datetime.strptime = datetime.datetime.strptime

        mock_post.return_value = MagicMock(
            raise_for_status=MagicMock(),
            json=MagicMock(return_value={"access_token": "tok", "api_domain": "https://api"}),
        )

        session = MagicMock()
        mock_session_cls.return_value = session

        items_page = MagicMock()
        items_page.raise_for_status = MagicMock()
        items_page.json.return_value = {
            "items": [{"item_id": "item-1"}],
            "page_context": {"has_more_page": False},
        }
        item_detail = MagicMock()
        item_detail.status_code = 200
        item_detail.raise_for_status = MagicMock()
        item_detail.json.return_value = {
            "item": {"item_id": "item-1", "sku": "SKU1", "name": "Widget", "brand": "Acme"}
        }

        invoices_page = MagicMock()
        invoices_page.raise_for_status = MagicMock()
        invoices_page.json.return_value = {
            "invoices": [{"invoice_id": "inv-1", "date": "2026-09-22", "status": "paid"}],
            "page_context": {"has_more_page": False},
        }
        invoice_detail = MagicMock()
        invoice_detail.status_code = 200
        invoice_detail.raise_for_status = MagicMock()
        invoice_detail.json.return_value = {
            "invoice": {
                "invoice_id": "inv-1",
                "date": "2026-09-22",
                "line_items": [{"item_id": "item-1", "quantity": 4}],
            }
        }

        def get_side_effect(url, params=None, headers=None, timeout=None):
            if url.endswith("/inventory/v1/items"):
                return items_page
            if url.endswith("/inventory/v1/items/item-1"):
                return item_detail
            if url.endswith("/inventory/v1/invoices"):
                return invoices_page
            if url.endswith("/inventory/v1/invoices/inv-1"):
                return invoice_detail
            raise AssertionError(f"Unexpected GET {url}")

        session.get.side_effect = get_side_effect

        output_path = "reports/Sales_by_item(sep).xlsx"
        if os.path.exists(output_path):
            os.remove(output_path)

        try:
            report.main()
            self.assertTrue(os.path.exists(output_path))

            import openpyxl

            workbook = openpyxl.load_workbook(output_path)
            sheet = workbook.active
            self.assertEqual(sheet.cell(row=1, column=1).value, "SKU")
            self.assertEqual(sheet.cell(row=2, column=1).value, "SKU1")
            self.assertEqual(sheet.cell(row=2, column=2).value, "Widget")
            self.assertEqual(sheet.cell(row=2, column=3).value, "Acme")
            self.assertEqual(sheet.cell(row=2, column=sheet.max_column).value, 4)
        finally:
            if os.path.exists(output_path):
                os.remove(output_path)


if __name__ == "__main__":
    unittest.main()
