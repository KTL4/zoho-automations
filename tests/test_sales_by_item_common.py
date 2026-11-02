import datetime
import os
import sys
import unittest
from unittest.mock import MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))

import sales_by_item_common as common  # noqa: E402


class ComputeWeekBucketsEndingByTests(unittest.TestCase):
    def test_reference_date_on_sunday_includes_that_weeks_own_end(self):
        # 2026-09-13 is itself a Sunday.
        buckets = common.compute_week_buckets_ending_by(datetime.date(2026, 9, 13), 1)
        self.assertEqual(buckets, [(datetime.date(2026, 9, 7), datetime.date(2026, 9, 13))])

    def test_reference_date_midweek_uses_previous_sunday(self):
        # 2026-09-16 is a Wednesday, inside the still-open Sep14-20 week.
        buckets = common.compute_week_buckets_ending_by(datetime.date(2026, 9, 16), 1)
        self.assertEqual(buckets, [(datetime.date(2026, 9, 7), datetime.date(2026, 9, 13))])

    def test_reference_date_on_monday_uses_previous_sunday(self):
        buckets = common.compute_week_buckets_ending_by(datetime.date(2026, 9, 14), 1)
        self.assertEqual(buckets, [(datetime.date(2026, 9, 7), datetime.date(2026, 9, 13))])

    def test_multiple_periods_are_chronological_and_contiguous(self):
        buckets = common.compute_week_buckets_ending_by(datetime.date(2026, 9, 13), 36)

        self.assertEqual(len(buckets), 36)
        self.assertEqual(buckets[-1], (datetime.date(2026, 9, 7), datetime.date(2026, 9, 13)))
        for (start, end) in buckets:
            self.assertEqual((end - start).days, 6)
        for i in range(1, len(buckets)):
            self.assertEqual(buckets[i][0] - buckets[i - 1][0], datetime.timedelta(days=7))

    def test_matches_monthly_reports_previous_full_week_semantics(self):
        # The monthly report anchors on (today - 1 day) to get "the last full
        # week strictly before today's week" -- verify against the values
        # already validated by two live production runs.
        today = datetime.date(2026, 10, 1)
        buckets = common.compute_week_buckets_ending_by(today - datetime.timedelta(days=1), 14)

        self.assertEqual(len(buckets), 14)
        self.assertEqual(buckets[-1], (datetime.date(2026, 9, 21), datetime.date(2026, 9, 27)))
        self.assertEqual(buckets[0], (datetime.date(2026, 6, 22), datetime.date(2026, 6, 28)))


class WeekLabelTests(unittest.TestCase):
    def test_label_uses_iso_week_number(self):
        self.assertEqual(common.week_label(datetime.date(2026, 9, 21)), "WK 39")


class AggregateNetSalesTests(unittest.TestCase):
    def setUp(self):
        self.buckets = common.compute_week_buckets_ending_by(datetime.date(2026, 9, 27), 14)

    def _cache(self, **entries):
        return entries

    def test_sums_invoice_quantities_into_correct_week_bucket(self):
        cache = {
            ("invoice", "inv-1"): {"date": "2026-09-22", "modified": "x", "lines": [("item-1", 3), ("item-2", 1)]},
            ("invoice", "inv-2"): {"date": "2026-09-25", "modified": "x", "lines": [("item-1", 2)]},  # same week
            ("invoice", "inv-3"): {"date": "2026-06-24", "modified": "x", "lines": [("item-1", 10)]},  # oldest bucket
        }
        sales = common.aggregate_net_sales(cache, self.buckets)

        self.assertEqual(sales["item-1"][-1], 5)
        self.assertEqual(sales["item-2"][-1], 1)
        self.assertEqual(sales["item-1"][0], 10)
        self.assertEqual(sum(sales["item-2"]), 1)

    def test_credit_notes_are_subtracted(self):
        cache = {
            ("invoice", "inv-1"): {"date": "2026-09-22", "modified": "x", "lines": [("item-1", 10)]},
            ("creditnote", "cn-1"): {"date": "2026-09-23", "modified": "x", "lines": [("item-1", 4)]},
        }
        sales = common.aggregate_net_sales(cache, self.buckets)
        self.assertEqual(sales["item-1"][-1], 6)

    def test_credit_note_nets_off_its_own_week_not_the_invoice_week(self):
        cache = {
            ("invoice", "inv-1"): {"date": "2026-06-24", "modified": "x", "lines": [("item-1", 10)]},  # oldest week
            ("creditnote", "cn-1"): {"date": "2026-09-22", "modified": "x", "lines": [("item-1", 4)]},  # most recent week
        }
        sales = common.aggregate_net_sales(cache, self.buckets)
        self.assertEqual(sales["item-1"][0], 10)
        self.assertEqual(sales["item-1"][-1], -4)

    def test_skips_entries_outside_window_and_with_bad_dates(self):
        cache = {
            ("invoice", "old"): {"date": "2020-01-01", "modified": "x", "lines": [("item-1", 99)]},
            ("invoice", "bad-date"): {"date": "not-a-date", "modified": "x", "lines": [("item-1", 5)]},
        }
        sales = common.aggregate_net_sales(cache, self.buckets)
        self.assertEqual(sales, {})


class ListDocumentsInWindowTests(unittest.TestCase):
    def _make_response(self, documents, has_more_page):
        response = MagicMock()
        response.raise_for_status = MagicMock()
        response.json.return_value = {
            "invoices": documents,
            "page_context": {"has_more_page": has_more_page},
        }
        return response

    def test_paginates_and_stops_past_window_start(self):
        window_start = datetime.date(2026, 9, 1)
        window_end = datetime.date(2026, 9, 30)

        page1 = self._make_response(
            [
                {"invoice_id": "future", "date": "2026-10-05", "status": "paid"},  # newer than window
                {"invoice_id": "inv-1", "date": "2026-09-20", "status": "paid", "last_modified_time": "t1"},
                {"invoice_id": "inv-draft", "date": "2026-09-15", "status": "draft"},
                {"invoice_id": "inv-void", "date": "2026-09-10", "status": "void"},
            ],
            has_more_page=True,
        )
        page2 = self._make_response(
            [
                {"invoice_id": "inv-2", "date": "2026-09-05", "status": "sent", "last_modified_time": "t2"},
                {"invoice_id": "too-old", "date": "2026-08-20", "status": "paid"},  # should stop here
                {"invoice_id": "never-fetched", "date": "2026-08-01", "status": "paid"},
            ],
            has_more_page=True,  # would keep going if not for the window_start cutoff
        )

        session = MagicMock()
        session.get.side_effect = [page1, page2]

        docs = common.list_documents_in_window(
            session, "https://api", "org1", "invoices", "invoices", "invoice_id", window_start, window_end
        )

        self.assertEqual(docs, {"inv-1": ("2026-09-20", "t1"), "inv-2": ("2026-09-05", "t2")})
        self.assertEqual(session.get.call_count, 2)

    def test_stops_when_no_more_pages(self):
        window_start = datetime.date(2026, 9, 1)
        window_end = datetime.date(2026, 9, 30)
        page1 = self._make_response(
            [{"invoice_id": "inv-1", "date": "2026-09-15", "status": "paid", "last_modified_time": "t1"}],
            has_more_page=False,
        )
        session = MagicMock()
        session.get.side_effect = [page1]

        docs = common.list_documents_in_window(
            session, "https://api", "org1", "invoices", "invoices", "invoice_id", window_start, window_end
        )
        self.assertEqual(docs, {"inv-1": ("2026-09-15", "t1")})
        self.assertEqual(session.get.call_count, 1)


class DocCacheTests(unittest.TestCase):
    def test_save_then_load_round_trips(self):
        cache = {
            ("invoice", "inv-1"): {"date": "2026-09-20", "modified": "t1", "lines": [("item-1", 3.0), ("item-2", 1.0)]},
            ("creditnote", "cn-1"): {"date": "2026-09-21", "modified": "t2", "lines": [("item-1", 1.0)]},
        }
        path = os.path.join(os.path.dirname(__file__), "_tmp_test_cache.csv")
        try:
            common.save_doc_cache(path, cache)
            loaded = common.load_doc_cache(path)
            self.assertEqual(loaded, cache)
        finally:
            if os.path.exists(path):
                os.remove(path)

    def test_document_with_no_lines_round_trips_as_cached_with_empty_lines(self):
        cache = {("invoice", "empty-inv"): {"date": "2026-09-20", "modified": "t1", "lines": []}}
        path = os.path.join(os.path.dirname(__file__), "_tmp_test_cache_empty.csv")
        try:
            common.save_doc_cache(path, cache)
            loaded = common.load_doc_cache(path)
            self.assertEqual(loaded, cache)
        finally:
            if os.path.exists(path):
                os.remove(path)

    def test_load_missing_file_returns_empty_cache(self):
        self.assertEqual(common.load_doc_cache("/nonexistent/path.csv"), {})


class SyncDocumentsInWindowTests(unittest.TestCase):
    def _list_response(self, documents, has_more_page=False):
        response = MagicMock()
        response.raise_for_status = MagicMock()
        response.json.return_value = {"invoices": documents, "page_context": {"has_more_page": has_more_page}}
        return response

    def _detail_response(self, invoice):
        response = MagicMock(status_code=200)
        response.raise_for_status = MagicMock()
        response.json.return_value = {"invoice": invoice}
        return response

    def _sync(self, session, cache, window_start=datetime.date(2026, 9, 1), window_end=datetime.date(2026, 9, 30)):
        token_store = MagicMock()
        token_store.get.return_value = "tok"
        rate_limiter = MagicMock()
        return common.sync_documents_in_window(
            session, "https://api", "org1", token_store, rate_limiter,
            cache, "invoice", "invoices", "invoices", "invoice_id", "invoice", window_start, window_end,
        )

    def test_new_documents_are_listed_then_fetched(self):
        session = MagicMock()
        session.get.side_effect = [
            self._list_response([{"invoice_id": "inv-1", "date": "2026-09-20", "status": "paid", "last_modified_time": "t1"}]),
            self._detail_response({"invoice_id": "inv-1", "line_items": [{"item_id": "item-1", "quantity": 3}]}),
        ]

        cache = self._sync(session, {})

        self.assertEqual(cache[("invoice", "inv-1")]["lines"], [("item-1", 3)])
        self.assertEqual(cache[("invoice", "inv-1")]["modified"], "t1")

    def test_unchanged_cached_document_is_not_refetched(self):
        session = MagicMock()
        session.get.side_effect = [
            self._list_response([{"invoice_id": "inv-1", "date": "2026-09-20", "status": "paid", "last_modified_time": "t1"}]),
            # No detail response queued -- a second session.get call would raise StopIteration, failing the test.
        ]
        cache = {("invoice", "inv-1"): {"date": "2026-09-20", "modified": "t1", "lines": [("item-1", 3)]}}

        result = self._sync(session, cache)

        self.assertEqual(session.get.call_count, 1)  # only the list call, no detail fetch
        self.assertEqual(result[("invoice", "inv-1")]["lines"], [("item-1", 3)])

    def test_changed_last_modified_time_triggers_refetch(self):
        session = MagicMock()
        session.get.side_effect = [
            self._list_response([{"invoice_id": "inv-1", "date": "2026-09-20", "status": "paid", "last_modified_time": "t2"}]),
            self._detail_response({"invoice_id": "inv-1", "line_items": [{"item_id": "item-1", "quantity": 7}]}),
        ]
        cache = {("invoice", "inv-1"): {"date": "2026-09-20", "modified": "t1", "lines": [("item-1", 3)]}}

        result = self._sync(session, cache)

        self.assertEqual(result[("invoice", "inv-1")]["lines"], [("item-1", 7)])
        self.assertEqual(result[("invoice", "inv-1")]["modified"], "t2")

    def test_voided_or_deleted_document_is_dropped_from_cache(self):
        session = MagicMock()
        session.get.side_effect = [self._list_response([])]  # now draft/void/gone -- not listed at all
        cache = {("invoice", "inv-1"): {"date": "2026-09-20", "modified": "t1", "lines": [("item-1", 3)]}}

        result = self._sync(session, cache)

        self.assertNotIn(("invoice", "inv-1"), result)

    def test_cached_entries_outside_the_window_are_pruned(self):
        session = MagicMock()
        session.get.side_effect = [self._list_response([])]
        cache = {("invoice", "old-inv"): {"date": "2026-01-01", "modified": "t1", "lines": [("item-1", 3)]}}

        result = self._sync(session, cache)

        self.assertNotIn(("invoice", "old-inv"), result)

    def test_other_doc_types_in_cache_are_left_alone(self):
        session = MagicMock()
        session.get.side_effect = [self._list_response([])]
        cache = {("creditnote", "cn-1"): {"date": "2026-09-20", "modified": "t1", "lines": [("item-1", 1)]}}

        result = self._sync(session, cache)

        self.assertIn(("creditnote", "cn-1"), result)


class ExtractBrandTests(unittest.TestCase):
    def test_prefers_top_level_brand_field(self):
        item = {"brand": "Acme", "custom_fields": [{"label": "Brand", "value": "Other"}]}
        self.assertEqual(common.extract_brand(item), "Acme")

    def test_falls_back_to_custom_field(self):
        item = {"custom_fields": [{"label": "Brand Name", "value": "Acme"}]}
        self.assertEqual(common.extract_brand(item), "Acme")

    def test_missing_brand_returns_empty_string(self):
        self.assertEqual(common.extract_brand({}), "")


class IsStockItemTests(unittest.TestCase):
    def test_inventory_type_is_kept(self):
        self.assertTrue(common.is_stock_item({"item_type": "inventory"}))

    def test_non_inventory_types_are_excluded(self):
        for item_type in ["purchases", "sales", "sales_and_purchases", "service"]:
            with self.subTest(item_type=item_type):
                self.assertFalse(common.is_stock_item({"item_type": item_type}))

    def test_case_insensitive(self):
        self.assertTrue(common.is_stock_item({"item_type": "Inventory"}))

    def test_missing_item_type_fails_open(self):
        self.assertTrue(common.is_stock_item({}))
        self.assertTrue(common.is_stock_item({"item_type": ""}))


class BuildItemCatalogTests(unittest.TestCase):
    def test_excludes_non_stock_items_from_catalog(self):
        session = MagicMock()

        items_page = MagicMock()
        items_page.raise_for_status = MagicMock()
        items_page.json.return_value = {
            "items": [{"item_id": "item-1"}, {"item_id": "item-2"}],
            "page_context": {"has_more_page": False},
        }

        stock_item_detail = MagicMock()
        stock_item_detail.raise_for_status = MagicMock()
        stock_item_detail.json.return_value = {
            "item": {"item_id": "item-1", "sku": "SKU1", "name": "Widget", "item_type": "inventory"}
        }
        container_detail = MagicMock()
        container_detail.raise_for_status = MagicMock()
        container_detail.json.return_value = {
            "item": {"item_id": "item-2", "sku": "N/A", "name": "20FT Container", "item_type": "purchases"}
        }

        def get_side_effect(url, params=None, headers=None, timeout=None):
            if url.endswith("/inventory/v1/items"):
                return items_page
            if url.endswith("/inventory/v1/items/item-1"):
                return stock_item_detail
            if url.endswith("/inventory/v1/items/item-2"):
                return container_detail
            raise AssertionError(f"Unexpected GET {url}")

        session.get.side_effect = get_side_effect

        token_store = MagicMock()
        token_store.get.return_value = "tok"
        rate_limiter = MagicMock()

        catalog = common.build_item_catalog(session, "https://api", "org1", token_store, rate_limiter)

        self.assertEqual(list(catalog.keys()), ["item-1"])
        self.assertEqual(catalog["item-1"]["Item Name"], "Widget")


class BuildRowsTests(unittest.TestCase):
    def test_excludes_zero_sales_items_and_sorts_remaining_by_name(self):
        catalog = {
            "item-1": {"SKU": "SKU1", "Item Name": "Zebra Widget", "Brand": "Acme"},
            "item-2": {"SKU": "SKU2", "Item Name": "Apple Widget", "Brand": "Acme"},
            "item-3": {"SKU": "SKU3", "Item Name": "No Sales Widget", "Brand": "Acme"},
        }
        buckets = [(datetime.date(2026, 1, 1), datetime.date(2026, 1, 7))]
        sales = {"item-1": [5], "item-2": [0]}

        rows = common.build_rows(catalog, sales, buckets)

        # item-2 (all-zero sales) and item-3 (no sales at all) are both omitted.
        self.assertEqual([row[1] for row in rows], ["Zebra Widget"])
        self.assertEqual(rows[0][3], 5)

    def test_item_with_any_nonzero_week_is_kept(self):
        catalog = {"item-1": {"SKU": "SKU1", "Item Name": "Widget", "Brand": "Acme"}}
        buckets = [
            (datetime.date(2026, 1, 1), datetime.date(2026, 1, 7)),
            (datetime.date(2026, 1, 8), datetime.date(2026, 1, 14)),
        ]
        sales = {"item-1": [0, 3]}

        rows = common.build_rows(catalog, sales, buckets)

        self.assertEqual(len(rows), 1)

    def test_sales_for_items_outside_catalog_are_dropped_with_warning(self):
        catalog = {"item-1": {"SKU": "SKU1", "Item Name": "Widget", "Brand": "Acme"}}
        buckets = [(datetime.date(2026, 1, 1), datetime.date(2026, 1, 7))]
        sales = {"item-1": [5], "discontinued-item": [10]}

        rows = common.build_rows(catalog, sales, buckets)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][1], "Widget")


class WriteExcelTests(unittest.TestCase):
    def test_writes_header_frozen_panes_and_number_format(self):
        import openpyxl
        from unittest.mock import patch

        rows = [
            ["SKU1", "Widget A", "Acme", 3, 0],
            ["SKU2", "Widget B", "Acme", 0, 7],
        ]
        week_columns = ["WK 1", "WK 2"]

        with patch("os.makedirs"):
            output_path = os.path.join(
                os.path.dirname(__file__), "_tmp_test_output.xlsx"
            )
            common.write_excel(rows, week_columns, output_path)

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


class FetchDetailUnauthorizedTests(unittest.TestCase):
    def test_401_forces_a_token_refresh_then_succeeds(self):
        unauthorized = MagicMock(status_code=401)
        ok = MagicMock(status_code=200)
        ok.json.return_value = {"invoice": {"invoice_id": "1"}}
        session = MagicMock()
        session.get.side_effect = [unauthorized, ok]
        token_store = MagicMock()
        token_store.get.return_value = "tok"
        limiter = MagicMock()

        result = common.fetch_detail(session, "https://api", "org", token_store, limiter,
                                     url="https://api/invoices/1", result_key="invoice", label="invoice 1")

        self.assertEqual(result, {"invoice_id": "1"})
        token_store.force_refresh.assert_called_once()


class ResolveTodayTests(unittest.TestCase):
    def test_uses_override_when_set(self):
        from unittest.mock import patch
        from zoneinfo import ZoneInfo

        with patch.dict(os.environ, {"REPORT_AS_OF_DATE": "2026-09-13"}, clear=False):
            today = common.resolve_today(ZoneInfo("UTC"))
        self.assertEqual(today, datetime.date(2026, 9, 13))

    def test_uses_real_now_when_unset(self):
        from unittest.mock import patch
        from zoneinfo import ZoneInfo

        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop("REPORT_AS_OF_DATE", None)
            today = common.resolve_today(ZoneInfo("UTC"))
        self.assertEqual(today, datetime.datetime.now(ZoneInfo("UTC")).date())


if __name__ == "__main__":
    unittest.main()
