import os
import shutil
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))

import item_runway_report as report  # noqa: E402
import sales_by_item_common as common  # noqa: E402

OUTPUT_PATH = "reports/Sales_by_item(sep).xlsx"
CACHE_PATH = "data/item_runway_cache_monthly.csv"
AS_OF_ENV = {
    "ZOHO_CLIENT_ID": "id",
    "ZOHO_CLIENT_SECRET": "secret",
    "ZOHO_REFRESH_TOKEN": "refresh",
    "ZOHO_ORGANIZATION_ID": "org1",
    "REPORT_TZ": "UTC",
    "RATE_LIMIT_PER_MINUTE": "6000",
    # Thursday -- that week's Monday is 2026-09-28, so "previous week" (the
    # most recent full week strictly before today's own week) is
    # 2026-09-21 .. 2026-09-27, entirely in September.
    "REPORT_AS_OF_DATE": "2026-10-01",
}


class _BackupRealFiles:
    """Backs up OUTPUT_PATH and CACHE_PATH (the real production paths) for
    the duration of a test and restores them afterward, rather than
    deleting them outright -- an earlier version of this test deleted the
    real committed report when run locally."""

    PATHS = [OUTPUT_PATH, CACHE_PATH]

    def __enter__(self):
        self._backups = {}
        for path in self.PATHS:
            if os.path.exists(path):
                backup = path + ".bak"
                shutil.move(path, backup)
                self._backups[path] = backup
        return self

    def __exit__(self, *exc_info):
        for path in self.PATHS:
            if os.path.exists(path):
                os.remove(path)
            if path in self._backups:
                shutil.move(self._backups[path], path)


def _items_page(*item_ids):
    response = MagicMock()
    response.raise_for_status = MagicMock()
    response.json.return_value = {
        "items": [{"item_id": i} for i in item_ids],
        "page_context": {"has_more_page": False},
    }
    return response


def _item_detail(item_id, sku, name, item_type, brand=None):
    response = MagicMock(status_code=200)
    response.raise_for_status = MagicMock()
    item = {"item_id": item_id, "sku": sku, "name": name, "item_type": item_type}
    if brand is not None:
        item["brand"] = brand
    response.json.return_value = {"item": item}
    return response


def _documents_page(key, documents):
    response = MagicMock()
    response.raise_for_status = MagicMock()
    response.json.return_value = {key: documents, "page_context": {"has_more_page": False}}
    return response


def _invoice_detail(invoice_id, date, line_items):
    response = MagicMock(status_code=200)
    response.raise_for_status = MagicMock()
    response.json.return_value = {"invoice": {"invoice_id": invoice_id, "date": date, "line_items": line_items}}
    return response


class MainIntegrationTest(unittest.TestCase):
    """End-to-end wiring check with every HTTP call mocked.

    Verifies this script's own specific behavior (14-period window anchored
    on the last full week *before* "today", month-name filename, its own
    cache file); the shared fetch/cache/aggregate/filter/format logic
    itself is covered in tests/test_sales_by_item_common.py.
    """

    @patch.dict(os.environ, AS_OF_ENV, clear=True)
    @patch("sales_by_item_common.requests.Session")
    @patch("sales_by_item_common.requests.post")
    def test_main_writes_expected_file_and_cache(self, mock_post, mock_session_cls):
        mock_post.return_value = MagicMock(
            raise_for_status=MagicMock(),
            json=MagicMock(return_value={"access_token": "tok", "api_domain": "https://api"}),
        )

        session = MagicMock()
        mock_session_cls.return_value = session

        items_page = _items_page("item-1", "item-2", "item-3")
        item_detail = _item_detail("item-1", "SKU1", "Widget", "inventory", brand="Acme")
        # Non-stock catalog entry (e.g. a freight/container line item) -- should
        # never appear in the output regardless of whether it has sales.
        container_detail = _item_detail("item-2", "N/A", "20FT Container", "purchases")
        # Real stock item with zero sales in the window -- should be omitted.
        no_sales_detail = _item_detail("item-3", "SKU3", "Unsold Widget", "inventory")

        invoices_page = _documents_page(
            "invoices", [{"invoice_id": "inv-1", "date": "2026-09-22", "status": "paid", "last_modified_time": "t1"}]
        )
        invoice_detail = _invoice_detail("inv-1", "2026-09-22", [{"item_id": "item-1", "quantity": 4}])
        credit_notes_page = _documents_page("creditnotes", [])  # none in this scenario

        def get_side_effect(url, params=None, headers=None, timeout=None):
            if url.endswith("/inventory/v1/items"):
                return items_page
            if url.endswith("/inventory/v1/items/item-1"):
                return item_detail
            if url.endswith("/inventory/v1/items/item-2"):
                return container_detail
            if url.endswith("/inventory/v1/items/item-3"):
                return no_sales_detail
            if url.endswith("/inventory/v1/invoices"):
                return invoices_page
            if url.endswith("/inventory/v1/invoices/inv-1"):
                return invoice_detail
            if url.endswith("/inventory/v1/creditnotes"):
                return credit_notes_page
            raise AssertionError(f"Unexpected GET {url}")

        session.get.side_effect = get_side_effect

        with _BackupRealFiles():
            report.main()
            self.assertTrue(os.path.exists(OUTPUT_PATH))
            self.assertTrue(os.path.exists(CACHE_PATH))

            import openpyxl

            workbook = openpyxl.load_workbook(OUTPUT_PATH)
            sheet = workbook.active
            self.assertEqual([c.value for c in sheet[1]][:3], ["SKU", "Item Name", "Brand"])
            self.assertEqual(len(list(sheet[1])), 3 + 14)  # 14-period window
            # Only item-1 survives: item-2 is a non-stock ("purchases") catalog
            # entry and item-3 has zero sales in the window.
            self.assertEqual(sheet.max_row, 2)
            self.assertEqual(sheet.cell(row=2, column=1).value, "SKU1")
            self.assertEqual(sheet.cell(row=2, column=2).value, "Widget")
            self.assertEqual(sheet.cell(row=2, column=3).value, "Acme")
            self.assertEqual(sheet.cell(row=2, column=sheet.max_column).value, 4)

            cache = common.load_doc_cache(CACHE_PATH)
            self.assertEqual(cache[("invoice", "inv-1")]["lines"], [("item-1", 4)])

    @patch.dict(os.environ, AS_OF_ENV, clear=True)
    @patch("sales_by_item_common.requests.Session")
    @patch("sales_by_item_common.requests.post")
    def test_second_run_skips_detail_fetch_for_unchanged_invoice(self, mock_post, mock_session_cls):
        """The whole point of the cache: an unchanged invoice is listed
        again (cheap) but its line-item detail is not re-fetched."""
        mock_post.return_value = MagicMock(
            raise_for_status=MagicMock(),
            json=MagicMock(return_value={"access_token": "tok", "api_domain": "https://api"}),
        )

        session = MagicMock()
        mock_session_cls.return_value = session

        items_page = _items_page("item-1")
        item_detail = _item_detail("item-1", "SKU1", "Widget", "inventory", brand="Acme")
        invoices_page = _documents_page(
            "invoices", [{"invoice_id": "inv-1", "date": "2026-09-22", "status": "paid", "last_modified_time": "t1"}]
        )
        invoice_detail = _invoice_detail("inv-1", "2026-09-22", [{"item_id": "item-1", "quantity": 4}])
        credit_notes_page = _documents_page("creditnotes", [])

        detail_fetch_count = {"n": 0}

        def get_side_effect(url, params=None, headers=None, timeout=None):
            if url.endswith("/inventory/v1/items"):
                return items_page
            if url.endswith("/inventory/v1/items/item-1"):
                return item_detail
            if url.endswith("/inventory/v1/invoices"):
                return invoices_page
            if url.endswith("/inventory/v1/invoices/inv-1"):
                detail_fetch_count["n"] += 1
                return invoice_detail
            if url.endswith("/inventory/v1/creditnotes"):
                return credit_notes_page
            raise AssertionError(f"Unexpected GET {url}")

        session.get.side_effect = get_side_effect

        with _BackupRealFiles():
            report.main()
            self.assertEqual(detail_fetch_count["n"], 1)

            report.main()  # same mocked listing, same last_modified_time
            self.assertEqual(detail_fetch_count["n"], 1, "invoice detail should not be re-fetched when unchanged")


if __name__ == "__main__":
    unittest.main()
