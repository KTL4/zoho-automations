import os
import shutil
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))

import item_runway_weekly_report as report  # noqa: E402
import sales_by_item_common as common  # noqa: E402

OUTPUT_PATH = "reports/Sales_by_item(WK2-WK37).xlsx"
CACHE_PATH = "data/item_runway_cache_weekly.csv"
AS_OF_ENV = {
    "ZOHO_CLIENT_ID": "id",
    "ZOHO_CLIENT_SECRET": "secret",
    "ZOHO_REFRESH_TOKEN": "refresh",
    "ZOHO_ORGANIZATION_ID": "org1",
    "REPORT_TZ": "UTC",
    "RATE_LIMIT_PER_MINUTE": "6000",
    "REPORT_AS_OF_DATE": "2026-09-13",  # a Sunday (ISO week 37)
}


class _BackupRealFiles:
    """Backs up OUTPUT_PATH and CACHE_PATH (the real production paths) for
    the duration of a test and restores them afterward, rather than
    deleting them outright, in case a live run has already committed them."""

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


def _mock_response(json_value):
    response = MagicMock(status_code=200)
    response.raise_for_status = MagicMock()
    response.json.return_value = json_value
    return response


class MainIntegrationTest(unittest.TestCase):
    """End-to-end wiring check with every HTTP call mocked.

    Verifies this script's own specific behavior (36-period window anchored
    on "today" itself, so a Sunday run includes the week ending that day,
    the WKx-WKy filename pattern, and its own cache file); the shared
    fetch/cache/aggregate/filter/format logic itself is covered in
    tests/test_sales_by_item_common.py.
    """

    @patch.dict(os.environ, AS_OF_ENV, clear=True)
    @patch("sales_by_item_common.requests.Session")
    @patch("sales_by_item_common.requests.post")
    def test_main_writes_expected_file_with_wk_range_filename(self, mock_post, mock_session_cls):
        mock_post.return_value = _mock_response({"access_token": "tok", "api_domain": "https://api"})

        session = MagicMock()
        mock_session_cls.return_value = session

        items_page = _mock_response({
            "items": [{"item_id": "item-1"}],
            "page_context": {"has_more_page": False},
        })
        item_detail = _mock_response({
            "item": {
                "item_id": "item-1",
                "sku": "SKU1",
                "name": "Widget",
                "brand": "Acme",
                "item_type": "inventory",
            }
        })
        invoices_page = _mock_response({
            "invoices": [{"invoice_id": "inv-1", "date": "2026-09-10", "status": "paid", "last_modified_time": "t1"}],
            "page_context": {"has_more_page": False},
        })
        invoice_detail = _mock_response({
            "invoice": {
                "invoice_id": "inv-1",
                "date": "2026-09-10",
                "line_items": [{"item_id": "item-1", "quantity": 7}],
            }
        })
        # A credit note the same week, partially netting off the invoice.
        credit_notes_page = _mock_response({
            "creditnotes": [{"creditnote_id": "cn-1", "date": "2026-09-11", "status": "closed", "last_modified_time": "t1"}],
            "page_context": {"has_more_page": False},
        })
        credit_note_detail = _mock_response({
            "creditnote": {
                "creditnote_id": "cn-1",
                "date": "2026-09-11",
                "line_items": [{"item_id": "item-1", "quantity": 2}],
            }
        })

        def get_side_effect(url, params=None, headers=None, timeout=None):
            if url.endswith("/inventory/v1/items"):
                return items_page
            if url.endswith("/inventory/v1/items/item-1"):
                return item_detail
            if url.endswith("/inventory/v1/invoices"):
                return invoices_page
            if url.endswith("/inventory/v1/invoices/inv-1"):
                return invoice_detail
            if url.endswith("/inventory/v1/creditnotes"):
                return credit_notes_page
            if url.endswith("/inventory/v1/creditnotes/cn-1"):
                return credit_note_detail
            raise AssertionError(f"Unexpected GET {url}")

        session.get.side_effect = get_side_effect

        # Reference date 2026-09-13 (a Sunday, ISO week 37) with a 36-period
        # window ending on that week means the oldest week is ISO week 2.
        with _BackupRealFiles():
            report.main()
            self.assertTrue(os.path.exists(OUTPUT_PATH))
            self.assertTrue(os.path.exists(CACHE_PATH))

            import openpyxl

            workbook = openpyxl.load_workbook(OUTPUT_PATH)
            sheet = workbook.active
            header = [c.value for c in sheet[1]]
            self.assertEqual(header[:3], ["SKU", "Item Name", "Brand"])
            self.assertEqual(len(header), 3 + 36)  # 36-period window
            self.assertEqual(header[3], "WK 2")  # oldest week first
            self.assertEqual(header[-1], "WK 37")  # week ending on the run's own Sunday
            self.assertEqual(sheet.cell(row=2, column=1).value, "SKU1")
            # Net quantity: 7 invoiced - 2 credited = 5, both dated in WK 37.
            self.assertEqual(sheet.cell(row=2, column=sheet.max_column).value, 5)

            cache = common.load_doc_cache(CACHE_PATH)
            self.assertEqual(cache[("invoice", "inv-1")]["lines"], [("item-1", 7)])
            self.assertEqual(cache[("creditnote", "cn-1")]["lines"], [("item-1", 2)])


if __name__ == "__main__":
    unittest.main()
