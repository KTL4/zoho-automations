import os
import sys
import unittest
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))

import weekly_units_report as W  # noqa: E402
import sales_by_item_common as C  # noqa: E402


def row(day, sku, qty, total, doc_type="invoice", name="Item", brand="UBL"):
    return {"date": day, "sku": sku, "quantity": str(qty), "total": str(total), "doc_type": doc_type,
            "item_name": name, "brand": brand}


class WeeklyUnitsTests(unittest.TestCase):
    def setUp(self):
        # two Monday-Sunday weeks: 26 Oct - 1 Nov and 2 - 8 Nov 2026
        self.buckets = C.compute_week_buckets_ending_by(date(2026, 11, 8), 2)

    def test_invoiced_units_by_week_without_credit_notes_or_charges(self):
        rows = [row("2026-10-26", "A1", 5, 500), row("2026-11-01", "A1", 2, 200),
                row("2026-11-07", "A1", 3, 300),
                row("2026-11-03", "A1", -1, -100, doc_type="creditnote"),
                row("2026-10-20", "A1", 9, 900),                       # before the window
                row("2026-11-02", "SHIPPING CHARGES", 1, 5000),
                row("2026-11-02", "COMBO", 4, 400, name="Ledger name", brand="AIWIBI")]
        out = {r[0]: r for r in W.build_rows(rows, self.buckets, {"A1": ("SOH name", "UBL")})}
        self.assertEqual(out["A1"], ["A1", "SOH name", "UBL", 7, 3])
        self.assertEqual(out["COMBO"], ["COMBO", "Ledger name", "AIWIBI", 0, 4])
        self.assertNotIn("SHIPPING CHARGES", out)

    def test_saturday_run_includes_the_week_ending_next_day(self):
        sat = date(2026, 11, 7)
        end = sat + (date(2026, 11, 8) - sat)
        buckets = C.compute_week_buckets_ending_by(sat + __import__("datetime").timedelta(days=6 - sat.weekday()), 36)
        self.assertEqual(buckets[-1], (date(2026, 11, 2), end))
        self.assertEqual(C.week_label(buckets[-1][0]), "WK 45")


if __name__ == "__main__":
    unittest.main()
