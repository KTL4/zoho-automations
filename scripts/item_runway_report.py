"""
Item Runway Report (monthly)

Replicates Zoho Inventory's "Sales by Item" report (Reports > Sales by Item,
date range "Previous week", compared with "Previous periods", 14 periods)
and writes it to an Excel file named Sales_by_item(<mon>).xlsx.

Run on the 1st of each month by .github/workflows/item-runway-report.yml.
Shares its core logic (auth, rate limiting, catalog fetch, invoice fetch,
aggregation, filtering, Excel formatting) with item_runway_weekly_report.py
via sales_by_item_common.py -- see that module's docstring for the full
picture (why there's no direct report API, what gets filtered and why, and
which assumptions still need verifying against the real Zoho UI report).

This script's own responsibility is just: a 14-period window anchored on
the last full week before "today" (i.e. the same "Previous week" Zoho's UI
would show if run today), and a Sales_by_item(<mon>).xlsx filename named
for the month that window's most recent week falls in.

Required/optional environment variables: see sales_by_item_common.py.
"""
import os
from datetime import timedelta
from zoneinfo import ZoneInfo

from sales_by_item_common import (
    DEFAULT_RATE_LIMIT_PER_MINUTE,
    compute_week_buckets_ending_by,
    resolve_today,
    run_report,
)

NUM_PERIODS = 14


def main():
    report_tz = ZoneInfo(os.environ.get("REPORT_TZ", "Africa/Nairobi"))
    today = resolve_today(report_tz)

    # "Previous week" relative to today, i.e. the last full week strictly
    # before today's own week -- so anchor on yesterday, not today itself.
    buckets = compute_week_buckets_ending_by(today - timedelta(days=1), NUM_PERIODS)

    output_month = buckets[-1][1].strftime("%b").lower()
    output_path = f"reports/Sales_by_item({output_month}).xlsx"

    rate_limit_per_minute = int(os.environ.get("RATE_LIMIT_PER_MINUTE", DEFAULT_RATE_LIMIT_PER_MINUTE))
    run_report(rate_limit_per_minute, buckets, output_path)


if __name__ == "__main__":
    main()
