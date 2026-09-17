"""
Item Runway Report (weekly)

Replicates Zoho Inventory's "Sales by Item" report over a rolling 36-week
window and writes it to an Excel file named Sales_by_item(WKx-WKy).xlsx,
where x and y are the ISO week numbers of the oldest and most recent weeks
in the window (e.g. Sales_by_item(WK1-WK36).xlsx).

Run every Sunday by .github/workflows/item-runway-weekly-report.yml. Shares
its core logic (auth, rate limiting, catalog fetch, invoice fetch,
aggregation, filtering, Excel formatting) with item_runway_report.py via
sales_by_item_common.py -- see that module's docstring for the full picture
(why there's no direct report API, what gets filtered and why, and which
assumptions still need verifying against the real Zoho UI report).

This script's own responsibility is just: a 36-period window anchored on
"today" itself rather than the last full week before it -- run on a Sunday,
that means the week ending on that same Sunday is included as the most
recent of the 36 (Zoho's own week hasn't fully closed until end of day, but
by the time this runs the data is expected to be final). Run on any other
day (e.g. a manual workflow_dispatch), it anchors on the most recent Sunday
on or before that day instead.

Required/optional environment variables: see sales_by_item_common.py.
"""
import os
from zoneinfo import ZoneInfo

from sales_by_item_common import (
    DEFAULT_RATE_LIMIT_PER_MINUTE,
    compute_week_buckets_ending_by,
    resolve_today,
    run_report,
)

NUM_PERIODS = 36


def main():
    report_tz = ZoneInfo(os.environ.get("REPORT_TZ", "Africa/Nairobi"))
    today = resolve_today(report_tz)

    # Anchored on today itself (not "today - 1 day" like the monthly
    # report) so that a run on Sunday includes the week ending that day.
    buckets = compute_week_buckets_ending_by(today, NUM_PERIODS)

    oldest_week = buckets[0][0].isocalendar()[1]
    newest_week = buckets[-1][0].isocalendar()[1]
    output_path = f"reports/Sales_by_item(WK{oldest_week}-WK{newest_week}).xlsx"

    rate_limit_per_minute = int(os.environ.get("RATE_LIMIT_PER_MINUTE", DEFAULT_RATE_LIMIT_PER_MINUTE))
    run_report(rate_limit_per_minute, buckets, output_path)


if __name__ == "__main__":
    main()
