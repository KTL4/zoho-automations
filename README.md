# zoho-automations

Automations that pull data from Zoho Inventory. Each automation lives in
`scripts/` and runs on a schedule via GitHub Actions (`.github/workflows/`).

## Setup

These repository secrets must be set (Settings → Secrets and variables →
Actions) before any workflow will run:

- `ZOHO_CLIENT_ID`
- `ZOHO_CLIENT_SECRET`
- `ZOHO_REFRESH_TOKEN`
- `ZOHO_ORGANIZATION_ID` — only required if the Zoho account has more than
  one organization; scripts auto-detect it otherwise.

These come from a Zoho API Console "Self Client" (https://api-console.zoho.com/)
with the `ZohoInventory.FullAccess.all` scope. The refresh token does not
expire unless revoked, but if any script starts failing with an
authentication error, generate a fresh one from the same Self Client.

## SOH Daily Report

`scripts/soh_daily_report.py`, run daily at 06:00 East Africa Time by
`.github/workflows/soh-daily-report.yml`.

Produces `reports/SOH_DD_MM_YYYY.xlsx` (dated for the previous day, since
the job runs after Store 1 closes) with columns: BAR CODE, SKU, Item Name,
SOH, Sales Price, Brand — one row per active item stocked at "Store 1",
with a bold header row, frozen header, sized columns, and number
formatting. The workflow commits the file straight into the `reports/`
folder on `main` (requires `permissions: contents: write`, already set in
the workflow) — no manual download from the Actions UI needed; anyone with
repo access has it via a normal `git pull`, and it stays in version
history for good.

**Note:** Zoho Inventory's API has no direct endpoint for the "Stock
Summary" report shown in the web UI, so this is built from the Items API's
per-warehouse stock figures instead. This should match the UI report in
standard setups, but hasn't been verified against the exact "Bills &
Invoices" stock-tracking mode — spot-check the first run's numbers against
the Zoho Inventory UI.

Because the per-warehouse stock breakdown is only on each item's *detail*
endpoint (not the list endpoint), the script makes one API call per active
item — expect the run time to scale with catalog size (the live Store 1
catalog has ~1,850 active items). It fetches details concurrently, paced
by a shared rate limiter (`RATE_LIMIT_PER_MINUTE`, default 60/min), logs
progress every 100 items, and logs a heartbeat every 20 seconds so a
stalled run is visible in real time in the workflow's logs rather than
going silent. Network errors and rate-limit (429) responses are retried
with backoff (honoring the `Retry-After` header when Zoho sends one); an
item that still fails after retries is logged and skipped rather than
crashing the whole report. The job has a 45-minute timeout as a safety cap
(1,850 items at 60/min is ~31 minutes before any retries, so there's
headroom).

**Root cause found on 2026-09-03:** the scheduled run crashed with
`requests.exceptions.HTTPError: 429` after being throttled by Zoho for
close to 30 minutes straight — the account's real request rate limit is
apparently well below what firing 6 concurrent, unpaced requests produces.
The earlier version only backed off for a few seconds per item, which
isn't enough to recover from *sustained* throttling. The rate limiter
above paces every request (across all worker threads) to a fixed budget
instead of letting the workers fire as fast as they can and hoping
short backoffs are enough. If 429s still show up frequently in the logs
after this change, lower `RATE_LIMIT_PER_MINUTE` further.

## Item Runway Report

`scripts/item_runway_report.py`, run at 06:00 East Africa Time on the 1st
of each month by `.github/workflows/item-runway-report.yml`. Replicates
Zoho Inventory's Reports > Sales by Item, with date range "Previous week"
compared against 14 previous periods.

Produces `reports/Sales_by_item(<mon>).xlsx` (e.g. `Sales_by_item(sep).xlsx`,
named for the month the most recent of the 14 weeks falls in) with columns
SKU, Item Name, Brand, then one column per week (`WK <n>`, oldest to
newest, labeled with that week's ISO week-of-year number). One row per
stock-tracked item that sold at least once in the window — non-stock
catalog entries (freight/container line items etc., see below) and items
with zero sales across all 14 weeks are both omitted. Same bold header /
frozen header / sized columns / number formatting as the SOH report, and
the workflow commits straight into `reports/` on `main` the same way.

**Assumptions baked into this report** (documented in the script's
docstring too) — verify the first run's numbers against the real Zoho UI
report before trusting the automation, per the project's testing
convention:
- "Sales by Item" quantity = **invoiced** quantity (draft/void invoices
  excluded), not booked sales-order quantity.
- Covers **all warehouses/locations combined**, not a single store.
- Weeks run **Monday-Sunday**.
- The item list is scoped to **active items** (`Status.Active`), matching
  the SOH report's convention, further restricted to `item_type ==
  "inventory"` to drop non-sellable/non-stock catalog entries (e.g. a
  "20FT Container" freight line item showed up in the first live run,
  which is what prompted this filter). An item missing the `item_type`
  field entirely is kept rather than dropped, to fail open — this is a
  best-effort heuristic and hasn't been verified against Zoho's exact
  field semantics, so double-check it isn't silently dropping real
  products.

**Note:** like the SOH report, Zoho Inventory's API has no direct "Sales
by Item" report endpoint — a Zoho community thread confirms this gap
generally ("this functionality exists in the Reports interface, but there
is no API available for it"). So this is built from transactional data
instead:

1. The full active item catalog is fetched (list + per-item detail, same
   approach as the SOH report) to get SKU/Item Name/Brand for every item,
   filtering out non-stock items as described above.
2. Every non-draft, non-void invoice dated in the 14-week window is listed
   (cheap — paginated, sorted newest-first, and pagination stops as soon
   as it walks past the window) and then fetched in full, one call per
   invoice, to get its line items.
3. Each line item's quantity is summed into the week its invoice date
   falls in.
4. Items with zero sales across the whole window are dropped from the
   final output.

Both the catalog fetch and the invoice-detail fetch reuse the exact same
rate limiter, heartbeat logging, and retry-with-backoff pattern that fixed
the SOH report's production 429 crash (see above) — invoice volume over
the window is not bounded the way the item catalog is, so this can end up
making significantly more API calls than the SOH report does. The job has
a generous 180-minute timeout to accommodate that.

**First live run (2026-09-11, against a 35-period window before it was
narrowed to 14):** completed successfully in ~1h42m against a catalog of
1,842 active items — same order of magnitude as the SOH report's ~1,850
items, so the catalog-fetch phase takes about as long (~30 min) as it does
there; the rest of the time was the invoice-detail phase. With the window
now 14 weeks instead of 35 (~40% of the date range), expect a materially
shorter invoice-fetch phase on the next run, though the exact volume still
depends on real invoice counts.
