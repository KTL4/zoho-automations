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

For posting reports into Zoho Cliq, one additional secret is needed:

- `ZOHO_CLIQ_REFRESH_TOKEN` — a refresh token generated from the *same* Self
  Client above, but with the `ZohoCliq.Webhooks.CREATE` scope instead (Self
  Clients can generate multiple refresh tokens, each tied to the scope it
  was requested with).

## SOH Daily Report

`scripts/soh_daily_report.py`, run daily at 06:00 East Africa Time by
`.github/workflows/soh-daily-report.yml`.

Produces `reports/SOH_DD_MM_YYYY.xlsx` (dated for the day the report is
fetched) with columns: BAR CODE, SKU, Item Name,
SOH, Sales Price, Brand — one row per active item stocked at "Store 1",
with a bold header row, frozen header, sized columns, and number
formatting. The workflow commits the file straight into the `reports/`
folder on `main` (requires `permissions: contents: write`, already set in
the workflow) — no manual download from the Actions UI needed; anyone with
repo access has it via a normal `git pull`, and it stays in version
history for good.

After committing, the workflow also posts the file into the Zoho Cliq
channel `#Sales & Marketing UG` (unique name `salesmarketingug`, org ID
`941419808`) via `scripts/post_to_cliq.py`, using the Cliq REST API's
"Share files in a Channel" endpoint
(`POST https://cliq.zoho.com/company/{org}/api/v2/channelsbyname/{channel}/files`,
scope `ZohoCliq.Webhooks.CREATE`). If the Cliq channel or org ever changes,
override `CLIQ_CHANNEL_UNIQUE_NAME` / `CLIQ_ORG_ID` env vars in the
workflow rather than editing the script's defaults.

**Note:** Zoho Inventory's API has no direct endpoint for the "Stock
Summary" report shown in the web UI, so this is built from the Items API's
per-warehouse stock figures instead. This should match the UI report in
standard setups, but hasn't been verified against the exact "Bills &
Invoices" stock-tracking mode — spot-check the first run's numbers against
the Zoho Inventory UI.

Because the per-warehouse stock breakdown is only on each item's *detail*
endpoint (not the list endpoint), the script fetches detail for every
active item — but in batches of up to 100 via Zoho's bulk `/itemdetails`
endpoint rather than one call per item (switched 2026-10-27; a live run
against ~1,795 active items now takes **18 API calls and ~24 seconds**,
down from ~1,800 calls and ~30 minutes). Batches fetch concurrently, paced
by a shared rate limiter (`RATE_LIMIT_PER_MINUTE`, default 60/min), log
progress after every batch, and log a heartbeat every 20 seconds so a
stalled run is visible in real time in the workflow's logs rather than
going silent. Network errors and rate-limit (429) responses are retried
with backoff (honoring the `Retry-After` header when Zoho sends one); if
a whole batch still fails after retries, its items are re-fetched one at a
time via the single-item endpoint instead (so one bad id doesn't cost the
other 99), and an item that still fails after that is logged and skipped
rather than crashing the whole report. The job keeps its 45-minute
timeout as a safety cap even though a normal run now finishes in well
under a minute.

**Root cause found on 2026-10-28:** the bulk `/itemdetails` endpoint doesn't
return a top-level `brand` field the way the single-item endpoint (and the
items list endpoint) do, so Brand came back blank for every row for one
day after the batching switch above. Since the items list call (already
made to collect active item ids) returns `brand` for free, `scripts/soh_daily_report.py`
now also builds an id-to-brand map from that call and patches it back
onto each item after the detail fetch — no extra API calls needed.

An alternative optimization — only re-fetching items whose
`last_modified_time` changed since the last run — was considered and
rejected: there's no confirmation that a stock movement (a sale, a
purchase receipt, a manual adjustment) actually updates an item's
`last_modified_time` in Zoho's data model, so that approach risked
silently serving stale stock figures for exactly the data this report
exists to report correctly. The bulk-batch approach gets the same API
call reduction without that correctness risk, since every item is still
fetched fresh every run.

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

## Item Runway Reports (Sales by Item)

Two schedules of the same underlying report, both replicating Zoho
Inventory's Reports > Sales by Item:

- **Monthly** — `scripts/item_runway_report.py`, run at 06:00 East Africa
  Time on the 1st of each month by `.github/workflows/item-runway-report.yml`.
  Date range "Previous week" compared against **14** previous periods.
  Produces `reports/Sales_by_item(<mon>).xlsx` (e.g. `Sales_by_item(sep).xlsx`,
  named for the month the most recent of the 14 weeks falls in).
- **Weekly** — `scripts/item_runway_weekly_report.py`, run at 06:00 East
  Africa Time every Sunday by `.github/workflows/item-runway-weekly-report.yml`.
  Rolling **36**-period window anchored on the week ending that Sunday
  itself (not the last full week *before* it, unlike the monthly report —
  see "Window anchoring" below). Produces
  `reports/Sales_by_item(WKx-WKy).xlsx`, e.g. `Sales_by_item(WK2-WK37).xlsx`,
  where `x`/`y` are the ISO week numbers of the oldest/newest weeks in the
  window.

Both share their entire data-fetching, filtering, and formatting logic via
`scripts/sales_by_item_common.py` — only the window length, window
anchoring, and output filename differ between the two thin entry-point
scripts. This is deliberate: a fix to, say, the rate limiter now benefits
both automations instead of needing to be applied twice and kept in sync
by hand.

Each output file has columns SKU, Item Name, Brand, then one column per
week (`WK <n>`, oldest to newest, labeled with that week's ISO
week-of-year number). One row per stock-tracked item that sold at least
once in the window — non-stock catalog entries (freight/container line
items etc., see below) and items with zero sales across the whole window
are both omitted. Same bold header / frozen header / sized columns /
number formatting as the SOH report, and both workflows commit straight
into `reports/` on `main` the same way (each using its own filename glob,
so a monthly run never touches a weekly-pattern file or vice versa).

**Assumptions baked into these reports** (documented in
`sales_by_item_common.py`'s docstring too) — verify a run's numbers
against the real Zoho UI report before trusting the automation, per the
project's testing convention:
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
2. Every non-draft, non-void invoice dated in the window is listed (cheap
   — paginated, sorted newest-first, and pagination stops as soon as it
   walks past the window) and then fetched in full, one call per invoice,
   to get its line items.
3. Each line item's quantity is summed into the week its invoice date
   falls in.
4. Items with zero sales across the whole window are dropped from the
   final output.

Both the catalog fetch and the invoice-detail fetch reuse the exact same
rate limiter, heartbeat logging, and retry-with-backoff pattern that fixed
the SOH report's production 429 crash (see above) — invoice volume over
the window is not bounded the way the item catalog is, so this can end up
making significantly more API calls than the SOH report does. Both jobs
have a generous 180-minute timeout to accommodate that.

**Window anchoring:** the monthly report anchors on "yesterday" so its
window is the last *full* week strictly before the day it runs (matching
Zoho's own "Previous week" semantics when viewed any day of the month).
The weekly report anchors on "today" itself, so a run on Sunday includes
the week ending that same Sunday as its most recent period — the
assumption being that by the time the job actually runs, that day's sales
data is final. Both are implemented via one shared function,
`compute_week_buckets_ending_by(reference_date, num_periods)`, which
returns the `num_periods` most recently completed Monday-Sunday weeks on
or before `reference_date`; only the reference date passed to it differs
(`today - 1 day` for monthly, `today` for weekly).

**Backfilling/testing a specific date:** both scripts read an optional
`REPORT_AS_OF_DATE` environment variable (`YYYY-MM-DD`) that overrides
"today" — useful for seeing what a run would have produced on a past date
without waiting for the schedule. Both workflows expose this as a
`workflow_dispatch` input (`as_of_date`) for manual runs; leave it blank
for a normal run using the real current date.

**Live run history:**
- **2026-09-11** (monthly script, still at 35 periods before it was
  narrowed to 14): completed in ~1h42m against a catalog of 1,842 active
  items — same order of magnitude as the SOH report's ~1,850 items, so the
  catalog-fetch phase takes about as long (~30 min) as it does there; the
  rest was the invoice-detail phase.
- **2026-09-14** (monthly script, narrowed to 14 periods, with the
  non-stock/zero-sales filters added): completed in ~1h04m, confirming
  both the shorter window and the filters worked as intended (789 rows,
  down from 1,842; no more all-zero rows; the "20FT Container" freight
  item no longer appears).

## Container Sellout Report

`scripts/container_sellout_report.py`, run at 06:15 East Africa Time every
Monday by `.github/workflows/container-sellout-report.yml` (also runnable
on demand from the Actions tab). Full walkthrough:
`docs/CONTAINER_SELLOUT_WORKFLOW.md`.

Produces `reports/Container_Sellout_DD_MM_YYYY.xlsx` with three tabs:

- **Containers** — one row per container that has ever arrived (one bill
  from a container supplier): arrival date, units in, units sold, other
  outflows, units on hand, **Sellout % (Zoho FIFO)**, **Sold % (Invoices)**,
  sold % by day 30/60/90/180, days to 50% and 80% sold, last sale, cost in
  and cost still on hand.
- **Container x SKU** — the same figures per SKU within each container.
- **Read Me** — the definitions, and whether the sales history is complete.

Zoho doesn't record which container a sale came from, so the report uses
FIFO (oldest stock sells first), the same way Zoho values stock:

1. **Lots.** Zoho's own FIFO Cost Lot Tracking report
   (`reports/fifocostlottracking`) gives every inward lot per SKU, with
   quantity in and quantity still on hand. Container lots are bills /
   purchase receives from the suppliers in `CONTAINER_SUPPLIER_PATTERNS`
   (UBL, Aiwibi -- billed as YIFAR or "AWB Inventory Receipt" -- and EFRIS
   STOCK by default), grouped by bill number. This
   alone gives the headline Sellout %, but it counts every outflow
   (transfers and write-offs as well as sales).
2. **Sales.** Every non-draft, non-void invoice since the first lot, at line
   level. Credit notes are not subtracted: Zoho's FIFO already books each
   return as a new inward lot, so netting them off too would count every
   return twice. This costs one API call per invoice, so the
   lines are cached in `data/sales_lines_<org>.csv` and committed. Each run
   only fetches invoices that are new or edited since the last one
   (`last_modified_time`), and drops voided ones.
3. **Timing.** Each SKU's daily sales are matched to its lots oldest-first,
   never more than Zoho says left that lot and never before it arrived.
   What Zoho consumed that invoices don't explain is "Other Outflows".

The raw lot pull is saved alongside, as `data/fifo_lots_<org>.csv`, so the
data behind any report can be checked.

**Saving API calls:** the full lot history is ~1,000 calls for KTL, and
almost all of it is old, sold-out stock that never changes. So a normal run
re-pulls only the last 365 days of lots (`LOTS_REFRESH_DAYS`), keeps older
lots from `data/fifo_lots_<org>.csv`, and re-lists only the last 90 days of
invoices for edits/voids (`SALES_RECHECK_DAYS`). Every 28 days
(`FULL_REFRESH_DAYS`), or when the `full_refresh` box is ticked on a manual
run, it re-pulls everything to catch rare late changes to old lots. The date
of the last full refresh is in `data/sellout_state_<org>.json`. A normal run
is roughly a quarter of the calls of a full one.

**Live run history:**
- **2026-10-29** (first run, KTL org 780891784): the lot pull worked and
  found containers C1 (Dec 2019) onward as "UBL Inventory Receipt" bills,
  plus Aiwibi/YIFAR bills, but it hit the then 500-page cap at mid-2024
  (most report rows are outward movements, ~28 inward lots per page) and
  then crashed with a 401: every "report still computing" retry had
  minted a new access token, and Zoho revoked the working one. Fixed by
  raising the cap, minting a token only for a stale first page, and
  refreshing on any 401. It also showed that 8,215 of the lots are credit
  notes, which is why returns are no longer netted off sales.
- **2026-10-30** (run 2): succeeded, PARTIAL (3,548 invoices left for the
  next run). 27,057 lots, 40 containers -- but it missed all 17 Aiwibi
  containers billed to vendor "AWB Inventory Receipt (USD)" (A04 onward,
  LCL W1/W2), because "AWB" wasn't a supplier pattern. Fixed; 57 containers.

**First run (backfill):** fetching every historical invoice can take
longer than one GitHub Actions job allows. The script stops fetching after
`SALES_FETCH_TIME_BUDGET_MINUTES` (300), saves what it has, still writes
the report with a PARTIAL note on the Read Me tab, and the next run carries
on. Trigger it manually a few times in a row until the PARTIAL note goes.
Sellout % (Zoho FIFO) is complete from the very first run.

**Verify before trusting it:** the field names of the FIFO lot report are
taken from the inventory-workflows toolkit's `pull_fifo_lots.ps1`, not yet
from a run in this repo. On the first run, check the log's lot-type
summary, and tie one container's Units In to its bill and its Units On
Hand to the stock in Zoho.
