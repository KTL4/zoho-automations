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
