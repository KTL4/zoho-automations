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

Produces `SOH_DD_MM_YYYY.xlsx` (dated for the previous day, since the job
runs after Store 1 closes) with columns: BAR CODE, SKU, Item Name, SOH,
Sales Price, Brand — one row per active item stocked at "Store 1". The file
is attached to the workflow run as a downloadable artifact (retained 90
days); it is not committed to the repo or emailed anywhere.

**Note:** Zoho Inventory's API has no direct endpoint for the "Stock
Summary" report shown in the web UI, so this is built from the Items API's
per-warehouse stock figures instead. This should match the UI report in
standard setups, but hasn't been verified against the exact "Bills &
Invoices" stock-tracking mode — spot-check the first run's numbers against
the Zoho Inventory UI.

Because the per-warehouse stock breakdown is only on each item's *detail*
endpoint (not the list endpoint), the script makes one API call per active
item — expect the run time to scale with catalog size (the live Store 1
catalog has ~1,850 active items). It fetches details concurrently, logs
progress every 100 items, and logs a heartbeat every 20 seconds so a
stalled run is visible in real time in the workflow's logs rather than
going silent. Network errors and rate-limit (429) responses are retried
with backoff; an item that still fails after retries is logged and
skipped rather than crashing the whole report. The job has a 45-minute
timeout as a safety cap.

A run on 2026-09-02 stalled completely after ~100/1850 items with no
errors logged, which the previous version of the script had no visibility
into (no heartbeat, no per-item error logging) — it just went silent until
the job timeout killed it. If this recurs, the heartbeat and per-item error
logs added since should show whether it's sustained Zoho rate-limiting
(steady 429 messages) or something else, which will tell us whether to
lower `DETAIL_FETCH_WORKERS`, add a rate limiter, or look at another cause.
