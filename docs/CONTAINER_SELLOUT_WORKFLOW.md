# Container Sellout Report — End-to-End Workflow

How much of every container that has ever arrived has sold, and how fast.

## Overview

```
 Trigger              Phase 1: Lots          Phase 2: Sales         Phase 3: Match       Output
┌────────────┐      ┌────────────────┐     ┌──────────────────┐   ┌───────────────┐   ┌──────────────────────┐
│ Saturday   │ ───▶ │ Zoho FIFO Cost │ ──▶ │ Invoices, line   │──▶│ FIFO: each    │──▶│ reports/Container_   │
│ 16:00 UTC  │      │ Lot Tracking   │     │ level            │   │ SKU's sales → │   │ Sellout_DD_MM_YYYY   │
│ or manual  │      │ (every lot ever│     │ (cached in data/,│   │ oldest lot    │   │ .xlsx                │
│ run        │      │  received)     │     │  only new/edited │   │ first         │   │ + data/*.csv         │
└────────────┘      └────────────────┘     │  fetched)        │   └───────────────┘   └──────────────────────┘
                                           └──────────────────┘
```

## 1. Trigger

`.github/workflows/container-sellout-report.yml`:
- Scheduled every **Saturday at 16:00 UTC** (19:00 EAT). Until 2026-11-02 it ran Mondays at 06:15 EAT.
- GitHub can start scheduled runs late (the SOH report has seen delays of hours). So a **backup** Claude Code
  scheduled task fires at **19:10 EAT every Saturday**. If that Saturday's report isn't on `main` and no run is
  in progress, it starts the workflow itself with `max_api_calls=700`. A scheduled run that starts late skips
  itself when that day's report already exists, so the two never both run.
- **Call limit:** weekly runs (scheduled or started by the backup) stop before going over **700** Zoho API calls (500 until 2026-11-09; the 7 Nov run used 492).
  The monthly full refresh (~1,100 calls) **always runs without a limit**: the script lifts it as soon as it
  sees a full refresh is due.
- Manual: GitHub → Actions → *Container Sellout Report* → **Run workflow**.
  The optional `as_of_date` input (YYYY-MM-DD) runs it as if it were that date.

Runs never overlap (`concurrency` group), because each run rewrites the sales cache.

## 2. Generate (`scripts/container_sellout_report.py`)

1. **Lots.** Pulls Zoho Inventory's FIFO Cost Lot Tracking report for all
   history (`LOTS_FROM_DATE`, default 2015-01-01). Every inward lot is
   kept, because opening stock and adjustments compete with containers
   under FIFO. A lot counts as a **container lot** when it is a bill or
   purchase receive from a contact matching `CONTAINER_SUPPLIER_PATTERNS`
   (default: Uncle Bill / UBL, Aiwibi / Aiwina / Yifar / AWB, EFRIS STOCK).
   Container = bill number. The log prints the 15 most common lot types
   and suppliers, so you can see what was or wasn't counted.
2. **Sales.** Lists every non-draft, non-void invoice since the first
   lot, then fetches line detail only for documents that
   are new or whose `last_modified_time` changed. The lines live in
   `data/sales_lines_<org>.csv`.
3. **Match.** For each SKU, net daily sales go to the oldest lot that had
   arrived by that date and still had room. A lot's room is what Zoho
   says left it (units in − on hand).
4. **Write** `reports/Container_Sellout_DD_MM_YYYY.xlsx`, dated for the run day.

### Normal runs vs full refreshes

| | Normal run | Full refresh |
|---|---|---|
| When | Every run, unless a full one is due | Every 28 days, or tick `full_refresh` on a manual run |
| Lots re-pulled | Last 180 days (~3,500 lots, ~150 calls for KTL) | All history (~27,000 lots, ~1,000 calls, ~1 hour) |
| Older lots | Kept from `data/fifo_lots_<org>.csv` | Re-pulled |
| Invoices re-listed for edits/voids | Last 90 days | All history |
| Invoice details fetched | Only new or edited invoices | Only new or edited invoices |

FIFO uses up the oldest stock first, so older lots are mostly sold out. On
2026-11-09, 1,080 lots older than 180 days still had stock (51,560 units);
their on-hand figures can lag by up to 28 days on a normal run, until the
next full refresh. (With the old 365-day window it was 546 lots, 34,922
units; the owner chose 180 days to save ~95 calls a week.) The state file `data/sellout_state_<org>.json` records the last
full refresh and whether the first invoice backfill has finished.

## 3. Commit

The workflow commits the report and `data/*.csv` to `main`. The commit
step runs **even if the report step failed**, so a long backfill never
loses the invoices it already fetched.

## 4. Use the report

Download the latest file from `reports/` and save it to
`C:\Users\user\Downloads\Claude outputs\` for upload to WorkDrive.

| Column | Meaning |
|---|---|
| Sellout % (Zoho FIFO) | (Units In − Units On Hand) ÷ Units In. Straight from Zoho. Includes transfers and write-offs |
| Sold % (Invoices) | Units invoiced that FIFO places in this container ÷ Units In. Returns aren't subtracted: Zoho books a credit note as a new inward lot |
| Other Outflows | What left the container that invoices don't explain: transfers, adjustments, damages |
| Sold % by Day N | Share invoiced within N days of arrival. Blank if the container is younger than N days |
| Days to 50% / 80% Sold | Days from arrival until that share was invoiced. Blank if not reached yet |
| Cost In / Cost On Hand | Base currency, at FIFO lot cost |

## Settings (workflow `env:`)

| Variable | Default | Purpose |
|---|---|---|
| `ZOHO_ORGANIZATION_ID` | secret | Which org (KTL / KSL). One org per run |
| `CONTAINER_SUPPLIER_PATTERNS` | UBL, Aiwibi, EFRIS STOCK names | Which contacts count as container suppliers |
| `LOTS_FROM_DATE` | 2015-01-01 | Earliest lot to pull |
| `SALES_FETCH_TIME_BUDGET_MINUTES` | 300 | Stop fetching invoices after this long and save progress |
| `LOTS_REFRESH_DAYS` | 180 (365 until 2026-11-09) | Lots re-pulled on a normal run |
| `SALES_RECHECK_DAYS` | 90 | Invoices re-listed on a normal run |
| `FULL_REFRESH_DAYS` | 28 | How often a full refresh runs |
| `FULL_REFRESH` | workflow input | `true` forces a full refresh |
| `MAX_API_CALLS` | 700 on scheduled runs; workflow input `max_api_calls` on manual runs, blank = none. Never applied to a full refresh | Hard cap on Zoho API requests (retries count). On reaching it the run stops, keeps the invoices fetched so far, writes no report, and fails with "STOPPED: ... limit". Every run logs "Zoho API calls this run: N" |
| `RATE_LIMIT_PER_MINUTE` | 60 | Zoho API calls per minute |

## Known quirks / design decisions

- **The lot pull alone takes about an hour for KTL.** Most rows in the
  FIFO report are outward movements, so each 200-row page holds only ~28
  inward lots; the first live run needed 500 pages to reach mid-2024.
- **Never mint a new token per retry.** Zoho revokes the oldest live access
  token once too many exist for one refresh token; the first live run
  (2026-10-29) crashed with a 401 that way. Tokens are refreshed only on a
  401 or a stale first page of the lot report.
- **Edited invoices with the same date and total aren't re-downloaded** on a normal run. On 2026-11-07, 91 of
  232 invoice downloads were older invoices whose only visible change was `last_modified_time` (most likely a
  payment being recorded). The list call already returns each invoice's total, so this check costs nothing. A
  full refresh re-downloads every edited invoice, which catches the rare edit that changes items but keeps the
  exact same total.
- **Credit notes are inward lots**, not negative sales (8,215 of KTL's
  lots are credit notes), so they are not subtracted from sales.

- **First run is a backfill.** It needs one API call per historical
  invoice. At 60 calls a minute, 300 minutes is about 18,000 invoices per
  run. Until it finishes, the Read Me tab says PARTIAL. Re-run manually
  until that goes. Sellout % (Zoho FIFO) is complete from run one.
- **Zoho's FIFO report can serve a stale "This Month" result** instead of
  the requested range (found by the inventory-workflows toolkit,
  2026-07-18). A suspiciously short first page triggers a fresh token and
  a retry.
- **Containers raised as POs but not yet received** (C34 onward are POs
  left "to be received / to be billed") aren't in stock in Zoho, so they
  don't appear until they're received or billed.
- **Timing ignores the dates of transfers and adjustments.** They show up
  as Other Outflows but don't move which lot a sale was matched to.
- **Matched by SKU**, because the lot report is keyed on SKU. A SKU renamed
  between containers splits into two.
