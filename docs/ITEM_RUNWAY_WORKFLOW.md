# Item Runway Reports — End-to-End Workflow

> **Retired 4 Nov 2026.** Both reports are paused: the schedules are commented out, and the weekly workflow is also
> disabled in GitHub Actions. Their Claude Code delivery tasks are deleted.
>
> | Was used for | Now comes from |
> |---|---|
> | Weekly units file `Sales_by_item(WKx-WKy).xlsx` for the Stock Health run | `scripts/weekly_units_report.py`, part of the Saturday Sales Ledger run. Same file name and layout. |
> | Sales data for item listing and ranking | `Sales_13.xlsx` and `ITEM_LISTING_FEED.xlsx` |
>
> See `docs/SALES_LEDGER_WORKFLOW.md`. Manual runs still work from the Actions tab if ever needed.
> One known fault if they are revived: the credit-note listing used a login token that had already expired, so credit
> notes were never subtracted.

Two schedules of the same underlying report ("Sales by Item"), both built
on the same incremental invoice/credit-note cache.

## Overview

```
 Trigger            Generate                              Commit        Deliver
┌────────────┐    ┌───────────────────────────────────┐  ┌─────────┐  ┌──────────────┐
│ Monthly:   │ ─▶ │ 1. Item catalog (full fetch)       │─▶│ Commit  │─▶│ reports/     │
│ 1st, 06:00 │    │ 2. Invoices + credit notes in the  │  │ report +│  │ Sales_by_    │
│ EAT        │    │    window: list in full (cheap),   │  │ cache to│  │ item(...)    │
│            │    │    fetch detail only for new/      │  │ main    │  │ .xlsx on main│
│ Weekly:    │    │    changed documents               │  └─────────┘  └──────────────┘
│ Sunday,    │    │ 3. Net quantity per item per week  │                       │
│ 06:00 EAT  │    │    (invoices minus credit notes)   │                       ▼
└────────────┘    └───────────────────────────────────┘             ┌──────────────────┐
                                                                      │ Claude Code      │
                                                                      │ routine, Sunday  │
                                                                      │ 16:00 EAT (weekly│
                                                                      │ report only)     │
                                                                      └──────────────────┘
```

## 1. Trigger

- **Monthly** — `.github/workflows/item-runway-report.yml`, cron `0 3 1 * *`
  (06:00 EAT on the 1st). Window: 14 periods, anchored on the last *full*
  week strictly before the run day (Zoho's own "Previous week" semantics).
- **Weekly** — `.github/workflows/item-runway-weekly-report.yml`, cron
  `0 3 * * 0` (06:00 EAT every Sunday). Window: 36 periods, anchored on
  "today" itself, so the week ending that Sunday is included.
- Both: manual `workflow_dispatch`, with an optional `as_of_date` input to
  backfill/preview what a run would have produced on a past date.
- Both: a `concurrency` group per workflow, since each run reads and
  rewrites its own cache file — two runs of the *same* workflow must never
  overlap. The monthly and weekly workflows use separate cache files (see
  below), so they're free to run close together (they have, e.g. both
  triggered within 15 minutes of each other on 2026-11-01) without
  conflicting.

## 2. Generate (`scripts/sales_by_item_common.py`, shared by both entry scripts)

1. **Item catalog.** Every active item, filtered to `item_type ==
   "inventory"` (drops non-sellable entries like freight/container line
   items, e.g. "20FT Container"). Fetched fresh every run — not cached,
   since it's small (~1,800 items) and fixed-cost regardless of window
   size, unlike invoices/credit notes below.
2. **Invoices and credit notes, incrementally cached.** For each document
   type: list every non-draft/non-void document dated in the window
   (cheap, unthrottled, paginated) and diff against the cache file:
   - A document not in the cache, or whose `last_modified_time` no longer
     matches the cache, gets its line-item detail fetched (the expensive,
     rate-limited call).
   - A document still matching the cache is skipped — no detail fetch.
   - A cached document no longer listed (voided/deleted) or that has aged
     out of the window is dropped from the cache.
   This is the same `last_modified_time`-diffing pattern already proven by
   `container_sellout_report.py`'s invoice sync — reused here rather than
   reinvented. It is **not** the item-level `last_modified_time` idea the
   SOH report's README explicitly considered and rejected (see there):
   that was about inferring a stock-quantity change from an *item's* own
   timestamp, unconfirmed to actually update for that; this is a
   *document's* own timestamp changing when the document itself is edited
   or voided, which Container Sellout already relies on successfully.
3. **Net quantity per item per week** = invoiced quantity minus credited
   quantity, bucketed by each document's own date (a credit note nets off
   the week *it* was issued, not the week of the original invoice).
4. Items with net-zero quantity across the whole window are dropped from
   the final output.
5. Write a formatted `.xlsx` (bold frozen header, sized columns, number
   formatting): SKU, Item Name, Brand, then one column per week (`WK <n>`,
   oldest to newest, labeled by ISO week number).

### Credit notes: inferred endpoint, defensive fallback

Zoho Inventory's credit note field/endpoint names (`creditnotes`,
`creditnote_id`) are inferred from this API's naming convention for its
sibling endpoints (`salesorders`, `purchaseorders`), not confirmed against
a live response from this account. If the credit note list call fails
(wrong endpoint, unexpected shape, etc.), the run logs a warning and
proceeds on invoices alone rather than crashing — so a wrong guess here
degrades to "credit notes aren't netted off yet" instead of breaking the
report. **Verify this on the first live run after this feature ships**,
the same way every other assumption in this project gets verified.

Net-of-credit-notes is the intended definition for *this* report's
"quantity sold". That's a different question from whether Zoho's own FIFO
stock accounting nets them off — it doesn't (see
`container_sellout_report.py`'s docstring): there, a credit note is booked
as a *new inward lot*, so subtracting it from sales *there* too would
double-count the return. Both can be true at once: physically, returned
stock re-enters the warehouse (Container Sellout's concern); commercially,
a "Sales by Item" report nets the return off what was sold (this report's
concern).

## 3. Commit

Each workflow commits its own report file and its own cache file
(`data/item_runway_cache_monthly.csv` / `data/item_runway_cache_weekly.csv`)
to `main`. The commit step runs even if the report step failed
(`if: always()`), so a run that dies partway through a long sync never
loses the invoices/credit notes it already fetched — the next run picks up
from the saved cache instead of starting over.

## 4. Deliver

- **Monthly report:** committed to `reports/` on `main`; no separate
  delivery step (download from the repo).
- **Weekly report:** in addition to the commit above, a Claude Code
  routine fires every Sunday at **16:00 EAT**, checks that morning's
  GitHub Actions run succeeded, downloads the resulting
  `reports/Sales_by_item(WKx-WKy).xlsx` from `main`, and sends it directly
  in chat. If the run failed or didn't happen, it investigates and reports
  that instead of sending a stale file.

## Settings (environment variables, both scripts)

| Variable | Default | Purpose |
|---|---|---|
| `ZOHO_ORGANIZATION_ID` | auto-detected | Required only if the account has more than one organization |
| `REPORT_TZ` | `Africa/Nairobi` | Timezone used to determine "today" |
| `REPORT_AS_OF_DATE` | unset | Overrides "today" (`YYYY-MM-DD`) for backfill/testing; exposed as the `as_of_date` workflow input |
| `RATE_LIMIT_PER_MINUTE` | 60 | Zoho API calls per minute for detail fetches (catalog + document sync share one limiter) |

## Known quirks / design decisions

- **The item catalog is fetched independently by both reports, every
  run** — about 1,800 duplicate calls on any day both happen to run (e.g.
  the 1st of the month, if it's also a Sunday). Not yet optimized; a
  candidate for the same `last_modified_time` caching treatment if call
  volume ever becomes a real constraint, but out of scope for the
  invoice/credit-note change that introduced this cache (2026-11-02).
- **Monthly and weekly reports keep separate cache files on purpose**,
  even though their windows overlap heavily (14 weeks is a subset of 36).
  This avoids the two independently-scheduled workflows ever writing to
  the same file in the same run; the cost is some duplicated cached
  invoice/credit-note lines between the two files, which is cheap compared
  to the API calls it saves.
- **First run after this feature ships starts with an empty cache**, so it
  costs the same as every run did before (one API call per document in the
  window). Savings show up from the second run onward.
- **Credit note field/endpoint names are inferred, not confirmed live**
  (see above) — this is the one piece of this change that most needs
  checking against a real run's logs and the Zoho UI.

## Live run history

- **2026-09-11 / 2026-09-14** (pre-caching): see the git history for the
  original call-count investigation that motivated this change — a 36-week
  weekly run cost ~6,490 detail calls (~1h42m); a 14-week monthly run cost
  ~4,031 (~1h04m). Both scale with invoice volume in the window, which is
  exactly what incremental caching avoids paying for on every run.
