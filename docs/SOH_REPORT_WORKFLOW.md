# Daily SOH Report — End-to-End Workflow

This documents how the Daily Stock on Hand (SOH) report for Store 1 goes
from Zoho Inventory to its two delivery points (this chat and Zoho Cliq)
with zero manual effort.

## Overview

```
 Trigger                 Generate                Commit            Deliver
┌─────────────┐        ┌──────────────┐        ┌─────────┐       ┌─────────────────┐
│ Claude Code │ ─────▶ │ Fetch active │ ─────▶ │ Commit  │ ───┬─▶│ Zoho Cliq        │
│ Remote      │        │ items from   │        │ to repo │    │  │ (#Sales &        │
│ routine     │        │ Zoho         │        │ (main)  │    │  │  Marketing UG)   │
│ (03:10 UTC  │        │ Inventory,   │        └─────────┘    │  └─────────────────┘
│  daily)     │        │ sort by SKU, │                       │
└─────────────┘        │ write .xlsx  │                       └─▶┌─────────────────┐
                        └──────────────┘                         │ This Claude Code │
┌─────────────┐             ▲                                    │ chat (checked    │
│ GitHub's own│ ────────────┘                                    │ daily)           │
│ `schedule:` │   (rare backup — historically fires              └─────────────────┘
│ cron (03:00 │    hours late, see "Known quirks" below)
│  UTC)       │
└─────────────┘
```

## 1. Trigger

Two triggers exist for the same GitHub Actions workflow
(`.github/workflows/soh-daily-report.yml`), but only one is actually relied on:

- **Primary: a Claude Code Remote routine** ("Deliver daily SOH report
  (direct trigger)"), which fires at **03:10 UTC** every day. It checks
  whether today's report already exists on `main`; if not, it immediately
  calls the GitHub Actions API (`workflow_dispatch`) to start the workflow,
  then self-schedules a ~35 minute follow-up to verify and deliver the
  result.
- **Backup: GitHub's native `schedule:` cron** (`0 3 * * *`, i.e. 03:00
  UTC), defined in the workflow file itself. This is left in place as a
  rare safety net but is **not** relied on for timing — see "Known
  quirks" below.

Either trigger runs the exact same job.

## 2. Generate (`scripts/soh_daily_report.py`)

1. Exchanges the stored Zoho refresh token for a short-lived access token.
2. Fetches every **active** item from Zoho Inventory: a paginated list
   call, then full detail in **batches of up to 100 items per call** via
   Zoho's bulk `/itemdetails` endpoint (the list endpoint doesn't expose
   per-warehouse stock, but fetching one-item-at-a-time cost ~1,800 API
   calls and ~30 minutes; batching cut that to ~18 calls and ~24 seconds).
   Batches fetch concurrently, rate-limited, with retry/backoff on 429s
   and network errors; any batch that fails outright falls back to
   fetching its items one at a time so a single bad id doesn't cost the
   other 99.
3. Filters to items stocked at warehouse **"Store 1"**.
4. Builds one row per item: `BAR CODE, SKU, Item Name, SOH, Sales Price,
   Brand`.
5. **Sorts rows alphabetically by SKU** (case-insensitive).
6. Writes a formatted `.xlsx` (bold frozen header, sized columns, number
   formatting) to `reports/SOH_<DD_MM_YYYY>.xlsx`, dated for the day the
   report is **fetched** (i.e. today, not the prior business day).

## 3. Commit to repo

The workflow commits the new/updated file straight into `reports/` on
`main` (`github-actions[bot]` as the author). It records whether the file
was **newly added** that run (vs. just refreshed with the same date) —
this flag gates the next step.

## 4. Post to Zoho Cliq

**Only runs if the report was newly added this run** (see "Known quirks"
below for why this guard exists). `scripts/post_to_cliq.py`:

1. Refreshes a Cliq-scoped access token (`ZohoCliq.Webhooks.CREATE`,
   using the same Zoho Self Client as Inventory, but a separate refresh
   token: `ZOHO_CLIQ_REFRESH_TOKEN`).
2. Uploads the `.xlsx` via Cliq's file-share API to the channel
   `#Sales & Marketing UG` (unique name `salesmarketingug`, org ID
   `941419808`).

## 5. Deliver to this chat

A separate Claude Code Remote routine checks in daily:

1. Pulls `main`, looks for a newly-added `reports/SOH_*.xlsx` file not
   already delivered this session.
2. If found: opens it with `openpyxl` to sanity-check row count and
   header, then sends it directly in this chat.
3. If the 03:10 UTC trigger hasn't produced a report within ~2 hours
   (rare, but can happen — e.g. Zoho API slowness), it dispatches the
   workflow itself as a catch-up and re-checks ~35 minutes later.

## Required secrets (GitHub → Settings → Secrets and variables → Actions)

| Secret | Purpose |
|---|---|
| `ZOHO_CLIENT_ID` / `ZOHO_CLIENT_SECRET` | Shared Zoho Self Client (Inventory + Cliq) |
| `ZOHO_REFRESH_TOKEN` | Zoho Inventory access, scope `ZohoInventory.FullAccess.all` |
| `ZOHO_ORGANIZATION_ID` | Only needed if the Zoho account has >1 organization |
| `ZOHO_CLIQ_REFRESH_TOKEN` | Zoho Cliq access, scope `ZohoCliq.Webhooks.CREATE` |

## Known quirks / design decisions

- **GitHub's native `schedule:` cron has been unreliable** — observed
  firing anywhere from on-time to 9+ hours late, sometimes not at all
  within a day. This is why a Claude Code Remote routine became the
  *real* trigger instead. When the native schedule does eventually fire
  (refreshing the same day's already-committed file), the commit step
  still runs, but the Cliq step is **skipped** — this prevents the same
  report from being posted to Cliq twice in one day (a bug that was
  caught and fixed after it happened once in production).
- **File naming uses the fetch date**, not "yesterday" — this was changed
  partway through the project's life; older reports in git history follow
  the old "dated for the previous day" convention.
- The Cliq upload step does **not** send a `title` parameter — Cliq's API
  rejects it (`extra_param_found`) on this endpoint variant.
