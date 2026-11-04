# Sales Ledger — Weekly Update Workflow

Keeps `reports/Sales_13.xlsx` up to date. The file has one row per invoice or
credit note line, for every sale since December 2019.

## What it does

| | |
|---|---|
| **When** | Every Saturday at 23:00 EAT (`.github/workflows/sales-ledger-weekly.yml`), or on demand from the Actions tab |
| **Output** | `reports/Sales_13.xlsx`, overwritten in place, so the repo holds one file with the latest data. Earlier versions stay in git history. |
| **Master data** | `data/sales_ledger_780891784/<year>.csv` (one file per year) and `data/sales_ledger_state_780891784.json` |
| **Script** | `scripts/sales_ledger.py` |
| **Cost** | ~400 Zoho API calls in a normal week. Scheduled runs are capped at 1,000. |

Columns: Invoice ID, Date, Sales Person Name, Customer Name, SKU, Item Name,
Brand, Category, Total, Quantity.

- **Total** includes VAT and is after line discounts (for a percentage discount, rate × quantity × (1 − discount %)). This matches the old export and Zoho's own Item Total + Item Tax on every line of the 21 Sep – 1 Nov export.
- **Credit notes** are their own rows on the credit note's date, with negative Total and Quantity.

## Each run

1. **List.** Lists every invoice in Zoho (all history, ~90 calls) and every
   credit note dated from 21 Sep 2026 onward.
2. **Compare.** Checks each listed document's `last_modified_time` against the ledger:
   - **New or edited** (price changed, line removed, quantity corrected): fetches the
     document's details (1 call each) and replaces all of its rows.
   - **Voided or deleted** (no longer listed): removes its rows.
   - **Unchanged:** does nothing.
3. **Save.** Rewrites the yearly CSVs and `reports/Sales_13.xlsx`, then commits
   both to `main`.

**Safety checks:**
- If Zoho suddenly lists fewer than 90% of the invoices the ledger holds, the run stops and changes nothing.
- If the call cap is hit, the rest are fetched next run. The log says `PARTIAL`.
- Each fetched invoice's lines are added up and checked against Zoho's
  invoice total. Any that don't match are listed in the log under `Check:`.

## Category

Category is chosen in this order:
1. `data/sales_category_overrides.csv`, for SKUs that have no category in Zoho.
   Add a row here for any new SKU that comes out blank.
2. The SKU's most recent category in the existing data.
3. The item's category in Zoho.

## History that can't be re-checked

The ledger was seeded on 2 Nov 2026 by `scripts/seed_sales_ledger.py` from:
- the old `Sales_13.xlsx` export (up to 20 Sep 2026)
- Zoho's invoice-line CSV export (21 Sep – 1 Nov 2026)

Excel had rounded the old invoice IDs. 15,833 old invoices were matched
back to their Zoho IDs by date, SKU and quantity, so edits to them are
picked up. The remaining rows stay exactly as exported (`doc_type = history`)
and are never re-checked:
- 15,047 invoice rows that couldn't be matched to a single invoice
- 13,353 old credit-note rows (the old export gave them no ID)

The 1,196 invoice IDs those unmatched rows could belong to are listed in the
state file under `frozen_invoice_ids`, so they are never fetched twice.

## Check after the first live run

- **Credit notes listed:** the log's "Listed N creditnotes" line is non-zero.
  The credit-note endpoint has never worked in this repo before. The Item
  Runway report's 401 on it was an expired login token, not a wrong endpoint.
- **No `Check:` warnings:** they would mean the line-total formula needs adjusting.
- **Spot-check:** compare one new invoice and one credit note against Zoho.

## Live run history

- **2026-11-04, run 1** (manual, 348 calls): listed 18,039 invoices and 140 credit notes and fetched 245 documents.
  - **Result:** October came to 691.9m, VAT incl., net of 128.4m in credit notes. That matches the team's own figure.
  - **Problem:** 40 discounted invoices came out too high. Line totals subtracted Zoho's `discount_amount`, but on a VAT-inclusive invoice that amount is sometimes the ex-VAT share.
  - **Fix:** a percentage discount is now applied to the VAT-inclusive rate. The 245 documents are marked to be fetched again on the next run.
