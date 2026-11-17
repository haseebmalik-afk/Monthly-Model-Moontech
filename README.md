# MoonTech Monthly Model — dashboard

Monthly exports go in, `data.json` is rebuilt by the pipeline, and `index.html` shows it. The HTML never computes a
revenue number from a raw export: everything it shows comes from `data.json` (built once a month) plus four live
Google Sheets. **If you only have `index.html` and `data.json`, you can change how things look but you cannot update a
month — the `pipeline/` folder is the part that turns exports into data.**

Update date today: see `pipeline/control.json` → `latest_month`. Rules: Kareem's `SPEC_dashboard-data.md` v0.25
(engine `pipeline/moontech_ref.py`) plus the post-handover decisions listed in the doc *Changes post handover*.
43/43 check figures must pass on every build; 2024 and 2025 are frozen and the build fails if they move.

## How inputs become outputs

```
monthly data pack (uploaded)            live Google Sheets (read at run time)
  data/pack/net_sales_by_client.xlsx      ROAS ladder (receivables + tracker)  → also snapshotted to data/pack/ladder.json
  data/pack/orders_by_creator.xlsx        CAC sheet (merchant + creator spend)
  data/pack/ads_by_creator.xlsx           2026 P&L (model tab)                 → also snapshotted to data/pack/2026_model_pnl.csv
  data/pack/revenue_by_merchant.xlsx      CSM ROAS mapping (ROAS factor)
  pipeline/performance_commission.csv (Level shoes, typed by finance)
  pipeline/control.json (update date, Ounass provisional/confirmed, capitalisation, cash)
            │
            ▼
  python3 pipeline/build.py      pack_loaders.py reads the pack → moontech_ref.py (Kareem's engine, §-rules) →
                                 build.py (periods, segments, S/M/L tiers, Ounass scenarios, checks, health)
            │
            ▼
  data.json  (base + Ounass add-back scenarios + meta: checks, health, data-pack card)
            │
            ▼
  index.html (reads data.json; computes only presentation-level views: Investor set, ARPA at period end,
             period picker, trend explorer; fetches the four live sheets, cut at the update date)
  preview.html = index.html with data.json inlined (python3 pipeline/bundle.py) — the file to share as a single page
```

Locked inputs that never change month to month:
- `pipeline/history_base.json` — the baked 2024–25 history (merchants, creators, P&L) and creator history to Aug-26.
- `pipeline/static_2024_25.json` — frozen 2024–25 values; `build.py` refuses to finish if any of them moves.
- `data/pack/orders_by_creator_2024_25.xlsx` — per-creator orders 2024–25 (one-time), so the creator status rules can
  look back 12 months.
- `pipeline/roas_segments.csv` — Enterprise / Growth / SME per ROAS brand (Kareem's segmentation file).
- `pipeline/merchant_master_seed.csv` — commission client master (aliases, segments, ROAS links).
- `pipeline/check_figures.csv` — the 43 check figures (re-baselined, with a note, whenever the update date moves).

## Monthly update (what finance does)

1. Put the month's four BI exports in `data/pack/` (same names as above; each export is the new month only — the pack
   files hold all 2026 months, the new month is appended/replaced).
2. Add the month's Level shoes performance commission to `pipeline/performance_commission.csv`.
3. Make sure the live 2026 P&L sheet has the month; snapshot it to `data/pack/2026_model_pnl.csv` and the ladder to
   `data/pack/ladder.json` + `ladder_milestones.json` (`window.mtLadderExport()` on the ROAS Ladder page returns both).
4. `pipeline/control.json`: `latest_month` = the new month; Ounass `provisional_months` (buffer applies) /
   `confirmed_months`.
5. `python3 pipeline/build.py` (needs `openpyxl`, `pyyaml`; ~2 minutes). It prints the check-figure result and the
   2024–25 lock, and writes `data.json`. Checks that depend on the update month are re-baselined with a dated note.
6. `python3 pipeline/bundle.py` → `preview.html`.
7. Cross-check before accepting: Revenue by Merchant values for every earlier month must equal the previous pack
   (any change is flagged); the Control page checklist must have no red item.
8. Commit `data.json`, `index.html`, `preview.html`, `pipeline/`. GitHub Pages serves `index.html` from the repo root.

## Definitions (one per metric — the page and the pipeline agree)

- Revenue = P&L Sales. Commission revenue per client from the Revenue by Merchant workbook (its own Adjustment row is
  ignored); ROAS revenue from the ladder; **Adjustment = P&L Sales − tracked commission − ladder ROAS**.
- One update date for the whole model: every source, live sheets included, is cut at `latest_month`.
- Commission client: New in the first month it passes the active test (annualised revenue ≥ $950, not 6 quiet months);
  Churned after 6 quiet months; Reactivated if it comes back.
- ROAS brand: New in its ladder Since month; Churned only at the ladder Churn month (no payment-gap rule). New Model =
  ROAS incl. Jetour. Size: margin on paid ÷ CAC — <1× Small, 1–2× Medium, >2× Large. Segments from
  `roas_segments.csv` (unmapped = SME · U.D).
- ARPA: commission side = TTM revenue ÷ active clients at period end; ROAS / New Model = ARR (last 3 months × 4) ÷
  active brands at period end.
- ARR = last 3 months, seasonality-adjusted, × 4 (engine §10). Investor set (`computeInvestor()` in `index.html`)
  follows Kareem's Investor Metrics file; New Model growth = last 3 months vs prior 3.
- Creator: Active = avg ≥ 5 net orders/month over the last 12 months (or since first order) and an order in the last
  6 months; New / Existing / Churned / Reactivated / Emerging from that; Major = annualised ≥ 5,000 (exit < 2,500).
  Whole orders per creator-month for the status rules; raw net orders for totals.
- Ounass buffer: "On · held" = books as they are; "Off · added back" adds AED 150k to Ounass and to P&L Sales for each
  provisional month (revenue, ARR, margins, NRR move; Adjustment does not).
- CAC: live CAC sheet; KPI tiles use the selected period (spend ÷ new brands in the period); Investor = since launch.

## Checking a modified `index.html` against the build

```
python3 pipeline/build.py
python3 pipeline/contract_check.py --mine path/to/your/index.html
```
Prints: data fields the page reads that the build does not produce; which JavaScript functions differ from the
reference and whether they sit on the input → output path; live-sheet links that differ. Presentation changes (CSS,
markup, labels) are separated from changes that alter how a number is read or computed.

## Layout

```
index.html            the dashboard (reads data.json + live sheets)
preview.html          single-file version (data inlined)
data.json             built dataset
pipeline/             build.py, pack_loaders.py, moontech_ref.py (engine), bundle.py, contract_check.py, config.yaml,
                      control.json, history_base.json, static_2024_25.json, check_figures.csv, roas_segments.csv,
                      performance_commission.csv, merchant_master_seed.csv
data/pack/            the monthly data pack (git-ignored while the repo is public)
```

Confidential — contains client revenue by name.
