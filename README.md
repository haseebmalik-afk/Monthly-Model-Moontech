# MoonTech Monthly Model — dashboard

Replaces the Excel/Google Sheets monthly model. Raw monthly extracts go in, `data.json` is rebuilt, and `index.html` shows the five pages (Investor dashboard, Annual summary, Monthly summary, Monthly actuals, Control & data quality).

Rules and definitions follow Kareem's `SPEC_dashboard-data.md` v0.25. The calculation layer is his reference engine (`pipeline/moontech_ref.py`, unchanged) plus `pipeline/build.py` (ROAS S/M/L tiers, Ounass buffer scenarios, page datasets). 43/43 check figures must pass on every build.

## Monthly update

1. Drop the new extracts into `data/` using the same file names and formats as today:
   - `Moontech-_Data_updates_2025_1.xlsx` — Datos revenue by merchant, orders by creator, CAC sheets
   - `MoonTech_Monthly_Model_6.xlsx` — P&L import (2026 Financials) and history snapshots
   - `moontech-roas-ladder-YYYY-MM-DD.csv` — the ROAS ladder (update the file name in `pipeline/moontech_ref.py` → `load_ladder` if it changes)
   - `data_5_creator_orders_*.xlsx` — Datos creator orders export
2. Edit `pipeline/control.json`: `latest_month`, Ounass `provisional_months` / `confirmed_months`, capitalisation per year. (Or export it from the dashboard Control page.)
3. Run `python3 pipeline/build.py` from the repo root (needs `openpyxl`, `pyyaml`; ~2 minutes). It prints the check-figure result and writes `data.json`.
4. Optional: `python3 pipeline/bundle.py` writes `preview.html` (single file with data inlined).
5. Commit `data.json`, `index.html`, `pipeline/control.json`. GitHub Pages serves `index.html` from the repo root.

## Layout

```
index.html            the dashboard (reads data.json)
data.json             built dataset: base + Ounass add-back scenarios
preview.html          single-file preview (built by pipeline/bundle.py)
pipeline/             build.py, moontech_ref.py (Kareem's engine), config.yaml, control.json, merchant master, check figures
data/                 monthly source extracts (not committed while the repo is public)
```

## Controls

- **View**: Segmented (Legacy and ROAS separately) or Company-wide (blended rows, split kept visible).
- **Ounass buffer**: On = system figure held; Off = AED 150k added back to each unconfirmed provisional month. Confirmed months are locked.
- **Development**: one switch per year on the Control page (Capitalised / Expensed / Split).

Confidential — contains client revenue by name.
