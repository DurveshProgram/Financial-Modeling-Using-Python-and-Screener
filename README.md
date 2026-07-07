# Financial Model Engine — Screener.in → Full 3-Statement Model + DCF

## What this does
Feed it ANY company's Screener.in export (Export to Excel button on any
screener.in company page) and it generates a complete, formula-linked,
professional financial model: Cover, Assumptions/WACC, Income Statement,
Balance Sheet, Cash Flow, Ratios, DCF (FCFF), Relative Valuation (Comps),
and a Valuation Dashboard — matching the structure of the TCS model you
already built.

## How to run it
```
cd financial_model_engine
python3 generate_model.py path/to/CompanyScreener.xlsx Output_Model.xlsx
python3 /mnt/skills/public/xlsx/scripts/recalc.py Output_Model.xlsx    # (if on this environment)
```
Or on your own machine: open `Output_Model.xlsx` in Excel — it will
recalculate automatically.

## How it works
- `screener_parser.py` — reads the "Data Sheet" tab that every Screener.in
  export contains (company meta, P&L, Balance Sheet, Cash Flow, derived
  share count) by locating section headers, so it works for any company,
  not just TCS.
- `build_model.py` — pours that data into `model_template.xlsx` (your
  original TCS model, now genericized): historical actuals become
  hardcoded blue inputs; every ratio, margin, DCF and comps formula stays
  a live Excel formula, untouched.
- `generate_model.py` — the CLI entry point; also holds the macro/WACC
  assumptions (risk-free rate, equity risk premium, terminal growth) that
  Claude sources from current published data each time this is refreshed.

## What's fetched from the web vs. defaulted to last-actual-year
- **Fetched (macro, genuinely market-wide):** risk-free rate (India 10-Yr
  G-Sec yield), equity risk premium (India country-risk-adjusted ERP).
  These aren't scraped live by the script itself — Claude looks them up
  and hardcodes them into `generate_model.py`'s MACRO dict each time you
  ask for a refresh, with the source noted in a cell comment.
- **Defaulted to last actual year (flat-lined):** every operating
  projection assumption — revenue growth, EBITDA margin, capex %, DSO,
  tax rate, etc. These are company-specific and not reliably fetchable
  for an arbitrary ticker, so the honest default is "hold at the most
  recent actual year," fully visible and editable in the Assumptions tab.
- **Flagged as "confirm manually":** beta and peer-comps multiples.
  A reliable levered beta and a fresh peer set really need a market-data
  terminal (Bloomberg/Refinitiv) or NSE/BSE data — I did not want to
  silently guess these. Cells are commented telling you exactly what to
  check and where.

## Known limitations (next things to harden)
- Assumes 10 historical annual years, matching Screener.in's current
  export format. If a company's Screener export has a different count,
  the script will warn you — the row/column structure will still need a
  quick manual check.
- Comps sheet peer rows are still the illustrative Infosys/HCLTech/Wipro
  entries — swap in the right peer set per sector before trusting the
  peer-multiple valuation.
- No live web-scraping *inside* the Python script (by design — scraping
  NSE/BSE/Screener for real-time data reliably needs an API key/ToS
  review). The macro refresh is a one-line dict Claude updates for you
  on request.
