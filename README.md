# Financial Model Engine — Screener.in → Full 3-Statement Model + DCF

## What this does
Feed it ANY company's Screener.in export (Export to Excel button on any
screener.in company page) and it generates a complete, formula-linked,
professional financial model: Cover, Assumptions/WACC, Income Statement,
Balance Sheet, Cash Flow, Ratios, DCF (FCFF), Relative Valuation (Comps),
and a Valuation Dashboard.

Two ways to run it:
1. **Web interface** (new) — drag-and-drop a file in the browser.
2. **Command line** — `python generate_model.py <file.xlsx> <output.xlsx>`

---

## 1. Web interface

```
cd webapp
pip install -r requirements.txt
python app.py
```
Open **http://127.0.0.1:5000**, drag in a Screener.in export, click
**Generate Model**, download the result. Nothing is uploaded anywhere
outside your own machine — `app.py` runs a local Flask server only.

- Uploaded files are deleted from `uploads/` immediately after processing.
- Generated workbooks land in `outputs/` (also deleted at your discretion —
  they're not auto-cleaned, since you'll usually want to keep them).
- `app.config["MAX_CONTENT_LENGTH"]` caps uploads at 16 MB (Screener exports
  are ~30-70 KB, so this just guards against the wrong file being dropped in).

This is a local dev server (`debug=True`), fine for personal use. If you
ever want to expose it beyond your own machine, put a real WSGI server
(gunicorn/waitress) in front of it and turn debug mode off first.

## 2. Command line
```
python generate_model.py path/to/CompanyScreener.xlsx Output_Model.xlsx
python /mnt/skills/public/xlsx/scripts/recalc.py Output_Model.xlsx    # (if on this environment)
```
Or just open `Output_Model.xlsx` in Excel — it recalculates automatically.

---

## How it works
- `screener_parser.py` — reads the "Data Sheet" tab that every Screener.in
  export contains (company meta, P&L, Balance Sheet, Cash Flow, historical
  share price, derived share count) by locating section headers, so it
  works for any company, not just one you've tested.
- `build_model.py` — pours that data into `model_template.xlsx`: historical
  actuals become hardcoded blue inputs; every ratio, margin, DCF and comps
  formula stays a live Excel formula, untouched.
- `generate_model.py` — CLI entry point; also holds the macro/WACC
  assumptions and orchestrates the market-data lookups.
- `market_data.py` — layered live-data fetcher (see below).
- `app.py` / `templates/` / `static/` — the Flask web interface.

---

## Data sources & sector classification — fixed this round

Three real bugs were found and fixed in this pass (all confirmed against
actual generated output for TCS / Infosys / Asian Paints, not just code
review):

### Bug 1 — sector silently defaulted to IT
Previously, if ticker resolution failed, `classify_sector()` fell back to
`DEFAULT_SECTOR_KEY = "IT"` no matter what the company actually was — so
**Asian Paints was assigned Infosys/HCLTech/Wipro as its peer set.**

**Fix:** `DEFAULT_SECTOR_KEY` is now `None`. Sector classification runs in
this order, and stops as soon as one layer succeeds:
1. **Offline name-keyword match** (`classify_sector_by_name`) — matches the
   Screener.in company name itself against a keyword list per sector.
   Zero network dependency, so it can't fail due to a bad connection, and
   it's what actually fixed Asian Paints (matched on "paint").
2. **NSE India's own sector/industry tags** (`get_nse_quote`) — the listing
   exchange's own classification, fetched live.
3. **Yahoo Finance sector/industry text** (`get_yahoo_profile`) — secondary
   live source.

If none of these resolve, the sector is **left unresolved** (`None`) —
`build_model.py` then leaves the Comps sheet peer rows as an explicit
`[Enter Peer N name — ...]` placeholder with a comment explaining why,
instead of guessing.

### Bug 2 — stale data leaking across companies
When a sector's peer list had fewer than 3 usable names left (e.g. Infosys
excluding itself from IT's 3-peer list, leaving only 2), the 3rd row in the
Comps sheet was never cleared — so it silently kept whatever static
placeholder value was baked into `model_template.xlsx` from the original
reference build (a stale "Wipro Ltd" entry showed up in an *Infosys* model).

**Fix:** `build_model.py` now clears **all 3 peer row slots** (value +
comment) before conditionally filling in however many live peers exist.

### Bug 3 — Yahoo Finance fetch not returning data (and Google Finance was fetched but never actually used)
`yfinance` calls were coming back empty (all peer metrics `0`, cost-of-
capital inputs on `FALLBACK`). This is a known, ongoing issue — Yahoo has
repeatedly tightened anti-bot/crumb requirements on its unauthenticated
endpoints, and it isn't something a client library can reliably work
around. Separately, the first version of this fix added a Google Finance
scraper but its price was fetched and then discarded — it was never
actually written into a peer row or the Comps sheet.

**Fix — three more layers, all of them now genuinely wired into the output:**
- **NSE India's public API** (`nseindia.com/api/*`) — free, no key, the
  actual exchange's own data. Used for sector/industry, current price,
  and 52-week range. Requires a session cookie from a normal page hit
  first (standard, documented behaviour of that API, not a workaround of
  anything private).
- **TradingView's scanner endpoint** (`scanner.tradingview.com`) — no
  official API, but this is the same unauthenticated endpoint TradingView's
  own screener page uses. One call returns price, market cap, P/E,
  EV/EBITDA, P/B, ROE, sector, industry, and 52-week range together — and
  it's **batched across all 3 peers in a single request**, rather than one
  call per peer. This is now the primary source for the Comps sheet.
- **Google Finance** (best-effort HTML scrape of the public quote page) —
  its page doesn't reliably expose fundamentals (P/E etc.) in scrapable
  form, so it's used only as a last-resort **current-price** fallback if
  NSE, TradingView, and Yahoo all come up empty. There's no official
  Google Finance API, so this is inherently fragile (Google can change
  markup without notice).

Fetch order per field, cheapest/most-authoritative first:
`sector/industry`: name-keyword → NSE → TradingView → Yahoo.
`peer P/E, EV/EBITDA, P/B, ROE, market cap`: TradingView (batched) → Yahoo
per-peer for whatever's still missing.
`current price / 52-week range`: NSE → TradingView → Yahoo → Google Finance.
`beta`, `equity risk premium`: Yahoo only (no other source here exposes
either).

Every field in the workbook still carries a cell comment saying exactly
which source(s) answered it (`LIVE: NSE (...)`, `LIVE: TradingView (...)`,
`PARTIAL — TradingView + Yahoo (...)`, or `FALLBACK: ...` with the reason)
— nothing is ever presented as live when it's actually a static default,
and a peer row that's only partially filled says so explicitly rather than
silently looking complete.


### Bug 4 — historical P/E, P/B, EV/EBITDA used the wrong company's stock price
`Ratios!31` (historical market price per share, which every historical
valuation multiple in that sheet is computed from) was never overwritten
by `build_model.py` — so **every non-TCS company silently inherited TCS's
actual 10-year share price history** for its historical P/E, P/B, and
EV/EBITDA ratios.

**Fix:** `screener_parser.py` now extracts the `PRICE:` row from the Data
Sheet (Screener's single-row historical closing-price series), and
`build_model.py` writes it into `Ratios!31`. Verified independently correct
per company (Asian Paints ₹1,073→₹1,667, Infosys ₹511→₹641, TCS ₹1,216→
₹1,826 across the same historical window).

### Also hardened
- `Comps!C10` (peer average) and the two "Implied Value" cells now use
  `IFERROR(...)` so an unresolved/incomplete peer set shows a clear
  **"n/a" / "Enter peer data"** text instead of a `#DIV/0!` error or —
  worse — a silent `0` that looks like a real (implausibly cheap) multiple.
- Peer metrics that couldn't be fetched are now left **blank (`None`)**,
  not coerced to `0` — a blank cell is correctly excluded from `AVERAGE()`;
  a `0` was silently corrupting the peer average.

---

## Sector coverage
`SECTOR_MAP` in `market_data.py` now covers 15 sectors (IT, Bank, NBFC,
Insurance, FMCG, Auto, Pharma, Energy, Metal, Paints/Chemicals, Cement,
Telecom, Infra/Construction, Consumer Durables, Realty), each with:
`index_label` (Nifty sub-index shown in the workbook), `keywords` (matched
against the company name, offline), `match` (matched against live
sector/industry text), and `peers` (3 representative large-caps). Extend
this table for sectors outside that list (media, textiles, capital goods,
etc.) the same way — add an entry with keywords + peers.

## Known limitations
- Assumes 10 historical annual years, matching Screener.in's current
  export format. A different count triggers a printed warning; the
  row/column structure will still need a quick manual check.
- Ticker resolution (NSE search → Yahoo search → alias table) can still
  fail for obscure/unlisted/newly-listed names — in which case sector
  classification falls back to the offline name-keyword match, and if
  *that* also doesn't match, peers are left as an explicit manual-entry
  placeholder (never guessed).
- ERP is a simplified price-only CAGR proxy, not a rigorous
  dividend-inclusive total-return premium — treat it as a reasonable live
  starting point, not a substitute for a proper ERP study.
- Google Finance and TradingView scraping/endpoints are inherently less
  stable than an official, versioned API — if either silently stops
  returning data after they change something on their end, that's
  expected; NSE and Yahoo remain the other two independent sources, and
  every field's cell comment tells you which source actually answered it.
- The Flask app is a local single-user dev server — fine for personal use,
  not intended to be exposed to the internet as-is.
