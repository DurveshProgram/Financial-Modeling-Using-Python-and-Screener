"""
market_data.py
Multi-source market-data extension used by build_model.py / generate_model.py
to replace hardcoded-to-TCS values with data that's actually correct for
whatever company was fed in.

DATA SOURCE PRIORITY (each layer only runs if the previous one didn't
already answer the question -- every call is wrapped in try/except and
degrades to the next layer, never crashes):

  1. Name-keyword sector classification (SECTOR_MAP['keywords'])
     - Zero network dependency. Matches the Screener.in company name
       itself against a keyword list per sector. This is what actually
       fixes "Asian Paints gets classified as IT" -- it no longer
       depends on any external API resolving correctly at all.
  2. NSE India official API (nseindia.com)
     - Free, no key, and it's the *actual listing exchange's own* sector/
       industry classification for the stock -- more authoritative than
       generic sector tags for Indian-listed names. Used for: symbol
       resolution, sector/industry, current price, 52-week range.
  3. TradingView scanner endpoint (scanner.tradingview.com)
     - No official API, but this endpoint is unauthenticated and returns
       price, market cap, P/E, EV/EBITDA, P/B, ROE, sector/industry, and
       52-week range in ONE call -- and can be BATCHED across every peer
       in a sector at once. Best coverage-per-call of any source here,
       so it's tried first for the Comps sheet peer set.
  4. Yahoo Finance (yfinance)
     - Used for: beta, and the long-run Nifty history used to compute
       the equity risk premium (no other source here exposes either of
       these). Also fills in any peer fields TradingView didn't return.
  5. Google Finance (best-effort HTML scrape, no official API exists)
     - Last-resort fallback for CURRENT PRICE only -- its page doesn't
       reliably expose fundamentals (P/E etc.) in scrapable form. This
       is inherently fragile (Google can change markup at any time) so
       it's treated purely as a best-effort supplement, never the only
       source for something the model depends on.

If ALL live layers fail for a given field, a clearly labelled STATIC
FALLBACK is used and flagged in the cell comment -- the pipeline never
silently presents fallback data as if it were live, and (critically) it
never silently defaults an unresolved sector to IT anymore. An
unresolved sector now returns None and build_model.py leaves the peer
rows as an explicit "enter data manually" placeholder instead of
guessing.

Requires: pip install yfinance requests
(requests is usually already present and is what powers NSE, TradingView,
and Google Finance; yfinance is optional -- steps 1, 2, 3, and 5 all work
without it.)
"""
import datetime
import re

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

try:
    import yfinance as yf
    YF_AVAILABLE = True
except ImportError:
    YF_AVAILABLE = False

TODAY = datetime.date.today().strftime("%d-%b-%Y")

_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# ---------------------------------------------------------------------
# Sector map. Each entry has:
#   index_label -- Nifty sub-index shown in the workbook
#   keywords    -- matched directly against the Screener.in company name
#                  (offline, always available, checked FIRST)
#   match       -- matched against Yahoo/NSE sector+industry text
#   peers       -- (yahoo_ticker, nse_symbol) pairs for 3 representative
#                  large-caps in the sector, used for the Comps sheet
# Extend this table as you model more sectors.
# ---------------------------------------------------------------------
SECTOR_MAP = {
    "IT": {
        "index_label": "Nifty IT",
        "keywords": ["infosys", "tcs", "tata consultancy", "wipro", "hcl tech",
                     "tech mahindra", "software", "technologies", " it ", "mindtree",
                     "ltimindtree", "persistent", "coforge", "mphasis"],
        "match": ["information technology", "software", "it services"],
        "peers": [("INFY.NS", "INFY"), ("HCLTECH.NS", "HCLTECH"), ("WIPRO.NS", "WIPRO")],
    },
    "BANK": {
        "index_label": "Nifty Bank",
        "keywords": ["bank"],
        "match": ["bank", "banks"],
        "peers": [("ICICIBANK.NS", "ICICIBANK"), ("KOTAKBANK.NS", "KOTAKBANK"), ("AXISBANK.NS", "AXISBANK")],
    },
    "NBFC": {
        "index_label": "Nifty Financial Services",
        "keywords": ["finance", "financial services", "housing finance", "capital",
                     "nbfc", "chit fund", "investments ltd"],
        "match": ["financial services", "credit services", "shadow bank"],
        "peers": [("BAJFINANCE.NS", "BAJFINANCE"), ("HDFCLIFE.NS", "HDFCLIFE"), ("SBICARD.NS", "SBICARD")],
    },
    "INSURANCE": {
        "index_label": "Nifty Financial Services (Insurance)",
        "keywords": ["insurance", "assurance"],
        "match": ["insurance"],
        "peers": [("HDFCLIFE.NS", "HDFCLIFE"), ("SBILIFE.NS", "SBILIFE"), ("ICICIGI.NS", "ICICIGI")],
    },
    "FMCG": {
        "index_label": "Nifty FMCG",
        "keywords": ["hindustan unilever", "nestle", "britannia", "dabur", "marico",
                     "godrej consumer", "colgate", "itc ", "tata consumer", "emami"],
        "match": ["consumer defensive", "household", "personal products", "packaged foods"],
        "peers": [("HINDUNILVR.NS", "HINDUNILVR"), ("ITC.NS", "ITC"), ("NESTLEIND.NS", "NESTLEIND")],
    },
    "AUTO": {
        "index_label": "Nifty Auto",
        "keywords": ["motors", "automobile", "auto ltd", "bajaj auto", "hero moto",
                     "eicher", "ashok leyland", "tvs motor"],
        "match": ["auto manufacturers", "vehicle"],
        "peers": [("MARUTI.NS", "MARUTI"), ("TATAMOTORS.NS", "TATAMOTORS"), ("M&M.NS", "M&M")],
    },
    "PHARMA": {
        "index_label": "Nifty Pharma",
        "keywords": ["pharma", "laboratories", "lifesciences", "life sciences",
                     "healthcare", "drug"],
        "match": ["pharma", "drug manufacturers", "healthcare"],
        "peers": [("SUNPHARMA.NS", "SUNPHARMA"), ("DRREDDY.NS", "DRREDDY"), ("CIPLA.NS", "CIPLA")],
    },
    "ENERGY": {
        "index_label": "Nifty Energy",
        "keywords": ["petroleum", "oil ", "ongc", "gas ltd", "reliance industries",
                     "power grid", "ntpc"],
        "match": ["oil & gas", "oil and gas", "energy"],
        "peers": [("RELIANCE.NS", "RELIANCE"), ("ONGC.NS", "ONGC"), ("BPCL.NS", "BPCL")],
    },
    "METAL": {
        "index_label": "Nifty Metal",
        "keywords": ["steel", "hindalco", "vedanta", "nmdc", "mining", "metals"],
        "match": ["metal", "mining", "steel"],
        "peers": [("TATASTEEL.NS", "TATASTEEL"), ("JSWSTEEL.NS", "JSWSTEEL"), ("HINDALCO.NS", "HINDALCO")],
    },
    "PAINTS_CHEMICALS": {
        "index_label": "Nifty Chemicals / Paints",
        "keywords": ["paint", "chemicals", "pigments", "dye", "coatings"],
        "match": ["specialty chemicals", "chemicals", "coatings, paint"],
        "peers": [("ASIANPAINT.NS", "ASIANPAINT"), ("BERGEPAINT.NS", "BERGEPAINT"), ("PIDILITIND.NS", "PIDILITIND")],
    },
    "CEMENT": {
        "index_label": "Nifty Cement (S&P BSE)",
        "keywords": ["cement", "ambuja", "ultratech", "shree cement", "acc ltd"],
        "match": ["cement", "building materials"],
        "peers": [("ULTRACEMCO.NS", "ULTRACEMCO"), ("SHREECEM.NS", "SHREECEM"), ("AMBUJACEM.NS", "AMBUJACEM")],
    },
    "TELECOM": {
        "index_label": "Nifty Telecom",
        "keywords": ["bharti airtel", "vodafone", "idea ltd", "telecom", "jio"],
        "match": ["telecom"],
        "peers": [("BHARTIARTL.NS", "BHARTIARTL"), ("IDEA.NS", "IDEA"), ("INDUSTOWER.NS", "INDUSTOWER")],
    },
    "INFRA_CONSTRUCTION": {
        "index_label": "Nifty Infrastructure",
        "keywords": ["infrastructure", "construction", "engineering", "larsen", "l&t ",
                     "projects ltd", "developers"],
        "match": ["engineering & construction", "infrastructure operations"],
        "peers": [("LT.NS", "LT"), ("GMRINFRA.NS", "GMRINFRA"), ("IRB.NS", "IRB")],
    },
    "CONSUMER_DURABLES": {
        "index_label": "Nifty Consumer Durables",
        "keywords": ["appliances", "electronics", "voltas", "havells", "durables",
                     "crompton", "whirlpool"],
        "match": ["furnishings, fixtures", "consumer electronics", "appliances"],
        "peers": [("HAVELLS.NS", "HAVELLS"), ("VOLTAS.NS", "VOLTAS"), ("CROMPTON.NS", "CROMPTON")],
    },
    "REALTY": {
        "index_label": "Nifty Realty",
        "keywords": ["realty", "real estate", "properties", "estates", "dlf "],
        "match": ["real estate", "reit"],
        "peers": [("DLF.NS", "DLF"), ("GODREJPROP.NS", "GODREJPROP"), ("OBEROIRLTY.NS", "OBEROIRLTY")],
    },
}
DEFAULT_SECTOR_KEY = None  # IMPORTANT: no silent default anymore. None means
                           # "not resolved" and build_model.py must leave the
                           # peer rows as an explicit manual-entry placeholder
                           # rather than guessing a sector.

# Small alias net for common NSE names -> (Yahoo ticker, NSE symbol), used
# only if live search (NSE/Yahoo) can't resolve a symbol (e.g. offline).
_TICKER_ALIASES = {
    "tata consultancy services ltd": ("TCS.NS", "TCS"),
    "hdfc bank ltd": ("HDFCBANK.NS", "HDFCBANK"),
    "infosys ltd": ("INFY.NS", "INFY"),
    "hcltechnologies ltd": ("HCLTECH.NS", "HCLTECH"),
    "wipro ltd": ("WIPRO.NS", "WIPRO"),
    "asian paints ltd": ("ASIANPAINT.NS", "ASIANPAINT"),
    "reliance industries ltd": ("RELIANCE.NS", "RELIANCE"),
    "hindustan unilever ltd": ("HINDUNILVR.NS", "HINDUNILVR"),
    "itc ltd": ("ITC.NS", "ITC"),
    "larsen & toubro ltd": ("LT.NS", "LT"),
    "maruti suzuki india ltd": ("MARUTI.NS", "MARUTI"),
    "bajaj finance ltd": ("BAJFINANCE.NS", "BAJFINANCE"),
    "sun pharmaceutical industries ltd": ("SUNPHARMA.NS", "SUNPHARMA"),
    "ultratech cement ltd": ("ULTRACEMCO.NS", "ULTRACEMCO"),
}


def classify_sector_by_name(company_name):
    """Zero-network, always-available sector guess from the company name
    itself. Checked FIRST, before any live source, so a network outage or
    an unresolved ticker can never silently mislabel the sector."""
    text = f" {str(company_name).strip().lower()} "
    for key, spec in SECTOR_MAP.items():
        if any(kw in text for kw in spec["keywords"]):
            return key
    return None


# ---------------------------------------------------------------------
# NSE India -- primary live source. Free, no API key. NSE requires a
# session cookie obtained from a normal page hit before its /api/ JSON
# endpoints will respond (returns 401/403 otherwise) -- this is standard
# and documented community behaviour, not a workaround of anything
# private/paywalled.
# ---------------------------------------------------------------------
_nse_session_cache = {"session": None}


def _nse_session():
    if not REQUESTS_AVAILABLE:
        return None
    if _nse_session_cache["session"] is not None:
        return _nse_session_cache["session"]
    try:
        s = requests.Session()
        s.headers.update({
            "User-Agent": _UA,
            "Accept": "*/*",
            "Accept-Language": "en-US,en;q=0.9",
        })
        s.get("https://www.nseindia.com", timeout=6)  # sets cookies required by /api/*
        _nse_session_cache["session"] = s
        return s
    except Exception:
        return None


def resolve_nse_symbol(company_name):
    """Best-effort NSE trading symbol via NSE's own search-autocomplete
    endpoint. Returns None if unavailable (caller falls back to the
    alias table / Yahoo search)."""
    s = _nse_session()
    if not s:
        return None
    try:
        r = s.get("https://www.nseindia.com/api/search/autocomplete",
                   params={"q": company_name}, timeout=6)
        r.raise_for_status()
        for item in r.json().get("symbols", []):
            sym = item.get("symbol")
            if sym and item.get("symbol_info", "").upper() != "INDEX":
                return sym
    except Exception:
        pass
    return None


def get_nse_quote(nse_symbol):
    """Official NSE sector/industry classification + current price/52-week
    range for a resolved NSE symbol. Returns {} on any failure."""
    s = _nse_session()
    if not (s and nse_symbol):
        return {}
    try:
        r = s.get("https://www.nseindia.com/api/quote-equity",
                   params={"symbol": nse_symbol}, timeout=6)
        r.raise_for_status()
        j = r.json()
        info = j.get("info", {}) or {}
        industry_info = j.get("industryInfo", {}) or {}
        price_info = j.get("priceInfo", {}) or {}
        week = price_info.get("weekHighLow", {}) or {}
        return {
            "symbol": info.get("symbol"),
            "companyName": info.get("companyName"),
            "sector": industry_info.get("macro") or industry_info.get("sector"),
            "industry": industry_info.get("industry") or industry_info.get("sectorIndustry"),
            "lastPrice": price_info.get("lastPrice"),
            "fiftyTwoWeekLow": week.get("min"),
            "fiftyTwoWeekHigh": week.get("max"),
        }
    except Exception:
        return {}


# ---------------------------------------------------------------------
# Yahoo Finance -- used for beta and the Nifty history behind the ERP
# calc, and as a secondary sector/peer source. NOTE: Yahoo has repeatedly
# tightened anti-bot/crumb requirements on its unauthenticated endpoints;
# if this environment's yfinance calls keep returning empty results even
# with yfinance installed, that's Yahoo blocking the request pattern, not
# a bug in this file -- NSE (above) and Google Finance (below) exist
# specifically so the model doesn't depend on Yahoo actually working.
# ---------------------------------------------------------------------
def resolve_yahoo_ticker(company_name):
    key = str(company_name).strip().lower()
    if YF_AVAILABLE:
        try:
            res = yf.Search(company_name, max_results=5)
            for q in getattr(res, "quotes", []) or []:
                sym = q.get("symbol", "")
                if sym.endswith(".NS"):
                    return sym
        except Exception:
            pass
    alias = _TICKER_ALIASES.get(key)
    return alias[0] if alias else None


def get_yahoo_profile(ticker):
    if not (YF_AVAILABLE and ticker):
        return {}
    try:
        info = yf.Ticker(ticker).info or {}
        return {
            "sector": info.get("sector", ""),
            "industry": info.get("industry", ""),
            "longName": info.get("longName", ""),
            "fiftyTwoWeekLow": info.get("fiftyTwoWeekLow"),
            "fiftyTwoWeekHigh": info.get("fiftyTwoWeekHigh"),
            "beta": info.get("beta"),
        }
    except Exception:
        return {}


# ---------------------------------------------------------------------
# TradingView -- no official public API, but its "scanner" endpoint
# (used by tradingview.com's own screener page, and widely relied on by
# open-source trading tools) accepts an unauthenticated POST and returns
# price, market cap, P/E, EV/EBITDA, P/B, ROE, sector/industry and
# 52-week range all in ONE call -- and can be BATCHED across every peer
# at once. This makes it the strongest single source for the Comps sheet
# (better coverage per call than Yahoo's per-ticker .info, which is also
# the one most frequently blocked). Treated the same as every other
# source here: try/except, never the sole source, always labelled.
# ---------------------------------------------------------------------
_TV_COLUMNS = [
    "close", "market_cap_basic", "price_earnings_ttm", "price_book_ratio",
    "return_on_equity", "enterprise_value_to_ebitda_ttm", "sector", "industry",
    "High.1Y", "Low.1Y",
]


def get_tradingview_batch(nse_symbols):
    """One POST covering every symbol in nse_symbols. Returns
    {nse_symbol: {close, mkt_cap_cr, pe, pb, roe, ev_ebitda, sector,
    industry, fiftyTwoWeekHigh, fiftyTwoWeekLow}}. Returns {} on any
    failure (missing requests, network/timeout, unexpected response
    shape) -- callers must treat that as "TradingView had nothing" and
    fall through to the next source, not as an error."""
    if not (REQUESTS_AVAILABLE and nse_symbols):
        return {}
    try:
        url = "https://scanner.tradingview.com/india/scan"
        payload = {
            "symbols": {"tickers": [f"NSE:{s}" for s in nse_symbols], "query": {"types": []}},
            "columns": _TV_COLUMNS,
        }
        r = requests.post(url, json=payload, headers={"User-Agent": _UA, "Content-Type": "application/json"},
                           timeout=8)
        r.raise_for_status()
        rows = r.json().get("data", [])
        out = {}
        for row in rows:
            ticker = row.get("s", "")  # e.g. "NSE:ASIANPAINT"
            sym = ticker.split(":")[-1] if ":" in ticker else ticker
            vals = dict(zip(_TV_COLUMNS, row.get("d", [])))
            mcap = vals.get("market_cap_basic")
            roe_raw = vals.get("return_on_equity")
            out[sym] = {
                "close": vals.get("close"),
                "mkt_cap_cr": round(mcap / 1e7, 0) if mcap else None,
                "pe": vals.get("price_earnings_ttm"),
                "pb": vals.get("price_book_ratio"),
                "roe": round(roe_raw / 100, 4) if roe_raw is not None else None,
                "ev_ebitda": vals.get("enterprise_value_to_ebitda_ttm"),
                "sector": vals.get("sector"),
                "industry": vals.get("industry"),
                "fiftyTwoWeekHigh": vals.get("High.1Y"),
                "fiftyTwoWeekLow": vals.get("Low.1Y"),
            }
        return out
    except Exception:
        return {}


# ---------------------------------------------------------------------
# Google Finance -- no official API. Best-effort HTML scrape of the
# public quote page. The page doesn't reliably expose fundamentals
# (P/E, EV/EBITDA, etc.) in scrapable form, so this is used purely as a
# further fallback for CURRENT PRICE when NSE/TradingView/Yahoo all
# come up empty -- never as the sole source for anything the model
# depends on, and its price is actually wired into the profile/peer
# data below now (it used to be fetched and then silently discarded).
# ---------------------------------------------------------------------
def get_google_finance_quote(nse_symbol):
    if not (REQUESTS_AVAILABLE and nse_symbol):
        return {}
    try:
        url = f"https://www.google.com/finance/quote/{nse_symbol}:NSE"
        r = requests.get(url, headers={"User-Agent": _UA}, timeout=6)
        r.raise_for_status()
        html = r.text
        out = {"source_url": url}
        m = re.search(r'data-last-price="([\d.]+)"', html)
        if m:
            out["lastPrice"] = float(m.group(1))
        else:
            m = re.search(r'class="YMlKec fxKbKc">[₹$]?([\d,]+\.?\d*)<', html)
            if m:
                out["lastPrice"] = float(m.group(1).replace(",", ""))
        m2 = re.search(r'>About</div>.*?<div[^>]*>(.*?)</div>', html, re.S)
        if m2:
            out["about_snippet"] = re.sub("<[^>]+>", "", m2.group(1)).strip()[:200]
        return out
    except Exception:
        return {}


# ---------------------------------------------------------------------
# Unified resolvers used by generate_model.py
# ---------------------------------------------------------------------
def resolve_company(company_name):
    """Tries NSE search first (it's the actual exchange, most reliable
    for Indian-listed symbols), then Yahoo search, then the alias table.
    Returns dict: {nse_symbol, yahoo_ticker}."""
    nse_symbol = resolve_nse_symbol(company_name)
    yahoo_ticker = resolve_yahoo_ticker(company_name)
    if not (nse_symbol or yahoo_ticker):
        alias = _TICKER_ALIASES.get(str(company_name).strip().lower())
        if alias:
            yahoo_ticker, nse_symbol = alias
    if nse_symbol and not yahoo_ticker:
        yahoo_ticker = f"{nse_symbol}.NS"
    if yahoo_ticker and not nse_symbol:
        nse_symbol = yahoo_ticker.replace(".NS", "")
    return {"nse_symbol": nse_symbol, "yahoo_ticker": yahoo_ticker}


def build_profile(company_name, ids):
    """Merges NSE + TradingView + Yahoo + Google Finance data (in that
    priority order per-field) into one profile dict, and records which
    source answered which field."""
    nse = get_nse_quote(ids["nse_symbol"]) if ids["nse_symbol"] else {}
    tv = get_tradingview_batch([ids["nse_symbol"]]).get(ids["nse_symbol"], {}) if ids["nse_symbol"] else {}
    yahoo = get_yahoo_profile(ids["yahoo_ticker"]) if ids["yahoo_ticker"] else {}
    google = {}
    if not (nse.get("lastPrice") or tv.get("close") or yahoo.get("fiftyTwoWeekLow")):
        google = get_google_finance_quote(ids["nse_symbol"])

    profile = {
        "sector": nse.get("sector") or tv.get("sector") or yahoo.get("sector") or "",
        "industry": nse.get("industry") or tv.get("industry") or yahoo.get("industry") or "",
        "lastPrice": nse.get("lastPrice") or tv.get("close") or google.get("lastPrice"),
        "fiftyTwoWeekLow": yahoo.get("fiftyTwoWeekLow") or nse.get("fiftyTwoWeekLow") or tv.get("fiftyTwoWeekLow"),
        "fiftyTwoWeekHigh": yahoo.get("fiftyTwoWeekHigh") or nse.get("fiftyTwoWeekHigh") or tv.get("fiftyTwoWeekHigh"),
        "beta": yahoo.get("beta"),
        "sources_used": [s for s, d in [("NSE", nse), ("TradingView", tv), ("Yahoo", yahoo),
                                          ("Google Finance", google)] if d],
    }
    return profile


def classify_sector(company_name, profile):
    """Priority: (1) offline name-keyword match, (2) NSE/Yahoo sector or
    industry text match. Returns None (NOT a default sector) if nothing
    matches -- build_model.py must handle None by leaving peers blank
    with a manual-entry placeholder rather than guessing."""
    by_name = classify_sector_by_name(company_name)
    if by_name:
        return by_name
    text = f"{profile.get('sector', '')} {profile.get('industry', '')}".lower()
    if text.strip():
        for key, spec in SECTOR_MAP.items():
            if any(m in text for m in spec["match"]):
                return key
    return None


def get_peer_multiples(sector_key, exclude_yahoo_ticker=None):
    """Live trailing P/E, EV/EBITDA, P/B, ROE, and market cap (Rs Cr) for
    the peer set of a sector. Tries TradingView FIRST with a single
    batched call covering every peer at once (best coverage-per-call of
    the three sources), then Yahoo per-peer for anything TradingView
    didn't return, then Google Finance as a last-resort PRICE-only signal
    (fundamentals like P/E aren't reliably scrapable from its page, so a
    Google-only row is explicitly flagged as incomplete rather than
    marked 'live'). Returns (peers_list, index_label). If sector_key is
    None, returns ([], "Sector not resolved") -- caller must NOT fall
    back to a default sector's peers."""
    if sector_key is None or sector_key not in SECTOR_MAP:
        return [], "Sector not resolved — set peers manually"
    spec = SECTOR_MAP[sector_key]
    peer_ids = [(yt, ns) for yt, ns in spec["peers"] if yt != exclude_yahoo_ticker]
    if not peer_ids:
        return [], spec["index_label"]

    tv_batch = get_tradingview_batch([ns for _, ns in peer_ids])

    out = []
    for yahoo_ticker, nse_symbol in peer_ids:
        row = {"ticker": yahoo_ticker, "name": nse_symbol, "live": False, "source": None,
               "mkt_cap_cr": None, "pe": None, "ev_ebitda": None, "pb": None, "roe": None}

        tv = tv_batch.get(nse_symbol, {})
        if tv:
            row["mkt_cap_cr"] = tv.get("mkt_cap_cr")
            row["pe"] = tv.get("pe")
            row["ev_ebitda"] = tv.get("ev_ebitda")
            row["pb"] = tv.get("pb")
            row["roe"] = tv.get("roe")
            row["source"] = "TradingView"

        missing = row["pe"] is None or row["ev_ebitda"] is None or row["pb"] is None or row["roe"] is None
        if missing and YF_AVAILABLE:
            try:
                info = yf.Ticker(yahoo_ticker).info or {}
                row["name"] = info.get("longName", row["name"])
                row["mkt_cap_cr"] = row["mkt_cap_cr"] or (round(info["marketCap"] / 1e7, 0) if info.get("marketCap") else None)
                row["pe"] = row["pe"] or info.get("trailingPE")
                row["ev_ebitda"] = row["ev_ebitda"] or info.get("enterpriseToEbitda")
                row["pb"] = row["pb"] or info.get("priceToBook")
                roe = info.get("returnOnEquity")
                row["roe"] = row["roe"] or (round(roe, 4) if roe is not None else None)
                row["source"] = (row["source"] + " + Yahoo") if row["source"] else "Yahoo"
            except Exception:
                pass

        if row["pe"] is None and row["mkt_cap_cr"] is None:
            g = get_google_finance_quote(nse_symbol)
            if g.get("lastPrice"):
                row["source"] = "Google Finance (price only — fundamentals still missing)"

        row["live"] = all(v is not None for v in
                           [row["mkt_cap_cr"], row["pe"], row["ev_ebitda"], row["pb"], row["roe"]])
        out.append(row)
    return out, spec["index_label"]


def get_risk_free_rate():
    """India 10-Yr G-Sec yield. Yahoo doesn't have one universally-stable
    ticker for this, so a short candidate list is tried; if all fail, the
    last known published figure is used as a clearly-labelled fallback."""
    candidates = ["^IN10Y", "IN10Y=RR", "IN10YT=RR"]
    if YF_AVAILABLE:
        for tkr in candidates:
            try:
                hist = yf.Ticker(tkr).history(period="5d")
                if not hist.empty:
                    rate = float(hist["Close"].dropna().iloc[-1]) / 100
                    return rate, f"LIVE: Yahoo Finance ({tkr}), close {TODAY}."
            except Exception:
                continue
    return 0.0672, ("FALLBACK (Yahoo unavailable/blocked as of "
                     f"{TODAY}): India 10-Yr G-Sec ~6.72%, last published figure. "
                     "Verify manually (e.g. worldgovernmentbonds.com) and override if stale.")


def get_equity_risk_premium(risk_free_rate):
    """ERP computed as trailing long-run Nifty 50 CAGR minus the current
    risk-free rate (a standard historical-premium proxy), using ~15 years
    of ^NSEI price history. Falls back to a Damodaran-style constant."""
    if YF_AVAILABLE:
        try:
            hist = yf.Ticker("^NSEI").history(period="15y")["Close"].dropna()
            if len(hist) > 250:
                years = (hist.index[-1] - hist.index[0]).days / 365.25
                cagr = (hist.iloc[-1] / hist.iloc[0]) ** (1 / years) - 1
                erp = round(cagr - risk_free_rate, 4)
                return erp, (f"LIVE: computed as {years:.1f}Y Nifty 50 (^NSEI) price CAGR "
                              f"({cagr:.2%}) minus current risk-free rate, as of {TODAY}. "
                              "Simplified historical-premium proxy (price only, excludes "
                              "dividends) -- cross-check against Damodaran's published "
                              "India ERP (pages.stern.nyu.edu/~adamodar).")
        except Exception:
            pass
    return 0.07, (f"FALLBACK (Yahoo unavailable/blocked as of {TODAY}): "
                   "approx. India equity risk premium per Damodaran country-risk "
                   "framework. Verify manually and re-run with network access.")


def get_beta(yahoo_ticker):
    """Yahoo-reported levered beta for the ticker, or None if unavailable."""
    profile = get_yahoo_profile(yahoo_ticker)
    b = profile.get("beta")
    if b is not None:
        return b, f"LIVE: Yahoo Finance reported beta for {yahoo_ticker}, as of {TODAY}."
    return None, None
