"""
generate_model.py
Usage:
    python generate_model.py <screener_export.xlsx> <output_model.xlsx>

Reads any Screener.in "export to Excel" file and produces a full,
professional three-statement model + DCF + comps + dashboard, in the
same structure as the reference TCS model, but generic to any company.
"""
import sys, datetime
from screener_parser import parse_screener_file
from build_model import build_model
import market_data as md

# ---- Macro / cost-of-capital assumptions -------------------------------
# Risk-free rate, ERP, and beta are now fetched LIVE from Yahoo Finance
# (see market_data.py) every time this script runs, instead of being
# hardcoded here. Only the pieces with no reliable live source (pre-tax
# cost of debt override, terminal growth) stay as editable static
# defaults — both are already surfaced as plain input cells in the
# Assumptions sheet so you can override them per-company anyway.
STATIC_DEFAULTS = {
    "pretax_cost_of_debt": 0.075,
    "terminal_growth": 0.05,
    "terminal_growth_source": "Default: long-run India nominal GDP growth proxy (~5%). Editable — should not exceed long-run nominal GDP growth of the economy the company primarily operates in.",
    "prepared_date": datetime.date.today().strftime("%B %Y"),
}


def build_macro_and_market(company_name):
    """Fetches risk-free rate, ERP, sector classification, peer multiples,
    and beta for the given company using the layered NSE -> TradingView -> Yahoo ->
    Google Finance -> offline-keyword pipeline in market_data.py. Returns
    (macro_dict, market_dict) ready for build_model()."""
    ids = md.resolve_company(company_name)                       # {nse_symbol, yahoo_ticker}
    profile = md.build_profile(company_name, ids)                # merged NSE+TradingView+Yahoo(+Google) fields
    sector_key = md.classify_sector(company_name, profile)       # name-keyword first, NEVER silently defaults

    rf, rf_source = md.get_risk_free_rate()
    erp, erp_source = md.get_equity_risk_premium(rf)
    beta, beta_source = md.get_beta(ids["yahoo_ticker"]) if ids["yahoo_ticker"] else (None, None)
    if beta is None:
        reason = (f"could not resolve a ticker for '{company_name}'" if not ids["yahoo_ticker"]
                   else f"live beta fetch failed for {ids['yahoo_ticker']} (Yahoo unavailable/blocked)")
        beta, beta_source = 0.90, (
            f"FALLBACK: default placeholder (broad-market-typical beta) — {reason}. "
            "Confirm actual 2Y/5Y levered beta from Bloomberg/NSE/a data terminal "
            "and override this cell.")

    peers, index_label = md.get_peer_multiples(sector_key, exclude_yahoo_ticker=ids["yahoo_ticker"])

    macro = {
        "risk_free_rate": rf,
        "risk_free_rate_source": rf_source,
        "equity_risk_premium": erp,
        "erp_source": erp_source,
        "beta": beta,
        "beta_source": beta_source,
        **STATIC_DEFAULTS,
    }
    market = {
        "ticker": ids["yahoo_ticker"],
        "nse_symbol": ids["nse_symbol"],
        "sector_key": sector_key,
        "index_label": index_label,
        "profile": profile,
        "peers": peers,
    }
    return macro, market


def main():
    if len(sys.argv) < 3:
        print("Usage: python generate_model.py <screener_export.xlsx> <output_model.xlsx>")
        sys.exit(1)
    screener_path, output_path = sys.argv[1], sys.argv[2]

    data = parse_screener_file(screener_path)
    if data["n_years"] != 10:
        print(f"NOTE: this export has {data['n_years']} historical years (template assumes 10). "
              f"Review the Assumptions/IS/BS/CF sheets' year headers before relying on the output.")

    if not md.REQUESTS_AVAILABLE:
        print("NOTE: 'requests' is not installed — run 'pip install requests' for NSE India / TradingView / "
              "Google Finance live data. Only offline name-keyword sector classification and "
              "static defaults will be used for this run.")
    if not md.YF_AVAILABLE:
        print("NOTE: yfinance is not installed — run 'pip install yfinance' for live "
              "beta / ERP / secondary sector data. Falling back to static "
              "defaults for this run (clearly flagged in the workbook's cell comments).")

    macro, market = build_macro_and_market(data["company_name"])
    if not market["ticker"]:
        print(f"NOTE: could not resolve any ticker (NSE or Yahoo) for '{data['company_name']}' — "
              "beta and 52-week range will use fallback/placeholder values. Check the cell "
              "comments on the Cover and Assumptions sheets.")
    if market["sector_key"] is None:
        print(f"NOTE: could not confidently classify the sector for '{data['company_name']}' from "
              "its name or live data — the Comps sheet peer rows are left as an explicit "
              "'enter manually' placeholder rather than guessing (no default-to-IT fallback).")

    wb, labels = build_model("model_template.xlsx", output_path, data, macro, market)
    wb.save(output_path)
    print(f"Built model for {data['company_name']} -> {output_path}")
    print("Years:", labels)
    if market["sector_key"]:
        print(f"Sector classification: {market['sector_key']} ({market['index_label']}) "
              f"-- sources used: {', '.join(market['profile'].get('sources_used', [])) or 'name-keyword match only'}")
    else:
        print("Sector classification: NOT RESOLVED -- see Comps sheet for manual peer entry")


if __name__ == "__main__":
    main()
