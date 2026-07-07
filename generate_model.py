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

# ---- Macro / cost-of-capital assumptions -------------------------------
# These are the pieces of a DCF that are genuinely macro (not
# company-specific line items) and so CAN be sourced from the web.
# Values below reflect current published data as of the run date;
# re-run with updated figures periodically.
MACRO = {
    "risk_free_rate": 0.0672,
    "risk_free_rate_source": "Source: India 10-Year G-Sec yield, TradingEconomics, 3-Jul-2026 (~6.72%). Update periodically.",
    "equity_risk_premium": 0.07,
    "erp_source": "Source: Approx. India equity risk premium (mature-market ERP + India country risk premium), per Damodaran country-risk framework, 2026 update. Editable — confirm against latest Damodaran dataset (pages.stern.nyu.edu/~adamodar).",
    "beta": 0.90,
    "beta_source": "Default placeholder (unlevered sector-typical beta for large-cap IT services). COULD NOT be reliably auto-fetched for an arbitrary ticker — confirm actual 2Y/5Y levered beta from Bloomberg/NSE/a data terminal and override this cell.",
    "pretax_cost_of_debt": 0.075,
    "terminal_growth": 0.05,
    "terminal_growth_source": "Default: long-run India nominal GDP growth proxy (~5%). Editable — should not exceed long-run nominal GDP growth of the economy the company primarily operates in.",
    "prepared_date": datetime.date.today().strftime("%B %Y"),
}


def main():
    if len(sys.argv) < 3:
        print("Usage: python generate_model.py <screener_export.xlsx> <output_model.xlsx>")
        sys.exit(1)
    screener_path, output_path = sys.argv[1], sys.argv[2]

    data = parse_screener_file(screener_path)
    if data["n_years"] != 10:
        print(f"NOTE: this export has {data['n_years']} historical years (template assumes 10). "
              f"Review the Assumptions/IS/BS/CF sheets' year headers before relying on the output.")

    wb, labels = build_model("model_template.xlsx", output_path, data, MACRO)
    wb.save(output_path)
    print(f"Built model for {data['company_name']} -> {output_path}")
    print("Years:", labels)


if __name__ == "__main__":
    main()
