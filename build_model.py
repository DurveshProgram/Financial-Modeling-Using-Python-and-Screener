"""
build_model.py
Takes parsed Screener.in data + macro assumptions and populates the
professional three-statement model template (Cover / Assumptions / IS /
BS / CF / Ratios / DCF / Comps / Dashboard), producing a fully-linked,
company-specific Excel model.

Design principles:
- Historical actuals (columns B:K, i.e. the years actually reported by
  Screener.in) are hardcoded BLUE inputs, sourced straight from the
  uploaded screener file.
- All ratios, margins, DCF math, comps math etc. remain live FORMULAS
  (black/green per model conventions) -- untouched from the template,
  since that logic is company-agnostic.
- Forward projection assumptions (columns L:P) default to the company's
  own last actual year (FY closing value), each still a plain, editable
  input cell -- NOT a python-computed hardcode -- so changing the
  historical data or the assumption cell itself recalculates the whole
  model inside Excel.
- Cost-of-capital inputs (risk-free rate, equity risk premium) are
  populated from current published market data fetched by Claude,
  with the source noted in a cell comment. Beta and peer comps are
  flagged clearly as items to confirm/override since they cannot be
  reliably scraped for an arbitrary ticker without a market-data API.
"""
import openpyxl
from openpyxl.styles import Font
from openpyxl.comments import Comment

BLUE = Font(color="0000FF")
COLS = list("BCDEFGHIJKLMNOP")  # B..K = 10 historical, L..P = 5 projected


def _set(ws, row, col_letter, value, blue=True, comment=None):
    cell = ws[f"{col_letter}{row}"]
    cell.value = value
    if blue:
        cell.font = Font(name=cell.font.name or "Calibri", color="0000FF")
    if comment:
        cell.comment = Comment(comment, "Model Engine")


def _fill_historical_row(ws, row, values, n_hist, source_note=None):
    for i, v in enumerate(values[-n_hist:]):
        col = COLS[i]
        _set(ws, row, col, v if v is not None else 0, blue=True,
             comment=source_note if i == 0 else None)


def _flatline_projection(ws, row, n_hist, n_proj=5):
    """Projection columns default to '=<last actual col>' then carry the
    prior projected column forward, i.e. flat at last-actual-year level.
    This is the explicit, inspectable default used whenever a live web
    figure is not applicable to an individual P&L/BS line item."""
    last_hist_col = COLS[n_hist - 1]
    proj_cols = COLS[n_hist:n_hist + n_proj]
    prev = last_hist_col
    for c in proj_cols:
        ws[f"{c}{row}"] = f"={prev}{row}"
        prev = c


def build_model(template_path, output_path, data, macro):
    wb = openpyxl.load_workbook(template_path)
    n = data["n_years"]
    fy = data["fy_labels"]
    proj_labels = [f"FY{int(fy[-1][2:]) + i}E" for i in range(1, 6)]
    all_labels = fy + proj_labels
    company = data["company_name"]

    # ---------------- COVER ----------------
    cov = wb["Cover"]
    cov["C3"] = f"{company.upper()} — INTEGRATED FINANCIAL MODEL & INTRINSIC VALUATION"
    cov["E14"] = data["current_price"]
    cov["E15"] = data["shares_outstanding_latest"]
    cov["E16"] = round(data["current_price"] * data["shares_outstanding_latest"], 2)
    cov["C9"] = "Prepared: " + macro["prepared_date"]

    # ---------------- Year headers on every statement sheet ----------------
    for sheet_name, header_rows in [("Assumptions", [5]), ("IS", [5]), ("BS", [5]),
                                     ("CF", [5, 14]), ("Ratios", [5]), ("DCF", []),
                                     ]:
        ws = wb[sheet_name]
        ws["A2"] = f"{company} (Consolidated)"
        for hr in header_rows:
            for i, col in enumerate(COLS[:len(all_labels)]):
                ws[f"{col}{hr}"] = all_labels[i]

    for sheet_name in ["IS", "BS", "CF", "Ratios"]:
        wb[sheet_name]["A2"] = f"{company} (Consolidated)"
    wb["Comps"].cell(row=6, column=1).value = company
    wb["Dashboard"]["A1"] = f"{company.upper()} — VALUATION DASHBOARD"

    # ---------------- IS: historical actuals ----------------
    isws = wb["IS"]
    src = "Source: Screener.in export, uploaded by user"
    _fill_historical_row(isws, 6, data["sales"], n, src)
    _fill_historical_row(isws, 10, data["employee_cost"], n)
    _fill_historical_row(isws, 11, data["cost_of_materials"], n,
                          "= Raw Material + Power & Fuel + Other Mfg. Exp. (Screener.in)")
    _fill_historical_row(isws, 12, data["selling_admin"], n)
    _fill_historical_row(isws, 13, data["other_expenses"], n)
    _fill_historical_row(isws, 17, data["other_income"], n)
    _fill_historical_row(isws, 18, data["depreciation"], n)
    _fill_historical_row(isws, 20, data["interest"], n)
    _fill_historical_row(isws, 22, data["tax"], n)
    _fill_historical_row(isws, 26, data["dividend_amount"], n)
    _fill_historical_row(isws, 30, [s if s else data["shares_outstanding_latest"]
                                     for s in data["shares_outstanding_series"]], n)

    # ---------------- BS: historical actuals ----------------
    bsws = wb["BS"]
    _fill_historical_row(bsws, 6, data["net_block"], n, src)
    _fill_historical_row(bsws, 7, data["cwip"], n)
    _fill_historical_row(bsws, 8, data["investments"], n)
    _fill_historical_row(bsws, 9, data["receivables"], n)
    _fill_historical_row(bsws, 10, data["inventory"], n)
    _fill_historical_row(bsws, 11, data["cash_bank"], n)
    _fill_historical_row(bsws, 12, data["other_assets"], n)
    _fill_historical_row(bsws, 16, data["equity_share_capital"], n)
    _fill_historical_row(bsws, 17, data["reserves"], n)
    _fill_historical_row(bsws, 18, data["borrowings"], n)
    _fill_historical_row(bsws, 19, data["other_liabilities"], n)

    # ---------------- CF: historical actuals (reported) ----------------
    cfws = wb["CF"]
    _fill_historical_row(cfws, 6, data["cfo"], n, src)
    _fill_historical_row(cfws, 7, data["cfi"], n)
    _fill_historical_row(cfws, 8, data["cff"], n)

    # ---------------- Assumptions: WACC / macro block ----------------
    aws = wb["Assumptions"]
    aws["A1"] = f"{company} — Model Assumptions & Cost of Capital"
    _set(aws, 19, "C", macro["risk_free_rate"], comment=macro["risk_free_rate_source"])
    _set(aws, 20, "C", macro["equity_risk_premium"], comment=macro["erp_source"])
    _set(aws, 21, "C", macro["beta"], comment=macro["beta_source"])
    _set(aws, 23, "C", macro["pretax_cost_of_debt"], comment="Default: FY latest effective interest rate on borrowings — override with issuer's actual cost of debt/credit spread if known.")
    _set(aws, 32, "C", macro["terminal_growth"], comment=macro["terminal_growth_source"])
    _set(aws, 35, "C", data["current_price"], comment=src)
    _set(aws, 37, "C", data["shares_outstanding_latest"], comment=src)
    _set(aws, 38, "C", data["face_value"], comment=src)

    # ---------------- Assumptions: projection defaults = flat at last actual ----------------
    for row in range(6, 17):
        _flatline_projection(aws, row, n)
        for c in COLS[n:n + 5]:
            aws[f"{c}{row}"].font = Font(color="0000FF")

    return wb, all_labels