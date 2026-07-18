"""
app.py
Flask web interface for the Screener.in -> Financial Model engine.
Drag-and-drop (or click-to-browse) a Screener.in "Export to Excel" file,
and get back a full three-statement + DCF + comps + dashboard model.

Run:
    pip install -r requirements.txt
    python app.py
Then open http://127.0.0.1:5000 in a browser.
"""
import os
import re
import uuid
import traceback
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_from_directory

from screener_parser import parse_screener_file
from build_model import build_model
from generate_model import build_macro_and_market
import market_data as md

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
TEMPLATE_XLSX = os.path.join(BASE_DIR, "model_template.xlsx")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload cap


def _safe_slug(name):
    slug = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")
    return slug or "Company"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/generate", methods=["POST"])
def generate():
    f = request.files.get("file")
    if not f or f.filename == "":
        return jsonify({"ok": False, "error": "No file received."}), 400
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"ok": False, "error": "Please upload the Screener.in 'Export to Excel' file (.xlsx)."}), 400

    job_id = uuid.uuid4().hex[:8]
    in_path = os.path.join(UPLOAD_DIR, f"{job_id}_{f.filename}")
    f.save(in_path)

    log = []
    try:
        log.append("Parsing Screener.in export…")
        data = parse_screener_file(in_path)
        company = data["company_name"]
        log.append(f"Detected company: {company}")

        if data["n_years"] != 10:
            log.append(f"NOTE: {data['n_years']} historical years found (template assumes 10) — "
                       "review year headers in the output before relying on it.")

        if not md.REQUESTS_AVAILABLE:
            log.append("NOTE: 'requests' not installed on the server — NSE/TradingView/Google Finance live "
                       "data skipped; offline name-keyword sector classification still applies.")
        if not md.YF_AVAILABLE:
            log.append("NOTE: 'yfinance' not installed on the server — beta/ERP/secondary sector "
                       "data will use static fallback values (flagged in the workbook).")

        log.append("Resolving ticker, sector, peers, and cost-of-capital inputs…")
        macro, market = build_macro_and_market(company)

        if market["sector_key"]:
            src = ", ".join(market["profile"].get("sources_used", [])) or "name-keyword match only"
            log.append(f"Sector: {market['sector_key']} ({market['index_label']}) — {src}")
        else:
            log.append("Sector could not be classified — Comps sheet left as a manual-entry "
                       "placeholder (no default/guessed sector applied).")

        log.append("Building workbook (Cover / Assumptions / IS / BS / CF / Ratios / DCF / Comps / Dashboard)…")
        out_name = f"{_safe_slug(company)}_Financial_Model_{job_id}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        wb, labels = build_model(TEMPLATE_XLSX, out_path, data, macro, market)
        wb.save(out_path)
        log.append("Done.")

        return jsonify({
            "ok": True,
            "company": company,
            "sector_key": market["sector_key"],
            "index_label": market["index_label"],
            "years": labels,
            "current_price": data["current_price"],
            "download_url": f"/download/{out_name}",
            "filename": out_name,
            "log": log,
            "generated_at": datetime.now().strftime("%d %b %Y, %H:%M"),
        })
    except Exception as e:
        traceback.print_exc()
        log.append(f"ERROR: {e}")
        return jsonify({"ok": False, "error": str(e), "log": log}), 500
    finally:
        try:
            os.remove(in_path)
        except OSError:
            pass


@app.route("/download/<path:fname>")
def download(fname):
    return send_from_directory(OUTPUT_DIR, fname, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
