"""
TaxLedger Pro — Web Server

FastAPI backend for the ESPP2 tax reporting web application.

Endpoints:
  GET  /                      → serve SPA (index.html)
  GET  /static/{path}         → static files
  GET  /api/settings          → get saved settings
  POST /api/settings          → save settings (year, currencies)
  POST /api/settings/rates    → upload Norges Bank exchange rate CSV
  POST /api/process           → upload PDFs + wire data → run tax calculation
  GET  /api/job/{job_id}      → poll job status
  GET  /api/results/{job_id}  → get full results JSON
  GET  /api/download/{job_id} → download xlsx report
"""

import json
import logging
import os
import tempfile
import traceback
import uuid
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

import simplejson
import uvicorn
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# App & paths
# ---------------------------------------------------------------------------

app = FastAPI(title="TaxLedger Pro")

STATIC_DIR = Path(__file__).parent / "static"
SETTINGS_FILE = Path.home() / ".espp2" / "settings.json"
SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)

# In-memory job store: job_id → {status, phase, pct, message, result, xlsx_bytes, account_id}
_jobs: dict = {}

# ---------------------------------------------------------------------------
# Settings helpers
# ---------------------------------------------------------------------------

_default_settings = {
    "year": date.today().year - 1,
    "source_currency": "USD",
    "destination_currency": "NOK",
    "rates_csv_path": None,
    "rates_days": 0,
}


def _load_settings() -> dict:
    if SETTINGS_FILE.exists():
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                s = dict(_default_settings)
                s.update(saved)
                return s
        except Exception:
            pass
    return dict(_default_settings)


def _save_settings(s: dict):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(s, f, indent=2)


# ---------------------------------------------------------------------------
# Static files & SPA
# ---------------------------------------------------------------------------

@app.get("/")
async def serve_index():
    index = STATIC_DIR / "index.html"
    if not index.exists():
        return Response("TaxLedger Pro — static files not found", status_code=500)
    return FileResponse(str(index))


app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


# ---------------------------------------------------------------------------
# Settings API
# ---------------------------------------------------------------------------

@app.get("/api/settings")
async def get_settings():
    s = _load_settings()
    from espp2.fmv import FMV
    rates_days = sum(len(v) for v in FMV._local_rates.values())
    return {
        "year": s["year"],
        "rates_loaded": rates_days > 0,
        "rates_days": rates_days,
    }


@app.post("/api/fetch-rates")
async def fetch_rates(year: int = Form(...)):
    """Fetch USD/NOK exchange rates from Norges Bank API for the given reporting year."""
    from espp2.fmv import FMV
    try:
        n = FMV.fetch_norges_bank_rates(year)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Norges Bank API error: {e}")
    s = _load_settings()
    s["year"] = year
    _save_settings(s)
    return {"ok": True, "days_loaded": n, "year": year}


@app.post("/api/parse-wires")
async def parse_wires(reporting_year_pdf: UploadFile = File(...)):
    """Parse WIRE entries from a Morgan Stanley PDF. Returns [{date, usd}] list."""
    import io as _io
    import asyncio
    from espp2.plugins.morgan_pdf import read as pdf_read
    from espp2.datamodels import EntryTypeEnum

    content = await reporting_year_pdf.read()
    filename = reporting_year_pdf.filename or ""

    # Run in thread executor so the event loop is not blocked by pdfplumber
    def _parse():
        transactions = pdf_read(_io.BytesIO(content), filename=filename)
        wires = [
            {"date": str(t.date), "usd": float(abs(t.amount.value))}
            for t in transactions.transactions
            if t.type == EntryTypeEnum.WIRE
        ]
        # Extract the reporting year from the PDF period (todate year)
        year = transactions.todate.year if transactions.todate else None
        return {"wires": wires, "year": year}

    loop = asyncio.get_event_loop()
    result = await loop.run_in_executor(None, _parse)
    return result


# ---------------------------------------------------------------------------
# Tax processing API
# ---------------------------------------------------------------------------

def _decimal_default(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError(f"Not serializable: {type(obj)}")


def _run_tax_calculation(job_id: str, reporting_pdf_path: str,
                         prior_pdf_path: Optional[str],
                         wires_data: list, year: int, account_id_hint: str):
    """Run in a background thread (called via asyncio executor)."""
    job = _jobs[job_id]

    def update(phase: int, pct: int, message: str):
        job.update({"phase": phase, "pct": pct, "message": message})

    try:
        update(1, 5, f"Parsing {year} PDF statement…")

        from espp2.transactions import normalize
        from espp2.main import (
            do_taxes, generate_previous_year_holdings,
            merge_transactions, ESPPErrorException,
        )
        from espp2.datamodels import Wires, Holdings
        from espp2.skatterapport import build_xlsx_with_skatterapport
        import simplejson as json
        from decimal import Decimal

        broker = "morgan"

        # Phase 1: Parse reporting year PDF
        reporting_transactions = normalize(reporting_pdf_path, broker)
        account_id = getattr(reporting_transactions, "__dict__", {}).get("account_id") or account_id_hint

        update(1, 20, f"Parsed {len(reporting_transactions.transactions)} transactions from {year} PDF")

        # Phase 2: Parse prior year PDF (if provided) to get opening balance
        update(2, 30, "Computing opening balance from prior year PDF…")

        holdfile = None
        if prior_pdf_path:
            prior_transactions = normalize(prior_pdf_path, broker)
            # Determine prior year
            prior_year = year - 1
            prior_years = {}
            for t in prior_transactions.transactions:
                prior_years[t.date.year] = 0

            holdings = generate_previous_year_holdings(
                broker=broker,
                years=sorted(prior_years.keys()),
                year=year,
                prev_holdings=None,
                transactions=prior_transactions,
                portfolio_engine=True,
            )
            update(2, 45, f"Opening balance computed (prior year holdings: {len(holdings.stocks)} positions)")
            # Write holdings to a temp JSON file
            import tempfile, io
            holdfile_path = tempfile.mktemp(suffix=".json")
            with open(holdfile_path, "w", encoding="utf-8") as hf:
                hf.write(holdings.model_dump_json(indent=2))
            holdfile = open(holdfile_path, "r", encoding="utf-8")

        # Phase 3: Convert wire data → Wires object
        update(3, 55, "Processing wire transfers…")

        wires_list = []
        for w in wires_data:
            try:
                wires_list.append({
                    "date": w["date"],
                    "currency": "USD",
                    "value": str(w.get("usd", 0)),
                    "nok_value": str(w.get("nok", 0)),
                })
            except Exception as e:
                logger.warning("Skipping wire entry %s: %s", w, e)

        wirefile = None
        if wires_list:
            import tempfile
            wire_tmp = tempfile.mktemp(suffix=".json")
            with open(wire_tmp, "w", encoding="utf-8") as wf:
                json.dump(wires_list, wf)
            wirefile = open(wire_tmp, "r", encoding="utf-8")

        update(3, 65, f"Running tax calculations for {year}…")

        # Phase 4: Run do_taxes
        result = do_taxes(
            broker=broker,
            transaction_files=[reporting_pdf_path],
            holdfile=holdfile,
            wirefile=wirefile,
            year=year,
            portfolio_engine=True,
        )

        update(4, 80, "Generating xlsx report…")

        # Build xlsx with Skattemeldingen sheet
        xlsx_bytes = build_xlsx_with_skatterapport(result.excel, result, year, account_id=account_id or "")

        update(4, 95, "Finalizing…")

        # Serialize summary for API response
        summary = result.summary
        report = result.report

        foreignshares = []
        for fs in summary.foreignshares:
            eoy_items = report.eoy_balance.get(year, [])
            symbol_items = [item for item in eoy_items if item.symbol == fs.symbol]
            wealth_usd = float(sum(item.amount.value for item in symbol_items))
            wealth_nok = int(round(float(fs.wealth)))
            exchange_rate = round(wealth_nok / wealth_usd, 4) if wealth_usd else 1.0
            gain_val = int(round(float(fs.taxable_gain)))

            cd = next((c for c in summary.credit_deduction if c.symbol == fs.symbol), None)

            foreignshares.append({
                "symbol": fs.symbol,
                "isin": fs.isin,
                "country": fs.country,
                "shares": float(round(float(fs.shares), 4)),
                "wealth_usd": round(wealth_usd, 2),
                "wealth_nok": wealth_nok,
                "exchange_rate": exchange_rate,
                "dividend_nok": int(round(float(fs.dividend))),
                "taxable_gain_nok": gain_val,
                "tax_deduction_used_nok": int(round(float(fs.tax_deduction_used))),
                "credit_deduction": {
                    "country": cd.country,
                    "income_tax_nok": int(round(float(cd.income_tax))),
                    "gross_dividend_nok": int(round(float(cd.gross_share_dividend))),
                    "tax_on_gross_nok": int(round(float(cd.tax_on_gross_share_dividend))),
                } if cd else None,
            })

        remaining = summary.cashsummary.remaining_cash
        # gain_aggregated = FX gain already included in share taxable_gain (Type A)
        # gain            = separate FX gain to report under Finans → Valuta (Type B)
        fx_gain_aggregated = int(round(float(summary.cashsummary.gain_aggregated)))
        fx_gain_separate   = int(round(float(summary.cashsummary.gain)))

        response_data = {
            "year": year,
            "account_id": account_id or "",
            "broker": "Morgan Stanley",
            "foreignshares": foreignshares,
            "fx_gain_aggregated_nok": fx_gain_aggregated,
            "fx_gain_nok": fx_gain_separate,
            "usd_balance": round(float(remaining.value), 2),
            "usd_balance_nok": int(round(float(remaining.nok_value))) if remaining.nok_value else 0,
        }

        job.update({
            "status": "done",
            "phase": 4,
            "pct": 100,
            "message": "Calculation complete",
            "result": response_data,
            "xlsx_bytes": xlsx_bytes,
        })

    except Exception as e:
        tb = traceback.format_exc()
        logger.error("Job %s failed: %s\n%s", job_id, e, tb)
        job.update({
            "status": "error",
            "phase": 0,
            "pct": 0,
            "message": str(e),
        })


@app.post("/api/process")
async def process(
    reporting_year_pdf: UploadFile = File(...),
    prior_year_pdf: Optional[UploadFile] = File(None),
    wires: str = Form("[]"),
    year: Optional[int] = Form(None),
):
    """Upload PDFs + wire data, start tax calculation, return job_id."""
    settings = _load_settings()
    if year is None:
        year = settings["year"]

    # Save uploaded files to temp dir
    tmpdir = tempfile.mkdtemp(prefix="taxledger_")

    reporting_path = os.path.join(tmpdir, f"reporting_{year}.pdf")
    with open(reporting_path, "wb") as f:
        f.write(await reporting_year_pdf.read())

    prior_path = None
    if prior_year_pdf and prior_year_pdf.filename:
        content = await prior_year_pdf.read()
        if content:
            prior_path = os.path.join(tmpdir, f"prior_{year - 1}.pdf")
            with open(prior_path, "wb") as f:
                f.write(content)

    try:
        wires_data = json.loads(wires)
    except Exception:
        wires_data = []

    job_id = str(uuid.uuid4())
    _jobs[job_id] = {
        "status": "running",
        "phase": 0,
        "pct": 0,
        "message": "Starting…",
        "result": None,
        "xlsx_bytes": None,
    }

    # Run in background thread so it doesn't block the event loop
    import asyncio
    loop = asyncio.get_event_loop()
    loop.run_in_executor(
        None,
        _run_tax_calculation,
        job_id, reporting_path, prior_path, wires_data, year, ""
    )

    return {"job_id": job_id}


@app.get("/api/job/{job_id}")
async def get_job_status(job_id: str):
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    return {
        "status": job["status"],
        "phase": job["phase"],
        "pct": job["pct"],
        "message": job["message"],
    }


@app.get("/api/results/{job_id}")
async def get_results(job_id: str):
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    if job["status"] != "done":
        raise HTTPException(status_code=202, detail="Job not complete")
    return job["result"]


@app.get("/api/download/{job_id}")
async def download_xlsx(job_id: str):
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    if job["status"] != "done" or not job.get("xlsx_bytes"):
        raise HTTPException(status_code=202, detail="Report not ready")
    year = job["result"].get("year", "")
    filename = f"Skattemeldingen_{year}.xlsx"
    return Response(
        content=job["xlsx_bytes"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def start():
    """Start the TaxLedger Pro web server."""
    import webbrowser
    import threading

    settings = _load_settings()

    # Auto-fetch exchange rates from Norges Bank on startup
    from espp2.fmv import FMV
    if not FMV._local_rates:
        try:
            year = settings.get("year", date.today().year - 1)
            n = FMV.fetch_norges_bank_rates(year)
            logger.info("Fetched %d exchange rate entries from Norges Bank at startup (year=%d)", n, year)
        except Exception as e:
            logger.warning("Could not fetch exchange rates at startup: %s", e)

    host = "127.0.0.1"
    port = 8000
    url = f"http://{host}:{port}"

    print(f"\n  TaxLedger Pro — starting at {url}\n")

    # Open browser after a short delay
    def open_browser():
        import time
        time.sleep(1.2)
        webbrowser.open(url)

    threading.Thread(target=open_browser, daemon=True).start()
    uvicorn.run(app, host=host, port=port, log_level="warning")


if __name__ == "__main__":
    start()
