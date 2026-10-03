from __future__ import annotations

import os
import re
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from pydantic import BaseModel

import test_brave_payment_working_normal_items as automation
import db as _db


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploaded_excel"
UPLOAD_DIR.mkdir(exist_ok=True)
AUTOMATION_SCRIPT = BASE_DIR / "test_brave_payment_working_normal_items.py"

app = FastAPI(title="CRM Excel Order Upload")

STATE: dict[str, Any] = {
    "uploaded_path": None,
    "filename": None,
    "preview": None,
    "summary": None,
    "process": None,
    "log": [],
    "started_at": None,
    "finished_at": None,
    "return_code": None,
    "error": None,
    "report_path": None,
}
STATE_LOCK = threading.Lock()


class PreviewLine(BaseModel):
    document_no: str
    customer_no: str
    item_no: str
    description: str
    quantity: Any
    unit_price: Any


def _safe_filename(filename: str) -> str:
    name = Path(filename).name
    name = re.sub(r"[^A-Za-z0-9._ -]", "_", name)
    if not name.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Please upload an Excel .xlsx or .xlsm file.")
    return name


def _append_log(message: str) -> None:
    with STATE_LOCK:
        STATE["log"].append(message.rstrip())
        STATE["log"] = STATE["log"][-500:]


def _run_automation(excel_path: Path) -> None:
    env = os.environ.copy()
    env["CRM_EXCEL_FILE"] = str(excel_path)
    env["PYTHONUNBUFFERED"] = "1"
    env["CRM_UPLOAD_MODE"] = "1"

    try:
        process = subprocess.Popen(
            [sys.executable, "-u", str(AUTOMATION_SCRIPT)],
            cwd=str(BASE_DIR),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )

        with STATE_LOCK:
            STATE["process"] = process
            STATE["started_at"] = time.time()
            STATE["finished_at"] = None
            STATE["return_code"] = None
            STATE["error"] = None

        assert process.stdout is not None
        for line in process.stdout:
            _append_log(line)

        return_code = process.wait()
        with STATE_LOCK:
            STATE["return_code"] = return_code
            STATE["finished_at"] = time.time()

    except Exception as exc:
        _append_log(f"PAGE ERROR: {type(exc).__name__}: {exc}")
        with STATE_LOCK:
            STATE["error"] = str(exc)
            STATE["finished_at"] = time.time()
            STATE["return_code"] = -1
    finally:
        with STATE_LOCK:
            STATE["process"] = None


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return HTMLResponse(PAGE_HTML)


@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)) -> JSONResponse:
    filename = _safe_filename(file.filename or "uploaded.xlsx")
    destination = UPLOAD_DIR / filename

    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")

    destination.write_bytes(data)

    try:
        orders = automation.load_excel_orders(destination)
    except Exception as exc:
        destination.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if not orders:
        raise HTTPException(status_code=400, detail="No usable order data was found in the Excel file.")

    preview: list[dict[str, Any]] = []
    total_lines = 0
    for order in orders.values():
        total_lines += len(order["lines"])
        for line in order["lines"]:
            if len(preview) >= 100:
                break
            preview.append(
                {
                    "document_no": order["document_no"],
                    "customer_no": order["customer_no"],
                    "item_no": line["item_no"],
                    "description": line["description"],
                    "quantity": line["quantity"],
                    "unit_price": line["unit_price_excel"],
                }
            )
        if len(preview) >= 100:
            break

    summary = {
        "documents": len(orders),
        "lines": total_lines,
        "previewed_lines": len(preview),
    }

    with STATE_LOCK:
        STATE["uploaded_path"] = str(destination)
        STATE["filename"] = filename
        STATE["preview"] = preview
        STATE["summary"] = summary
        STATE["log"] = [
            f"Excel validated successfully: {filename}",
            f"Usable documents: {len(orders)}",
            f"Usable item lines: {total_lines}",
        ]
        STATE["error"] = None

    return JSONResponse(
        {
            "ok": True,
            "filename": filename,
            "summary": summary,
            "preview": preview,
        }
    )


@app.post("/start")
def start_automation() -> JSONResponse:
    with STATE_LOCK:
        uploaded_path = STATE["uploaded_path"]
        process = STATE["process"]

    if process is not None and process.poll() is None:
        raise HTTPException(status_code=409, detail="CRM automation is already running.")

    if not uploaded_path:
        raise HTTPException(status_code=400, detail="Upload and validate an Excel file first.")

    excel_path = Path(uploaded_path)
    if not excel_path.exists():
        raise HTTPException(status_code=400, detail="The uploaded Excel file is no longer available.")

    with STATE_LOCK:
        STATE["log"] = ["Starting existing CRM automation...", f"Excel source: {excel_path.name}"]
        STATE["return_code"] = None
        STATE["error"] = None

    thread = threading.Thread(target=_run_automation, args=(excel_path,), daemon=True)
    thread.start()

    return JSONResponse({"ok": True, "message": "CRM automation started."})


@app.get("/status")
def status() -> JSONResponse:
    with STATE_LOCK:
        process = STATE["process"]
        running = process is not None and process.poll() is None
        return_code = STATE["return_code"]
        return JSONResponse(
            {
                "running": running,
                "filename": STATE["filename"],
                "summary": STATE["summary"],
                "return_code": return_code,
                "error": STATE["error"],
                "log": STATE["log"][-120:],
            }
        )
def _latest_posting_report():
    downloads = Path.home() / "Downloads"

    reports = list(
        downloads.glob("CRM Posting Report - *.xlsx")
    )

    if not reports:
        return None

    return max(
        reports,
        key=lambda path: path.stat().st_mtime
    )


@app.get("/download-report")
def download_report():
    report_path = _latest_posting_report()

    if report_path is None:
        raise HTTPException(
            status_code=404,
            detail="No CRM posting report is available yet."
        )

    return FileResponse(
        path=report_path,
        filename=report_path.name,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

@app.get("/history")
def history() -> JSONResponse:
    """Return the most recent 100 runs."""
    runs = _db.get_all_runs(limit=100)
    return JSONResponse({"runs": runs})


@app.get("/history/{run_id}")
def run_detail(run_id: int) -> JSONResponse:
    """Return full detail for a single run — metadata + posting rows + summary."""
    run = _db.get_run(run_id)

    if run is None:
        raise HTTPException(status_code=404, detail=f"Run {run_id} not found.")

    rows = _db.get_posting_report_rows(run_id)
    summary = _db.get_posting_report_summary(run_id)
    excel_lines = _db.get_excel_order_lines(run_id)

    return JSONResponse({
        "run": run,
        "summary": summary,
        "posting_rows": rows,
        "excel_lines": excel_lines,
    })


@app.get("/history/{run_id}/export")
def export_run_report(run_id: int):
    """Download the Excel posting report file for a specific run."""
    run = _db.get_run(run_id)

    if run is None:
        raise HTTPException(status_code=404, detail=f"Run {run_id} not found.")

    report_path = run.get("report_path")

    if not report_path or not Path(report_path).exists():
        raise HTTPException(
            status_code=404,
            detail="No report file found for this run.",
        )

    return FileResponse(
        path=report_path,
        filename=Path(report_path).name,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )


# ============================================================
# ODATA ENDPOINTS — Power BI Web connector compatible
# ============================================================
# Power BI can consume these via:
#   Home -> Get Data -> Web -> http://127.0.0.1:8000/odata/runs
# Each endpoint returns { "value": [...] } which Power BI
# recognises as an OData collection automatically.
# ============================================================

@app.get("/odata/runs")
def odata_runs() -> JSONResponse:
    """All automation runs — one row per CRM session."""
    runs = _db.get_all_runs(limit=10000)
    return JSONResponse({"@odata.context": "runs", "value": runs})


@app.get("/odata/posting_rows")
def odata_posting_rows(run_id: int | None = None) -> JSONResponse:
    """
    All posting report rows across every run.
    Optional ?run_id=N filter to limit to a single run.
    """
    with _db.get_connection() as conn:
        if run_id is not None:
            rows = conn.execute(
                "SELECT * FROM posting_report_rows WHERE run_id = ? ORDER BY id",
                (run_id,),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM posting_report_rows ORDER BY id"
            ).fetchall()
    return JSONResponse({"@odata.context": "posting_rows", "value": [dict(r) for r in rows]})


@app.get("/odata/excel_lines")
def odata_excel_lines(run_id: int | None = None) -> JSONResponse:
    """
    All Excel order lines across every run.
    Optional ?run_id=N filter to limit to a single run.
    """
    with _db.get_connection() as conn:
        if run_id is not None:
            rows = conn.execute(
                "SELECT * FROM excel_order_lines WHERE run_id = ? ORDER BY id",
                (run_id,),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM excel_order_lines ORDER BY id"
            ).fetchall()
    return JSONResponse({"@odata.context": "excel_lines", "value": [dict(r) for r in rows]})


@app.get("/odata")
def odata_service_document() -> JSONResponse:
    """OData service document — lists available entity sets."""
    base = "http://127.0.0.1:8000/odata"
    return JSONResponse({
        "@odata.context": f"{base}/$metadata",
        "value": [
            {"name": "runs",          "url": f"{base}/runs"},
            {"name": "posting_rows",  "url": f"{base}/posting_rows"},
            {"name": "excel_lines",   "url": f"{base}/excel_lines"},
        ],
    })


PAGE_HTML = r"""
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>CRM Excel Order Upload</title>
<style>
  :root { font-family: Inter, Segoe UI, Arial, sans-serif; color: #172033; background: #f4f7fb; }
  * { box-sizing: border-box; }
  body { margin: 0; min-height: 100vh; }
  .wrap { max-width: 1120px; margin: 0 auto; padding: 42px 24px 60px; }
  .hero { margin-bottom: 24px; }
  .eyebrow { font-size: 13px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; opacity: .65; }
  h1 { margin: 8px 0; font-size: 34px; }
  .sub { margin: 0; color: #5b6577; }
  .card { background: white; border: 1px solid #e1e7f0; border-radius: 18px; padding: 24px; margin-top: 20px; box-shadow: 0 8px 28px rgba(24, 39, 75, .06); }
  .drop { border: 2px dashed #9eabc0; border-radius: 16px; padding: 46px 24px; text-align: center; cursor: pointer; background: #fbfcfe; transition: .15s ease; }
  .drop.drag { border-color: #5a6f99; background: #f1f5fb; }
  .drop strong { display: block; font-size: 18px; margin-bottom: 8px; }
  .drop span { color: #6a7383; }
  input[type=file] { display: none; }
  .actions { display: flex; gap: 12px; flex-wrap: wrap; margin-top: 18px; }
  button { border: 0; border-radius: 10px; padding: 12px 18px; font-weight: 700; cursor: pointer; }
  #browse, #start { background: #172033; color: white; }
  button:disabled { opacity: .45; cursor: not-allowed; }
  .status { margin-top: 16px; min-height: 24px; color: #4c5668; white-space: pre-wrap; }
  .stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin: 18px 0; }
  .stat { background: #f7f9fc; border: 1px solid #e4e8ef; border-radius: 12px; padding: 16px; }
  .stat b { display: block; font-size: 24px; margin-top: 4px; }
  .table-wrap { overflow: auto; border: 1px solid #e3e7ee; border-radius: 12px; }
  table { width: 100%; border-collapse: collapse; min-width: 820px; }
  th, td { padding: 11px 12px; border-bottom: 1px solid #edf0f4; text-align: left; font-size: 13px; }
  th { background: #f8f9fb; position: sticky; top: 0; }
  tr:last-child td { border-bottom: 0; }
  .pill { display: inline-block; padding: 5px 9px; border-radius: 999px; background: #edf2f8; font-size: 12px; }
  .pill-success { background: #e8f5e9; color: #2e7d32; }
  .pill-error   { background: #fce4ec; color: #c00000; }
  .pill-warn    { background: #fff8e1; color: #f57f17; }
  .pill-running { background: #e3f2fd; color: #1565c0; }
  .log { background: #10141c; color: #dce4f2; border-radius: 12px; padding: 16px; min-height: 240px; max-height: 420px; overflow: auto; font: 12px/1.55 Consolas, monospace; white-space: pre-wrap; }
  .hidden { display: none; }
  .tabs { display: flex; gap: 8px; margin-bottom: 4px; }
  .tab { background: #e8edf5; border: 0; border-radius: 10px 10px 0 0; padding: 10px 22px; font-size: 14px; font-weight: 700; cursor: pointer; color: #4c5668; }
  .tab-active { background: white; color: #172033; border: 1px solid #e1e7f0; border-bottom: 1px solid white; margin-bottom: -1px; }
  .run-row { border: 1px solid #e3e7ee; border-radius: 12px; padding: 14px 16px; margin-bottom: 10px; cursor: pointer; transition: .12s; }
  .run-row:hover { background: #f4f7fb; border-color: #b0bcce; }
  #detailPanel { margin-top: 16px; }
  @media (max-width: 760px) { .stats { grid-template-columns: 1fr; } h1 { font-size: 28px; } }
</style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <div class="eyebrow">CRM Order Automation</div>
    <h1>Upload Excel Order File</h1>
    <p class="sub">Drop the Excel export here, validate the usable order lines, then start the existing CRM posting workflow.</p>
  </div>

  <div class="tabs">
    <button class="tab tab-active" id="uploadTab">Upload &amp; Post</button>
    <button class="tab" id="historyTab">Run History</button>
    <button class="tab" id="pbiTab">&#128200; Power BI</button>
  </div>

  <div id="uploadPanel">

  <div class="card">
    <label id="drop" class="drop" for="file">
      <strong>Drag & drop your Excel file here</strong>
      <span>or click to browse — .xlsx / .xlsm</span>
    </label>
    <input id="file" type="file" accept=".xlsx,.xlsm">
    <div class="actions">
  <button id="browse" type="button">Choose Excel File</button>
  <button id="start" type="button" disabled>Start CRM Posting</button>
  <button id="downloadReport" type="button" style="display:none;">
    Download Posting Report
  </button>
</div>
    <div id="status" class="status">No file loaded.</div>
  </div>

  <div id="previewCard" class="card hidden">
    <h2>Validated Excel Data</h2>
    <div class="stats">
      <div class="stat">Documents<b id="documents">0</b></div>
      <div class="stat">Usable item lines<b id="lines">0</b></div>
      <div class="stat">Preview rows<b id="previewed">0</b></div>
    </div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>Document</th><th>Customer</th><th>Item</th><th>Description</th><th>Qty</th><th>Excel Unit Price</th></tr></thead>
        <tbody id="rows"></tbody>
      </table>
    </div>
  </div>

  <div class="card">
    <h2>CRM Automation Log</h2>
    <div id="log" class="log">Waiting for the CRM automation to start...</div>
  </div>
</div>

<div id="historyPanel" class="hidden">
  <div class="card">
    <h2>Run History</h2>
    <div id="historyList"><p style="color:#6a7383">Click the History tab to load.</p></div>
    <div id="detailPanel" class="hidden"></div>
  </div>
</div>

<div id="pbiPanel" class="hidden">
  <div class="card">
    <h2>&#128200; Power BI Integration</h2>
    <p style="color:#5b6577;margin-top:0">Connect Power BI Desktop directly to the live database using the endpoints below.</p>

    <div class="stats" style="margin-bottom:0">
      <div class="stat">
        Runs
        <b style="font-size:14px;word-break:break-all">
          <a href="/odata/runs" target="_blank" style="color:#172033">/odata/runs</a>
        </b>
        <span style="font-size:12px;color:#6a7383">One row per automation session</span>
      </div>
      <div class="stat">
        Posting Rows
        <b style="font-size:14px;word-break:break-all">
          <a href="/odata/posting_rows" target="_blank" style="color:#172033">/odata/posting_rows</a>
        </b>
        <span style="font-size:12px;color:#6a7383">Every CRM posting line with inventory status</span>
      </div>
      <div class="stat">
        Excel Lines
        <b style="font-size:14px;word-break:break-all">
          <a href="/odata/excel_lines" target="_blank" style="color:#172033">/odata/excel_lines</a>
        </b>
        <span style="font-size:12px;color:#6a7383">Every Excel order line loaded per run</span>
      </div>
    </div>
  </div>

  <div class="card">
    <h3 style="margin-top:0">Quick Connect Steps</h3>
    <ol style="color:#4c5668;line-height:2">
      <li>Open <b>Power BI Desktop</b></li>
      <li>Click <b>Home → Transform Data</b> to open Power Query Editor</li>
      <li>Click <b>Home → New Source → Blank Query</b></li>
      <li>Click <b>View → Advanced Editor</b></li>
      <li>Paste one of the M queries below, click <b>Done</b></li>
      <li>Rename the query in the left panel (<code>Runs</code>, <code>PostingRows</code>, <code>ExcelLines</code>)</li>
      <li>Repeat for each table, then click <b>Close &amp; Apply</b></li>
      <li>In <b>Model view</b> link <code>PostingRows[run_id]</code> → <code>Runs[id]</code> and <code>ExcelLines[run_id]</code> → <code>Runs[id]</code></li>
    </ol>
  </div>

  <div class="card">
    <h3 style="margin-top:0">M Query — Runs</h3>
    <div style="position:relative">
      <button onclick="copyQuery('qRuns')" style="position:absolute;top:8px;right:8px;background:#172033;color:#fff;padding:6px 12px;font-size:12px;border-radius:6px">Copy</button>
      <pre id="qRuns" class="log" style="min-height:auto;max-height:220px;margin:0;font-size:12px">let
    Source     = Json.Document(Web.Contents("http://127.0.0.1:8000/odata/runs")),
    Value      = Source[value],
    ToTable    = Table.FromList(Value, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Expanded   = Table.ExpandRecordColumn(ToTable, "Column1",
                     {"id","started_at","finished_at","excel_filename","status",
                      "error_message","total_documents","total_lines",
                      "report_path","customer_name","created_at"},
                     {"id","started_at","finished_at","excel_filename","status",
                      "error_message","total_documents","total_lines",
                      "report_path","customer_name","created_at"}),
    TypedTable = Table.TransformColumnTypes(Expanded,{
                     {"id", Int64.Type}, {"started_at", type datetime},
                     {"finished_at", type datetime}, {"excel_filename", type text},
                     {"status", type text}, {"error_message", type text},
                     {"total_documents", Int64.Type}, {"total_lines", Int64.Type},
                     {"report_path", type text}, {"customer_name", type text},
                     {"created_at", type datetime}})
in TypedTable</pre>
    </div>
  </div>

  <div class="card">
    <h3 style="margin-top:0">M Query — PostingRows</h3>
    <div style="position:relative">
      <button onclick="copyQuery('qPosting')" style="position:absolute;top:8px;right:8px;background:#172033;color:#fff;padding:6px 12px;font-size:12px;border-radius:6px">Copy</button>
      <pre id="qPosting" class="log" style="min-height:auto;max-height:220px;margin:0;font-size:12px">let
    Source     = Json.Document(Web.Contents("http://127.0.0.1:8000/odata/posting_rows")),
    Value      = Source[value],
    ToTable    = Table.FromList(Value, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Expanded   = Table.ExpandRecordColumn(ToTable, "Column1",
                     {"id","run_id","posting_date","customer_no","document_no",
                      "excel_item_no","excel_description","excel_qty","excel_unit_price",
                      "crm_item_no","crm_product","crm_inventory","allocated_qty",
                      "crm_unit_price","inventory_status"},
                     {"id","run_id","posting_date","customer_no","document_no",
                      "excel_item_no","excel_description","excel_qty","excel_unit_price",
                      "crm_item_no","crm_product","crm_inventory","allocated_qty",
                      "crm_unit_price","inventory_status"}),
    TypedTable = Table.TransformColumnTypes(Expanded,{
                     {"id", Int64.Type}, {"run_id", Int64.Type},
                     {"posting_date", type date}, {"customer_no", type text},
                     {"document_no", type text}, {"excel_item_no", type text},
                     {"excel_description", type text}, {"excel_qty", type number},
                     {"excel_unit_price", type number}, {"crm_item_no", type text},
                     {"crm_product", type text}, {"crm_inventory", type number},
                     {"allocated_qty", type number}, {"crm_unit_price", type number},
                     {"inventory_status", type text}})
in TypedTable</pre>
    </div>
  </div>

  <div class="card">
    <h3 style="margin-top:0">M Query — ExcelLines</h3>
    <div style="position:relative">
      <button onclick="copyQuery('qExcel')" style="position:absolute;top:8px;right:8px;background:#172033;color:#fff;padding:6px 12px;font-size:12px;border-radius:6px">Copy</button>
      <pre id="qExcel" class="log" style="min-height:auto;max-height:220px;margin:0;font-size:12px">let
    Source     = Json.Document(Web.Contents("http://127.0.0.1:8000/odata/excel_lines")),
    Value      = Source[value],
    ToTable    = Table.FromList(Value, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Expanded   = Table.ExpandRecordColumn(ToTable, "Column1",
                     {"id","run_id","document_no","customer_no","item_no",
                      "description","quantity","unit_price","amount",
                      "unit_of_measure","location_code","crm_search_text"},
                     {"id","run_id","document_no","customer_no","item_no",
                      "description","quantity","unit_price","amount",
                      "unit_of_measure","location_code","crm_search_text"}),
    TypedTable = Table.TransformColumnTypes(Expanded,{
                     {"id", Int64.Type}, {"run_id", Int64.Type},
                     {"document_no", type text}, {"customer_no", type text},
                     {"item_no", type text}, {"description", type text},
                     {"quantity", type number}, {"unit_price", type number},
                     {"amount", type number}, {"unit_of_measure", type text},
                     {"location_code", type text}, {"crm_search_text", type text}})
in TypedTable</pre>
    </div>
  </div>

  <div class="card">
    <h3 style="margin-top:0">Suggested DAX Measures</h3>
    <div style="position:relative">
      <button onclick="copyQuery('qDax')" style="position:absolute;top:8px;right:8px;background:#172033;color:#fff;padding:6px 12px;font-size:12px;border-radius:6px">Copy</button>
      <pre id="qDax" class="log" style="min-height:auto;max-height:260px;margin:0;font-size:12px">-- Paste each measure into: Home -> New Measure (on PostingRows table)

Total Lines = COUNTROWS(PostingRows)

Enough Inventory =
    CALCULATE(COUNTROWS(PostingRows),
              PostingRows[inventory_status] = "ENOUGH INVENTORY")

Insufficient Inventory =
    CALCULATE(COUNTROWS(PostingRows),
              PostingRows[inventory_status] = "INSUFFICIENT INVENTORY")

Zero Inventory =
    CALCULATE(COUNTROWS(PostingRows),
              PostingRows[inventory_status] = "ZERO INVENTORY")

Fulfilment Rate % =
    DIVIDE([Enough Inventory], [Total Lines], 0) * 100

Total Allocated Qty = SUM(PostingRows[allocated_qty])

Total Excel Qty = SUM(PostingRows[excel_qty])

Shortfall Qty = [Total Excel Qty] - [Total Allocated Qty]</pre>
    </div>
  </div>
</div>

</div>

<script>
const fileInput = document.getElementById('file');
const drop = document.getElementById('drop');
const browse = document.getElementById('browse');
const start = document.getElementById('start');
const downloadReport = document.getElementById('downloadReport');
const status = document.getElementById('status');
const previewCard = document.getElementById('previewCard');
const rows = document.getElementById('rows');
const log = document.getElementById('log');

browse.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', () => {
  if (fileInput.files.length) uploadFile(fileInput.files[0]);
});

['dragenter','dragover'].forEach(evt => drop.addEventListener(evt, e => {
  e.preventDefault(); drop.classList.add('drag');
}));
['dragleave','drop'].forEach(evt => drop.addEventListener(evt, e => {
  e.preventDefault(); drop.classList.remove('drag');
}));
drop.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (file) uploadFile(file);
});

async function uploadFile(file) {
  const form = new FormData();
  form.append('file', file);
  status.textContent = `Uploading and validating ${file.name}...`;
  start.disabled = true;
  try {
    const response = await fetch('/upload', { method: 'POST', body: form });
    const data = await response.json();
    if (!response.ok) throw new Error(data.detail || 'Upload failed.');
    renderPreview(data);
    status.textContent = `Validated: ${data.filename}`;
    start.disabled = false;
  } catch (err) {
    previewCard.classList.add('hidden');
    status.textContent = `ERROR: ${err.message}`;
  }
}

function renderPreview(data) {
  previewCard.classList.remove('hidden');
  document.getElementById('documents').textContent = data.summary.documents;
  document.getElementById('lines').textContent = data.summary.lines;
  document.getElementById('previewed').textContent = data.summary.previewed_lines;
  rows.innerHTML = data.preview.map(r => `
    <tr>
      <td>${esc(r.document_no)}</td>
      <td>${esc(r.customer_no)}</td>
      <td><span class="pill">${esc(r.item_no)}</span></td>
      <td>${esc(r.description)}</td>
      <td>${esc(r.quantity)}</td>
      <td>${esc(r.unit_price)}</td>
    </tr>`).join('');
}

start.addEventListener('click', async () => {
  start.disabled = true;
  status.textContent = 'Starting the existing CRM automation...';
  try {
    const response = await fetch('/start', { method: 'POST' });
    const data = await response.json();
    if (!response.ok) throw new Error(data.detail || 'Could not start automation.');
    status.textContent = 'CRM automation started. Complete manual CRM login/MFA if prompted.';
  } catch (err) {
    status.textContent = `ERROR: ${err.message}`;
    start.disabled = false;
  }
});

function esc(value) {
  return String(value ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
}

async function pollStatus() {
  try {
    const response = await fetch('/status');
    const data = await response.json();
    if (data.log && data.log.length) {
      log.textContent = data.log.join('\n');
      log.scrollTop = log.scrollHeight;
    }
    if (data.running) {
      start.disabled = true;
      status.textContent = 'CRM automation is running...';
    } else if (data.return_code !== null) {
      start.disabled = false;

      if (data.return_code === 0) {
        status.textContent = 'CRM automation finished.';
        downloadReport.style.display = 'inline-block';
      } else {
        status.textContent =
          `CRM automation stopped with exit code ${data.return_code}.`;
        downloadReport.style.display = 'none';
      }
}
  } catch (_) {}
}
downloadReport.addEventListener('click', () => {
  window.location.href = `/download-report?t=${Date.now()}`;
});
setInterval(pollStatus, 1000);
pollStatus();

// ============================================================
// HISTORY TAB
// ============================================================

const historyTab   = document.getElementById('historyTab');
const uploadTab    = document.getElementById('uploadTab');
const pbiTab       = document.getElementById('pbiTab');
const historyPanel = document.getElementById('historyPanel');
const uploadPanel  = document.getElementById('uploadPanel');
const pbiPanel     = document.getElementById('pbiPanel');
const historyList  = document.getElementById('historyList');
const detailPanel  = document.getElementById('detailPanel');

pbiTab.addEventListener('click', () => {
  pbiTab.classList.add('tab-active');
  uploadTab.classList.remove('tab-active');
  historyTab.classList.remove('tab-active');
  pbiPanel.classList.remove('hidden');
  uploadPanel.classList.add('hidden');
  historyPanel.classList.add('hidden');
});

function copyQuery(id) {
  const text = document.getElementById(id).innerText;
  navigator.clipboard.writeText(text).then(() => {
    const btn = document.querySelector(`button[onclick="copyQuery('${id}')"]`);
    const orig = btn.textContent;
    btn.textContent = 'Copied!';
    setTimeout(() => { btn.textContent = orig; }, 1500);
  });
}

uploadTab.addEventListener('click', () => {
  uploadTab.classList.add('tab-active');
  historyTab.classList.remove('tab-active');
  pbiTab.classList.remove('tab-active');
  uploadPanel.classList.remove('hidden');
  historyPanel.classList.add('hidden');
  pbiPanel.classList.add('hidden');
});

historyTab.addEventListener('click', () => {
  historyTab.classList.add('tab-active');
  uploadTab.classList.remove('tab-active');
  pbiTab.classList.remove('tab-active');
  historyPanel.classList.remove('hidden');
  uploadPanel.classList.add('hidden');
  pbiPanel.classList.add('hidden');
  loadHistory();
});

async function loadHistory() {
  historyList.innerHTML = '<p style="color:#6a7383">Loading...</p>';
  detailPanel.classList.add('hidden');
  try {
    const res  = await fetch('/history');
    const data = await res.json();
    if (!data.runs.length) {
      historyList.innerHTML = '<p style="color:#6a7383">No runs recorded yet.</p>';
      return;
    }
    historyList.innerHTML = data.runs.map(r => `
      <div class="run-row" onclick="loadRunDetail(${r.id})">
        <div style="display:flex;justify-content:space-between;align-items:center">
          <span>
            <b>Run #${r.id}</b>
            <span class="pill pill-${r.status}">${r.status}</span>
          </span>
          <span style="font-size:12px;color:#6a7383">${r.started_at || ''}</span>
        </div>
        <div style="font-size:13px;margin-top:4px;color:#4c5668">
          ${esc(r.excel_filename || '—')} &nbsp;|&nbsp;
          ${r.total_documents ?? '?'} docs &nbsp;|&nbsp;
          ${r.total_lines ?? '?'} lines &nbsp;|&nbsp;
          Customer: ${esc(r.customer_name || '—')}
        </div>
        ${r.error_message ? `<div style="font-size:12px;color:#c00000;margin-top:2px">${esc(r.error_message)}</div>` : ''}
      </div>`).join('');
  } catch (err) {
    historyList.innerHTML = `<p style="color:#c00000">Failed to load history: ${err.message}</p>`;
  }
}

async function loadRunDetail(runId) {
  detailPanel.classList.remove('hidden');
  detailPanel.innerHTML = '<p style="color:#6a7383">Loading run detail...</p>';
  try {
    const res  = await fetch(`/history/${runId}`);
    const data = await res.json();
    const r    = data.run;
    const s    = data.summary;
    const rows = data.posting_rows;

    detailPanel.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
        <h3 style="margin:0">Run #${r.id} — ${esc(r.excel_filename || '—')}</h3>
        ${r.report_path ? `<button onclick="window.location='/history/${r.id}/export'" style="background:#172033;color:#fff;border:0;border-radius:8px;padding:8px 14px;font-weight:700;cursor:pointer">Download Report</button>` : ''}
      </div>
      <div class="stats" style="margin-bottom:16px">
        <div class="stat">Status<b><span class="pill pill-${r.status}">${r.status}</span></b></div>
        <div class="stat">Started<b>${r.started_at || '—'}</b></div>
        <div class="stat">Finished<b>${r.finished_at || '—'}</b></div>
        <div class="stat">Documents<b>${r.total_documents ?? '—'}</b></div>
        <div class="stat">Lines<b>${r.total_lines ?? '—'}</b></div>
        <div class="stat">Customer<b>${esc(r.customer_name || '—')}</b></div>
      </div>
      <div class="stats" style="margin-bottom:16px">
        <div class="stat" style="background:#e8f5e9">Enough Inventory<b style="color:#2e7d32">${s.enough ?? 0}</b></div>
        <div class="stat" style="background:#fff8e1">Insufficient<b style="color:#f57f17">${s.insufficient ?? 0}</b></div>
        <div class="stat" style="background:#fce4ec">Zero Inventory<b style="color:#c00000">${s.zero ?? 0}</b></div>
      </div>
      ${r.error_message ? `<div style="background:#fff0f0;border:1px solid #f5c6c6;border-radius:8px;padding:12px;margin-bottom:12px;color:#c00000;font-size:13px">${esc(r.error_message)}</div>` : ''}
      <div class="table-wrap">
        <table>
          <thead><tr>
            <th>Date</th><th>Customer</th><th>Document</th>
            <th>Excel Item</th><th>Description</th><th>Excel Qty</th>
            <th>CRM Item</th><th>CRM Product</th>
            <th>Inventory</th><th>Allocated</th><th>Status</th>
          </tr></thead>
          <tbody>
            ${rows.map(row => `<tr>
              <td>${esc(row.posting_date)}</td>
              <td>${esc(row.customer_no)}</td>
              <td>${esc(row.document_no)}</td>
              <td><span class="pill">${esc(row.excel_item_no)}</span></td>
              <td style="color:${row.inventory_status==='ZERO INVENTORY'||row.inventory_status==='INSUFFICIENT INVENTORY'?'#c00000':'inherit'}">${esc(row.excel_description)}</td>
              <td>${esc(row.excel_qty)}</td>
              <td><span class="pill">${esc(row.crm_item_no)}</span></td>
              <td>${esc(row.crm_product)}</td>
              <td>${esc(row.crm_inventory)}</td>
              <td>${esc(row.allocated_qty)}</td>
              <td><span class="pill pill-${row.inventory_status==='ENOUGH INVENTORY'?'success':row.inventory_status==='ZERO INVENTORY'?'error':'warn'}">${esc(row.inventory_status)}</span></td>
            </tr>`).join('')}
          </tbody>
        </table>
      </div>`;
  } catch (err) {
    detailPanel.innerHTML = `<p style="color:#c00000">Failed to load run detail: ${err.message}</p>`;
  }
}

</script>
</body>
</html>
"""


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000)
