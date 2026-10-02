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
  .log { background: #10141c; color: #dce4f2; border-radius: 12px; padding: 16px; min-height: 240px; max-height: 420px; overflow: auto; font: 12px/1.55 Consolas, monospace; white-space: pre-wrap; }
  .hidden { display: none; }
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
</script>
</body>
</html>
"""


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000)
