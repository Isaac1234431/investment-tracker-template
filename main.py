from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Any
import base64
import io
import os
import httpx
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.formula import ArrayFormula
from datetime import datetime
from collections import defaultdict

app = FastAPI()

# ── CORS — allows the UI (any origin) to call this service directly ──────────
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Constants ────────────────────────────────────────────────────────────────
TEMPLATE_URL = os.environ.get("TEMPLATE_URL", "")  # Set in Render env vars
N8N_WEBHOOK_URL = os.environ.get("N8N_WEBHOOK_URL", "")  # Set in Render env vars
DATE_FMT     = "mmm/dd/yyyy"
BLUE_FILL    = PatternFill("solid", fgColor="1F4E79")
WHITE_BOLD   = Font(bold=True, color="FFFFFF")
WHITE_FONT   = Font(color="FFFFFF")
PEND_COLS    = ["C","D","E","F","G","H","I","J","K","L","M","N","O","P"]

BUY_VARIANTS  = {"buy", "purchase", "bought"}
SELL_VARIANTS = {"sell", "sale", "sold", "redemption"}

# ── Request model ─────────────────────────────────────────────────────────────
class BuildRequest(BaseModel):
    extract_result: dict[str, Any]

# ── Helpers ───────────────────────────────────────────────────────────────────
def normalize_trans_type(raw: str) -> str:
    lower = raw.lower().strip()
    if lower in BUY_VARIANTS:  return "Buy"
    if lower in SELL_VARIANTS: return "Sell"
    return raw  # preserve original for non-trade types

def build_symbol_lookup(holdings):
    return {h["description"].upper().strip(): h["symbol"]
            for h in holdings if h.get("symbol") and h.get("description")}

def resolve_symbol(description, symbol, lookup):
    if symbol: return symbol
    return lookup.get((description or "").upper().strip(), "")

def parse_date(s):
    if not s: return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
        try: return datetime.strptime(s, fmt)
        except: pass
    return s

def set_date(cell, value):
    cell.value = value
    if isinstance(value, datetime):
        cell.number_format = DATE_FMT

def clear_row(ws, row, cols):
    for col in cols:
        cell = ws[f"{col}{row}"]
        cell.value = None
        cell.fill = PatternFill(fill_type=None)
        cell.font = Font()

def apply_blue_row(ws, row, cols, label=None, label_col=None):
    for col in cols:
        cell = ws[f"{col}{row}"]
        cell.fill = BLUE_FILL
        cell.font = WHITE_BOLD if (label and col == label_col) else WHITE_FONT
    if label and label_col:
        ws[f"{label_col}{row}"].value = label

# ── Core builder ──────────────────────────────────────────────────────────────
def build_excel(extract: dict, template_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))

    holdings     = extract.get("holdings", [])
    transactions = extract.get("transactions", [])
    lookup       = build_symbol_lookup(holdings)

    # Normalize all trans_type values
    for t in transactions:
        t["trans_type"] = normalize_trans_type(t.get("trans_type", ""))

    confirmed = [t for t in transactions if not t.get("is_pending")]
    pending   = [t for t in transactions if t.get("is_pending")]

    # ── 1. Reference Template ────────────────────────────────────────────────
    ref = wb["Reference Template"]

    for r in range(5, 100):
        clear_row(ref, r, PEND_COLS)

    conf_start = 5
    for i, t in enumerate(confirmed):
        r = conf_start + i
        set_date(ref[f"D{r}"], parse_date(t.get("date")))
        ref[f"E{r}"].value = t.get("trans_type", "")
        ref[f"F{r}"].value = resolve_symbol(t.get("description"), t.get("symbol"), lookup)
        ref[f"G{r}"].value = t.get("description", "")
        ref[f"H{r}"].value = t.get("currency", "")
        ref[f"I{r}"].value = t.get("price_per_share")
        ref[f"J{r}"].value = t.get("shares")
        ref[f"K{r}"].value = t.get("amount")
        ref[f"L{r}"].value = t.get("charges") or None

    pend_title_row  = conf_start + len(confirmed) + 2
    pend_data_start = pend_title_row + 1
    apply_blue_row(ref, pend_title_row, PEND_COLS,
                   label="Pending Transactions", label_col="C")

    for i, t in enumerate(pending):
        r = pend_data_start + i
        ref[f"C{r}"].value = "Pending"
        set_date(ref[f"D{r}"], parse_date(t.get("date")))
        ref[f"E{r}"].value = t.get("trans_type", "")
        ref[f"F{r}"].value = resolve_symbol(t.get("description"), t.get("symbol"), lookup)
        ref[f"G{r}"].value = t.get("description", "")
        ref[f"H{r}"].value = t.get("currency", "")
        ref[f"I{r}"].value = t.get("price_per_share")
        ref[f"J{r}"].value = t.get("shares")
        set_date(ref[f"N{r}"], parse_date(t.get("settle_date")))
        ref[f"P{r}"].value = t.get("amount")

    # ── 2. Individual stock tabs ─────────────────────────────────────────────
    template_sheet = wb["Individual Stock Template"]

    tickers = sorted(set(
        resolve_symbol(t.get("description",""), t.get("symbol"), lookup)
        for t in transactions
        if t.get("trans_type","").lower() in ("buy","sell")
        and resolve_symbol(t.get("description",""), t.get("symbol"), lookup)
    ))

    for ticker in tickers:
        new_ws = wb.copy_worksheet(template_sheet)
        new_ws.title = ticker[:31]
        new_ws["D2"] = ticker
        # FILTER formula in C6 stays intact — auto-pulls from Reference Template
        # Update output summary to use dynamic last-row detection
        new_ws["R5"] = "=IFERROR(INDEX(L:L,MATCH(9.99E+307,IF(J6:J999<>\"\",ROW(J6:J999)),1)),0)"
        new_ws["R6"] = "=IFERROR(INDEX(M:M,MATCH(9.99E+307,IF(J6:J999<>\"\",ROW(J6:J999)),1)),0)"
        new_ws["R7"] = "=R5*R6"

    # Re-order sheets alphabetically
    fixed = ["Summary", "Reference Template", "Individual Stock Template"]
    stock_tabs = sorted([s for s in wb.sheetnames if s not in fixed])
    for i, name in enumerate(fixed + stock_tabs):
        wb.move_sheet(name, offset=wb.sheetnames.index(name) - i)

    # ── 3. Summary sheet ─────────────────────────────────────────────────────
    summary = wb["Summary"]
    for i, h in enumerate(sorted(holdings, key=lambda x: x.get("symbol",""))):
        r = 4 + i
        ticker = h.get("symbol") or lookup.get((h.get("description","")).upper().strip(),"")
        summary[f"C{r}"].value = h.get("description","")
        summary[f"D{r}"].value = ticker
        set_date(summary[f"E{r}"], parse_date(extract.get("statement_period_end")))
        book = (h.get("cost_per_unit") or 0) * (h.get("quantity") or 0)
        summary[f"G{r}"].value = book if h.get("cost_per_unit") else None
        summary[f"H{r}"].value = h.get("quantity")
        summary[f"I{r}"].value = h.get("cost_per_unit")
        summary[f"S{r}"].value = book if h.get("cost_per_unit") else None
        summary[f"T{r}"].value = h.get("quantity")
        summary[f"U{r}"].value = h.get("cost_per_unit")
        summary[f"W{r}"].value = h.get("market_price")
        summary[f"X{r}"].value = h.get("market_value")
        if h.get("market_value") is not None and h.get("cost_per_unit") is not None:
            summary[f"Y{r}"].value = f"=X{r}-G{r}"

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ── Routes ────────────────────────────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/trigger")
async def trigger_workflow(file: UploadFile = File(...)):
    """
    Receives a PDF from the UI and forwards it to the n8n webhook.
    Acts as a CORS-safe proxy — the UI never calls n8n directly.
    """
    if not N8N_WEBHOOK_URL:
        raise HTTPException(500, "N8N_WEBHOOK_URL environment variable not set")

    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Only PDF files are accepted")

    pdf_bytes = await file.read()

    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.post(
            N8N_WEBHOOK_URL,
            files={"Investment_Statement": (file.filename, pdf_bytes, "application/pdf")},
        )
        if resp.status_code not in (200, 202):
            raise HTTPException(502, f"n8n webhook returned {resp.status_code}")

    return JSONResponse({"status": "received", "filename": file.filename})


@app.post("/build-excel")
async def build_excel_endpoint(req: BuildRequest):
    # Fetch template from URL set in env var (upload template to GitHub or S3)
    if not TEMPLATE_URL:
        raise HTTPException(500, "TEMPLATE_URL environment variable not set")

    async with httpx.AsyncClient() as client:
        resp = await client.get(TEMPLATE_URL)
        if resp.status_code != 200:
            raise HTTPException(500, f"Failed to fetch template: {resp.status_code}")
        template_bytes = resp.content

    excel_bytes = build_excel(req.extract_result, template_bytes)
    b64 = base64.b64encode(excel_bytes).decode("utf-8")

    period = req.extract_result.get("statement_period_end", "output")
    broker = req.extract_result.get("broker_name", "Investment").replace(" ", "_")
    filename = f"{broker}_{period}.xlsx"

    return JSONResponse({
        "filename": filename,
        "data": b64,
        "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    })
