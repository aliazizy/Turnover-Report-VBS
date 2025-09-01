# main.py
from fastapi.responses import StreamingResponse, JSONResponse
import io, json
from typing import List, Dict, Any

from fastapi import FastAPI, UploadFile, File, HTTPException, Depends, Request
from fastapi.middleware.cors import CORSMiddleware

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import logging
import time

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(name)s - %(message)s"
)
logger = logging.getLogger("turnover-api")
app = FastAPI(
    title="Turnover Report Generator",
    version="1.0.0",
    description=(
        "Upload a turnover JSON and receive a styled Excel file.\n\n"
        "- A1: dateTimeUser (yellow, bold)\n"
        "- A2: caption (bold, larger)\n"
        "- A3: legend (italic)\n"
        "- Row 5: headers (blue fill, white bold, centered) + Excel Table\n"
        "- Formats: Turnover EUR = `#,##0.00`, Percent = `0.0%`"
    ),
    contact={"name": "VOSAIO Engineering"},
    license_info={"name": "MIT"},
    openapi_url="/openapi.json",   # optional: explicit
    docs_url="/docs",              # Swagger UI
    redoc_url="/redoc",            # ReDoc
)

@app.middleware("http")
async def log_requests(request: Request, call_next):
    idem = f"{time.time_ns()}"
    logger.info(f"rid={idem} start request path={request.url.path} method={request.method}")

    start_time = time.time()
    try:
        response = await call_next(request)
    except Exception as ex:
        logger.exception(f"rid={idem} exception: {ex}")
        raise

    process_time = (time.time() - start_time) * 1000
    logger.info(
        f"rid={idem} completed_in={process_time:.2f}ms "
        f"status_code={response.status_code} path={request.url.path}"
    )
    return response

# Optional: enable CORS (here: allow all; tighten in prod)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # replace with your domains in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def build_workbook(payload: Dict[str, Any]) -> io.BytesIO:
    """
    Build the Excel workbook with the exact visual formatting:
      - A1: dateTimeUser (yellow, bold)
      - A2: caption (bold, larger)
      - A3: legend (italic)
      - Row 5: headers (blue fill, white bold, centered) + Excel Table with filters & banded rows
      - Numbers: Turnover EUR -> #,##0.00, Percent -> 0.0% (value normalized /100)
      - Auto column widths
    Returns:
      BytesIO ready to stream as .xlsx
    """
    # Validate minimal shape
    for key in ("caption", "dateTimeUser", "legend", "columns", "rows"):
        if key not in payload:
            raise ValueError(f"Missing required key: '{key}'")

    caption: str = payload["caption"]
    date_time_user: str = payload["dateTimeUser"]

    legend_items = payload["legend"]
    if not isinstance(legend_items, list) or not legend_items:
        legend_text = ""
    else:
        legend_text = " | ".join([f"{item.get('label','')}: {item.get('value','')}" for item in legend_items])

    # Columns and rows
    col_defs: List[Dict[str, str]] = payload["columns"]
    rows: List[Dict[str, Any]] = payload["rows"]

    captions: List[str] = [c["caption"] for c in col_defs]  # display headers
    keys: List[str] = [c["name"] for c in col_defs]         # JSON keys

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    # Header area
    c = ws["A1"]; c.value = date_time_user
    c.fill = PatternFill("solid", fgColor="FFFF00"); c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="left")

    c = ws["A2"]; c.value = caption
    c.font = Font(size=14, bold=True); c.alignment = Alignment(horizontal="left")

    c = ws["A3"]; c.value = legend_text
    c.font = Font(italic=True); c.alignment = Alignment(horizontal="left")

    # Table headers
    header_row = 5
    for col_idx, header in enumerate(captions, start=1):
        hc = ws.cell(row=header_row, column=col_idx, value=header)
        hc.font = Font(bold=True, color="FFFFFF")
        hc.fill = PatternFill("solid", fgColor="4F81BD")  # blue
        hc.alignment = Alignment(horizontal="center")

    # Data rows
    percent_col_idx = None
    for idx, cap in enumerate(captions, start=1):
        if cap.strip() == "%":
            percent_col_idx = idx
            break

    start_row = header_row + 1
    for r_idx, row in enumerate(rows, start=start_row):
        for c_idx, key in enumerate(keys, start=1):
            val = row.get(key, "")
            if percent_col_idx is not None and c_idx == percent_col_idx:
                try:
                    if val is not None and val != "":
                        val = float(val) / 100.0
                except Exception:
                    pass
            ws.cell(row=r_idx, column=c_idx).value = val

    end_row = header_row + len(rows)
    end_col = len(captions)

    # Column widths
    for idx, col_def in enumerate(col_defs, start=1):
        cap = col_def["caption"]; key = col_def["name"]
        max_len_in_data = 0
        for row in rows:
            val = row.get(key, "")
            length = len(str(val))
            if length > max_len_in_data:
                max_len_in_data = length
        ws.column_dimensions[get_column_letter(idx)].width = max(len(str(cap)), max_len_in_data) + 2

    # Number formats
    turnover_idx = None
    for idx, cap in enumerate(captions, start=1):
        if cap.replace("\r", "").replace("\n", "").strip().lower().startswith("turnover"):
            turnover_idx = idx
            break

    if turnover_idx:
        for r in range(start_row, end_row + 1):
            ws.cell(row=r, column=turnover_idx).number_format = "#,##0.00"

    if percent_col_idx:
        for r in range(start_row, end_row + 1):
            ws.cell(row=r, column=percent_col_idx).number_format = "0.0%"

    # Excel Table with filters & banding
    if end_row >= header_row:
        ref = f"A{header_row}:{get_column_letter(end_col)}{max(end_row, header_row)}"
        table = Table(displayName="TurnoverTable", ref=ref)
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = style
        ws.add_table(table)

    # Save to bytes
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.post(
    "/report",
    summary="Upload turnover JSON and receive styled Excel file",
    tags=["Report"],
    responses={
       200: {"content": {"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {}}},
       400: {"description": "Bad Request"},
       401: {"description": "Unauthorized"},
       403: {"description": "Forbidden"},
       500: {"description": "Server Error"},
    },
    openapi_extra={
        "requestBody": {
            "content": {
                "multipart/form-data": {
                    "schema": {
                        "type": "object",
                        "properties": {
                            "file": {"type": "string", "format": "binary"}
                        },
                        "required": ["file"]
                    }
                }
            }
        }
    },
)
async def create_report(file: UploadFile = File(..., description="Upload a `turnover.json` file")):
    """
    **Uploads** a turnover JSON as *multipart/form-data* and **returns** an Excel file with:
    - colored header row (blue)
    - filters & banded rows (Excel Table)
    - currency & percent formatting

    **Field**: `file` → the JSON file to convert.
    """
    logger.info(f"Processing /report upload: filename={file.filename}, content_type={file.content_type}")

    if file.content_type not in ("application/json", "text/json", "application/octet-stream"):
        raise HTTPException(status_code=400, detail="Please upload a JSON file.")

    try:
        raw = await file.read()
        payload = json.loads(raw.decode("utf-8"))
        xlsx_bytes = build_workbook(payload)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="Invalid JSON.")
    except Exception as ex:
        raise HTTPException(status_code=500, detail=f"Failed to generate report: {ex}")

    headers = {"Content-Disposition": 'attachment; filename="turnover-report.xlsx"'}
    logger.info("/report generated Excel successfully")

    return StreamingResponse(
        xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/", include_in_schema=False)
def root():
    return JSONResponse({"status": "ok", "upload_to": "/report", "docs": "/docs", "redoc": "/redoc"})
