from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
import io, json
from typing import List, Dict, Any

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

app = FastAPI(title="Turnover Report Generator", version="1.0.0")


def build_workbook(payload: Dict[str, Any]) -> io.BytesIO:
    """
    Build the Excel workbook with the exact visual formatting:
      - A1: dateTimeUser (yellow, bold)
      - A2: caption (bold, larger)
      - A3: legend (italic)
      - Row 5: headers (blue fill, white bold, centered) + Excel Table with filters & banded rows
      - Numbers: Turnover EUR -> #,##0.00, Percent -> 0.0% (value normalized /100)
      - Auto column widths
    Returns the workbook bytes in a BytesIO.
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
        # Join all legend pairs like: "Invoice date: 01/01/2025 - 01/12/2025"
        legend_text = " | ".join([f"{item.get('label','')}: {item.get('value','')}" for item in legend_items])

    # Columns and rows
    col_defs: List[Dict[str, str]] = payload["columns"]
    rows: List[Dict[str, Any]] = payload["rows"]

    # Prepare header captions (display) and keys (JSON field names)
    captions: List[str] = [c["caption"] for c in col_defs]
    keys: List[str] = [c["name"] for c in col_defs]

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    # ----- Header area -----
    # A1: dateTimeUser (yellow fill, bold)
    c = ws["A1"]
    c.value = date_time_user
    c.fill = PatternFill("solid", fgColor="FFFF00")
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="left")

    # A2: caption (bigger, bold)
    c = ws["A2"]
    c.value = caption
    c.font = Font(size=14, bold=True)
    c.alignment = Alignment(horizontal="left")

    # A3: legend (italic)
    c = ws["A3"]
    c.value = legend_text
    c.font = Font(italic=True)
    c.alignment = Alignment(horizontal="left")

    # ----- Table headers -----
    header_row = 5
    for col_idx, header in enumerate(captions, start=1):
        hc = ws.cell(row=header_row, column=col_idx, value=header)
        hc.font = Font(bold=True, color="FFFFFF")
        hc.fill = PatternFill("solid", fgColor="4F81BD")  # blue
        hc.alignment = Alignment(horizontal="center")

    # ----- Data rows -----
    # Normalize % column to fraction if present (Excel % format expects 0.064 for 6.4%)
    percent_col_idx = None
    for idx, cap in enumerate(captions, start=1):
        if cap.strip() == "%":
            percent_col_idx = idx
            break

    start_row = header_row + 1
    for r_idx, row in enumerate(rows, start=start_row):
        for c_idx, key in enumerate(keys, start=1):
            val = row.get(key, "")
            # If this is the percent column, convert 6.4 -> 0.064
            if percent_col_idx is not None and c_idx == percent_col_idx:
                try:
                    if val is not None and val != "":
                        val = float(val) / 100.0
                except Exception:
                    # Leave as-is if not numeric; Excel will just show raw
                    pass
            ws.cell(row=r_idx, column=c_idx).value = val

    end_row = header_row + len(rows)
    end_col = len(captions)

    # ----- Column widths -----
    for idx, col_def in enumerate(col_defs, start=1):
        cap = col_def["caption"]
        key = col_def["name"]
        max_len_in_data = 0
        for row in rows:
            val = row.get(key, "")
            length = len(str(val))
            if length > max_len_in_data:
                max_len_in_data = length
        width = max(len(str(cap)), max_len_in_data) + 2
        ws.column_dimensions[get_column_letter(idx)].width = width

    # ----- Number formats -----
    # Turnover column (caption may contain newline "Turnover\r\nEUR")
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

    # ----- Excel Table with filters & banding -----
    if end_row >= header_row:  # safeguard
        ref = f"A{header_row}:{get_column_letter(end_col)}{end_row if end_row >= header_row else header_row}"
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


@app.post("/report", summary="Upload turnover JSON and receive styled Excel file")
async def create_report(file: UploadFile = File(...)):
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
        # For production: log ex
        raise HTTPException(status_code=500, detail=f"Failed to generate report: {ex}")

    headers = {
        "Content-Disposition": 'attachment; filename="turnover-report.xlsx"'
    }
    return StreamingResponse(
        xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/", include_in_schema=False)
def root():
    return JSONResponse({"status": "ok", "upload_to": "/report", "method": "POST (multipart/form-data)"} )
