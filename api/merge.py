from io import BytesIO
from fastapi import FastAPI, UploadFile, File, Form, Response, HTTPException
from openpyxl import load_workbook
from datetime import datetime
import json

app = FastAPI()

SHEET_NAME = "1_Data Entry"
START_ROW = 6

INPUT_COLUMNS = {
    "ticker": "E",
    "buyDate": "G",
    "sharesBought": "H",
    "costPerShare": "I",
    "sellDate": "J",
    "sharesSold": "K",
    "salePricePerShare": "L",
    "note": "N",
}

def clear_cell(ws, cell_ref: str):
    ws[cell_ref].value = None

def parse_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except:
        return None

@app.get("/")
def root():
    return {"status": "ok"}

@app.post("/api/merge")
async def merge_post(
    workbook: UploadFile = File(...),
    data: str = Form(...)
):
    try:
        rows = json.loads(data)
        if not isinstance(rows, list):
            raise HTTPException(status_code=400, detail="Data must be a JSON array")

        content = await workbook.read()
        wb = load_workbook(BytesIO(content))

        if SHEET_NAME not in wb.sheetnames:
            raise HTTPException(status_code=400, detail=f'Sheet "{SHEET_NAME}" not found')

        ws = wb[SHEET_NAME]

        # Force Excel FULL recalculation
        wb.calculation_properties.fullCalcOnLoad = True

        # Clear input cells
        for row_num in range(6, 506):
            for col in INPUT_COLUMNS.values():
                clear_cell(ws, f"{col}{row_num}")

        # Write rows
        for i, row in enumerate(rows, start=START_ROW):
            if i > 505:
                raise HTTPException(status_code=400, detail="Too many rows")

            for field, col in INPUT_COLUMNS.items():
                value = row.get(field)
                if value in (None, ""):
                    continue

                cell = ws[f"{col}{i}"]

                # NUMBERS
                if field in ("sharesBought", "costPerShare", "sharesSold", "salePricePerShare"):
                    cell.value = float(value)

                # DATES (CRITICAL FIX)
                elif field in ("buyDate", "sellDate"):
                    dt = parse_date(value)
                    if dt:
                        cell.value = dt
                        cell.number_format = "m/d/yyyy"

                # TEXT
                else:
                    cell.value = value

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return Response(
            content=output.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": 'attachment; filename="merged.xlsx"',
                "Access-Control-Allow-Origin": "*",
            },
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
