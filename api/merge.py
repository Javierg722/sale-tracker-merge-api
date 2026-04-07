from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import Response
import json
import traceback
from openpyxl import load_workbook
from io import BytesIO

app = FastAPI()

@app.get("/")
def root():
    return {"status": "ok"}

@app.post("/api/merge")
async def merge(
    workbook: UploadFile = File(...),
    data: str = Form(...)
):
    try:
        rows = json.loads(data)

        contents = await workbook.read()
        wb = load_workbook(filename=BytesIO(contents))
        ws = wb["1_Data Entry"]

        START_ROW = 6

        for i, row in enumerate(rows):
            excel_row = START_ROW + i

            ws[f"E{excel_row}"] = row.get("ticker")
            ws[f"G{excel_row}"] = row.get("buyDate")
            ws[f"H{excel_row}"] = row.get("sharesBought")
            ws[f"I{excel_row}"] = row.get("costPerShare")
            ws[f"J{excel_row}"] = row.get("sellDate")
            ws[f"K{excel_row}"] = row.get("sharesSold")
            ws[f"L{excel_row}"] = row.get("salePricePerShare")
            ws[f"N{excel_row}"] = row.get("note")

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return Response(
            content=output.read(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=merged.xlsx"}
        )

    except Exception as e:
        return {
            "error": str(e),
            "trace": traceback.format_exc()
        }
