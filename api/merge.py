from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/api/merge")
def merge_get():
    return JSONResponse(
        status_code=405,
        content={"detail": "Method not allowed"}
    )

@app.post("/api/merge")
async def merge(
    workbook: UploadFile = File(...),
    data: str = Form(...)
):
    return JSONResponse(
        status_code=200,
        content={
            "received_workbook": workbook.filename if workbook else None,
            "data_length": len(data) if data else 0
        }
    )
