from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import io
import os

app = FastAPI(title="Excel Viewer API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    # Serve HTML file
    return FileResponse("index.html")

@app.get("/api")
def api_info():
    return {
        "name": "Excel Viewer API",
        "version": "1.0",
        "status": "running"
    }

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    try:
        if not file.filename.endswith(('.xlsx', '.xls')):
            return JSONResponse(
                content={"error": "File must be Excel format"},
                status_code=400
            )
        
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))
        
        return {
            "success": True,
            "filename": file.filename,
            "rows": len(df),
            "columns": len(df.columns),
            "column_names": df.columns.tolist(),
            "preview": df.head(50).to_dict('records')
        }
    except Exception as e:
        return JSONResponse(
            content={"error": str(e)},
            status_code=500
        )
