from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import io
import traceback
import json
from datetime import datetime, date

app = FastAPI(title="Excel Viewer API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/", response_class=HTMLResponse)
def root():
    try:
        with open("index.html", "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        return f"<h1>Error loading page: {str(e)}</h1>"

@app.get("/api")
def api_info():
    return {
        "name": "Excel Viewer API",
        "version": "1.0",
        "status": "running"
    }

def convert_to_serializable(obj):
    """Convert pandas/numpy types to JSON serializable types"""
    if pd.isna(obj):
        return None
    elif isinstance(obj, (datetime, date)):
        return obj.isoformat()
    elif isinstance(obj, (pd.Timestamp)):
        return obj.isoformat()
    elif hasattr(obj, 'item'):  # numpy types
        return obj.item()
    else:
        return obj

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    try:
        if not file.filename.endswith(('.xlsx', '.xls')):
            return JSONResponse(
                content={
                    "success": False,
                    "error": "File must be .xlsx or .xls format"
                },
                status_code=400
            )
        
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents), engine='openpyxl')
        
        # Convert ALL data to JSON-serializable format
        preview_data = []
        for _, row in df.head(50).iterrows():
            row_dict = {}
            for col in df.columns:
                row_dict[col] = convert_to_serializable(row[col])
            preview_data.append(row_dict)
        
        return JSONResponse(
            content={
                "success": True,
                "filename": file.filename,
                "rows": int(len(df)),
                "columns": int(len(df.columns)),
                "column_names": [str(col) for col in df.columns],
                "preview": preview_data
            }
        )
        
    except Exception as e:
        error_details = traceback.format_exc()
        print(f"Error: {error_details}")
        
        return JSONResponse(
            content={
                "success": False,
                "error": f"Failed to process file: {str(e)}"
            },
            status_code=500
        )
