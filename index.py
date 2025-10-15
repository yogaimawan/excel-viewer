from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import io
import traceback

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
        
        # Convert datetime columns to string
        for col in df.columns:
            if df[col].dtype == 'datetime64[ns]':
                df[col] = df[col].astype(str)
        
        return JSONResponse(
            content={
                "success": True,
                "filename": file.filename,
                "rows": len(df),
                "columns": len(df.columns),
                "column_names": df.columns.tolist(),
                "preview": df.head(50).to_dict('records')
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
