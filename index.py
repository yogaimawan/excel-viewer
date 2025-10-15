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
        # Validate file extension
        if not file.filename.endswith(('.xlsx', '.xls')):
            return JSONResponse(
                content={
                    "success": False,
                    "error": "File must be .xlsx or .xls format"
                },
                status_code=400
            )
        
        # Read file contents
        contents = await file.read()
        
        # Read Excel with ALL columns as string
        df = pd.read_excel(
            io.BytesIO(contents), 
            engine='openpyxl',
            dtype=str,  # Force semua kolom jadi string
            keep_default_na=False  # Jangan convert empty cells ke NaN
        )
        
        # Force column names jadi string (handle datetime headers)
        df.columns = [str(col) for col in df.columns]
        
        # Replace NaN dengan empty string
        df = df.fillna('')
        
        # Convert to dict (semua udah string, aman untuk JSON)
        preview_data = df.head(50).to_dict('records')
        
        # Return response
        return JSONResponse(
            content={
                "success": True,
                "filename": file.filename,
                "rows": int(len(df)),
                "columns": int(len(df.columns)),
                "column_names": list(df.columns),
                "preview": preview_data
            }
        )
        
    except Exception as e:
        # Log detailed error
        error_details = traceback.format_exc()
        print(f"Error processing file:")
        print(error_details)
        
        return JSONResponse(
            content={
                "success": False,
                "error": f"Failed to process file: {str(e)}"
            },
            status_code=500
        )
