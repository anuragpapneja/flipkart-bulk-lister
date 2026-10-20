"""
FastAPI Backend for Flipkart Bulk Listing Automation Tool.
Handles in-place binary Excel modification (.xls and .xlsx), batch AI generation,
and exact binary download preservation.
"""

import os
import uuid
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from .processor import FlipkartBinaryProcessor
from .gemini_engine import call_gemini_api
from .sample_generator import create_sample_flipkart_workbook

app = FastAPI(
    title="Flipkart Bulk Listing Automation Tool",
    description="Automate Flipkart Bulk Listings by modifying binary Excel files in-place (.xls/.xlsx).",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

STORAGE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "storage")
os.makedirs(STORAGE_DIR, exist_ok=True)

def load_env_file():
    """Load key-value pairs from .env into os.environ if present."""
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for p in [os.path.join(base_dir, ".env"), os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")]:
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith("#") and "=" in line:
                            k, v = line.split("=", 1)
                            k = k.strip()
                            v = v.strip().strip("'\"")
                            if k and k not in os.environ:
                                os.environ[k] = v
            except Exception:
                pass

load_env_file()

PROCESSED_CACHE = {}
LATEST_FILE_ID = None

@app.get("/api/health")
def health_check():
    has_key = bool((os.environ.get("GEMINI_API_KEY") or os.environ.get("ANTIGRAVITY_API_KEY") or "").strip())
    return {
        "status": "ok",
        "app": "Flipkart Bulk Listing Generator (In-Place Binary Engine)",
        "has_api_key": has_key
    }

@app.get("/api/sample")
def download_sample_template():
    """Generate and download a ready-to-test Flipkart bulk listing template."""
    sample_bytes = create_sample_flipkart_workbook()
    return Response(
        content=sample_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="flipkart_sample_seeded_template.xlsx"'}
    )

@app.post("/api/analyze")
async def analyze_file(file: UploadFile = File(...)):
    """Analyze uploaded Excel sheet (.xls or .xlsx) and return metadata."""
    if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        raise HTTPException(status_code=400, detail="Please upload a valid Excel file (.xls, .xlsx, .xlsm)")

    contents = await file.read()
    try:
        processor = FlipkartBinaryProcessor(contents, filename=file.filename)
        analysis = processor.analyze()
        return {"success": True, "analysis": analysis}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to analyze Excel file: {str(e)}")

@app.post("/api/process")
async def process_file(
    file: UploadFile = File(...),
    api_key: Optional[str] = Form(None),
    model: str = Form("gemini-2.5-flash"),
    preserve_existing: bool = Form(False)
):
    """
    Process uploaded Excel file strictly in-place:
    1. Opens uploaded binary workbook directly (.xls via xlrd/xlutils, .xlsx via openpyxl).
    2. Mirrors standard attributes from Master Seed Row (Row 3/4) into subsequent active rows.
    3. Strictly PRESERVES existing Seller SKU IDs and Image URLs.
    4. Overrides ONLY Model Name (<80 chars, 100% unique) and Search Keywords (5 terms with ::).
    5. Returns exact same file format (.xls or .xlsx) so CTRL+SHIFT+S passes with 0 errors.
    """
    if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        raise HTTPException(status_code=400, detail="Please upload a valid Excel file (.xls, .xlsx, .xlsm)")

    contents = await file.read()
    try:
        processor = FlipkartBinaryProcessor(contents, filename=file.filename)
        analysis = processor.analyze()

        active_count = analysis["subsequent_active_rows_count"]
        if active_count == 0:
            raise HTTPException(
                status_code=400,
                detail="No subsequent active rows detected (Row 4 onwards with a Seller SKU ID or Main Image URL)."
            )

        sheet_name = analysis.get("sheet_name", "")
        category_title = sheet_name.replace("_", " ").replace("-", " ").strip().title()
        master_model = analysis.get("master_model_name", "")
        master_keywords = analysis.get("master_keywords", "")
        product_title = master_model or category_title or "Product"

        env_key = os.environ.get("GEMINI_API_KEY", "")
        effective_key = (api_key or "").strip() or env_key.strip()

        # Batch AI generation (Model Name < 80 chars, 100% unique, 5 keywords separated by ::)
        generated_data = call_gemini_api(
            api_key=effective_key,
            master_model=master_model,
            master_title=product_title,
            row_count=active_count,
            master_keywords=master_keywords,
            model=model
        )

        # In-place binary modification (NO reconstruction from scratch)
        processed_bytes, audit, out_filename = processor.process(
            generated_data=generated_data,
            preserve_existing=preserve_existing
        )

        # Determine MIME type based on extension
        if out_filename.lower().endswith(".xls"):
            mime_type = "application/vnd.ms-excel"
        elif out_filename.lower().endswith(".xlsm"):
            mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12"
        else:
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        global LATEST_FILE_ID
        file_id = str(uuid.uuid4())
        LATEST_FILE_ID = file_id
        PROCESSED_CACHE[file_id] = {
            "bytes": processed_bytes,
            "filename": out_filename,
            "media_type": mime_type
        }

        # Persist to disk so files are never lost on reload
        try:
            bin_path = os.path.join(STORAGE_DIR, f"{file_id}.bin")
            meta_path = os.path.join(STORAGE_DIR, f"{file_id}.json")
            with open(bin_path, "wb") as f:
                f.write(processed_bytes)
            import json
            with open(meta_path, "w", encoding="utf-8") as f:
                json.dump({"filename": out_filename, "media_type": mime_type}, f)
        except Exception as e:
            print(f"[Warning] Failed to write persistent cache: {e}")

        return {
            "success": True,
            "file_id": file_id,
            "filename": out_filename,
            "audit": audit,
            "source_engine": generated_data[0].get("source", "hybrid") if generated_data else "unknown"
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing workbook: {str(e)}")

@app.get("/api/download/latest")
def download_latest_file():
    """Download the most recently generated workbook."""
    global LATEST_FILE_ID
    if not LATEST_FILE_ID:
        raise HTTPException(status_code=404, detail="No processed workbook available yet.")
    return download_processed_file(LATEST_FILE_ID)

@app.get("/api/download/{file_id}/{filename}")
@app.get("/api/download/{file_id}")
def download_processed_file(file_id: str, filename: Optional[str] = None):
    """Download the finalized processed workbook with exact format preserved."""
    item = PROCESSED_CACHE.get(file_id)
    if not item:
        # Check disk storage
        bin_path = os.path.join(STORAGE_DIR, f"{file_id}.bin")
        meta_path = os.path.join(STORAGE_DIR, f"{file_id}.json")
        if os.path.exists(bin_path) and os.path.exists(meta_path):
            try:
                import json
                with open(meta_path, "r", encoding="utf-8") as f:
                    meta = json.load(f)
                with open(bin_path, "rb") as f:
                    content_bytes = f.read()
                item = {
                    "bytes": content_bytes,
                    "filename": meta["filename"],
                    "media_type": meta["media_type"]
                }
                PROCESSED_CACHE[file_id] = item
            except Exception as e:
                raise HTTPException(status_code=500, detail=f"Failed to read cached file: {e}")

    if not item:
        raise HTTPException(status_code=404, detail="File expired or not found. Please re-generate.")

    file_bytes = item["bytes"]
    out_name = filename or item["filename"]
    media_type = item["media_type"]

    return Response(
        content=file_bytes,
        media_type=media_type,
        headers={
            "Content-Disposition": f'attachment; filename="{out_name}"; filename*={out_name}',
            "Content-Length": str(len(file_bytes)),
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0"
        }
    )

# Mount frontend static directory
frontend_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "frontend")
if os.path.exists(frontend_dir):
    app.mount("/", StaticFiles(directory=frontend_dir, html=True), name="static")
