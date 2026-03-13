"""
FastAPI server for SAP SF Permission Comparer.

Serves the frontend and provides an API endpoint for comparing
T3 vs PROD permission PDFs using the logic from comparer.py.
"""

import os
import tempfile

from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from comparer import (
    extract_text_from_pdf,
    parse_sections,
    compare_permissions,
    compare_raw_lines,
    parse_permission_lines,
)

app = FastAPI(title="SAP SF Permission Comparer")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

FRONTEND_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "frontend")


@app.get("/")
async def serve_frontend():
    """Serve the frontend single-page application."""
    return FileResponse(os.path.join(FRONTEND_DIR, "index.html"))


@app.post("/api/compare")
async def compare(t3: UploadFile = File(...), prod: UploadFile = File(...)):
    """
    Accept two PDF uploads (t3 and prod), run the comparison pipeline,
    and return the structured differences and raw diff as JSON.
    """
    # Write uploads to temp files (pdfplumber needs file paths)
    t3_tmp = None
    prod_tmp = None
    try:
        t3_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        t3_tmp.write(await t3.read())
        t3_tmp.close()

        prod_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        prod_tmp.write(await prod.read())
        prod_tmp.close()

        # Extract text
        t3_text = extract_text_from_pdf(t3_tmp.name)
        prod_text = extract_text_from_pdf(prod_tmp.name)

        # Parse sections
        t3_sections = parse_sections(t3_text)
        prod_sections = parse_sections(prod_text)

        # Compare
        differences = compare_permissions(t3_sections, prod_sections)
        raw_diff = compare_raw_lines(t3_text, prod_text)

        # Normalize difference keys to camelCase for the frontend
        normalized_differences = []
        for diff in differences:
            normalized_differences.append({
                "section": diff["section"],
                "item": diff["item"],
                "type": diff["type"],
                "t3Value": diff["t3_value"],
                "prodValue": diff["prod_value"],
            })

        return {
            "differences": normalized_differences,
            "rawDiff": {
                "onlyInT3": raw_diff["only_in_t3"],
                "onlyInProd": raw_diff["only_in_prod"],
            },
            "t3FileName": t3.filename,
            "prodFileName": prod.filename,
        }

    finally:
        # Clean up temp files
        if t3_tmp and os.path.exists(t3_tmp.name):
            os.unlink(t3_tmp.name)
        if prod_tmp and os.path.exists(prod_tmp.name):
            os.unlink(prod_tmp.name)
