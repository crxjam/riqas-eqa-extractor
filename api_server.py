from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from starlette.background import BackgroundTask
from fastapi.staticfiles import StaticFiles

from pathlib import Path
import shutil
import tempfile
from datetime import date
from typing import List
from io import BytesIO

# import your existing functions WITHOUT changing main.py
from main import (
    process_riqas_pdf_into_workbook,
    load_internal_tea_map,
    read_pdf_text,
    parse_metadata,
)

app = FastAPI()

# -----------------------
# API routes FIRST
# -----------------------

@app.get("/health", include_in_schema=False)
def health():
    return {"status": "ok"}

@app.get("/__routes", include_in_schema=False)
def __routes():
    return sorted([
        f"{r.path}::{','.join(sorted(getattr(r, 'methods', []) or []))}"
        for r in app.router.routes
    ])

@app.post("/process")
async def process(
    pdfs: List[UploadFile] = File(...),
    template: UploadFile = File(...),
    tea: UploadFile = File(...),
):
    tmpdir = tempfile.mkdtemp(prefix="rqxheqa_")
    tmp = Path(tmpdir)
    pdf_dir = tmp / "pdfs"
    pdf_dir.mkdir(parents=True, exist_ok=True)

    def cleanup_dir(d: str):
        try:
            shutil.rmtree(d, ignore_errors=True)
        except Exception:
            pass

    try:
        template_path = tmp / (template.filename or "template.xlsx")
        template_path.write_bytes(await template.read())

        tea_path = tmp / (tea.filename or "tea.xlsx")
        tea_path.write_bytes(await tea.read())

        pdf_paths: List[Path] = []
        for up in pdfs:
            p = pdf_dir / (up.filename or "input.pdf")
            p.write_bytes(await up.read())
            pdf_paths.append(p)

        tea_map = load_internal_tea_map(tea_path)

        def get_pdf_report_date(p: Path):
            txt = read_pdf_text(p)
            meta = parse_metadata(txt)
            return meta.get("report_date") or date.min

        pdf_paths_sorted = sorted(pdf_paths, key=get_pdf_report_date)

        rolling_out = tmp / "RIQAS_EQA_Rolling_History.xlsx"
        shutil.copyfile(template_path, rolling_out)

        errors = []
        ok = 0

        for pdf in pdf_paths_sorted:
            try:
                process_riqas_pdf_into_workbook(
                    pdf_path=str(pdf),
                    xlsx_path=str(rolling_out),
                    out_path=str(rolling_out),
                    internal_tea_map=tea_map,
                )
                ok += 1
            except Exception as e:
                errors.append(f"{pdf.name}: {repr(e)}")

        if errors:
            cleanup_dir(tmpdir)
            return JSONResponse(
                {"ok": ok, "total": len(pdf_paths_sorted), "errors": errors},
                status_code=400,
            )

        data = rolling_out.read_bytes()
        buf = BytesIO(data)
        buf.seek(0)

        headers = {
            "Content-Disposition": 'attachment; filename="EQA_Rolling_History.xlsx"',
            "Cache-Control": "no-store",
        }

        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
            background=BackgroundTask(cleanup_dir, tmpdir),
        )

    except Exception as e:
        cleanup_dir(tmpdir)
        return JSONResponse({"error": repr(e)}, status_code=500)

# -----------------------
# CORS
# -----------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "http://localhost:5174",
        "http://127.0.0.1:5174",
        "https://crxjam.github.io",
    ],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)

# -----------------------
# Frontend LAST (so it can't steal /process)
# -----------------------
DIST_DIR = Path(__file__).resolve().parent / "dist"
if DIST_DIR.exists():
    app.mount("/", StaticFiles(directory=str(DIST_DIR), html=True), name="frontend")
