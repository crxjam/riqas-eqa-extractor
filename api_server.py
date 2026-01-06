from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask

from pathlib import Path
import os
import shutil
import tempfile
from datetime import date
from typing import List

# import your existing functions WITHOUT changing main.py
from main import (
    process_riqas_pdf_into_workbook,
    load_internal_tea_map,
    read_pdf_text,
    parse_metadata,
)

app = FastAPI()

@app.get("/")
def root():
    return {"status": "ok", "docs": "/docs"}

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "https://crxjam.github.io",
    ],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)


def cleanup_dir(tmpdir: str):
    """Delete temp working directory after response is sent."""
    try:
        shutil.rmtree(tmpdir, ignore_errors=True)
    except Exception:
        pass


@app.post("/process")
async def process(
    pdfs: List[UploadFile] = File(...),
    template: UploadFile = File(...),
    tea: UploadFile = File(...),
):
    """
    Accept:
      - multiple PDFs (pdfs)
      - one Excel template (template)
      - one Internal TEa xlsx/csv (tea)
    Return:
      - one rolling workbook (RIQAS_EQA_Rolling_History.xlsx)
    """

    # IMPORTANT: mkdtemp() persists until we delete it ourselves
    tmpdir = tempfile.mkdtemp(prefix="rqxheqa_")
    tmp = Path(tmpdir)
    pdf_dir = tmp / "pdfs"
    pdf_dir.mkdir(parents=True, exist_ok=True)

    try:
        # save template
        template_path = tmp / (template.filename or "template.xlsx")
        template_path.write_bytes(await template.read())

        # save tea file
        tea_path = tmp / (tea.filename or "tea.xlsx")
        tea_path.write_bytes(await tea.read())

        # save PDFs
        pdf_paths: List[Path] = []
        for up in pdfs:
            p = pdf_dir / (up.filename or "input.pdf")
            p.write_bytes(await up.read())
            pdf_paths.append(p)

        # load TEa map
        tea_map = load_internal_tea_map(tea_path)

        # sort PDFs by report date
        def get_pdf_report_date(p: Path):
            txt = read_pdf_text(p)
            meta = parse_metadata(txt)
            return meta.get("report_date") or date.min

        pdf_paths_sorted = sorted(pdf_paths, key=get_pdf_report_date)

        # create rolling workbook from template
        rolling_out = tmp / "RIQAS_EQA_Rolling_History.xlsx"
        shutil.copyfile(template_path, rolling_out)

        # process each PDF into the rolling workbook
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
            # keep tempdir for debugging? (we’ll still clean it)
            return JSONResponse(
                {"ok": ok, "total": len(pdf_paths_sorted), "errors": errors},
                status_code=400,
            )

        # sanity check before returning
        if not rolling_out.exists():
            return JSONResponse(
                {"error": "Output workbook was not created", "expected_path": str(rolling_out)},
                status_code=500,
            )

        # Return file AND only delete tmpdir after response finishes
        return FileResponse(
            path=str(rolling_out),
            filename="RIQAS_EQA_Rolling_History.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            background=BackgroundTask(cleanup_dir, tmpdir),
        )

    except Exception as e:
        # clean up on failure too
        cleanup_dir(tmpdir)
        return JSONResponse({"error": repr(e)}, status_code=500)
