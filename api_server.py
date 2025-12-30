from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
import tempfile, os, shutil
from pathlib import Path

# import your existing functions WITHOUT changing main.py
from main import process_riqas_pdf_into_workbook, load_internal_tea_map, read_pdf_text, parse_metadata
import pandas as pd


app = FastAPI()

@app.post("/process")
async def process(
    pdfs: list[UploadFile] = File(...),
    template: UploadFile = File(...),
    tea: UploadFile = File(...),
):
    """
    Accept:
      - multiple PDFs
      - one Excel template
      - one Internal TEa xlsx/csv
    Return:
      - one rolling workbook (RIQAS_EQA_Rolling_History.xlsx)
    """
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pdf_dir = tmp / "pdfs"
        pdf_dir.mkdir(parents=True, exist_ok=True)

        # save template
        template_path = tmp / template.filename
        with open(template_path, "wb") as f:
            f.write(await template.read())

        # save tea file
        tea_path = tmp / tea.filename
        with open(tea_path, "wb") as f:
            f.write(await tea.read())

        # save PDFs
        pdf_paths = []
        for up in pdfs:
            p = pdf_dir / up.filename
            with open(p, "wb") as f:
                f.write(await up.read())
            pdf_paths.append(p)

        # load TEa map
        tea_map = load_internal_tea_map(tea_path)

        # sort PDFs by report date (same logic as your GUI)
        from datetime import date
        def get_pdf_report_date(p: Path):
            txt = read_pdf_text(p)
            meta = parse_metadata(txt)
            return meta.get("report_date") or date.min

        pdf_paths_sorted = sorted(pdf_paths, key=get_pdf_report_date)

        # create rolling workbook once
        rolling_out = tmp / "RIQAS_EQA_Rolling_History.xlsx"
        shutil.copyfile(template_path, rolling_out)

        # process each PDF into same rolling workbook
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
            return JSONResponse({"ok": ok, "total": len(pdf_paths_sorted), "errors": errors}, status_code=400)

        # return the file
        return FileResponse(
            path=str(rolling_out),
            filename="RIQAS_EQA_Rolling_History.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
