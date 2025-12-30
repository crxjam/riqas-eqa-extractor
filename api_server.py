from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import tempfile, os, shutil

# ✅ IMPORTANT: import the function that runs your existing pipeline
# You must point this at the function that creates the output workbook for ONE PDF.
from main import process_one_pdf  # <-- rename this to your real function

app = FastAPI()

# allow calls from your GitHub Pages site
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # later we can lock this to https://crxjam.github.io
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/process")
async def process(pdf: UploadFile = File(...)):
    # save uploaded PDF to a temp folder
    with tempfile.TemporaryDirectory() as td:
        pdf_path = os.path.join(td, pdf.filename)
        with open(pdf_path, "wb") as f:
            f.write(await pdf.read())

        # run your EXACT existing script logic
        out_xlsx_path = process_one_pdf(pdf_path)

        # return the excel file as a download
        return FileResponse(
            out_xlsx_path,
            filename=os.path.basename(out_xlsx_path),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
