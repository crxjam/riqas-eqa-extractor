import re
import datetime
from pathlib import Path

import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from typing import Optional



############################################
# 1. PDF PARSING
############################################

def read_pdf_text(pdf_path: Path) -> str:
    chunks = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            chunks.append(page.extract_text() or "")
    return "\n".join(chunks)



def parse_metadata(full_text: str) -> dict:
    """
    Pull top-level metadata like Cycle, Sample, Date, Lab name, Instrument, Panel.
    These strings are present in the PDF header/intro.

    Example we look for:
    "CYCLE 18    SAMPLE 10    13/10/2025"
    "MONTHLY HAEMATOLOGY"
    "NHLS Red Cross Childrens Hospital"
    """
    # Cycle / Sample / Date
    m = re.search(r"CYCLE\s+(\d+)\s+SAMPLE\s+(\d+)\s+(\d{2}/\d{2}/\d{4})", full_text)
    if m:
        cycle_no = int(m.group(1))
        sample_no = int(m.group(2))
        report_date_str = m.group(3)
        report_date = datetime.datetime.strptime(report_date_str, "%d/%m/%Y").date()
    else:
        cycle_no = None
        sample_no = None
        report_date = None

    # Panel / Scheme name e.g. "MONTHLY HAEMATOLOGY"
    panel_match = re.search(r"(MONTHLY\s+[A-Z ]+)", full_text)
    panel_name = panel_match.group(1).strip() if panel_match else None

    # Lab name: first meaningful line usually near top of page 1
    # We'll take the first line that contains "NHLS" or "Hospital"
    lab_name = None
    for line in full_text.splitlines()[:50]:
        line_clean = line.strip()
        if len(line_clean) < 3:
            continue
        if "NHLS" in line_clean or "Hospital" in line_clean or "HOSPITAL" in line_clean:
            lab_name = line_clean
            break

    # Instrument / Method: often visible near header as something like "Sysmex XN2000"
    # We'll try to grab a Sysmex / Roche / Beckman string if present.
    instrument_match = re.search(
        r"(Sysmex\s+[A-Za-z0-9\-/ ]+|Beckman\s+Coulter\s+[A-Za-z0-9\-/ ]+|Roche\s+[A-Za-z0-9\-/ ]+)",
        full_text[:2000]
    )
    instrument_name = instrument_match.group(1).strip() if instrument_match else None

    return {
        "cycle_no": cycle_no,
        "sample_no": sample_no,
        "report_date": report_date,
        "panel_name": panel_name,
        "lab_name": lab_name,
        "instrument_name": instrument_name,
    }


def extract_tdpa(full_text: str, idx: int) -> Optional[float]:
    """
    TEa / performance tolerance usually appears on the page as:
    'TDPA = 3.7%'
    We'll search ~2000 chars after the analyte label first.
    """
    window = full_text[idx:idx + 2500]
    m = re.search(r"TDPA\s*=\s*([0-9.]+)%", window)
    if m:
        return float(m.group(1))
    # fallback: look backwards
    window_back = full_text[max(0, idx - 2500):idx]
    m2 = re.search(r"TDPA\s*=\s*([0-9.]+)%", window_back)
    if m2:
        return float(m2.group(1))
    return None


from typing import Optional, Dict

def parse_single_analyte(full_text: str, analyte_label: str) -> Optional[Dict]:
    idx = full_text.find(analyte_label)
    if idx == -1:
        return None

    # Work on a window AFTER the analyte label (that’s where the keyed fields are)
    post = full_text[idx: idx + 2000]

    # Your Result and SDI appear on the same line in your sample:
    # "Your Result 353.000 SDI -6.11"
    m_yr_sdi = re.search(r"Your Result\s+([-\d.]+)\s+SDI\s+([-\d.]+)", post)
    if not m_yr_sdi:
        return None
    yr = float(m_yr_sdi.group(1))
    sdi = float(m_yr_sdi.group(2))

    # Mean for Comparison and TS appear on the same line:
    # "Mean for Comparison 521.531 TS 10"
    m_mean_ts = re.search(r"Mean for Comparison\s+([-\d.]+)\s+TS\s+([-\d.]+)", post)
    if not m_mean_ts:
        return None
    mean_comp = float(m_mean_ts.group(1))
    ts = float(m_mean_ts.group(2))

    # %DEV line:
    # "%DEV -32.3 Sample Number"  (so take the number right after %DEV)
    m_dev = re.search(r"%DEV\s+([-\d.]+)", post)
    dev = float(m_dev.group(1)) if m_dev else None

    # TDPA:
    m_tdpa = re.search(r"TDPA\s*=\s*([0-9.]+)%", post)
    tdpa = float(m_tdpa.group(1)) if m_tdpa else None

    return {
        "AnalyteRaw": analyte_label,
        "Peer Mean": mean_comp,
        "Your Result": yr,
        "%DEV": dev,
        "SDI": sdi,
        "Target Score": ts,
        "TDPA_limit_percent": tdpa,
    }

def find_analyte_labels(full_text: str) -> list[str]:
    """
    Find analyte header labels that appear right after 'Mean for Comparison'
    e.g. 'Amylase, U/l @ 37°C', 'Calcium, mmol/l', 'Chloride, mmol/l'
    """
    labels = []
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]

    for i, ln in enumerate(lines[:-1]):
        if ln.lower() == "mean for comparison":
            cand = lines[i + 1]

            # Filter out junk headers
            if cand.lower().startswith("laboratory ref"):
                continue
            if cand.lower().startswith("cycle"):
                continue

            # Heuristic: analyte line usually contains a comma + unit-ish
            # allow U/l, mmol/l, umol/l, mg/l, g/l, etc.
            if re.search(r",\s*[A-Za-zµ/%]+\s*/\s*[A-Za-z]+", cand):
                labels.append(cand)

    # de-dup while preserving order
    seen = set()
    out = []
    for x in labels:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out

def extract_all_analytes(full_text: str) -> pd.DataFrame:
    labels = find_analyte_labels(full_text)

    rows = []
    for label in labels:
        rec = parse_single_analyte(full_text, label)
        if rec:
            rows.append(rec)

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df["diff_from_mean"] = df["Your Result"] - df["Peer Mean"]
    return df

############################################
# 2. RISK LOGIC (current + historical)
############################################

def flag_bias(last3_diffs: pd.Series) -> bool:
    """
    Bias definition: ≥3 results on the same side of the mean.
    We interpret this as last 3 diffs all >0 or all <0.
    """
    if len(last3_diffs) < 3:
        return False
    return (last3_diffs.gt(0).all()) or (last3_diffs.lt(0).all())


def flag_trend(last3_sdi: pd.Series) -> bool:
    """
    Trend definition: ≥3 results moving in same direction within ±1 SD.
    We'll check:
    - all abs(SDI) <= 1
    - strictly increasing OR strictly decreasing across the last 3 points
    """
    if len(last3_sdi) < 3:
        return False
    if not (last3_sdi.abs() <= 1).all():
        return False
    inc = last3_sdi.iloc[0] < last3_sdi.iloc[1] < last3_sdi.iloc[2]
    dec = last3_sdi.iloc[0] > last3_sdi.iloc[1] > last3_sdi.iloc[2]
    return inc or dec


def base_risk_for_row(row: pd.Series) -> str:
    abs_sdi = abs(row["SDI"])
    ts = row["Target Score"]
    abs_dev = abs(row["%DEV"]) if pd.notnull(row["%DEV"]) else None
    tea_limit = row["TDPA_limit_percent"]

    high_hits = []
    if abs_sdi >= 2:
        high_hits.append("SDI≥2")
    if ts < 40:
        high_hits.append("TS<40")
    if tea_limit is not None and abs_dev is not None and abs_dev > tea_limit:
        high_hits.append(">%TEa")

    if high_hits:
        return "High"

    # Moderate: borderline but not failing
    if (1 <= abs_sdi < 2) or (41 <= ts <= 50):
        return "Moderate"

    return "Low"



def escalate_for_history(hist_df: pd.DataFrame, analyte_name: str, new_row: pd.Series,
                         this_cycle_risk: str) -> tuple[str, bool, bool]:
    """
    Incorporate historical cycles to:
      - upgrade to Critical if recurrent failure (>2 cycles)
      - identify bias/trend over last 3 cycles
    hist_df: full Cycle_History INCLUDING the new row (so last row is current)
    We'll subset hist for this analyte.
    """

    sub = (
        hist_df[hist_df["Analyte"] == analyte_name]
        .sort_values(["Report_Date"])
        .reset_index(drop=True)
    )

    # Bias / Trend flags from last 3 points:
    bias_last3 = flag_bias(sub["diff_from_mean"].tail(3))
    trend_last3 = flag_trend(sub["SDI"].tail(3))

    upgraded_risk = this_cycle_risk

    # Recurrent failure rule:
    # "Recurrent failure (>2 cycles) / Persistent failure -> Critical"
    # We interpret: if this cycle is High AND any High/Critical in the immediately previous cycle,
    # OR if we already had ≥2 High/Critical in the recent past.
    recent = sub["Risk_Category_BaseOnly"].tail(3).tolist()
    # Count how many High-or-worse in the last 3 (including now)
    high_like_count = sum(r in ["High", "Critical"] for r in recent)
    if high_like_count >= 2 and this_cycle_risk in ["High", "Critical"]:
        upgraded_risk = "Critical"

    return upgraded_risk, bias_last3, trend_last3


def build_comment_and_action(final_risk: str,
                             bias_flagged: bool,
                             trend_flagged: bool) -> tuple[str, str]:
    """
    Gives:
    - Interpretation / Comments (goes into Result Summary, auditable narrative)
    - Required Action (for escalation sheet)
    """
    if final_risk == "Critical":
        action = "Escalate to LM; notify QA"
        comment = "Persistent failure (≥2 cycles). Immediate escalation required."
        return comment, action

    if final_risk == "High":
        action = "Investigate immediately; complete RCA and alert LM"
        comment = "Outside TEa / SDI ≥2 or low target score. Immediate RCA required."
        return comment, action

    # Moderate
    if final_risk == "Moderate":
        action = "Monitor next EQA cycles; review calibration"
        if bias_flagged and trend_flagged:
            comment = "Developing bias and trend (≥3 results consistent direction within ±1 SD). Monitor closely."
        elif bias_flagged:
            comment = "Developing bias (≥3 results on same side of mean). Monitor next cycle."
        elif trend_flagged:
            comment = "Developing trend (drift over ≥3 cycles within ±1 SD). Monitor next cycle."
        else:
            comment = "Slight bias/trend or borderline target score (41–50). Monitor."
        return comment, action

    # Low
    action = "File record; continue routine QC"
    comment = "Acceptable: within TEa, SDI<2, TS>50. No significant bias/trend."
    return comment, action


############################################
# 3. EXCEL UPDATE
############################################

def ensure_cycle_history_sheet(wb: Workbook):
    """Create Cycle_History sheet with headers if it doesn't exist yet."""
    if "Cycle_History" not in wb.sheetnames:
        ws = wb.create_sheet("Cycle_History")
        headers = [
            "Cycle_No",
            "Sample_No",
            "Report_Date",        # as date
            "Analyte",
            "Peer_Mean",
            "Your_Result",
            "%DEV",
            "SDI",
            "Target_Score",
            "TEa_or_TDPA(%)",
            "diff_from_mean",
            "Risk_Category_BaseOnly",
            "Risk_Category_Final",
            "Bias_Flag_Last3",
            "Trend_Flag_Last3",
            "Required_Action",
            "Comment",
        ]
        for col_i, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_i).value = header


def ws_to_dataframe(ws) -> pd.DataFrame:
    """Helper: read an openpyxl worksheet into a DataFrame (header row is row 1)."""
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(list(row))
    df = pd.DataFrame(data[1:], columns=data[0])
    # Drop completely empty rows
    df = df.dropna(how="all")
    return df


def append_df_to_worksheet(ws, df: pd.DataFrame):
    """Append rows of df to ws (no header)."""
    for _, row in df.iterrows():
        ws.append(list(row))


def fill_header_information_sheet(wb: Workbook, meta: dict):
    """
    Populate 'Header Information' sheet:
    A1: 'Laboratory Name' -> put value in B1, etc.
    We'll just write into column B for the known rows.
    """
    if "Header Information" not in wb.sheetnames:
        return
    ws = wb["Header Information"]

    mapping = {
        "Laboratory Name": meta.get("lab_name"),
        "Instrument/Method": meta.get("instrument_name"),
        "Analyte / Panel": meta.get("panel_name"),
        "Cycle / Sample No.": f"Cycle {meta.get('cycle_no')} / Sample {meta.get('sample_no')}",
        "Date of Report": meta.get("report_date").strftime("%d/%m/%Y") if meta.get("report_date") else None,
        "EQA Scheme": "RIQAS (Randox International Quality Assessment Scheme)",  # already present
    }

    # loop rows, put in column B (col 2)
    for row in range(1, 20):
        label = ws.cell(row=row, column=1).value
        if label in mapping and mapping[label] is not None:
            ws.cell(row=row, column=2).value = mapping[label]


def map_parameter_name(analyte_raw: str) -> str:
    """
    Convert the RIQAS analyte label into the row label used in 'Result Summary'.
    This mapping can be expanded as needed.
    """
    mapping_table = {
        "Haemoglobin, g/dl": "Haemoglobin (g/dL)",
        "Haematocrit (HCT), L/L": "Haematocrit (L/L)",
        "MCH, pg": "MCH (pg)",
        "MCHC, g/dl": "MCHC (g/dL)",
        "MCV, fL": "MCV (fL)",
        "Mean Platelet Volume, fL": "MPV (fL)",
        "Plateletcrit, %": "Plateletcrit (%)",
        "Platelets (Impedance Count), X 10 9/L": "Platelets (x10^9/L)",
        "RBC (Impedance Count), X 10 12/L": "RBC (x10^12/L)",
        "Red Cell Dist. Width CV, %": "RDW-CV (%)",
        "WBC (Optical Count), X 10 9/L": "WBC (x10^9/L)",
    }
    return mapping_table.get(analyte_raw, analyte_raw)


def update_result_summary_sheet(wb: Workbook,
                                current_cycle_rows: pd.DataFrame,
                                final_eval: pd.DataFrame):
    """
    'Result Summary' sheet has this header:
        A: Parameter
        B: Peer Group Mean
        C: Your Result
        D: %Deviation
        E: SDI
        F: Target Score
        G: Acceptable Limit (TEa or RIQAS)
        H: Acceptable? (Y/N)
        I: Interpretation / Comments

    We'll overwrite rows 2..N with the latest cycle values.
    We'll also inject a standardised comment aligned to SANAS language.
    """
    if "Result Summary" not in wb.sheetnames:
        return

    ws = wb["Result Summary"]

    # First: clear old numeric cells in rows 2-50 for cols B..I
    for row in range(2, 60):
        for col in range(2, 10):  # B..I
            ws.cell(row=row, column=col).value = None

    # Build a lookup of risk/action/comment from final_eval
    # final_eval columns:
    #   Analyte, Risk_Category_Final, Comment, Required_Action
    eval_lookup = (
        final_eval.set_index("Parameter_Name")
                  [["Risk_Category_Final", "Comment", "Required_Action", "TDPA_limit_percent"]]
        .to_dict(orient="index")
    )

    # Now walk the sheet and if Parameter cell matches, fill.
    for row in range(2, 200):
        param_cell_val = ws.cell(row=row, column=1).value
        if not param_cell_val:
            continue

        # find matching record in current_cycle_rows
        subset = current_cycle_rows[current_cycle_rows["Parameter_Name"] == param_cell_val]
        if subset.empty:
            continue
        rec = subset.iloc[0]

        # acceptable? logic for column H:
        # Acceptable if:
        #   SDI < 2
        #   Target Score > 50
        #   abs(%DEV) <= TEa
        abs_sdi = abs(rec["SDI"])
        ts = rec["Target Score"]
        abs_dev = abs(rec["%DEV"]) if pd.notnull(rec["%DEV"]) else None
        tea_lim = rec["TDPA_limit_percent"]
        acceptable = (
            (abs_sdi < 2) and
            (ts > 50) and
            (tea_lim is not None and abs_dev is not None and abs_dev <= tea_lim)
        )
        acceptable_flag = "Y" if acceptable else "N"

        # Pull risk-derived narrative
        risk_info = eval_lookup.get(param_cell_val, {})
        narrative_comment = risk_info.get("Comment", "")
        tea_display = rec["TDPA_limit_percent"]

        ws.cell(row=row, column=2).value = float(rec["Peer Mean"])
        ws.cell(row=row, column=3).value = float(rec["Your Result"])
        ws.cell(row=row, column=4).value = float(rec["%DEV"]) if pd.notnull(rec["%DEV"]) else None
        ws.cell(row=row, column=5).value = float(rec["SDI"])
        ws.cell(row=row, column=6).value = float(rec["Target Score"])
        ws.cell(row=row, column=7).value = tea_display
        ws.cell(row=row, column=8).value = acceptable_flag
        ws.cell(row=row, column=9).value = narrative_comment


def update_cycle_history_sheet(wb: Workbook,
                               meta: dict,
                               enriched_rows: pd.DataFrame):
    """
    Append new rows (the latest cycle) into Cycle_History.
    Cycle_History is the auditable longitudinal evidence for SANAS:
      - shows consecutive bias, trend, recurrence
      - shows final risk and action
      - supports "Critical" escalation
    """
    ws = wb["Cycle_History"]

    append_df_to_worksheet(ws, enriched_rows)


############################################
# 4. MAIN ORCHESTRATION
############################################

def process_riqas_pdf_into_workbook(pdf_path: str, xlsx_path: str, out_path: Optional[str] = None):
    """
    Full pipeline:
      - Read PDF
      - Extract analyte performance
      - Load workbook
      - Update Header Information
      - Update Result Summary (this cycle)
      - Update/append Cycle_History (all cycles, classified with history-aware risk)
      - Save workbook (in-place or to new file)
    """
    pdf_path = Path(pdf_path)
    xlsx_path = Path(xlsx_path)
    out_path = Path(out_path) if out_path else xlsx_path

    # --- parse PDF ---
    full_text = read_pdf_text(pdf_path)
    meta = parse_metadata(full_text)
    df = extract_all_analytes(full_text)

    # Attach metadata columns now (cycle/date/sample etc)
    df["Cycle_No"] = meta["cycle_no"]
    df["Sample_No"] = meta["sample_no"]
    df["Report_Date"] = meta["report_date"]

    # Rename for workbook logic
    df["Parameter_Name"] = df["AnalyteRaw"]

    # --- open workbook / init sheets ---
    if xlsx_path.exists():
        wb = load_workbook(xlsx_path)
    else:
        wb = Workbook()

    ensure_cycle_history_sheet(wb)

    # We'll convert Cycle_History (existing) to df_hist for historical risk calc
    hist_ws = wb["Cycle_History"]
    hist_df_existing = ws_to_dataframe(hist_ws) if hist_ws.max_row > 1 else pd.DataFrame(columns=[
        "Cycle_No",
        "Sample_No",
        "Report_Date",
        "Analyte",
        "Peer_Mean",
        "Your_Result",
        "%DEV",
        "SDI",
        "Target_Score",
        "TEa_or_TDPA(%)",
        "diff_from_mean",
        "Risk_Category_BaseOnly",
        "Risk_Category_Final",
        "Bias_Flag_Last3",
        "Trend_Flag_Last3",
        "Required_Action",
        "Comment",
    ])

    # --- build risk per analyte, including historical escalation ---
    enriched_rows_for_history = []
    summary_eval_rows = []

    for _, row in df.iterrows():
        analyte_name = row["Parameter_Name"]

        # Build a temp frame with existing + this new measurement for THIS analyte
        sub_hist = hist_df_existing[hist_df_existing["Analyte"] == analyte_name].copy()

        new_point = {
            "Cycle_No": row["Cycle_No"],
            "Sample_No": row["Sample_No"],
            "Report_Date": row["Report_Date"],
            "Analyte": analyte_name,
            "Peer_Mean": row["Peer Mean"],
            "Your_Result": row["Your Result"],
            "%DEV": row["%DEV"],
            "SDI": row["SDI"],
            "Target_Score": row["Target Score"],
            "TEa_or_TDPA(%)": row["TDPA_limit_percent"],
            "diff_from_mean": row["diff_from_mean"],
        }

        new_point_df = pd.DataFrame([new_point])

        # If there's no prior history for this analyte, just use the new row
        if sub_hist.empty or sub_hist.dropna(how="all").empty:
            simulated_hist = new_point_df.copy()
        else:
            simulated_hist = pd.concat(
                [sub_hist, new_point_df],
                ignore_index=True,
                sort=False,
                copy=False,
            )

        simulated_hist = simulated_hist.sort_values(["Report_Date"]).reset_index(drop=True)

        # STEP 1: base risk for this single cycle
        base_risk_label = base_risk_for_row(row)

        # We'll set this base risk to last row (i.e. this new cycle),
        # then escalate using historical context in simulated_hist
        simulated_hist["Risk_Category_BaseOnly"] = None
        simulated_hist.loc[simulated_hist.index[-1], "Risk_Category_BaseOnly"] = base_risk_label

        final_risk_label, bias_last3, trend_last3 = escalate_for_history(
            simulated_hist,
            analyte_name,
            row,
            base_risk_label,
        )

        comment_txt, action_txt = build_comment_and_action(
            final_risk_label,
            bias_last3,
            trend_last3
        )

        simulated_hist.loc[simulated_hist.index[-1], "Risk_Category_Final"] = final_risk_label
        simulated_hist.loc[simulated_hist.index[-1], "Bias_Flag_Last3"] = bool(bias_last3)
        simulated_hist.loc[simulated_hist.index[-1], "Trend_Flag_Last3"] = bool(trend_last3)
        simulated_hist.loc[simulated_hist.index[-1], "Required_Action"] = action_txt
        simulated_hist.loc[simulated_hist.index[-1], "Comment"] = comment_txt

        # harvest just the *new row* (the last one we added)
        new_hist_row = simulated_hist.iloc[[-1]]
        enriched_rows_for_history.append(new_hist_row)

        # also collect info for populating Result Summary sheet
        summary_eval_rows.append({
            "Parameter_Name": analyte_name,
            "Risk_Category_Final": final_risk_label,
            "Comment": comment_txt,
            "Required_Action": action_txt,
            "TDPA_limit_percent": row["TDPA_limit_percent"],
        })

    # concat all analytes new rows
    enriched_rows_for_history_df = pd.concat(enriched_rows_for_history, ignore_index=True)

    # --- write Header Information sheet ---
    fill_header_information_sheet(wb, meta)

    # --- write Result Summary sheet (for THIS cycle) ---
    summary_eval_df = pd.DataFrame(summary_eval_rows)
    update_result_summary_sheet(
        wb,
        current_cycle_rows=df,
        final_eval=summary_eval_df
    )

    # --- append to Cycle_History sheet (for ALL cycles, longitudinal audit trail) ---
    update_cycle_history_sheet(
        wb,
        meta,
        enriched_rows_for_history_df[
            [
                "Cycle_No",
                "Sample_No",
                "Report_Date",
                "Analyte",
                "Peer_Mean",
                "Your_Result",
                "%DEV",
                "SDI",
                "Target_Score",
                "TEa_or_TDPA(%)",
                "diff_from_mean",
                "Risk_Category_BaseOnly",
                "Risk_Category_Final",
                "Bias_Flag_Last3",
                "Trend_Flag_Last3",
                "Required_Action",
                "Comment",
            ]
        ]
    )

    # --- save workbook ---
    wb.save(out_path)

############################################
# 5. GUI LAUNCHER (file pickers for PDFs / template / output)
############################################
if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog, messagebox
    from pathlib import Path

    def run_gui():
        root = tk.Tk()
        root.withdraw()
        root.update_idletasks()

        # 1) Select one or more RIQAS PDFs
        pdf_paths = filedialog.askopenfilenames(
            title="Select RIQAS PDF file(s)",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not pdf_paths:
            messagebox.showinfo("Cancelled", "No PDFs selected.")
            return

        # 2) Select the Excel template
        template_path = filedialog.askopenfilename(
            title="Select Excel Template (.xlsx)",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not template_path:
            messagebox.showinfo("Cancelled", "No Excel template selected.")
            return

        # 3) Choose output folder
        out_dir = filedialog.askdirectory(
            title="Select output folder for generated .xlsx files"
        )
        if not out_dir:
            messagebox.showinfo("Cancelled", "No output folder selected.")
            return

        # Normalize paths
        pdf_list = [Path(p) for p in root.tk.splitlist(pdf_paths)]
        template = Path(template_path)
        out_dir = Path(out_dir)

        # 4) Process each selected PDF
        errors = []
        ok_count = 0
        for pdf in pdf_list:
            out_path = out_dir / f"{pdf.stem}-extracted.xlsx"
            try:
                process_riqas_pdf_into_workbook(
                    pdf_path=str(pdf),
                    xlsx_path=str(template),
                    out_path=str(out_path),
                )
                ok_count += 1
            except Exception as e:
                errors.append(f"{pdf.name}: {e!r}")

        # 5) Result dialog
        if errors:
            msg_lines = [
                f"Created {ok_count} / {len(pdf_list)} workbook(s) in:",
                str(out_dir),
                "",
                "Errors:",
                *errors
            ]
            messagebox.showerror("Done with errors", "\n".join(msg_lines[:50]))
        else:
            messagebox.showinfo(
                "Success",
                f"Created {ok_count} workbook(s) in:\n{out_dir}"
            )

    run_gui()