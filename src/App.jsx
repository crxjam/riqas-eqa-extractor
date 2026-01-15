import { useMemo, useState } from "react";

const API_BASE = "https://riqas-eqa-extractor.onrender.com";

export default function App() {
  const [pdfs, setPdfs] = useState([]);
  const [template, setTemplate] = useState(null);
  const [tea, setTea] = useState(null);
  const [status, setStatus] = useState("");
  const [busy, setBusy] = useState(false);

  const canSubmit = useMemo(
    () => pdfs.length > 0 && template && tea && !busy,
    [pdfs, template, tea, busy]
  );

  async function upload(e) {
    e.preventDefault();

    try {
      setBusy(true);
      setStatus("Uploading and processing…");

      const fd = new FormData();
      pdfs.forEach((f) => fd.append("pdfs", f));
      fd.append("template", template);
      fd.append("tea", tea);

      const res = await fetch(`${API_BASE}/process`, {
        method: "POST",
        body: fd,
      });

      if (!res.ok) {
        throw new Error(await res.text());
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = "EQA_Rolling_History.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();

      window.URL.revokeObjectURL(url);
      setStatus("Done. Workbook downloaded.");
    } catch (err) {
      setStatus(String(err));
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={{ padding: 20, fontFamily: "system-ui, sans-serif" }}>
      <div
        style={{
          maxWidth: 820,
          margin: "40px auto",
          padding: 24,
          borderRadius: 14,
          border: "1px solid rgba(0,0,0,0.1)",
          boxShadow: "0 10px 25px rgba(0,0,0,0.08)",
        }}
      >
        <div style={{ display: "flex", gap: 10, alignItems: "baseline" }}>
          <h1 style={{ margin: 0 }}>EQA Extractor</h1>
          <span style={{ fontStyle: "italic", opacity: 0.65 }}>
            Developed by James Croxford
          </span>
        </div>

        <p style={{ opacity: 0.85 }}>
          Upload EQA PDF reports, your Excel template, and your Internal TEa
          (Total Allowable Error) table. A single rolling workbook will be returned.
        </p>

        <form onSubmit={upload}>
          <label><b>1) EQA PDF reports</b></label><br />
          <input type="file" accept="application/pdf" multiple
            onChange={(e) => setPdfs(Array.from(e.target.files || []))}
          />
          <p style={{ fontSize: 13, opacity: 0.7 }}>
            Drag and drop one or more EQA PDFs. Multiple PDFs are processed in date order.
          </p>

          <label><b>2) Excel template (.xlsx)</b></label><br />
          <input type="file" accept=".xlsx"
            onChange={(e) => setTemplate(e.target.files[0])}
          />
          <p style={{ fontSize: 13, opacity: 0.7 }}>
            The EQA template containing sheets such as Header Information,
            Result Summary, and Cycle_History.
          </p>

          <label><b>3) Internal TEa / TAE table (.xlsx or .csv)</b></label><br />
          <input type="file" accept=".xlsx,.csv"
            onChange={(e) => setTea(e.target.files[0])}
          />
          <p style={{ fontSize: 13, opacity: 0.7 }}>
            Must include columns for analyte name and Internal TEa values.
          </p>

          <button
            type="submit"
            disabled={!canSubmit}
            style={{
              marginTop: 16,
              padding: "10px 14px",
              fontWeight: 700,
              opacity: canSubmit ? 1 : 0.5,
            }}
          >
            {busy ? "Processing…" : "Process & Download"}
          </button>
        </form>

        {status && <p style={{ marginTop: 12 }}>{status}</p>}
      </div>
    </div>
  );
}

