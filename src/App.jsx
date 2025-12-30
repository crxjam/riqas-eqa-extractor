import { useMemo, useState } from "react";

export default function App() {
  const [pdfs, setPdfs] = useState([]);
  const [template, setTemplate] = useState(null);
  const [tea, setTea] = useState(null);

  const [status, setStatus] = useState("");
  const [busy, setBusy] = useState(false);

  const canRun = useMemo(() => {
    return pdfs.length > 0 && !!template && !!tea && !busy;
  }, [pdfs, template, tea, busy]);

  async function run() {
    if (!canRun) return;

    setBusy(true);
    setStatus(`Uploading ${pdfs.length} PDF(s)…`);

    try {
      const form = new FormData();

      // Names MUST match your FastAPI signature: pdfs, template, tea
      for (const f of pdfs) form.append("pdfs", f);
      form.append("template", template);
      form.append("tea", tea);

      const res = await fetch("http://127.0.0.1:8000/process", {
        method: "POST",
        body: form,
      });

      if (!res.ok) {
        const txt = await res.text();
        throw new Error(`Server error ${res.status}: ${txt}`);
      }

      // If server returns JSON, show it. Otherwise download the file.
      const ct = res.headers.get("content-type") || "";

      if (ct.includes("application/json")) {
        const data = await res.json();
        setStatus(`Done ✅\n${JSON.stringify(data, null, 2)}`);
        return;
      }

      const blob = await res.blob();
      const cd = res.headers.get("content-disposition") || "";
      const match = cd.match(/filename="?([^"]+)"?/i);
      const filename = match?.[1] || (pdfs.length > 1 ? "riqas_outputs.zip" : "riqas_output.xlsx");

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      setStatus("Done ✅ Download started.");
    } catch (e) {
      setStatus(e?.message || "Failed");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={{ maxWidth: 880, margin: "40px auto", fontFamily: "system-ui" }}>
      <h1 style={{ fontSize: 40, marginBottom: 6 }}>RIQAS EQA Extractor</h1>
      <p style={{ opacity: 0.8, marginTop: 0 }}>
        Upload one or more RIQAS PDFs + your template + TEA file, then process.
      </p>

      <div style={{ padding: 16, border: "1px solid #ddd", borderRadius: 12 }}>
        <div style={{ display: "grid", gap: 14 }}>
          <div>
            <label style={{ display: "block", fontWeight: 600, marginBottom: 6 }}>
              PDFs (you can select multiple)
            </label>
            <input
              type="file"
              accept="application/pdf"
              multiple
              onChange={(e) => setPdfs(Array.from(e.target.files || []))}
            />
            <div style={{ marginTop: 6, opacity: 0.75 }}>
              Selected: {pdfs.length ? pdfs.map((f) => f.name).join(", ") : "none"}
            </div>
          </div>

          <div>
            <label style={{ display: "block", fontWeight: 600, marginBottom: 6 }}>
              Template file
            </label>
            <input
              type="file"
              onChange={(e) => setTemplate(e.target.files?.[0] || null)}
            />
            <div style={{ marginTop: 6, opacity: 0.75 }}>
              Selected: {template ? template.name : "none"}
            </div>
          </div>

          <div>
            <label style={{ display: "block", fontWeight: 600, marginBottom: 6 }}>
              TEA file
            </label>
            <input
              type="file"
              onChange={(e) => setTea(e.target.files?.[0] || null)}
            />
            <div style={{ marginTop: 6, opacity: 0.75 }}>
              Selected: {tea ? tea.name : "none"}
            </div>
          </div>

          <div>
            <button
              onClick={run}
              disabled={!canRun}
              style={{
                padding: "10px 14px",
                borderRadius: 10,
                border: "1px solid #333",
                background: canRun ? "#fff" : "#eee",
                cursor: canRun ? "pointer" : "not-allowed",
              }}
            >
              {busy ? "Processing…" : "Process"}
            </button>

            {!canRun && (
              <div style={{ marginTop: 10, opacity: 0.7, fontSize: 13 }}>
                Required: at least 1 PDF + template + TEA.
              </div>
            )}
          </div>

          {status && (
            <pre style={{ margin: 0, padding: 12, background: "#f7f7f7", borderRadius: 10, whiteSpace: "pre-wrap" }}>
              {status}
            </pre>
          )}
        </div>
      </div>

      <p style={{ marginTop: 14, opacity: 0.6, fontSize: 13 }}>
        Backend must be running on <code>http://127.0.0.1:8000</code>.
      </p>
    </div>
  );
}
