import { useState } from "react";

export default function App() {
  const [pdfs, setPdfs] = useState([]);
  const [template, setTemplate] = useState(null);
  const [tea, setTea] = useState(null);
  const [loading, setLoading] = useState(false);

  async function handleProcess() {
    if (!pdfs.length || !template || !tea) {
      alert("Please select PDFs, template, and TEA file.");
      return;
    }

    const form = new FormData();
    pdfs.forEach((p) => form.append("pdfs", p));
    form.append("template", template);
    form.append("tea", tea);

    setLoading(true);

    try {
      const API = import.meta.env.VITE_API_BASE || "http://127.0.0.1:8000";

      const res = await fetch(`${API}/process`, {
        method: "POST",
        body: form,
      });

      if (!res.ok) {
        const txt = await res.text();
        alert("Server error:\n" + txt);
        return;
      }

      // Download Excel (read body ONCE)
      const blob = await res.blob();

      // filename from header (needs CORS expose on backend)
      const cd = res.headers.get("content-disposition") || "";
      const match = cd.match(/filename\*?=(?:UTF-8''|")?([^";\n]+)"?/i);

      let filename = match
        ? decodeURIComponent(match[1])
        : "RIQAS_EQA_Rolling_History.xlsx";

      if (!filename.toLowerCase().endsWith(".xlsx")) filename += ".xlsx";

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error(err);
      alert("Failed:\n" + (err?.message || String(err)));
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{ padding: "2rem", fontFamily: "Arial" }}>
      <h1>RIQAS EQA Extractor</h1>

      <label>
        PDFs (multiple):
        <input
          type="file"
          multiple
          accept=".pdf"
          onChange={(e) => setPdfs([...e.target.files])}
        />
      </label>

      <br />
      <br />

      <label>
        Template (.xlsx):
        <input
          type="file"
          accept=".xlsx"
          onChange={(e) => setTemplate(e.target.files[0])}
        />
      </label>

      <br />
      <br />

      <label>
        TEA file:
        <input
          type="file"
          accept=".xlsx,.csv"
          onChange={(e) => setTea(e.target.files[0])}
        />
      </label>

      <br />
      <br />

      <button onClick={handleProcess} disabled={loading}>
        {loading ? "Processing..." : "Process"}
      </button>
    </div>
  );
}
