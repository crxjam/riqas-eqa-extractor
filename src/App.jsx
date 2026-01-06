import { useState } from "react";

export default function App() {
  const [pdfs, setPdfs] = useState([]);
  const [template, setTemplate] = useState(null);
  const [tea, setTea] = useState(null);
  const [status, setStatus] = useState("");
  const [loading, setLoading] = useState(false);

  async function handleSubmit() {
    if (!pdfs.length || !template || !tea) {
      alert("Please select PDFs, template, and TEA file");
      return;
    }

    const form = new FormData();
    pdfs.forEach((p) => form.append("pdfs", p));
    form.append("template", template);
    form.append("tea", tea);

    setLoading(true);
    setStatus("");

    try {
      const API = import.meta.env.VITE_API_BASE || "http://127.0.0.1:8000";
      const res = await fetch(`${API}/process`, {
        method: "POST",
        body: form,
      });

      if (!res.ok) {
        const txt = await res.text();
        throw new Error(txt);
      }

      const blob = await res.blob();
      const cd = res.headers.get("content-disposition") || "";
      const match = cd.match(/filename="?([^"]+)"?/i);
      const filename = match?.[1] || "RIQAS_EQA_Rolling_History.xlsx";

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.click();
      window.URL.revokeObjectURL(url);

      setStatus("✔ File generated successfully");
    } catch (err) {
      console.error(err);
      setStatus("❌ " + err.message);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{ padding: "2rem", fontFamily: "Arial, sans-serif" }}>
      <h1>RIQAS EQA Extractor</h1>

      <label>
        PDFs (multiple):
        <input
          type="file"
          accept=".pdf"
          multiple
          onChange={(e) => setPdfs(Array.from(e.target.files))}
        />
      </label>

      <br /><br />

      <label>
        Template (.xlsx):
        <input
          type="file"
          accept=".xlsx"
          onChange={(e) => setTemplate(e.target.files[0])}
        />
      </label>

      <br /><br />

      <label>
        TEA file (.xlsx / .csv):
        <input
          type="file"
          accept=".xlsx,.csv"
          onChange={(e) => setTea(e.target.files[0])}
        />
      </label>

      <br /><br />

      <button onClick={handleSubmit} disabled={loading}>
        {loading ? "Processing…" : "Process"}
      </button>

      <p>{status}</p>
    </div>
  );
}
