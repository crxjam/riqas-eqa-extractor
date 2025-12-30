function App() {
  return (
    <div style={{ padding: '2rem', fontFamily: 'Arial, sans-serif' }}>
      <h1>RIQAS EQA Extractor</h1>

      <p>
        Upload a RIQAS PDF report and extract analyte results automatically.
      </p>

      <input type="file" accept=".pdf" />

      <p style={{ marginTop: '2rem', color: '#666' }}>
        (PDF parsing logic coming next)
      </p>
    </div>
  )
}

export default App

