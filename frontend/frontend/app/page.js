"use client";

import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_BASE_URL || "https://visual-data-backend-260065454054.asia-south1.run.app";
const MAX_FILE_SIZE = 200 * 1024 * 1024;
const ACCEPTED_FILES = ".pdf,.png,.jpg,.jpeg,.webp,.bmp,.gif,.txt,.md,.xml,.xlsx,.xls,.docx";

function formatValue(value) {
  if (value === null || value === undefined || value === "") return "-";
  return String(value);
}

function fileSizeLabel(file) {
  if (!file) return "No file selected";
  const size = file.size;
  if (size < 1024) return `${file.name} • ${size} B`;
  if (size < 1024 * 1024) return `${file.name} • ${(size / 1024).toFixed(1)} KB`;
  return `${file.name} • ${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function downloadBlob(content, fileName, type) {
  const blob = new Blob([content], { type });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

function exportExcel(result) {
  if (!result) return;
  const workbook = XLSX.utils.book_new();
  const fieldRows = (result.key_value_pairs || []).map((pair) => ({
    label: pair.field_label || "",
    value: pair.normalized_value || pair.value || "",
    type: pair.value_type || "",
    confidence: pair.confidence || "",
    evidence: pair.source_text || "",
  }));

  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(fieldRows.length ? fieldRows : [{ message: "No extracted fields" }]),
    "Fields"
  );

  (result.tables || []).forEach((table, index) => {
    const rows = table.rows?.length ? table.rows : [{ message: "No rows" }];
    const safeName = `${index + 1}_${table.table_name || "Table"}`.slice(0, 31);
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), safeName);
  });

  const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  downloadBlob(buffer, "visual-data-export.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

function statusTone(status) {
  if (status === "success") return "success";
  if (status === "unreadable") return "warning";
  return "neutral";
}

function SummaryCards({ cards }) {
  if (!cards?.length) return null;
  return (
    <div className="summary-grid">
      {cards.map((card, index) => (
        <article key={`${card.label}-${index}`} className="summary-card">
          <span>{card.label}</span>
          <strong>{formatValue(card.value)}</strong>
        </article>
      ))}
    </div>
  );
}

function PairTable({ pairs }) {
  if (!pairs?.length) return null;
  return (
    <section className="section-card">
      <div className="section-head compact">
        <div>
          <h3>Extracted fields</h3>
          <p>Mapped directly from the file.</p>
        </div>
        <span className="count-pill">{pairs.length} fields</span>
      </div>
      <div className="table-wrap">
        <table className="data-table">
          <thead>
            <tr>
              <th>Label</th>
              <th>Value</th>
              <th>Type</th>
              <th>Confidence</th>
              <th>Evidence</th>
            </tr>
          </thead>
          <tbody>
            {pairs.map((pair, index) => (
              <tr key={`${pair.field_key || pair.field_label}-${index}`}>
                <td>{formatValue(pair.field_label)}</td>
                <td>{formatValue(pair.normalized_value || pair.value)}</td>
                <td>{formatValue(pair.value_type)}</td>
                <td>
                  <span className={`confidence-pill ${pair.confidence || "medium"}`}>{formatValue(pair.confidence || "medium")}</span>
                </td>
                <td>{formatValue(pair.source_text)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function DynamicTables({ tables }) {
  if (!tables?.length) return null;
  return (
    <div className="stack">
      {tables.map((table, index) => (
        <section className="section-card" key={`${table.table_name || "table"}-${index}`}>
          <div className="section-head compact">
            <div>
              <h3>{formatValue(table.table_name || `Table ${index + 1}`)}</h3>
              {table.description ? <p>{table.description}</p> : null}
            </div>
            <span className="count-pill">{table.rows?.length || 0} rows</span>
          </div>
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  {(table.columns || []).map((column, colIndex) => (
                    <th key={`${column}-${colIndex}`}>{formatValue(column)}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {(table.rows || []).map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {(table.columns || []).map((column, colIndex) => (
                      <td key={`${column}-${colIndex}`}>{formatValue(row?.[column])}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      ))}
    </div>
  );
}

export default function HomePage() {
  const inputRef = useRef(null);
  const [inputMode, setInputMode] = useState("file");
  const [extractionMode, setExtractionMode] = useState("smart");
  const [file, setFile] = useState(null);
  const [textTitle, setTextTitle] = useState("Pasted Text");
  const [textInput, setTextInput] = useState("");
  const [result, setResult] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [dragActive, setDragActive] = useState(false);

  const fileLabel = useMemo(() => fileSizeLabel(file), [file]);
  const tone = result ? statusTone(result.status) : "neutral";

  function pickFile(nextFile) {
    setFile(nextFile || null);
    setError("");
  }

  function openPicker() {
    inputRef.current?.click();
  }

  function onDrop(event) {
    event.preventDefault();
    setDragActive(false);
    const nextFile = event.dataTransfer?.files?.[0];
    if (nextFile) pickFile(nextFile);
  }

  async function handleExtract(event) {
    event.preventDefault();
    setLoading(true);
    setError("");
    setResult(null);

    try {
      if (inputMode === "text") {
        const cleaned = textInput.trim();
        if (!cleaned) throw new Error("Paste some text first.");
        const res = await fetch(`${API_BASE_URL}/extract-text`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: cleaned, title: textTitle.trim() || "Pasted Text", extraction_mode: extractionMode }),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data?.detail || data?.message || "Extraction failed.");
        setResult(data.result);
        return;
      }

      if (!file) throw new Error("Choose a file first.");
      if (file.size > MAX_FILE_SIZE) throw new Error("File exceeds 200 MB limit.");
      const formData = new FormData();
      formData.append("file", file);
      formData.append("extraction_mode", extractionMode);

      const res = await fetch(`${API_BASE_URL}/extract`, { method: "POST", body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data?.detail || data?.message || "Extraction failed.");
      setResult(data.result);
    } catch (err) {
      setError(err.message || "Something went wrong.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="page-shell">
      <header className="topbar">
        <div className="brand-mark">V</div>
        <div className="topbar-copy">
          <h1>Visual Data AI Assistant</h1>
          <p>Structured extraction for all kinds of data.</p>
        </div>
      </header>

      <main className="workspace">
        <aside className="panel control-panel">
          <form onSubmit={handleExtract} className="control-stack">
            <div className="toggle-row">
              <button type="button" className={`toggle-btn ${inputMode === "file" ? "active" : ""}`} onClick={() => setInputMode("file")}>
                File upload
              </button>
              <button type="button" className={`toggle-btn ${inputMode === "text" ? "active" : ""}`} onClick={() => setInputMode("text")}>
                Paste text
              </button>
            </div>

            <div className="field-block">
              <label className="field-label">Extraction mode</label>
              <select className="field-input field-select" value={extractionMode} onChange={(event) => setExtractionMode(event.target.value)}>
                <option value="smart">Smart extract</option>
                <option value="key_value">Key-value focus</option>
                <option value="table">Table focus</option>
              </select>
            </div>

            {inputMode === "file" ? (
              <div className="field-block">
                <input ref={inputRef} type="file" accept={ACCEPTED_FILES} className="hidden-file-input" onChange={(event) => pickFile(event.target.files?.[0] || null)} />
                <div
                  className={`dropzone ${dragActive ? "drag-active" : ""}`}
                  onDragOver={(event) => {
                    event.preventDefault();
                    setDragActive(true);
                  }}
                  onDragLeave={(event) => {
                    event.preventDefault();
                    setDragActive(false);
                  }}
                  onDrop={onDrop}
                >
                  <div className="drop-icon">📁</div>
                  <div className="drop-copy">
                    <strong>{file ? file.name : "Click or drag a file here"}</strong>
                    <span>{fileLabel}</span>
                    <small>PDF · Image · DOCX · XLSX · TXT · XML · up to 200 MB</small>
                  </div>
                  <button type="button" className="btn-secondary" onClick={openPicker}>
                    Choose file
                  </button>
                </div>
              </div>
            ) : (
              <>
                <div className="field-block">
                  <label className="field-label">Document title</label>
                  <input className="field-input" value={textTitle} onChange={(event) => setTextTitle(event.target.value)} placeholder="Pasted Text" />
                </div>
                <div className="field-block">
                  <label className="field-label">Paste content</label>
                  <textarea
                    className="field-input field-textarea"
                    value={textInput}
                    onChange={(event) => setTextInput(event.target.value)}
                    placeholder="Paste report text, marks, forms, list items, or notes..."
                    rows={10}
                  />
                </div>
              </>
            )}

            <button className="btn-primary" type="submit" disabled={loading}>
              {loading ? "Extracting..." : "Extract data"}
            </button>

            {error ? <div className="error-box">{error}</div> : null}

            <div className="info-card">
              <h3>Extraction modes</h3>
              <ul>
                <li>Smart Extract → Best for mixed and unstructured documents.</li>
                <li>Key-Value Focus → Best for forms, cards, and report fields.</li>
                <li>Table Focus → Best for rows, columns, and tabular data.</li>
              </ul>
            </div>
          </form>
        </aside>

        <section className="panel results-panel">
          <div className="results-head">
            <div>
              <h2>Results</h2>
              <p>Fields and tables appear here after extraction</p>
            </div>
            <button className="btn-secondary" onClick={() => exportExcel(result)} disabled={!result}>
              Export Excel
            </button>
          </div>

          <div className="results-body">
            {loading ? (
              <div className="empty-state loading-state">
                <div className="spinner" />
                <h3>Extracting data</h3>
                <p>Please wait while the file is being processed.</p>
              </div>
            ) : !result ? (
              <div className="empty-state">
                <div className="empty-icon">📊</div>
                <h3>No results yet</h3>
                <p>Upload a file or paste text to see extracted fields and tables here.</p>
              </div>
            ) : (
              <div className="stack">
                <div className={`status-banner ${tone}`}>
                  <strong>{formatValue(result.status).replaceAll("_", " ")}</strong>
                  <p>{formatValue(result.message)}</p>
                </div>

                <div className="chip-row">
                  {[result.meta?.file_name, result.meta?.file_kind, result.meta?.file_size, result.meta?.extraction_mode]
                    .filter(Boolean)
                    .map((item, index) => (
                      <span key={`${item}-${index}`} className="chip">{item}</span>
                    ))}
                </div>

                <section className="hero-card">
                  <div className="section-head">
                    <div>
                      <h3>{formatValue(result.title)}</h3>
                      <p>{formatValue(result.summary_text || "Structured output is ready.")}</p>
                    </div>
                  </div>
                  <SummaryCards cards={result.summary_cards} />
                </section>

                {result.warnings?.length ? (
                  <section className="section-card warning-card">
                    <div className="section-head">
                      <div>
                        <h3>Warnings</h3>
                        <p>Parsing notes and quality hints.</p>
                      </div>
                    </div>
                    <ul className="warning-list">
                      {result.warnings.map((warning, index) => (
                        <li key={`${warning}-${index}`}>{warning}</li>
                      ))}
                    </ul>
                  </section>
                ) : null}

                <PairTable pairs={result.key_value_pairs} />
                <DynamicTables tables={result.tables} />
              </div>
            )}
          </div>
        </section>
      </main>
    </div>
  );
}
