import React, { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { Upload, CheckCircle2, AlertTriangle, XCircle, HelpCircle, Download, Loader2, ArrowRight, RotateCcw } from "lucide-react";

// Colors matched from the HotTopics.ht screenshot
const PURPLE = "#2B0B5E";
const PURPLE_DEEP = "#1D0740";
const PINK = "#E31C79";
const WHITE = "#FFFFFF";
const INK = "#1B0E3D";
const LINE = "#E4E1EC";
const SLATE = "#6C6580";
const GREEN = "#2FA36B";
const YELLOW = "#F0A93A";

function parseFile(file) {
  return new Promise((resolve, reject) => {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv") {
      file.text().then((text) => {
        const result = Papa.parse(text, { header: true, skipEmptyLines: true });
        resolve(result.data);
      }).catch(reject);
    } else {
      file.arrayBuffer().then((buf) => {
        const wb = XLSX.read(buf, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json(sheet, { defval: "" }));
      }).catch(reject);
    }
  });
}

function norm(v) {
  return (v || "").toString().trim().toLowerCase();
}

const GlobalStyle = () => (
  <style>{`
    @keyframes fadeInUp {
      from { opacity: 0; transform: translateY(14px); }
      to { opacity: 1; transform: translateY(0); }
    }
    @keyframes floatBlob {
      0%, 100% { transform: translate(0, 0) scale(1); }
      50% { transform: translate(-10px, 12px) scale(1.04); }
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    @keyframes shimmer {
      0% { background-position: -200px 0; }
      100% { background-position: 200px 0; }
    }
    @keyframes popIn {
      0% { opacity: 0; transform: scale(0.9); }
      100% { opacity: 1; transform: scale(1); }
    }
    .stage-enter { animation: fadeInUp 0.45s ease both; }
    .blob { animation: floatBlob 7s ease-in-out infinite; }
    .btn-anim { transition: transform 0.15s ease, box-shadow 0.15s ease, background 0.15s ease; }
    .btn-anim:hover:not(:disabled) { transform: translateY(-2px); box-shadow: 0 6px 18px rgba(227,28,121,0.35); }
    .btn-anim:active:not(:disabled) { transform: translateY(0); }
    .filter-chip { transition: all 0.15s ease; }
    .filter-chip:hover { transform: translateY(-1px); }
    .row-anim { animation: fadeInUp 0.3s ease both; }
    .dropzone { transition: border-color 0.2s ease, background 0.2s ease, transform 0.2s ease; cursor: pointer; }
    .dropzone:hover { transform: translateY(-2px); }

    /* Keyboard accessibility: visible focus ring on every interactive element */
    button:focus-visible, select:focus-visible, [tabindex]:focus-visible, .dropzone:focus-visible {
      outline: 2px solid #E31C79;
      outline-offset: 2px;
      border-radius: 4px;
    }
    button, select { cursor: pointer; }
    button:disabled { cursor: not-allowed; }

    /* Respect users who've asked for less motion */
    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after {
        animation-duration: 0.001ms !important;
        animation-iteration-count: 1 !important;
        transition-duration: 0.001ms !important;
        scroll-behavior: auto !important;
      }
    }

    /* Responsive breakpoints: stack two-column layouts on small screens */
    @media (max-width: 640px) {
      .responsive-grid { grid-template-columns: 1fr !important; }
    }
  `}</style>
);

function Dropzone({ label, file, columns, onFile, sample }) {
  const inputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);

  return (
    <div
      className="dropzone"
      role="button"
      tabIndex={0}
      aria-label={`Upload ${label}`}
      onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        const f = e.dataTransfer.files[0];
        if (f) onFile(f);
      }}
      onClick={() => inputRef.current.click()}
      onKeyDown={(e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          inputRef.current.click();
        }
      }}
      style={{
        border: `1.5px dashed ${dragOver ? PINK : LINE}`,
        borderRadius: 10,
        padding: "28px 20px",
        textAlign: "center",
        cursor: "pointer",
        background: dragOver ? "rgba(227,28,121,0.06)" : WHITE,
        minHeight: 140,
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        gap: 8,
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".csv,.xlsx,.xls"
        style={{ display: "none" }}
        onChange={(e) => e.target.files[0] && onFile(e.target.files[0])}
      />
      {file ? (
        <>
          <CheckCircle2 size={22} color={GREEN} style={{ animation: "popIn 0.3s ease" }} />
          <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 13, color: INK, fontWeight: 700 }}>
            {file.name}
          </div>
          <div style={{ fontSize: 12, color: SLATE }}>
            {columns ? `${columns.length} columns · ${sample} rows` : "reading…"}
          </div>
          <div style={{ fontSize: 11, color: PINK, marginTop: 2, fontWeight: 600 }}>click to replace</div>
        </>
      ) : (
        <>
          <Upload size={20} color={SLATE} />
          <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12, letterSpacing: 0.5, color: INK, fontWeight: 700, textTransform: "uppercase" }}>
            {label}
          </div>
          <div style={{ fontSize: 12, color: SLATE }}>drop a .csv or .xlsx file, or click to browse</div>
        </>
      )}
    </div>
  );
}

function FieldSelect({ label, columns, value, onChange }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
      <label style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: 0.6, color: SLATE, fontFamily: "'JetBrains Mono', monospace", fontWeight: 600 }}>
        {label}
      </label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        style={{
          padding: "8px 10px",
          borderRadius: 6,
          border: `1px solid ${LINE}`,
          background: WHITE,
          fontSize: 13,
          color: INK,
          fontFamily: "inherit",
        }}
      >
        <option value="">— not used —</option>
        {columns.map((c) => (
          <option key={c} value={c}>{c}</option>
        ))}
      </select>
    </div>
  );
}

const STATUS_META = {
  changed: { label: "Changed", color: YELLOW, Icon: AlertTriangle },
  no_change: { label: "No change", color: GREEN, Icon: CheckCircle2 },
  not_found: { label: "Not found", color: PINK, Icon: XCircle },
  needs_manual_linkedin_url: { label: "Needs LinkedIn URL", color: SLATE, Icon: HelpCircle },
  error: { label: "Check failed", color: PINK, Icon: XCircle },
};

export default function App() {
  const [stage, setStage] = useState("upload");
  const [oldFile, setOldFile] = useState(null);
  const [newFile, setNewFile] = useState(null);
  const [oldRows, setOldRows] = useState(null);
  const [newRows, setNewRows] = useState(null);
  const [oldMap, setOldMap] = useState({ name: "", title: "", company: "", linkedin: "" });
  const [newMap, setNewMap] = useState({ title: "", company: "", linkedin: "" });
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const [results, setResults] = useState([]);
  const [filter, setFilter] = useState("all");
  const [error, setError] = useState("");

  const handleOldFile = useCallback(async (f) => {
    setOldFile(f); setError("");
    try {
      const rows = await parseFile(f);
      setOldRows(rows);
      if (rows.length) {
        const cols = Object.keys(rows[0]);
        setOldMap({
          name: cols.find((c) => /name/i.test(c)) || "",
          title: cols.find((c) => /title|role/i.test(c)) || "",
          company: cols.find((c) => /company|employer|org/i.test(c)) || "",
          linkedin: cols.find((c) => /linkedin/i.test(c)) || "",
        });
      }
    } catch (e) {
      setError("Couldn't read that file. Try exporting it as .csv or .xlsx.");
    }
  }, []);

  const handleNewFile = useCallback(async (f) => {
    setNewFile(f); setError("");
    try {
      const rows = await parseFile(f);
      setNewRows(rows);
      if (rows.length) {
        const cols = Object.keys(rows[0]);
        setNewMap({
          title: cols.find((c) => /title|role/i.test(c)) || "",
          company: cols.find((c) => /company|employer|org/i.test(c)) || "",
          linkedin: cols.find((c) => /linkedin/i.test(c)) || "",
        });
      }
    } catch (e) {
      setError("Couldn't read that file. Try exporting it as .csv or .xlsx.");
    }
  }, []);

  function goToMapping() {
    if (!oldRows || !newRows) return;
    setStage("mapping");
  }

  async function runComparison() {
    setError("");
    setStage("running");

    const newByLinkedin = new Map();
    newRows.forEach((r) => {
      const li = norm(r[newMap.linkedin]);
      if (li) newByLinkedin.set(li, r);
    });

    const merged = oldRows.map((r) => {
      const name = r[oldMap.name] || "";
      const linkedin_url = r[oldMap.linkedin] || "";
      const li = norm(linkedin_url);
      const match = li ? newByLinkedin.get(li) : null;
      return {
        name,
        old_title: r[oldMap.title] || "",
        old_company: r[oldMap.company] || "",
        linkedin_url,
        new_title: match ? (match[newMap.title] || "") : "",
        new_company: match ? (match[newMap.company] || "") : "",
        has_match: !!match,
      };
    });

    const needsCheck = merged.filter((c) => c.linkedin_url && c.has_match);
    const needsFlag = merged.filter((c) => !c.linkedin_url);
    const noMatchFound = merged.filter((c) => c.linkedin_url && !c.has_match);

    setProgress({ done: 0, total: needsCheck.length });

    const today = new Date().toISOString().slice(0, 10);
    const checked = [];
    const BATCH = 8;

    for (let i = 0; i < needsCheck.length; i += BATCH) {
      const batch = needsCheck.slice(i, i + BATCH);
      try {
        const response = await fetch("/api/compare", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ contacts: batch }),
        });
        const data = await response.json();
        const parsed = data.results || batch.map(() => ({ status: "error", resolved_title: "", resolved_company: "" }));
        batch.forEach((c, idx) => {
          const p = parsed[idx] || { status: "error", resolved_title: "", resolved_company: "" };
          checked.push({ ...c, status: p.status, resolved_title: p.resolved_title, resolved_company: p.resolved_company, last_verified: today });
        });
      } catch (e) {
        batch.forEach((c) => {
          checked.push({ ...c, status: "error", resolved_title: "", resolved_company: "", last_verified: today });
        });
      }
      setProgress({ done: Math.min(i + BATCH, needsCheck.length), total: needsCheck.length });
    }

    const flagged = needsFlag.map((c) => ({ ...c, status: "needs_manual_linkedin_url", resolved_title: "", resolved_company: "", last_verified: today }));
    const unmatched = noMatchFound.map((c) => ({ ...c, status: "not_found", resolved_title: "", resolved_company: "", last_verified: today }));

    setResults([...checked, ...flagged, ...unmatched]);
    setStage("results");
  }

  function exportCsv() {
    const headers = ["Name", "Old Title", "Old Company", "New Title", "New Company", "Status", "LinkedIn URL", "Last Verified"];
    const rows = results.map((r) => [
      r.name, r.old_title, r.old_company, r.resolved_title, r.resolved_company,
      STATUS_META[r.status]?.label || r.status, r.linkedin_url, r.last_verified,
    ]);
    const csv = [headers, ...rows].map((row) => row.map((v) => `"${(v || "").toString().replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `contact-refresh-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  function reset() {
    setStage("upload");
    setOldFile(null); setNewFile(null);
    setOldRows(null); setNewRows(null);
    setResults([]); setFilter("all"); setError("");
  }

  const counts = results.reduce((acc, r) => { acc[r.status] = (acc[r.status] || 0) + 1; return acc; }, {});
  const filtered = filter === "all" ? results : results.filter((r) => r.status === filter);

  return (
    <div style={{ fontFamily: "'Inter', -apple-system, sans-serif", background: "#F7F6FB", color: INK, minHeight: "100vh" }}>
      <GlobalStyle />

      {/* Top bar - matches HotTopics deep purple header with pink accent button */}
      <div style={{ background: PURPLE, color: WHITE, padding: "18px 24px", position: "relative", overflow: "hidden" }}>
        <div
          className="blob"
          style={{
            position: "absolute", right: -60, top: -60, width: 220, height: 220, borderRadius: "50%",
            background: `radial-gradient(circle, ${PINK}55, transparent 70%)`, pointerEvents: "none",
          }}
        />
        <div style={{ maxWidth: 1040, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", position: "relative" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 10, height: 10, borderRadius: "50%", background: PINK }} />
            <div style={{ fontWeight: 800, fontSize: 16, letterSpacing: -0.3 }}>Contact Refresh</div>
          </div>
          {stage === "results" && (
            <button className="btn-anim" onClick={reset} style={{
              display: "flex", alignItems: "center", gap: 6, background: "rgba(255,255,255,0.08)",
              border: "1px solid rgba(255,255,255,0.22)", borderRadius: 6, padding: "7px 14px",
              fontSize: 13, color: WHITE, cursor: "pointer", fontFamily: "inherit",
            }}>
              <RotateCcw size={14} /> New check
            </button>
          )}
        </div>
      </div>

      <div style={{ maxWidth: 1040, margin: "0 auto", padding: "36px 24px 60px" }}>
        {error && (
          <div style={{ background: "rgba(227,28,121,0.08)", border: `1px solid ${PINK}`, color: PINK, padding: "10px 14px", borderRadius: 6, fontSize: 13, marginBottom: 20 }}>
            {error}
          </div>
        )}

        {stage === "upload" && (
          <div key="upload" className="stage-enter">
            <h1 style={{ fontSize: 28, fontWeight: 800, letterSpacing: -0.5, margin: "0 0 8px", color: PURPLE }}>Check your contacts</h1>
            <p style={{ fontSize: 14, color: SLATE, marginBottom: 24, maxWidth: 620, lineHeight: 1.5 }}>
              Upload the old contact list and the freshly pulled data. Every record gets checked and logged.
            </p>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              <Dropzone label="Old data" file={oldFile} columns={oldRows && Object.keys(oldRows[0] || {})} sample={oldRows?.length} onFile={handleOldFile} />
              <Dropzone label="New data" file={newFile} columns={newRows && Object.keys(newRows[0] || {})} sample={newRows?.length} onFile={handleNewFile} />
            </div>
            <button
              className="btn-anim"
              disabled={!oldRows || !newRows}
              onClick={goToMapping}
              style={{
                marginTop: 24, display: "flex", alignItems: "center", gap: 8,
                background: (!oldRows || !newRows) ? LINE : PURPLE, color: (!oldRows || !newRows) ? SLATE : WHITE,
                border: "none", borderRadius: 8, padding: "11px 20px", fontSize: 13, fontWeight: 700,
                cursor: (!oldRows || !newRows) ? "not-allowed" : "pointer", fontFamily: "inherit",
              }}
            >
              Map columns <ArrowRight size={15} />
            </button>
          </div>
        )}

        {stage === "mapping" && (
          <div key="mapping" className="stage-enter">
            <h1 style={{ fontSize: 28, fontWeight: 800, letterSpacing: -0.5, margin: "0 0 8px", color: PURPLE }}>Match your columns</h1>
            <p style={{ fontSize: 14, color: SLATE, marginBottom: 24, maxWidth: 620, lineHeight: 1.5 }}>
              Contacts are matched between files by LinkedIn URL.
            </p>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 28 }}>
              <div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 11, letterSpacing: 0.6, textTransform: "uppercase", color: PINK, fontWeight: 700, marginBottom: 10 }}>Old data — {oldFile?.name}</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                  <FieldSelect label="Name" columns={Object.keys(oldRows[0] || {})} value={oldMap.name} onChange={(v) => setOldMap({ ...oldMap, name: v })} />
                  <FieldSelect label="Title" columns={Object.keys(oldRows[0] || {})} value={oldMap.title} onChange={(v) => setOldMap({ ...oldMap, title: v })} />
                  <FieldSelect label="Company" columns={Object.keys(oldRows[0] || {})} value={oldMap.company} onChange={(v) => setOldMap({ ...oldMap, company: v })} />
                  <FieldSelect label="LinkedIn URL" columns={Object.keys(oldRows[0] || {})} value={oldMap.linkedin} onChange={(v) => setOldMap({ ...oldMap, linkedin: v })} />
                </div>
              </div>
              <div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 11, letterSpacing: 0.6, textTransform: "uppercase", color: PINK, fontWeight: 700, marginBottom: 10 }}>New data — {newFile?.name}</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                  <FieldSelect label="Title" columns={Object.keys(newRows[0] || {})} value={newMap.title} onChange={(v) => setNewMap({ ...newMap, title: v })} />
                  <FieldSelect label="Company" columns={Object.keys(newRows[0] || {})} value={newMap.company} onChange={(v) => setNewMap({ ...newMap, company: v })} />
                  <FieldSelect label="LinkedIn URL" columns={Object.keys(newRows[0] || {})} value={newMap.linkedin} onChange={(v) => setNewMap({ ...newMap, linkedin: v })} />
                </div>
              </div>
            </div>
            <button
              className="btn-anim"
              disabled={!oldMap.linkedin || !newMap.linkedin}
              onClick={runComparison}
              style={{
                marginTop: 28, display: "flex", alignItems: "center", gap: 8,
                background: (!oldMap.linkedin || !newMap.linkedin) ? LINE : PINK,
                color: (!oldMap.linkedin || !newMap.linkedin) ? SLATE : WHITE,
                border: "none", borderRadius: 8, padding: "11px 20px", fontSize: 13, fontWeight: 700,
                cursor: (!oldMap.linkedin || !newMap.linkedin) ? "not-allowed" : "pointer", fontFamily: "inherit",
              }}
            >
              Run check <ArrowRight size={15} />
            </button>
          </div>
        )}

        {stage === "running" && (
          <div key="running" className="stage-enter" style={{ display: "flex", flexDirection: "column", alignItems: "center", padding: "80px 0", gap: 16 }}>
            <Loader2 size={28} color={PINK} style={{ animation: "spin 1s linear infinite" }} />
            <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 13, color: INK }}>
              Checking contact {progress.done} of {progress.total}…
            </div>
            <div style={{ width: 280, height: 6, background: LINE, borderRadius: 3, overflow: "hidden", position: "relative" }}>
              <div style={{
                width: `${progress.total ? (progress.done / progress.total) * 100 : 0}%`, height: "100%",
                background: `linear-gradient(90deg, ${PURPLE}, ${PINK})`, transition: "width 0.2s",
                backgroundSize: "200px 100%", animation: "shimmer 1.2s linear infinite",
              }} />
            </div>
          </div>
        )}

        {stage === "results" && (
          <div key="results" className="stage-enter">
            <div style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
              {["all", "changed", "no_change", "not_found", "needs_manual_linkedin_url", "error"].map((f) => {
                if (f !== "all" && !counts[f]) return null;
                const meta = STATUS_META[f];
                const active = filter === f;
                return (
                  <button
                    key={f}
                    className="filter-chip"
                    onClick={() => setFilter(f)}
                    style={{
                      display: "flex", alignItems: "center", gap: 6,
                      padding: "7px 12px", borderRadius: 6, fontSize: 12, fontWeight: 700,
                      border: `1px solid ${active ? PURPLE : LINE}`,
                      background: active ? PURPLE : WHITE,
                      color: active ? WHITE : INK,
                      cursor: "pointer", fontFamily: "'JetBrains Mono', monospace",
                    }}
                  >
                    {f === "all" ? `ALL (${results.length})` : `${meta.label.toUpperCase()} (${counts[f]})`}
                  </button>
                );
              })}
              <button className="btn-anim" onClick={exportCsv} style={{
                marginLeft: "auto", display: "flex", alignItems: "center", gap: 6,
                background: PINK, color: WHITE, border: "none", borderRadius: 6,
                padding: "7px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit",
              }}>
                <Download size={13} /> Export CSV
              </button>
            </div>

            <div style={{ border: `1px solid ${LINE}`, borderRadius: 10, overflow: "hidden", background: WHITE }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ background: PURPLE, color: WHITE }}>
                    {["Name", "Old", "New", "Status", "Verified"].map((h) => (
                      <th key={h} style={{ textAlign: "left", padding: "10px 14px", fontFamily: "'JetBrains Mono', monospace", fontSize: 11, letterSpacing: 0.6, textTransform: "uppercase", fontWeight: 700 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((r, i) => {
                    const meta = STATUS_META[r.status] || STATUS_META.error;
                    const Icon = meta.Icon;
                    return (
                      <tr key={i} className="row-anim" style={{ borderTop: `1px solid ${LINE}`, animationDelay: `${Math.min(i, 20) * 0.02}s` }}>
                        <td style={{ padding: "10px 14px", fontWeight: 700 }}>{r.name}</td>
                        <td style={{ padding: "10px 14px", color: SLATE }}>{r.old_title}{r.old_company ? `, ${r.old_company}` : ""}</td>
                        <td style={{ padding: "10px 14px" }}>
                          {r.status === "changed" ? `${r.resolved_title}, ${r.resolved_company}` : r.status === "no_change" ? "—" : ""}
                        </td>
                        <td style={{ padding: "10px 14px" }}>
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 5, color: meta.color, fontWeight: 700, fontSize: 12 }}>
                            <Icon size={13} /> {meta.label}
                          </span>
                        </td>
                        <td style={{ padding: "10px 14px", color: SLATE, fontFamily: "'JetBrains Mono', monospace", fontSize: 12 }}>{r.last_verified}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
