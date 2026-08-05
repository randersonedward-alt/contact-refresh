import React, { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { Upload, CheckCircle2, AlertTriangle, XCircle, HelpCircle, Download, Loader2, ArrowRight, RotateCcw } from "lucide-react";

// Colors matched from the HotTopics.ht screenshot
const PURPLE = "#2B0B5E";
const PURPLE_DEEP = "#1D0740";
const PURPLE_BRIGHT = "#4B21A6";
const PINK = "#E31C79";
const PINK_DEEP = "#B3115C";
const WHITE = "#FFFFFF";
const INK = "#1B0E3D";
const LINE = "#E4E1EC";
const SLATE = "#6C6580";
const GREEN = "#2FA36B";
const YELLOW = "#F0A93A";
const DISPLAY = "'Space Grotesk', 'Inter', -apple-system, sans-serif";
const MONO = "'JetBrains Mono', monospace";
const GRAD_BRAND = `linear-gradient(135deg, ${PURPLE_BRIGHT}, ${PURPLE_DEEP})`;
const GRAD_ACCENT = `linear-gradient(135deg, ${PINK}, ${PINK_DEEP})`;

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

// LinkedIn URLs are often saved inconsistently (http vs https, www vs no www,
// trailing slash, tracking params). This strips all of that down to just the
// meaningful part (e.g. "in/johndoe") so old vs new data actually matches.
function normUrl(v) {
  let s = (v || "").toString().trim().toLowerCase();
  if (!s) return "";
  s = s.replace(/^https?:\/\//, "");
  s = s.replace(/^www\./, "");
  s = s.split("?")[0];
  s = s.replace(/\/+$/, "");
  s = s.replace(/^linkedin\.com\//, "");
  return s;
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
    @keyframes meshDrift {
      0%, 100% { transform: translate(0, 0) rotate(0deg); }
      50% { transform: translate(-2%, 2%) rotate(3deg); }
    }
    @keyframes cardIn {
      from { opacity: 0; transform: translateY(20px) scale(0.99); }
      to { opacity: 1; transform: translateY(0) scale(1); }
    }
    .stage-enter { animation: fadeInUp 0.45s ease both; }
    .blob { animation: floatBlob 7s ease-in-out infinite; }
    .mesh { animation: meshDrift 14s ease-in-out infinite; }
    .main-card { animation: cardIn 0.5s cubic-bezier(0.16, 1, 0.3, 1) both; }
    .btn-anim { transition: transform 0.15s ease, box-shadow 0.2s ease, filter 0.15s ease; }
    .btn-anim:hover:not(:disabled) { transform: translateY(-2px); box-shadow: 0 10px 24px rgba(227,28,121,0.32); filter: brightness(1.05); }
    .btn-anim:active:not(:disabled) { transform: translateY(0); filter: brightness(0.98); }
    .filter-chip { transition: all 0.18s ease; }
    .filter-chip:hover { transform: translateY(-1px); box-shadow: 0 4px 10px rgba(43,11,94,0.12); }
    .row-anim { animation: fadeInUp 0.3s ease both; }
    .row-hover:hover { background: rgba(43,11,94,0.03) !important; }
    .dropzone { transition: border-color 0.2s ease, background 0.2s ease, transform 0.2s ease, box-shadow 0.2s ease; cursor: pointer; }
    .dropzone:hover { transform: translateY(-3px); box-shadow: 0 12px 28px rgba(43,11,94,0.1); }
    .stat-card { transition: transform 0.2s ease, box-shadow 0.2s ease; }
    .stat-card:hover { transform: translateY(-2px); box-shadow: 0 10px 22px rgba(43,11,94,0.1); }

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
        borderRadius: 14,
        padding: "30px 20px",
        textAlign: "center",
        cursor: "pointer",
        background: dragOver ? "rgba(227,28,121,0.06)" : file ? "#FBFAFD" : WHITE,
        minHeight: 150,
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        gap: 9,
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
          <div style={{
            width: 38, height: 38, borderRadius: "50%", background: `${GREEN}18`,
            display: "flex", alignItems: "center", justifyContent: "center", animation: "popIn 0.3s ease",
          }}>
            <CheckCircle2 size={19} color={GREEN} />
          </div>
          <div style={{ fontFamily: MONO, fontSize: 13, color: INK, fontWeight: 700 }}>
            {file.name}
          </div>
          <div style={{ fontSize: 12, color: SLATE }}>
            {columns ? `${columns.length} columns · ${sample} rows` : "reading…"}
          </div>
          <div style={{ fontSize: 11, color: PINK, marginTop: 2, fontWeight: 600 }}>click to replace</div>
        </>
      ) : (
        <>
          <div style={{
            width: 38, height: 38, borderRadius: "50%", background: "#F1EFF8",
            display: "flex", alignItems: "center", justifyContent: "center",
          }}>
            <Upload size={17} color={PURPLE_BRIGHT} />
          </div>
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
  duplicate_name: { label: "Same name found — check manually", color: SLATE, Icon: HelpCircle },
  error: { label: "Check failed", color: PINK, Icon: XCircle },
};

export default function App() {
  const [stage, setStage] = useState("upload");
  const [oldFile, setOldFile] = useState(null);
  const [newFile, setNewFile] = useState(null);
  const [oldRows, setOldRows] = useState(null);
  const [newRows, setNewRows] = useState(null);
  const [oldMap, setOldMap] = useState({ recordId: "", firstName: "", lastName: "", title: "", company: "", linkedin: "" });
  const [newMap, setNewMap] = useState({ firstName: "", lastName: "", title: "", company: "", linkedin: "" });
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
          recordId: cols.find((c) => /hubspot.*id|record.*id|contact.*id/i.test(c)) || "",
          firstName: cols.find((c) => /^first[\s_-]*name$/i.test(c)) || cols.find((c) => /first/i.test(c) && /name/i.test(c)) || "",
          lastName: cols.find((c) => /^last[\s_-]*name$/i.test(c)) || cols.find((c) => /last|surname/i.test(c) && /name/i.test(c)) || "",
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
          firstName: cols.find((c) => /^first[\s_-]*name$/i.test(c)) || cols.find((c) => /first/i.test(c) && /name/i.test(c)) || "",
          lastName: cols.find((c) => /^last[\s_-]*name$/i.test(c)) || cols.find((c) => /last|surname/i.test(c) && /name/i.test(c)) || "",
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

    // Group new-file records by normalized name, so we can tell when a name
    // is unique (safe to match) vs. shared by more than one person (ambiguous).
    // Group new-file records by normalized full name, so we can tell when a
    // name is unique (safe to match) vs. shared by more than one person (ambiguous).
    const newByLinkedin = new Map();
    const newByName = new Map();
    newRows.forEach((r) => {
      const li = normUrl(r[newMap.linkedin]);
      if (li) newByLinkedin.set(li, r);
      const nm = norm(`${r[newMap.firstName] || ""} ${r[newMap.lastName] || ""}`);
      if (nm) {
        if (!newByName.has(nm)) newByName.set(nm, []);
        newByName.get(nm).push(r);
      }
    });

    const merged = oldRows.map((r) => {
      const first_name = r[oldMap.firstName] || "";
      const last_name = r[oldMap.lastName] || "";
      const record_id = oldMap.recordId ? (r[oldMap.recordId] || "") : "";
      const linkedin_url = r[oldMap.linkedin] || "";
      const li = normUrl(linkedin_url);
      const nm = norm(`${first_name} ${last_name}`);

      let match = null;
      let ambiguous = false;

      // LinkedIn URL is an unambiguous identifier when present, so it wins
      // even if the name also happens to collide with someone else.
      if (li && newByLinkedin.has(li)) {
        match = newByLinkedin.get(li);
      } else if (nm && newByName.has(nm)) {
        const candidates = newByName.get(nm);
        if (candidates.length === 1) {
          match = candidates[0];
        } else {
          ambiguous = true; // more than one person shares this name — don't guess
        }
      }

      return {
        record_id,
        first_name,
        last_name,
        name: `${first_name} ${last_name}`.trim(),
        old_title: r[oldMap.title] || "",
        old_company: r[oldMap.company] || "",
        linkedin_url,
        new_title: match ? (match[newMap.title] || "") : "",
        new_company: match ? (match[newMap.company] || "") : "",
        has_match: !!match,
        ambiguous,
      };
    });

    const needsCheck = merged.filter((c) => c.has_match);
    const duplicateNames = merged.filter((c) => c.ambiguous);
    const noMatchFound = merged.filter((c) => !c.has_match && !c.ambiguous);

    setProgress({ done: 0, total: needsCheck.length });

    const today = new Date().toISOString().slice(0, 10);

    // Direct comparison — no external API needed. A contact is "changed" if
    // their title or company text differs from what's on file (case/spacing
    // insensitive); otherwise they're "no change".
    const checked = needsCheck.map((c) => {
      const titleSame = norm(c.old_title) === norm(c.new_title);
      const companySame = norm(c.old_company) === norm(c.new_company);
      const status = (titleSame && companySame) ? "no_change" : "changed";
      return { ...c, status, resolved_title: c.new_title, resolved_company: c.new_company, last_verified: today };
    });
    setProgress({ done: needsCheck.length, total: needsCheck.length });

    const duplicates = duplicateNames.map((c) => ({ ...c, status: "duplicate_name", resolved_title: "", resolved_company: "", last_verified: today }));
    const unmatched = noMatchFound.map((c) => ({ ...c, status: "not_found", resolved_title: "", resolved_company: "", last_verified: today }));

    setResults([...checked, ...duplicates, ...unmatched]);
    setStage("results");
  }

  function exportWorkbook() {
    const toRow = (r) => ({
      "Record ID": r.record_id,
      "First Name": r.first_name,
      "Last Name": r.last_name,
      "Old Title": r.old_title,
      "Old Company": r.old_company,
      "New Title": r.status === "changed" ? r.resolved_title : "",
      "New Company": r.status === "changed" ? r.resolved_company : "",
      "Status": STATUS_META[r.status]?.label || r.status,
      "LinkedIn URL": r.linkedin_url,
      "Last Verified": r.last_verified,
    });

    const allSheet = XLSX.utils.json_to_sheet(results.map(toRow));
    const changesSheet = XLSX.utils.json_to_sheet(results.filter((r) => r.status === "changed").map(toRow));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, allSheet, "All contacts");
    XLSX.utils.book_append_sheet(wb, changesSheet, "Changes only");

    XLSX.writeFile(wb, `contact-refresh-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function exportNotFound() {
    const toRow = (r) => ({
      "Record ID": r.record_id,
      "First Name": r.first_name,
      "Last Name": r.last_name,
      "Old Title": r.old_title,
      "Old Company": r.old_company,
      "LinkedIn URL": r.linkedin_url,
      "Last Verified": r.last_verified,
    });

    const notFoundSheet = XLSX.utils.json_to_sheet(results.filter((r) => r.status === "not_found").map(toRow));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, notFoundSheet, "Not found");

    XLSX.writeFile(wb, `contact-refresh-not-found-${new Date().toISOString().slice(0, 10)}.xlsx`);
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
    <div style={{
      fontFamily: "'Inter', -apple-system, sans-serif", color: INK, minHeight: "100vh",
      background: `
        radial-gradient(1100px 500px at 12% -10%, rgba(75,33,166,0.10), transparent 60%),
        radial-gradient(900px 500px at 100% 0%, rgba(227,28,121,0.08), transparent 55%),
        #F6F4FB
      `,
    }}>
      <GlobalStyle />

      {/* Top bar - deep purple gradient header with drifting pink glow, matches HotTopics brand */}
      <div style={{ background: GRAD_BRAND, color: WHITE, padding: "22px 24px 64px", position: "relative", overflow: "hidden" }}>
        <div
          className="blob mesh"
          style={{
            position: "absolute", right: -80, top: -100, width: 320, height: 320, borderRadius: "50%",
            background: `radial-gradient(circle, ${PINK}4d, transparent 70%)`, pointerEvents: "none",
          }}
        />
        <div
          className="mesh"
          style={{
            position: "absolute", left: "20%", bottom: -140, width: 260, height: 260, borderRadius: "50%",
            background: `radial-gradient(circle, #ffffff14, transparent 70%)`, pointerEvents: "none", animationDelay: "-6s",
          }}
        />
        <div style={{ maxWidth: 1040, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", position: "relative" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{
              width: 30, height: 30, borderRadius: 9, background: GRAD_ACCENT,
              display: "flex", alignItems: "center", justifyContent: "center",
              boxShadow: "0 4px 14px rgba(227,28,121,0.45)",
            }}>
              <div style={{ width: 8, height: 8, borderRadius: "50%", background: WHITE }} />
            </div>
            <div style={{ fontFamily: DISPLAY, fontWeight: 700, fontSize: 17, letterSpacing: -0.2 }}>Contact Refresh</div>
          </div>
          {stage === "results" && (
            <button className="btn-anim" onClick={reset} style={{
              display: "flex", alignItems: "center", gap: 6, background: "rgba(255,255,255,0.1)",
              border: "1px solid rgba(255,255,255,0.24)", borderRadius: 7, padding: "7px 14px",
              fontSize: 13, color: WHITE, cursor: "pointer", fontFamily: "inherit", backdropFilter: "blur(6px)",
            }}>
              <RotateCcw size={14} /> New check
            </button>
          )}
        </div>
        <div style={{ maxWidth: 1040, margin: "14px auto 0", position: "relative", fontSize: 13.5, color: "rgba(255,255,255,0.72)" }}>
          Automated CRM verification for your C-suite contact list.
        </div>
      </div>

      <div style={{ maxWidth: 1040, margin: "-38px auto 0", padding: "0 24px 60px", position: "relative" }}>
        <div className="main-card" style={{
          background: WHITE, borderRadius: 18, border: `1px solid ${LINE}`,
          boxShadow: "0 20px 50px rgba(43,11,94,0.12), 0 2px 8px rgba(43,11,94,0.06)",
          padding: "32px 32px 40px",
        }}>
        {error && (
          <div style={{ background: "rgba(227,28,121,0.08)", border: `1px solid ${PINK}`, color: PINK, padding: "10px 14px", borderRadius: 6, fontSize: 13, marginBottom: 20 }}>
            {error}
          </div>
        )}

        {stage === "upload" && (
          <div key="upload" className="stage-enter">
            <h1 style={{ fontFamily: DISPLAY, fontSize: 30, fontWeight: 700, letterSpacing: -0.6, margin: "0 0 8px", color: PURPLE }}>Check your contacts</h1>
            <p style={{ fontSize: 14, color: SLATE, marginBottom: 28, maxWidth: 620, lineHeight: 1.6 }}>
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
                marginTop: 26, display: "flex", alignItems: "center", gap: 8,
                background: (!oldRows || !newRows) ? LINE : GRAD_BRAND, color: (!oldRows || !newRows) ? SLATE : WHITE,
                border: "none", borderRadius: 9, padding: "12px 22px", fontSize: 13.5, fontWeight: 700,
                boxShadow: (!oldRows || !newRows) ? "none" : "0 6px 16px rgba(43,11,94,0.28)",
                cursor: (!oldRows || !newRows) ? "not-allowed" : "pointer", fontFamily: "inherit",
              }}
            >
              Map columns <ArrowRight size={15} />
            </button>
          </div>
        )}

        {stage === "mapping" && (
          <div key="mapping" className="stage-enter">
            <h1 style={{ fontFamily: DISPLAY, fontSize: 30, fontWeight: 700, letterSpacing: -0.6, margin: "0 0 8px", color: PURPLE }}>Match your columns</h1>
            <p style={{ fontSize: 14, color: SLATE, marginBottom: 28, maxWidth: 620, lineHeight: 1.6 }}>
              Contacts are matched by name. If a LinkedIn URL is mapped on both sides, it's used to break ties when two people share a name.
            </p>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 28 }}>
              <div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 11, letterSpacing: 0.6, textTransform: "uppercase", color: PINK, fontWeight: 700, marginBottom: 10 }}>Old data — {oldFile?.name}</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                  <FieldSelect label="Record ID" columns={Object.keys(oldRows[0] || {})} value={oldMap.recordId} onChange={(v) => setOldMap({ ...oldMap, recordId: v })} />
                  <FieldSelect label="First Name" columns={Object.keys(oldRows[0] || {})} value={oldMap.firstName} onChange={(v) => setOldMap({ ...oldMap, firstName: v })} />
                  <FieldSelect label="Last Name" columns={Object.keys(oldRows[0] || {})} value={oldMap.lastName} onChange={(v) => setOldMap({ ...oldMap, lastName: v })} />
                  <FieldSelect label="Title" columns={Object.keys(oldRows[0] || {})} value={oldMap.title} onChange={(v) => setOldMap({ ...oldMap, title: v })} />
                  <FieldSelect label="Company" columns={Object.keys(oldRows[0] || {})} value={oldMap.company} onChange={(v) => setOldMap({ ...oldMap, company: v })} />
                  <FieldSelect label="LinkedIn URL" columns={Object.keys(oldRows[0] || {})} value={oldMap.linkedin} onChange={(v) => setOldMap({ ...oldMap, linkedin: v })} />
                </div>
              </div>
              <div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 11, letterSpacing: 0.6, textTransform: "uppercase", color: PINK, fontWeight: 700, marginBottom: 10 }}>New data — {newFile?.name}</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                  <FieldSelect label="First Name" columns={Object.keys(newRows[0] || {})} value={newMap.firstName} onChange={(v) => setNewMap({ ...newMap, firstName: v })} />
                  <FieldSelect label="Last Name" columns={Object.keys(newRows[0] || {})} value={newMap.lastName} onChange={(v) => setNewMap({ ...newMap, lastName: v })} />
                  <FieldSelect label="Title" columns={Object.keys(newRows[0] || {})} value={newMap.title} onChange={(v) => setNewMap({ ...newMap, title: v })} />
                  <FieldSelect label="Company" columns={Object.keys(newRows[0] || {})} value={newMap.company} onChange={(v) => setNewMap({ ...newMap, company: v })} />
                  <FieldSelect label="LinkedIn URL" columns={Object.keys(newRows[0] || {})} value={newMap.linkedin} onChange={(v) => setNewMap({ ...newMap, linkedin: v })} />
                </div>
              </div>
            </div>
            <button
              className="btn-anim"
              disabled={!oldMap.firstName || !newMap.firstName}
              onClick={runComparison}
              style={{
                marginTop: 30, display: "flex", alignItems: "center", gap: 8,
                background: (!oldMap.firstName || !newMap.firstName) ? LINE : GRAD_ACCENT,
                color: (!oldMap.firstName || !newMap.firstName) ? SLATE : WHITE,
                border: "none", borderRadius: 9, padding: "12px 22px", fontSize: 13.5, fontWeight: 700,
                boxShadow: (!oldMap.firstName || !newMap.firstName) ? "none" : "0 6px 16px rgba(227,28,121,0.32)",
                cursor: (!oldMap.firstName || !newMap.firstName) ? "not-allowed" : "pointer", fontFamily: "inherit",
              }}
            >
              Run check <ArrowRight size={15} />
            </button>
          </div>
        )}

        {stage === "running" && (
          <div key="running" className="stage-enter" style={{ display: "flex", flexDirection: "column", alignItems: "center", padding: "88px 0", gap: 18 }}>
            <div style={{
              width: 54, height: 54, borderRadius: "50%", background: GRAD_BRAND,
              display: "flex", alignItems: "center", justifyContent: "center",
              boxShadow: "0 10px 26px rgba(43,11,94,0.28)",
            }}>
              <Loader2 size={24} color={WHITE} style={{ animation: "spin 1s linear infinite" }} />
            </div>
            <div style={{ fontFamily: DISPLAY, fontSize: 15, fontWeight: 600, color: INK }}>
              Checking contact {progress.done} of {progress.total}…
            </div>
            <div style={{ width: 300, height: 7, background: LINE, borderRadius: 4, overflow: "hidden", position: "relative" }}>
              <div style={{
                width: `${progress.total ? (progress.done / progress.total) * 100 : 0}%`, height: "100%",
                background: `linear-gradient(90deg, ${PURPLE_BRIGHT}, ${PINK})`, transition: "width 0.2s",
                backgroundSize: "200px 100%", animation: "shimmer 1.2s linear infinite", borderRadius: 4,
              }} />
            </div>
          </div>
        )}

        {stage === "results" && (
          <div key="results" className="stage-enter">
            <h1 style={{ fontFamily: DISPLAY, fontSize: 26, fontWeight: 700, letterSpacing: -0.5, margin: "0 0 20px", color: PURPLE }}>Results</h1>

            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, marginBottom: 24 }}>
              {["changed", "no_change", "not_found", "duplicate_name"].map((f) => {
                const meta = STATUS_META[f];
                if (!meta) return null;
                const Icon = meta.Icon;
                return (
                  <div key={f} className="stat-card" style={{
                    border: `1px solid ${LINE}`, borderRadius: 12, padding: "14px 14px 12px",
                    background: WHITE, display: "flex", flexDirection: "column", gap: 8,
                  }}>
                    <Icon size={16} color={meta.color} />
                    <div style={{ fontFamily: DISPLAY, fontSize: 22, fontWeight: 700, color: INK, lineHeight: 1 }}>{counts[f] || 0}</div>
                    <div style={{ fontSize: 11.5, color: SLATE, fontWeight: 600, lineHeight: 1.3 }}>{meta.label}</div>
                  </div>
                );
              })}
            </div>

            <div style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap", alignItems: "center" }}>
              {["all", "changed", "no_change", "not_found", "duplicate_name", "error"].map((f) => {
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
                      padding: "7px 13px", borderRadius: 20, fontSize: 12, fontWeight: 700,
                      border: `1px solid ${active ? "transparent" : LINE}`,
                      background: active ? GRAD_BRAND : WHITE,
                      color: active ? WHITE : INK,
                      cursor: "pointer", fontFamily: MONO,
                    }}
                  >
                    {f === "all" ? `ALL (${results.length})` : `${meta.label.toUpperCase()} (${counts[f]})`}
                  </button>
                );
              })}
              {counts.not_found > 0 && (
                <button className="btn-anim" onClick={exportNotFound} style={{
                  marginLeft: "auto", display: "flex", alignItems: "center", gap: 6,
                  background: WHITE, color: PINK, border: `1.5px solid ${PINK}`, borderRadius: 20,
                  padding: "8px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit",
                }}>
                  <Download size={13} /> Export not found ({counts.not_found})
                </button>
              )}
              <button className="btn-anim" onClick={exportWorkbook} style={{
                marginLeft: counts.not_found > 0 ? 0 : "auto", display: "flex", alignItems: "center", gap: 6,
                background: GRAD_ACCENT, color: WHITE, border: "none", borderRadius: 20,
                padding: "8px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit",
                boxShadow: "0 4px 12px rgba(227,28,121,0.28)",
              }}>
                <Download size={13} /> Export (All + Changes tabs)
              </button>
            </div>

            <div style={{ border: `1px solid ${LINE}`, borderRadius: 12, overflow: "hidden", background: WHITE, boxShadow: "0 4px 16px rgba(43,11,94,0.06)" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ background: GRAD_BRAND, color: WHITE }}>
                    {["Name", "Old", "New", "Status", "Verified"].map((h) => (
                      <th key={h} style={{ textAlign: "left", padding: "12px 14px", fontFamily: MONO, fontSize: 11, letterSpacing: 0.6, textTransform: "uppercase", fontWeight: 700 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((r, i) => {
                    const meta = STATUS_META[r.status] || STATUS_META.error;
                    const Icon = meta.Icon;
                    return (
                      <tr key={i} className="row-anim row-hover" style={{
                        borderTop: `1px solid ${LINE}`, animationDelay: `${Math.min(i, 20) * 0.02}s`,
                        background: i % 2 ? "#FBFAFD" : WHITE, transition: "background 0.15s ease",
                      }}>
                        <td style={{ padding: "11px 14px", fontWeight: 700 }}>{r.name}</td>
                        <td style={{ padding: "11px 14px", color: SLATE }}>{r.old_title}{r.old_company ? `, ${r.old_company}` : ""}</td>
                        <td style={{ padding: "11px 14px" }}>
                          {r.status === "changed" ? `${r.resolved_title}, ${r.resolved_company}` : r.status === "no_change" ? "—" : ""}
                        </td>
                        <td style={{ padding: "11px 14px" }}>
                          <span style={{
                            display: "inline-flex", alignItems: "center", gap: 5, color: meta.color, fontWeight: 700, fontSize: 11.5,
                            background: `${meta.color}14`, padding: "4px 9px", borderRadius: 20,
                          }}>
                            <Icon size={12} /> {meta.label}
                          </span>
                        </td>
                        <td style={{ padding: "11px 14px", color: SLATE, fontFamily: MONO, fontSize: 12 }}>{r.last_verified}</td>
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
    </div>
  );
}
