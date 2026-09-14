import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

// ── Course assignment template: download blank / upload / edit / insights ──
// Columns match the paper form used to hand out course evaluations:
// اسم المقرر - كود المقرر - عضو هيئة التدريس - عضو الهيئة المعاونة -
// القسم العلمي - هل يوجد لاب (نعم-لا) - القائم بالمراجعة

const COLUMNS = [
  { key: "courseName", label: "اسم المقرر" },
  { key: "courseCode", label: "كود المقرر" },
  { key: "instructor", label: "عضو هيئة التدريس" },
  { key: "assistant", label: "عضو الهيئة المعاونة" },
  { key: "department", label: "القسم العلمي" },
  { key: "hasLab", label: "هل يوجد لاب (نعم-لا)" },
  { key: "reviewer", label: "القائم بالمراجعة" },
];
const HEADER_ROW = COLUMNS.map(c => c.label);

function emptyRow() {
  return { courseName: "", courseCode: "", instructor: "", assistant: "", department: "", hasLab: "", reviewer: "" };
}

function normHeader(s) {
  return (s || "").toString().replace(/\(.*?\)/g, "").replace(/\s+/g, "").trim();
}

// Matches an uploaded header row to our column keys — by normalized text
// first, falling back to plain column position if a header isn't recognized
// (so a file with slightly different header wording still imports cleanly).
function mapHeaders(headerRow) {
  const normTargets = COLUMNS.map(c => normHeader(c.label));
  return (headerRow || []).map((h, i) => {
    const nh = normHeader(h);
    const idx = normTargets.findIndex(t => t === nh || t.includes(nh) || nh.includes(t));
    return idx >= 0 ? COLUMNS[idx].key : (COLUMNS[i] ? COLUMNS[i].key : null);
  });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function rowsToWorkbook(rows) {
  const aoa = [HEADER_ROW, ...rows.map(r => COLUMNS.map(c => r[c.key] || ""))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = COLUMNS.map(() => ({ wch: 22 }));
  ws["!sheetViews"] = [{ rightToLeft: true }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "قالب المقررات");
  return wb;
}

function normLab(v) {
  const t = (v || "").toString().trim();
  if (/^(نعم|yes|y|true|1)$/i.test(t)) return "نعم";
  if (/^(لا|no|n|false|0)$/i.test(t)) return "لا";
  return t;
}

export default function CourseTemplateTool() {
  const [rows, setRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const inputRef = useRef(null);

  function downloadEmptyTemplate() {
    const wb = rowsToWorkbook([]);
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBlob(new Blob([wbout], { type: "application/octet-stream" }), "قالب_بيانات_المقررات.xlsx");
  }

  function downloadEditedFile() {
    const wb = rowsToWorkbook(rows);
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBlob(new Blob([wbout], { type: "application/octet-stream" }), `بيانات_المقررات_${fileName || "معدّلة"}.xlsx`.replace(/\.xlsx\.xlsx$/, ".xlsx"));
  }

  function handleUpload(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setFileName(file.name.replace(/\.(xlsx|xls|csv)$/i, ""));
    file.arrayBuffer().then(buf => {
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (!aoa.length) return;
      const keys = mapHeaders(aoa[0]);
      const newRows = aoa.slice(1)
        .filter(r => r.some(cell => (cell ?? "").toString().trim() !== ""))
        .map(r => {
          const row = emptyRow();
          keys.forEach((key, i) => {
            if (key) row[key] = key === "hasLab" ? normLab(r[i]) : (r[i] ?? "").toString().trim();
          });
          return row;
        });
      setRows(newRows);
    });
  }

  function updateCell(idx, key, value) {
    setRows(prev => prev.map((r, i) => i === idx ? { ...r, [key]: value } : r));
  }
  function addRow() {
    setRows(prev => [...prev, emptyRow()]);
  }
  function removeRow(idx) {
    setRows(prev => prev.filter((_, i) => i !== idx));
  }

  // ── insights ──
  const insights = useMemo(() => {
    const byDept = new Map();
    const byInstructor = new Map();
    const byAssistant = new Map();
    const byReviewer = new Map();
    const missing = [];
    const codeCounts = new Map();

    rows.forEach((r, i) => {
      if (r.department) {
        const d = byDept.get(r.department) || { total: 0, lab: 0 };
        d.total++;
        if (r.hasLab === "نعم") d.lab++;
        byDept.set(r.department, d);
      }
      if (r.instructor) byInstructor.set(r.instructor, (byInstructor.get(r.instructor) || 0) + 1);
      if (r.assistant) byAssistant.set(r.assistant, (byAssistant.get(r.assistant) || 0) + 1);
      if (r.reviewer) byReviewer.set(r.reviewer, (byReviewer.get(r.reviewer) || 0) + 1);

      const missingFields = [];
      if (!r.courseName) missingFields.push("اسم المقرر");
      if (!r.courseCode) missingFields.push("كود المقرر");
      if (!r.department) missingFields.push("القسم العلمي");
      if (missingFields.length) missing.push({ idx: i, row: r, missingFields });

      if (r.courseCode) {
        const key = r.courseCode.trim().toUpperCase();
        codeCounts.set(key, (codeCounts.get(key) || 0) + 1);
      }
    });

    const duplicates = rows.filter(r => r.courseCode && codeCounts.get(r.courseCode.trim().toUpperCase()) > 1);

    const toSortedList = m => [...m.entries()].sort((a, b) => b[1] - a[1]);

    return {
      byDept: [...byDept.entries()].sort((a, b) => b[1].total - a[1].total),
      deptDetail: byDept,
      byInstructor: toSortedList(byInstructor),
      byAssistant: toSortedList(byAssistant),
      byReviewer: toSortedList(byReviewer),
      missing,
      duplicates,
    };
  }, [rows]);

  const maxDeptTotal = Math.max(1, ...[...insights.deptDetail.values()].map(d => d.total));
  const maxWorkload = Math.max(1, ...insights.byInstructor.map(([, n]) => n), ...insights.byAssistant.map(([, n]) => n), ...insights.byReviewer.map(([, n]) => n));

  return (
    <div>
      <div style={{ marginBottom: 22 }}>
        <div style={{ color: "#fff", fontSize: 20, fontWeight: 900 }}>📋 قالب بيانات المقررات</div>
        <div style={{ color: "rgba(255,255,255,.55)", fontSize: 13, marginTop: 4 }}>
          نموذج موحّد لبيانات كل مقرر (المحاضر، المعاون، القسم، اللاب، القائم بالمراجعة) — حمّل القالب فارغًا أو ارفع نسخة موجودة وعدّل فيها مباشرة.
        </div>
      </div>

      <div className="card" style={{ padding: 22, marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button className="btn btn-primary btn-sm" onClick={downloadEmptyTemplate}>⬇ تحميل قالب فارغ</button>
          <button className="btn btn-ghost btn-sm" onClick={() => inputRef.current?.click()}>📤 رفع ملف موجود</button>
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={handleUpload} />
          {rows.length > 0 && (
            <>
              <button className="btn btn-ghost btn-sm" onClick={addRow}>+ إضافة صف</button>
              <button className="btn btn-primary btn-sm" onClick={downloadEditedFile}>⬇ تنزيل نسخة معدّلة</button>
            </>
          )}
        </div>
      </div>

      {rows.length > 0 && (
        <>
          <div className="card" style={{ padding: 22, marginBottom: 18 }}>
            <div style={{ color: "#fff", fontWeight: 700, fontSize: 15, marginBottom: 12 }}>البيانات ({rows.length} مقرر)</div>
            <div style={{ overflowX: "auto" }}>
              <table className="mini-table">
                <thead>
                  <tr>
                    {COLUMNS.map(c => <th key={c.key}>{c.label}</th>)}
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, idx) => (
                    <tr key={idx}>
                      {COLUMNS.map(c => (
                        <td key={c.key}>
                          {c.key === "hasLab" ? (
                            <button
                              className="btn btn-ghost btn-sm"
                              style={{ padding: "4px 12px", color: r.hasLab === "نعم" ? "#3fbf85" : r.hasLab === "لا" ? "#e0555a" : "rgba(255,255,255,.5)" }}
                              onClick={() => updateCell(idx, "hasLab", r.hasLab === "نعم" ? "لا" : r.hasLab === "لا" ? "" : "نعم")}
                            >
                              {r.hasLab || "—"}
                            </button>
                          ) : (
                            <input
                              className="input" style={{ padding: "6px 8px", fontSize: 12, minWidth: 100 }}
                              value={r[c.key]}
                              onChange={e => updateCell(idx, c.key, e.target.value)}
                            />
                          )}
                        </td>
                      ))}
                      <td><button className="btn btn-ghost btn-sm" onClick={() => removeRow(idx)} aria-label={`حذف صف المقرر ${r.courseName || idx + 1}`} title="حذف الصف" style={{ color: "#e0555a" }}>✕</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* insights */}
          <div className="card" style={{ padding: 22, marginBottom: 18 }}>
            <div style={{ color: "#fff", fontWeight: 700, fontSize: 15, marginBottom: 14 }}>📊 توزيع حسب القسم</div>
            {insights.byDept.length === 0 ? (
              <div style={{ color: "rgba(255,255,255,.4)", fontSize: 13 }}>لا توجد بيانات قسم بعد.</div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {insights.byDept.map(([dept]) => {
                  const d = insights.deptDetail.get(dept);
                  const pct = (d.total / maxDeptTotal) * 100;
                  return (
                    <div key={dept}>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#e8f0fe", marginBottom: 4 }}>
                        <span>{dept}</span>
                        <span style={{ color: "rgba(255,255,255,.55)" }}>{d.total} مقرر · {d.lab} فيه لاب</span>
                      </div>
                      <div style={{ background: "rgba(255,255,255,.08)", borderRadius: 6, height: 8, overflow: "hidden" }}>
                        <div style={{ width: `${pct}%`, height: "100%", background: "linear-gradient(90deg,#1abc9c,#16a085)" }} />
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          <div className="card" style={{ padding: 22, marginBottom: 18 }}>
            <div style={{ color: "#fff", fontWeight: 700, fontSize: 15, marginBottom: 14 }}>👤 عبء العمل لكل عضو</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 20 }}>
              <WorkloadList title="عضو هيئة التدريس" list={insights.byInstructor} max={maxWorkload} />
              <WorkloadList title="عضو الهيئة المعاونة" list={insights.byAssistant} max={maxWorkload} />
              <WorkloadList title="القائم بالمراجعة" list={insights.byReviewer} max={maxWorkload} />
            </div>
          </div>

          <div className="card" style={{ padding: 22 }}>
            <div style={{ color: "#fff", fontWeight: 700, fontSize: 15, marginBottom: 14 }}>⚠ بيانات ناقصة / تكرار</div>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 14 }}>
              <Stat n={insights.missing.length} l="صفوف فيها بيانات ناقصة" cls={insights.missing.length ? "warn" : "good"} />
              <Stat n={insights.duplicates.length} l="صفوف كود مقرر مكرر" cls={insights.duplicates.length ? "bad" : "good"} />
            </div>
            {insights.missing.length > 0 && (
              <div style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 12, color: "rgba(255,255,255,.5)", marginBottom: 6 }}>بيانات ناقصة:</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  {insights.missing.slice(0, 30).map((m, i) => (
                    <div key={i} style={{ fontSize: 12, color: "#e0a336" }}>
                      صف {m.idx + 1}: {m.row.courseName || m.row.courseCode || "—"} — ناقص: {m.missingFields.join("، ")}
                    </div>
                  ))}
                </div>
              </div>
            )}
            {insights.duplicates.length > 0 && (
              <div>
                <div style={{ fontSize: 12, color: "rgba(255,255,255,.5)", marginBottom: 6 }}>أكواد مكررة:</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  {insights.duplicates.map((r, i) => (
                    <div key={i} style={{ fontSize: 12, color: "#e0555a" }}>{r.courseCode} — {r.courseName || "—"}</div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

function WorkloadList({ title, list, max }) {
  return (
    <div>
      <div style={{ fontSize: 13, color: "#e8f0fe", fontWeight: 700, marginBottom: 8 }}>{title}</div>
      {list.length === 0 ? (
        <div style={{ color: "rgba(255,255,255,.35)", fontSize: 12 }}>لا يوجد</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {list.slice(0, 10).map(([name, n]) => (
            <div key={name}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5, color: "rgba(255,255,255,.7)", marginBottom: 2 }}>
                <span>{name}</span><span>{n}</span>
              </div>
              <div style={{ background: "rgba(255,255,255,.08)", borderRadius: 5, height: 6, overflow: "hidden" }}>
                <div style={{ width: `${(n / max) * 100}%`, height: "100%", background: "linear-gradient(90deg,#2874a6,#1abc9c)" }} />
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function Stat({ n, l, cls }) {
  const color = cls === "good" ? "#3fbf85" : cls === "bad" ? "#e0555a" : cls === "warn" ? "#e0a336" : "#fff";
  return (
    <div style={{ background: "rgba(255,255,255,.05)", border: "1px solid rgba(255,255,255,.1)", borderRadius: 8, padding: "10px 16px", minWidth: 110 }}>
      <div style={{ fontSize: 20, fontWeight: 700, color }}>{n}</div>
      <div style={{ color: "rgba(255,255,255,.5)", fontSize: 11, marginTop: 2 }}>{l}</div>
    </div>
  );
}
