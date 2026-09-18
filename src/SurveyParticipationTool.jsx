import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { QualityIcon, QualityPageHeader } from "./UiElements.jsx";

// ── Survey participation report analyzer ────────────────────────────────────
// Reads the system's survey-performance export (COURSE_CODE, COURSE_DESCR_EN,
// NoOfVotes) and flags courses that need manual attention before running the
// full evaluation: courses nobody voted on, courses with no students enrolled
// at all, and courses with a suspiciously low vote count (< threshold) that
// need a manual check for whether they're a "regular" course.

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

// NoOfVotes comes as a "voted/total" fraction string (e.g. "0/25"). Falls
// back to treating it as a plain number (voted only, total unknown) if no
// "/" is present, so odd exports don't just crash the parser.
function parseVotes(v) {
  const s = (v ?? "").toString().trim();
  if (s.includes("/")) {
    const [a, b] = s.split("/");
    const voted = parseInt(a, 10);
    const total = parseInt(b, 10);
    return { voted: Number.isFinite(voted) ? voted : 0, total: Number.isFinite(total) ? total : null };
  }
  const n = parseInt(s, 10);
  return { voted: Number.isFinite(n) ? n : 0, total: null };
}

function classify(voted, total, threshold) {
  if (total === 0) return "none-enrolled";
  if (voted === 0) return "not-voted";
  if (voted < threshold) return "suspicious";
  return "regular";
}

const STATUS = {
  "none-enrolled": { label: "لا يوجد طلاب مسجلين", color: "#94a3b8" },
  "not-voted": { label: "لم يتم التقييم", color: "#fb7185" },
  "suspicious": { label: "مشكوك في انتظامها", color: "#fcd34d" },
  "regular": { label: "طبيعية", color: "#34d399" },
};

function courseRowsToWorkbook(rows, sheetName) {
  const aoa = [["كود المقرر", "اسم المقرر", "عدد المصوتين", "الإجمالي", "الحالة"],
    ...rows.map(r => [r.code, r.name, r.voted, r.total ?? "", STATUS[r.status].label])];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 16 }, { wch: 45 }, { wch: 12 }, { wch: 10 }, { wch: 20 }];
  ws["!autofilter"] = { ref: ws["!ref"] };
  ws["!sheetViews"] = [{ rightToLeft: true }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return wb;
}

function exportRows(rows, sheetName, filename) {
  if (!rows.length) return false;
  const wb = courseRowsToWorkbook(rows, sheetName);
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array", compression: true });
  downloadBlob(new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename);
  return true;
}

export default function SurveyParticipationTool() {
  const [rows, setRows] = useState([]);
  const [threshold, setThreshold] = useState(10);
  const [filter, setFilter] = useState("all");
  const [status, setStatus] = useState({ msg: "", kind: "" });
  const inputRef = useRef(null);

  function handleUpload(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    file.arrayBuffer().then(buf => {
      try {
        const wb = XLSX.read(buf, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const parsed = data.map(r => {
          const code = r.COURSE_CODE ?? r.course_code ?? r.CourseCode ?? "";
          const name = r.COURSE_DESCR_EN ?? r.course_descr_en ?? r.CourseDescrEn ?? "";
          const { voted, total } = parseVotes(r.NoOfVotes ?? r.noOfVotes ?? r.NOOFVOTES ?? "");
          return { code, name, voted, total };
        }).filter(r => r.code || r.name);
        setRows(parsed);
        setStatus({ msg: `تم تحليل ${parsed.length} مقرر.`, kind: "ok" });
      } catch (err) {
        setStatus({ msg: `تعذّرت قراءة الملف: ${err.message}`, kind: "err" });
      }
    });
  }

  const classified = useMemo(
    () => rows.map(r => ({ ...r, status: classify(r.voted, r.total, threshold) })),
    [rows, threshold]
  );

  const buckets = useMemo(() => ({
    "none-enrolled": classified.filter(r => r.status === "none-enrolled"),
    "not-voted": classified.filter(r => r.status === "not-voted"),
    "suspicious": classified.filter(r => r.status === "suspicious"),
    "regular": classified.filter(r => r.status === "regular"),
  }), [classified]);

  const filled = classified.filter(r => r.voted > 0);
  const notFilled = classified.filter(r => r.voted === 0);

  const visible = filter === "all" ? classified : classified.filter(r => r.status === filter);
  const handleExport = (exportRowsData, sheetName, filename) => {
    if (exportRows(exportRowsData, sheetName, filename)) {
      setStatus({ msg: `✓ تم تنزيل ${filename}`, kind: "ok" });
    }
  };

  return (
    <div>
      <QualityPageHeader
        icon="chart"
        eyebrow="متابعة المشاركة"
        title="أداء الاستبيانات"
        description="راجع نسب مشاركة الطلاب وحدد المقررات التي تحتاج متابعة قبل إصدار التقارير."
      />

      <div className="card" style={{ padding: 22, marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <button className="btn btn-primary btn-sm qa-icon-button" onClick={() => inputRef.current?.click()}><QualityIcon name="upload" size={15} /> رفع ملف تقرير الأداء</button>
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={handleUpload} />
          <label style={{ fontSize: 12, color: "#8ea3b4", display: "flex", alignItems: "center", gap: 6 }}>
            حد "الشك" (أقل من)
            <input className="input" type="number" min={1} value={threshold}
              onChange={e => setThreshold(Math.max(1, parseInt(e.target.value, 10) || 1))}
              style={{ width: 70, padding: "4px 8px" }} />
          </label>
        </div>
        {status.msg && (
          <div style={{ marginTop: 12, fontSize: 13, color: status.kind === "err" ? "#fb7185" : "#34d399" }}>{status.msg}</div>
        )}
      </div>

      {rows.length > 0 && (
        <div className="card" style={{ padding: 22 }}>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 18 }}>
            <Stat n={classified.length} l="إجمالي المواد" />
            <Stat n={buckets["regular"].length} l="طبيعية" cls="good" />
            <Stat n={buckets["suspicious"].length} l={`مشكوك فيها (أقل من ${threshold})`} cls="warn" />
            <Stat n={buckets["not-voted"].length} l="لم يتم التقييم" cls="bad" />
            <Stat n={buckets["none-enrolled"].length} l="لا يوجد طلاب" />
          </div>

          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 18 }}>
            <button className="btn btn-ghost btn-sm qa-icon-button" disabled={!filled.length}
              onClick={() => handleExport(filled, "تم التقييم", "مواد_تم_تقييمها.xlsx")}>
              <QualityIcon name="download" size={14} /> تنزيل المواد اللي فيها طلبة ملوها ({filled.length})
            </button>
            <button className="btn btn-ghost btn-sm qa-icon-button" disabled={!notFilled.length}
              onClick={() => handleExport(notFilled, "لم يتم التقييم", "مواد_لم_تُقيّم.xlsx")}>
              <QualityIcon name="download" size={14} /> تنزيل المواد اللي مفيش طلبة ملوها ({notFilled.length})
            </button>
            <button className="btn btn-ghost btn-sm qa-icon-button" disabled={!buckets["suspicious"].length}
              onClick={() => handleExport(buckets["suspicious"], "مشكوك فيها", "مواد_مشكوك_في_انتظامها.xlsx")}>
              <QualityIcon name="download" size={14} /> تنزيل المواد المشكوك فيها ({buckets["suspicious"].length})
            </button>
            <button className="btn btn-primary btn-sm qa-icon-button" disabled={!classified.length}
              onClick={() => handleExport(classified, "تقرير كامل", "تقرير_أداء_الاستبيانات_كامل.xlsx")}>
              <QualityIcon name="download" size={14} /> تنزيل التقرير الكامل
            </button>
          </div>

          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
            {[["all", "الكل"], ["regular", "طبيعية"], ["suspicious", "مشكوك فيها"], ["not-voted", "لم يتم التقييم"], ["none-enrolled", "لا يوجد طلاب"]].map(([key, label]) => (
              <button key={key} onClick={() => setFilter(key)} className={`btn btn-sm ${filter === key ? "btn-primary" : "btn-ghost"}`}>
                {label}
              </button>
            ))}
          </div>

          <div style={{ overflowX: "auto" }}>
            <table className="mini-table">
              <thead>
                <tr>
                  <th>كود المقرر</th>
                  <th>اسم المقرر</th>
                  <th>المصوتين</th>
                  <th>الإجمالي</th>
                  <th>الحالة</th>
                </tr>
              </thead>
              <tbody>
                {visible.length === 0 && (
                  <tr><td colSpan={5} style={{ color: "rgba(255,255,255,.4)", padding: 20 }}>لا توجد نتائج مطابقة لهذا الفلتر</td></tr>
                )}
                {visible.map((r, i) => (
                  <tr key={i}>
                    <td>{r.code || "—"}</td>
                    <td className="rtl-td">{r.name || "—"}</td>
                    <td>{r.voted}</td>
                    <td>{r.total ?? "—"}</td>
                    <td>
                      <span style={{ display: "inline-block", padding: "2px 10px", borderRadius: 999, fontSize: 11, fontWeight: 600, color: STATUS[r.status].color, background: `${STATUS[r.status].color}22`, border: `1px solid ${STATUS[r.status].color}55` }}>
                        {STATUS[r.status].label}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function Stat({ n, l, cls }) {
  const color = cls === "good" ? "#34d399" : cls === "bad" ? "#fb7185" : cls === "warn" ? "#fcd34d" : "#fff";
  return (
    <div style={{ background: "rgba(255,255,255,.05)", border: "1px solid rgba(255,255,255,.1)", borderRadius: 8, padding: "10px 16px", minWidth: 110 }}>
      <div style={{ fontSize: 20, fontWeight: 700, color }}>{n}</div>
      <div style={{ color: "rgba(255,255,255,.5)", fontSize: 11, marginTop: 2 }}>{l}</div>
    </div>
  );
}
