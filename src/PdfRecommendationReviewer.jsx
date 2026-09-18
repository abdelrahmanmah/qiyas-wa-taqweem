import { useRef, useState } from "react";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { QualityIcon, QualityPageHeader } from "./UiElements.jsx";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

// ── PDF recommendation reviewer ─────────────────────────────────────────────
// Ports the standalone "pdf_rec_reviewer" tool: open a folder of course-review
// PDFs, render the last page of each as a thumbnail (that's where the
// reviewer's recommendation section lives), let the user quickly mark which
// ones actually contain a recommendation, and export the marked list to Excel.

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

const CONCURRENCY = 4;

export default function PdfRecommendationReviewer() {
  const [items, setItems] = useState([]);
  const [renderState, setRenderState] = useState({}); // idx -> "loading" | "done" | "error"
  const [thumbs, setThumbs] = useState({}); // idx -> dataURL
  const [checked, setChecked] = useState({}); // idx -> bool
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const inputRef = useRef(null);

  const total = items.length;
  const checkedCount = Object.values(checked).filter(Boolean).length;

  async function handleFolder(e) {
    const files = [...e.target.files].filter(f => f.name.toLowerCase().endsWith(".pdf"))
      .sort((a, b) => a.name.localeCompare(b.name, "ar"));
    e.target.value = "";
    if (!files.length) return;

    setItems(files);
    setThumbs({}); setChecked({}); setRenderState({});
    setProgress({ done: 0, total: files.length });
    renderAll(files);
  }

  async function renderAll(files) {
    const queue = files.map((_, i) => i);
    let done = 0;
    const worker = async () => {
      while (queue.length) {
        const i = queue.shift();
        setRenderState(s => ({ ...s, [i]: "loading" }));
        try {
          const buf = await files[i].arrayBuffer();
          const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
          const page = await pdf.getPage(pdf.numPages);
          const base = page.getViewport({ scale: 1 });
          const scale = Math.min(2.0, 300 / base.width);
          const vp = page.getViewport({ scale });
          const canvas = document.createElement("canvas");
          canvas.width = vp.width;
          canvas.height = vp.height;
          await page.render({ canvasContext: canvas.getContext("2d"), viewport: vp }).promise;
          const dataUrl = canvas.toDataURL("image/png");
          setThumbs(t => ({ ...t, [i]: dataUrl }));
          setRenderState(s => ({ ...s, [i]: "done" }));
        } catch {
          setRenderState(s => ({ ...s, [i]: "error" }));
        }
        done++;
        setProgress({ done, total: files.length });
      }
    };
    await Promise.all(Array.from({ length: CONCURRENCY }, worker));
  }

  function toggle(i) {
    setChecked(c => ({ ...c, [i]: !c[i] }));
  }
  function selectAll() {
    setChecked(Object.fromEntries(items.map((_, i) => [i, true])));
  }
  function clearAll() {
    setChecked({});
  }

  function doExport() {
    const selected = items.map((f, i) => ({ f, i })).filter(({ i }) => checked[i]);
    if (!selected.length) return;
    const data = [["#", "اسم المقرر", "اسم الملف"]];
    selected.forEach(({ f }, i) => data.push([i + 1, f.name.replace(/\.pdf$/i, ""), f.name]));
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [{ wch: 5 }, { wch: 50 }, { wch: 55 }];
    ws["!sheetViews"] = [{ rightToLeft: true }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "توصيات المقررات");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBlob(new Blob([wbout], { type: "application/octet-stream" }), `توصيات_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  return (
    <div>
      <QualityPageHeader
        icon="report"
        eyebrow="مراجعة تقارير المقررات"
        title="مراجعة التوصيات"
        description="اعرض الصفحة الأخيرة من ملفات PDF وحدد المقررات التي تحتوي على توصيات ثم صدّر القائمة إلى Excel."
      />

      <div className="card" style={{ padding: 22, marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <button className="btn btn-primary btn-sm qa-icon-button" onClick={() => inputRef.current?.click()}><QualityIcon name="folder" size={15} /> فتح مجلد</button>
          <input
            ref={inputRef} type="file" webkitdirectory="" multiple style={{ display: "none" }}
            onChange={handleFolder}
          />
          {total > 0 && (
            <>
              <button className="btn btn-ghost btn-sm" onClick={selectAll}>تحديد الكل</button>
              <button className="btn btn-ghost btn-sm" onClick={clearAll}>مسح الكل</button>
              <button className="btn btn-primary btn-sm qa-icon-button" disabled={!checkedCount} onClick={doExport}><QualityIcon name="download" size={15} /> تصدير Excel ({checkedCount})</button>
            </>
          )}
        </div>
        {total === 0 && (
          <div style={{ marginTop: 14, fontSize: 12, color: "rgba(255,255,255,.4)" }}>
            <span className="qa-icon-label"><QualityIcon name="info" size={14} /> يعمل اختيار المجلد في Chrome وEdge فقط.</span>
          </div>
        )}
        {total > 0 && (
          <div style={{ marginTop: 14 }}>
            <div style={{ fontSize: 12, color: "rgba(255,255,255,.5)", marginBottom: 6 }}>
              إجمالي: {total} · ✅ محدد: {checkedCount} · تم عرض {progress.done}/{progress.total}
            </div>
            <div style={{ height: 3, background: "rgba(255,255,255,.1)", borderRadius: 3, overflow: "hidden" }}>
              <div style={{ width: `${(progress.done / Math.max(1, progress.total)) * 100}%`, height: "100%", background: "linear-gradient(90deg,#34d399,#60a5fa)", transition: "width .3s" }} />
            </div>
          </div>
        )}
      </div>

      {total > 0 && (
        <div style={{
          display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(190px, 1fr))",
          gap: 12,
        }}>
          {items.map((f, i) => (
            <PdfCard
              key={i}
              name={f.name}
              state={renderState[i]}
              thumb={thumbs[i]}
              on={!!checked[i]}
              onToggle={() => toggle(i)}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function PdfCard({ name, state, thumb, on, onToggle }) {
  const nameNoExt = name.replace(/\.pdf$/i, "");
  return (
    <div
      className="card"
      onClick={onToggle}
      style={{
        padding: 0, overflow: "hidden", cursor: "pointer",
        borderColor: on ? "#34d399" : undefined,
        boxShadow: on ? "0 0 0 1px #34d399, 0 4px 20px rgba(52,211,153,.18)" : undefined,
        background: on ? "rgba(52,211,153,.08)" : undefined,
        position: "relative",
      }}
    >
      {on && (
        <div style={{
          position: "absolute", top: 8, insetInlineEnd: 8, background: "#34d399", color: "#031A0C",
          fontSize: 11, fontWeight: 800, padding: "3px 9px", borderRadius: 20, zIndex: 2,
        }}>✓ توصيات</div>
      )}
      <div style={{ width: "100%", background: "#08080D", minHeight: 180, display: "flex", alignItems: "center", justifyContent: "center" }}>
        {state === "done" && thumb ? (
          <img src={thumb} alt={nameNoExt} style={{ width: "100%", display: "block" }} />
        ) : state === "error" ? (
          <div style={{ padding: 20, color: "rgba(255,255,255,.4)", fontSize: 12, textAlign: "center" }}>⚠️<br />تعذّر تحميل الملف</div>
        ) : (
          <div style={{ width: "100%", height: 220, background: "linear-gradient(110deg, rgba(255,255,255,.05) 30%, rgba(255,255,255,.1) 50%, rgba(255,255,255,.05) 70%)" }} />
        )}
      </div>
      <div style={{ padding: "8px 10px", display: "flex", alignItems: "center", gap: 8, borderTop: "1px solid rgba(255,255,255,.08)" }}>
        <input type="checkbox" checked={on} onChange={onToggle} onClick={e => e.stopPropagation()} style={{ accentColor: "#34d399", cursor: "pointer" }} />
        <span title={name} style={{
          fontSize: 11.5, color: on ? "#34d399" : "rgba(255,255,255,.6)", fontWeight: on ? 600 : 400,
          flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", direction: "ltr", textAlign: "left",
        }}>{nameNoExt}</span>
      </div>
    </div>
  );
}
