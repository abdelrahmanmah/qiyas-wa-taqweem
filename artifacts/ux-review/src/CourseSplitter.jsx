import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";

// ── UMIS merged-report course splitter ──────────────────────────────────────
// Ports the standalone "course_eval_splitte" tool into the app: takes a
// merged UMIS course-evaluation Excel export (one file, many courses stacked
// with a repeating block-header row), splits it into one workbook per course,
// optionally matches each against a reference course list and a
// department-distribution list, and offers per-course or zipped-by-department
// downloads.

// The exact Arabic marker text that repeats once per course block, matching
// the verified-working standalone tool (course_eval_splitte V3r.html). Still
// user-editable in the UI (see `markerText` state below) in case a different
// UMIS report template/version uses different text — `findMarkerCandidates()`
// scans the sheet for other repeating strings so the user can pick the right
// one without needing to open the file in Excel.
const DEFAULT_MARKER = "بنود الاستبيان";
const BLOCK_OFFSET = 10; // rows between a course's block start and its marker row — also template-dependent

// ── text/matching helpers (ported 1:1 from the standalone tool) ────────────
function normCode(s) {
  return (s || "").toString().toUpperCase().replace(/[^A-Z0-9]/g, "");
}
function normName(s) {
  return (s || "").toString().toLowerCase().replace(/[^a-z0-9؀-ۿ]+/g, " ").replace(/\s+/g, " ").trim();
}
function cleanTitle(s) {
  return (s || "").toString().replace(/\n/g, " ").replace(/\s+/g, " ").replace(/^[.\s]+/, "").trim();
}
// Levenshtein-based similarity ratio, 0..1 (1 = identical)
function similarity(a, b) {
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  const m = a.length, n = b.length;
  const dp = new Array(n + 1);
  for (let j = 0; j <= n; j++) dp[j] = j;
  for (let i = 1; i <= m; i++) {
    let prev = dp[0]; dp[0] = i;
    for (let j = 1; j <= n; j++) {
      const tmp = dp[j];
      dp[j] = a[i - 1] === b[j - 1] ? prev : 1 + Math.min(prev, dp[j], dp[j - 1]);
      prev = tmp;
    }
  }
  return 1 - dp[n] / Math.max(m, n);
}
function findDepartment(code, name, deptByName, deptByCode, deptNameKeys) {
  const ck = normCode(code), nk = normName(name);
  // 1. exact name match (most reliable — codes vary between systems, names don't)
  if (nk && deptByName.has(nk)) return { dept: deptByName.get(nk).dept, method: "name", score: 0.95 };
  // 2. exact code match
  if (ck && deptByCode.has(ck)) return { dept: deptByCode.get(ck).dept, method: "code", score: 1 };
  // 3. fuzzy name match
  if (nk) {
    let best = null, bestScore = 0;
    for (const key of deptNameKeys) {
      const s = similarity(nk, key);
      if (s > bestScore) { bestScore = s; best = key; }
    }
    if (best && bestScore >= 0.8) return { dept: deptByName.get(best).dept, method: "fuzzy-name", score: bestScore };
  }
  return null;
}
function safeFileName(s) {
  return (s || "").toString().replace(/[\\/:*?"<>|]/g, "_").trim();
}
function bytesToSize(b) {
  if (b < 1024) return b + " B";
  if (b < 1024 * 1024) return (b / 1024).toFixed(0) + " KB";
  return (b / 1024 / 1024).toFixed(1) + " MB";
}
function courseFileName(c) {
  const t = safeFileName(c.cleanTitle || "مادة");
  return `${safeFileName(c.code)} - ${t}.xlsx`.slice(0, 150);
}
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}
function downloadCourse(c) {
  const wbout = XLSX.write(c.wb, { bookType: "xlsx", type: "array" });
  downloadBlob(new Blob([wbout], { type: "application/octet-stream" }), courseFileName(c));
}

// ── parse reference course list ─────────────────────────────────────────────
async function parseRefFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  const idx = new Map();
  rows.forEach(r => {
    const code = r.COURSE_CODE || r.course_code || r.CourseCode;
    if (!code) return;
    const key = normCode(code);
    if (!idx.has(key)) {
      idx.set(key, {
        code,
        descrEn: cleanTitle(r.COURSE_DESCR_EN || ""),
        descrAr: cleanTitle(r.COURSE_DESCR_AR || ""),
        faculty: r.FACULTY_AR || r.FACULTY_EN || "",
      });
    }
  });
  return idx;
}

// ── parse department-distribution list ──────────────────────────────────────
async function parseDeptFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  const byName = new Map();
  const byCode = new Map();
  rows.forEach(r => {
    const name = r["اسم المقرر"] || r["اسم المادة"] || "";
    const code = r["كود المقرر"] || r["كود المادة"] || "";
    const dept = r["القسم العلمي"] || r["القسم"] || "";
    if (!dept) return;
    const nk = normName(name);
    const ck = normCode(code);
    if (nk && !byName.has(nk)) byName.set(nk, { code, dept });
    if (ck && !byCode.has(ck)) byCode.set(ck, { name, dept });
  });
  return { byName, byCode };
}

// Scans a sheet for strings that repeat a plausible number of times (2..200)
// — candidates for the block marker when the configured one matches nothing.
// Returns the top few, most-frequent first.
function findMarkerCandidates(wb) {
  const shName = wb.SheetNames[0];
  const ws = wb.Sheets[shName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
  const counts = new Map();
  rows.forEach(row => {
    row.forEach(cell => {
      if (typeof cell !== "string") return;
      const t = cell.trim();
      if (t.length < 3 || t.length > 60) return;
      counts.set(t, (counts.get(t) || 0) + 1);
    });
  });
  return [...counts.entries()]
    .filter(([, n]) => n >= 2 && n <= 200)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8)
    .map(([text, count]) => ({ text, count }));
}

// ── find blocks & split a merged workbook into one workbook per course ─────
function extractCoursesFromWorkbook(wb, sourceName, markerText, blockOffset) {
  const shName = wb.SheetNames[0];
  const ws = wb.Sheets[shName];
  const ref = XLSX.utils.decode_range(ws["!ref"]);
  const totalRows = ref.e.r + 1;
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
  const merges = ws["!merges"] || [];
  const cols = ws["!cols"];

  const markerNorm = markerText.replace(/\s/g, "");
  const markerRows = [];
  if (markerNorm) {
    rows.forEach((row, r) => {
      for (const cell of row) {
        if (typeof cell === "string" && cell.replace(/\s/g, "").includes(markerNorm)) {
          markerRows.push(r);
          break;
        }
      }
    });
  }

  if (markerRows.length === 0) {
    const candidates = findMarkerCandidates(wb);
    return {
      ok: false,
      error: `لم يتم العثور على النص "${markerText}" في أي صف داخل هذا الملف.`,
      candidates,
      courses: [],
    };
  }

  const firstBlockStart = markerRows[0] - blockOffset; // "السنة الدراسية" row
  const headerEnd = Math.max(0, firstBlockStart - 1); // fixed top header rows 0..headerEnd
  const headerRowCount = headerEnd + 1;

  const results = [];

  markerRows.forEach((m, i) => {
    const blockStart = m - blockOffset;
    const blockEnd = (i < markerRows.length - 1) ? (markerRows[i + 1] - blockOffset - 1) : (totalRows - 1);
    if (blockStart < 0 || blockEnd < blockStart) return;

    // course title: search a few rows around (marker-8 .. marker-1) for "... (CODE)"
    let rawTitle = "", code = "", nameOnly = "";
    for (let r = Math.max(0, m - 8); r < m; r++) {
      const rowArr = rows[r] || [];
      for (const cell of rowArr) {
        if (typeof cell === "string") {
          const match = cell.match(/^([\s\S]*)\(([A-Za-z0-9.\-]{3,15})\)\s*$/);
          if (match) { rawTitle = cell; nameOnly = match[1]; code = match[2]; break; }
        }
      }
      if (code) break;
    }
    if (!code) code = "UNKNOWN_" + (i + 1);

    const newRows = [];
    for (let r = 0; r <= headerEnd; r++) newRows.push(rows[r] ? rows[r].slice() : []);
    for (let r = blockStart; r <= blockEnd; r++) newRows.push(rows[r] ? rows[r].slice() : []);

    const newWs = XLSX.utils.aoa_to_sheet(newRows);

    const offset = headerRowCount - blockStart;
    const newMerges = [];
    merges.forEach(mg => {
      if (mg.s.r <= headerEnd && mg.e.r <= headerEnd) {
        newMerges.push({ s: { r: mg.s.r, c: mg.s.c }, e: { r: mg.e.r, c: mg.e.c } });
      } else if (mg.s.r >= blockStart && mg.e.r <= blockEnd) {
        newMerges.push({ s: { r: mg.s.r + offset, c: mg.s.c }, e: { r: mg.e.r + offset, c: mg.e.c } });
      }
    });
    newWs["!merges"] = newMerges;
    if (cols) newWs["!cols"] = cols;

    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, newWs, "Sheet1");

    results.push({ code, rawTitle, cleanTitle: cleanTitle(nameOnly), wb: newWb, sourceFile: sourceName });
  });

  return { ok: true, courses: results };
}

// ── small file-list row (used for the 3 drop zones) ─────────────────────────
function FileItem({ name, size, onRemove }) {
  return (
    <div style={{
      display: "flex", alignItems: "center", justifyContent: "space-between",
      background: "rgba(255,255,255,.05)", border: "1px solid rgba(255,255,255,.1)",
      borderRadius: 8, padding: "8px 12px", fontSize: 13, marginTop: 6,
    }}>
      <span style={{ color: "#e8f0fe" }}>📄 {name}</span>
      <span style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <span style={{ color: "rgba(255,255,255,.45)", fontSize: 11 }}>{bytesToSize(size)}</span>
        {onRemove && <button onClick={onRemove} aria-label={`إزالة الملف ${name}`} title={`إزالة ${name}`} style={{ background: "none", border: "none", color: "#e0555a", cursor: "pointer", fontSize: 15, padding: "0 4px" }}>✕</button>}
      </span>
    </div>
  );
}

function DropZone({ icon, main, sub, onFiles, multiple, inputRef }) {
  const [drag, setDrag] = useState(false);
  return (
    <div
      className={`upload-zone ${drag ? "drag" : ""}`}
      style={{ padding: "26px 20px" }}
      onClick={() => inputRef.current?.click()}
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); onFiles(Array.from(e.dataTransfer.files)); }}
    >
      <input ref={inputRef} type="file" accept=".xlsx,.xls" multiple={multiple} style={{ display: "none" }}
        onChange={e => {
          const picked = Array.from(e.target.files);
          e.target.value = "";
          onFiles(picked);
        }} />
      <div style={{ fontSize: 26, marginBottom: 4, opacity: .85 }}>{icon}</div>
      <div style={{ fontSize: 14, color: "#e8f0fe" }}>{main}</div>
      <div style={{ fontSize: 12, color: "rgba(255,255,255,.45)", marginTop: 4 }}>{sub}</div>
    </div>
  );
}

// ── main component ───────────────────────────────────────────────────────────
export default function CourseSplitter() {
  const [refFile, setRefFile] = useState(null);
  const [deptFile, setDeptFile] = useState(null);
  const [mergedFiles, setMergedFiles] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [status, setStatus] = useState({ msg: "", kind: "" });
  const [extractedCourses, setExtractedCourses] = useState([]);
  const [refIndex, setRefIndex] = useState(new Map());
  const [deptInfo, setDeptInfo] = useState({ byName: new Map(), byCode: new Map() });
  const [showResults, setShowResults] = useState(false);
  const [sectionFilter, setSectionFilter] = useState("all");
  const [markerText, setMarkerText] = useState(DEFAULT_MARKER);
  const [blockOffset, setBlockOffset] = useState(BLOCK_OFFSET);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [markerCandidates, setMarkerCandidates] = useState([]);

  const refInputRef = useRef(null);
  const deptInputRef = useRef(null);
  const mergedInputRef = useRef(null);

  const hasDept = deptInfo.byName.size > 0 || deptInfo.byCode.size > 0;

  function resetAll() {
    setRefFile(null); setDeptFile(null); setMergedFiles([]);
    setExtractedCourses([]); setRefIndex(new Map());
    setDeptInfo({ byName: new Map(), byCode: new Map() });
    setShowResults(false); setSectionFilter("all");
    setStatus({ msg: "", kind: "" });
    setMarkerCandidates([]); setMarkerText(DEFAULT_MARKER); setBlockOffset(BLOCK_OFFSET);
  }

  async function handleProcess() {
    if (mergedFiles.length === 0) return;
    setProcessing(true);
    setMarkerCandidates([]);
    setStatus({ msg: refFile ? "جارٍ المعالجة..." : "⚠ لسه ما رفعتش \"قائمة المواد الأساسية\" — هيتم التقسيم بره، بس مش هيتعرف أي مادة ناقصة.", kind: refFile ? "" : "err" });

    try {
      const newRefIndex = refFile ? await parseRefFile(refFile) : new Map();
      const newDeptInfo = deptFile ? await parseDeptFile(deptFile) : { byName: new Map(), byCode: new Map() };
      const deptNameKeys = [...newDeptInfo.byName.keys()];

      let courses = [];
      let allCandidates = [];
      for (const f of mergedFiles) {
        const buf = await f.arrayBuffer();
        let wb;
        try {
          wb = XLSX.read(buf, { type: "array", cellText: false });
        } catch (e) {
          setStatus({ msg: `تعذّرت قراءة الملف: ${f.name} — ${e.message}`, kind: "err" });
          continue;
        }
        const out = extractCoursesFromWorkbook(wb, f.name, markerText, blockOffset);
        if (!out.ok) {
          setStatus({ msg: `⚠ ${f.name}: ${out.error}`, kind: "err" });
          if (out.candidates?.length) allCandidates = out.candidates;
          continue;
        }
        courses.push(...out.courses);
      }

      if (courses.length === 0 && allCandidates.length > 0) {
        setMarkerCandidates(allCandidates);
        setStatus({ msg: "لم يتم العثور على أي مادة بالنص الحالي — جرّب أحد الأنماط المتكررة المقترحة تحت (أو عدّل النص يدويًا في «إعدادات متقدمة»).", kind: "err" });
      }

      // mark duplicates (same code found more than once across files)
      const seen = new Map();
      courses.forEach(c => {
        const key = normCode(c.code);
        if (!seen.has(key)) seen.set(key, []);
        seen.get(key).push(c);
      });
      courses.forEach(c => {
        const key = normCode(c.code);
        c.dup = seen.get(key).length > 1;
        const refEntry = newRefIndex.get(key);
        c.matched = !!refEntry;
        if (refEntry && !c.cleanTitle) c.cleanTitle = refEntry.descrEn;

        const deptMatch = (newDeptInfo.byName.size || newDeptInfo.byCode.size)
          ? findDepartment(c.code, c.cleanTitle, newDeptInfo.byName, newDeptInfo.byCode, deptNameKeys)
          : null;
        c.department = deptMatch ? deptMatch.dept : null;
        c.deptMethod = deptMatch ? deptMatch.method : null;
        c.deptScore = deptMatch ? deptMatch.score : 0;
      });

      setRefIndex(newRefIndex);
      setDeptInfo(newDeptInfo);
      setExtractedCourses(courses);
      setShowResults(true);
      setStatus({ msg: `تم استخراج ${courses.length} مادة من ${mergedFiles.length} ملف.`, kind: "ok" });
    } catch (err) {
      setStatus({ msg: `حصل خطأ غير متوقع: ${err.message}`, kind: "err" });
    } finally {
      setProcessing(false);
    }
  }

  async function handleZipAll() {
    if (extractedCourses.length === 0) return;
    const zip = new JSZip();
    extractedCourses.forEach(c => {
      const wbout = XLSX.write(c.wb, { bookType: "xlsx", type: "array" });
      if (hasDept) {
        const folderName = safeFileName(c.department || "غير محدد - يحتاج مراجعة");
        zip.folder(folderName).file(courseFileName(c), wbout);
      } else {
        zip.file(courseFileName(c), wbout);
      }
    });
    const blob = await zip.generateAsync({ type: "blob" });
    downloadBlob(blob, "مواد_مقسّمة.zip");
  }

  function handleDownloadReport() {
    const dupCourses = extractedCourses.filter(c => c.dup);
    const matchedCodes = new Set(extractedCourses.filter(c => c.matched).map(c => normCode(c.code)));
    const missingList = [];
    if (refIndex.size > 0) refIndex.forEach((v, k) => { if (!matchedCodes.has(k)) missingList.push(v); });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(extractedCourses.map(c => ({
      "كود المادة": c.code,
      "اسم المادة": c.cleanTitle || "",
      "القسم": c.department || "",
      "الحالة": c.matched ? "موجودة ومطابقة" : refIndex.size > 0 ? "موجودة — غير مطابقة بالقائمة" : "موجودة",
      "مكررة؟": c.dup ? "نعم" : "لا",
      "الملف المصدر": c.sourceFile || "",
    }))), "موجودة");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(missingList.map(m => ({
      "كود المادة": m.code,
      "اسم المادة": m.descrEn || m.descrAr || "",
    }))), "ناقصة");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dupCourses.map(c => ({
      "كود المادة": c.code,
      "اسم المادة": c.cleanTitle || "",
      "الملف المصدر": c.sourceFile || "",
    }))), "مكررة");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBlob(new Blob([wbout], { type: "application/octet-stream" }), "تقرير_تقسيم_المواد.xlsx");
  }

  // ── results derived data ──
  const matchedCodes = new Set(extractedCourses.filter(c => c.matched).map(c => normCode(c.code)));
  const missing = [];
  if (refIndex.size > 0) refIndex.forEach((v, k) => { if (!matchedCodes.has(k)) missing.push(v); });
  const dupCount = extractedCourses.filter(c => c.dup).length;
  const noDeptCount = extractedCourses.filter(c => !c.department).length;
  const uncertainCount = extractedCourses.filter(c => c.deptMethod === "fuzzy-name").length;

const doneCourses = extractedCourses;
  const dupCourses = extractedCourses.filter(c => c.dup);

  return (
    <div>
      <div style={{ marginBottom: 22 }}>
        <div style={{
          display: "inline-block", fontSize: 11, color: "#1abc9c",
          background: "rgba(26,188,156,.12)", border: "1px solid rgba(26,188,156,.35)",
          padding: "3px 12px", borderRadius: 999, marginBottom: 10, letterSpacing: .3,
        }}>ERU · UMIS Course Evaluation</div>
        <div style={{ color: "#fff", fontSize: 22, fontWeight: 900 }}>🧩 تقسيم ملفات تقييم المقررات</div>
        <div style={{ color: "rgba(255,255,255,.55)", fontSize: 13, marginTop: 4 }}>
          يحوّل ملف الـ Excel المجمّع (كل المواد في ملف واحد) إلى ملف مستقل لكل مادة، يقارنها بقائمة المواد المتوقعة، ويصنّفها حسب القسم العلمي.
        </div>
      </div>

      {/* Step 1: reference list */}
      <div className="card" style={{ padding: 22, marginBottom: 18, borderColor: "rgba(26,188,156,.35)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
          <span style={{ width: 24, height: 24, borderRadius: "50%", background: "rgba(255,255,255,.08)", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 12, color: "#1abc9c" }}>1</span>
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 15 }}>قائمة المواد الأساسية</span>
          <span style={{ fontSize: 11, color: "#1abc9c", background: "rgba(26,188,156,.12)", border: "1px solid rgba(26,188,156,.35)", padding: "2px 10px", borderRadius: 999, fontWeight: 700 }}>المرجع الرئيسي</span>
        </div>
        <p style={{ color: "rgba(255,255,255,.5)", fontSize: 13, margin: "0 0 14px" }}>
          ده الأساس اللي هيتحدد بيه أي مادة من ضمن استبيانات فعلًا، وأي مادة ناقصة. يجب أن تحتوي على أعمدة COURSE_CODE و COURSE_DESCR_EN.
        </p>
        <DropZone icon="📋" main="اضغط لاختيار الملف أو اسحبه هنا" sub="xlsx / xls — يحتوي على أعمدة COURSE_CODE و COURSE_DESCR"
          onFiles={files => { if (files.length) setRefFile(files[0]); }} inputRef={refInputRef} />
        {refFile && <FileItem name={refFile.name} size={refFile.size} onRemove={() => setRefFile(null)} />}
      </div>

      {/* Step 2: department distribution (optional) */}
      <div className="card" style={{ padding: 22, marginBottom: 18 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
          <span style={{ width: 24, height: 24, borderRadius: "50%", background: "rgba(255,255,255,.08)", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 12, color: "#1abc9c" }}>2</span>
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 15 }}>توزيع المقررات على الأقسام (اختياري)</span>
        </div>
        <p style={{ color: "rgba(255,255,255,.5)", fontSize: 13, margin: "0 0 14px" }}>
          يحدد قسم كل مادة. المطابقة تتم بالاسم أولًا (كود المادة ممكن يختلف بين الأنظمة)، ثم بالكود.
        </p>
        <DropZone icon="🗂️" main="اضغط لاختيار الملف أو اسحبه هنا" sub="xlsx / xls — يحتوي على أعمدة اسم المقرر، كود المقرر، القسم العلمي"
          onFiles={files => { if (files.length) setDeptFile(files[0]); }} inputRef={deptInputRef} />
        {deptFile && <FileItem name={deptFile.name} size={deptFile.size} onRemove={() => setDeptFile(null)} />}
      </div>

      {/* Step 3: merged files */}
      <div className="card" style={{ padding: 22, marginBottom: 18 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
          <span style={{ width: 24, height: 24, borderRadius: "50%", background: "rgba(255,255,255,.08)", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 12, color: "#1abc9c" }}>3</span>
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 15 }}>الملفات المجمّعة من السيستم</span>
        </div>
        <p style={{ color: "rgba(255,255,255,.5)", fontSize: 13, margin: "0 0 14px" }}>يمكن رفع أكثر من ملف مرة واحدة (مثلاً ملف لكل قسم/كلية).</p>
        <DropZone icon="📦" main="اضغط لاختيار ملف أو أكثر أو اسحبها هنا" sub="xlsx / xls — نفس تنسيق تقرير الـ RDLC الصادر من UMIS"
          onFiles={files => { setMergedFiles(prev => [...prev, ...Array.from(files)]); }} multiple inputRef={mergedInputRef} />
        {mergedFiles.map((f, idx) => (
          <FileItem key={idx} name={f.name} size={f.size} onRemove={() => setMergedFiles(prev => prev.filter((_, i) => i !== idx))} />
        ))}

        <button onClick={() => setShowAdvanced(v => !v)} style={{
          background: "none", border: "none", color: "#8ea3b4", fontSize: 12, cursor: "pointer",
          marginTop: 14, padding: 0, textDecoration: "underline",
        }}>
          {showAdvanced ? "▲ إخفاء الإعدادات المتقدمة" : "⚙ إعدادات متقدمة (تخصيص نص التقسيم)"}
        </button>
        {showAdvanced && (
          <div style={{ background: "rgba(255,255,255,.04)", border: "1px solid rgba(255,255,255,.1)", borderRadius: 10, padding: 16, marginTop: 10, display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ fontSize: 12, color: "rgba(255,255,255,.5)" }}>
              الملف بيتقسّم بالبحث عن نص متكرر يفصل بين كل مادة والتانية. لو التقسيم مش شغال، غيّر النص هنا لأي نمط من الأنماط المقترحة تحت (تظهر بعد أول محاولة معالجة).
            </div>
            <label style={{ fontSize: 12, color: "#8ea3b4" }}>
              نص الفصل بين المواد
              <input className="input" value={markerText} onChange={e => setMarkerText(e.target.value)} style={{ marginTop: 4 }} />
            </label>
            <label style={{ fontSize: 12, color: "#8ea3b4" }}>
              عدد الصفوف بين بداية بيانات المادة ونص الفصل
              <input className="input" type="number" min={0} value={blockOffset}
                onChange={e => setBlockOffset(Math.max(0, parseInt(e.target.value, 10) || 0))} style={{ marginTop: 4 }} />
            </label>
            {markerCandidates.length > 0 && (
              <div>
                <div style={{ fontSize: 12, color: "#8ea3b4", marginBottom: 6 }}>أنماط متكررة موجودة فعليًا في الملف (اختر واحد وأعد المعالجة):</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  {markerCandidates.map((c, i) => (
                    <button key={i} onClick={() => setMarkerText(c.text)} className="btn btn-ghost btn-sm"
                      style={{ justifyContent: "space-between", display: "flex", textAlign: "right" }}>
                      <span>{c.text}</span>
                      <span style={{ color: "#8ea3b4" }}>× {c.count}</span>
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 20 }}>
          <button className="btn btn-primary" disabled={mergedFiles.length === 0 || processing} onClick={handleProcess}>
            {processing ? "⏳ جارٍ المعالجة..." : "🧩 قسّم الملفات وابدأ المطابقة"}
          </button>
          <button className="btn btn-ghost" onClick={resetAll}>تفريغ الكل</button>
        </div>
        {status.msg && (
          <div style={{ marginTop: 14, fontSize: 13, color: status.kind === "err" ? "#e0555a" : status.kind === "ok" ? "#3fbf85" : "rgba(255,255,255,.5)" }}>
            {status.msg}
          </div>
        )}
      </div>

      {/* Step 4: results */}
      {showResults && (
        <div className="card" style={{ padding: 22 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
            <span style={{ width: 24, height: 24, borderRadius: "50%", background: "rgba(255,255,255,.08)", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 12, color: "#1abc9c" }}>4</span>
            <span style={{ color: "#fff", fontWeight: 700, fontSize: 15 }}>النتائج</span>
          </div>

          {/* summary stats */}
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 16 }}>
            <Stat n={extractedCourses.length} l="مادة تم تقسيمها" />
            {refIndex.size > 0 && <Stat n={`${matchedCodes.size} / ${refIndex.size}`} l="تمت مطابقتها بالقائمة" cls="good" />}
            {refIndex.size > 0 && <Stat n={missing.length} l="لم تُرفع بعد" cls={missing.length ? "bad" : "good"} />}
            {dupCount > 0 && <Stat n={dupCount} l="نسخ مكررة" cls="warn" />}
            {hasDept && noDeptCount > 0 && <Stat n={noDeptCount} l="بدون قسم محدد" cls="bad" />}
            {hasDept && uncertainCount > 0 && <Stat n={uncertainCount} l="مطابقة قسم تقريبية" cls="warn" />}
          </div>

          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 20 }}>
            <button className="btn btn-primary" onClick={handleZipAll}>
              ⬇ تحميل كل المواد (ZIP {hasDept ? "مقسّم حسب القسم" : ""})
            </button>
            <button className="btn btn-ghost" onClick={handleDownloadReport}>
              📊 تنزيل تقرير Excel (موجودة / ناقصة / مكررة)
            </button>
          </div>

          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 18 }}>
            {[
              ["all", "الكل"],
              ["done", `✅ موجودة (${doneCourses.length})`],
              ...(refIndex.size > 0 ? [["missing", `❌ ناقصة (${missing.length})`]] : []),
              ["dup", `⚠ مكررة (${dupCourses.length})`],
            ].map(([key, label]) => (
              <button key={key} onClick={() => setSectionFilter(key)}
                className={`btn btn-sm ${sectionFilter === key ? "btn-primary" : "btn-ghost"}`}>
                {label}
              </button>
            ))}
          </div>

          {/* ✅ done */}
          {(sectionFilter === "all" || sectionFilter === "done") && <ResultSection
            title="✅ المواد التي تم تقسيمها"
            count={doneCourses.length}
            emptyMsg="لا توجد مواد تم تقسيمها بعد."
            accent="#3fbf85"
          >
            {doneCourses.length > 0 && (
              <table className="mini-table">
                <thead>
                  <tr>
                    <th>كود المادة</th>
                    <th>اسم المادة</th>
                    <th>القسم</th>
                    <th>الحالة</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {doneCourses.map((c, i) => (
                    <tr key={"c" + i}>
                      <td>{c.code}</td>
                      <td className="rtl-td">{c.cleanTitle || "—"}</td>
                      <td>
                        {c.department
                          ? <>{c.department}{c.deptMethod === "fuzzy-name" && <span title="مطابقة تقريبية بالاسم — راجعها" style={{ color: "#e0a336" }}> ⚠</span>}</>
                          : <span style={{ color: "rgba(255,255,255,.4)" }}>—</span>}
                      </td>
                      <td>
                        <Badge kind="done">{c.matched ? "تم تقسيمها ومطابقتها" : refIndex.size > 0 ? "تم تقسيمها — غير موجودة بالقائمة" : "تم تقسيمها"}</Badge>
                        {c.dup && <Badge kind="dup"> مكررة</Badge>}
                      </td>
                      <td><button className="btn btn-ghost btn-sm" onClick={() => downloadCourse(c)}>⬇ تحميل</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </ResultSection>}

          {/* ❌ missing */}
          {(sectionFilter === "all" || sectionFilter === "missing") && refIndex.size > 0 && (
            <ResultSection
              title="❌ المواد الناقصة (لم تُرفع بعد)"
              count={missing.length}
              emptyMsg="كل المواد الموجودة بالقائمة الأساسية تم رفعها — لا يوجد نقص."
              accent="#e0555a"
            >
              {missing.length > 0 && (
                <table className="mini-table">
                  <thead>
                    <tr>
                      <th>كود المادة</th>
                      <th>اسم المادة</th>
                      <th>الحالة</th>
                    </tr>
                  </thead>
                  <tbody>
                    {missing.map((m, i) => (
                      <tr key={"m" + i}>
                        <td>{m.code}</td>
                        <td className="rtl-td">{m.descrEn || m.descrAr || "—"}</td>
                        <td><Badge kind="missing">لم تُرفع بعد</Badge></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </ResultSection>
          )}

          {/* ⚠ duplicates */}
          {(sectionFilter === "all" || sectionFilter === "dup") && <ResultSection
            title="⚠ المواد المكررة"
            count={dupCourses.length}
            emptyMsg="لا توجد أي مادة مكررة بين الملفات المرفوعة."
            accent="#e0a336"
          >
            {dupCourses.length > 0 && (
              <table className="mini-table">
                <thead>
                  <tr>
                    <th>كود المادة</th>
                    <th>اسم المادة</th>
                    <th>الملف المصدر</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {dupCourses.map((c, i) => (
                    <tr key={"d" + i}>
                      <td>{c.code}</td>
                      <td className="rtl-td">{c.cleanTitle || "—"}</td>
                      <td className="rtl-td">{c.sourceFile}</td>
                      <td><button className="btn btn-ghost btn-sm" onClick={() => downloadCourse(c)}>⬇ تحميل</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </ResultSection>}
        </div>
      )}

      <div style={{ color: "rgba(255,255,255,.35)", fontSize: 11, textAlign: "center", marginTop: 22 }}>
        كل المعالجة تتم داخل المتصفح فقط — لا يتم رفع أي ملف لأي سيرفر.
      </div>
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

function ResultSection({ title, count, emptyMsg, accent, children }) {
  return (
    <div style={{ marginBottom: 22 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <span style={{ color: "#fff", fontWeight: 700, fontSize: 14 }}>{title}</span>
        <span style={{
          fontSize: 11, fontWeight: 700, color: accent,
          background: `${accent}22`, border: `1px solid ${accent}66`,
          padding: "1px 9px", borderRadius: 999,
        }}>{count}</span>
      </div>
      {count === 0
        ? <div style={{ color: "rgba(255,255,255,.4)", fontSize: 13, padding: "10px 2px" }}>{emptyMsg}</div>
        : <div style={{ overflowX: "auto" }}>{children}</div>}
    </div>
  );
}

function Badge({ kind, children }) {
  const styles = {
    done: { background: "rgba(63,191,133,.15)", color: "#3fbf85", border: "1px solid rgba(63,191,133,.4)" },
    dup: { background: "rgba(224,163,54,.12)", color: "#e0a336", border: "1px solid rgba(224,163,54,.4)" },
    missing: { background: "rgba(224,85,90,.12)", color: "#e0555a", border: "1px solid rgba(224,85,90,.4)" },
  };
  return (
    <span style={{ display: "inline-block", padding: "2px 10px", borderRadius: 999, fontSize: 11, fontWeight: 600, marginInlineStart: 4, ...styles[kind] }}>
      {children}
    </span>
  );
}
