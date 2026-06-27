import { useState, useRef, useCallback, useEffect, useMemo } from "react";
import { SCHEMAS, analyze, analyzeRows, prepareData, readExcel, detectSurveyType, buildComparison } from "./engine/analyze.js";
import { buildAnnualDocx, buildComparisonDocx, DEFAULT_SETTINGS } from "./engine/buildDocx.js";
import AiChat, { PROVIDERS, DEFAULT_AI_SETTINGS } from "./AiChat.jsx";
import EnhancedReportView from "./EnhancedReportView.jsx";

// ── Styles ────────────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Cairo',sans-serif;direction:rtl}

/* ── Buttons ── */
.btn{display:inline-flex;align-items:center;gap:8px;padding:12px 28px;border-radius:50px;
  font-family:'Cairo',sans-serif;font-weight:700;font-size:15px;cursor:pointer;border:none;
  transition:all .22s;white-space:nowrap}
.btn-primary{background:linear-gradient(135deg,#1abc9c,#16a085);color:#fff}
.btn-primary:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(26,188,156,.4)}
.btn-blue{background:linear-gradient(135deg,#2874a6,#1a3a5c);color:#fff}
.btn-blue:hover{transform:translateY(-2px);box-shadow:0 8px 20px rgba(40,116,166,.4)}
.btn-ghost{background:rgba(255,255,255,.1);color:#e8f0fe;border:1px solid rgba(255,255,255,.18)}
.btn-ghost:hover{background:rgba(255,255,255,.18);border-color:rgba(255,255,255,.3)}
.btn-danger{background:rgba(231,76,60,.15);color:#e87c70;border:1px solid rgba(231,76,60,.25)}
.btn-danger:hover{background:rgba(231,76,60,.28);color:#ff6b5b}
.btn-sm{padding:7px 16px;font-size:13px}

/* ── Cards ── */
.card{background:rgba(255,255,255,.07);backdrop-filter:blur(16px);
  border:1px solid rgba(255,255,255,.12);border-radius:20px}

/* ── Upload zone ── */
.upload-zone{border:2px dashed rgba(255,255,255,.3);border-radius:16px;padding:48px 32px;
  text-align:center;cursor:pointer;transition:all .25s;background:rgba(255,255,255,.04)}
.upload-zone:hover,.upload-zone.drag{border-color:#1abc9c;background:rgba(26,188,156,.08)}

/* ── Mini tables ── */
.mini-table{width:100%;border-collapse:collapse;font-size:13px;color:#e8f0fe}
.mini-table th{background:rgba(255,255,255,.1);padding:9px 10px;text-align:center;font-weight:700;color:#d0e8ff}
.mini-table td{padding:8px 10px;border-bottom:1px solid rgba(255,255,255,.06);text-align:center}
.mini-table tr:nth-child(even) td{background:rgba(255,255,255,.03)}
.mini-table tr:hover td{background:rgba(26,188,156,.06)}
.mini-table .rtl-td{text-align:right}

/* ── Form inputs ── */
.input{
  width:100%;padding:11px 16px;border-radius:10px;
  border:1px solid rgba(255,255,255,.18);
  background:#0d1f33;color:#e8f0fe;
  font-size:14px;font-family:'Cairo',sans-serif;
  outline:none;direction:rtl;
  transition:border-color .2s,box-shadow .2s,background .2s}
.input:focus{border-color:#1abc9c;background:#0f2438;box-shadow:0 0 0 3px rgba(26,188,156,.15)}
.input:hover:not(:focus){border-color:rgba(255,255,255,.3);background:#0f2236}
.input::placeholder{color:rgba(255,255,255,.28)}
.input[type="number"]{-moz-appearance:textfield}
.input[type="number"]::-webkit-inner-spin-button,.input[type="number"]::-webkit-outer-spin-button{opacity:.5}

/* ── Select dropdowns — the key fix ── */
select.input{
  appearance:none;-webkit-appearance:none;cursor:pointer;
  padding-left:38px;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='11' height='7' viewBox='0 0 11 7'%3E%3Cpath d='M1 1l4.5 4.5L10 1' stroke='%231abc9c' stroke-width='1.8' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E");
  background-repeat:no-repeat;
  background-position:left 14px center;
}
select.input option{
  background:#0d1f33;
  color:#e8f0fe;
  padding:10px 14px;
  font-family:'Cairo',sans-serif;
  font-size:14px;
}
select.input option:checked{background:#1a5276;color:#fff}

/* ── Textarea ── */
.textarea{
  width:100%;padding:11px 16px;border-radius:10px;
  border:1px solid rgba(255,255,255,.18);
  background:#0d1f33;color:#e8f0fe;
  font-size:13px;font-family:'Cairo',sans-serif;
  outline:none;direction:rtl;resize:vertical;min-height:80px;line-height:1.7;
  transition:border-color .2s,box-shadow .2s}
.textarea:focus{border-color:#1abc9c;background:#0f2438;box-shadow:0 0 0 3px rgba(26,188,156,.15)}
.textarea::placeholder{color:rgba(255,255,255,.28)}

/* ── Labels ── */
.label{color:rgba(200,220,255,.75);font-size:12.5px;display:block;margin-bottom:7px;font-weight:700;letter-spacing:.3px}

/* ── Stat cards ── */
.stat-card{
  background:rgba(255,255,255,.08);border-radius:16px;padding:18px 22px;text-align:center;
  border:1px solid rgba(255,255,255,.08);transition:background .2s}
.stat-card:hover{background:rgba(255,255,255,.12)}

/* ── Spinner ── */
.spin{animation:spin 1s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}

/* ── Tutorial overlay ── */
@keyframes tutSlideIn{from{opacity:0;transform:translateY(32px) scale(.97)}to{opacity:1;transform:translateY(0) scale(1)}}
@keyframes tutSlideNext{from{opacity:0;transform:translateX(-28px)}to{opacity:1;transform:translateX(0)}}
@keyframes tutSlidePrev{from{opacity:0;transform:translateX(28px)}to{opacity:1;transform:translateX(0)}}
.tut-card{animation:tutSlideIn .38s cubic-bezier(.22,1,.36,1)}
.tut-slide-next{animation:tutSlideNext .28s ease}
.tut-slide-prev{animation:tutSlidePrev .28s ease}
.tut-dot{width:8px;height:8px;border-radius:50%;transition:all .22s;cursor:pointer}
.tut-nav{width:40px;height:40px;border-radius:50%;border:1px solid rgba(255,255,255,.15);background:rgba(255,255,255,.07);color:rgba(255,255,255,.7);font-size:18px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .18s}
.tut-nav:hover{background:rgba(255,255,255,.15);border-color:rgba(255,255,255,.3);color:#fff}
.tut-nav:disabled{opacity:.25;cursor:not-allowed}

/* ── Processing overlay ── */
@keyframes fadeInUp{from{opacity:0;transform:translateY(24px)}to{opacity:1;transform:translateY(0)}}
@keyframes stepPop{0%{transform:scale(1)}45%{transform:scale(1.18)}100%{transform:scale(1)}}
@keyframes activePulse{0%,100%{opacity:1;box-shadow:0 0 0 0 rgba(26,188,156,.4)}60%{opacity:.7;box-shadow:0 0 0 6px rgba(26,188,156,0)}}
@keyframes progressFlow{0%{background-position:0% 50%}100%{background-position:200% 50%}}
.proc-overlay{animation:fadeInUp .35s cubic-bezier(.22,1,.36,1)}
.proc-step-active .proc-icon{animation:activePulse 1.1s ease infinite}
.proc-step-done   .proc-icon{animation:stepPop .3s ease}

/* ── Badges ── */
.badge{display:inline-block;padding:4px 14px;border-radius:20px;font-size:12px;font-weight:700}

/* ── Step bar ── */
.step-bar{display:flex;gap:0;margin-bottom:32px;border-radius:14px;overflow:hidden;
  border:1px solid rgba(255,255,255,.08)}
.step-item{flex:1;padding:11px 6px;text-align:center;font-size:11.5px;font-weight:700;
  background:rgba(255,255,255,.04);color:rgba(255,255,255,.35);transition:all .2s;
  border-left:1px solid rgba(255,255,255,.06)}
.step-item:last-child{border-left:none}
.step-item.active{background:rgba(26,188,156,.18);color:#1abc9c;border-color:rgba(26,188,156,.2)}
.step-item.done{background:rgba(26,188,156,.07);color:rgba(26,188,156,.55)}

/* ── Type/mode cards ── */
.type-card{flex:1;padding:24px 18px;border-radius:16px;border:2px solid rgba(255,255,255,.1);
  background:rgba(255,255,255,.04);cursor:pointer;text-align:center;transition:all .22s}
.type-card:hover{border-color:rgba(26,188,156,.45);background:rgba(26,188,156,.06);
  transform:translateY(-2px)}
.type-card.selected{border-color:#1abc9c;background:rgba(26,188,156,.11);
  box-shadow:0 4px 20px rgba(26,188,156,.15)}

/* ── Settings ── */
.settings-section{
  background:rgba(255,255,255,.04);border-radius:14px;padding:22px 24px;margin-bottom:18px;
  border:1px solid rgba(255,255,255,.07)}
.settings-section-title{
  color:#1abc9c;font-size:14px;font-weight:900;margin-bottom:18px;
  padding-bottom:10px;border-bottom:1px solid rgba(26,188,156,.18);letter-spacing:.3px}
.toggle-row{display:flex;align-items:center;justify-content:space-between;
  padding:10px 0;border-bottom:1px solid rgba(255,255,255,.05)}
.toggle-row:last-child{border-bottom:none}
.toggle-row:hover{background:rgba(255,255,255,.02);border-radius:8px;padding-inline:6px;margin-inline:-6px}
.toggle{width:46px;height:25px;border-radius:13px;border:none;cursor:pointer;transition:background .2s;
  position:relative;flex-shrink:0}
.toggle.on{background:#1abc9c}
.toggle.off{background:rgba(255,255,255,.18)}
.toggle::after{content:'';position:absolute;width:19px;height:19px;border-radius:50%;
  background:#fff;top:3px;transition:left .18s;box-shadow:0 1px 3px rgba(0,0,0,.3)}
.toggle.on::after{left:24px}
.toggle.off::after{left:3px}

/* ── Logo preview ── */
.logo-preview{width:60px;height:60px;object-fit:contain;border-radius:8px;
  border:1px solid rgba(255,255,255,.15);background:rgba(255,255,255,.05)}

/* ── Scrollbar ── */
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:rgba(255,255,255,.04)}
::-webkit-scrollbar-thumb{background:rgba(255,255,255,.2);border-radius:3px}
::-webkit-scrollbar-thumb:hover{background:rgba(255,255,255,.35)}
`;

const STEPS = ["نوع التقرير", "رفع الملف", "التحقق", "معاينة البيانات", "بيانات التقرير", "النتائج"];
const SETTINGS_KEY = "eruQA_settings_v1";
const AI_KEY       = "eruQA_ai_v1";

const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID ?? "";
const DRIVE_SCOPE      = "https://www.googleapis.com/auth/drive.readonly";
const GSHEETS_MIME     = "application/vnd.google-apps.spreadsheet";
const DRIVE_FILE_MIMES = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  GSHEETS_MIME,
  "text/csv",
];

function loadGisScript() {
  return new Promise(resolve => {
    if (window.google?.accounts?.oauth2) { resolve(); return; }
    const s = document.createElement("script");
    s.src = "https://accounts.google.com/gsi/client";
    s.onload = resolve;
    document.head.appendChild(s);
  });
}

const PROGRAMS = [
  "محاسبة",
  "اقتصاد",
  "إدارة",
  "علوم سياسية",
  "تكنولوجيا أعمال",
];

// ── Helpers ───────────────────────────────────────────────────────────────────
function dirColor(tier) {
  if (tier === "excellent") return "#0d6e3a";
  if (tier === "good")      return "#1a5276";
  if (tier === "neutral")   return "#784212";
  return "#922b21";
}

function downloadBlob(blob, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}

function readFileAsBuffer(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload  = e => resolve(e.target.result);
    r.onerror = () => reject(new Error("فشل قراءة الملف"));
    r.readAsArrayBuffer(file);
  });
}

function downloadPDF(result, meta) {
  const is5 = result.scaleType === "likert-5";
  const { n, axes, totalQuestions, overallAgreePct, overallMean, overallDirection, schemaLabel, byDegree, byDepartment } = result;

  const badgeClass = tier =>
    tier === "excellent" ? "badge-excellent" :
    tier === "good"      ? "badge-good" :
    tier === "neutral"   ? "badge-neutral" : "badge-low";

  const axesRows = axes.map((ax, i) => `
    <tr>
      <td class="num">${i + 1}</td>
      <td class="rtl">${ax.name}</td>
      ${is5 ? `<td class="num">${ax.axisMean ?? ""}</td>` : ""}
      <td class="num pct">${ax.axisAgreePct}%</td>
      <td class="num"><span class="badge ${badgeClass(ax.tier)}">${ax.direction}</span></td>
    </tr>`).join("");

  const degreeRows = byDegree && Object.keys(byDegree).length
    ? Object.entries(byDegree).sort((a, b) => b[1] - a[1]).map(([k, v]) =>
        `<tr><td class="num">${v}</td><td class="num">${(v / n * 100).toFixed(1)}%</td><td class="rtl">${k}</td></tr>`
      ).join("") : "";

  const deptRows = byDepartment && Object.keys(byDepartment).length
    ? Object.entries(byDepartment).sort((a, b) => b[1] - a[1]).map(([k, v]) =>
        `<tr><td class="num">${v}</td><td class="num">${(v / n * 100).toFixed(1)}%</td><td class="rtl">${k}</td></tr>`
      ).join("") : "";

  const html = `<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
<meta charset="UTF-8"/>
<title>تقرير ${schemaLabel ?? ""} — ${meta.year ?? ""}</title>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet"/>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Cairo',Arial,sans-serif;direction:rtl;color:#222;background:#fff;padding:16mm}
h1{color:#1F3864;font-size:22pt;text-align:center;margin-bottom:6pt}
h2{color:#1F3864;font-size:14pt;margin:18pt 0 8pt;padding-bottom:4pt;border-bottom:2pt solid #2E75B6}
.cover{text-align:center;margin-bottom:20pt;padding-bottom:16pt;border-bottom:1pt solid #BDD7EE}
.cover .sub{font-size:13pt;color:#555;margin:5pt 0}
.cover .prog{font-size:11pt;color:#2E75B6;font-weight:700}
.stats{display:flex;gap:10pt;margin:12pt 0}
.stat{flex:1;text-align:center;padding:9pt;border:1pt solid #BDD7EE;border-radius:6pt}
.stat .n{font-size:20pt;font-weight:900;color:#1a5276}
.stat .l{font-size:8pt;color:#666;margin-top:2pt}
table{width:100%;border-collapse:collapse;margin:8pt 0;font-size:10pt}
th{background:#1F3864;color:#fff;padding:6pt;text-align:center;font-weight:700}
th.rtl{text-align:right}
td{padding:5pt 6pt;border:1pt solid #ccc}
td.num{text-align:center}
td.rtl{text-align:right}
td.pct{font-weight:700;color:#1a5276}
tr:nth-child(even) td{background:#f5f9ff}
.total-row td{background:#BDD7EE!important;font-weight:bold}
.badge{display:inline-block;padding:2pt 7pt;border-radius:8pt;font-size:9pt;font-weight:700}
.badge-excellent{background:#d4edda;color:#155724}
.badge-good{background:#cce5ff;color:#004085}
.badge-neutral{background:#fff3cd;color:#856404}
.badge-low{background:#f8d7da;color:#721c24}
@media print{@page{size:A4;margin:15mm}body{padding:0}}
</style>
</head>
<body>
<div class="cover">
  <h1>نتائج تحليل استبيان<br/>${schemaLabel ?? ""}</h1>
  <div class="sub">العام الدراسي: ${meta.year ?? ""}</div>
  ${meta.program ? `<div class="prog">${meta.program}</div>` : ""}
</div>

<div class="stats">
  <div class="stat"><div class="n">${n}</div><div class="l">عدد المستجيبين</div></div>
  <div class="stat"><div class="n">${axes.length}</div><div class="l">عدد المحاور</div></div>
  <div class="stat"><div class="n">${totalQuestions}</div><div class="l">عدد الأسئلة</div></div>
  <div class="stat"><div class="n">${overallAgreePct}%</div><div class="l">نسبة الموافقة</div></div>
</div>

${degreeRows ? `<h2>توزيع المشاركين حسب المسمى الوظيفي</h2>
<table><thead><tr><th>العدد</th><th>النسبة</th><th class="rtl">المسمى</th></tr></thead>
<tbody>${degreeRows}</tbody></table>` : ""}

${deptRows ? `<h2>توزيع المشاركين حسب القسم</h2>
<table><thead><tr><th>العدد</th><th>النسبة</th><th class="rtl">القسم</th></tr></thead>
<tbody>${deptRows}</tbody></table>` : ""}

<h2>ملخص المحاور</h2>
<table>
<thead><tr>
  <th style="width:32pt">م</th>
  <th class="rtl">المحور</th>
  ${is5 ? "<th style=\"width:60pt\">المتوسط</th>" : ""}
  <th style="width:80pt">نسبة الموافقة</th>
  <th style="width:70pt">الاتجاه</th>
</tr></thead>
<tbody>
${axesRows}
<tr class="total-row">
  <td class="num" colspan="${is5 ? 2 : 2}" style="text-align:right;font-size:12pt">الإجمالي</td>
  ${is5 ? `<td class="num">${overallMean ?? ""}</td>` : ""}
  <td class="num pct" style="font-size:13pt">${overallAgreePct}%</td>
  <td class="num"><span class="badge badge-excellent">${overallDirection}</span></td>
</tr>
</tbody>
</table>
</body></html>`;

  const w = window.open("", "_blank", "width=900,height=700");
  w.document.write(html);
  w.document.close();
  w.focus();
  setTimeout(() => { w.print(); }, 800);
}

// Filename → program detector. Inline Arabic normalization so we don't depend on engine internals.
function detectProgramFromFilename(filename) {
  const n = String(filename ?? "")
    .replace(/[أإآ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .toLowerCase();
  if (n.includes("تكنولوجيا") && n.includes("اعمال"))  return "تكنولوجيا أعمال";
  if (n.includes("علوم")     && n.includes("سياسي")) return "علوم سياسية";
  if (n.includes("سياسي")) return "علوم سياسية";
  if (n.includes("محاسب"))  return "محاسبة";
  if (n.includes("اقتصاد")) return "اقتصاد";
  if (n.includes("ادار"))   return "إدارة";
  return null;
}

function detectYearFromFilename(filename) {
  const m = String(filename).match(/(\d{4}[-_]\d{4}|\d{4}[-_]\d{2})/);
  if (!m) return null;
  const parts = m[1].split(/[-_]/);
  if (parts[1].length === 2) return `${parts[0]}-20${parts[1]}`;
  return `${parts[0]}-${parts[1]}`;
}

function detectTypeHintFromFilename(filename) {
  const n = String(filename ?? "")
    .replace(/[أإآ]/g, "ا").replace(/ى/g, "ي").replace(/ة/g, "ه").toLowerCase();
  for (const s of Object.values(SCHEMAS)) {
    if (s.fileHints?.some(h =>
      n.includes(h.replace(/[أإآ]/g, "ا").replace(/ى/g, "ي").replace(/ة/g, "ه").toLowerCase())
    )) return s.id;
  }
  return null;
}

function uniqueCol(rows, colIdx) {
  if (colIdx == null) return [];
  const set = new Set();
  rows.forEach(r => {
    const v = r[colIdx];
    if (v != null && String(v).trim()) set.add(String(v));
  });
  return [...set].sort((a, b) => a.localeCompare(b, "ar"));
}

function computeFilteredRows(allRows, removed, metaCols, filters) {
  if (!allRows) return [];
  const { deptIdx, degreeIdx } = metaCols ?? {};
  return allRows.filter((row, idx) => {
    if (removed.has(idx)) return false;
    if (filters.dept   && deptIdx   != null && String(row[deptIdx])   !== filters.dept)   return false;
    if (filters.degree && degreeIdx != null && String(row[degreeIdx]) !== filters.degree) return false;
    return true;
  });
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
  } catch { /* ignore */ }
  return { ...DEFAULT_SETTINGS };
}

function saveSettings(s) {
  try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(s)); } catch { /* ignore */ }
}

function loadAiSettings() {
  try {
    const raw = localStorage.getItem(AI_KEY);
    if (raw) return { ...DEFAULT_AI_SETTINGS, ...JSON.parse(raw) };
  } catch { /* ignore */ }
  return { ...DEFAULT_AI_SETTINGS };
}

function saveAiSettings(s) {
  try { localStorage.setItem(AI_KEY, JSON.stringify(s)); } catch { /* ignore */ }
}

// ── SettingsPanel ─────────────────────────────────────────────────────────────
function SettingsPanel({ settings, onChange, aiSettings, onAiChange }) {
  const logoInputRef = useRef();

  const set = (key, value) => {
    const next = { ...settings, [key]: value };
    onChange(next);
    saveSettings(next);
  };

  const handleLogoUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => set("logoDataUrl", ev.target.result);
    reader.readAsDataURL(file);
  };

  const reset = () => {
    onChange({ ...DEFAULT_SETTINGS });
    saveSettings({ ...DEFAULT_SETTINGS });
  };

  const F = (key, label, placeholder, type = "text") => (
    <div>
      <label className="label">{label}</label>
      <input
        className="input"
        type={type}
        value={settings[key] ?? ""}
        placeholder={placeholder}
        onChange={e => set(key, e.target.value)}
      />
    </div>
  );

  const T = (key, label, placeholder) => (
    <div>
      <label className="label">{label}</label>
      <textarea
        className="textarea"
        value={settings[key] ?? ""}
        placeholder={placeholder}
        onChange={e => set(key, e.target.value)}
      />
    </div>
  );

  const Toggle = (key, label) => (
    <div className="toggle-row">
      <span style={{ color: "#e8f0fe", fontSize: 14 }}>{label}</span>
      <button
        className={`toggle ${settings[key] ? "on" : "off"}`}
        onClick={() => set(key, !settings[key])}
      />
    </div>
  );

  return (
    <div>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 28 }}>
        <div style={{ color: "#fff", fontSize: 22, fontWeight: 900 }}>⚙ الإعدادات</div>
        <button className="btn btn-danger btn-sm" onClick={reset}>إعادة تعيين للافتراضي</button>
      </div>

      {/* Section 1 — Institutional Identity */}
      <div className="settings-section">
        <div className="settings-section-title">🏛 هوية المؤسسة (ترويسة التقرير)</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          {F("uniName",       "اسم الجامعة",   "Egyptian Russian University")}
          {F("facultyName",   "اسم الكلية",    "Faculty of Management...")}
          {F("unitName",      "اسم الوحدة",    "Quality Assurance unit (QAU)")}
          {F("committeeName", "اسم اللجنة",    "Measurement and Evaluation Committee")}
        </div>

        {/* Logo */}
        <div style={{ marginTop: 16 }}>
          <label className="label">شعار الجامعة</label>
          <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
            {settings.logoDataUrl
              ? <img src={settings.logoDataUrl} className="logo-preview" alt="logo" />
              : <div className="logo-preview" style={{ display: "flex", alignItems: "center", justifyContent: "center", color: "rgba(255,255,255,.3)", fontSize: 11 }}>لا يوجد</div>
            }
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              <button className="btn btn-ghost btn-sm" onClick={() => logoInputRef.current.click()}>
                📎 رفع شعار
              </button>
              {settings.logoDataUrl && (
                <button className="btn btn-danger btn-sm" onClick={() => set("logoDataUrl", null)}>
                  حذف الشعار
                </button>
              )}
            </div>
            <input ref={logoInputRef} type="file" accept="image/*" style={{ display: "none" }}
              onChange={handleLogoUpload} />
          </div>
          <div style={{ color: "rgba(255,255,255,.3)", fontSize: 11, marginTop: 6 }}>
            PNG أو JPG — يظهر في الترويسة على كل صفحة. إذا لم يُرفع يُستخدم logo.png الافتراضي.
          </div>
        </div>
      </div>

      {/* Section 2 — Signature Block */}
      <div className="settings-section">
        <div className="settings-section-title">✍ كتلة التوقيع (نهاية التقرير)</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          {F("qmName", "رئيس وحدة القياس والتقويم", "د/ ...")}
          {F("quName", "رئيس وحدة الجودة",           "د/ ...")}
        </div>
      </div>

      {/* Section 3 — Footer */}
      <div className="settings-section">
        <div className="settings-section-title">📄 تذييل الصفحة</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {T("vision",  "رؤية الوحدة",       "رؤية الوحدة: ...")}
          {T("mission", "رسالة الوحدة",      "رسالة الوحدة: ...")}
          {F("email",   "البريد الإلكتروني", "E-mail: ...")}
        </div>
      </div>

      {/* Section 4 — Report Options */}
      <div className="settings-section">
        <div className="settings-section-title">📋 خيارات التقرير</div>
        {Toggle("includeEvaluatorsTable", "تضمين جدول القائم بالتقييم")}
        {Toggle("includeParticipants",    "تضمين توزيع المشاركين")}
        {Toggle("includeRecommendations", "تضمين قسم التوصيات")}
        {settings.includeRecommendations && (
          <div style={{ marginTop: 14 }}>
            <label className="label">عدد المحاور في التوصيات</label>
            <input
              className="input"
              type="number"
              min={1}
              max={17}
              style={{ width: 120 }}
              value={settings.recommendationsCount ?? 5}
              onChange={e => set("recommendationsCount", Math.max(1, parseInt(e.target.value) || 5))}
            />
          </div>
        )}
      </div>

      {/* Section 5 — Report Design */}
      <div className="settings-section">
        <div className="settings-section-title">🎨 تصميم التقرير (Word)</div>

        {/* Font */}
        <div style={{ marginBottom: 16 }}>
          <label className="label">الخط</label>
          <select
            className="input"
            value={settings.reportFont ?? "Arial"}
            onChange={e => set("reportFont", e.target.value)}
            style={{ cursor: "pointer" }}
          >
            {["Arial", "Times New Roman", "Calibri", "Tahoma"].map(f => (
              <option key={f} value={f}>{f}</option>
            ))}
          </select>
        </div>

        {/* Direction */}
        <div style={{ marginBottom: 16 }}>
          <label className="label">اتجاه التقرير</label>
          <div style={{ display: "flex", gap: 10 }}>
            {[{ val: "rtl", label: "من اليمين للشمال (RTL) ← عربي" },
              { val: "ltr", label: "من الشمال لليمين (LTR) → English" }].map(opt => (
              <button key={opt.val}
                onClick={() => set("reportDirection", opt.val)}
                style={{
                  flex: 1, padding: "9px 8px", borderRadius: 10, cursor: "pointer",
                  fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 12,
                  border: `2px solid ${settings.reportDirection === opt.val ? "#1abc9c" : "rgba(255,255,255,.15)"}`,
                  background: settings.reportDirection === opt.val ? "rgba(26,188,156,.15)" : "rgba(255,255,255,.04)",
                  color: settings.reportDirection === opt.val ? "#1abc9c" : "rgba(255,255,255,.55)",
                }}>{opt.label}</button>
            ))}
          </div>
        </div>

        {/* Table text & number alignment */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
          <div>
            <label className="label">محاذاة نص الجداول</label>
            <select className="input" value={settings.textAlign ?? "right"}
              onChange={e => set("textAlign", e.target.value)} style={{ cursor: "pointer" }}>
              <option value="right">يمين</option>
              <option value="center">وسط</option>
              <option value="left">يسار</option>
            </select>
          </div>
          <div>
            <label className="label">محاذاة الأرقام في الجداول</label>
            <select className="input" value={settings.numAlign ?? "center"}
              onChange={e => set("numAlign", e.target.value)} style={{ cursor: "pointer" }}>
              <option value="center">وسط</option>
              <option value="right">يمين</option>
              <option value="left">يسار</option>
            </select>
          </div>
        </div>

        {/* Color theme */}
        <div>
          <label className="label">ثيم الألوان</label>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            {[
              { id: "default", label: "أزرق (افتراضي)", primary: "#1F3864", secondary: "#2E75B6" },
              { id: "green",   label: "أخضر",           primary: "#1a4731", secondary: "#2d6a4f" },
              { id: "purple",  label: "بنفسجي",         primary: "#4a235a", secondary: "#7d3c98" },
              { id: "dark",    label: "رمادي",          primary: "#1c1c1c", secondary: "#444444" },
            ].map(t => (
              <button key={t.id}
                onClick={() => set("colorTheme", t.id)}
                style={{
                  display: "flex", alignItems: "center", gap: 8,
                  padding: "9px 14px", borderRadius: 10, cursor: "pointer",
                  fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 13,
                  border: `2px solid ${settings.colorTheme === t.id ? "#1abc9c" : "rgba(255,255,255,.15)"}`,
                  background: settings.colorTheme === t.id ? "rgba(26,188,156,.12)" : "rgba(255,255,255,.04)",
                  color: settings.colorTheme === t.id ? "#1abc9c" : "rgba(255,255,255,.6)",
                }}>
                <span style={{
                  display: "inline-block", width: 18, height: 18, borderRadius: 4,
                  background: `linear-gradient(135deg,${t.primary},${t.secondary})`,
                  border: "1px solid rgba(255,255,255,.2)", flexShrink: 0,
                }} />
                {t.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Section 6 — Survey Responsible Persons */}
      <div className="settings-section">
        <div className="settings-section-title">👤 المسئولون عن الاستبيانات</div>
        <div style={{ color: "rgba(255,255,255,.4)", fontSize: 12, marginBottom: 16 }}>
          يُملأ حقل "أعده" تلقائياً عند اكتشاف نوع الاستبيان
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {Object.values(SCHEMAS).map(sc => (
            <div key={sc.id} style={{ display: "grid", gridTemplateColumns: "190px 1fr", gap: 12, alignItems: "center" }}>
              <div style={{ color: "#e8f0fe", fontSize: 13, fontWeight: 600 }}>
                {sc.icon ?? "📋"} {sc.label}
              </div>
              <input
                className="input"
                value={settings.surveyResponsible?.[sc.id] ?? ""}
                placeholder="اسم المسئول — مثال: د/ أحمد محمد"
                onChange={e => set("surveyResponsible", {
                  ...(settings.surveyResponsible ?? {}),
                  [sc.id]: e.target.value,
                })}
              />
            </div>
          ))}
        </div>
      </div>

      {/* Section 7 — AI Assistant */}
      <div className="settings-section">
        <div className="settings-section-title">🤖 المساعد الذكي</div>

        {/* Provider selector */}
        <div style={{ marginBottom: 16 }}>
          <label className="label">المزود</label>
          <div style={{ display: "flex", gap: 8 }}>
            {Object.entries(PROVIDERS).map(([id, p]) => (
              <button
                key={id}
                onClick={() => {
                  const next = { ...aiSettings, provider: id, model: p.default };
                  onAiChange(next);
                  saveAiSettings(next);
                }}
                style={{
                  flex: 1, padding: "9px 6px", borderRadius: 10, cursor: "pointer",
                  fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 13,
                  border: `2px solid ${aiSettings.provider === id ? p.color : "rgba(255,255,255,.15)"}`,
                  background: aiSettings.provider === id ? `${p.color}22` : "rgba(255,255,255,.04)",
                  color: aiSettings.provider === id ? p.color : "rgba(255,255,255,.55)",
                  transition: "all .2s",
                }}
              >
                {p.icon} {p.label}
              </button>
            ))}
          </div>
        </div>

        {/* Model selector */}
        <div style={{ marginBottom: 16 }}>
          <label className="label">الموديل</label>
          <input
            className="input"
            list={`model-list-${aiSettings.provider}`}
            value={aiSettings.model}
            placeholder="اكتب اسم الموديل أو اختر من القائمة"
            onChange={e => {
              const next = { ...aiSettings, model: e.target.value };
              onAiChange(next);
              saveAiSettings(next);
            }}
            style={{ direction: "ltr", textAlign: "left" }}
          />
          <datalist id={`model-list-${aiSettings.provider}`}>
            {(PROVIDERS[aiSettings.provider]?.models || []).map(m => (
              <option key={m.id} value={m.id}>{m.label}</option>
            ))}
          </datalist>
          <div style={{ color: "rgba(255,255,255,.28)", fontSize: 11, marginTop: 6 }}>
            اكتب اسم أي موديل مدعوم أو اختر من الاقتراحات
          </div>
        </div>

        {/* API key */}
        <div style={{ marginBottom: 16 }}>
          <label className="label">مفتاح API</label>
          <input
            className="input"
            type="password"
            value={aiSettings.key}
            placeholder={`sk-... أو gsk-... حسب المزود`}
            onChange={e => {
              const next = { ...aiSettings, key: e.target.value };
              onAiChange(next);
              saveAiSettings(next);
            }}
          />
          <div style={{ color: "rgba(255,255,255,.28)", fontSize: 11, marginTop: 6 }}>
            يُخزّن في المتصفح فقط · بديلاً يمكن تعيين متغير البيئة{" "}
            <code style={{ color: "#1abc9c" }}>{aiSettings.provider.toUpperCase()}_API_KEY</code>
            {" "}في ملف .env
          </div>
        </div>

        {/* Surveys folder */}
        <div>
          <label className="label">مجلد الاستبيانات (اختياري)</label>
          <input
            className="input"
            type="text"
            value={aiSettings.surveysFolder ?? ""}
            placeholder="مثال: C:\جودة\استبيانات"
            onChange={e => {
              const next = { ...aiSettings, surveysFolder: e.target.value };
              onAiChange(next);
              saveAiSettings(next);
            }}
          />
          <div style={{ color: "rgba(255,255,255,.28)", fontSize: 11, marginTop: 6 }}>
            المسار الكامل للمجلد المحلي الذي يحتوي ملفات Excel —
            يتيح للمساعد قراءة الملفات مباشرة دون رفع يدوي
          </div>
        </div>
      </div>
    </div>
  );
}

// ── BatchItem ─────────────────────────────────────────────────────────────────
function BatchItem({ item, year, settings }) {
  const [downloading, setDownloading] = useState(false);

  const doDownload = async () => {
    if (!item.result || !item.type) return;
    setDownloading(true);
    try {
      const meta = {
        year,
        program:    item.program ?? "",
        preparedBy: settings.surveyResponsible?.[item.type] ?? "",
      };
      const blob = await buildAnnualDocx(item.result, meta, settings);
      const prog  = item.program ? `_${item.program}` : "";
      const label = SCHEMAS[item.type]?.label ?? "تقرير";
      downloadBlob(blob, `تقرير_${label}${prog}_${year}.docx`);
    } catch (e) { console.error(e); }
    finally    { setDownloading(false); }
  };

  const bg     = item.status === "done"       ? "rgba(26,188,156,.08)"
               : item.status === "error"      ? "rgba(231,76,60,.07)"
               : item.status === "processing" ? "rgba(255,255,255,.05)"
               :                               "rgba(255,255,255,.02)";
  const border = item.status === "done"       ? "rgba(26,188,156,.32)"
               : item.status === "error"      ? "rgba(231,76,60,.22)"
               : item.status === "processing" ? "rgba(255,255,255,.18)"
               :                               "rgba(255,255,255,.07)";

  return (
    <div style={{
      display: "flex", alignItems: "center", gap: 14,
      background: bg, border: `1px solid ${border}`,
      borderRadius: 12, padding: "13px 18px", transition: "background .3s",
    }}>
      {/* Status icon */}
      <div style={{ fontSize: 20, flexShrink: 0, width: 24, textAlign: "center" }}>
        {item.status === "done"       ? "✅"
       : item.status === "error"      ? "❌"
       : item.status === "processing" ? (
          <svg className="spin" width="20" height="20" viewBox="0 0 20 20" style={{ display: "block" }}>
            <circle cx="10" cy="10" r="7" fill="none" stroke="#1abc9c" strokeWidth="2.5" strokeDasharray="32 12"/>
          </svg>
        ) : "🕐"}
      </div>

      {/* Details */}
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ color: "#fff", fontWeight: 700, fontSize: 13,
                      overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          {item.file.name}
        </div>
        <div style={{ fontSize: 11, marginTop: 3 }}>
          {item.status === "done" && (
            <span style={{ display: "flex", gap: 10, flexWrap: "wrap", color: "rgba(255,255,255,.5)" }}>
              <span style={{ color: "#1abc9c", fontWeight: 700 }}>
                {SCHEMAS[item.type]?.icon} {SCHEMAS[item.type]?.label}
              </span>
              <span>·</span><span>{item.result.n} استجابة</span>
              {item.program && <><span>·</span><span>{item.program}</span></>}
              <span>·</span>
              <span style={{ color: "#1abc9c", fontWeight: 700 }}>{item.result.overallAgreePct}% موافقة</span>
            </span>
          )}
          {item.status === "processing" && <span style={{ color: "#1abc9c" }}>جاري التحليل…</span>}
          {item.status === "error"      && <span style={{ color: "#e74c3c" }}>{item.error}</span>}
          {item.status === "pending"    && <span style={{ color: "rgba(255,255,255,.3)" }}>في الانتظار</span>}
        </div>
      </div>

      {/* Download */}
      {item.status === "done" && (
        <button className="btn btn-ghost btn-sm" onClick={doDownload} disabled={downloading}
          style={{ flexShrink: 0, fontSize: 12, whiteSpace: "nowrap", opacity: downloading ? 0.6 : 1 }}>
          {downloading
            ? <svg className="spin" width="14" height="14" viewBox="0 0 14 14">
                <circle cx="7" cy="7" r="5" fill="none" stroke="white" strokeWidth="2" strokeDasharray="22 8"/>
              </svg>
            : "⬇ Word"}
        </button>
      )}
    </div>
  );
}

// ── BatchProcessor ────────────────────────────────────────────────────────────
function BatchProcessor({ files, year, onYearChange, settings, onBack }) {
  const [items, setItems] = useState(() =>
    files.map(f => ({ file: f, status: "pending", result: null, error: null, type: null, program: null }))
  );
  const [running,        setRunning]        = useState(false);
  const [allDone,        setAllDone]        = useState(false);
  const [downloadingAll, setDownloadingAll] = useState(false);

  const doneCount  = items.filter(i => i.status === "done").length;
  const errorCount = items.filter(i => i.status === "error").length;
  const totalDone  = doneCount + errorCount;

  const startProcessing = async () => {
    setRunning(true);
    for (let idx = 0; idx < files.length; idx++) {
      setItems(prev => prev.map((it, i) => i === idx ? { ...it, status: "processing" } : it));
      await new Promise(r => setTimeout(r, 30));  // let React paint
      try {
        const buf          = await readFileAsBuffer(files[idx]);
        const rows         = readExcel(buf);
        const detectedType = detectSurveyType(files[idx].name, rows[0]) ?? "faculty";
        const s            = SCHEMAS[detectedType];
        const result       = analyze(rows, s);
        const program      = detectProgramFromFilename(files[idx].name);
        setItems(prev => prev.map((it, i) => i === idx
          ? { ...it, status: "done", result, type: detectedType, program }
          : it
        ));
      } catch (err) {
        setItems(prev => prev.map((it, i) => i === idx
          ? { ...it, status: "error", error: err.message ?? "خطأ في القراءة" }
          : it
        ));
      }
    }
    setRunning(false);
    setAllDone(true);
  };

  const downloadAll = async () => {
    setDownloadingAll(true);
    for (const it of items) {
      if (it.status === "done" && it.result) {
        try {
          const meta  = { year, program: it.program ?? "", preparedBy: settings.surveyResponsible?.[it.type] ?? "" };
          const blob  = await buildAnnualDocx(it.result, meta, settings);
          const prog  = it.program ? `_${it.program}` : "";
          const label = SCHEMAS[it.type]?.label ?? "تقرير";
          downloadBlob(blob, `تقرير_${label}${prog}_${year}.docx`);
          await new Promise(r => setTimeout(r, 450));
        } catch (e) { console.error(e); }
      }
    }
    setDownloadingAll(false);
  };

  return (
    <div className="card" style={{ padding: 36 }}>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between",
                    marginBottom: 24, gap: 16 }}>
        <div>
          <div style={{ color: "#fff", fontSize: 20, fontWeight: 900 }}>
            ⚡ معالجة دفعية — {files.length} ملف
          </div>
          <div style={{ color: "rgba(255,255,255,.45)", fontSize: 13, marginTop: 4 }}>
            يُكتشف نوع كل استبيان تلقائياً · التحليل متسلسل
          </div>
        </div>
        {!running && (
          <button className="btn btn-ghost btn-sm" onClick={onBack}>→ إعادة الاختيار</button>
        )}
      </div>

      {/* Year input (before start) */}
      {!running && !allDone && (
        <div style={{ marginBottom: 24, maxWidth: 320 }}>
          <label className="label">العام الأكاديمي (مشترك لجميع الملفات)</label>
          <input className="input" value={year}
            onChange={e => onYearChange(e.target.value)}
            placeholder="مثال: 2024-2025" />
        </div>
      )}

      {/* Progress bar */}
      {(running || allDone) && (
        <div style={{ marginBottom: 22 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 7 }}>
            <span style={{ color: "rgba(255,255,255,.6)", fontSize: 12.5 }}>
              {totalDone} من {files.length} ملف
            </span>
            <span style={{ fontSize: 12.5, fontWeight: 700 }}>
              <span style={{ color: "#1abc9c" }}>{doneCount} ناجح </span>
              {errorCount > 0 && <span style={{ color: "#e74c3c" }}>· {errorCount} فشل</span>}
            </span>
          </div>
          <div style={{ background: "rgba(255,255,255,.1)", borderRadius: 6, height: 7, overflow: "hidden" }}>
            <div style={{
              height: "100%", borderRadius: 6, transition: "width .4s ease",
              background: "linear-gradient(90deg,#1abc9c,#2874a6)",
              width: `${files.length ? (totalDone / files.length) * 100 : 0}%`,
            }} />
          </div>
        </div>
      )}

      {/* File list */}
      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 28,
                    maxHeight: 480, overflowY: "auto" }}>
        {items.map((it, i) => (
          <BatchItem key={i} item={it} year={year} settings={settings} />
        ))}
      </div>

      {/* Actions */}
      <div style={{ display: "flex", justifyContent: "center", gap: 12, flexWrap: "wrap" }}>
        {!running && !allDone && (
          <button className="btn btn-primary" style={{ fontSize: 16, padding: "13px 44px" }}
            onClick={startProcessing}>
            ▶ بدء التحليل
          </button>
        )}
        {allDone && doneCount > 1 && (
          <button className="btn btn-blue" style={{ fontSize: 15, padding: "12px 32px" }}
            disabled={downloadingAll} onClick={downloadAll}>
            {downloadingAll
              ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="white" strokeWidth="2" strokeDasharray="28 10"/></svg> جاري التحميل…</>
              : `⬇ تحميل الكل (${doneCount} ملف Word)`}
          </button>
        )}
        {allDone && (
          <button className="btn btn-ghost" style={{ fontSize: 15, padding: "12px 28px" }}
            onClick={onBack}>
            ↩ رفع ملفات جديدة
          </button>
        )}
      </div>
    </div>
  );
}

// ── DriveDashboard ────────────────────────────────────────────────────────────
function DriveDashboard({ files, token, onSelectFile }) {
  const [stats, setStats] = useState(() =>
    files.map(f => ({
      id: f.id, name: f.name, mimeType: f.mimeType,
      modifiedTime: f.modifiedTime,
      type: detectTypeHintFromFilename(f.name),
      year: detectYearFromFilename(f.name),
      program: detectProgramFromFilename(f.name),
      count: null, status: "pending",
    }))
  );
  const [loading, setLoading] = useState(false);
  const [done, setDone] = useState(0);

  const loadStats = async () => {
    if (loading) return;
    setLoading(true);
    setDone(0);
    setStats(prev => prev.map(s => ({ ...s, status: "loading", count: null })));
    const BATCH = 4;
    for (let i = 0; i < files.length; i += BATCH) {
      const batch = files.slice(i, i + BATCH);
      await Promise.all(batch.map(async (f, bi) => {
        const idx = i + bi;
        try {
          const isGSheet = f.mimeType === GSHEETS_MIME;
          const url = isGSheet
            ? `https://www.googleapis.com/drive/v3/files/${f.id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
            : `https://www.googleapis.com/drive/v3/files/${f.id}?alt=media`;
          const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
          if (!res.ok) throw new Error();
          const buf = await res.arrayBuffer();
          const rows = readExcel(buf);
          const count = Math.max(0, rows.length - 1);
          setStats(prev => prev.map((s, si) => si === idx ? { ...s, count, status: "done" } : s));
        } catch {
          setStats(prev => prev.map((s, si) => si === idx ? { ...s, status: "error" } : s));
        }
        setDone(prev => prev + 1);
      }));
    }
    setLoading(false);
  };

  const doneStats   = stats.filter(s => s.status === "done");
  const totalResp   = doneStats.reduce((a, s) => a + (s.count ?? 0), 0);

  const aggregate = (key) => {
    const map = {};
    stats.forEach(s => {
      const k = s[key] ?? "غير محدد";
      if (!map[k]) map[k] = { files: 0, responses: 0 };
      map[k].files++;
      if (s.count != null) map[k].responses += s.count;
    });
    return map;
  };
  const byType    = aggregate("type");
  const byProgram = aggregate("program");
  const byYear    = aggregate("year");

  const maxProg = Math.max(...Object.values(byProgram).map(v => v.files), 1);
  const maxYear = Math.max(...Object.values(byYear).map(v => v.files), 1);

  const barBtn = (active, label, onClick) => (
    <button onClick={onClick} style={{
      padding: "6px 12px", borderRadius: 8, border: "none", cursor: "pointer",
      fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 12,
      background: active ? "rgba(26,188,156,.25)" : "rgba(255,255,255,.08)",
      color: active ? "#1abc9c" : "rgba(255,255,255,.5)",
      borderBottom: `2px solid ${active ? "#1abc9c" : "transparent"}`,
    }}>{label}</button>
  );

  return (
    <div style={{ textAlign: "right" }}>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
        <div>
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 13 }}>
            {files.length} ملف
          </span>
          {doneStats.length > 0 && (
            <span style={{ color: "#1abc9c", fontWeight: 700, fontSize: 13, marginRight: 10 }}>
              · {totalResp.toLocaleString()} استجابة إجمالاً
            </span>
          )}
          {loading && (
            <div style={{ color: "rgba(255,255,255,.4)", fontSize: 11, marginTop: 2 }}>
              جاري تحليل {done} من {files.length}…
            </div>
          )}
        </div>
        {!loading && done < files.length && (
          <button onClick={loadStats} style={{
            padding: "8px 16px", background: "linear-gradient(135deg,#1abc9c,#16a085)",
            border: "none", borderRadius: 20, color: "#fff",
            fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 12, cursor: "pointer",
          }}>
            📊 تحليل الملفات ({files.length})
          </button>
        )}
        {loading && (
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <svg className="spin" width="16" height="16" viewBox="0 0 16 16">
              <circle cx="8" cy="8" r="6" fill="none" stroke="#1abc9c" strokeWidth="2.5" strokeDasharray="24 8"/>
            </svg>
            <span style={{ color: "rgba(255,255,255,.5)", fontSize: 12 }}>
              {Math.round(done / files.length * 100)}%
            </span>
          </div>
        )}
      </div>

      {/* Progress bar */}
      {(loading || done > 0) && (
        <div style={{ background: "rgba(255,255,255,.1)", borderRadius: 4, height: 4, marginBottom: 14, overflow: "hidden" }}>
          <div style={{
            height: "100%", background: "linear-gradient(90deg,#1abc9c,#2874a6)", borderRadius: 4,
            transition: "width .4s", width: `${Math.round(done / files.length * 100)}%`,
          }} />
        </div>
      )}

      {/* Cards by survey type */}
      <div style={{ marginBottom: 16 }}>
        <div style={{ color: "rgba(255,255,255,.45)", fontSize: 11, fontWeight: 700, marginBottom: 8, letterSpacing: .4 }}>
          حسب نوع الاستبيان
        </div>
        <div style={{ display: "flex", gap: 8, overflowX: "auto", paddingBottom: 4 }}>
          {Object.entries(byType).map(([tid, d]) => {
            const sc = SCHEMAS[tid];
            return (
              <div key={tid} style={{
                background: "rgba(255,255,255,.06)", border: "1px solid rgba(255,255,255,.1)",
                borderRadius: 12, padding: "12px 14px", textAlign: "center", minWidth: 110, flexShrink: 0,
              }}>
                <div style={{ fontSize: 20, marginBottom: 3 }}>{sc?.icon ?? "📋"}</div>
                <div style={{ color: "#e8f0fe", fontWeight: 700, fontSize: 11, marginBottom: 4,
                              whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 100 }}>
                  {sc?.label ?? "غير معروف"}
                </div>
                <div style={{ color: "#1abc9c", fontWeight: 900, fontSize: 20 }}>{d.files}</div>
                <div style={{ color: "rgba(255,255,255,.4)", fontSize: 10 }}>ملف</div>
                {d.responses > 0 && (
                  <>
                    <div style={{ color: "#f0b27a", fontWeight: 700, fontSize: 14, marginTop: 4 }}>
                      {d.responses.toLocaleString()}
                    </div>
                    <div style={{ color: "rgba(255,255,255,.35)", fontSize: 10 }}>استجابة</div>
                  </>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* Bar charts */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
        {/* By program */}
        <div style={{ background: "rgba(255,255,255,.04)", borderRadius: 12, padding: "12px 14px", border: "1px solid rgba(255,255,255,.07)" }}>
          <div style={{ color: "rgba(255,255,255,.45)", fontSize: 11, fontWeight: 700, marginBottom: 10, letterSpacing: .4 }}>البرامج</div>
          {Object.entries(byProgram).sort((a, b) => b[1].files - a[1].files).map(([prog, d]) => (
            <div key={prog} style={{ marginBottom: 7 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
                <span style={{ color: "#e8f0fe", fontSize: 11 }}>{prog}</span>
                <span style={{ color: "#1abc9c", fontSize: 11, fontWeight: 700 }}>
                  {d.files}{d.responses > 0 ? ` · ${d.responses.toLocaleString()}` : ""}
                </span>
              </div>
              <div style={{ background: "rgba(255,255,255,.08)", borderRadius: 3, height: 6, overflow: "hidden" }}>
                <div style={{
                  height: "100%", borderRadius: 3, transition: "width .5s",
                  background: "linear-gradient(90deg,#1abc9c,#2874a6)",
                  width: `${d.files / maxProg * 100}%`,
                }} />
              </div>
            </div>
          ))}
        </div>

        {/* By year */}
        <div style={{ background: "rgba(255,255,255,.04)", borderRadius: 12, padding: "12px 14px", border: "1px solid rgba(255,255,255,.07)" }}>
          <div style={{ color: "rgba(255,255,255,.45)", fontSize: 11, fontWeight: 700, marginBottom: 10, letterSpacing: .4 }}>السنوات الدراسية</div>
          {Object.entries(byYear).sort((a, b) => b[0].localeCompare(a[0])).map(([yr, d]) => (
            <div key={yr} style={{ marginBottom: 7 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
                <span style={{ color: "#e8f0fe", fontSize: 11 }}>{yr}</span>
                <span style={{ color: "#a29bfe", fontSize: 11, fontWeight: 700 }}>
                  {d.files}{d.responses > 0 ? ` · ${d.responses.toLocaleString()}` : ""}
                </span>
              </div>
              <div style={{ background: "rgba(255,255,255,.08)", borderRadius: 3, height: 6, overflow: "hidden" }}>
                <div style={{
                  height: "100%", borderRadius: 3, transition: "width .5s",
                  background: "linear-gradient(90deg,#8e44ad,#2874a6)",
                  width: `${d.files / maxYear * 100}%`,
                }} />
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Detailed table */}
      <div style={{ color: "rgba(255,255,255,.45)", fontSize: 11, fontWeight: 700, marginBottom: 8, letterSpacing: .4 }}>
        تفاصيل الملفات
      </div>
      <div style={{ maxHeight: 260, overflowY: "auto", border: "1px solid rgba(255,255,255,.1)", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr style={{ background: "rgba(255,255,255,.08)", position: "sticky", top: 0, zIndex: 1 }}>
              {["الملف","النوع","البرنامج","السنة","الاستجابات",""].map((h, i) => (
                <th key={i} style={{ padding: "7px 10px", color: "rgba(255,255,255,.55)", fontWeight: 700,
                                     textAlign: i === 0 ? "right" : "center", whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {stats.map((s, i) => (
              <tr key={s.id} style={{ borderTop: "1px solid rgba(255,255,255,.05)", cursor: "default" }}>
                <td style={{ padding: "7px 10px", color: "#e8f0fe", maxWidth: 170,
                             overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {s.name}
                </td>
                <td style={{ padding: "7px 10px", textAlign: "center" }}>
                  {s.type
                    ? <span style={{ fontSize: 10, color: "#1abc9c", background: "rgba(26,188,156,.12)", borderRadius: 6, padding: "2px 7px", whiteSpace: "nowrap" }}>
                        {SCHEMAS[s.type]?.label ?? s.type}
                      </span>
                    : <span style={{ color: "rgba(255,255,255,.2)" }}>—</span>}
                </td>
                <td style={{ padding: "7px 10px", textAlign: "center", color: "#d6eaf8", fontSize: 11, whiteSpace: "nowrap" }}>
                  {s.program ?? <span style={{ color: "rgba(255,255,255,.2)" }}>—</span>}
                </td>
                <td style={{ padding: "7px 10px", textAlign: "center", color: "#d6eaf8", fontSize: 11, whiteSpace: "nowrap" }}>
                  {s.year ?? <span style={{ color: "rgba(255,255,255,.2)" }}>—</span>}
                </td>
                <td style={{ padding: "7px 10px", textAlign: "center" }}>
                  {s.status === "loading"
                    ? <svg className="spin" width="13" height="13" viewBox="0 0 13 13">
                        <circle cx="6.5" cy="6.5" r="5" fill="none" stroke="#1abc9c" strokeWidth="2" strokeDasharray="18 8"/>
                      </svg>
                    : s.status === "done"
                    ? <span style={{ color: "#1abc9c", fontWeight: 700 }}>{s.count?.toLocaleString()}</span>
                    : s.status === "error"
                    ? <span style={{ color: "#e74c3c", fontSize: 10 }}>خطأ</span>
                    : <span style={{ color: "rgba(255,255,255,.2)" }}>—</span>}
                </td>
                <td style={{ padding: "6px 8px", textAlign: "center" }}>
                  <button onClick={() => onSelectFile(files[i])} style={{
                    padding: "3px 10px", background: "rgba(26,188,156,.13)",
                    border: "1px solid rgba(26,188,156,.3)", borderRadius: 7,
                    color: "#1abc9c", cursor: "pointer", fontSize: 11,
                    fontFamily: "'Cairo',sans-serif", fontWeight: 700,
                  }}>تحليل</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── LoadingOverlay ────────────────────────────────────────────────────────────
// ── TutorialOverlay ───────────────────────────────────────────────────────────
const TUTORIAL_SLIDES = [
  {
    icon: "🎉",
    color: "#1abc9c",
    title: "مرحباً بك في محلل الاستبيانات",
    subtitle: "دليل سريع – 6 خطوات",
    body: "هذا التطبيق يحلّل نتائج استبيانات وحدة الجودة ويولّد تقارير Word كاملة بالإحصاءات والتوصيات.",
    tips: [
      "يدعم 5 أنواع استبيانات: طلاب، هيئة تدريس، هيئة معاونة، خريجين، منسقي برامج",
      "يولّد تقارير Word جاهزة للطباعة مع كل الجداول والمحاور",
      "يدعم تقارير المقارنة بين 2 أو 3 سنوات دراسية",
      "يمكن ربطه بـ Google Drive مباشرةً",
    ],
  },
  {
    icon: "📂",
    color: "#2874a6",
    title: "الخطوة ١ – رفع ملف الاستبيان",
    subtitle: "رفع ملف · Google Drive",
    body: "ارفع ملف Excel الصادر من Google Forms عبر السحب والإفلات، أو اربط حسابك على Google Drive واختر الملف مباشرةً.",
    tips: [
      "يدعم .xlsx / .xls / .csv",
      "يكتشف نوع الاستبيان تلقائياً من اسم الملف والأعمدة",
      "لو اخترت أكثر من ملف تفتح وضع المعالجة الدفعية",
      "في Drive يمكن تصفية الملفات حسب النوع والسنة والبرنامج",
    ],
  },
  {
    icon: "✅",
    color: "#1abc9c",
    title: "الخطوة ٢ – التحقق من نوع الاستبيان",
    subtitle: "مراجعة الاكتشاف التلقائي",
    body: "يعرض التطبيق نوع الاستبيان الذي اكتشفه مع مستوى الثقة. يمكنك تأكيده أو تغييره يدوياً.",
    tips: [
      "الاكتشاف يعتمد على اسم الملف وترويسات الأعمدة",
      "لو النوع مش صح، غيّره من القائمة المنسدلة",
      "اختر نمط الاستبيان: مقياس 3 درجات أو 5 درجات",
    ],
  },
  {
    icon: "📋",
    color: "#9b59b6",
    title: "الخطوة ٣ – نوع التقرير",
    subtitle: "سنوي · مقارنة سنتين · مقارنة 3 سنوات",
    body: "اختر إذا كنت تريد تقريراً لسنة واحدة، أو مقارنة نتائج سنتين أو ثلاث سنوات في نفس التقرير.",
    tips: [
      "التقرير السنوي يحلّل ملفاً واحداً بالكامل",
      "تقرير المقارنة يرصد التطور ويظهر أسهم تحسّن/تراجع",
      "في المقارنة ارفع ملف لكل سنة",
    ],
  },
  {
    icon: "👁",
    color: "#e67e22",
    title: "الخطوة ٤ – معاينة وتنقية البيانات",
    subtitle: "مراجعة الصفوف قبل التحليل",
    body: "راجع بيانات المشاركين وأزِل الصفوف الخاطئة أو المكررة قبل الشروع في التحليل.",
    tips: [
      "التكرارات تُحذف تلقائياً في الخطوة السابقة",
      "يمكن تصفية البيانات حسب القسم أو الدرجة الوظيفية",
      "حدّد الصفوف يدوياً ثم اضغط 'حذف المحدد'",
    ],
  },
  {
    icon: "📄",
    color: "#1abc9c",
    title: "الخطوات ٥ و ٦ – البيانات والنتائج",
    subtitle: "أدخل معلومات التقرير ثم حمّله",
    body: "أدخل بيانات التقرير (السنة، البرنامج، المُعد، المراجع) ثم شاهد النتائج الكاملة وحمّل ملف Word.",
    tips: [
      "البيانات المُدخلة تُحفظ تلقائياً لأول مرة تُعبّأ",
      "التقرير يشمل: المقدمة، المشاركين، النتائج بالمحاور، الملخص، التوصيات",
      "يمكن تخصيص الألوان والخطوط وتوقيع التقرير من الإعدادات ⚙",
      "زر 'عرض محسّن' يفتح عرضاً تفاعلياً داخل التطبيق",
    ],
  },
];

function TutorialOverlay({ onClose }) {
  const [idx,  setIdx]  = useState(0);
  const [dir,  setDir]  = useState("next"); // "next" | "prev"
  const [anim, setAnim] = useState(false);
  const slide = TUTORIAL_SLIDES[idx];
  const total = TUTORIAL_SLIDES.length;

  const go = (next) => {
    if (next < 0 || next >= total) return;
    setDir(next > idx ? "next" : "prev");
    setAnim(false);
    // micro-delay so the class is stripped then re-added
    requestAnimationFrame(() => {
      setIdx(next);
      setAnim(true);
    });
  };

  return (
    <div
      onClick={e => { if (e.target === e.currentTarget) onClose(); }}
      style={{
        position: "fixed", inset: 0, zIndex: 9100,
        background: "rgba(4,12,24,.88)", backdropFilter: "blur(16px)",
        display: "flex", alignItems: "center", justifyContent: "center",
        fontFamily: "'Cairo',sans-serif", padding: 20,
      }}
    >
      <div className="tut-card" style={{
        background: "linear-gradient(155deg,rgba(255,255,255,.08),rgba(255,255,255,.03))",
        border: "1px solid rgba(255,255,255,.13)",
        borderRadius: 28, width: "100%", maxWidth: 560,
        boxShadow: "0 48px 96px rgba(0,0,0,.55), 0 0 0 1px rgba(255,255,255,.04)",
        overflow: "hidden",
        position: "relative",
      }}>
        {/* Accent top bar */}
        <div style={{ height: 4, background: `linear-gradient(90deg,${slide.color},${slide.color}88)` }} />

        {/* Close */}
        <button onClick={onClose} style={{
          position: "absolute", top: 16, left: 16,
          width: 32, height: 32, borderRadius: 8,
          border: "1px solid rgba(255,255,255,.15)",
          background: "rgba(255,255,255,.07)",
          color: "rgba(255,255,255,.5)", fontSize: 16,
          cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
          transition: "all .18s",
        }}
          onMouseOver={e => { e.currentTarget.style.background = "rgba(231,76,60,.2)"; e.currentTarget.style.color = "#ff6b6b"; }}
          onMouseOut={e  => { e.currentTarget.style.background = "rgba(255,255,255,.07)"; e.currentTarget.style.color = "rgba(255,255,255,.5)"; }}
        >✕</button>

        {/* Slide content */}
        <div
          key={idx}
          className={anim ? (dir === "next" ? "tut-slide-next" : "tut-slide-prev") : ""}
          style={{ padding: "28px 36px 24px", textAlign: "center" }}
        >
          {/* Icon */}
          <div style={{
            width: 72, height: 72, borderRadius: 20, margin: "0 auto 18px",
            background: `linear-gradient(135deg,${slide.color}30,${slide.color}12)`,
            border: `1px solid ${slide.color}45`,
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: 34,
          }}>{slide.icon}</div>

          {/* Title */}
          <div style={{ color: "#fff", fontWeight: 900, fontSize: 20, marginBottom: 4, lineHeight: 1.4 }}>
            {slide.title}
          </div>
          <div style={{ color: slide.color, fontSize: 12, fontWeight: 700, marginBottom: 16, letterSpacing: .5 }}>
            {slide.subtitle}
          </div>

          {/* Body */}
          <div style={{
            color: "rgba(255,255,255,.65)", fontSize: 14, lineHeight: 1.8,
            marginBottom: 20, textAlign: "right",
          }}>{slide.body}</div>

          {/* Tips */}
          <div style={{
            background: "rgba(255,255,255,.05)", borderRadius: 14,
            border: "1px solid rgba(255,255,255,.08)",
            padding: "14px 18px", textAlign: "right",
          }}>
            {slide.tips.map((t, i) => (
              <div key={i} style={{
                display: "flex", alignItems: "flex-start", gap: 10,
                padding: "6px 0",
                borderBottom: i < slide.tips.length - 1 ? "1px solid rgba(255,255,255,.05)" : "none",
              }}>
                <div style={{
                  width: 20, height: 20, borderRadius: 6, flexShrink: 0, marginTop: 1,
                  background: `${slide.color}25`, border: `1px solid ${slide.color}40`,
                  display: "flex", alignItems: "center", justifyContent: "center",
                  fontSize: 10, color: slide.color, fontWeight: 900,
                }}>{i + 1}</div>
                <div style={{ color: "rgba(255,255,255,.7)", fontSize: 13, lineHeight: 1.7 }}>{t}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Footer navigation */}
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "16px 36px 24px",
          borderTop: "1px solid rgba(255,255,255,.07)",
        }}>
          {/* Prev */}
          <button className="tut-nav" onClick={() => go(idx - 1)} disabled={idx === 0}>‹</button>

          {/* Dots */}
          <div style={{ display: "flex", gap: 7, alignItems: "center" }}>
            {TUTORIAL_SLIDES.map((_, i) => (
              <div
                key={i}
                className="tut-dot"
                onClick={() => go(i)}
                style={{
                  background: i === idx ? slide.color : "rgba(255,255,255,.2)",
                  transform: i === idx ? "scale(1.35)" : "scale(1)",
                }}
              />
            ))}
          </div>

          {/* Next / Done */}
          {idx < total - 1 ? (
            <button className="tut-nav" onClick={() => go(idx + 1)}>›</button>
          ) : (
            <button
              onClick={onClose}
              style={{
                padding: "8px 20px", borderRadius: 22, border: "none",
                background: `linear-gradient(135deg,#1abc9c,#16a085)`,
                color: "#fff", fontFamily: "'Cairo',sans-serif",
                fontWeight: 700, fontSize: 13, cursor: "pointer",
              }}
            >ابدأ الآن ✓</button>
          )}
        </div>
      </div>
    </div>
  );
}

// ── ProcessingOverlay ─────────────────────────────────────────────────────────
function ProcessingOverlay({ steps, filename }) {
  const doneCount = steps.filter(s => s.status === "done").length;
  const pct = Math.round((doneCount / steps.length) * 100);

  return (
    <div style={{
      position: "fixed", inset: 0, zIndex: 8900,
      background: "rgba(5,14,28,.93)", backdropFilter: "blur(14px)",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "'Cairo',sans-serif",
    }}>
      <div className="proc-overlay" style={{
        background: "linear-gradient(145deg,rgba(255,255,255,.07),rgba(255,255,255,.03))",
        border: "1px solid rgba(255,255,255,.13)",
        borderRadius: 24, padding: "36px 44px",
        minWidth: 380, maxWidth: 460,
        boxShadow: "0 40px 80px rgba(0,0,0,.6), 0 0 0 1px rgba(26,188,156,.08)",
      }}>
        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{
            width: 64, height: 64, borderRadius: 18,
            background: "linear-gradient(135deg,#1abc9c22,#2874a622)",
            border: "1px solid rgba(26,188,156,.3)",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: 28, margin: "0 auto 14px",
          }}>📊</div>
          <div style={{ color: "#fff", fontWeight: 900, fontSize: 18, marginBottom: 6 }}>جاري المعالجة…</div>
          {filename && (
            <div style={{
              color: "rgba(255,255,255,.35)", fontSize: 11,
              maxWidth: 320, margin: "0 auto",
              overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
            }}>{filename}</div>
          )}
        </div>

        {/* Steps */}
        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 24 }}>
          {steps.map((s, i) => (
            <div
              key={i}
              className={`proc-step-${s.status}`}
              style={{
                display: "flex", alignItems: "center", gap: 14,
                padding: "10px 16px", borderRadius: 12,
                background: s.status === "done"   ? "rgba(26,188,156,.1)"  :
                            s.status === "active" ? "rgba(255,255,255,.07)" : "transparent",
                border: `1px solid ${
                  s.status === "done"   ? "rgba(26,188,156,.28)"  :
                  s.status === "active" ? "rgba(255,255,255,.12)" : "transparent"
                }`,
                transition: "background .3s, border-color .3s",
              }}
            >
              {/* Icon */}
              <div className="proc-icon" style={{
                width: 28, height: 28, borderRadius: 8, flexShrink: 0,
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: 14,
                background: s.status === "done"   ? "rgba(26,188,156,.22)" :
                            s.status === "active" ? "rgba(255,255,255,.1)"  : "rgba(255,255,255,.04)",
                border: `1px solid ${
                  s.status === "done"   ? "rgba(26,188,156,.4)"  :
                  s.status === "active" ? "rgba(255,255,255,.18)" : "rgba(255,255,255,.07)"
                }`,
              }}>
                {s.status === "done"   ? <span style={{ color: "#1abc9c", fontWeight: 900, fontSize: 13 }}>✓</span> :
                 s.status === "active" ? (
                   <svg className="spin" width="14" height="14" viewBox="0 0 14 14">
                     <circle cx="7" cy="7" r="5.5" fill="none" stroke="rgba(255,255,255,.2)" strokeWidth="2"/>
                     <path d="M7 1.5a5.5 5.5 0 0 1 5.5 5.5" fill="none" stroke="#1abc9c" strokeWidth="2" strokeLinecap="round"/>
                   </svg>
                 ) : <span style={{ color: "rgba(255,255,255,.18)", fontSize: 11 }}>○</span>}
              </div>

              {/* Label */}
              <div style={{
                color: s.status === "done"   ? "#1abc9c"               :
                       s.status === "active" ? "#fff"                   : "rgba(255,255,255,.28)",
                fontSize: 14,
                fontWeight: s.status === "active" ? 700 : 600,
                transition: "color .3s",
                flex: 1,
              }}>{s.label}</div>

              {/* Timing badge for done */}
              {s.status === "done" && (
                <div style={{
                  fontSize: 10, color: "rgba(26,188,156,.6)",
                  background: "rgba(26,188,156,.1)", borderRadius: 6,
                  padding: "2px 7px", fontWeight: 700,
                }}>تم</div>
              )}
            </div>
          ))}
        </div>

        {/* Progress bar */}
        <div style={{ height: 5, background: "rgba(255,255,255,.07)", borderRadius: 3, overflow: "hidden" }}>
          <div style={{
            height: "100%", borderRadius: 3,
            background: "linear-gradient(90deg,#1abc9c,#2874a6,#1abc9c)",
            backgroundSize: "200% 100%",
            width: `${pct}%`,
            transition: "width .45s cubic-bezier(.4,0,.2,1)",
            animation: pct < 100 ? "progressFlow 1.8s linear infinite" : "none",
          }}/>
        </div>
        <div style={{ textAlign: "center", marginTop: 8, color: "rgba(255,255,255,.3)", fontSize: 11 }}>
          {pct}%
        </div>
      </div>
    </div>
  );
}

function LoadingOverlay({ message = "جاري المعالجة…" }) {
  return (
    <div style={{
      position: "fixed", inset: 0, zIndex: 8500,
      background: "rgba(9,20,38,.90)", backdropFilter: "blur(12px)",
      display: "flex", flexDirection: "column", alignItems: "center",
      justifyContent: "center", gap: 24, fontFamily: "'Cairo',sans-serif",
    }}>
      <div style={{
        width: 68, height: 68, borderRadius: "50%",
        border: "4px solid rgba(26,188,156,.2)",
        borderTop: "4px solid #1abc9c",
        animation: "spin 0.85s linear infinite",
      }} />
      <div style={{ color: "#fff", fontSize: 18, fontWeight: 700, textAlign: "center" }}>{message}</div>
      <div style={{ color: "rgba(255,255,255,.35)", fontSize: 12 }}>يرجى الانتظار…</div>
    </div>
  );
}

// ── StepBar ───────────────────────────────────────────────────────────────────
function StepBar({ step }) {
  return (
    <div className="step-bar">
      {STEPS.map((label, i) => (
        <div key={i} className={`step-item ${i === step ? "active" : i < step ? "done" : ""}`}>
          {i < step ? "✓ " : ""}{label}
        </div>
      ))}
    </div>
  );
}

// ── SurveyTypePicker ──────────────────────────────────────────────────────────
function SurveyTypePicker({ value, onChange }) {
  const TYPES = Object.values(SCHEMAS).map(s => {
    const isL5 = s.scale.type === "likert-5";
    const scaleLabel = isL5 ? "مقياس 5 درجات" : "مقياس 3 درجات";
    const autoDesc = s.axes.length === 1
      ? `${s.axes[0].questions.length} عبارة – ${scaleLabel}`
      : `${s.axes.length} محوراً – ${scaleLabel}`;
    return {
      id:    s.id,
      icon:  s.icon  || "📋",
      label: s.label,
      desc:  s.desc  || autoDesc,
    };
  });
  return (
    <div className="card" style={{ padding: 36 }}>
      <div style={{ color: "#fff", fontSize: 22, fontWeight: 900, marginBottom: 8, textAlign: "center" }}>اختر نوع الاستبيان</div>
      <div style={{ color: "rgba(255,255,255,.45)", fontSize: 13, marginBottom: 28, textAlign: "center" }}>
        يمكن اكتشاف النوع تلقائياً من الملف أيضاً
      </div>
      <div style={{ display: "flex", gap: 16 }}>
        {TYPES.map(t => (
          <div key={t.id} className={`type-card ${value === t.id ? "selected" : ""}`} onClick={() => onChange(t.id)}>
            <div style={{ fontSize: 40, marginBottom: 10 }}>{t.icon}</div>
            <div style={{ color: "#fff", fontWeight: 700, fontSize: 15, marginBottom: 6 }}>{t.label}</div>
            <div style={{ color: "rgba(255,255,255,.45)", fontSize: 12 }}>{t.desc}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── ReportModePicker ──────────────────────────────────────────────────────────
function ReportModePicker({ value, onChange }) {
  const MODES = [
    { id: "annual",   icon: "📊", label: "تقرير سنة واحدة", desc: "تحليل عام دراسي محدد" },
    { id: "compare2", icon: "📈", label: "مقارنة سنتين",     desc: "مقارنة عامين دراسيين" },
    { id: "compare3", icon: "📉", label: "مقارنة 3 سنوات",   desc: "مقارنة ثلاثة أعوام" },
  ];
  return (
    <div className="card" style={{ padding: 36 }}>
      <div style={{ color: "#fff", fontSize: 22, fontWeight: 900, marginBottom: 28, textAlign: "center" }}>نوع التقرير</div>
      <div style={{ display: "flex", gap: 16 }}>
        {MODES.map(m => (
          <div key={m.id} className={`type-card ${value === m.id ? "selected" : ""}`} onClick={() => onChange(m.id)}>
            <div style={{ fontSize: 36, marginBottom: 10 }}>{m.icon}</div>
            <div style={{ color: "#fff", fontWeight: 700, fontSize: 15, marginBottom: 6 }}>{m.label}</div>
            <div style={{ color: "rgba(255,255,255,.45)", fontSize: 12 }}>{m.desc}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── FileSlot ──────────────────────────────────────────────────────────────────
function FileSlot({ slot, onChange, label, driveToken, driveFiles, driveLoading, onConnectDrive }) {
  const ref = useRef();
  const [slotTab, setSlotTab] = useState("local");
  const [driveSearch, setDriveSearch] = useState("");
  const [downloading, setDownloading] = useState(false);
  const [slotDriveSelected, setSlotDriveSelected] = useState(null);

  const pickFromDrive = async (file) => {
    if (!driveToken || downloading) return;
    setDownloading(true);
    try {
      const isGSheet = file.mimeType === GSHEETS_MIME;
      const url = isGSheet
        ? `https://www.googleapis.com/drive/v3/files/${file.id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
        : `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`;
      const res = await fetch(url, { headers: { Authorization: `Bearer ${driveToken}` } });
      if (!res.ok) throw new Error(`خطأ ${res.status}`);
      const arrayBuf = await res.arrayBuffer();
      const filename = isGSheet ? (file.name.endsWith(".xlsx") ? file.name : file.name + ".xlsx") : file.name;
      let detectedType = detectTypeHintFromFilename(filename);
      try {
        const rows = readExcel(arrayBuf);
        detectedType = detectSurveyType(filename, rows[0]) ?? detectedType;
      } catch { /* keep hint */ }
      const prog = detectProgramFromFilename(filename);
      const f = new File([arrayBuf], filename, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const yr = detectYearFromFilename(filename);
      onChange({ ...slot, _file: f, fileName: filename, ...(yr ? { year: yr } : {}), _type: detectedType, _program: prog });
      setSlotDriveSelected(file);
    } catch {
      // silent
    } finally {
      setDownloading(false);
    }
  };

  const filtered = (driveFiles ?? []).filter(f =>
    !driveSearch || f.name.toLowerCase().includes(driveSearch.toLowerCase())
  );

  return (
    <div style={{
      background: "rgba(255,255,255,.06)", borderRadius: 14, padding: 18,
      border: slot._file ? "1px solid #1abc9c" : "1px solid rgba(255,255,255,.1)"
    }}>
      <div style={{ marginBottom: 10 }}>
        <label className="label">{label}</label>
        <input className="input" value={slot.year}
          onChange={e => onChange({ ...slot, year: e.target.value })}
          placeholder="مثال: 2024-2025" />
      </div>

      {/* Source tabs */}
      <div style={{
        display: "flex", borderRadius: 8, overflow: "hidden",
        border: "1px solid rgba(255,255,255,.1)", background: "rgba(255,255,255,.04)", marginBottom: 10,
      }}>
        {[{ id: "local", label: "📁 جهازك" }, { id: "drive", label: "☁️ Drive" }].map(t => (
          <button key={t.id} onClick={() => setSlotTab(t.id)} style={{
            flex: 1, padding: "7px 4px", border: "none", cursor: "pointer",
            fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 12,
            background: slotTab === t.id ? "rgba(26,188,156,.2)" : "transparent",
            color: slotTab === t.id ? "#1abc9c" : "rgba(255,255,255,.45)",
            borderBottom: slotTab === t.id ? "2px solid #1abc9c" : "2px solid transparent",
            transition: "all .2s",
          }}>{t.label}</button>
        ))}
      </div>

      {/* Local upload */}
      {slotTab === "local" && (
        <>
          <div onClick={() => ref.current.click()} style={{
            border: "1.5px dashed rgba(255,255,255,.25)", borderRadius: 10, padding: "16px 10px",
            textAlign: "center", cursor: "pointer", background: "rgba(255,255,255,.03)",
            borderColor: slot._file ? "#1abc9c" : "rgba(255,255,255,.25)",
          }}>
            <div style={{ fontSize: 28, marginBottom: 6 }}>{slot._file ? "✅" : "📂"}</div>
            <div style={{ color: slot._file ? "#1abc9c" : "rgba(255,255,255,.6)", fontSize: 12, fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {slot.fileName || "اضغط لاختيار ملف"}
            </div>
            {slot._file && (
              <div style={{ marginTop: 6 }}>
                <div style={{ display: "flex", gap: 5, flexWrap: "wrap", justifyContent: "center", marginBottom: slot._file ? 5 : 0 }}>
                  {slot._type && SCHEMAS[slot._type] ? (
                    <span style={{ fontSize: 10, color: "#1abc9c", background: "rgba(26,188,156,.15)", borderRadius: 6, padding: "2px 7px", fontWeight: 700 }}>
                      {SCHEMAS[slot._type].icon} {SCHEMAS[slot._type].label}
                    </span>
                  ) : (
                    <span style={{ fontSize: 10, color: "#ffd54f", background: "rgba(255,193,7,.12)", borderRadius: 6, padding: "2px 7px", fontWeight: 700 }}>
                      ⚠ نوع غير محدد
                    </span>
                  )}
                  {slot._program && (
                    <span style={{ fontSize: 10, color: "#d6eaf8", background: "rgba(255,255,255,.08)", borderRadius: 6, padding: "2px 7px" }}>
                      {slot._program}
                    </span>
                  )}
                </div>
              </div>
            )}
          </div>
          <input ref={ref} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
            onChange={async e => {
              const f = e.target.files[0];
              if (!f) return;
              const yr = detectYearFromFilename(f.name);
              const prog = detectProgramFromFilename(f.name);
              let detectedType = detectTypeHintFromFilename(f.name);
              try {
                const buf = await f.arrayBuffer();
                const rows = readExcel(buf);
                detectedType = detectSurveyType(f.name, rows[0]) ?? detectedType;
              } catch { /* keep hint */ }
              onChange({ ...slot, _file: f, fileName: f.name, ...(yr ? { year: yr } : {}), _program: prog, _type: detectedType });
            }} />
          {slot._file && (
            <div style={{ marginTop: 8 }}>
              <div style={{ color: "rgba(255,255,255,.4)", fontSize: 10, marginBottom: 4, textAlign: "center" }}>
                النوع غير صحيح؟ اختر:
              </div>
              <div style={{ display: "flex", gap: 4, flexWrap: "wrap", justifyContent: "center" }}>
                {Object.values(SCHEMAS).map(sc => (
                  <button key={sc.id} onClick={e => { e.stopPropagation(); onChange({ ...slot, _type: sc.id }); }}
                    style={{
                      padding: "3px 9px", borderRadius: 8, cursor: "pointer",
                      fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 10,
                      border: `1.5px solid ${slot._type === sc.id ? "#1abc9c" : "rgba(255,255,255,.15)"}`,
                      background: slot._type === sc.id ? "rgba(26,188,156,.18)" : "rgba(255,255,255,.04)",
                      color: slot._type === sc.id ? "#1abc9c" : "rgba(255,255,255,.5)",
                      transition: "all .15s",
                    }}>
                    {sc.icon} {sc.label}
                  </button>
                ))}
              </div>
            </div>
          )}
        </>
      )}

      {/* Drive picker */}
      {slotTab === "drive" && (
        !driveToken ? (
          <div style={{ textAlign: "center", padding: "14px 0" }}>
            <button onClick={onConnectDrive} style={{
              display: "inline-flex", alignItems: "center", gap: 8,
              padding: "10px 20px", borderRadius: 50, border: "1px solid #dadce0",
              background: "#fff", color: "#3c4043", cursor: "pointer",
              fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 13,
            }}>
              <svg width="16" height="16" viewBox="0 0 48 48">
                <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
                <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
                <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
                <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.18 1.48-4.97 2.31-8.16 2.31-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
              </svg>
              ربط Google Drive
            </button>
          </div>
        ) : driveLoading ? (
          <div style={{ textAlign: "center", padding: "14px 0", color: "rgba(255,255,255,.5)", fontSize: 12 }}>
            جاري تحميل الملفات…
          </div>
        ) : (
          <>
            <input value={driveSearch} onChange={e => setDriveSearch(e.target.value)}
              placeholder="🔍 ابحث…"
              style={{
                width: "100%", background: "#0d1f33", color: "#e8f0fe",
                border: "1px solid rgba(255,255,255,.18)", borderRadius: 8,
                padding: "7px 10px", fontFamily: "'Cairo',sans-serif",
                fontSize: 12, direction: "rtl", marginBottom: 6, outline: "none",
              }} />
            <div style={{ maxHeight: 160, overflowY: "auto", display: "flex", flexDirection: "column", gap: 4 }}>
              {filtered.length === 0 ? (
                <div style={{ color: "rgba(255,255,255,.3)", fontSize: 12, textAlign: "center", padding: 12 }}>لا توجد ملفات</div>
              ) : filtered.map(f => {
                const isSel = slotDriveSelected?.id === f.id;
                return (
                  <div key={f.id} onClick={() => pickFromDrive(f)} style={{
                    display: "flex", alignItems: "center", gap: 8,
                    background: isSel ? "rgba(26,188,156,.18)" : "rgba(255,255,255,.04)",
                    border: `1px solid ${isSel ? "rgba(26,188,156,.5)" : "rgba(255,255,255,.08)"}`,
                    borderRadius: 8, padding: "7px 10px",
                    cursor: downloading ? "not-allowed" : "pointer", transition: "all .15s",
                  }}>
                    <span style={{ fontSize: 14 }}>{f.mimeType === GSHEETS_MIME ? "📊" : "📄"}</span>
                    <div style={{ flex: 1, overflow: "hidden" }}>
                      <div style={{ color: "#e8f0fe", fontSize: 11, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</div>
                    </div>
                    {isSel && <span style={{ color: "#1abc9c", fontSize: 14 }}>✅</span>}
                    {downloading && isSel && (
                      <svg className="spin" width="12" height="12" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="#1abc9c" strokeWidth="2" strokeDasharray="22 8"/></svg>
                    )}
                  </div>
                );
              })}
            </div>
            {slot._file && (
              <div style={{ marginTop: 8 }}>
                <div style={{ color: "rgba(255,255,255,.4)", fontSize: 10, marginBottom: 4, textAlign: "center" }}>
                  النوع غير صحيح؟ اختر:
                </div>
                <div style={{ display: "flex", gap: 4, flexWrap: "wrap", justifyContent: "center" }}>
                  {Object.values(SCHEMAS).map(sc => (
                    <button key={sc.id} onClick={() => onChange({ ...slot, _type: sc.id })}
                      style={{
                        padding: "3px 9px", borderRadius: 8, cursor: "pointer",
                        fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 10,
                        border: `1.5px solid ${slot._type === sc.id ? "#1abc9c" : "rgba(255,255,255,.15)"}`,
                        background: slot._type === sc.id ? "rgba(26,188,156,.18)" : "rgba(255,255,255,.04)",
                        color: slot._type === sc.id ? "#1abc9c" : "rgba(255,255,255,.5)",
                        transition: "all .15s",
                      }}>
                      {sc.icon} {sc.label}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </>
        )
      )}
    </div>
  );
}

// ── MetadataForm ──────────────────────────────────────────────────────────────
function MetadataForm({ meta, onChange }) {
  const F = (key, label, placeholder) => (
    <div>
      <label className="label">{label}</label>
      <input className="input" value={meta[key] ?? ""} placeholder={placeholder}
        onChange={e => onChange({ ...meta, [key]: e.target.value })} />
    </div>
  );
  return (
    <div className="card" style={{ padding: 32 }}>
      <div style={{ color: "#fff", fontSize: 20, fontWeight: 900, marginBottom: 20, textAlign: "center" }}>بيانات التقرير</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <div>
          <label className="label">اسم البرنامج / القسم</label>
          <ProgramPicker value={meta.program ?? ""}
            onChange={v => onChange({ ...meta, program: v })} />
        </div>
        {F("year",       "العام الأكاديمي", "مثال: 2024-2025")}
        {F("preparedBy", "أعده",            "اسم معد التقرير")}
        {F("reviewer",   "راجعه",           "اسم المراجع")}
      </div>
    </div>
  );
}

// ── ProgramPicker ─────────────────────────────────────────────────────────────
function ProgramPicker({ value, onChange }) {
  const isCustom = !!value && !PROGRAMS.includes(value);
  const [mode, setMode] = useState(isCustom ? "other" : "list");

  useEffect(() => {
    // Re-sync if parent value changes (e.g. auto-detect from filename)
    if (value && PROGRAMS.includes(value) && mode !== "list") setMode("list");
    if (value && !PROGRAMS.includes(value) && mode !== "other") setMode("other");
  }, [value]); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <div>
      <select className="input"
        value={mode === "other" ? "__other__" : value}
        onChange={e => {
          if (e.target.value === "__other__") { setMode("other"); onChange(""); }
          else { setMode("list"); onChange(e.target.value); }
        }}>
        <option value="">— اختر البرنامج —</option>
        {PROGRAMS.map(p => <option key={p} value={p}>{p}</option>)}
        <option value="__other__">أخرى…</option>
      </select>
      {mode === "other" && (
        <input className="input" type="text" value={value}
          placeholder="اكتب اسم البرنامج"
          onChange={e => onChange(e.target.value)}
          style={{ marginTop: 8 }} />
      )}
    </div>
  );
}

// ── FilterDropdown ────────────────────────────────────────────────────────────
function FilterDropdown({ label, value, options, onChange }) {
  if (!options.length) return null;
  return (
    <label style={{ color: "rgba(255,255,255,.8)", fontSize: 13, display: "flex",
                    alignItems: "center", gap: 8 }}>
      {label}:
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ background: "#0f2035", color: "#e8f0fe",
                 border: "1px solid rgba(255,255,255,.25)", borderRadius: 8,
                 padding: "7px 32px 7px 12px", fontFamily: "inherit", fontSize: 13,
                 cursor: "pointer", outline: "none",
                 appearance: "none", WebkitAppearance: "none",
                 backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='7' viewBox='0 0 10 7'%3E%3Cpath d='M1 1l4 4 4-4' stroke='%231abc9c' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E\")",
                 backgroundRepeat: "no-repeat", backgroundPosition: "left 10px center" }}>
        <option value="" style={{ background: "#0f2035", color: "#e8f0fe" }}>— الكل —</option>
        {options.map(o => <option key={o} value={o} style={{ background: "#0f2035", color: "#e8f0fe" }}>{o}</option>)}
      </select>
    </label>
  );
}

// ── DataPreviewTable ──────────────────────────────────────────────────────────
function DataPreviewTable({ headers, rows, removed, onToggleRow,
                            deptIdx, degreeIdx, filters, onFilterChange }) {
  const deptOptions  = useMemo(() => uniqueCol(rows, deptIdx),   [rows, deptIdx]);
  const degreeOpts   = useMemo(() => uniqueCol(rows, degreeIdx), [rows, degreeIdx]);

  const visibleRows = rows.map((r, i) => ({ row: r, idx: i }))
    .filter(({ idx })  => !removed.has(idx))
    .filter(({ row }) => !filters.dept   || String(row[deptIdx])   === filters.dept)
    .filter(({ row }) => !filters.degree || String(row[degreeIdx]) === filters.degree);

  const cellTh = { padding: "8px 10px", fontSize: 11, fontWeight: 700, textAlign: "right",
                   borderBottom: "1px solid rgba(255,255,255,.15)", whiteSpace: "nowrap" };
  const cellTd = { padding: "6px 10px", textAlign: "right", color: "rgba(255,255,255,.85)",
                   verticalAlign: "top" };

  return (
    <div>
      <div style={{ display: "flex", gap: 12, marginBottom: 12, flexWrap: "wrap",
                    alignItems: "center" }}>
        <FilterDropdown label="القسم" value={filters.dept}
          options={deptOptions}
          onChange={v => onFilterChange({ ...filters, dept: v })} />
        <FilterDropdown label="الدرجة / المسمى الوظيفي" value={filters.degree}
          options={degreeOpts}
          onChange={v => onFilterChange({ ...filters, degree: v })} />
        <button className="btn btn-ghost btn-sm"
          onClick={() => onFilterChange({ dept: "", degree: "" })}>
          ↩ إزالة الفلاتر
        </button>
      </div>

      <div style={{ maxHeight: 500, overflow: "auto",
                    border: "1px solid rgba(255,255,255,.15)", borderRadius: 8,
                    background: "rgba(0,0,0,.2)" }}>
        <table style={{ width: "100%", fontSize: 12, color: "#fff",
                        borderCollapse: "collapse", direction: "rtl" }}>
          <thead style={{ position: "sticky", top: 0, background: "#1a3a5c", zIndex: 1 }}>
            <tr>
              <th style={{ ...cellTh, textAlign: "center", width: 60 }}>إجراء</th>
              {headers.map((h, i) => <th key={i} style={cellTh}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {visibleRows.map(({ row, idx }) => (
              <tr key={idx} style={{ borderTop: "1px solid rgba(255,255,255,.05)" }}>
                <td style={{ ...cellTd, textAlign: "center" }}>
                  <button className="btn btn-danger btn-sm" title="حذف هذا الصف"
                    style={{ padding: "4px 10px", fontSize: 14 }}
                    onClick={() => onToggleRow(idx)}>🗑️</button>
                </td>
                {headers.map((_, ci) => (
                  <td key={ci} style={cellTd}>
                    {row[ci] == null ? "" : String(row[ci])}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={{ color: "rgba(255,255,255,.5)", fontSize: 12, marginTop: 8,
                    textAlign: "center" }}>
        عدد الصفوف المعروضة: {visibleRows.length} من إجمالي {rows.length}
        {" "}(تم حذف {removed.size})
      </div>
    </div>
  );
}

// ── ResultsPreview ────────────────────────────────────────────────────────────
function ResultsPreview({ result }) {
  const is5 = result.scaleType === "likert-5";
  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14, marginBottom: 24 }}>
        {[
          { n: result.n,                    l: "عدد المستجيبين",   icon: "👥" },
          { n: result.axes.length,          l: "عدد المحاور",      icon: "📋" },
          { n: result.totalQuestions,       l: "عدد الأسئلة",      icon: "❓" },
          { n: `${result.overallAgreePct}%`, l: "نسبة الموافقة",   icon: "✅" },
        ].map((s, i) => (
          <div key={i} className="stat-card">
            <div style={{ fontSize: 26, marginBottom: 6 }}>{s.icon}</div>
            <div style={{ color: "#1abc9c", fontSize: 28, fontWeight: 900 }}>{s.n}</div>
            <div style={{ color: "rgba(255,255,255,.55)", fontSize: 11, marginTop: 4 }}>{s.l}</div>
          </div>
        ))}
      </div>

      {Object.keys(result.byDegree).length > 0 && (
        <div className="card" style={{ padding: 22, marginBottom: 18 }}>
          <div style={{ color: "#fff", fontSize: 16, fontWeight: 700, marginBottom: 12, borderBottom: "1px solid rgba(255,255,255,.1)", paddingBottom: 8 }}>
            توزيع المشاركين
          </div>
          <table className="mini-table">
            <thead><tr><th>العدد</th><th>النسبة</th><th style={{ textAlign: "right" }}>الدرجة / الوظيفة</th></tr></thead>
            <tbody>
              {Object.entries(result.byDegree).sort((a, b) => b[1] - a[1]).map(([k, v], i) => (
                <tr key={i}>
                  <td>{v}</td>
                  <td style={{ color: "#1abc9c", fontWeight: 700 }}>{(v / result.n * 100).toFixed(1)}%</td>
                  <td className="rtl-td" style={{ color: "#e8f0fe" }}>{k}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div className="card" style={{ padding: 22 }}>
        <div style={{ color: "#fff", fontSize: 16, fontWeight: 700, marginBottom: 12, borderBottom: "1px solid rgba(255,255,255,.1)", paddingBottom: 8 }}>
          ملخص المحاور
        </div>
        <table className="mini-table">
          <thead>
            <tr>
              <th style={{ width: 36 }}>م</th>
              <th style={{ textAlign: "right" }}>المحور</th>
              {is5 && <th>المتوسط</th>}
              <th>نسبة الموافقة</th>
              <th>الاتجاه</th>
            </tr>
          </thead>
          <tbody>
            {result.axes.map((ax, i) => (
              <tr key={i}>
                <td style={{ color: "#1abc9c", fontWeight: 700 }}>{i + 1}</td>
                <td className="rtl-td" style={{ fontSize: 12 }}>{ax.name}</td>
                {is5 && <td style={{ fontWeight: 700, color: "#d6eaf8" }}>{ax.axisMean}</td>}
                <td style={{ fontWeight: 700, color: "#1abc9c" }}>{ax.axisAgreePct}%</td>
                <td>
                  <span className="badge" style={{ background: `${dirColor(ax.tier)}22`, color: dirColor(ax.tier) }}>
                    {ax.direction}
                  </span>
                </td>
              </tr>
            ))}
            <tr>
              <td colSpan={is5 ? 2 : 2} style={{ textAlign: "right", fontWeight: 900, color: "#fff", background: "rgba(26,188,156,.18)", padding: "10px" }}>
                الإجمالي
              </td>
              {is5 && <td style={{ fontWeight: 900, color: "#1abc9c", background: "rgba(26,188,156,.18)" }}>{result.overallMean}</td>}
              <td style={{ fontWeight: 900, color: "#1abc9c", fontSize: 15, background: "rgba(26,188,156,.18)" }}>
                {result.overallAgreePct}%
              </td>
              <td style={{ background: "rgba(26,188,156,.18)" }}>
                <span className="badge" style={{ background: "rgba(13,110,58,.35)", color: "#1abc9c" }}>
                  {result.overallDirection}
                </span>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── ComparisonPreview ─────────────────────────────────────────────────────────
function ComparisonPreview({ comparison }) {
  const { axes, overall } = comparison;
  const years = overall.map(o => o.year);
  return (
    <div className="card" style={{ padding: 24 }}>
      <div style={{ color: "#fff", fontSize: 16, fontWeight: 700, marginBottom: 16 }}>جدول المقارنة</div>
      <div style={{ overflowX: "auto" }}>
        <table className="mini-table">
          <thead>
            <tr>
              <th style={{ width: 40 }}>م</th>
              <th style={{ textAlign: "right" }}>المحور</th>
              {years.map((y, i) => <th key={i}>{y}</th>)}
              <th>الاتجاه</th>
            </tr>
          </thead>
          <tbody>
            {axes.map((ax, ai) => (
              <tr key={ai}>
                <td style={{ color: "#1abc9c", fontWeight: 700 }}>{ai + 1}</td>
                <td className="rtl-td" style={{ fontSize: 12 }}>{ax.name}</td>
                {ax.years.map((y, yi) => (
                  <td key={yi} style={{ fontWeight: 700, color: "#d6eaf8" }}>
                    {y.agreePct !== null ? `${y.agreePct}%` : "—"}
                  </td>
                ))}
                <td>
                  <span className="badge" style={{
                    background: ax.trend === "تحسن" ? "rgba(13,110,58,.3)" : ax.trend === "تراجع" ? "rgba(146,43,33,.3)" : "rgba(120,66,18,.2)",
                    color: ax.trend === "تحسن" ? "#1abc9c" : ax.trend === "تراجع" ? "#e74c3c" : "#f0b27a"
                  }}>{ax.trend}</span>
                </td>
              </tr>
            ))}
            <tr>
              <td colSpan={2} style={{ textAlign: "right", fontWeight: 900, color: "#fff", background: "rgba(26,188,156,.15)", padding: "9px 10px" }}>الإجمالي</td>
              {overall.map((o, i) => (
                <td key={i} style={{ fontWeight: 900, color: "#1abc9c", fontSize: 14, background: "rgba(26,188,156,.15)" }}>
                  {o.agreePct}%
                </td>
              ))}
              <td style={{ background: "rgba(26,188,156,.15)" }}></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

const LOCAL_STEPS = ["قراءة البيانات", "تنظيف البيانات", "كشف نوع الاستبيان", "حساب المؤشرات", "تجهيز العرض"];
const DRIVE_STEPS = ["تحميل الملف من Drive", "قراءة البيانات", "تنظيف البيانات", "كشف نوع الاستبيان", "تجهيز العرض"];

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [step, setStep]         = useState(0);
  const [surveyType, setSurveyType] = useState("faculty");
  const [mode, setMode]         = useState("annual");
  const [meta, setMeta]         = useState({ year: "2024-2025", program: "", preparedBy: "", reviewer: "" });
  const [processing, setProcessing] = useState(false);
  const [error, setError]       = useState("");
  const [showSettings, setShowSettings] = useState(false);
  const [showTutorial, setShowTutorial] = useState(false);
  const [settings, setSettings]         = useState(loadSettings);
  const [aiSettings, setAiSettings]     = useState(loadAiSettings);

  const [singleFile, setSingleFile]     = useState(null);
  const [singleResult, setSingleResult] = useState(null);
  const [showEnhancedView, setShowEnhancedView] = useState(false);
  const singleRef = useRef();
  const [dragging, setDragging]         = useState(false);

  // File-detection state (step 0 → 1)
  const [detectedAutoType, setDetectedAutoType] = useState(null);
  const [rawRows, setRawRows]                   = useState(null);

  // Processing animation overlay
  const [procSteps, setProcSteps] = useState(null);
  const [procFile,  setProcFile]  = useState("");

  const runStepAnim = useCallback((labels, startAt = 0, msPerStep = 480, onDone) => {
    const build = (cur) => labels.map((label, i) => ({
      label,
      status: i < cur ? "done" : i === cur ? "active" : "pending",
    }));
    setProcSteps(build(startAt));
    let cur = startAt;
    const tick = () => {
      cur++;
      if (cur >= labels.length) {
        setProcSteps(labels.map(label => ({ label, status: "done" })));
        setTimeout(() => { setProcSteps(null); onDone(); }, 380);
        return;
      }
      setProcSteps(build(cur));
      setTimeout(tick, msPerStep);
    };
    setTimeout(tick, msPerStep);
  }, []);

  // Batch mode (multiple files selected in step 0)
  const [batchMode,  setBatchMode]  = useState(false);
  const [batchFiles, setBatchFiles] = useState([]);
  const [batchYear,  setBatchYear]  = useState("2024-2025");

  // Upload source tab
  const [uploadTab,         setUploadTab]         = useState("upload");
  const [driveToken,        setDriveToken]        = useState(null);
  const [driveFiles,        setDriveFiles]        = useState([]);
  const [driveSearch,       setDriveSearch]       = useState("");
  const [driveLoading,      setDriveLoading]      = useState(false);
  const [driveSelected,     setDriveSelected]     = useState(null);
  const [driveProcessing,   setDriveProcessing]   = useState(false);
  const [driveFilterType,   setDriveFilterType]   = useState("");
  const [driveFilterYear,   setDriveFilterYear]   = useState("");
  const [driveFilterProgram,setDriveFilterProgram]= useState("");
  const [driveDashboard,    setDriveDashboard]    = useState(false);
  const tokenClientRef = useRef(null);

  // Annual-mode preview state (Step 3)
  const [singleHeaders, setSingleHeaders]     = useState(null);
  const [singleAllRows, setSingleAllRows]     = useState(null);
  const [singleSchema, setSingleSchema]       = useState(null);
  const [singleMetaCols, setSingleMetaCols]   = useState(null);
  const [singleRemoved, setSingleRemoved]     = useState(new Set());
  const [singleFilters, setSingleFilters]     = useState({ dept: "", degree: "" });

  const slotCount = mode === "compare2" ? 2 : 3;
  const defaultSlots = () => [
    { year: "2022-2023", fileName: "", result: null, _file: null, _type: null, _program: null },
    { year: "2023-2024", fileName: "", result: null, _file: null, _type: null, _program: null },
    { year: "2024-2025", fileName: "", result: null, _file: null, _type: null, _program: null },
  ];
  const [slots, setSlots]       = useState(defaultSlots());
  const [comparison, setComparison] = useState(null);

  const schema = SCHEMAS[surveyType];

  const processBuffer = useCallback((buf, filename) => {
    const rows = readExcel(buf);
    const type = detectSurveyType(filename, rows[0]) ?? surveyType;
    const s = SCHEMAS[type] ?? schema;
    return analyze(rows, s);
  }, [surveyType, schema]);

  const fetchDriveFiles = useCallback(async (token) => {
    setDriveLoading(true);
    try {
      const q = DRIVE_FILE_MIMES.map(m => `mimeType='${m}'`).join(" or ");
      const params = new URLSearchParams({
        q: `(${q}) and trashed=false`,
        fields: "files(id,name,mimeType,modifiedTime)",
        orderBy: "modifiedTime desc",
        pageSize: "50",
      });
      const res = await fetch(`https://www.googleapis.com/drive/v3/files?${params}`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error.message);
      setDriveFiles(data.files || []);
    } catch (e) {
      setError("خطأ في تحميل ملفات Drive: " + e.message);
    } finally {
      setDriveLoading(false);
    }
  }, []);

  const connectDrive = useCallback(async () => {
    if (!GOOGLE_CLIENT_ID) return;
    await loadGisScript();
    if (!tokenClientRef.current) {
      tokenClientRef.current = window.google.accounts.oauth2.initTokenClient({
        client_id: GOOGLE_CLIENT_ID,
        scope: DRIVE_SCOPE,
        prompt: "select_account",
        callback: async (resp) => {
          if (resp.error) { setError("فشل الاتصال بـ Google Drive: " + resp.error); return; }
          setDriveToken(resp.access_token);
          await fetchDriveFiles(resp.access_token);
        },
      });
    }
    tokenClientRef.current.requestAccessToken({ prompt: "select_account" });
  }, [fetchDriveFiles]);

  const handleDriveFileSelect = useCallback(async (file) => {
    if (!driveToken || !file) return;
    setDriveProcessing(true);
    setError("");
    setProcFile(file.name);
    // Show step 0 (download) as active immediately
    setProcSteps(DRIVE_STEPS.map((label, i) => ({ label, status: i === 0 ? "active" : "pending" })));
    try {
      const isGSheet = file.mimeType === GSHEETS_MIME;
      const url = isGSheet
        ? `https://www.googleapis.com/drive/v3/files/${file.id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
        : `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`;
      const res = await fetch(url, { headers: { Authorization: `Bearer ${driveToken}` } });
      if (!res.ok) throw new Error(`خطأ في التحميل (${res.status})`);
      const buf = await res.arrayBuffer();
      const filename = isGSheet
        ? (file.name.endsWith(".xlsx") ? file.name : file.name + ".xlsx")
        : file.name;

      setSingleFile({ name: filename });
      const rows = readExcel(buf);
      setRawRows(rows);
      const detected = detectSurveyType(filename, rows[0]) ?? null;
      setDetectedAutoType(detected);
      setSurveyType(detected ?? "faculty");
      const prog = detectProgramFromFilename(filename);
      if (prog) setMeta(m => ({ ...m, program: prog }));

      // Animate remaining steps (download already done = step 0 done)
      runStepAnim(DRIVE_STEPS, 1, 400, () => {
        setDriveProcessing(false);
        setStep(2);
      });
    } catch (err) {
      setProcSteps(null);
      setDriveProcessing(false);
      setError(err.message ?? "خطأ في تحميل الملف من Drive");
    }
  }, [driveToken, runStepAnim]);

  // Step 0: read file headers → auto-detect → advance to step 1
  const handleFileSelected = useCallback((file) => {
    setSingleFile(file);
    setError("");
    setProcFile(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const rows = readExcel(e.target.result);
        setRawRows(rows);
        const detected = detectSurveyType(file.name, rows[0]) ?? null;
        setDetectedAutoType(detected);
        setSurveyType(detected ?? "faculty");
        const prog = detectProgramFromFilename(file.name);
        if (prog) setMeta(m => ({ ...m, program: prog }));
        // Animate all 5 steps then advance
        runStepAnim(LOCAL_STEPS, 0, 520, () => setStep(2));
      } catch (err) {
        setProcSteps(null);
        setError(err.message ?? "تعذّر قراءة الملف. تأكد أنه ملف Excel صحيح.");
      }
    };
    reader.readAsArrayBuffer(file);
  }, [runStepAnim]);

  // Step 1: user confirms the detected type → prepare data → step 2
  const handleValidationConfirm = useCallback(() => {
    if (!rawRows || !surveyType) return;
    setProcessing(true); setError("");
    try {
      const s = SCHEMAS[surveyType];
      const { headers, dataRows, metaCols } = prepareData(rawRows, s);
      setSingleHeaders(headers);
      setSingleAllRows(dataRows);
      setSingleSchema(s);
      setSingleMetaCols(metaCols);
      setSingleRemoved(new Set());
      setSingleFilters({ dept: "", degree: "" });
      // Auto-fill preparedBy from settings responsible person (if empty)
      const responsible = settings.surveyResponsible?.[surveyType];
      if (responsible) setMeta(m => ({ ...m, preparedBy: m.preparedBy || responsible }));
      setStep(3);
    } catch (err) {
      setError(err.message ?? "خطأ في تحضير البيانات.");
    } finally {
      setProcessing(false);
    }
  }, [rawRows, surveyType, settings]);

  // Step 3 (comparison): process all files (slot 0 = singleFile from step 0)
  const handleProcessSingle = useCallback(() => {
    if (!singleFile) return;
    setProcessing(true); setError("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const rows = readExcel(e.target.result);
        const type = detectSurveyType(singleFile.name, rows[0]) ?? surveyType;
        const s = SCHEMAS[type] ?? schema;
        const { headers, dataRows, metaCols } = prepareData(rows, s);
        setSingleHeaders(headers);
        setSingleAllRows(dataRows);
        setSingleSchema(s);
        setSingleMetaCols(metaCols);
        setSingleRemoved(new Set());
        setSingleFilters({ dept: "", degree: "" });
        // Auto-detect program from filename — only if user hasn't already filled it
        const detected = detectProgramFromFilename(singleFile.name);
        if (detected) setMeta(m => ({ ...m, program: detected }));
        setStep(3);
      } catch (err) {
        setError(err.message ?? "حدث خطأ في تحليل الملف.");
      } finally { setProcessing(false); }
    };
    reader.readAsArrayBuffer(singleFile);
  }, [singleFile, surveyType, schema]);

  const handleProcessMulti = useCallback(() => {
    const activeSlots = slots.slice(0, slotCount);
    if (!activeSlots.every(s => s._file)) { setError("ارفع ملفاً لكل سنة."); return; }

    // Validate unique years
    const years = activeSlots.map(s => (s.year ?? "").trim());
    if (new Set(years).size !== activeSlots.length) {
      setError("كل ملف يجب أن يكون لسنة دراسية مختلفة — تحقق من السنوات المدخلة.");
      return;
    }

    // Validate same survey type (when all slots have a detected type)
    const slotTypes = activeSlots.map(s => s._type).filter(Boolean);
    if (slotTypes.length === activeSlots.length) {
      const uniqueTypes = new Set(slotTypes);
      if (uniqueTypes.size > 1) {
        setError(`الملفات لأنواع استبيانات مختلفة: ${[...uniqueTypes].map(t => SCHEMAS[t]?.label ?? t).join(" / ")} — تأكد أن جميع الملفات لنفس نوع الاستبيان.`);
        return;
      }
    }

    setProcessing(true); setError("");
    let pending = activeSlots.length;
    const results = Array(activeSlots.length).fill(null);
    let firstDetectedType = null;
    const firstDetectedProgram = activeSlots.find(s => s._program)?._program ?? null;

    activeSlots.forEach((slot, idx) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const rows = readExcel(e.target.result);
          const type = detectSurveyType(slot._file.name, rows[0]) ?? surveyType;
          if (!firstDetectedType) firstDetectedType = type;
          const s = SCHEMAS[type] ?? schema;
          const result = analyze(rows, s);
          results[idx] = { year: slot.year, result };
        } catch (err) {
          setError(err.message ?? "خطأ في ملف " + slot._file.name);
        } finally {
          if (--pending === 0) {
            setProcessing(false);
            if (firstDetectedType) setSurveyType(firstDetectedType);
            if (firstDetectedProgram) setMeta(m => ({ ...m, program: m.program || firstDetectedProgram }));
            const validResults = results.filter(Boolean);
            if (validResults.length > 0) {
              const cmp = buildComparison(validResults);
              setComparison({ ...cmp, slots: results });
              setStep(4);
            }
          }
        }
      };
      reader.readAsArrayBuffer(slot._file);
    });
  }, [slots, slotCount, surveyType, schema]);

  const handleDropSingle = (e) => {
    e.preventDefault(); setDragging(false);
    const files = [...e.dataTransfer.files].filter(f => /\.(xlsx|xls|csv)$/i.test(f.name));
    if (!files.length) return;
    if (files.length === 1) { handleFileSelected(files[0]); }
    else                    { setBatchFiles(files); setBatchMode(true); }
  };

  const downloadAnnual = async () => {
    setProcessing(true);
    try {
      const blob = await buildAnnualDocx(singleResult, { ...meta }, settings);
      const dept = meta.program ? `_${meta.program}` : "";
      downloadBlob(blob, `تقرير_${schema.label}${dept}_${meta.year}.docx`);
    } finally { setProcessing(false); }
  };

  const downloadComparison = async () => {
    setProcessing(true);
    try {
      const blob = await buildComparisonDocx(comparison, { ...meta }, settings);
      const dept = meta.program ? `_${meta.program}` : "";
      const years = comparison.slots?.map(s => s.year).filter(Boolean).join("_و_") ?? meta.year;
      downloadBlob(blob, `مقارنة_${schema.label}${dept}_${years}.docx`);
    } finally { setProcessing(false); }
  };

  const reset = () => {
    setStep(0); setSingleFile(null); setSingleResult(null); setComparison(null);
    setSlots(defaultSlots()); setError(""); setProcessing(false); setShowSettings(false);
    setSingleHeaders(null); setSingleAllRows(null); setSingleSchema(null);
    setSingleMetaCols(null); setSingleRemoved(new Set());
    setSingleFilters({ dept: "", degree: "" });
    setDetecting(false); setDetectedAutoType(null); setRawRows(null);
    setBatchMode(false); setBatchFiles([]);
    setUploadTab("upload"); setDriveSelected(null); setDriveDashboard(false);
    setDriveFilterType(""); setDriveFilterYear(""); setDriveFilterProgram("");
  };

  const isAnnual = mode === "annual";
  const canProceed = isAnnual ? !!singleFile : slots.slice(0, slotCount).every(s => s._file);

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg,#0f2035 0%,#1a3a5c 50%,#0d3b2e 100%)",
      fontFamily: "'Cairo',sans-serif", direction: "rtl",
    }}>
      <style>{CSS}</style>

      {/* Header */}
      <div style={{ padding: "18px 32px", borderBottom: "1px solid rgba(255,255,255,.1)", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{ width: 42, height: 42, borderRadius: 11, background: "linear-gradient(135deg,#1abc9c,#2874a6)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>📊</div>
          <div>
            <div style={{ color: "#fff", fontSize: 18, fontWeight: 900 }}>محلل الاستبيانات الأكاديمية</div>
            <div style={{ color: "rgba(255,255,255,.4)", fontSize: 11 }}>ERU – وحدة ضمان الجودة</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          {/* Help button */}
          <button
            onClick={() => setShowTutorial(true)}
            title="دليل الاستخدام"
            style={{
              width: 36, height: 36, borderRadius: "50%",
              border: "1px solid rgba(255,255,255,.18)",
              background: "rgba(255,255,255,.08)",
              color: "rgba(255,255,255,.7)", fontWeight: 900, fontSize: 16,
              cursor: "pointer", display: "flex", alignItems: "center",
              justifyContent: "center", transition: "all .18s",
              fontFamily: "'Cairo',sans-serif",
            }}
            onMouseOver={e => { e.currentTarget.style.background = "rgba(26,188,156,.25)"; e.currentTarget.style.borderColor = "rgba(26,188,156,.5)"; e.currentTarget.style.color = "#1abc9c"; }}
            onMouseOut={e  => { e.currentTarget.style.background = "rgba(255,255,255,.08)"; e.currentTarget.style.borderColor = "rgba(255,255,255,.18)"; e.currentTarget.style.color = "rgba(255,255,255,.7)"; }}
          >?</button>
          <button
            className={`btn btn-sm ${showSettings ? "btn-primary" : "btn-ghost"}`}
            onClick={() => setShowSettings(v => !v)}
          >
            ⚙ الإعدادات
          </button>
          {(step > 0 || showSettings) && (
            <button className="btn btn-ghost btn-sm" onClick={reset}>↩ بدء جديد</button>
          )}
        </div>
      </div>

      <div style={{ maxWidth: 1400, margin: "0 auto", padding: "28px 32px" }}>

        {/* ── Settings view ── */}
        {showSettings ? (
          <div className="card" style={{ padding: 32 }}>
            <SettingsPanel
              settings={settings} onChange={setSettings}
              aiSettings={aiSettings} onAiChange={setAiSettings}
            />
            <div style={{ textAlign: "center", marginTop: 24 }}>
              <button className="btn btn-primary" onClick={() => setShowSettings(false)}>
                ✓ حفظ والعودة
              </button>
            </div>
          </div>
        ) : batchMode ? (
          <BatchProcessor
            files={batchFiles}
            year={batchYear}
            onYearChange={setBatchYear}
            settings={settings}
            onBack={() => { setBatchMode(false); setBatchFiles([]); setError(""); }}
          />
        ) : (
          <>
            <StepBar step={step} />

            {/* ══ STEP 0 — Choose report mode ══ */}
            {step === 0 && (
              <div>
                <ReportModePicker value={mode} onChange={setMode} />
                <div style={{ display: "flex", justifyContent: "center", marginTop: 24 }}>
                  <button className="btn btn-primary" style={{ fontSize: 16, padding: "13px 36px" }}
                    onClick={() => { setError(""); setStep(1); }}>
                    التالي ←
                  </button>
                </div>
              </div>
            )}

            {/* ══ STEP 1 — Upload file(s) ══ */}
            {step === 1 && isAnnual && (
              <div className="card" style={{ padding: uploadTab === "drive" && driveToken ? "24px 32px" : 44, textAlign: "center" }}>
                {!(uploadTab === "drive" && driveToken) && (<>
                  <div style={{ fontSize: 52, marginBottom: 16 }}>📂</div>
                  <div style={{ color: "#fff", fontSize: 22, fontWeight: 900, marginBottom: 8 }}>
                    ارفع ملف الاستبيان
                  </div>
                  <div style={{ color: "rgba(255,255,255,.45)", fontSize: 13, marginBottom: 24 }}>
                    يُكتشف نوع الاستبيان تلقائياً من الأعمدة · يدعم .xlsx / .xls / .csv
                  </div>
                </>)}

                {/* Source tabs */}
                <div style={{
                  display: "flex", maxWidth: uploadTab === "drive" && driveToken ? "100%" : 520,
                  margin: uploadTab === "drive" && driveToken ? "0 0 20px" : "0 auto 28px",
                  borderRadius: 12, overflow: "hidden",
                  border: "1px solid rgba(255,255,255,.1)", background: "rgba(255,255,255,.04)",
                }}>
                  {[
                    { id: "upload", label: "📁 رفع ملف" },
                    { id: "drive",  label: "☁️ Google Drive" },
                  ].map(t => (
                    <button key={t.id} onClick={() => setUploadTab(t.id)} style={{
                      flex: 1, padding: "13px 8px", border: "none", cursor: "pointer",
                      fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 14,
                      background: uploadTab === t.id ? "rgba(26,188,156,.2)" : "transparent",
                      color:      uploadTab === t.id ? "#1abc9c" : "rgba(255,255,255,.45)",
                      borderBottom: uploadTab === t.id ? "2px solid #1abc9c" : "2px solid transparent",
                      transition: "all .2s",
                    }}>{t.label}</button>
                  ))}
                </div>

                {/* ── Upload tab ── */}
                {uploadTab === "upload" && (
                  <div
                    className={`upload-zone ${dragging ? "drag" : ""}`}
                    style={{ maxWidth: 640, margin: "0 auto" }}
                    onDragOver={e => { e.preventDefault(); setDragging(true); }}
                    onDragLeave={() => setDragging(false)}
                    onDrop={handleDropSingle}
                    onClick={() => singleRef.current.click()}
                  >
                    <div style={{ fontSize: 48, marginBottom: 12 }}>📊</div>
                    <div style={{ color: "#fff", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>
                      اسحب الملف هنا أو اضغط للاختيار
                    </div>
                    <div style={{ color: "rgba(255,255,255,.35)", fontSize: 12 }}>
                      ملف استجابات Google Forms · يمكن اختيار أكثر من ملف للمعالجة الدفعية
                    </div>
                    <input ref={singleRef} type="file" accept=".xlsx,.xls,.csv" multiple
                      style={{ display: "none" }}
                      onChange={e => {
                        const files = [...e.target.files];
                        if (!files.length) return;
                        if (files.length === 1) { handleFileSelected(files[0]); }
                        else                    { setBatchFiles(files); setBatchMode(true); }
                      }} />
                  </div>
                )}

                {/* ── Google Drive tab ── */}
                {uploadTab === "drive" && (
                  <div style={{ maxWidth: driveToken ? "100%" : 520, margin: driveToken ? "0" : "0 auto", textAlign: "right" }}>
                    {!GOOGLE_CLIENT_ID ? (
                      /* No Client ID configured */
                      <div style={{
                        background: "rgba(243,156,18,.07)", border: "1px solid rgba(243,156,18,.3)",
                        borderRadius: 14, padding: "22px 24px", fontSize: 13, color: "#f1c40f", lineHeight: 1.9,
                      }}>
                        ⚠️ لم يتم إعداد Google Client ID بعد. اتبع الخطوات:<br/>
                        <span style={{ color: "rgba(255,255,255,.6)", fontSize: 12 }}>
                          ١. إنشاء مشروع في Google Cloud Console<br/>
                          ٢. تفعيل Google Drive API<br/>
                          ٣. إنشاء OAuth 2.0 Client ID (نوع: Web application)<br/>
                          ٤. إضافة <code style={{ color: "#1abc9c" }}>http://localhost:3000</code> في Authorized JavaScript Origins<br/>
                          ٥. نسخ الـ Client ID وإضافته في ملف <code style={{ color: "#1abc9c" }}>.env.local</code>:<br/>
                          <code style={{ color: "#1abc9c", fontSize: 11 }}>VITE_GOOGLE_CLIENT_ID=your_client_id_here</code><br/>
                          ٦. إعادة تشغيل الـ dev server
                        </span>
                      </div>
                    ) : !driveToken ? (
                      /* Not connected */
                      <div style={{ textAlign: "center", padding: "28px 0" }}>
                        <div style={{ fontSize: 44, marginBottom: 14 }}>☁️</div>
                        <div style={{ color: "rgba(255,255,255,.5)", fontSize: 13, marginBottom: 22 }}>
                          اربط حسابك على Google Drive لاختيار ملف الاستبيان مباشرة
                        </div>
                        <button onClick={connectDrive} style={{
                          display: "inline-flex", alignItems: "center", gap: 10,
                          padding: "13px 28px", borderRadius: 50, border: "1px solid #dadce0",
                          background: "#fff", color: "#3c4043", cursor: "pointer",
                          fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 15,
                          boxShadow: "0 2px 8px rgba(0,0,0,.2)", transition: "all .2s",
                        }}
                          onMouseOver={e => e.currentTarget.style.boxShadow = "0 4px 16px rgba(0,0,0,.3)"}
                          onMouseOut={e  => e.currentTarget.style.boxShadow = "0 2px 8px rgba(0,0,0,.2)"}
                        >
                          <svg width="20" height="20" viewBox="0 0 48 48">
                            <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
                            <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
                            <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
                            <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.18 1.48-4.97 2.31-8.16 2.31-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
                          </svg>
                          ربط Google Drive
                        </button>
                      </div>
                    ) : driveLoading ? (
                      /* Loading files */
                      <div style={{ textAlign: "center", padding: "36px 0", color: "rgba(255,255,255,.5)" }}>
                        <svg className="spin" width="38" height="38" viewBox="0 0 38 38" style={{ display: "block", margin: "0 auto 14px" }}>
                          <circle cx="19" cy="19" r="15" fill="none" stroke="#1abc9c" strokeWidth="3" strokeDasharray="60 20"/>
                        </svg>
                        جاري تحميل الملفات…
                      </div>
                    ) : (
                      /* File list */
                      <div>
                        {/* Connected bar */}
                        <div style={{
                          display: "flex", alignItems: "center", justifyContent: "space-between",
                          background: "rgba(26,188,156,.1)", border: "1px solid rgba(26,188,156,.25)",
                          borderRadius: 10, padding: "10px 14px", marginBottom: 12,
                        }}>
                          <span style={{ color: "#1abc9c", fontWeight: 700, fontSize: 13 }}>✅ متصل بـ Google Drive</span>
                          <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                            {/* View toggle */}
                            <div style={{ display: "flex", background: "rgba(255,255,255,.06)", borderRadius: 8, overflow: "hidden", border: "1px solid rgba(255,255,255,.1)" }}>
                              {[
                                { id: false, label: "📁 قائمة" },
                                { id: true,  label: "📊 لوحة"  },
                              ].map(v => (
                                <button key={String(v.id)} onClick={() => setDriveDashboard(v.id)} style={{
                                  padding: "4px 11px", border: "none", cursor: "pointer",
                                  fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 11,
                                  background: driveDashboard === v.id ? "rgba(26,188,156,.3)" : "transparent",
                                  color: driveDashboard === v.id ? "#1abc9c" : "rgba(255,255,255,.45)",
                                  transition: "all .15s",
                                }}>{v.label}</button>
                              ))}
                            </div>
                            <button onClick={() => fetchDriveFiles(driveToken)} style={{
                              background: "rgba(255,255,255,.08)", border: "1px solid rgba(255,255,255,.15)",
                              borderRadius: 8, padding: "5px 11px", color: "rgba(255,255,255,.6)",
                              cursor: "pointer", fontSize: 12, fontFamily: "'Cairo',sans-serif",
                            }}>🔄</button>
                            <button onClick={() => { setDriveToken(null); setDriveFiles([]); setDriveSelected(null); setDriveDashboard(false); }} style={{
                              background: "rgba(231,76,60,.1)", border: "1px solid rgba(231,76,60,.25)",
                              borderRadius: 8, padding: "5px 11px", color: "#ff6b6b",
                              cursor: "pointer", fontSize: 12, fontFamily: "'Cairo',sans-serif",
                            }}>قطع</button>
                          </div>
                        </div>

                        {/* Dashboard view */}
                        {driveDashboard ? (
                          <DriveDashboard
                            files={driveFiles}
                            token={driveToken}
                            onSelectFile={(f) => { setDriveDashboard(false); handleDriveFileSelect(f); }}
                          />
                        ) : (<>
                        {/* Filters row */}
                        {(() => {
                          const yearSet = new Set();
                          driveFiles.forEach(f => { const y = detectYearFromFilename(f.name); if (y) yearSet.add(y); });
                          const availableYears = [...yearSet].sort().reverse();
                          const selStyle = (active) => ({
                            flex: 1, minWidth: 0, background: "#0d1f33",
                            color: active ? "#1abc9c" : "rgba(255,255,255,.5)",
                            border: `1px solid ${active ? "rgba(26,188,156,.5)" : "rgba(255,255,255,.18)"}`,
                            borderRadius: 9, padding: "8px 10px",
                            fontFamily: "'Cairo',sans-serif", fontSize: 12,
                            cursor: "pointer", outline: "none", appearance: "none",
                          });
                          const hasFilter = driveFilterType || driveFilterYear || driveFilterProgram;
                          return (
                            <div style={{ display: "flex", gap: 6, marginBottom: 10, flexWrap: "wrap" }}>
                              <select value={driveFilterType}
                                onChange={e => { setDriveFilterType(e.target.value); setDriveSelected(null); }}
                                style={selStyle(!!driveFilterType)}>
                                <option value="">كل الاستبيانات</option>
                                {Object.values(SCHEMAS).map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                              </select>
                              <select value={driveFilterYear}
                                onChange={e => { setDriveFilterYear(e.target.value); setDriveSelected(null); }}
                                style={selStyle(!!driveFilterYear)}>
                                <option value="">كل السنوات</option>
                                {availableYears.map(y => <option key={y} value={y}>{y}</option>)}
                              </select>
                              <select value={driveFilterProgram}
                                onChange={e => { setDriveFilterProgram(e.target.value); setDriveSelected(null); }}
                                style={selStyle(!!driveFilterProgram)}>
                                <option value="">كل البرامج</option>
                                {PROGRAMS.map(p => <option key={p} value={p}>{p}</option>)}
                              </select>
                              {hasFilter && (
                                <button
                                  onClick={() => { setDriveFilterType(""); setDriveFilterYear(""); setDriveFilterProgram(""); setDriveSelected(null); }}
                                  style={{ padding: "7px 11px", background: "rgba(231,76,60,.12)", border: "1px solid rgba(231,76,60,.25)", borderRadius: 9, color: "#ff8a80", cursor: "pointer", fontSize: 12, fontFamily: "'Cairo',sans-serif", whiteSpace: "nowrap" }}>
                                  ↩ مسح
                                </button>
                              )}
                            </div>
                          );
                        })()}

                        {/* Search */}
                        <input
                          value={driveSearch}
                          onChange={e => { setDriveSearch(e.target.value); setDriveSelected(null); }}
                          placeholder="🔍 ابحث باسم الملف…"
                          style={{
                            width: "100%", background: "#0d1f33", color: "#e8f0fe",
                            border: "1px solid rgba(255,255,255,.18)", borderRadius: 10,
                            padding: "10px 14px", fontFamily: "'Cairo',sans-serif",
                            fontSize: 13, direction: "rtl", marginBottom: 8, outline: "none",
                          }}
                        />

                        {/* Files */}
                        {(() => {
                          const filtered = driveFiles.filter(f => {
                            if (driveSearch && !f.name.toLowerCase().includes(driveSearch.toLowerCase())) return false;
                            if (driveFilterType && detectTypeHintFromFilename(f.name) !== driveFilterType) return false;
                            if (driveFilterYear && detectYearFromFilename(f.name) !== driveFilterYear) return false;
                            if (driveFilterProgram && detectProgramFromFilename(f.name) !== driveFilterProgram) return false;
                            return true;
                          });
                          const hasAnyFilter = driveSearch || driveFilterType || driveFilterYear || driveFilterProgram;
                          return (
                            <>
                              {hasAnyFilter && (
                                <div style={{ color: "rgba(255,255,255,.35)", fontSize: 11, marginBottom: 6, textAlign: "right" }}>
                                  {filtered.length} من {driveFiles.length} ملف
                                </div>
                              )}
                              {!filtered.length ? (
                                <div style={{ textAlign: "center", color: "rgba(255,255,255,.3)", padding: 24, fontSize: 13 }}>
                                  {driveFiles.length ? "لا توجد ملفات تطابق الفلاتر" : "لا توجد ملفات Excel أو CSV في Drive"}
                                </div>
                              ) : (
                                <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 260, overflowY: "auto" }}>
                                  {filtered.map(f => {
                                    const icon = f.mimeType === GSHEETS_MIME ? "📊" : f.mimeType === "text/csv" ? "📋" : "📄";
                                    const isSelected = driveSelected?.id === f.id;
                                    return (
                                      <div key={f.id} onClick={() => setDriveSelected(f)} style={{
                                        display: "flex", alignItems: "center", gap: 10,
                                        background: isSelected ? "rgba(26,188,156,.15)" : "rgba(255,255,255,.04)",
                                        border: `1px solid ${isSelected ? "rgba(26,188,156,.5)" : "rgba(255,255,255,.08)"}`,
                                        borderRadius: 10, padding: "10px 14px", cursor: "pointer", transition: "all .15s",
                                      }}>
                                        <span style={{ fontSize: 18, flexShrink: 0 }}>{icon}</span>
                                        <div style={{ flex: 1, textAlign: "right" }}>
                                          <div style={{ color: "#e8f0fe", fontSize: 13 }}>{f.name}</div>
                                          {(() => {
                                            const prog = detectProgramFromFilename(f.name);
                                            const yr   = detectYearFromFilename(f.name);
                                            const type = detectTypeHintFromFilename(f.name);
                                            const tags = [
                                              type && SCHEMAS[type]?.label,
                                              yr,
                                              prog,
                                            ].filter(Boolean);
                                            return tags.length ? (
                                              <div style={{ display: "flex", gap: 6, marginTop: 3, flexWrap: "wrap" }}>
                                                {tags.map((t, i) => (
                                                  <span key={i} style={{ fontSize: 10, color: "#1abc9c", background: "rgba(26,188,156,.12)", borderRadius: 6, padding: "1px 7px" }}>{t}</span>
                                                ))}
                                              </div>
                                            ) : null;
                                          })()}
                                        </div>
                                        <span style={{ color: "rgba(255,255,255,.3)", fontSize: 11, whiteSpace: "nowrap", flexShrink: 0 }}>
                                          {f.modifiedTime?.slice(0, 10) ?? ""}
                                        </span>
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </>
                          );
                        })()}

                        {/* Confirm button */}
                        {driveSelected && (
                          <button
                            className="btn btn-primary"
                            style={{ width: "100%", marginTop: 14, justifyContent: "center", opacity: driveProcessing ? 0.7 : 1 }}
                            disabled={driveProcessing}
                            onClick={() => handleDriveFileSelect(driveSelected)}
                          >
                            {driveProcessing
                              ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="white" strokeWidth="2" strokeDasharray="22 8"/></svg> جاري التحميل…</>
                              : `✔ تحليل: ${driveSelected.name}`}
                          </button>
                        )}
                        </>)}
                      </div>
                    )}
                  </div>
                )}

                {error && (
                  <div style={{ color: "#e74c3c", marginTop: 20, fontSize: 13,
                                background: "rgba(231,76,60,.08)", borderRadius: 10, padding: "12px 16px" }}>
                    {error}
                  </div>
                )}

                <div style={{ textAlign: "center", marginTop: 20 }}>
                  <button className="btn btn-ghost btn-sm" onClick={() => { setError(""); setStep(0); }}>→ السابق</button>
                </div>
              </div>
            )}

            {/* Comparison: upload all slots in step 1 */}
            {step === 1 && !isAnnual && (
              <div className="card" style={{ padding: 36 }}>
                <div style={{ color: "#fff", fontSize: 20, fontWeight: 900, textAlign: "center", marginBottom: 8 }}>
                  رفع ملفات الاستبيانات
                </div>
                <div style={{ color: "rgba(255,255,255,.4)", fontSize: 13, textAlign: "center", marginBottom: 24 }}>
                  {mode === "compare2" ? "حدد ملف لكل سنة من السنتين" : "حدد ملف لكل سنة من السنوات الثلاث"}
                </div>

                <div style={{ display: "grid", gridTemplateColumns: `repeat(${slotCount}, 1fr)`, gap: 16 }}>
                  {slots.slice(0, slotCount).map((slot, idx) => (
                    <FileSlot
                      key={idx}
                      slot={slot}
                      label={`العام الأكاديمي ${idx + 1}`}
                      onChange={updated => setSlots(prev => prev.map((s, i) => i === idx ? updated : s))}
                      driveToken={driveToken}
                      driveFiles={driveFiles}
                      driveLoading={driveLoading}
                      onConnectDrive={connectDrive}
                    />
                  ))}
                </div>

                {/* Inline conflict warnings */}
                {(() => {
                  const active = slots.slice(0, slotCount).filter(s => s._file);
                  if (active.length < 2) return null;
                  const warnings = [];
                  const yrs = active.map(s => (s.year ?? "").trim()).filter(Boolean);
                  if (new Set(yrs).size < yrs.length)
                    warnings.push("⚠ بعض الملفات لها نفس السنة الدراسية — يجب أن تكون السنوات مختلفة.");
                  const types = active.map(s => s._type).filter(Boolean);
                  if (types.length === active.length && new Set(types).size > 1)
                    warnings.push(`⚠ أنواع استبيانات مختلفة: ${[...new Set(types)].map(t => SCHEMAS[t]?.label).join(" / ")} — يجب أن تكون الملفات لنفس نوع الاستبيان.`);
                  const progs = active.map(s => s._program).filter(Boolean);
                  if (progs.length > 1 && new Set(progs).size > 1)
                    warnings.push(`⚠ برامج مختلفة محتملة: ${[...new Set(progs)].join(" / ")} — تأكد أن الملفات لنفس البرنامج.`);
                  if (!warnings.length) return null;
                  return (
                    <div style={{ background: "rgba(243,156,18,.08)", border: "1px solid rgba(243,156,18,.3)", borderRadius: 10, padding: "12px 16px", marginTop: 16 }}>
                      {warnings.map((w, i) => (
                        <div key={i} style={{ color: "#ffd54f", fontSize: 12, fontWeight: 700, lineHeight: 1.7 }}>{w}</div>
                      ))}
                    </div>
                  );
                })()}

                {error && (
                  <div style={{ color: "#e74c3c", marginTop: 16, textAlign: "center", fontSize: 13 }}>
                    {error}
                  </div>
                )}

                <div style={{ display: "flex", justifyContent: "center", gap: 12, marginTop: 28 }}>
                  <button className="btn btn-ghost" onClick={() => { setError(""); setStep(0); }}>→ السابق</button>
                  <button
                    className="btn btn-primary"
                    disabled={!slots.slice(0, slotCount).every(s => s._file) || processing}
                    style={{ opacity: (!slots.slice(0, slotCount).every(s => s._file) || processing) ? 0.5 : 1 }}
                    onClick={handleProcessMulti}
                  >
                    {processing
                      ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="white" strokeWidth="2" strokeDasharray="28 10"/></svg> جاري التحليل…</>
                      : "تحليل ←"}
                  </button>
                </div>
              </div>
            )}

            {/* ══ STEP 2 — Validation (annual only) ══ */}
            {step === 2 && rawRows && (() => {
              const s = SCHEMAS[surveyType];
              const totalQ = s?.axes.reduce((acc, ax) => acc + ax.questions.length, 0) ?? 0;
              const dataRowCount = rawRows.length - 1;
              const colHeaders = rawRows[0] ?? [];
              return (
                <div className="card" style={{ padding: 36 }}>
                  <div style={{ color: "#fff", fontSize: 20, fontWeight: 900,
                                textAlign: "center", marginBottom: 6 }}>
                    التحقق من نوع الاستبيان
                  </div>
                  <div style={{ color: "rgba(255,255,255,.4)", fontSize: 13,
                                textAlign: "center", marginBottom: 28 }}>
                    تحقق من صحة الاكتشاف التلقائي قبل المتابعة
                  </div>

                  {/* Detected schema card */}
                  <div style={{
                    background: detectedAutoType ? "rgba(26,188,156,.10)" : "rgba(255,193,7,.07)",
                    border: `1.5px solid ${detectedAutoType ? "rgba(26,188,156,.4)" : "rgba(255,193,7,.4)"}`,
                    borderRadius: 16, padding: "20px 24px", marginBottom: 24,
                    display: "flex", alignItems: "center", gap: 20,
                  }}>
                    <div style={{ fontSize: 52, lineHeight: 1 }}>{s?.icon ?? "📋"}</div>
                    <div style={{ flex: 1 }}>
                      <div style={{
                        display: "inline-block", padding: "3px 12px", borderRadius: 20, fontSize: 11,
                        fontWeight: 700, marginBottom: 8,
                        background: detectedAutoType ? "rgba(26,188,156,.25)" : "rgba(255,193,7,.2)",
                        color: detectedAutoType ? "#1abc9c" : "#ffd54f",
                      }}>
                        {detectedAutoType ? "✓ تم الاكتشاف التلقائي" : "⚠ لم يتم الاكتشاف — تم اختيار الافتراضي"}
                      </div>
                      <div style={{ color: "#fff", fontSize: 20, fontWeight: 900, marginBottom: 4 }}>
                        {s?.label}
                      </div>
                      <div style={{ color: "rgba(255,255,255,.55)", fontSize: 13, display: "flex", gap: 16, flexWrap: "wrap" }}>
                        <span>{s?.scale.type === "likert-5" ? "مقياس 5 درجات" : "مقياس 3 درجات"}</span>
                        <span>·</span>
                        <span>{s?.axes.length} محور</span>
                        <span>·</span>
                        <span>{totalQ} سؤال</span>
                      </div>
                    </div>
                  </div>

                  {/* File stats */}
                  <div style={{
                    display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12, marginBottom: 24,
                  }}>
                    {[
                      { icon: "📄", label: "الملف", value: singleFile?.name ?? "—" },
                      { icon: "👥", label: "عدد الاستجابات", value: `${dataRowCount} صف` },
                      { icon: "📊", label: "عدد الأعمدة", value: `${colHeaders.length} عمود` },
                    ].map((item, i) => (
                      <div key={i} style={{
                        background: "rgba(255,255,255,.05)", borderRadius: 12, padding: "14px 16px",
                        border: "1px solid rgba(255,255,255,.08)",
                      }}>
                        <div style={{ fontSize: 22, marginBottom: 4 }}>{item.icon}</div>
                        <div style={{ color: "#1abc9c", fontWeight: 700, fontSize: 14,
                                      overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                          {item.value}
                        </div>
                        <div style={{ color: "rgba(255,255,255,.4)", fontSize: 11, marginTop: 2 }}>{item.label}</div>
                      </div>
                    ))}
                  </div>

                  {/* Column headers preview */}
                  <div style={{ marginBottom: 24 }}>
                    <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12, fontWeight: 700,
                                  marginBottom: 10, letterSpacing: .3 }}>
                      أول الأعمدة المكتشفة في الملف:
                    </div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                      {colHeaders.slice(0, 9).map((h, i) => (
                        <span key={i} style={{
                          background: "rgba(255,255,255,.07)", border: "1px solid rgba(255,255,255,.1)",
                          borderRadius: 8, padding: "4px 10px", fontSize: 11.5,
                          color: "rgba(255,255,255,.75)",
                        }}>{String(h)}</span>
                      ))}
                      {colHeaders.length > 9 && (
                        <span style={{
                          background: "rgba(26,188,156,.12)", border: "1px solid rgba(26,188,156,.25)",
                          borderRadius: 8, padding: "4px 10px", fontSize: 11.5, color: "#1abc9c",
                        }}>+{colHeaders.length - 9} أخرى</span>
                      )}
                    </div>
                  </div>

                  {/* Override section */}
                  <div style={{
                    background: "rgba(255,255,255,.03)", border: "1px solid rgba(255,255,255,.07)",
                    borderRadius: 14, padding: "18px 20px", marginBottom: 24,
                  }}>
                    <div style={{ color: "rgba(255,255,255,.6)", fontSize: 13, fontWeight: 700,
                                  marginBottom: 12 }}>
                      النوع غير صحيح؟ اختر يدوياً:
                    </div>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                      {Object.values(SCHEMAS).map(sc => (
                        <button key={sc.id}
                          onClick={() => setSurveyType(sc.id)}
                          style={{
                            padding: "8px 18px", borderRadius: 10, cursor: "pointer",
                            fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 13,
                            transition: "all .2s",
                            border: `2px solid ${surveyType === sc.id ? "#1abc9c" : "rgba(255,255,255,.15)"}`,
                            background: surveyType === sc.id ? "rgba(26,188,156,.15)" : "rgba(255,255,255,.04)",
                            color: surveyType === sc.id ? "#1abc9c" : "rgba(255,255,255,.6)",
                          }}>
                          {sc.icon ?? "📋"} {sc.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  {error && (
                    <div style={{ color: "#e74c3c", marginBottom: 16, textAlign: "center", fontSize: 13 }}>
                      {error}
                    </div>
                  )}

                  <div style={{ display: "flex", justifyContent: "center", gap: 12 }}>
                    <button className="btn btn-ghost" onClick={() => {
                      setStep(1); setRawRows(null); setSingleFile(null);
                      setDetectedAutoType(null); setError("");
                    }}>
                      → إعادة الرفع
                    </button>
                    <button className="btn btn-primary" style={{ fontSize: 16, padding: "13px 36px" }}
                      disabled={processing}
                      onClick={handleValidationConfirm}>
                      {processing
                        ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="white" strokeWidth="2" strokeDasharray="28 10"/></svg> جاري التحضير…</>
                        : "✓ تأكيد ومتابعة ←"}
                    </button>
                  </div>
                </div>
              );
            })()}

            {/* ══ STEP 3 — Data Preview (annual) ══ */}
            {step === 3 && isAnnual && singleAllRows && singleSchema && (() => {
              const filteredRows = computeFilteredRows(
                singleAllRows, singleRemoved, singleMetaCols, singleFilters
              );
              const liveResult = filteredRows.length
                ? analyzeRows(singleHeaders, filteredRows, singleSchema, singleMetaCols)
                : null;
              return (
                <div className="card" style={{ padding: 28 }}>
                  <div style={{ color: "#fff", fontSize: 20, fontWeight: 900,
                                textAlign: "center", marginBottom: 10 }}>
                    🔍 معاينة البيانات
                  </div>
                  <div style={{ color: "rgba(255,255,255,.5)", fontSize: 13,
                                textAlign: "center", marginBottom: 20 }}>
                    راجع البيانات واحذف الصفوف غير المرغوبة قبل توليد التقرير
                  </div>

                  <DataPreviewTable
                    headers={singleHeaders}
                    rows={singleAllRows}
                    removed={singleRemoved}
                    onToggleRow={idx => setSingleRemoved(prev => {
                      const next = new Set(prev);
                      if (next.has(idx)) next.delete(idx); else next.add(idx);
                      return next;
                    })}
                    deptIdx={singleMetaCols?.deptIdx}
                    degreeIdx={singleMetaCols?.degreeIdx}
                    filters={singleFilters}
                    onFilterChange={setSingleFilters}
                  />

                  <div style={{ marginTop: 36 }}>
                    <div style={{ color: "#fff", fontSize: 18, fontWeight: 900,
                                  marginBottom: 14, textAlign: "center" }}>
                      📈 لوحة التحليل التفاعلية
                    </div>
                    {liveResult ? (
                      <ResultsPreview result={liveResult} />
                    ) : (
                      <div style={{ color: "#e74c3c", textAlign: "center", padding: 20,
                                    background: "rgba(231,76,60,.08)", borderRadius: 10 }}>
                        لا توجد صفوف بعد تطبيق الفلاتر — أزل بعض الفلاتر للمتابعة.
                      </div>
                    )}
                  </div>

                  <div style={{ display: "flex", justifyContent: "center", gap: 12, marginTop: 28 }}>
                    <button className="btn btn-ghost" onClick={() => setStep(2)}>→ السابق</button>
                    <button className="btn btn-primary"
                      disabled={!liveResult}
                      style={{ opacity: liveResult ? 1 : 0.5 }}
                      onClick={() => { setSingleResult(liveResult); setStep(4); }}>
                      متابعة ←
                    </button>
                  </div>
                </div>
              );
            })()}

            {/* ══ STEP 4 — Metadata ══ */}
            {step === 4 && (
              <div>
                <MetadataForm meta={meta} onChange={setMeta} />
                <div style={{ display: "flex", justifyContent: "center", gap: 12, marginTop: 24 }}>
                  <button className="btn btn-ghost" onClick={() => setStep(isAnnual ? 3 : 1)}>→ السابق</button>
                  <button className="btn btn-primary" onClick={() => setStep(5)}>عرض النتائج ←</button>
                </div>
              </div>
            )}

            {/* ══ STEP 5 — Results ══ */}
            {step === 5 && (
              <div>
                {isAnnual && singleResult && (
                  <div>
                    <div style={{ color: "rgba(255,255,255,.6)", fontSize: 14, marginBottom: 18, textAlign: "center" }}>
                      <span style={{ color: "#1abc9c", fontWeight: 700 }}>{schema.label}</span>
                      {meta.program ? ` — ${meta.program}` : ""}
                      {" | "}<span style={{ color: "#d6eaf8" }}>{meta.year}</span>
                    </div>
                    <ResultsPreview result={singleResult} />
                  </div>
                )}

                {!isAnnual && comparison && (
                  <div>
                    <div style={{ color: "rgba(255,255,255,.6)", fontSize: 14, marginBottom: 18, textAlign: "center" }}>
                      <span style={{ color: "#1abc9c", fontWeight: 700 }}>{schema.label}</span>
                      {meta.program ? ` — ${meta.program}` : ""}
                      {" | مقارنة "}
                      {comparison.slots.map(s => s.year).join(" / ")}
                    </div>
                    <ComparisonPreview comparison={comparison} />
                  </div>
                )}

                <div style={{ textAlign: "center", padding: "28px 0", display: "flex", flexDirection: "column", alignItems: "center", gap: 14 }}>
                  <div style={{ display: "flex", gap: 12, flexWrap: "wrap", justifyContent: "center" }}>
                    {isAnnual ? (
                      <button className="btn btn-blue" style={{ fontSize: 16, padding: "14px 36px" }}
                        disabled={processing} onClick={downloadAnnual}>
                        {processing
                          ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="white" strokeWidth="2" strokeDasharray="28 10"/></svg> جاري الإنشاء…</>
                          : "📄 تحميل التقرير (Word)"}
                      </button>
                    ) : (
                      <button className="btn btn-blue" style={{ fontSize: 16, padding: "14px 36px" }}
                        disabled={processing} onClick={downloadComparison}>
                        📄 تحميل تقرير المقارنة (Word)
                      </button>
                    )}
                    {isAnnual && singleResult && (
                      <button className="btn btn-ghost" style={{ fontSize: 16, padding: "14px 36px" }}
                        onClick={() => downloadPDF(singleResult, meta)}>
                        🖨️ تحميل PDF
                      </button>
                    )}
                    {isAnnual && singleResult && (
                      <button
                        className="btn btn-ghost"
                        style={{ fontSize: 16, padding: "14px 36px", background: "rgba(30,58,138,.25)", borderColor: "rgba(37,99,235,.45)", color: "#93c5fd" }}
                        onClick={() => setShowEnhancedView(true)}
                      >
                        ✨ معاينة محسّنة
                      </button>
                    )}
                  </div>
                  <div style={{ color: "rgba(255,255,255,.35)", fontSize: 12 }}>
                    Word: تقرير كامل مع جداول ومحاور • PDF: ملخص للطباعة السريعة
                  </div>
                </div>
              </div>
            )}
          </>
        )}
      </div>

      {/* ── Tutorial overlay ── */}
      {showTutorial && <TutorialOverlay onClose={() => setShowTutorial(false)} />}

      {/* ── Processing steps overlay ── */}
      {procSteps && <ProcessingOverlay steps={procSteps} filename={procFile} />}

      {/* ── Enhanced Report View overlay (isolated, zero impact on existing flow) ── */}
      {showEnhancedView && singleResult && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 9000,
          background: "#f1f5f9",
          overflowY: "auto",
        }}>
          {/* Close bar */}
          <div style={{
            position: "sticky", top: 0, zIndex: 9001,
            background: "#1e3a8a",
            padding: "10px 24px",
            display: "flex", alignItems: "center", justifyContent: "space-between",
            boxShadow: "0 2px 12px rgba(0,0,0,.25)",
          }}>
            <span style={{ color: "#fff", fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 15 }}>
              ✨ معاينة التقرير المحسّن
            </span>
            <button
              onClick={() => setShowEnhancedView(false)}
              style={{
                background: "rgba(255,255,255,.15)", border: "1px solid rgba(255,255,255,.3)",
                color: "#fff", borderRadius: 8, padding: "7px 20px",
                fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 14, cursor: "pointer",
              }}
            >
              ✕ إغلاق
            </button>
          </div>
          <EnhancedReportView result={singleResult} meta={meta} settings={settings} />
        </div>
      )}

      {/* AI Chat — floating panel, always visible */}
      <AiChat
        currentResult={singleResult}
        aiSettings={aiSettings}
        docSettings={settings}
        onAnalysisComplete={(result, type) => {
          setSingleResult(result);
          if (type) setSurveyType(type);
        }}
      />
    </div>
  );
}
