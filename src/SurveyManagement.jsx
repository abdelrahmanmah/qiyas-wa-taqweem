import { useState, useRef } from "react";
import {
  SCALE_TYPES, QUESTION_TYPES, SURVEY_STATUSES,
  createSurvey, createSection, createQuestion, createMetadata, createReportTexts,
  loadCustomSurveys, saveCustomSurvey, deleteCustomSurvey,
  duplicateCustomSurvey, setCustomSurveyStatus, toRuntimeSchemaShape,
  getBuiltInSurveyCatalog, createCustomSurveyFromBuiltIn,
  saveBuiltInReportTexts,
  importSurveyStructureFromRows,
  buildSurveyTemplateBlob, importSurveyFromTemplateArrayBuffer,
  buildSurveysBackupBlob, importSurveysBackup,
} from "./engine/customSurveyModel.js";
import { readExcel } from "./engine/analyze.js";
import { InlineNotice, QualityIcon, QualityPageHeader } from "./UiElements.jsx";

const SURVEY_CATALOG_CSS = `
.survey-catalog-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:15px}
.survey-catalog-card{position:relative;display:flex;flex-direction:column;min-height:270px;padding:20px;border-radius:18px;
  border:1px solid rgba(255,255,255,.1);background:linear-gradient(145deg,rgba(255,255,255,.085),rgba(255,255,255,.035));
  transition:transform .2s,border-color .2s,box-shadow .2s}
.survey-catalog-card:hover{transform:translateY(-3px);border-color:rgba(26,188,156,.35);box-shadow:0 18px 34px rgba(3,14,28,.16)}
.survey-card-icon{width:44px;height:44px;border-radius:13px;display:grid;place-items:center;font-size:20px;
  background:rgba(26,188,156,.12);border:1px solid rgba(26,188,156,.2);margin-bottom:14px}
.survey-card-title{color:#fff;font-size:15px;font-weight:900;line-height:1.55;margin-bottom:6px}
.survey-card-desc{color:rgba(255,255,255,.46);font-size:11.5px;line-height:1.7;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.survey-card-metrics{display:flex;gap:7px;flex-wrap:wrap;margin:15px 0}
.survey-card-metric{padding:5px 8px;border-radius:8px;background:rgba(255,255,255,.055);color:rgba(255,255,255,.58);font-size:10.5px}
.survey-card-actions{display:flex;gap:7px;flex-wrap:wrap;margin-top:auto;padding-top:14px;border-top:1px solid rgba(255,255,255,.07)}
.survey-catalog-toolbar{display:flex;gap:10px;align-items:center;justify-content:space-between;flex-wrap:wrap;margin-bottom:18px}
.survey-search{position:relative;flex:1;min-width:240px;max-width:460px}.survey-search .input{padding-right:40px}
.survey-search::before{content:'⌕';position:absolute;right:14px;top:8px;color:rgba(255,255,255,.35);font-size:20px;z-index:1}
.survey-filter-group{display:flex;gap:6px;padding:5px;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.035);border-radius:12px}
.survey-empty{grid-column:1/-1;padding:46px;text-align:center}
.quick-create-grid{display:grid;grid-template-columns:1fr 1fr;gap:13px}
.report-text-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.survey-editor{--editor-line:rgba(255,255,255,.09)}
.editor-topbar{position:sticky;top:74px;z-index:25;display:flex;align-items:center;justify-content:space-between;gap:18px;
  padding:16px 18px;margin:-4px 0 18px;border:1px solid var(--editor-line);border-radius:18px;
  background:rgba(9,27,45,.9);backdrop-filter:blur(18px);box-shadow:0 12px 30px rgba(2,12,27,.16)}
.editor-title-row{display:flex;align-items:center;gap:12px;min-width:0}.editor-title-icon{width:42px;height:42px;border-radius:13px;
  display:grid;place-items:center;background:linear-gradient(145deg,rgba(26,188,156,.23),rgba(40,116,166,.18));color:#5eead4;font-size:19px;flex:0 0 auto}
.editor-title{color:#fff;font-weight:900;font-size:18px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.editor-kicker{color:rgba(255,255,255,.4);font-size:10.5px;margin-top:2px}
.editor-progress{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:22px}
.editor-step{position:relative;display:flex;align-items:center;gap:10px;padding:12px 14px;border:1px solid var(--editor-line);border-radius:14px;
  color:rgba(255,255,255,.43);background:rgba(255,255,255,.03);font-family:inherit;text-align:right;cursor:pointer;transition:.2s}
.editor-step:hover{background:rgba(255,255,255,.06);color:#fff}.editor-step.active{color:#fff;border-color:rgba(26,188,156,.45);background:rgba(26,188,156,.11);box-shadow:inset 0 -2px #1abc9c}
.editor-step.done{color:rgba(110,231,207,.75)}.editor-step-number{width:29px;height:29px;border-radius:9px;display:grid;place-items:center;
  background:rgba(255,255,255,.07);font-size:11px;font-weight:900;flex:0 0 auto}.editor-step.active .editor-step-number{background:#1abc9c;color:#fff}
.editor-step-label{font-size:12.5px;font-weight:800}.editor-step-hint{display:block;font-size:9.5px;opacity:.62;margin-top:1px}
.sections-toolbar{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;margin-bottom:14px}
.sections-summary{display:flex;gap:8px;flex-wrap:wrap}.editor-count-chip{padding:7px 11px;border-radius:10px;background:rgba(255,255,255,.05);border:1px solid var(--editor-line);color:rgba(255,255,255,.56);font-size:11px}
.section-editor-card{overflow:hidden;margin-bottom:16px;border:1px solid rgba(255,255,255,.11);border-radius:20px;background:linear-gradient(145deg,rgba(255,255,255,.075),rgba(255,255,255,.035));box-shadow:0 12px 28px rgba(2,12,27,.11)}
.section-editor-head{display:grid;grid-template-columns:auto minmax(0,1fr) auto;align-items:start;gap:13px;padding:18px 18px 15px;border-bottom:1px solid var(--editor-line)}
.section-index{width:42px;height:42px;border-radius:13px;display:grid;place-items:center;background:rgba(26,188,156,.14);border:1px solid rgba(26,188,156,.25);color:#5eead4;font-size:13px;font-weight:900}
.section-fields{display:grid;grid-template-columns:minmax(230px,.75fr) minmax(280px,1.25fr);gap:10px}.section-name{font-size:15px!important;font-weight:900!important}
.editor-icon-actions{display:flex;gap:6px}.editor-icon-btn{width:34px;height:34px;padding:0;border-radius:10px;border:1px solid rgba(255,255,255,.1);background:rgba(255,255,255,.05);color:rgba(255,255,255,.65);font-family:inherit;cursor:pointer;display:grid;place-items:center;transition:.18s}
.editor-icon-btn:hover:not(:disabled){background:rgba(255,255,255,.11);color:#fff}.editor-icon-btn:disabled{opacity:.24;cursor:not-allowed}.editor-icon-btn.danger{color:#fb7185;border-color:rgba(251,113,133,.2)}.editor-icon-btn.danger:hover{background:rgba(251,113,133,.12)}
.section-body{padding:17px}.section-body-title{display:flex;align-items:center;justify-content:space-between;margin-bottom:11px;color:rgba(255,255,255,.52);font-size:11px;font-weight:800}
.question-editor-card{display:grid;grid-template-columns:auto minmax(0,1fr);gap:12px;padding:14px;margin-bottom:10px;border:1px solid rgba(255,255,255,.085);border-radius:15px;background:rgba(5,22,38,.35);transition:border-color .18s,background .18s}
.question-editor-card:focus-within{border-color:rgba(26,188,156,.42);background:rgba(8,29,47,.6)}
.question-order{display:flex;flex-direction:column;align-items:center;gap:5px}.question-number{width:31px;height:31px;border-radius:9px;display:grid;place-items:center;background:rgba(96,165,250,.12);color:#93c5fd;font-size:11px;font-weight:900}
.question-content{min-width:0}.question-main{display:grid;grid-template-columns:minmax(300px,1fr) minmax(190px,.4fr) auto;gap:9px;align-items:center}
.question-text{font-size:14px!important;font-weight:700}.required-toggle{padding:10px 14px!important;border-radius:11px!important;justify-content:center}
.question-advanced{margin-top:9px;border-top:1px dashed rgba(255,255,255,.08);padding-top:9px}.question-advanced summary{width:max-content;color:rgba(255,255,255,.43);font-size:10.5px;font-weight:700;cursor:pointer;user-select:none}.question-advanced-grid{display:grid;grid-template-columns:1fr 1fr 110px;gap:9px;margin-top:10px}
.add-question-btn{width:100%;justify-content:center;border-style:dashed!important;border-radius:13px!important;padding:11px!important;color:#6ee7cf!important;background:rgba(26,188,156,.045)!important}
.add-section-card{width:100%;min-height:92px;display:flex;align-items:center;justify-content:center;gap:10px;border:1.5px dashed rgba(96,165,250,.34);border-radius:18px;background:rgba(96,165,250,.045);color:#93c5fd;font-family:inherit;font-weight:800;cursor:pointer;transition:.2s}.add-section-card:hover{background:rgba(96,165,250,.09);border-color:#60a5fa;transform:translateY(-2px)}
.editor-footer{position:sticky;bottom:14px;z-index:20;display:flex;justify-content:space-between;align-items:center;gap:12px;margin-top:24px;padding:12px 14px;border:1px solid rgba(255,255,255,.1);border-radius:16px;background:rgba(9,27,45,.92);backdrop-filter:blur(18px);box-shadow:0 16px 38px rgba(2,12,27,.3)}
.editor-empty{padding:44px 20px;text-align:center;border:1.5px dashed rgba(255,255,255,.13);border-radius:18px;background:rgba(255,255,255,.025)}
.editor-general-card{padding:22px;display:grid;grid-template-columns:1fr 1fr;gap:16px}.editor-general-wide{grid-column:1/-1}.metadata-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.metadata-grid>*:last-child{grid-column:1/-1}
@media(max-width:1000px){.survey-catalog-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:900px){.editor-progress{grid-template-columns:repeat(2,1fr)}.section-fields{grid-template-columns:1fr}.question-main{grid-template-columns:1fr 1fr}.question-text{grid-column:1/-1}.question-advanced-grid,.editor-general-card,.metadata-grid{grid-template-columns:1fr}.question-advanced-grid>*,.metadata-grid>*:last-child{grid-column:auto}}
@media(max-width:650px){.survey-catalog-grid,.quick-create-grid,.report-text-grid{grid-template-columns:1fr}.survey-catalog-card{min-height:250px}.editor-topbar{position:relative;top:auto;align-items:flex-start}.editor-title-icon{display:none}.editor-progress{grid-template-columns:1fr 1fr}.editor-step{padding:9px}.editor-step-hint{display:none}.section-editor-head{grid-template-columns:auto 1fr}.section-fields{grid-column:1/-1;grid-row:2}.editor-icon-actions{grid-column:2;grid-row:1;justify-content:flex-end}.question-editor-card{grid-template-columns:1fr}.question-order{flex-direction:row;justify-content:space-between}.question-main,.question-advanced-grid{grid-template-columns:1fr}.question-text{grid-column:auto}.editor-footer{bottom:8px}}
`;

// ── small helpers ─────────────────────────────────────────────────────────────
function moveItem(arr, idx, dir) {
  const next = [...arr];
  const target = idx + dir;
  if (target < 0 || target >= next.length) return arr;
  [next[idx], next[target]] = [next[target], next[idx]];
  return next;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function countQuestions(survey) {
  return (survey.sections ?? []).reduce((sum, s) => sum + (s.questions?.length ?? 0), 0);
}

function fmtDate(iso) {
  if (!iso) return "—";
  try { return new Date(iso).toLocaleDateString("ar-EG", { year: "numeric", month: "short", day: "numeric" }); }
  catch { return iso; }
}

// ── Status badge ──────────────────────────────────────────────────────────────
function StatusBadge({ status }) {
  const isActive = status === "active";
  return (
    <span style={{
      display: "inline-block", padding: "4px 12px", borderRadius: 20, fontSize: 11, fontWeight: 700,
      background: isActive ? "rgba(26,188,156,.18)" : "rgba(255,193,7,.15)",
      color: isActive ? "#1abc9c" : "#ffd54f",
      border: `1px solid ${isActive ? "rgba(26,188,156,.4)" : "rgba(255,193,7,.35)"}`,
    }}>
      {isActive ? "● نشط" : "● مسودة"}
    </span>
  );
}

// ── Confirm dialog (lightweight inline modal) ─────────────────────────────────
function ConfirmDialog({ title, message, onConfirm, onCancel }) {
  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(5,12,22,.72)", zIndex: 500,
      display: "flex", alignItems: "center", justifyContent: "center", padding: 20,
    }} onClick={e => { if (e.target === e.currentTarget) onCancel(); }}>
      <div className="card" style={{ maxWidth: 420, padding: 28, textAlign: "center" }}>
        <div className="qa-status-icon danger" style={{ width: 42, height: 42, margin: "0 auto 12px" }}><QualityIcon name="warning" size={21} /></div>
        <div style={{ color: "#fff", fontWeight: 900, fontSize: 17, marginBottom: 8 }}>{title}</div>
        <div style={{ color: "rgba(255,255,255,.6)", fontSize: 13, marginBottom: 22 }}>{message}</div>
        <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
          <button className="btn btn-ghost btn-sm" onClick={onCancel}>إلغاء</button>
          <button className="btn btn-danger btn-sm" onClick={onConfirm}>تأكيد الحذف</button>
        </div>
      </div>
    </div>
  );
}

function QuickCreateDialog({ onCreate, onAdvanced, onCancel }) {
  const [name, setName] = useState("");
  const [surveyType, setSurveyType] = useState("");
  const [scaleType, setScaleType] = useState("likert-5");
  const submit = () => {
    if (!name.trim()) return;
    onCreate({ name: name.trim(), surveyType: surveyType.trim() || name.trim(), scaleType });
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(3,12,24,.78)", zIndex: 500, display: "grid", placeItems: "center", padding: 20 }}
      onClick={e => { if (e.target === e.currentTarget) onCancel(); }}>
      <div className="card" style={{ width: "min(620px,100%)", padding: 28, boxShadow: "0 28px 80px rgba(0,0,0,.35)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 14, marginBottom: 22 }}>
          <div><div style={{ color: "#fff", fontWeight: 900, fontSize: 20 }}>إنشاء استبيان جديد</div><div style={{ color: "rgba(255,255,255,.45)", fontSize: 12, marginTop: 4 }}>أدخل الأساسيات وسنجهز لك أول محور وسؤال تلقائيًا.</div></div>
          <button className="btn btn-ghost btn-sm" onClick={onCancel} aria-label="إغلاق">✕</button>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <div><label className="label">اسم الاستبيان *</label><input autoFocus className="input" placeholder="مثال: رضا جهات التوظيف عن الخريجين" value={name} onChange={e => setName(e.target.value)} onKeyDown={e => { if (e.key === "Enter") submit(); }} /></div>
          <div className="quick-create-grid">
            <div><label className="label">الفئة أو الغرض</label><input className="input" placeholder="طلاب، خريجون، جهات توظيف..." value={surveyType} onChange={e => setSurveyType(e.target.value)} /></div>
            <div><label className="label">مقياس الإجابة</label><select className="input" value={scaleType} onChange={e => setScaleType(e.target.value)}>{Object.entries(SCALE_TYPES).map(([id, sc]) => <option key={id} value={id}>{sc.label}</option>)}</select></div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", marginTop: 22, paddingTop: 18, borderTop: "1px solid rgba(255,255,255,.08)" }}>
          <button className="btn btn-ghost btn-sm" onClick={onAdvanced}>إنشاء فارغ بإعدادات متقدمة</button>
          <button className="btn btn-primary" disabled={!name.trim()} onClick={submit}>إنشاء والبدء في إضافة الأسئلة ←</button>
        </div>
      </div>
    </div>
  );
}

// ── List view ─────────────────────────────────────────────────────────────────
function SurveyListView({
  surveys, builtInSurveys, onCreate, onEdit, onDelete, onDuplicate, onToggleStatus, onCloneBuiltIn, onEditReportTexts,
  onDownloadTemplate, onImportTemplate, onExportBackup, onImportBackup,
}) {
  const [confirmId, setConfirmId] = useState(null);
  const [toolbarMsg, setToolbarMsg] = useState("");
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("all");
  const templateFileRef = useRef();
  const backupFileRef = useRef();
  const allSurveys = [...builtInSurveys, ...surveys];
  const normalizedSearch = search.trim().toLowerCase();
  const visibleSurveys = allSurveys.filter(s => {
    if (filter === "builtin" && !s.builtIn) return false;
    if (filter === "custom" && s.builtIn) return false;
    if (filter === "draft" && (s.builtIn || s.status !== "draft")) return false;
    if (!normalizedSearch) return true;
    return `${s.name} ${s.description ?? ""} ${s.surveyType ?? ""}`.toLowerCase().includes(normalizedSearch);
  });

  const handleTemplateFile = async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      await onImportTemplate(file);
      setToolbarMsg(`✓ تم إنشاء استبيان جديد من القالب «${file.name}» — افتحه من القائمة لمراجعته.`);
    } catch (err) {
      setToolbarMsg("✗ تعذّر استيراد القالب: " + err.message);
    }
  };

  const handleBackupFile = async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      const text = await file.text();
      const count = onImportBackup(text);
      setToolbarMsg(`✓ تم استرجاع النسخة الاحتياطية — العدد الإجمالي الآن: ${count} استبيان.`);
    } catch (err) {
      setToolbarMsg("✗ تعذّر استيراد النسخة الاحتياطية: " + err.message);
    }
  };

  return (
    <div>
      <style>{SURVEY_CATALOG_CSS}</style>
      <QualityPageHeader
        icon="report"
        eyebrow="تصميم وإدارة النماذج"
        title="مكتبة الاستبيانات"
        description="راجع النماذج المتاحة، عدّل نصوصها، وأنشئ الاستبيان الناقص بنفس هوية وحدة الجودة."
        actions={<button className="btn btn-primary qa-icon-button" onClick={onCreate}><QualityIcon name="plus" size={16} /> استبيان جديد</button>}
      />

      <div style={{
        display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center",
        background: "rgba(255,255,255,.04)", border: "1px solid rgba(255,255,255,.08)",
        borderRadius: 10, padding: 12, marginBottom: 20,
      }}>
        <span style={{ color: "rgba(255,255,255,.45)", fontSize: 11.5, fontWeight: 700 }}>قالب Excel:</span>
        <button className="btn btn-ghost btn-sm qa-icon-button" onClick={onDownloadTemplate}><QualityIcon name="download" size={14} /> تحميل القالب</button>
        <input ref={templateFileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleTemplateFile} />
        <button className="btn btn-blue btn-sm qa-icon-button" onClick={() => templateFileRef.current.click()}><QualityIcon name="upload" size={14} /> استيراد استبيان من قالب</button>

        <span style={{ width: 1, height: 20, background: "rgba(255,255,255,.12)", margin: "0 6px" }} />

        <span style={{ color: "rgba(255,255,255,.45)", fontSize: 11.5, fontWeight: 700 }}>نسخة احتياطية:</span>
        <button className="btn btn-ghost btn-sm qa-icon-button" onClick={onExportBackup}><QualityIcon name="download" size={14} /> تصدير نسخة احتياطية (JSON)</button>
        <input ref={backupFileRef} type="file" accept=".json" style={{ display: "none" }} onChange={handleBackupFile} />
        <button className="btn btn-ghost btn-sm qa-icon-button" onClick={() => backupFileRef.current.click()}><QualityIcon name="upload" size={14} /> استيراد نسخة احتياطية</button>
      </div>

      {toolbarMsg && (
        toolbarMsg.startsWith("✗")
          ? <InlineNotice style={{ marginBottom: 14 }}>{toolbarMsg.replace(/^✗\s*/, "")}</InlineNotice>
          : <div style={{ color: "#1abc9c", fontSize: 12.5, marginBottom: 14 }}>{toolbarMsg}</div>
      )}

      <div className="survey-catalog-toolbar">
        <div className="survey-search"><input className="input" value={search} onChange={e => setSearch(e.target.value)} placeholder="ابحث باسم الاستبيان أو الفئة..." aria-label="البحث في الاستبيانات" /></div>
        <div className="survey-filter-group" aria-label="تصفية الاستبيانات">
          {[{ id: "all", label: `الكل (${allSurveys.length})` }, { id: "builtin", label: `مدمج (${builtInSurveys.length})` }, { id: "custom", label: `مخصص (${surveys.length})` }, { id: "draft", label: "مسودات" }].map(item => (
            <button key={item.id} className={`btn btn-sm ${filter === item.id ? "btn-primary" : "btn-ghost"}`} style={{ padding: "7px 12px" }} onClick={() => setFilter(item.id)}>{item.label}</button>
          ))}
        </div>
      </div>

      <div className="survey-catalog-grid">
        {visibleSurveys.length === 0 ? (
          <div className="card survey-empty"><div style={{ fontSize: 38, marginBottom: 8 }}>⌕</div><div style={{ color: "rgba(255,255,255,.6)", fontSize: 14 }}>لا توجد استبيانات مطابقة للبحث أو الفلتر.</div></div>
        ) : visibleSurveys.map((s, index) => (
          <article key={`${s.builtIn ? "builtin" : "custom"}-${s.id}`} className="survey-catalog-card">
            <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 10 }}>
              <span className="survey-card-icon"><QualityIcon name={["report", "chart", "users", "table", "sheet"][index % 5]} size={22} /></span>
              {s.builtIn ? <span className="badge" style={{ background: "rgba(96,165,250,.13)", color: "#93c5fd", border: "1px solid rgba(96,165,250,.22)" }}>مدمج بالنظام</span> : <StatusBadge status={s.status} />}
            </div>
            <div className="survey-card-title">{s.name || "(بدون اسم)"}</div>
            <div className="survey-card-desc">{s.description || (s.builtIn ? "جاهز للتحليل وإصدار التقرير مباشرة." : "استبيان مخصص بدون وصف.")}</div>
            <div className="survey-card-metrics">
              <span className="survey-card-metric">{s.isFlat ? "بدون محاور" : `${(s.sections ?? []).length} محاور`}</span>
              <span className="survey-card-metric">{countQuestions(s)} سؤال</span>
              <span className="survey-card-metric">{s.scaleType === "likert-3" ? "مقياس ثلاثي" : "مقياس خماسي"}</span>
              {s.programs?.length > 0 && <span className="survey-card-metric">{s.programs.length} برامج</span>}
            </div>
            <div className="survey-card-actions">
              {s.builtIn ? (
                <><button className="btn btn-primary btn-sm" onClick={() => onEditReportTexts(s)}>ضبط نصوص التقرير</button><button className="btn btn-ghost btn-sm" onClick={() => onCloneBuiltIn(s.id)}>استخدم كنسخة</button></>
              ) : (<>
                <button className="btn btn-primary btn-sm" onClick={() => onEdit(s)}>تعديل</button>
                <button className="btn btn-ghost btn-sm" onClick={() => onDuplicate(s.id)}>نسخ</button>
                <button className="btn btn-ghost btn-sm" onClick={() => onToggleStatus(s.id, s.status === "active" ? "draft" : "active")}>{s.status === "active" ? "تعطيل" : "تفعيل"}</button>
                <button className="btn btn-danger btn-sm" onClick={() => setConfirmId(s.id)} aria-label={`حذف ${s.name}`}>حذف</button>
              </>)}
            </div>
          </article>
        ))}
      </div>

      {confirmId && (
        <ConfirmDialog
          title="حذف الاستبيان؟"
          message="سيتم حذف تعريف هذا الاستبيان بشكل نهائي. لا يمكن التراجع عن هذا الإجراء."
          onCancel={() => setConfirmId(null)}
          onConfirm={() => { onDelete(confirmId); setConfirmId(null); }}
        />
      )}
    </div>
  );
}

// ── Question row editor ───────────────────────────────────────────────────────
function QuestionRow({ q, index, total, onChange, onDelete, onDuplicate, onMove, sectionId, allSections, onMoveToSection }) {
  const [moveTarget, setMoveTarget] = useState("");
  const otherSections = allSections.filter(s => s.id !== sectionId);

  return (
    <div className="question-editor-card">
      <div className="question-order">
        <span className="question-number">{index + 1}</span>
        <button className="editor-icon-btn" disabled={index === 0} onClick={() => onMove(-1)} title="تحريك السؤال لأعلى" aria-label="تحريك السؤال لأعلى">↑</button>
        <button className="editor-icon-btn" disabled={index === total - 1} onClick={() => onMove(1)} title="تحريك السؤال لأسفل" aria-label="تحريك السؤال لأسفل">↓</button>
      </div>
      <div className="question-content">
        <div className="question-main">
          <input className="input question-text" placeholder={`اكتب نص السؤال ${index + 1} هنا...`} value={q.text}
            onChange={e => onChange({ ...q, text: e.target.value })} />
          <select className="input" value={q.type} onChange={e => onChange({ ...q, type: e.target.value })} aria-label={`نوع السؤال ${index + 1}`}>
            {QUESTION_TYPES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
          </select>
          <button className={`btn btn-sm required-toggle ${q.required ? "btn-primary" : "btn-ghost"}`} onClick={() => onChange({ ...q, required: !q.required })} title="تغيير حالة السؤال بين إلزامي واختياري">
            {q.required ? "✓ إلزامي" : "اختياري"}
          </button>
        </div>
        <details className="question-advanced">
          <summary>إعدادات متقدمة ونقل السؤال</summary>
          <div className="question-advanced-grid">
            <input className="input" placeholder="اسم العمود في Excel" value={q.excelColumn} onChange={e => onChange({ ...q, excelColumn: e.target.value })} />
            <input className="input" placeholder="الفئة (اختياري)" value={q.category ?? ""} onChange={e => onChange({ ...q, category: e.target.value })} />
            <input className="input" type="number" placeholder="الوزن" value={q.weight ?? ""} onChange={e => onChange({ ...q, weight: e.target.value === "" ? null : Number(e.target.value) })} />
          </div>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap", marginTop: 9 }}>
            {otherSections.length > 0 ? <div style={{ display: "flex", gap: 8, flex: "1 1 300px" }}>
              <select className="input" style={{ flex: 1 }} value={moveTarget} onChange={e => setMoveTarget(e.target.value)}>
                <option value="">نقل هذا السؤال إلى محور آخر...</option>
                {otherSections.map(s => <option key={s.id} value={s.id}>{s.name || "(بدون اسم)"}</option>)}
              </select>
              <button className="btn btn-ghost btn-sm" disabled={!moveTarget} onClick={() => { onMoveToSection(moveTarget); setMoveTarget(""); }}>نقل</button>
            </div> : <span />}
            <div style={{ display: "flex", gap: 7 }}><button className="btn btn-ghost btn-sm" onClick={onDuplicate}>نسخ السؤال</button><button className="btn btn-danger btn-sm" onClick={onDelete}>حذف السؤال</button></div>
          </div>
        </details>
      </div>
    </div>
  );
}

// ── Section editor ────────────────────────────────────────────────────────────
function SectionBlock({ section, index, total, onChange, onDelete, onMove, allSections, onMoveQuestionToSection }) {
  const [expanded, setExpanded] = useState(true);
  const updateQuestion = (qIdx, q) => {
    const questions = [...section.questions];
    questions[qIdx] = q;
    onChange({ ...section, questions });
  };
  const deleteQuestion = (qIdx) => {
    onChange({ ...section, questions: section.questions.filter((_, i) => i !== qIdx) });
  };
  const moveQuestion = (qIdx, dir) => {
    onChange({ ...section, questions: moveItem(section.questions, qIdx, dir) });
  };
  const addQuestion = () => {
    onChange({ ...section, questions: [...section.questions, createQuestion()] });
  };
  const duplicateQuestion = (qIdx) => {
    const source = section.questions[qIdx];
    const { id: _sourceId, ...copyable } = source;
    const questions = [...section.questions];
    questions.splice(qIdx + 1, 0, createQuestion({ ...copyable, text: `${source.text} (نسخة)` }));
    onChange({ ...section, questions });
  };

  return (
    <section className="section-editor-card">
      <div className="section-editor-head">
        <button className="section-index" style={{ cursor: "pointer" }} onClick={() => setExpanded(v => !v)} title={expanded ? "طي المحور" : "فتح المحور"}>{String(index + 1).padStart(2, "0")}</button>
        <div className="section-fields">
          <div><label className="label">اسم المحور</label><input className="input section-name" placeholder={`مثال: المحور ${index + 1}`} value={section.name} onChange={e => onChange({ ...section, name: e.target.value })} /></div>
          <div><label className="label">وصف مختصر <span style={{ fontWeight: 400, opacity: .65 }}>(اختياري)</span></label><input className="input" placeholder="اشرح ما الذي يقيسه هذا المحور..." value={section.description ?? ""} onChange={e => onChange({ ...section, description: e.target.value })} /></div>
        </div>
        <div className="editor-icon-actions">
          <button className="editor-icon-btn" onClick={() => setExpanded(v => !v)} title={expanded ? "طي المحور" : "فتح المحور"} aria-label={expanded ? "طي المحور" : "فتح المحور"}>{expanded ? "⌃" : "⌄"}</button>
          <button className="editor-icon-btn" disabled={index === 0} onClick={() => onMove(-1)} title="تحريك المحور لأعلى" aria-label="تحريك المحور لأعلى">↑</button>
          <button className="editor-icon-btn" disabled={index === total - 1} onClick={() => onMove(1)} title="تحريك المحور لأسفل" aria-label="تحريك المحور لأسفل">↓</button>
          <button className="editor-icon-btn danger" onClick={onDelete} title="حذف المحور" aria-label="حذف المحور">✕</button>
        </div>
      </div>
      {expanded && <div className="section-body">
        <div className="section-body-title"><span>الأسئلة داخل المحور</span><span className="editor-count-chip">{section.questions.length} سؤال</span></div>
        {section.questions.length === 0 && <div style={{ padding: "22px", textAlign: "center", color: "rgba(255,255,255,.38)", fontSize: 12 }}>لا توجد أسئلة في هذا المحور بعد.</div>}
        {section.questions.map((q, qi) => (
          <QuestionRow key={q.id} q={q} index={qi} total={section.questions.length}
            onChange={nq => updateQuestion(qi, nq)} onDelete={() => deleteQuestion(qi)} onDuplicate={() => duplicateQuestion(qi)} onMove={dir => moveQuestion(qi, dir)}
            sectionId={section.id} allSections={allSections} onMoveToSection={toSectionId => onMoveQuestionToSection(qi, toSectionId)} />
        ))}
        <button className="btn btn-ghost btn-sm add-question-btn" onClick={addQuestion}>＋ إضافة سؤال جديد داخل هذا المحور</button>
      </div>}
    </section>
  );
}

// ── Preview tab ───────────────────────────────────────────────────────────────
function PreviewTab({ survey }) {
  const compiled = toRuntimeSchemaShape(survey);
  return (
    <div className="card" style={{ padding: 24 }}>
      <div style={{ color: "#fff", fontWeight: 900, fontSize: 19, marginBottom: 4 }}>{compiled.label}</div>
      <div style={{ color: "rgba(255,255,255,.5)", fontSize: 12.5, marginBottom: 4 }}>{survey.description}</div>
      <div style={{ display: "flex", gap: 14, flexWrap: "wrap", color: "rgba(255,255,255,.45)", fontSize: 12, marginBottom: 18 }}>
        <span>الإصدار: {compiled.version}</span>
        <span><StatusBadge status={compiled.status} /></span>
        <span>{SCALE_TYPES[survey.scaleType]?.label}</span>
        <span>{compiled.axes.length} محاور · {countQuestions(survey)} سؤال</span>
      </div>

      {compiled.axes.length === 0 && (
        <div style={{ color: "rgba(255,255,255,.4)", fontSize: 13 }}>لا توجد محاور بعد. أضف محوراً من تبويب "المحاور".</div>
      )}

      {compiled.axes.map((ax, ai) => (
        <div key={ax.id} style={{ marginBottom: 18 }}>
          <div style={{ color: "#1abc9c", fontWeight: 800, fontSize: 14.5, marginBottom: 8, borderBottom: "1px solid rgba(26,188,156,.25)", paddingBottom: 6 }}>
            {ai + 1}. {ax.name}
          </div>
          {ax.questions.length === 0 ? (
            <div style={{ color: "rgba(255,255,255,.35)", fontSize: 12.5 }}>لا توجد أسئلة في هذا المحور.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {ax.questions.map(q => (
                <div key={q.id} style={{ display: "flex", gap: 10, alignItems: "center", fontSize: 13, color: "rgba(255,255,255,.8)" }}>
                  <span style={{ color: "rgba(255,255,255,.35)", minWidth: 26 }}>{q.seq}.</span>
                  <span style={{ flex: 1 }}>{q.text || "(بدون نص)"}</span>
                  {q.excelColumn && <span style={{ fontSize: 10.5, color: "rgba(255,255,255,.4)" }}>[{q.excelColumn}]</span>}
                  <span style={{ fontSize: 10.5, color: "#1abc9c" }}>{QUESTION_TYPES.find(t => t.id === q.type)?.label ?? q.type}</span>
                  {q.required && <span style={{ fontSize: 10.5, color: "#ffd54f" }}>إلزامي</span>}
                </div>
              ))}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

const REPORT_TEXT_FIELDS = [
  { key: "reportTitle", label: "العنوان الرئيسي للغلاف", placeholder: "نتائج تحليل استبيان", rows: 1 },
  { key: "reportSubtitle", label: "العنوان الفرعي للغلاف", placeholder: "قياس آراء ورضا {اسم_الاستبيان}", rows: 1 },
  { key: "introduction", label: "مقدمة خاصة بالتقرير", placeholder: "اكتب مقدمة توضّح هدف هذا الاستبيان ونطاق التقرير...", rows: 4, wide: true },
  { key: "variablesText", label: "نص متغيرات الاستبيان", placeholder: "اشتمل الاستبيان على ({عدد_الأسئلة}) عبارة تتمثل في ({عدد_المحاور}) محور رئيسي.", rows: 3, wide: true },
  { key: "methodologyText", label: "نص المعالجة الإحصائية", placeholder: "النسب والمتوسط الحسابي.", rows: 3, wide: true },
  { key: "resultsHeading", label: "عنوان قسم النتائج", placeholder: "خامساً: عرض النتائج وتحليلها ومناقشتها", rows: 1 },
  { key: "summaryHeading", label: "عنوان ملخص النتائج", placeholder: "رابعاً: ملخص النتائج", rows: 1 },
  { key: "recommendationsHeading", label: "عنوان قسم التوصيات", placeholder: "أخيراً: التوصيات", rows: 1 },
  { key: "noRecommendationsText", label: "النص عند عدم وجود توصيات", placeholder: "لا توجد توصيات.", rows: 1 },
  { key: "recommendationTemplate", label: "صيغة التوصية التلقائية", placeholder: "مراجعة محور \"{اسم_المحور}\" لأنه سجل نسبة موافقة {النسبة}%.", rows: 3, wide: true },
];

function ReportTextsTab({ survey, onChange }) {
  const texts = createReportTexts(survey.reportTexts);
  const set = (key, value) => onChange({ ...survey, reportTexts: { ...texts, [key]: value } });
  return (
    <div>
      <div className="card" style={{ padding: 20, marginBottom: 16, background: "rgba(26,188,156,.07)", borderColor: "rgba(26,188,156,.22)" }}>
        <div style={{ color: "#fff", fontSize: 14, fontWeight: 900, marginBottom: 5 }}>نصوص خاصة بهذا الاستبيان</div>
        <div style={{ color: "rgba(255,255,255,.52)", fontSize: 12, lineHeight: 1.8 }}>أي حقل تتركه فارغًا سيستخدم النص الافتراضي للنظام. يمكنك استخدام المتغيرات <b style={{ color: "#5eead4" }}>{`{اسم_الاستبيان} · {عدد_المشاركين} · {عدد_الأسئلة} · {عدد_المحاور} · {العام} · {البرنامج}`}</b>.</div>
      </div>
      <div className="report-text-grid">
        {REPORT_TEXT_FIELDS.map(field => (
          <div key={field.key} className="card" style={{ padding: 17, gridColumn: field.wide ? "1 / -1" : undefined }}>
            <label className="label">{field.label}</label>
            {field.rows === 1 ? (
              <input className="input" value={texts[field.key]} placeholder={field.placeholder} onChange={e => set(field.key, e.target.value)} />
            ) : (
              <textarea className="textarea" rows={field.rows} value={texts[field.key]} placeholder={field.placeholder} onChange={e => set(field.key, e.target.value)} />
            )}
            <div style={{ color: "rgba(255,255,255,.3)", fontSize: 10.5, marginTop: 6 }}>الافتراضي: {field.placeholder}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

function BuiltInReportTextsEditor({ survey: initial, onSave, onCancel }) {
  const [survey, setSurvey] = useState({ ...initial, reportTexts: createReportTexts(initial.reportTexts) });
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 18 }}>
        <div><div style={{ color: "#fff", fontSize: 20, fontWeight: 900 }}>نصوص تقرير: {survey.name}</div><div style={{ color: "rgba(255,255,255,.45)", fontSize: 12, marginTop: 4 }}>تخصيص النصوص فقط؛ أسئلة ومحاور الاستبيان المدمج ستظل كما هي.</div></div>
        <div style={{ display: "flex", gap: 8 }}><button className="btn btn-ghost btn-sm" onClick={onCancel}>إلغاء</button><button className="btn btn-primary btn-sm" onClick={() => onSave(survey.id, survey.reportTexts)}>حفظ نصوص التقرير</button></div>
      </div>
      <ReportTextsTab survey={survey} onChange={setSurvey} />
    </div>
  );
}

// ── comma-separated-list helpers for metadata column fields ───────────────────
function listToText(arr) { return (arr ?? []).join("، "); }
function textToList(text) { return text.split(/[،,]/).map(s => s.trim()).filter(Boolean); }

// ── Import-from-Excel panel (General step) ────────────────────────────────────
function ImportFromExcel({ onImported }) {
  const fileRef = useRef();
  const [summary, setSummary] = useState(null);
  const [error, setError] = useState("");

  const handleFile = async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setError("");
    try {
      const buf = await file.arrayBuffer();
      const rows = readExcel(buf);
      const imported = importSurveyStructureFromRows(rows);
      if (!imported || imported.questions.length === 0) {
        setError("لم يتم العثور على أعمدة أسئلة صالحة في هذا الملف.");
        return;
      }
      onImported(imported);
      const metaCount = Object.values(imported.metadata).reduce((s, a) => s + a.length, 0);
      setSummary({ fileName: file.name, questionCount: imported.questions.length, metaCount, scaleType: imported.scaleType });
    } catch (err) {
      setError("تعذّر قراءة الملف: " + err.message);
    }
  };

  return (
    <div style={{
      background: "rgba(26,188,156,.06)", border: "1px dashed rgba(26,188,156,.35)",
      borderRadius: 12, padding: 16, marginBottom: 18,
    }}>
      <div className="qa-icon-label" style={{ color: "#fff", fontWeight: 800, fontSize: 13.5, marginBottom: 4 }}><QualityIcon name="upload" size={16} /> استيراد تلقائي من ملف Excel (اختياري)</div>
      <div style={{ color: "rgba(255,255,255,.5)", fontSize: 12, marginBottom: 10 }}>
        يكتشف النظام الأسئلة من صفوف العناوين تلقائياً، ويستثني أعمدة المعلومات العامة (الاسم، البريد، الوظيفة/الدرجة، القسم/التخصص، الملاحظات).
        كل الأسئلة المكتشفة توضع في محور واحد جديد — يمكنك بعد ذلك تقسيمها على عدة محاور من تبويب "المحاور".
      </div>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={handleFile} />
      <button className="btn btn-blue btn-sm" onClick={() => fileRef.current.click()}>اختر ملف Excel…</button>
      {summary && (
        <div style={{ color: "#1abc9c", fontSize: 12, marginTop: 10 }}>
          ✓ تم استيراد {summary.questionCount} سؤال و {summary.metaCount} عمود معلومات عامة من «{summary.fileName}» (مقياس مكتشف: {SCALE_TYPES[summary.scaleType]?.label}).
        </div>
      )}
      <InlineNotice text={error} style={{ marginTop: 10 }} />
    </div>
  );
}

const EDITOR_STEPS = [
  { id: "general", label: "البيانات الأساسية", hint: "الاسم والمقياس" },
  { id: "sections", label: "المحاور والأسئلة", hint: "بناء المحتوى" },
  { id: "report", label: "نصوص التقرير", hint: "تخصيص المخرجات" },
  { id: "preview", label: "المعاينة", hint: "مراجعة نهائية" },
];

// ── Editor view (General → Sections → Preview, step-by-step) ──────────────────
function SurveyEditorView({ survey: initial, onSave, onCancel }) {
  const [survey, setSurvey] = useState(() => {
    const { _initialTab, ...clean } = initial;
    return { ...clean, metadata: createMetadata(clean.metadata), reportTexts: createReportTexts(clean.reportTexts) };
  });
  const [tab, setTab] = useState(initial._initialTab || "general");

  const updateSection = (idx, sec) => {
    const sections = [...survey.sections];
    sections[idx] = sec;
    setSurvey({ ...survey, sections });
  };
  const deleteSection = (idx) => setSurvey({ ...survey, sections: survey.sections.filter((_, i) => i !== idx) });
  const moveSection = (idx, dir) => setSurvey({ ...survey, sections: moveItem(survey.sections, idx, dir) });
  const addSection = () => setSurvey({ ...survey, sections: [...survey.sections, createSection()] });

  const moveQuestionToSection = (fromIdx, qIdx, toSectionId) => {
    setSurvey(s => {
      const sections = s.sections.map(sec => ({ ...sec, questions: [...sec.questions] }));
      const toIdx = sections.findIndex(sec => sec.id === toSectionId);
      if (toIdx === -1 || toIdx === fromIdx) return s;
      const [moved] = sections[fromIdx].questions.splice(qIdx, 1);
      sections[toIdx].questions.push(moved);
      return { ...s, sections };
    });
  };

  const updateMetaList = (field, text) => setSurvey({ ...survey, metadata: { ...survey.metadata, [field]: textToList(text) } });

  const handleImported = (imported) => {
    setSurvey(s => ({
      ...s,
      scaleType: imported.scaleType,
      metadata: imported.metadata,
      sections: [...s.sections, createSection({ name: `محور مستورد ${s.sections.length + 1}`, questions: imported.questions })],
    }));
  };

  const canSave = survey.name.trim().length > 0;
  const stepIdx = EDITOR_STEPS.findIndex(s => s.id === tab);
  const goTo = (id) => setTab(id);
  const goNext = () => goTo(EDITOR_STEPS[Math.min(stepIdx + 1, EDITOR_STEPS.length - 1)].id);
  const goBack = () => goTo(EDITOR_STEPS[Math.max(stepIdx - 1, 0)].id);

  return (
    <div className="survey-editor">
      <style>{SURVEY_CATALOG_CSS}</style>
      <div className="editor-topbar">
        <div className="editor-title-row">
          <span className="editor-title-icon"><QualityIcon name="edit" size={21} /></span>
          <div style={{ minWidth: 0 }}>
            <div className="editor-title">{survey.name || "استبيان جديد"}</div>
            <div className="editor-kicker">محرر الاستبيان · {survey.sections.length} محاور · {countQuestions(survey)} سؤال</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="btn btn-ghost btn-sm" onClick={onCancel}>إلغاء</button>
          <button className="btn btn-primary btn-sm" disabled={!canSave} onClick={() => onSave(survey)} title="حفظ سريع من أي خطوة">حفظ التغييرات</button>
        </div>
      </div>

      <div className="editor-progress" aria-label="خطوات إنشاء الاستبيان">
        {EDITOR_STEPS.map((t, i) => (
          <button key={t.id} className={`editor-step ${tab === t.id ? "active" : stepIdx > i ? "done" : ""}`} onClick={() => goTo(t.id)} aria-current={tab === t.id ? "step" : undefined}>
            <span className="editor-step-number">{stepIdx > i ? "✓" : i + 1}</span>
            <span><span className="editor-step-label">{t.label}</span><span className="editor-step-hint">{t.hint}</span></span>
          </button>
        ))}
      </div>

      {tab === "general" && (
        <div>
          <ImportFromExcel onImported={handleImported} />
          <div className="card editor-general-card">
            <div className="editor-general-wide">
              <label className="label">اسم الاستبيان *</label>
              <input className="input" value={survey.name} onChange={e => setSurvey({ ...survey, name: e.target.value })} />
            </div>
            <div className="editor-general-wide">
              <label className="label">وصف الاستبيان</label>
              <textarea className="textarea" rows={3} value={survey.description}
                onChange={e => setSurvey({ ...survey, description: e.target.value })} />
            </div>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <div style={{ flex: "1 1 140px" }}>
                <label className="label">الإصدار</label>
                <input className="input" value={survey.version} onChange={e => setSurvey({ ...survey, version: e.target.value })} />
              </div>
              <div style={{ flex: "1 1 140px" }}>
                <label className="label">الحالة</label>
                <select className="input" value={survey.status} onChange={e => setSurvey({ ...survey, status: e.target.value })}>
                  {SURVEY_STATUSES.map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                </select>
              </div>
            </div>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <div style={{ flex: "1 1 200px" }}>
                <label className="label">نوع الاستبيان</label>
                <input className="input" placeholder="مثال: رضا الطلاب، خريجين..." value={survey.surveyType}
                  onChange={e => setSurvey({ ...survey, surveyType: e.target.value })} />
              </div>
              <div style={{ flex: "1 1 200px" }}>
                <label className="label">نوع المقياس</label>
                <select className="input" value={survey.scaleType} onChange={e => setSurvey({ ...survey, scaleType: e.target.value })}>
                  {Object.entries(SCALE_TYPES).map(([id, sc]) => <option key={id} value={id}>{sc.label}</option>)}
                </select>
              </div>
            </div>

            <div className="editor-general-wide" style={{ borderTop: "1px solid rgba(255,255,255,.08)", paddingTop: 14 }}>
              <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12, fontWeight: 700, marginBottom: 10 }}>
                معلومات عامة (تُستثنى تلقائياً من الأسئلة عند التحليل)
              </div>
              <div className="metadata-grid">
                <div>
                  <label className="label">أعمدة الاسم</label>
                  <input className="input" placeholder="مثال: الاسم" value={listToText(survey.metadata.nameCol)}
                    onChange={e => updateMetaList("nameCol", e.target.value)} />
                </div>
                <div>
                  <label className="label">أعمدة البريد الإلكتروني</label>
                  <input className="input" placeholder="مثال: Email" value={listToText(survey.metadata.emailCol)}
                    onChange={e => updateMetaList("emailCol", e.target.value)} />
                </div>
                <div>
                  <label className="label">أعمدة الوظيفة / الدرجة</label>
                  <input className="input" placeholder="مثال: الوظيفة، الدرجة العلمية" value={listToText(survey.metadata.degreeCol)}
                    onChange={e => updateMetaList("degreeCol", e.target.value)} />
                </div>
                <div>
                  <label className="label">أعمدة القسم / التخصص</label>
                  <input className="input" placeholder="مثال: القسم" value={listToText(survey.metadata.departmentCol)}
                    onChange={e => updateMetaList("departmentCol", e.target.value)} />
                </div>
                <div>
                  <label className="label">أعمدة نصية مفتوحة (ملاحظات/مقترحات)</label>
                  <input className="input" placeholder="مثال: مقترحات أخرى" value={listToText(survey.metadata.freeTextCols)}
                    onChange={e => updateMetaList("freeTextCols", e.target.value)} />
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {tab === "sections" && (
        <div>
          <div className="sections-toolbar">
            <div><div style={{ color: "#fff", fontSize: 16, fontWeight: 900 }}>هيكل الاستبيان</div><div style={{ color: "rgba(255,255,255,.42)", fontSize: 11.5, marginTop: 3 }}>قسّم الأسئلة إلى محاور واضحة؛ الإعدادات الأقل استخدامًا موجودة داخل كل سؤال.</div></div>
            <div className="sections-summary"><span className="editor-count-chip">{survey.sections.length} محاور</span><span className="editor-count-chip">{countQuestions(survey)} سؤال</span><button className="btn btn-blue btn-sm" onClick={addSection}>＋ محور جديد</button></div>
          </div>
          {survey.sections.length === 0 && <div className="editor-empty"><div className="qa-status-icon info" style={{ width: 46, height: 46, margin: "0 auto 10px" }}><QualityIcon name="table" size={22} /></div><div style={{ color: "#fff", fontWeight: 800, marginBottom: 5 }}>ابدأ بإضافة أول محور</div><div style={{ color: "rgba(255,255,255,.42)", fontSize: 12, marginBottom: 16 }}>كل محور يجمع مجموعة أسئلة تقيس جانبًا محددًا.</div><button className="btn btn-primary qa-icon-button" onClick={addSection}><QualityIcon name="plus" size={16} /> إضافة أول محور</button></div>}
          {survey.sections.map((sec, i) => (
            <SectionBlock
              key={sec.id} section={sec} index={i} total={survey.sections.length}
              onChange={s => updateSection(i, s)}
              onDelete={() => deleteSection(i)}
              onMove={dir => moveSection(i, dir)}
              allSections={survey.sections.map(s => ({ id: s.id, name: s.name }))}
              onMoveQuestionToSection={(qIdx, toSectionId) => moveQuestionToSection(i, qIdx, toSectionId)}
            />
          ))}
          {survey.sections.length > 0 && <button className="add-section-card" onClick={addSection}><span style={{ fontSize: 22 }}>＋</span><span>إضافة محور آخر</span></button>}
        </div>
      )}

      {tab === "report" && <ReportTextsTab survey={survey} onChange={setSurvey} />}

      {tab === "preview" && <PreviewTab survey={survey} />}

      {/* Step navigation: Next after each step, Save (distinct from the top one) on the last */}
      <div className="editor-footer">
        <div>
          {stepIdx > 0 && <button className="btn btn-ghost btn-sm" onClick={goBack}>→ الخطوة السابقة</button>}
        </div>
        <span style={{ color: "rgba(255,255,255,.35)", fontSize: 10.5 }}>الخطوة {stepIdx + 1} من {EDITOR_STEPS.length}</span>
        <div>
          {stepIdx < EDITOR_STEPS.length - 1 ? (
            <button
              className="btn btn-blue btn-sm"
              disabled={stepIdx === 0 && !canSave}
              title={stepIdx === 0 && !canSave ? "أدخل اسم الاستبيان للمتابعة" : undefined}
              onClick={goNext}
            >
              التالي: {EDITOR_STEPS[stepIdx + 1].label} ←
            </button>
          ) : (
            <button className="btn btn-primary" disabled={!canSave} onClick={() => onSave(survey)}>
              💾 حفظ الاستبيان
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// ── Root component ────────────────────────────────────────────────────────────
export default function SurveyManagement() {
  const [surveys, setSurveys] = useState(loadCustomSurveys);
  const [editing, setEditing] = useState(null); // survey object being edited, or null = list view
  const [editingBuiltInTexts, setEditingBuiltInTexts] = useState(null);
  const [showQuickCreate, setShowQuickCreate] = useState(false);
  const builtInSurveys = getBuiltInSurveyCatalog();

  const refresh = () => setSurveys(loadCustomSurveys());

  const handleCreate = () => setShowQuickCreate(true);
  const handleAdvancedCreate = () => { setShowQuickCreate(false); setEditing(createSurvey()); };
  const handleQuickCreate = ({ name, surveyType, scaleType }) => {
    setShowQuickCreate(false);
    setEditing(createSurvey({
      name, surveyType, scaleType,
      sections: [createSection({ name: "المحور الأول", questions: [createQuestion()] })],
      _initialTab: "sections",
    }));
  };
  const handleEdit = (survey) => setEditing(survey);
  const handleSave = (survey) => { saveCustomSurvey(survey); refresh(); setEditing(null); };
  const handleCancel = () => setEditing(null);
  const handleDelete = (id) => { deleteCustomSurvey(id); refresh(); };
  const handleDuplicate = (id) => { duplicateCustomSurvey(id); refresh(); };
  const handleToggleStatus = (id, status) => { setCustomSurveyStatus(id, status); refresh(); };
  const handleCloneBuiltIn = (id) => {
    const survey = createCustomSurveyFromBuiltIn(id);
    if (survey) setEditing({ ...survey, _initialTab: "general" });
  };
  const handleSaveBuiltInTexts = (id, texts) => {
    saveBuiltInReportTexts(id, texts);
    setEditingBuiltInTexts(null);
  };

  const handleDownloadTemplate = () => downloadBlob(buildSurveyTemplateBlob(), "قالب_استبيان.xlsx");

  const handleImportTemplate = async (file) => {
    const buf = await file.arrayBuffer();
    const survey = importSurveyFromTemplateArrayBuffer(buf);
    if (!survey.sections.length) throw new Error("لم يتم العثور على أسئلة في تبويب \"المحاور والأسئلة\" بالملف.");
    saveCustomSurvey(survey);
    refresh();
  };

  const handleExportBackup = () => downloadBlob(buildSurveysBackupBlob(), `نسخة_احتياطية_الاستبيانات_${new Date().toISOString().slice(0, 10)}.json`);

  const handleImportBackup = (text) => {
    const count = importSurveysBackup(text);
    refresh();
    return count;
  };

  if (editing) {
    return <SurveyEditorView survey={editing} onSave={handleSave} onCancel={handleCancel} />;
  }
  if (editingBuiltInTexts) {
    return <BuiltInReportTextsEditor survey={editingBuiltInTexts} onSave={handleSaveBuiltInTexts} onCancel={() => setEditingBuiltInTexts(null)} />;
  }

  return (
    <>
      <SurveyListView
        surveys={surveys}
        builtInSurveys={builtInSurveys}
        onCreate={handleCreate}
        onEdit={handleEdit}
        onDelete={handleDelete}
        onDuplicate={handleDuplicate}
        onToggleStatus={handleToggleStatus}
        onCloneBuiltIn={handleCloneBuiltIn}
        onEditReportTexts={setEditingBuiltInTexts}
        onDownloadTemplate={handleDownloadTemplate}
        onImportTemplate={handleImportTemplate}
        onExportBackup={handleExportBackup}
        onImportBackup={handleImportBackup}
      />
      {showQuickCreate && <QuickCreateDialog onCreate={handleQuickCreate} onAdvanced={handleAdvancedCreate} onCancel={() => setShowQuickCreate(false)} />}
    </>
  );
}
