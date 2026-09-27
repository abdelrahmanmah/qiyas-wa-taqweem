import { QualityIcon } from "./UiElements.jsx";
import { ANALYSIS_MODES, ANALYSIS_STAGES, analysisModeLabel, analysisSourceLabel } from "./analysisWorkflow.js";

const STAGES = [
  { id: ANALYSIS_STAGES.SELECTING, label: "اختيار البيانات", icon: "upload" },
  { id: ANALYSIS_STAGES.REVIEWING, label: "مراجعة الإعدادات", icon: "report" },
  { id: ANALYSIS_STAGES.CLEANING, label: "معالجة البيانات", icon: "table" },
  { id: ANALYSIS_STAGES.PROCESSING, label: "تشغيل التحليل", icon: "analytics" },
  { id: ANALYSIS_STAGES.RESULTS, label: "مركز النتائج", icon: "check" },
];

export function AnalysisJourneyHeader({ stage, request, nextLabel }) {
  const current = Math.max(0, STAGES.findIndex(item => item.id === stage));
  return (
    <div className="analysis-journey-shell">
      <div className="analysis-context-bar">
        <div>
          <span className="analysis-context-kicker">رحلة التحليل الموحدة</span>
          <strong>{analysisModeLabel(request?.mode, request?.items?.length || 0)}</strong>
        </div>
        <div className="analysis-context-meta">
          <span>المصدر: <b>{analysisSourceLabel(request?.source)}</b></span>
          <span>العناصر: <b>{request?.items?.length || 0}</b></span>
          {nextLabel && <span>التالي: <b>{nextLabel}</b></span>}
        </div>
      </div>
      <div className="analysis-stage-track" role="list" aria-label="مراحل التحليل">
        {STAGES.map((item, index) => {
          const state = index === current ? "active" : index < current ? "done" : "pending";
          return <div key={item.id} role="listitem" className={`analysis-stage ${state}`} aria-current={state === "active" ? "step" : undefined}>
            <span className="analysis-stage-icon"><QualityIcon name={state === "done" ? "check" : item.icon} size={17} /></span>
            <span>{item.label}</span>
          </div>;
        })}
      </div>
    </div>
  );
}

export function AnalysisStart({ onFiles, onComparison, onSemester }) {
  const choices = [
    { id: "files", icon: "upload", title: "ملف أو مجموعة ملفات", copy: "اختر من الجهاز أو Google Drive. يحدد النظام تلقائيًا إن كان التحليل فرديًا أو جماعيًا.", action: onFiles },
    { id: "comparison", icon: "table", title: "مقارنة سنوات", copy: "قارن نتائج النوع نفسه خلال سنتين أو ثلاث سنوات دراسية.", action: onComparison },
    { id: "semester", icon: "calendar", title: "استبيانات فصل دراسي", copy: "اختر النماذج حسب السنة والفصل ثم راجع إعدادات التقارير مرة واحدة.", action: onSemester },
  ];
  return <div className="analysis-start-card card">
    <div className="analysis-start-heading"><span>ابدأ من نوع البيانات</span><h1>ماذا تريد أن تحلّل؟</h1><p>اختر نقطة البداية فقط، وسنقودك في رحلة واحدة حتى مركز النتائج.</p></div>
    <div className="analysis-start-grid">
      {choices.map(choice => <button key={choice.id} type="button" className="analysis-start-option" onClick={choice.action}>
        <span className="analysis-start-icon"><QualityIcon name={choice.icon} size={25} /></span>
        <strong>{choice.title}</strong><small>{choice.copy}</small><i aria-hidden="true">←</i>
      </button>)}
    </div>
  </div>;
}

export function AnalysisReview({ request, meta, onMetaChange, issues = [], onBack, onContinue, busy = false, children }) {
  const itemNames = request?.items?.map(item => item.name || item.file?.name).filter(Boolean) || [];
  return <div className="card analysis-review-card">
    <div className="analysis-review-head">
      <div><span>مراجعة مختصرة</span><h2>تأكد من البيانات قبل بدء التحليل</h2><p>يمكنك تعديل بيانات التقرير هنا دون الرجوع إلى بداية الرحلة.</p></div>
      <span className="analysis-review-count">{itemNames.length} {itemNames.length === 1 ? "استبيان" : "استبيانات"}</span>
    </div>
    <div className="analysis-review-layout">
      <section>
        <h3>البيانات المحددة</h3>
        <div className="analysis-review-items">
          {itemNames.map((name, index) => <div key={`${name}-${index}`}><QualityIcon name="report" size={15} /><span>{name}</span></div>)}
        </div>
      </section>
      <section>
        <h3>بيانات التقرير</h3>
        <div className="analysis-review-fields">
          <label><span>العام الأكاديمي</span><input className="input" value={meta.year || ""} onChange={e => onMetaChange({ ...meta, year: e.target.value })} /></label>
          <label><span>البرنامج / القسم</span><input className="input" value={meta.program || ""} onChange={e => onMetaChange({ ...meta, program: e.target.value })} /></label>
          <label><span>مُعدّ التقرير *</span><input className="input" value={meta.preparedBy || ""} onChange={e => onMetaChange({ ...meta, preparedBy: e.target.value })} /></label>
          <label><span>مراجع التقرير *</span><input className="input" value={meta.reviewer || ""} onChange={e => onMetaChange({ ...meta, reviewer: e.target.value })} /></label>
        </div>
        {children}
      </section>
    </div>
    {!!issues.length && <div className="analysis-review-issues" role="alert">{issues.map(issue => <div key={issue}><QualityIcon name="warning" size={15} />{issue}</div>)}</div>}
    <div className="analysis-review-actions"><button type="button" className="btn btn-ghost" onClick={onBack}>→ رجوع</button><button type="button" className="btn btn-primary" disabled={busy} onClick={onContinue}>{busy ? "جاري التحضير…" : request?.mode === ANALYSIS_MODES.COMPARISON ? "إنشاء المقارنة ←" : "متابعة التحليل ←"}</button></div>
  </div>;
}
