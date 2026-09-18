import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { GoogleDriveIcon, QualityIcon } from "./UiElements.jsx";
import { SEMESTERS } from "./engine/semesterSurveyModel.js";
import { DRIVE_LIBRARY_CSS } from "./DriveLibrary.jsx";

function Icon({ name, size = 20 }) {
  return <QualityIcon name={name} size={size} />;
}

function normalize(value) {
  return String(value ?? "").toLowerCase().replace(/[أإآ]/g, "ا").replace(/ى/g, "ي").replace(/ة/g, "ه").trim();
}

function currentAcademicYear() {
  const now = new Date();
  const start = now.getMonth() >= 8 ? now.getFullYear() : now.getFullYear() - 1;
  return `${start}-${start + 1}`;
}

function currentSemester() {
  const month = new Date().getMonth() + 1;
  if (month >= 9 || month === 1) return "خريف";
  if (month >= 2 && month <= 6) return "ربيع";
  return "صيف";
}

function academicYearOptions() {
  const currentStart = Number(currentAcademicYear().slice(0, 4));
  return Array.from({ length: 8 }, (_, index) => `${currentStart - index}-${currentStart - index + 1}`);
}

function formatDate(value) {
  const date = new Date(value || 0);
  if (!value || Number.isNaN(date.getTime())) return "تاريخ غير متاح";
  return new Intl.DateTimeFormat("ar-EG", { day: "numeric", month: "short", year: "numeric" }).format(date);
}

export default function SurveySummary({ connected, connecting, onConnect, onLoad, onOpenFile, onAnalyzeSurvey, schemas }) {
  const [year, setYear] = useState(currentAcademicYear);
  const [semester, setSemester] = useState(currentSemester);
  const [summary, setSummary] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [progress, setProgress] = useState({ stage: "idle", current: 0, total: 0 });
  const [status, setStatus] = useState("all");
  const [surveyType, setSurveyType] = useState("all");
  const [query, setQuery] = useState("");
  const cacheRef = useRef(new Map());
  const requestRef = useRef(0);
  const years = useMemo(academicYearOptions, []);

  const loadSummary = useCallback(async (force = false) => {
    if (!connected || !year || !semester) return;
    const key = `${year}::${semester}`;
    const requestId = ++requestRef.current;
    setError("");
    setStatus("all");
    setSurveyType("all");
    setQuery("");
    if (!force && cacheRef.current.has(key)) {
      setSummary(cacheRef.current.get(key));
      setProgress({ stage: "done", current: 0, total: 0 });
      setLoading(false);
      return;
    }
    setLoading(true);
    setSummary({ year, semester, items: [] });
    setProgress({ stage: "scan", current: 0, total: 0 });
    try {
      const result = await onLoad(year, semester, next => {
        if (requestId !== requestRef.current) return;
        setProgress(next);
        if (next.items) setSummary({ year, semester, items: next.items });
      });
      if (requestId !== requestRef.current) return;
      cacheRef.current.set(key, result);
      setSummary(result);
      setProgress({ stage: "done", current: result.items.length, total: result.items.length });
    } catch (loadError) {
      if (requestId === requestRef.current) setError(loadError?.message || "تعذر تجهيز ملخص الاستبيانات.");
    } finally {
      if (requestId === requestRef.current) setLoading(false);
    }
  }, [connected, onLoad, semester, year]);

  useEffect(() => {
    if (connected) loadSummary();
  }, [connected, loadSummary]);

  const items = summary?.year === year && summary?.semester === semester ? (summary.items || []) : [];
  const stats = useMemo(() => ({
    total: items.length,
    analyzed: items.filter(item => item.analyzed).length,
    pending: items.filter(item => !item.analyzed).length,
    responses: items.reduce((sum, item) => sum + (Number.isFinite(item.responses) ? item.responses : 0), 0),
  }), [items]);
  const filteredItems = useMemo(() => {
    const needle = normalize(query);
    return items.filter(item => {
      if (status === "analyzed" && !item.analyzed) return false;
      if (status === "pending" && item.analyzed) return false;
      if (surveyType !== "all" && item.typeId !== surveyType) return false;
      const typeLabel = item.typeId ? (schemas?.[item.typeId]?.label || item.typeId) : "";
      return !needle || normalize([item.file?.name, item.program, typeLabel].join(" ")).includes(needle);
    });
  }, [items, query, schemas, status, surveyType]);

  const availableSurveyTypes = useMemo(() => {
    const ids = new Set(items.map(item => item.typeId).filter(Boolean));
    return [...ids]
      .map(id => ({ id, label: schemas?.[id]?.label || id }))
      .sort((a, b) => a.label.localeCompare(b.label, "ar"));
  }, [items, schemas]);

  if (!connected) {
    return (
      <section className="drive-library drive-library-disconnected">
        <style>{DRIVE_LIBRARY_CSS}</style>
        <div className="drive-library-connect-card">
          <span className="drive-library-drive-mark"><GoogleDriveIcon size={50} /></span>
          <span className="drive-library-kicker">ملخص الاستبيانات</span>
          <h1>اعرف حالة كل استبيان في ثوانٍ</h1>
          <p>اربط Google Drive، ثم اختر السنة والفصل لعرض ما تم تحليله وما ينتظر التحليل وعدد المشاركين.</p>
          <button className="drive-library-primary" type="button" onClick={onConnect} disabled={connecting}>
            <GoogleDriveIcon size={21} /> {connecting ? "جاري الاتصال…" : "ربط Google Drive"}
          </button>
        </div>
      </section>
    );
  }

  return (
    <section className="drive-library" aria-label="ملخص الاستبيانات">
      <style>{DRIVE_LIBRARY_CSS}</style>
      <header className="drive-library-hero">
        <div className="drive-library-hero-copy">
          <span className="drive-library-kicker"><Icon name="dashboard" size={18} /> متابعة سريعة</span>
          <h1>ملخص الاستبيانات</h1>
          <p>اختر السنة والفصل؛ وسيظهر تلقائياً ما تم تحليله وما لم يُحلل وعدد المشاركين في كل استبيان.</p>
        </div>
      </header>

      <section className="drive-quick-summary" style={{ marginTop: 16 }}>
        <div className="drive-quick-controls">
          <label><span>السنة الدراسية</span><select value={year} onChange={event => setYear(event.target.value)}>{years.map(value => <option key={value}>{value}</option>)}</select></label>
          <label><span>الفصل الدراسي</span><select value={semester} onChange={event => setSemester(event.target.value)}>{SEMESTERS.map(value => <option key={value}>{value}</option>)}</select></label>
          <button type="button" onClick={() => loadSummary(true)} disabled={loading}>
            <span className={loading ? "drive-library-spin" : ""}><Icon name="refresh" size={16} /></span> تحديث الملخص
          </button>
        </div>

        {loading && (
          <div className="drive-quick-progress" role="status">
            <div><span>{progress.stage === "scan" ? `جاري حصر استبيانات فصل ${semester}…` : `جاري حساب أعداد المشاركين (${progress.current} من ${progress.total})`}</span>{progress.total > 0 && <b>{Math.round((progress.current / progress.total) * 100)}%</b>}</div>
            <i><span className={progress.stage === "scan" ? "indeterminate" : ""} style={progress.total > 0 ? { width: `${Math.round((progress.current / progress.total) * 100)}%` } : undefined} /></i>
          </div>
        )}
        {error && <div className="drive-library-error" role="alert"><span>!</span><p>{error}</p><button type="button" onClick={() => loadSummary(true)}>إعادة المحاولة</button></div>}

        {!error && (
          <>
            <div className="drive-quick-stats">
              <div><span className="blue"><Icon name="folder" size={18} /></span><b>{stats.total.toLocaleString("ar-EG")}</b><small>إجمالي الاستبيانات</small></div>
              <div><span className="green"><Icon name="check" size={18} /></span><b>{stats.analyzed.toLocaleString("ar-EG")}</b><small>تم تحليلها</small></div>
              <div><span className="amber"><Icon name="clock" size={18} /></span><b>{stats.pending.toLocaleString("ar-EG")}</b><small>لم تُحلل بعد</small></div>
              <div><span className="purple"><Icon name="users" size={18} /></span><b>{stats.responses.toLocaleString("ar-EG")}</b><small>{loading ? "مشارك محسوب حتى الآن" : "إجمالي المشاركين"}</small></div>
            </div>

            <div className="drive-quick-toolbar">
              <div className="drive-quick-status-tabs">
                <button className={status === "all" ? "active" : ""} onClick={() => setStatus("all")}>الكل <b>{stats.total}</b></button>
                <button className={status === "analyzed" ? "active" : ""} onClick={() => setStatus("analyzed")}>تم التحليل <b>{stats.analyzed}</b></button>
                <button className={status === "pending" ? "active" : ""} onClick={() => setStatus("pending")}>لم يُحلل <b>{stats.pending}</b></button>
              </div>
              <div className="drive-quick-filter-tools">
                <label className="drive-quick-type-filter">
                  <Icon name="filter" size={15} />
                  <span className="sr-only">فلترة حسب نوع الاستبيان</span>
                  <select value={surveyType} onChange={event => setSurveyType(event.target.value)} aria-label="فلترة حسب نوع الاستبيان">
                    <option value="all">كل أنواع الاستبيانات</option>
                    {availableSurveyTypes.map(type => <option key={type.id} value={type.id}>{type.label}</option>)}
                  </select>
                </label>
                <label className="drive-quick-search"><Icon name="search" size={16} /><input value={query} onChange={event => setQuery(event.target.value)} placeholder="ابحث باسم الاستبيان أو البرنامج…" /></label>
              </div>
            </div>

            {!loading && items.length === 0 ? (
              <div className="drive-quick-empty"><Icon name="folder" size={25} /><b>لا توجد استبيانات لهذا الفصل</b><span>تأكد أن أسماء ملفات الردود تحتوي على {semester} و{year}.</span></div>
            ) : filteredItems.length === 0 && items.length > 0 ? (
              <div className="drive-quick-empty"><Icon name="search" size={25} /><b>لا توجد نتائج مطابقة</b><span>غيّر حالة التحليل أو كلمة البحث.</span></div>
            ) : (
              <div className="drive-quick-list">
                <div className="drive-quick-list-head"><span>الحالة</span><span>الاستبيان</span><span>النوع / البرنامج</span><span>المشاركون</span><span>الإجراء</span></div>
                {filteredItems.map(item => {
                  const typeLabel = item.typeId ? (schemas?.[item.typeId]?.label || item.typeId) : "نوع غير محدد";
                  return (
                    <div className="drive-quick-row" key={item.file.id}>
                      <span><i className={`drive-quick-state ${item.analyzed ? "done" : "pending"}`}><Icon name={item.analyzed ? "check" : "clock"} size={13} />{item.analyzed ? "تم التحليل" : "لم يُحلل"}</i></span>
                      <span className="drive-quick-name"><b title={item.file.name}>{item.file.name}</b><small>{formatDate(item.file.modifiedTime)}</small></span>
                      <span className="drive-quick-context"><b>{typeLabel}</b><small>{item.program || "كلية / برنامج غير محدد"}</small></span>
                      <span className="drive-quick-responses">{Number.isFinite(item.responses) ? <><b>{item.responses.toLocaleString("ar-EG")}</b><small>مشارك</small></> : <em title={item.responseError}>{loading ? "…" : "—"}</em>}</span>
                      <span className="drive-quick-actions">
                        {item.analyzed
                          ? <button className="report" type="button" onClick={() => onOpenFile(item.report)}><Icon name="open" size={14} /> فتح التقرير</button>
                          : <button className="analyze" type="button" onClick={() => onAnalyzeSurvey(item.file)}><Icon name="chart" size={14} /> تحليل الآن</button>}
                        <button className="source" type="button" onClick={() => onOpenFile(item.file)}>فتح الردود</button>
                      </span>
                    </div>
                  );
                })}
              </div>
            )}
          </>
        )}
      </section>
    </section>
  );
}
