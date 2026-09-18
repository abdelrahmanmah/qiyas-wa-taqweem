import { useMemo, useState } from "react";
import { GoogleDriveIcon, QualityIcon } from "./UiElements.jsx";

const SHEET_MIMES = new Set([
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "application/vnd.google-apps.spreadsheet",
  "text/csv",
]);

const REPORT_MIMES = new Set([
  "application/pdf",
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.google-apps.document",
]);

const FORM_MIME = "application/vnd.google-apps.form";

function Icon({ name, size = 20 }) {
  return <QualityIcon name={name} size={size} />;
}

function getFileKind(file) {
  if (SHEET_MIMES.has(file.mimeType) || file.mimeType === FORM_MIME) return "survey";
  if (REPORT_MIMES.has(file.mimeType)) return "report";
  return "other";
}

function getFormat(file) {
  const formats = {
    "application/vnd.google-apps.spreadsheet": "Google Sheets",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "Excel",
    "application/vnd.ms-excel": "Excel",
    "text/csv": "CSV",
    "application/vnd.google-apps.form": "Google Form",
    "application/pdf": "PDF",
    "application/vnd.google-apps.document": "Google Docs",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "Word",
    "application/msword": "Word",
  };
  return formats[file.mimeType] ?? "ملف";
}

function fileAccent(file) {
  const format = getFormat(file);
  if (format === "PDF") return { color: "#fb7185", bg: "rgba(251,113,133,.12)" };
  if (format === "Word" || format === "Google Docs") return { color: "#60a5fa", bg: "rgba(96,165,250,.12)" };
  if (format === "Google Form") return { color: "#c084fc", bg: "rgba(192,132,252,.12)" };
  return { color: "#34d399", bg: "rgba(52,211,153,.12)" };
}

function formatDate(value) {
  if (!value) return "تاريخ غير متاح";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "تاريخ غير متاح";
  return new Intl.DateTimeFormat("ar-EG", { day: "numeric", month: "short", year: "numeric" }).format(date);
}

function normalize(value) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[أإآ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .trim();
}

function StatCard({ label, value, icon, tone }) {
  return (
    <div className="drive-library-stat" style={{ "--stat-tone": tone }}>
      <span className="drive-library-stat-icon"><Icon name={icon} size={21} /></span>
      <span><b>{value.toLocaleString("ar-EG")}</b><small>{label}</small></span>
    </div>
  );
}

export default function DriveLibrary({
  connected,
  connecting,
  account,
  files,
  loading,
  loadingMore,
  hasMore,
  error,
  onConnect,
  onRefresh,
  onLoadMore,
  onOpenFile,
  onAnalyzeSurvey,
  detectYear,
  detectProgram,
  detectType,
  schemas,
}) {
  const [query, setQuery] = useState("");
  const [kind, setKind] = useState("all");
  const [year, setYear] = useState("");
  const [program, setProgram] = useState("");
  const [surveyType, setSurveyType] = useState("");
  const [format, setFormat] = useState("");
  const [sort, setSort] = useState("newest");
  const [view, setView] = useState("grid");

  const enriched = useMemo(() => files.map(file => {
    const fileKind = getFileKind(file);
    const typeId = fileKind === "survey" ? detectType?.(file.name) : null;
    return {
      ...file,
      kind: fileKind,
      format: getFormat(file),
      year: detectYear?.(file.name) ?? "",
      program: detectProgram?.(file.name) ?? "",
      typeId: typeId ?? "",
      typeLabel: typeId ? (schemas?.[typeId]?.label ?? typeId) : "",
      analyzable: SHEET_MIMES.has(file.mimeType),
    };
  }), [files, detectYear, detectProgram, detectType, schemas]);

  const options = useMemo(() => ({
    years: [...new Set(enriched.map(file => file.year).filter(Boolean))].sort().reverse(),
    programs: [...new Set(enriched.map(file => file.program).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ar")),
    surveyTypes: [...new Map(enriched.filter(file => file.typeId).map(file => [file.typeId, file.typeLabel])).entries()],
    formats: [...new Set(enriched.map(file => file.format))].sort((a, b) => a.localeCompare(b, "ar")),
  }), [enriched]);


  const filtered = useMemo(() => {
    const needle = normalize(query);
    return enriched
      .filter(file => kind === "all" || file.kind === kind)
      .filter(file => !year || file.year === year)
      .filter(file => !program || file.program === program)
      .filter(file => !surveyType || file.typeId === surveyType)
      .filter(file => !format || file.format === format)
      .filter(file => !needle || normalize([file.name, file.year, file.program, file.typeLabel, file.format].join(" ")).includes(needle))
      .sort((a, b) => {
        if (sort === "name") return a.name.localeCompare(b.name, "ar");
        const diff = new Date(b.modifiedTime || 0) - new Date(a.modifiedTime || 0);
        return sort === "oldest" ? -diff : diff;
      });
  }, [enriched, query, kind, year, program, surveyType, format, sort]);

  const counts = useMemo(() => {
    const recentLimit = Date.now() - (30 * 24 * 60 * 60 * 1000);
    return {
      total: enriched.length,
      surveys: enriched.filter(file => file.kind === "survey").length,
      reports: enriched.filter(file => file.kind === "report").length,
      recent: enriched.filter(file => new Date(file.modifiedTime || 0).getTime() >= recentLimit).length,
    };
  }, [enriched]);


  const hasFilters = Boolean(query || kind !== "all" || year || program || surveyType || format || sort !== "newest");
  const clearFilters = () => {
    setQuery(""); setKind("all"); setYear(""); setProgram(""); setSurveyType(""); setFormat(""); setSort("newest");
  };

  if (!connected) {
    return (
      <section className="drive-library drive-library-disconnected">
        <style>{DRIVE_LIBRARY_CSS}</style>
        <div className="drive-library-connect-card">
          <span className="drive-library-drive-mark"><GoogleDriveIcon size={50} /></span>
          <span className="drive-library-kicker">مكتبة الملفات الذكية</span>
          <h1>كل تقاريرك واستبياناتك في مكان واحد</h1>
          <p>اربط حساب Google Drive لعرض الملفات، معرفة أعداد التقارير والاستبيانات، والوصول إليها بفلاتر أسرع.</p>
          <button className="drive-library-primary" type="button" onClick={onConnect} disabled={connecting}>
            <GoogleDriveIcon size={21} />
            {connecting ? "جاري الاتصال…" : "ربط Google Drive"}
          </button>
          <small>يُستخدم الاتصال للعرض والوصول إلى ملفاتك فقط داخل هذه المساحة.</small>
          {error && <div className="drive-library-error" role="alert"><span>!</span><p>{error}</p></div>}
        </div>
      </section>
    );
  }

  return (
    <section className="drive-library" aria-label="مكتبة Google Drive">
      <style>{DRIVE_LIBRARY_CSS}</style>

      <header className="drive-library-hero">
        <div className="drive-library-hero-copy">
          <span className="drive-library-kicker"><GoogleDriveIcon size={19} /> مساحة Drive المتصلة</span>
          <h1>مكتبة التقارير والاستبيانات</h1>
          <p>ابحث، صفِّ، وافتح ملفات الجودة من واجهة واحدة منظمة وسريعة.</p>
          {account && <span className="drive-library-account"><i /> {account}</span>}
        </div>
        <div className="drive-library-hero-actions">
          <button className="drive-library-refresh" type="button" onClick={onRefresh} disabled={loading || loadingMore}>
            <span className={loading ? "drive-library-spin" : ""}><Icon name="refresh" size={18} /></span>
            {loading ? "جاري التحديث" : "تحديث المكتبة"}
          </button>
        </div>
      </header>

      {error && (
        <div className="drive-library-error" role="alert">
          <span>!</span><p>{error}</p><button type="button" onClick={onRefresh}>إعادة المحاولة</button>
        </div>
      )}

      <div className="drive-library-stats" aria-label="ملخص الملفات">
        <StatCard label={hasMore ? "الملفات المحملة" : "إجمالي الملفات"} value={counts.total} icon="folder" tone="#60a5fa" />
        <StatCard label={hasMore ? "استبيانات محملة" : "الاستبيانات"} value={counts.surveys} icon="chart" tone="#34d399" />
        <StatCard label={hasMore ? "تقارير محملة" : "التقارير"} value={counts.reports} icon="file" tone="#fb7185" />
        <StatCard label="معدلة آخر 30 يومًا" value={counts.recent} icon="refresh" tone="#c084fc" />
      </div>

      <div className="drive-library-controls">
        <div className="drive-library-search-row">
          <label className="drive-library-search">
            <Icon name="search" size={19} />
            <input value={query} onChange={event => setQuery(event.target.value)} placeholder="ابحث باسم الملف، البرنامج، السنة أو نوع الاستبيان…" />
            {query && <button type="button" onClick={() => setQuery("")} aria-label="مسح البحث"><Icon name="close" size={16} /></button>}
          </label>
          <div className="drive-library-kind-tabs" aria-label="تصنيف الملفات">
            <button className={kind === "all" ? "active" : ""} onClick={() => setKind("all")}>الكل <b>{counts.total}</b></button>
            <button className={kind === "survey" ? "active" : ""} onClick={() => setKind("survey")}>استبيانات <b>{counts.surveys}</b></button>
            <button className={kind === "report" ? "active" : ""} onClick={() => setKind("report")}>تقارير <b>{counts.reports}</b></button>
          </div>
        </div>

        <div className="drive-library-filter-row">
          <select value={surveyType} onChange={event => setSurveyType(event.target.value)} aria-label="نوع الاستبيان">
            <option value="">كل أنواع الاستبيانات</option>
            {options.surveyTypes.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
          </select>
          <select value={year} onChange={event => setYear(event.target.value)} aria-label="السنة الدراسية">
            <option value="">كل السنوات</option>
            {options.years.map(value => <option key={value} value={value}>{value}</option>)}
          </select>
          <select value={program} onChange={event => setProgram(event.target.value)} aria-label="البرنامج">
            <option value="">كل البرامج</option>
            {options.programs.map(value => <option key={value} value={value}>{value}</option>)}
          </select>
          <select value={format} onChange={event => setFormat(event.target.value)} aria-label="صيغة الملف">
            <option value="">كل الصيغ</option>
            {options.formats.map(value => <option key={value} value={value}>{value}</option>)}
          </select>
          <select value={sort} onChange={event => setSort(event.target.value)} aria-label="ترتيب الملفات">
            <option value="newest">الأحدث أولًا</option>
            <option value="oldest">الأقدم أولًا</option>
            <option value="name">الاسم أبجديًا</option>
          </select>
          {hasFilters && <button type="button" className="drive-library-clear" onClick={clearFilters}>مسح الفلاتر</button>}
        </div>
      </div>

      <div className="drive-library-results-head">
        <div><b>{filtered.length.toLocaleString("ar-EG")}</b> ملف مطابق <small>من {counts.total.toLocaleString("ar-EG")} ملف محمل</small></div>
        <div className="drive-library-view-toggle" aria-label="طريقة العرض">
          <button className={view === "grid" ? "active" : ""} onClick={() => setView("grid")} aria-label="عرض شبكي"><Icon name="grid" size={17} /></button>
          <button className={view === "list" ? "active" : ""} onClick={() => setView("list")} aria-label="عرض قائمة"><Icon name="list" size={18} /></button>
        </div>
      </div>

      {loading && enriched.length === 0 ? (
        <div className="drive-library-loading" role="status">
          <span className="drive-library-spinner" />
          <b>نجمع ملفات Drive ونرتبها لك…</b>
          <small>قد يستغرق ذلك لحظات حسب عدد الملفات.</small>
        </div>
      ) : filtered.length === 0 ? (
        <div className="drive-library-empty">
          <span><Icon name="search" size={28} /></span>
          <h2>{enriched.length ? "لا توجد ملفات تطابق هذه الفلاتر" : "لم نجد تقارير أو استبيانات بعد"}</h2>
          <p>{enriched.length ? "جرّب مسح بعض الفلاتر أو استخدم كلمة بحث مختلفة." : "اضغط تحديث المكتبة بعد إضافة الملفات إلى Google Drive."}</p>
          <div className="drive-library-empty-actions">
            {hasFilters && <button type="button" onClick={clearFilters}>عرض كل الملفات المحملة</button>}
            {hasMore && <button type="button" onClick={onLoadMore} disabled={loadingMore}>{loadingMore ? "جاري التحميل…" : "البحث في ملفات أكثر"}</button>}
          </div>
        </div>
      ) : (
        <>
          <div className={`drive-library-files ${view === "list" ? "list" : "grid"}`}>
            {filtered.map(file => {
              const accent = fileAccent(file);
              return (
                <article className="drive-library-file" key={file.id} style={{ "--file-color": accent.color, "--file-bg": accent.bg }}>
                  <div className="drive-library-file-top">
                    <span className="drive-library-file-icon"><Icon name={file.kind === "survey" ? "chart" : "file"} size={24} /></span>
                    <span className={`drive-library-kind ${file.kind}`}>{file.kind === "survey" ? "استبيان" : "تقرير"}</span>
                  </div>
                  <div className="drive-library-file-copy">
                    <h2 title={file.name}>{file.name}</h2>
                    <p>آخر تعديل: {formatDate(file.modifiedTime)}</p>
                    <div className="drive-library-file-tags">
                      <span>{file.format}</span>
                      {file.year && <span>{file.year}</span>}
                      {file.program && <span>{file.program}</span>}
                      {file.typeLabel && <span title={file.typeLabel}>{file.typeLabel}</span>}
                    </div>
                  </div>
                  <div className="drive-library-file-actions">
                    <button type="button" className="drive-library-open" onClick={() => onOpenFile(file)}><Icon name="open" size={16} /> فتح في Drive</button>
                    {file.analyzable && <button type="button" className="drive-library-analyze" onClick={() => onAnalyzeSurvey(file)}><Icon name="chart" size={16} /> تحليل الاستبيان</button>}
                  </div>
                </article>
              );
            })}
          </div>
          <div className="drive-library-pagination" aria-live="polite">
            {hasMore ? (
              <button type="button" onClick={onLoadMore} disabled={loadingMore || loading}>
                {loadingMore ? <><span className="drive-library-mini-spinner" /> جاري تحميل الدفعة التالية…</> : <>رؤية المزيد <small>تحميل الدفعة التالية فقط</small></>}
              </button>
            ) : (
              <span>تم عرض كل الملفات المتاحة</span>
            )}
          </div>
        </>
      )}
    </section>
  );
}

export const DRIVE_LIBRARY_CSS = `
.drive-library{direction:rtl;color:#e8f0fe;text-align:right;animation:driveLibraryIn .34s ease both}
@keyframes driveLibraryIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
.drive-library-hero{position:relative;overflow:hidden;display:flex;align-items:flex-end;justify-content:space-between;gap:24px;padding:30px;border:1px solid rgba(96,165,250,.18);border-radius:24px;background:radial-gradient(circle at 12% 0,rgba(96,165,250,.18),transparent 34%),linear-gradient(135deg,rgba(20,53,82,.94),rgba(9,31,49,.96));box-shadow:0 18px 55px rgba(2,12,25,.18)}
.drive-library-hero:after{content:"";position:absolute;width:240px;height:240px;border:1px solid rgba(52,211,153,.13);border-radius:50%;left:-70px;bottom:-150px;box-shadow:0 0 0 36px rgba(52,211,153,.025),0 0 0 72px rgba(52,211,153,.018)}
.drive-library-hero-copy{position:relative;z-index:1}.drive-library-kicker{display:inline-flex;align-items:center;gap:8px;color:#93c5fd;font-size:12px;font-weight:800;letter-spacing:.02em}
.drive-library-hero h1{margin:9px 0 6px;color:#fff;font-size:clamp(25px,3vw,36px);line-height:1.35}.drive-library-hero p{margin:0;color:rgba(232,240,254,.66);font-size:14px}
.drive-library-account{display:inline-flex;align-items:center;gap:7px;margin-top:14px;padding:6px 11px;border:1px solid rgba(52,211,153,.15);border-radius:999px;background:rgba(52,211,153,.07);color:#a7f3d0;font-size:11px;font-weight:700;direction:ltr}.drive-library-account i{width:7px;height:7px;border-radius:50%;background:#34d399;box-shadow:0 0 0 4px rgba(52,211,153,.12)}
.drive-library-hero-actions{position:relative;z-index:1;display:flex;align-items:center;gap:8px}.drive-library-refresh,.drive-library-primary,.drive-library-quick-button{position:relative;z-index:1;display:inline-flex;align-items:center;justify-content:center;gap:9px;border:0;border-radius:12px;font-family:'Cairo',sans-serif;font-weight:800;cursor:pointer;transition:.2s ease}
.drive-library-refresh{padding:11px 16px;color:#dbeafe;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12)}.drive-library-refresh:hover:not(:disabled){background:rgba(255,255,255,.14);transform:translateY(-1px)}.drive-library-refresh:disabled{opacity:.6;cursor:wait}
.drive-library-quick-button{padding:11px 17px;color:#06291f;background:linear-gradient(135deg,#6ee7b7,#34d399);box-shadow:0 9px 24px rgba(52,211,153,.19)}.drive-library-quick-button:hover{transform:translateY(-2px);box-shadow:0 13px 30px rgba(52,211,153,.28)}
.drive-library-spin{display:flex;animation:driveLibrarySpin 1s linear infinite}@keyframes driveLibrarySpin{to{transform:rotate(-360deg)}}
.drive-library-error{display:flex;align-items:center;gap:10px;margin:13px 0;padding:11px 13px;border:1px solid rgba(251,113,133,.2);border-radius:12px;background:rgba(251,113,133,.08);color:#fecdd3;font-size:11.5px}.drive-library-error>span{width:23px;height:23px;display:grid;place-items:center;flex:0 0 auto;border-radius:7px;background:rgba(251,113,133,.14);font-weight:900}.drive-library-error p{flex:1;margin:0}.drive-library-error button{padding:6px 10px;border:1px solid rgba(251,113,133,.22);border-radius:8px;background:rgba(251,113,133,.09);color:#fecdd3;font-family:'Cairo',sans-serif;font-size:10.5px;font-weight:800;cursor:pointer}
.drive-quick-summary{margin:16px 0;padding:20px;border:1px solid rgba(52,211,153,.17);border-radius:20px;background:linear-gradient(145deg,rgba(7,31,46,.93),rgba(8,25,42,.9));box-shadow:0 18px 45px rgba(1,10,20,.16);animation:driveLibraryIn .28s ease both}.drive-quick-head{display:flex;align-items:flex-start;justify-content:space-between;gap:20px}.drive-quick-head h2{margin:7px 0 3px;color:#fff;font-size:20px}.drive-quick-head p{margin:0;color:rgba(255,255,255,.46);font-size:11.5px}.drive-quick-close{width:34px;height:34px;display:grid;place-items:center;flex:0 0 auto;border:1px solid rgba(255,255,255,.09);border-radius:10px;background:rgba(255,255,255,.05);color:rgba(255,255,255,.55);cursor:pointer}.drive-quick-close:hover{color:#fff;background:rgba(255,255,255,.1)}
.drive-quick-controls{display:flex;align-items:flex-end;gap:9px;margin-top:16px;padding:12px;border:1px solid rgba(255,255,255,.07);border-radius:13px;background:rgba(255,255,255,.025)}.drive-quick-controls label{display:flex;flex-direction:column;gap:5px}.drive-quick-controls label span{color:rgba(255,255,255,.42);font-size:9.5px;font-weight:700}.drive-quick-controls select{height:38px;min-width:180px;padding:0 12px;border:1px solid rgba(255,255,255,.11);border-radius:9px;outline:0;background:#102940;color:#fff;font-family:'Cairo',sans-serif;font-size:12px}.drive-quick-controls>button{height:38px;display:inline-flex;align-items:center;gap:7px;padding:0 13px;border:1px solid rgba(96,165,250,.17);border-radius:9px;background:rgba(96,165,250,.09);color:#bfdbfe;font-family:'Cairo',sans-serif;font-size:11px;font-weight:800;cursor:pointer}.drive-quick-controls>button:disabled{opacity:.55;cursor:wait}
.drive-quick-progress{margin-top:11px;padding:10px 12px;border-radius:11px;background:rgba(96,165,250,.07)}.drive-quick-progress>div{display:flex;align-items:center;justify-content:space-between;gap:12px;color:#bfdbfe;font-size:10.5px}.drive-quick-progress>div b{font-size:10px}.drive-quick-progress>i{display:block;height:4px;margin-top:7px;overflow:hidden;border-radius:9px;background:rgba(255,255,255,.08)}.drive-quick-progress>i span{display:block;height:100%;border-radius:inherit;background:linear-gradient(90deg,#34d399,#60a5fa);transition:width .25s}.drive-quick-progress>i span.indeterminate{width:34%;animation:driveQuickProgress 1.2s ease-in-out infinite}@keyframes driveQuickProgress{0%{transform:translateX(220%)}100%{transform:translateX(-320%)}}
.drive-quick-stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:9px;margin-top:13px}.drive-quick-stats>div{display:grid;grid-template-columns:auto 1fr;grid-template-rows:auto auto;column-gap:9px;align-items:center;padding:12px;border:1px solid rgba(255,255,255,.07);border-radius:12px;background:rgba(255,255,255,.035)}.drive-quick-stats>div>span{grid-row:1/3;width:34px;height:34px;display:grid;place-items:center;border-radius:10px}.drive-quick-stats span.blue{color:#93c5fd;background:rgba(96,165,250,.11)}.drive-quick-stats span.green{color:#6ee7b7;background:rgba(52,211,153,.11)}.drive-quick-stats span.amber{color:#fcd34d;background:rgba(252,211,77,.1)}.drive-quick-stats span.purple{color:#d8b4fe;background:rgba(192,132,252,.1)}.drive-quick-stats b{color:#fff;font-size:18px;line-height:1.1}.drive-quick-stats small{color:rgba(255,255,255,.38);font-size:9px;white-space:nowrap}
.drive-quick-toolbar{display:flex;align-items:center;justify-content:space-between;gap:12px;margin:15px 0 9px}.drive-quick-status-tabs{display:flex;gap:5px;padding:3px;border-radius:10px;background:rgba(255,255,255,.04)}.drive-quick-status-tabs button{display:flex;align-items:center;gap:6px;padding:7px 10px;border:0;border-radius:8px;background:transparent;color:rgba(255,255,255,.48);font-family:'Cairo',sans-serif;font-size:10px;font-weight:800;cursor:pointer}.drive-quick-status-tabs button b{padding:0 5px;border-radius:99px;background:rgba(255,255,255,.06);font-size:9px}.drive-quick-status-tabs button.active{color:#fff;background:rgba(52,211,153,.12)}.drive-quick-filter-tools{display:flex;align-items:center;gap:7px;min-width:0}.drive-quick-type-filter,.drive-quick-search{height:37px;display:flex;align-items:center;gap:7px;padding:0 10px;border:1px solid rgba(255,255,255,.09);border-radius:9px;color:rgba(255,255,255,.35);background:rgba(3,17,29,.48)}.drive-quick-type-filter{min-width:190px;color:#93c5fd}.drive-quick-type-filter select{width:100%;border:0;outline:0;background:#091d30;color:#e8f0fe;font-family:'Cairo',sans-serif;font-size:10.5px;cursor:pointer}.drive-quick-search{min-width:220px}.drive-quick-search input{width:100%;border:0;outline:0;background:transparent;color:#fff;font-family:'Cairo',sans-serif;font-size:10.5px}.drive-quick-search input::placeholder{color:rgba(255,255,255,.28)}
.drive-quick-list{max-height:500px;overflow-y:auto;overflow-x:hidden;border:1px solid rgba(255,255,255,.075);border-radius:13px;scrollbar-gutter:stable}.drive-quick-list-head,.drive-quick-row{display:grid;grid-template-columns:120px minmax(230px,1.6fr) minmax(160px,1fr) 90px minmax(190px,.8fr);gap:10px;align-items:center}.drive-quick-list-head{position:sticky;top:0;z-index:2;padding:8px 12px;background:#122d45;color:rgba(219,234,254,.7);font-size:9.5px;font-weight:800;box-shadow:0 1px 0 rgba(255,255,255,.08)}.drive-quick-row{min-height:68px;padding:10px 12px;border-top:1px solid rgba(255,255,255,.055);background:rgba(255,255,255,.018);transition:.18s}.drive-quick-row:hover{background:rgba(255,255,255,.04)}.drive-quick-state{width:max-content;display:inline-flex;align-items:center;gap:5px;padding:5px 8px;border-radius:99px;font-style:normal;font-size:9px;font-weight:900}.drive-quick-state.done{color:#6ee7b7;background:rgba(52,211,153,.1)}.drive-quick-state.pending{color:#fcd34d;background:rgba(252,211,77,.09)}.drive-quick-name,.drive-quick-context,.drive-quick-responses{display:flex;flex-direction:column;min-width:0}.drive-quick-name b,.drive-quick-context b{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#eef6ff;font-size:10.5px}.drive-quick-name small,.drive-quick-context small{margin-top:3px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:rgba(255,255,255,.35);font-size:8.8px}.drive-quick-responses{align-items:center}.drive-quick-responses b{color:#fff;font-size:15px}.drive-quick-responses small{color:rgba(255,255,255,.34);font-size:8px}.drive-quick-responses em{color:rgba(255,255,255,.35);font-style:normal}.drive-quick-actions{display:flex;align-items:center;justify-content:flex-end;gap:5px}.drive-quick-actions button{height:31px;display:inline-flex;align-items:center;justify-content:center;gap:5px;padding:0 9px;border-radius:8px;font-family:'Cairo',sans-serif;font-size:9px;font-weight:900;cursor:pointer}.drive-quick-actions .report{border:1px solid rgba(96,165,250,.18);background:rgba(96,165,250,.1);color:#bfdbfe}.drive-quick-actions .analyze{border:1px solid rgba(52,211,153,.2);background:rgba(52,211,153,.11);color:#6ee7b7}.drive-quick-actions .source{padding:0 10px;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.055);color:#dce8f7;white-space:nowrap}.drive-quick-actions .source:hover{border-color:rgba(147,197,253,.3);background:rgba(96,165,250,.1);color:#fff}.drive-quick-empty{min-height:145px;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;border:1px dashed rgba(255,255,255,.09);border-radius:13px;color:#93c5fd}.drive-quick-empty b{margin-top:8px;color:#e8f0fe;font-size:12px}.drive-quick-empty span{margin-top:3px;color:rgba(255,255,255,.37);font-size:9.5px}.sr-only{position:absolute!important;width:1px!important;height:1px!important;padding:0!important;margin:-1px!important;overflow:hidden!important;clip:rect(0,0,0,0)!important;white-space:nowrap!important;border:0!important}
.drive-library-stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:13px;margin:16px 0}
.drive-library-stat{display:flex;align-items:center;gap:13px;min-width:0;padding:17px 18px;border:1px solid rgba(255,255,255,.09);border-radius:17px;background:rgba(255,255,255,.055);backdrop-filter:blur(12px)}
.drive-library-stat-icon{width:43px;height:43px;display:grid;place-items:center;flex:0 0 auto;border-radius:13px;color:var(--stat-tone);background:color-mix(in srgb,var(--stat-tone) 13%,transparent)}.drive-library-stat span:last-child{display:flex;flex-direction:column;min-width:0}.drive-library-stat b{color:#fff;font-size:22px;line-height:1.1}.drive-library-stat small{margin-top:5px;color:rgba(255,255,255,.52);font-size:11px;white-space:nowrap}
.drive-library-controls{padding:16px;border:1px solid rgba(255,255,255,.09);border-radius:18px;background:rgba(7,25,42,.47)}
.drive-library-search-row{display:grid;grid-template-columns:minmax(280px,1fr) auto;gap:14px;align-items:center}.drive-library-search{height:45px;display:flex;align-items:center;gap:10px;padding:0 14px;border:1px solid rgba(255,255,255,.12);border-radius:12px;background:rgba(4,18,31,.58);color:rgba(255,255,255,.42);transition:.2s}.drive-library-search:focus-within{border-color:rgba(52,211,153,.55);box-shadow:0 0 0 3px rgba(52,211,153,.08)}.drive-library-search input{width:100%;border:0;outline:0;background:transparent;color:#fff;font-family:'Cairo',sans-serif;font-size:13px}.drive-library-search input::placeholder{color:rgba(255,255,255,.34)}.drive-library-search button{display:grid;place-items:center;border:0;background:transparent;color:rgba(255,255,255,.45);cursor:pointer}
.drive-library-kind-tabs{display:flex;gap:5px;padding:4px;border-radius:12px;background:rgba(255,255,255,.055)}.drive-library-kind-tabs button{display:flex;align-items:center;gap:7px;padding:8px 12px;border:0;border-radius:9px;background:transparent;color:rgba(255,255,255,.56);font-family:'Cairo',sans-serif;font-size:12px;font-weight:700;cursor:pointer;transition:.2s}.drive-library-kind-tabs button b{min-width:20px;padding:1px 5px;border-radius:99px;background:rgba(255,255,255,.07);font-size:10px}.drive-library-kind-tabs button.active{color:#fff;background:rgba(52,211,153,.16);box-shadow:inset 0 0 0 1px rgba(52,211,153,.2)}
.drive-library-filter-row{display:grid;grid-template-columns:repeat(5,minmax(120px,1fr)) auto;gap:9px;margin-top:11px}.drive-library-filter-row select{min-width:0;height:39px;padding:0 11px;border:1px solid rgba(255,255,255,.1);border-radius:10px;outline:0;background:#102940;color:#dce8f7;font-family:'Cairo',sans-serif;font-size:11.5px;cursor:pointer}.drive-library-filter-row select:focus{border-color:#34d399}.drive-library-clear{padding:0 12px;border:1px solid rgba(251,113,133,.2);border-radius:10px;background:rgba(251,113,133,.08);color:#fda4af;font-family:'Cairo',sans-serif;font-size:11.5px;font-weight:700;cursor:pointer}
.drive-library-results-head{display:flex;align-items:center;justify-content:space-between;margin:19px 2px 11px;color:rgba(255,255,255,.52);font-size:12px}.drive-library-results-head b{color:#fff;font-size:14px}.drive-library-results-head small{margin-right:6px;color:rgba(255,255,255,.3);font-size:10px}.drive-library-view-toggle{display:flex;padding:3px;border-radius:9px;background:rgba(255,255,255,.06)}.drive-library-view-toggle button{width:34px;height:30px;display:grid;place-items:center;border:0;border-radius:7px;background:transparent;color:rgba(255,255,255,.45);cursor:pointer}.drive-library-view-toggle button.active{background:rgba(96,165,250,.16);color:#bfdbfe}
.drive-library-files.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:13px}.drive-library-file{position:relative;display:flex;flex-direction:column;min-width:0;padding:17px;border:1px solid rgba(255,255,255,.09);border-radius:17px;background:linear-gradient(145deg,rgba(255,255,255,.068),rgba(255,255,255,.035));transition:transform .2s,border-color .2s,box-shadow .2s}.drive-library-file:hover{transform:translateY(-3px);border-color:color-mix(in srgb,var(--file-color) 42%,transparent);box-shadow:0 15px 38px rgba(1,10,20,.22)}.drive-library-file-top{display:flex;align-items:center;justify-content:space-between;gap:12px}.drive-library-file-icon{width:45px;height:45px;display:grid;place-items:center;border-radius:13px;background:var(--file-bg);color:var(--file-color)}.drive-library-kind{padding:4px 9px;border-radius:99px;font-size:10px;font-weight:800}.drive-library-kind.survey{color:#6ee7b7;background:rgba(52,211,153,.1)}.drive-library-kind.report{color:#fda4af;background:rgba(251,113,133,.1)}
.drive-library-file-copy{min-width:0;flex:1}.drive-library-file h2{display:-webkit-box;overflow:hidden;-webkit-box-orient:vertical;-webkit-line-clamp:2;min-height:49px;margin:15px 0 6px;color:#f8fbff;font-size:14px;line-height:1.75}.drive-library-file p{color:rgba(255,255,255,.38);font-size:10.5px}.drive-library-file-tags{display:flex;flex-wrap:wrap;gap:6px;margin:12px 0 16px}.drive-library-file-tags span{max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:4px 8px;border:1px solid rgba(255,255,255,.07);border-radius:7px;background:rgba(255,255,255,.035);color:rgba(255,255,255,.58);font-size:9.5px}
.drive-library-file-actions{display:flex;gap:7px;margin-top:auto;padding-top:12px;border-top:1px solid rgba(255,255,255,.065)}.drive-library-file-actions button{display:inline-flex;align-items:center;justify-content:center;gap:6px;min-height:34px;padding:6px 10px;border-radius:9px;font-family:'Cairo',sans-serif;font-size:10.5px;font-weight:800;cursor:pointer;transition:.18s}.drive-library-open{flex:1;border:1px solid rgba(255,255,255,.11);background:rgba(255,255,255,.055);color:#dbeafe}.drive-library-open:hover{background:rgba(255,255,255,.1)}.drive-library-analyze{flex:1;border:1px solid rgba(52,211,153,.19);background:rgba(52,211,153,.1);color:#6ee7b7}.drive-library-analyze:hover{background:rgba(52,211,153,.18)}
.drive-library-files.list{display:flex;flex-direction:column;gap:8px}.drive-library-files.list .drive-library-file{display:grid;grid-template-columns:auto minmax(220px,1fr) minmax(190px,auto);align-items:center;gap:15px;padding:12px 14px}.drive-library-files.list .drive-library-file-top{flex-direction:column;gap:5px}.drive-library-files.list .drive-library-file-icon{width:39px;height:39px}.drive-library-files.list .drive-library-kind{padding:2px 7px}.drive-library-files.list .drive-library-file h2{min-height:auto;margin:0 0 3px;-webkit-line-clamp:1}.drive-library-files.list .drive-library-file-tags{margin:7px 0 0}.drive-library-files.list .drive-library-file-actions{margin:0;padding:0;border:0;min-width:190px}
.drive-library-loading,.drive-library-empty{min-height:270px;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;border:1px dashed rgba(255,255,255,.12);border-radius:18px;background:rgba(255,255,255,.025)}.drive-library-loading b,.drive-library-empty h2{margin:14px 0 4px;color:#fff;font-size:15px}.drive-library-loading small,.drive-library-empty p{color:rgba(255,255,255,.45);font-size:11.5px}.drive-library-spinner{width:36px;height:36px;border:3px solid rgba(255,255,255,.09);border-top-color:#34d399;border-radius:50%;animation:driveLibrarySpin .85s linear infinite}.drive-library-empty>span{width:56px;height:56px;display:grid;place-items:center;border-radius:17px;color:#93c5fd;background:rgba(96,165,250,.1)}.drive-library-empty-actions{display:flex;gap:8px;flex-wrap:wrap;justify-content:center}.drive-library-empty button{margin-top:14px;padding:8px 15px;border:1px solid rgba(96,165,250,.2);border-radius:9px;background:rgba(96,165,250,.1);color:#bfdbfe;font-family:'Cairo',sans-serif;font-weight:700;cursor:pointer}.drive-library-empty button:disabled{opacity:.55;cursor:wait}
.drive-library-pagination{display:flex;justify-content:center;padding:24px 0 5px}.drive-library-pagination>span{padding:9px 14px;color:rgba(255,255,255,.34);font-size:10.5px}.drive-library-pagination button{min-width:235px;min-height:47px;display:inline-flex;align-items:center;justify-content:center;gap:9px;padding:9px 20px;border:1px solid rgba(52,211,153,.2);border-radius:13px;background:linear-gradient(135deg,rgba(52,211,153,.14),rgba(96,165,250,.08));color:#a7f3d0;font-family:'Cairo',sans-serif;font-size:12px;font-weight:900;cursor:pointer;transition:.2s}.drive-library-pagination button:hover:not(:disabled){transform:translateY(-2px);border-color:rgba(52,211,153,.38);box-shadow:0 10px 25px rgba(3,18,30,.2)}.drive-library-pagination button:disabled{opacity:.58;cursor:wait}.drive-library-pagination button small{color:rgba(255,255,255,.38);font-size:9.5px;font-weight:600}.drive-library-mini-spinner{width:17px;height:17px;border:2px solid rgba(255,255,255,.15);border-top-color:#6ee7b7;border-radius:50%;animation:driveLibrarySpin .8s linear infinite}
.drive-library-disconnected{min-height:calc(100vh - 150px);display:grid;place-items:center}.drive-library-connect-card{position:relative;overflow:hidden;max-width:670px;padding:48px 42px;border:1px solid rgba(96,165,250,.17);border-radius:28px;background:radial-gradient(circle at 50% 0,rgba(96,165,250,.17),transparent 36%),rgba(8,29,47,.72);text-align:center;box-shadow:0 25px 70px rgba(0,0,0,.22)}.drive-library-drive-mark{width:86px;height:86px;display:grid;place-items:center;margin:0 auto 17px;border:1px solid rgba(255,255,255,.1);border-radius:25px;background:rgba(255,255,255,.07)}.drive-library-connect-card h1{margin:10px 0;color:#fff;font-size:30px}.drive-library-connect-card p{max-width:530px;margin:0 auto;color:rgba(255,255,255,.57);font-size:14px;line-height:1.9}.drive-library-primary{margin:24px auto 13px;padding:12px 22px;background:linear-gradient(135deg,#34d399,#168f78);color:#fff;box-shadow:0 12px 28px rgba(22,143,120,.23)}.drive-library-primary:hover:not(:disabled){transform:translateY(-2px);box-shadow:0 15px 32px rgba(22,143,120,.35)}.drive-library-primary:disabled{opacity:.6}.drive-library-connect-card>small{display:block;color:rgba(255,255,255,.32);font-size:10.5px}
@media(max-width:1080px){.drive-library-stats,.drive-quick-stats{grid-template-columns:repeat(2,1fr)}.drive-library-files.grid{grid-template-columns:repeat(2,minmax(0,1fr))}.drive-library-filter-row{grid-template-columns:repeat(3,1fr)}.drive-library-clear{min-height:39px}.drive-library-files.list .drive-library-file{grid-template-columns:auto minmax(180px,1fr)}.drive-library-files.list .drive-library-file-actions{grid-column:2;min-width:0}.drive-quick-toolbar{align-items:stretch;flex-direction:column}.drive-quick-filter-tools{width:100%}.drive-quick-filter-tools>*{flex:1}.drive-quick-list{overflow:auto}.drive-quick-list-head,.drive-quick-row{min-width:920px}}
@media(max-width:720px){.drive-library-hero{align-items:flex-start;flex-direction:column;padding:22px}.drive-library-hero-actions{width:100%;flex-direction:column}.drive-library-refresh,.drive-library-quick-button{width:100%}.drive-library-stats{grid-template-columns:1fr 1fr;gap:8px}.drive-library-stat{padding:13px}.drive-library-stat-icon{width:38px;height:38px}.drive-library-stat b{font-size:18px}.drive-library-search-row{grid-template-columns:1fr}.drive-library-kind-tabs{overflow-x:auto}.drive-library-kind-tabs button{flex:1;justify-content:center;white-space:nowrap}.drive-library-filter-row{grid-template-columns:1fr 1fr}.drive-library-files.grid{grid-template-columns:1fr}.drive-library-files.list .drive-library-file{display:flex;align-items:stretch}.drive-library-files.list .drive-library-file-top{flex-direction:row}.drive-library-files.list .drive-library-file-actions{min-width:0}.drive-library-connect-card{padding:36px 22px}.drive-library-connect-card h1{font-size:25px}.drive-quick-summary{padding:15px}.drive-quick-head{gap:10px}.drive-quick-head h2{font-size:17px}.drive-quick-controls{align-items:stretch;flex-direction:column}.drive-quick-controls select,.drive-quick-controls>button{width:100%}.drive-quick-toolbar{align-items:stretch;flex-direction:column}.drive-quick-status-tabs{overflow-x:auto}.drive-quick-status-tabs button{flex:1;justify-content:center;white-space:nowrap}.drive-quick-filter-tools{align-items:stretch;flex-direction:column}.drive-quick-type-filter,.drive-quick-search{min-width:0;width:100%}}
@media(max-width:460px){.drive-library-stats{grid-template-columns:1fr}.drive-library-filter-row{grid-template-columns:1fr}.drive-library-file-actions{flex-direction:column}}
`;
