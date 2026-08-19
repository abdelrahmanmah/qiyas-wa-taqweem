import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  loadGisScript, initSemesterTokenClient,
  SEMESTER_TOKEN_KEY, saveStoredToken, getStoredToken,
  TEMPLATE_FOLDER_ID, ROOT_SURVEYS_FOLDER_ID,
  SEMESTERS, loadDepartments, saveDepartments, expandDepartmentUnits,
  listFormsInFolder, listSubfolders,
  getForm, listAllResponses, computeResponseStats,
  buildGenerationJobs, runGenerationJobs, listSemesterSurveysWithStats,
  yearFolderName, isValidAcademicYear,
  editorResponsesUrl, responsesToRows, departmentFromSurveyName,
} from "./engine/semesterSurveyModel.js";
import { analyze } from "./engine/analyze.js";
import { buildAnnualDocx, DEFAULT_SETTINGS } from "./engine/buildDocx.js";
import { getAllAnalysisSchemas, detectAnySurveyType } from "./engine/customSurveyModel.js";

const SETTINGS_KEY = "eruQA_settings_v1"; // same key App.jsx's SettingsPanel writes to

function loadReportSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
  } catch { /* ignore */ }
  return { ...DEFAULT_SETTINGS };
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Shared "links export" helpers — used by both the Generate results table and the
// Dashboard's survey table, wherever a list of { name, formUrl, responsesUrl } exists.
function buildOrganizedMessage(rows, title) {
  const lines = [`📋 ${title}`, ""];
  rows.forEach((r, i) => {
    lines.push(`${i + 1}. ${r.name}`);
    lines.push(`   🔗 النموذج: ${r.formUrl}`);
    lines.push(`   📊 الردود: ${r.responsesUrl}`);
    lines.push("");
  });
  return lines.join("\n").trim();
}

function buildLinksWorkbookBlob(rows) {
  const data = rows.map(r => ({
    "الاستبيان": r.name,
    "رابط النموذج": r.formUrl,
    "رابط الردود": r.responsesUrl,
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = [{ wch: 45 }, { wch: 55 }, { wch: 55 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "روابط الاستبيانات");
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function LinksExportButtons({ rows, title, filename, pushToast }) {
  if (!rows.length) return null;
  return (
    <div style={{ display: "flex", gap: 10 }}>
      <button className="btn btn-ghost btn-sm" onClick={() => {
        copyToClipboard(buildOrganizedMessage(rows, title), pushToast);
      }}>📋 نسخ الروابط كرسالة</button>
      <button className="btn btn-ghost btn-sm" onClick={() => {
        downloadBlob(buildLinksWorkbookBlob(rows), filename);
        pushToast("تم تنزيل ملف Excel.", "success");
      }}>⬇ تنزيل Excel</button>
    </div>
  );
}

const CSS = `
@keyframes ssgShimmer{0%{background-position:-300px 0}100%{background-position:300px 0}}
@keyframes ssgToastIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
@keyframes ssgProgressFlow{0%{background-position:0 0}100%{background-position:40px 0}}
@keyframes ssgTipFade{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}
.ssg-progress-track{background:rgba(255,255,255,.08);border-radius:6px;height:10px;overflow:hidden}
.ssg-progress-fill{height:100%;border-radius:6px;transition:width .4s ease;
  background-image:linear-gradient(135deg,rgba(255,255,255,.2) 25%,transparent 25%,transparent 50%,
    rgba(255,255,255,.2) 50%,rgba(255,255,255,.2) 75%,transparent 75%,transparent);
  background-color:#1abc9c;background-size:28px 28px;animation:ssgProgressFlow 1s linear infinite}
.ssg-tip{animation:ssgTipFade .35s ease-out}
.ssg-skel-row{height:38px;border-radius:10px;margin-bottom:8px;
  background:linear-gradient(90deg,rgba(255,255,255,.05) 25%,rgba(255,255,255,.12) 37%,rgba(255,255,255,.05) 63%);
  background-size:600px 100%;animation:ssgShimmer 1.4s infinite linear}
.ssg-tpl-row{display:flex;align-items:center;justify-content:space-between;gap:14px;
  padding:12px 16px;border-radius:12px;background:rgba(255,255,255,.04);
  border:1px solid rgba(255,255,255,.08);margin-bottom:8px}
.ssg-tpl-row input[type=checkbox]{width:17px;height:17px;accent-color:#1abc9c;cursor:pointer}
.ssg-tpl-dept{display:flex;align-items:center;gap:8px;font-size:12px;color:rgba(255,255,255,.55);white-space:nowrap}
.ssg-stat-card{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);
  border-radius:14px;padding:16px 18px}
.ssg-toast{animation:ssgToastIn .25s ease-out;padding:12px 18px;border-radius:12px;
  font-size:13px;font-weight:700;color:#fff;box-shadow:0 8px 24px rgba(0,0,0,.35);min-width:220px}
`;

// ---------- toasts ----------
function useToasts() {
  const [toasts, setToasts] = useState([]);
  const push = useCallback((text, kind = "success") => {
    const id = Date.now() + Math.random();
    setToasts(t => [...t, { id, text, kind }]);
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 4000);
  }, []);
  const dismiss = useCallback(id => setToasts(t => t.filter(x => x.id !== id)), []);
  return { toasts, push, dismiss };
}

function ToastStack({ toasts, onDismiss }) {
  if (!toasts.length) return null;
  return (
    <div style={{ position: "fixed", bottom: 24, left: 24, zIndex: 9999, display: "flex", flexDirection: "column", gap: 10 }}>
      {toasts.map(t => (
        <div
          key={t.id}
          className="ssg-toast"
          style={{ background: t.kind === "error" ? "linear-gradient(135deg,#e74c3c,#c0392b)" : "linear-gradient(135deg,#1abc9c,#16a085)", cursor: "pointer" }}
          onClick={() => onDismiss(t.id)}
        >
          {t.text}
        </div>
      ))}
    </div>
  );
}

// ---------- auth ----------
// Persists the access token across reloads (localStorage, ~1hr lifetime) and, once it
// expires, tries a silent requestAccessToken({prompt:""}) in the background before ever
// falling back to the visible "الاتصال بـ Google" button — so a returning user only sees
// the account picker once, not on every visit. See semesterSurveyModel.js for why.
function useSemesterAuth() {
  const [token, setToken] = useState(() => getStoredToken(SEMESTER_TOKEN_KEY));
  const [connecting, setConnecting] = useState(false);
  const [authError, setAuthError] = useState("");
  const clientRef = useRef(null);
  const explicitRef = useRef(false);
  const triedSilent = useRef(false);

  const ensureClient = useCallback(async () => {
    await loadGisScript();
    if (!clientRef.current) {
      clientRef.current = initSemesterTokenClient({
        onToken: (t, expiresIn) => {
          setToken(t);
          saveStoredToken(SEMESTER_TOKEN_KEY, t, expiresIn);
          setConnecting(false);
        },
        onError: e => {
          setConnecting(false);
          // A silent background attempt failing is normal (no prior session/consent) —
          // only surface an error banner for an explicit, user-clicked connect.
          if (explicitRef.current) setAuthError("فشل الاتصال بـ Google: " + e);
        },
      });
    }
    return clientRef.current;
  }, []);

  useEffect(() => {
    if (token || triedSilent.current) return;
    triedSilent.current = true;
    (async () => {
      try {
        const client = await ensureClient();
        explicitRef.current = false;
        client.requestAccessToken({ prompt: "" });
      } catch { /* ignore — falls through to the manual connect button */ }
    })();
  }, [token, ensureClient]);

  const connect = useCallback(async () => {
    setAuthError(""); setConnecting(true);
    try {
      const client = await ensureClient();
      explicitRef.current = true;
      client.requestAccessToken({ prompt: "select_account" });
    } catch (e) {
      setAuthError(e.message); setConnecting(false);
    }
  }, [ensureClient]);

  return { token, connecting, authError, connect };
}

// ---------- small shared pieces ----------
function Skeleton({ rows = 4 }) {
  return (
    <div>
      {Array.from({ length: rows }).map((_, i) => <div key={i} className="ssg-skel-row" />)}
    </div>
  );
}

function EmptyState({ icon = "📭", text }) {
  return (
    <div style={{ textAlign: "center", padding: "36px 20px", color: "rgba(255,255,255,.45)" }}>
      <div style={{ fontSize: 34, marginBottom: 8 }}>{icon}</div>
      <div style={{ fontSize: 13 }}>{text}</div>
    </div>
  );
}

function ErrorBanner({ text }) {
  if (!text) return null;
  return (
    <div style={{ color: "#ff6b5b", fontSize: 13, background: "rgba(231,76,60,.08)", borderRadius: 10, padding: "10px 14px", marginTop: 10 }}>
      {text}
    </div>
  );
}

function BarChart({ data, gradient = "linear-gradient(90deg,#1abc9c,#2874a6)" }) {
  const max = Math.max(1, ...data.map(d => d.value));
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      {data.map(d => (
        <div key={d.label}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "rgba(255,255,255,.7)", marginBottom: 4 }}>
            <span>{d.label}</span><span>{d.value}</span>
          </div>
          <div style={{ background: "rgba(255,255,255,.08)", borderRadius: 3, height: 8, overflow: "hidden" }}>
            <div style={{ height: "100%", borderRadius: 3, transition: "width .5s", background: gradient, width: `${(d.value / max) * 100}%` }} />
          </div>
        </div>
      ))}
    </div>
  );
}

// Cycles through short status lines while a long batch operation runs, so the wait
// doesn't feel like a frozen screen.
function useRotatingTip(active, tips, intervalMs = 2600) {
  const [idx, setIdx] = useState(0);
  useEffect(() => {
    if (!active) { setIdx(0); return; }
    const t = setInterval(() => setIdx(i => (i + 1) % tips.length), intervalMs);
    return () => clearInterval(t);
  }, [active, tips, intervalMs]);
  return tips[idx];
}

function ProgressBar({ done, total, label, tip }) {
  const pct = total ? Math.round((done / total) * 100) : 0;
  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "rgba(255,255,255,.7)", marginBottom: 6 }}>
        <span>{label}</span>
        <span>{done} / {total} — {pct}%</span>
      </div>
      <div className="ssg-progress-track">
        <div className="ssg-progress-fill" style={{ width: `${pct}%` }} />
      </div>
      {tip && <div key={tip} className="ssg-tip" style={{ textAlign: "center", color: "rgba(255,255,255,.5)", fontSize: 12, marginTop: 10 }}>{tip}</div>}
    </div>
  );
}

function copyToClipboard(text, pushToast) {
  navigator.clipboard?.writeText(text)
    .then(() => pushToast("تم نسخ الرابط.", "success"))
    .catch(() => pushToast("تعذّر نسخ الرابط.", "error"));
}

function formatWhen(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  const now = new Date();
  const days = Math.floor((now - d) / 86400000);
  if (days <= 0) return "اليوم";
  if (days === 1) return "أمس";
  return d.toLocaleDateString("ar-EG", { year: "numeric", month: "short", day: "numeric" });
}

// ---------- Generate Surveys ----------
const GENERATE_TIPS = [
  "🔄 جاري نسخ النماذج من القوالب...",
  "📁 جاري تنظيم مجلدات السنة والفصل الدراسي على Google Drive...",
  "✏️ جاري تحديث عنوان كل استبيان...",
  "🚀 جاري نشر الاستبيانات ليصبح بإمكانها استقبال الردود...",
  "⏳ لحظات ونكون قد انتهينا...",
];

function GenerateSurveysView({ token, pushToast }) {
  const [templates, setTemplates] = useState(null); // null = loading
  const [loadError, setLoadError] = useState("");
  const [year, setYear] = useState("");
  const [yearOptions, setYearOptions] = useState([]);
  const [semester, setSemester] = useState(SEMESTERS[0]);
  const [validationMsg, setValidationMsg] = useState("");
  const [jobs, setJobs] = useState(null);
  const [generating, setGenerating] = useState(false);
  const [results, setResults] = useState(null);
  const tip = useRotatingTip(generating, GENERATE_TIPS);

  useEffect(() => {
    let cancelled = false;
    setTemplates(null); setLoadError("");
    listFormsInFolder(token, TEMPLATE_FOLDER_ID)
      .then(files => { if (!cancelled) setTemplates(files.map(f => ({ ...f, selected: false, mode: "general" }))); })
      .catch(e => { if (!cancelled) setLoadError("فشل تحميل القوالب: " + e.message); });
    return () => { cancelled = true; };
  }, [token]);

  // Suggestions only — the year field stays free-text so a new year can always be typed.
  useEffect(() => {
    listSubfolders(token, ROOT_SURVEYS_FOLDER_ID)
      .then(fs => setYearOptions(fs.map(f => f.name)))
      .catch(() => {});
  }, [token]);

  const allSelected = !!templates?.length && templates.every(t => t.selected);
  const toggleAll = () => setTemplates(ts => ts.map(t => ({ ...t, selected: !allSelected })));
  const toggleOne = id => setTemplates(ts => ts.map(t => t.id === id ? { ...t, selected: !t.selected } : t));
  const setTemplateMode = (id, mode) => setTemplates(ts => ts.map(t => t.id === id ? { ...t, mode } : t));
  const selectedCount = templates?.filter(t => t.selected).length ?? 0;

  async function handleGenerate() {
    setValidationMsg("");
    if (selectedCount === 0) { setValidationMsg("الرجاء اختيار استبيان واحد على الأقل."); return; }
    if (!isValidAcademicYear(year)) { setValidationMsg("الرجاء إدخال السنة الدراسية بالصيغة 2026/2027."); return; }

    const departments = loadDepartments();
    const selected = templates.filter(t => t.selected);
    const jobList = buildGenerationJobs(selected, departments);

    setJobs(jobList.map(j => ({
      key: j.key, status: "pending",
      label: j.department ? `${j.template.name} — ${j.department}` : j.template.name,
    })));
    setResults(null);
    setGenerating(true);

    try {
      const { successRows, failCount } = await runGenerationJobs(
        token, jobList, { year, semester },
        (key, status) => setJobs(js => js.map(j => j.key === key ? { ...j, status } : j)),
      );
      setResults(successRows);
      if (successRows.length) pushToast(`تم إنشاء ${successRows.length} استبيان بنجاح.`, "success");
      if (failCount) pushToast(`فشل إنشاء ${failCount} استبيان.`, "error");
    } catch (e) {
      pushToast("فشل إنشاء مجلدات الفصل الدراسي: " + e.message, "error");
    } finally {
      setGenerating(false);
    }
  }

  return (
    <div>
      <div className="card" style={{ padding: 24, marginBottom: 20 }}>
        <div style={{ color: "#fff", fontSize: 15, fontWeight: 800, marginBottom: 14 }}>1) اختيار القوالب</div>
        {templates === null ? (
          <Skeleton rows={5} />
        ) : loadError ? (
          <ErrorBanner text={loadError} />
        ) : templates.length === 0 ? (
          <EmptyState text="لا توجد نماذج Google Forms داخل مجلد القوالب المحدد." />
        ) : (
          <>
            <label className="ssg-tpl-row" style={{ cursor: "pointer" }}>
              <span style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <input type="checkbox" checked={allSelected} onChange={toggleAll} />
                <span style={{ color: "#fff", fontWeight: 700 }}>تحديد الكل</span>
              </span>
            </label>
            {templates.map(t => (
              <div key={t.id} className="ssg-tpl-row">
                <label style={{ display: "flex", alignItems: "center", gap: 10, flex: 1, cursor: "pointer" }}>
                  <input type="checkbox" checked={t.selected} onChange={() => toggleOne(t.id)} />
                  <span style={{ color: "#e8f0fe", fontSize: 14 }}>{t.name}</span>
                </label>
                <select
                  value={t.mode} disabled={!t.selected} onChange={e => setTemplateMode(t.id, e.target.value)}
                  title="نسخة عامة واحدة، أو نسخة لكل قسم، أو نسخة لكل قسم وبرنامج (للأقسام التي لها برامج)"
                  style={{
                    opacity: t.selected ? 1 : .45, cursor: t.selected ? "pointer" : "default",
                    padding: "6px 10px", borderRadius: 8, border: "1px solid rgba(255,255,255,.18)",
                    background: "rgba(255,255,255,.06)", color: "#e8f0fe", fontFamily: "'Cairo',sans-serif", fontSize: 12,
                  }}
                >
                  <option value="general" style={{ color: "#000" }}>نسخة عامة</option>
                  <option value="departments" style={{ color: "#000" }}>الأقسام</option>
                  <option value="programs" style={{ color: "#000" }}>الأقسام والبرامج</option>
                </select>
              </div>
            ))}
          </>
        )}
      </div>

      <div className="card" style={{ padding: 24, marginBottom: 20 }}>
        <div style={{ color: "#fff", fontSize: 15, fontWeight: 800, marginBottom: 14 }}>2) بيانات الفصل الدراسي</div>
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
          <div style={{ flex: "1 1 200px" }}>
            <div style={{ color: "rgba(255,255,255,.6)", fontSize: 12, marginBottom: 6 }}>السنة الدراسية</div>
            <input
              value={year} onChange={e => setYear(e.target.value)} placeholder="2026/2027"
              disabled={generating} list="ssg-year-options" autoComplete="off"
              style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: "1px solid rgba(255,255,255,.18)", background: "rgba(255,255,255,.06)", color: "#fff", fontFamily: "'Cairo',sans-serif", fontSize: 14 }}
            />
            <datalist id="ssg-year-options">
              {yearOptions.map(y => <option key={y} value={y} />)}
            </datalist>
          </div>
          <div style={{ flex: "1 1 200px" }}>
            <div style={{ color: "rgba(255,255,255,.6)", fontSize: 12, marginBottom: 6 }}>الفصل الدراسي</div>
            <select
              value={semester} onChange={e => setSemester(e.target.value)} disabled={generating}
              style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: "1px solid rgba(255,255,255,.18)", background: "rgba(255,255,255,.06)", color: "#fff", fontFamily: "'Cairo',sans-serif", fontSize: 14 }}
            >
              {SEMESTERS.map(s => <option key={s} value={s} style={{ color: "#000" }}>{s}</option>)}
            </select>
          </div>
        </div>
      </div>

      <div className="card" style={{ padding: 24 }}>
        <button className="btn btn-primary" disabled={generating || selectedCount === 0} onClick={handleGenerate}>
          {generating ? "⏳ جاري الإنشاء..." : `⚙ إنشاء (${selectedCount})`}
        </button>
        <ErrorBanner text={validationMsg} />

        {jobs && (
          <div style={{ marginTop: 20 }}>
            {generating && (
              <ProgressBar
                done={jobs.filter(j => j.status === "done" || j.status === "error").length}
                total={jobs.length}
                label="جاري إنشاء الاستبيانات..."
                tip={tip}
              />
            )}
            {jobs.map(j => (
              <div key={j.key} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 4px", fontSize: 13, color: "rgba(255,255,255,.8)" }}>
                <span>{j.status === "done" ? "✔" : j.status === "error" ? "✖" : j.status === "active" ? "⏳" : "○"}</span>
                <span>{j.label}</span>
              </div>
            ))}
          </div>
        )}

        {results && results.length > 0 && (
          <div style={{ marginTop: 24 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
              <div style={{ color: "#1abc9c", fontWeight: 800 }}>
                تم إنشاء {results.length} استبيان بنجاح.
              </div>
              <LinksExportButtons
                rows={results.map(r => ({ name: r.name, formUrl: r.formUrl, responsesUrl: editorResponsesUrl(r.formId) }))}
                title={`استبيانات ${semester} ${year}`}
                filename={`روابط_استبيانات_${semester}_${yearFolderName(year)}.xlsx`}
                pushToast={pushToast}
              />
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ color: "rgba(255,255,255,.5)", textAlign: "right" }}>
                    <th style={{ padding: "8px 10px" }}>الاستبيان</th>
                    <th style={{ padding: "8px 10px" }}>النموذج</th>
                    <th style={{ padding: "8px 10px" }}>الردود</th>
                    <th style={{ padding: "8px 10px" }}>روابط</th>
                  </tr>
                </thead>
                <tbody>
                  {results.map(r => (
                    <tr key={r.key} style={{ borderTop: "1px solid rgba(255,255,255,.08)", color: "#e8f0fe" }}>
                      <td style={{ padding: "8px 10px" }}>{r.name}</td>
                      <td style={{ padding: "8px 10px" }}><a href={r.formUrl} target="_blank" rel="noreferrer" style={{ color: "#1abc9c" }}>فتح النموذج</a></td>
                      <td style={{ padding: "8px 10px" }}><a href={editorResponsesUrl(r.formId)} target="_blank" rel="noreferrer" style={{ color: "#1abc9c" }}>فتح الردود</a></td>
                      <td style={{ padding: "8px 10px", display: "flex", gap: 10 }}>
                        <button className="btn btn-ghost btn-sm" onClick={() => copyToClipboard(r.formUrl, pushToast)}>نسخ رابط النموذج</button>
                        <button className="btn btn-ghost btn-sm" onClick={() => copyToClipboard(editorResponsesUrl(r.formId), pushToast)}>نسخ رابط الردود</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

const ANALYZE_TIPS = [
  "🔎 جاري جلب ردود كل استبيان...",
  "📊 جاري حساب المؤشرات الإحصائية...",
  "📄 جاري إنشاء تقرير Word لكل استبيان...",
  "💾 جاري تنزيل التقارير...",
  "⏳ لحظات ونكون قد انتهينا...",
];

// ---------- Dashboard ----------
function DashboardView({ token, pushToast }) {
  const [years, setYears] = useState(null);
  const [year, setYear] = useState("");
  const [semesters, setSemesters] = useState(null);
  const [semester, setSemester] = useState("");
  const [surveys, setSurveys] = useState(null);
  const [loadError, setLoadError] = useState("");
  const [refreshKey, setRefreshKey] = useState(0);

  useEffect(() => {
    let cancelled = false;
    setYears(null); setLoadError("");
    listSubfolders(token, ROOT_SURVEYS_FOLDER_ID)
      .then(fs => { if (cancelled) return; setYears(fs); setYear(y => y || fs[0]?.name || ""); })
      .catch(e => { if (!cancelled) setLoadError("فشل تحميل السنوات الدراسية: " + e.message); });
    return () => { cancelled = true; };
  }, [token, refreshKey]);

  const yearFolder = years?.find(f => f.name === year) ?? null;

  useEffect(() => {
    if (!yearFolder) { setSemesters(null); return; }
    let cancelled = false;
    setSemesters(null);
    listSubfolders(token, yearFolder.id)
      .then(fs => { if (cancelled) return; setSemesters(fs); setSemester(s => (fs.some(f => f.name === s) ? s : (fs[0]?.name || ""))); })
      .catch(e => { if (!cancelled) setLoadError("فشل تحميل الفصول الدراسية: " + e.message); });
    return () => { cancelled = true; };
  }, [yearFolder?.id, refreshKey]);

  const semFolder = semesters?.find(f => f.name === semester) ?? null;

  const loadSurveys = useCallback(() => {
    if (!semFolder) { setSurveys([]); return; }
    setSurveys(null); setLoadError("");
    // Each survey type now lives in its own subfolder under the semester folder
    // (see GenerateSurveysView) — listSemesterSurveysWithStats walks that structure.
    listSemesterSurveysWithStats(token, semFolder.id)
      .then(rows => setSurveys(rows))
      .catch(e => setLoadError("فشل تحميل الاستبيانات: " + e.message));
  }, [token, semFolder?.id]);

  useEffect(() => { loadSurveys(); }, [loadSurveys, refreshKey]);

  async function refreshOne(id) {
    try {
      const responses = await listAllResponses(token, id);
      const stats = computeResponseStats(responses);
      setSurveys(rows => rows.map(r => r.id === id ? { ...r, responses: stats.count, last: stats.lastSubmittedTime, error: null } : r));
      pushToast("تم تحديث بيانات الاستبيان.", "success");
    } catch (e) {
      pushToast("فشل تحديث الاستبيان: " + e.message, "error");
    }
  }

  const totalSurveys = surveys?.length ?? 0;
  const totalResponses = surveys?.reduce((s, r) => s + r.responses, 0) ?? 0;
  const avg = totalSurveys ? Math.round((totalResponses / totalSurveys) * 10) / 10 : 0;
  const lastResponse = surveys?.reduce((max, r) => (r.last && (!max || r.last > max)) ? r.last : max, null) ?? null;

  // ---- Analyze all: reuse the app's existing analysis engine (analyze.js) + Word
  // builder (buildDocx.js) unmodified, by converting each form's responses into the
  // same [header, ...rows] shape a real Excel export would have. ----
  const [showAnalyzeForm, setShowAnalyzeForm] = useState(false);
  const [preparedBy, setPreparedBy] = useState("");
  const [reviewer, setReviewer] = useState("");
  const [analyzing, setAnalyzing] = useState(false);
  const [analyzeProgress, setAnalyzeProgress] = useState(null);
  const analyzeTip = useRotatingTip(analyzing, ANALYZE_TIPS);

  async function runAnalyzeAll() {
    setAnalyzing(true);
    setAnalyzeProgress(surveys.map(s => ({ id: s.id, name: s.name, status: "pending" })));
    const schemas = getAllAnalysisSchemas();
    const settings = loadReportSettings();
    let successCount = 0, skipCount = 0;

    for (const survey of surveys) {
      setAnalyzeProgress(p => p.map(x => x.id === survey.id ? { ...x, status: "active" } : x));
      try {
        const [form, responses] = await Promise.all([getForm(token, survey.id), listAllResponses(token, survey.id)]);
        const rows = responsesToRows(form, responses);
        const schemaId = detectAnySurveyType(survey.name, rows[0]);
        const schema = schemaId ? schemas[schemaId] : null;
        if (!schema) {
          setAnalyzeProgress(p => p.map(x => x.id === survey.id ? { ...x, status: "skipped" } : x));
          skipCount++;
          continue;
        }
        const result = analyze(rows, schema);
        const meta = { year, program: departmentFromSurveyName(survey.name), preparedBy, reviewer };
        const blob = await buildAnnualDocx(result, meta, settings);
        downloadBlob(blob, `تقرير_${schema.label}_${survey.name}.docx`);
        setAnalyzeProgress(p => p.map(x => x.id === survey.id ? { ...x, status: "done" } : x));
        successCount++;
        await new Promise(r => setTimeout(r, 450)); // let the browser process each download separately
      } catch (e) {
        setAnalyzeProgress(p => p.map(x => x.id === survey.id ? { ...x, status: "error" } : x));
      }
    }

    setAnalyzing(false);
    if (successCount) pushToast(`تم إنشاء ${successCount} تقرير بنجاح.`, "success");
    if (skipCount) pushToast(`تعذّر التعرف على نوع ${skipCount} استبيان — لم يتم تحليله.`, "error");
  }

  return (
    <div>
      <div className="card" style={{ padding: 24, marginBottom: 20 }}>
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
          <div style={{ flex: "1 1 200px" }}>
            <div style={{ color: "rgba(255,255,255,.6)", fontSize: 12, marginBottom: 6 }}>السنة الدراسية</div>
            <select value={year} onChange={e => setYear(e.target.value)}
              style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: "1px solid rgba(255,255,255,.18)", background: "rgba(255,255,255,.06)", color: "#fff", fontFamily: "'Cairo',sans-serif", fontSize: 14 }}>
              {(years ?? []).map(f => <option key={f.id} value={f.name} style={{ color: "#000" }}>{f.name}</option>)}
            </select>
          </div>
          <div style={{ flex: "1 1 200px" }}>
            <div style={{ color: "rgba(255,255,255,.6)", fontSize: 12, marginBottom: 6 }}>الفصل الدراسي</div>
            <select value={semester} onChange={e => setSemester(e.target.value)}
              style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: "1px solid rgba(255,255,255,.18)", background: "rgba(255,255,255,.06)", color: "#fff", fontFamily: "'Cairo',sans-serif", fontSize: 14 }}>
              {(semesters ?? []).map(f => <option key={f.id} value={f.name} style={{ color: "#000" }}>{f.name}</option>)}
            </select>
          </div>
          <button className="btn btn-ghost btn-sm" onClick={() => setRefreshKey(k => k + 1)}>↻ تحديث</button>
          {!!surveys?.length && (
            <button className="btn btn-primary btn-sm" disabled={analyzing} onClick={() => setShowAnalyzeForm(v => !v)}>
              🔍 تحليل الكل ({surveys.length})
            </button>
          )}
        </div>
        <ErrorBanner text={loadError} />

        {showAnalyzeForm && !analyzing && (
          <div style={{ marginTop: 18, paddingTop: 18, borderTop: "1px solid rgba(255,255,255,.1)" }}>
            <div style={{ color: "#fff", fontSize: 13, fontWeight: 700, marginBottom: 10 }}>
              بيانات مشتركة لجميع التقارير — {surveys.length} استبيان
            </div>
            <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
              <div style={{ flex: "1 1 200px" }}>
                <div style={{ color: "rgba(255,255,255,.6)", fontSize: 12, marginBottom: 6 }}>اسم المعد</div>
                <input value={preparedBy} onChange={e => setPreparedBy(e.target.value)}
                  style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: "1px solid rgba(255,255,255,.18)", background: "rgba(255,255,255,.06)", color: "#fff", fontFamily: "'Cairo',sans-serif", fontSize: 14 }} />
              </div>
              <div style={{ flex: "1 1 200px" }}>
                <div style={{ color: "rgba(255,255,255,.6)", fontSize: 12, marginBottom: 6 }}>المراجع</div>
                <input value={reviewer} onChange={e => setReviewer(e.target.value)}
                  style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: "1px solid rgba(255,255,255,.18)", background: "rgba(255,255,255,.06)", color: "#fff", fontFamily: "'Cairo',sans-serif", fontSize: 14 }} />
              </div>
            </div>
            <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
              <button className="btn btn-primary btn-sm" onClick={() => { setShowAnalyzeForm(false); runAnalyzeAll(); }}>▶ تشغيل التحليل</button>
              <button className="btn btn-ghost btn-sm" onClick={() => setShowAnalyzeForm(false)}>إلغاء</button>
            </div>
          </div>
        )}

        {(analyzing || analyzeProgress) && (
          <div style={{ marginTop: 18, paddingTop: 18, borderTop: "1px solid rgba(255,255,255,.1)" }}>
            {analyzing && (
              <ProgressBar
                done={analyzeProgress.filter(p => p.status === "done" || p.status === "error" || p.status === "skipped").length}
                total={analyzeProgress.length}
                label="جاري تحليل الاستبيانات..."
                tip={analyzeTip}
              />
            )}
            {analyzeProgress.map(p => (
              <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "6px 4px", fontSize: 13, color: "rgba(255,255,255,.8)" }}>
                <span>{p.status === "done" ? "✔" : p.status === "error" ? "✖" : p.status === "skipped" ? "⚠" : p.status === "active" ? "⏳" : "○"}</span>
                <span>{p.name}</span>
                {p.status === "skipped" && <span style={{ color: "rgba(255,255,255,.4)", fontSize: 11 }}>(نوع غير معروف)</span>}
              </div>
            ))}
          </div>
        )}
      </div>

      {!years?.length ? (
        <div className="card" style={{ padding: 24 }}>
          <EmptyState text="لا توجد سنوات دراسية بعد داخل مجلد الاستبيانات الرئيسي." />
        </div>
      ) : surveys === null ? (
        <div className="card" style={{ padding: 24 }}><Skeleton rows={4} /></div>
      ) : surveys.length === 0 ? (
        <div className="card" style={{ padding: 24 }}>
          <EmptyState text="لا توجد استبيانات لهذا الفصل بعد." />
        </div>
      ) : (
        <>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 14, marginBottom: 20 }}>
            <div className="ssg-stat-card">
              <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12 }}>إجمالي الاستبيانات</div>
              <div style={{ color: "#fff", fontSize: 26, fontWeight: 900 }}>{totalSurveys}</div>
            </div>
            <div className="ssg-stat-card">
              <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12 }}>إجمالي الردود</div>
              <div style={{ color: "#fff", fontSize: 26, fontWeight: 900 }}>{totalResponses}</div>
            </div>
            <div className="ssg-stat-card">
              <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12 }}>متوسط الردود لكل استبيان</div>
              <div style={{ color: "#fff", fontSize: 26, fontWeight: 900 }}>{avg}</div>
            </div>
            <div className="ssg-stat-card">
              <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12 }}>آخر رد مستلم</div>
              <div style={{ color: "#fff", fontSize: 20, fontWeight: 900 }}>{formatWhen(lastResponse)}</div>
            </div>
          </div>

          <div className="card" style={{ padding: 24, marginBottom: 20 }}>
            <div style={{ color: "#fff", fontSize: 15, fontWeight: 800, marginBottom: 16 }}>عدد الردود لكل استبيان</div>
            <BarChart data={surveys.map(s => ({ label: s.name, value: s.responses }))} />
          </div>

          <div className="card" style={{ padding: 24 }}>
            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 14 }}>
              <LinksExportButtons
                rows={surveys.map(r => ({ name: r.name, formUrl: r.formUrl, responsesUrl: editorResponsesUrl(r.id) }))}
                title={`استبيانات ${semester} ${year}`}
                filename={`روابط_استبيانات_${semester}_${yearFolderName(year)}.xlsx`}
                pushToast={pushToast}
              />
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ color: "rgba(255,255,255,.5)", textAlign: "right" }}>
                    <th style={{ padding: "8px 10px" }}>الاستبيان</th>
                    <th style={{ padding: "8px 10px" }}>النوع</th>
                    <th style={{ padding: "8px 10px" }}>الردود</th>
                    <th style={{ padding: "8px 10px" }}>آخر رد</th>
                    <th style={{ padding: "8px 10px" }}>إجراءات</th>
                  </tr>
                </thead>
                <tbody>
                  {surveys.map(r => (
                    <tr key={r.id} style={{ borderTop: "1px solid rgba(255,255,255,.08)", color: "#e8f0fe" }}>
                      <td style={{ padding: "8px 10px" }}>{r.name}{r.error && <span style={{ color: "#ff6b5b", fontSize: 11 }}> (تعذّر التحميل)</span>}</td>
                      <td style={{ padding: "8px 10px", color: "rgba(255,255,255,.6)" }}>{r.surveyType}</td>
                      <td style={{ padding: "8px 10px" }}>{r.responses}</td>
                      <td style={{ padding: "8px 10px" }}>{formatWhen(r.last)}</td>
                      <td style={{ padding: "8px 10px", display: "flex", gap: 10 }}>
                        <a href={r.formUrl} target="_blank" rel="noreferrer" style={{ color: "#1abc9c" }}>فتح النموذج</a>
                        <a href={editorResponsesUrl(r.id)} target="_blank" rel="noreferrer" style={{ color: "#1abc9c" }}>فتح الردود</a>
                        <button className="btn btn-ghost btn-sm" onClick={() => refreshOne(r.id)}>↻ تحديث</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

// ---------- root ----------
// ---------- Departments & Programs settings ----------
function DepartmentsView({ pushToast }) {
  const [departments, setDepartments] = useState(loadDepartments);
  const [newDeptName, setNewDeptName] = useState("");
  const [programDrafts, setProgramDrafts] = useState({}); // deptId -> current input text

  function persist(next, message) {
    setDepartments(next);
    saveDepartments(next);
    if (message) pushToast(message, "success");
  }

  function addDepartment() {
    const name = newDeptName.trim();
    if (!name) return;
    if (departments.some(d => d.name === name)) { pushToast("هذا القسم موجود بالفعل.", "error"); return; }
    persist([...departments, { id: `dept-${Date.now()}`, name, programs: [] }], "تم إضافة القسم.");
    setNewDeptName("");
  }

  function renameDepartment(id, name) {
    setDepartments(ds => ds.map(d => d.id === id ? { ...d, name } : d));
  }

  function commitRename(id) {
    saveDepartments(departments);
  }

  function removeDepartment(id) {
    persist(departments.filter(d => d.id !== id), "تم حذف القسم.");
  }

  function addProgram(deptId) {
    const text = (programDrafts[deptId] || "").trim();
    if (!text) return;
    persist(departments.map(d => d.id === deptId ? { ...d, programs: [...d.programs, text] } : d));
    setProgramDrafts(p => ({ ...p, [deptId]: "" }));
  }

  function removeProgram(deptId, program) {
    persist(departments.map(d => d.id === deptId ? { ...d, programs: d.programs.filter(p => p !== program) } : d));
  }

  return (
    <div>
      <div className="card" style={{ padding: 24, marginBottom: 20 }}>
        <div style={{ color: "#fff", fontSize: 13, marginBottom: 14 }}>
          الأقسام التي بلا برامج (كالمحاسبة والاقتصاد) تُنشئ نسخة واحدة عند تفعيل "نسخة لكل قسم/برنامج".
          الأقسام التي لها برامج (كتكنولوجيا الأعمال) تُنشئ نسخة لكل برنامج بدلاً من نسخة عامة للقسم.
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <input
            value={newDeptName} onChange={e => setNewDeptName(e.target.value)}
            placeholder="اسم قسم جديد" onKeyDown={e => e.key === "Enter" && addDepartment()}
            style={{ flex: 1, padding: "9px 14px", borderRadius: 10, border: "1px solid rgba(255,255,255,.18)", background: "rgba(255,255,255,.06)", color: "#fff", fontFamily: "'Cairo',sans-serif", fontSize: 13 }}
          />
          <button className="btn btn-primary btn-sm" onClick={addDepartment}>+ إضافة قسم</button>
        </div>
      </div>

      {departments.map(d => (
        <div key={d.id} className="card" style={{ padding: 20, marginBottom: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
            <input
              value={d.name} onChange={e => renameDepartment(d.id, e.target.value)} onBlur={() => commitRename(d.id)}
              style={{ flex: 1, padding: "8px 12px", borderRadius: 8, border: "1px solid rgba(255,255,255,.18)", background: "rgba(255,255,255,.06)", color: "#fff", fontFamily: "'Cairo',sans-serif", fontSize: 14, fontWeight: 700 }}
            />
            <button className="btn btn-danger btn-sm" onClick={() => removeDepartment(d.id)}>✕ حذف القسم</button>
          </div>

          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: d.programs.length ? 12 : 0 }}>
            {d.programs.map(p => (
              <span key={p} style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(255,255,255,.08)", borderRadius: 20, padding: "5px 6px 5px 12px", fontSize: 12, color: "#e8f0fe" }}>
                {p}
                <button onClick={() => removeProgram(d.id, p)} aria-label={`إزالة البرنامج ${p}`} title={`إزالة ${p}`} style={{ background: "rgba(231,76,60,.2)", border: "none", borderRadius: "50%", width: 18, height: 18, color: "#ff6b5b", cursor: "pointer", fontSize: 11, lineHeight: 1 }}>✕</button>
              </span>
            ))}
          </div>

          <div style={{ display: "flex", gap: 8 }}>
            <input
              value={programDrafts[d.id] || ""} onChange={e => setProgramDrafts(p => ({ ...p, [d.id]: e.target.value }))}
              placeholder="اسم برنامج جديد (اختياري)" onKeyDown={e => e.key === "Enter" && addProgram(d.id)}
              style={{ flex: 1, padding: "7px 12px", borderRadius: 8, border: "1px solid rgba(255,255,255,.14)", background: "rgba(255,255,255,.04)", color: "#fff", fontFamily: "'Cairo',sans-serif", fontSize: 12.5 }}
            />
            <button className="btn btn-ghost btn-sm" onClick={() => addProgram(d.id)}>+ إضافة برنامج</button>
          </div>
        </div>
      ))}
    </div>
  );
}

export default function SemesterSurveys() {
  const [tab, setTab] = useState("generate");
  const auth = useSemesterAuth();
  const { toasts, push, dismiss } = useToasts();
  const configured = !!TEMPLATE_FOLDER_ID && !!ROOT_SURVEYS_FOLDER_ID;

  return (
    <div>
      <style>{CSS}</style>
      <ToastStack toasts={toasts} onDismiss={dismiss} />

      <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
        <button className={`btn btn-sm ${tab === "generate" ? "btn-primary" : "btn-ghost"}`} onClick={() => setTab("generate")}>➕ إنشاء استبيانات</button>
        <button className={`btn btn-sm ${tab === "dashboard" ? "btn-primary" : "btn-ghost"}`} onClick={() => setTab("dashboard")}>📊 لوحة المتابعة</button>
        <button className={`btn btn-sm ${tab === "departments" ? "btn-primary" : "btn-ghost"}`} onClick={() => setTab("departments")}>⚙ الأقسام والبرامج</button>
      </div>

      {tab === "departments" ? (
        <DepartmentsView pushToast={push} />
      ) : !configured ? (
        <div className="card" style={{ padding: 24, color: "#ff6b5b", fontSize: 13 }}>
          ⚠ لم يتم ضبط معرفات مجلدات Google Drive. يرجى إضافة <code>VITE_GOOGLE_TEMPLATE_FOLDER_ID</code> و<code>VITE_GOOGLE_ROOT_SURVEYS_FOLDER_ID</code> إلى ملف <code>.env.local</code> وإعادة تشغيل الخادم.
        </div>
      ) : !auth.token ? (
        <div className="card" style={{ padding: 32, textAlign: "center" }}>
          <div style={{ fontSize: 34, marginBottom: 10 }}>🔐</div>
          <div style={{ color: "#fff", fontSize: 15, fontWeight: 800, marginBottom: 16 }}>الاستبيانات الفصلية تحتاج الاتصال بحساب Google</div>
          <button className="btn btn-primary" disabled={auth.connecting} onClick={auth.connect}>
            {auth.connecting ? "⏳ جاري الاتصال..." : "الاتصال بـ Google"}
          </button>
          <ErrorBanner text={auth.authError} />
        </div>
      ) : tab === "generate" ? (
        <GenerateSurveysView token={auth.token} pushToast={push} />
      ) : (
        <DashboardView token={auth.token} pushToast={push} />
      )}
    </div>
  );
}
