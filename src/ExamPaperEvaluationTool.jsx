import { useMemo, useRef, useState } from "react";
import { GoogleDriveIcon, InlineNotice, QualityIcon } from "./UiElements.jsx";
import {
  buildExamPaperReportDocx,
  examPaperReportFilename,
  summarizeExamPaperEvaluations,
} from "./engine/examPaperReport.js";
import {
  ROOT_SURVEYS_FOLDER_ID,
  getOrCreateFolder,
  yearFolderName,
} from "./engine/semesterSurveyModel.js";
import {
  EXAM_PAPER_DEPARTMENTS,
  buildExamPaperCoursesTemplate,
  examPaperCourseKey,
  readExamPaperCoursesFile,
} from "./engine/examPaperTemplate.js";

const ISSUE_SUGGESTIONS = [
  "اسم القسم",
  "اسم البرنامج",
  "الفرقة",
  "الفصل الدراسي",
  "تاريخ الامتحان",
  "زمن الامتحان",
  "الدرجة الكلية وتوزيع الدرجات",
  "ترقيم الصفحات",
  "عبارة انتهاء الأسئلة",
  "توقيع لجنة الممتحنين",
  "وضوح الطباعة وتنسيق الورقة",
];

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1200);
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, character => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", "\"": "&quot;",
  })[character]);
}

function reportPrintHtml(report) {
  const summary = summarizeExamPaperEvaluations(report.courses);
  const tableRows = (rows, columns) => rows.length
    ? rows.map((course, index) => `<tr>${columns.map(column => `<td>${escapeHtml(column(course, index))}</td>`).join("")}</tr>`).join("")
    : `<tr><td colspan="${columns.length}">لا توجد مقررات في هذا القسم</td></tr>`;
  return `<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>${escapeHtml(examPaperReportFilename(report).replace(/\.docx$/i, ""))}</title><style>
    @page{size:A4;margin:17mm 14mm}body{font-family:Arial,Tahoma,sans-serif;color:#111;line-height:1.55;margin:0}header{text-align:center;margin-bottom:24px}header img{width:70px;height:70px;object-fit:contain}header b,header span{display:block}h1{text-align:center;font-size:20px;margin:20px 0 26px}h2{font-size:15px;margin:22px 0 8px;page-break-after:avoid}table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px}th,td{border:1px solid #9aa9b5;padding:7px 6px;text-align:center;vertical-align:middle}th{background:#1f4e78;color:#fff}tbody tr:nth-child(even){background:#f4f8fc}.signature{text-align:center;margin-top:28px}.footer{font-size:9px;color:#555;border-top:1px solid #bbb;margin-top:28px;padding-top:7px}@media print{button{display:none}thead{display:table-header-group}tr{page-break-inside:avoid}}</style></head><body>
    <header><img src="/logo.png"><b>كلية الإدارة والاقتصاد وتكنولوجيا الأعمال</b><span>لجنة القياس والتقويم</span></header>
    <h1>تقرير تقييم الورقة الامتحانية من حيث الشكل لقسم ${escapeHtml(report.department)} للفصل الدراسي ${escapeHtml(report.semester)} ${escapeHtml(report.year)}</h1>
    <h2>أولًا: بيان بعدد ونسبة مقررات القسم التي تم تقييم الورقة الامتحانية لها من حيث الشكل</h2>
    <table><thead><tr><th>مقررات القسم</th><th>العدد</th><th>النسبة</th></tr></thead><tbody><tr><td>المقررات المستوفاة لجميع العناصر</td><td>${summary.completed.length}</td><td>${summary.completedPct}</td></tr><tr><td>المقررات غير المستوفاة لبعض العناصر</td><td>${summary.incomplete.length}</td><td>${summary.incompletePct}</td></tr><tr><td>الإجمالي</td><td>${summary.total}</td><td>100%</td></tr></tbody></table>
    <h2>ثانيًا: المقررات المستوفاة لجميع عناصر تقييم الورقة الامتحانية من حيث الشكل</h2>
    <table><thead><tr><th>م</th><th>كود المقرر</th><th>المقرر</th><th>تقييم مخرجات التعلم</th></tr></thead><tbody>${tableRows(summary.completed, [(c,i)=>i+1,c=>c.code,c=>c.name,c=>`${c.learningOutcomes}%`])}</tbody></table>
    <h2>ثالثًا: المقررات غير المستوفاة لبعض عناصر تقييم الورقة الامتحانية من حيث الشكل</h2>
    <table><thead><tr><th>م</th><th>كود المقرر</th><th>المقرر</th><th>العناصر غير المستوفاة</th><th>تقييم مخرجات التعلم</th></tr></thead><tbody>${tableRows(summary.incomplete, [(c,i)=>i+1,c=>c.code,c=>c.name,c=>c.missingItems,c=>`${c.learningOutcomes}%`])}</tbody></table>
    <h2>القائمون بالتقييم</h2><table><thead><tr><th>م</th><th>الاسم</th><th>الوظيفة</th></tr></thead><tbody>${report.evaluators.map((evaluator,index)=>`<tr><td>${index+1}</td><td>${escapeHtml(evaluator.name)}</td><td>${escapeHtml(evaluator.role)}</td></tr>`).join("")}</tbody></table>
    <div class="signature"><b>رئيس لجنة القياس والتقويم</b><div>${escapeHtml(report.committeeHead || report.evaluators[0]?.name || "")}</div></div>
    <div class="footer">رؤية الكلية ورسالتها: تقديم تعليم متميز يواكب التطورات المعرفية والتقنية، ودعم البحث العلمي وخدمة المجتمع في إطار القيم الأخلاقية والمسؤولية المهنية.</div>
    <script>window.addEventListener('load',()=>setTimeout(()=>window.print(),250));<\/script></body></html>`;
}

async function uploadBlobToFolder(token, folderId, blob, filename) {
  const metadata = { name: filename, mimeType: blob.type || "application/vnd.openxmlformats-officedocument.wordprocessingml.document", parents: [folderId] };
  const body = new FormData();
  body.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json; charset=UTF-8" }));
  body.append("file", blob, filename);
  const response = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}` },
    body,
  });
  const data = await response.json();
  if (!response.ok || data.error) throw new Error(data.error?.message || `فشل رفع التقرير (${response.status}).`);
  return data;
}

export default function ExamPaperEvaluationTool({ googleAuth = {}, reportSettings = {} }) {
  const inputRef = useRef(null);
  const [courses, setCourses] = useState([]);
  const [evaluations, setEvaluations] = useState({});
  const [department, setDepartment] = useState("");
  const [semester, setSemester] = useState("ربيع");
  const [year, setYear] = useState("2025-2026");
  const [sourceName, setSourceName] = useState("");
  const [evaluators, setEvaluators] = useState([
    { name: "", role: "رئيس لجنة القياس والتقويم" },
    { name: "", role: "عضو لجنة القياس والتقويم" },
  ]);
  const [committeeHead, setCommitteeHead] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState("");
  const [driveFile, setDriveFile] = useState(null);

  const fileDepartments = useMemo(() => [...new Set(courses.map(course => course.department))], [courses]);
  const visibleCourses = useMemo(() => courses.filter(course => course.department === department), [courses, department]);
  const evaluatedCourses = useMemo(() => visibleCourses.map(course => ({
    ...course,
    status: evaluations[examPaperCourseKey(course)]?.status || "",
    learningOutcomes: evaluations[examPaperCourseKey(course)]?.learningOutcomes ?? "",
    missingItems: evaluations[examPaperCourseKey(course)]?.missingItems || "",
  })), [visibleCourses, evaluations]);
  const progress = useMemo(() => {
    if (!evaluatedCourses.length) return 0;
    const ready = evaluatedCourses.filter(course => course.status && course.learningOutcomes !== "" && (course.status !== "incomplete" || course.missingItems.trim())).length;
    return Math.round(ready * 100 / evaluatedCourses.length);
  }, [evaluatedCourses]);
  const summary = useMemo(() => summarizeExamPaperEvaluations(evaluatedCourses), [evaluatedCourses]);

  const updateEvaluation = (course, patch) => {
    const key = examPaperCourseKey(course);
    setEvaluations(current => ({ ...current, [key]: { ...current[key], ...patch } }));
    setDriveFile(null);
  };

  const handleTemplateDownload = () => {
    downloadBlob(buildExamPaperCoursesTemplate(), "قالب مقررات تقييم الورقة الامتحانية.xlsx");
    setMessage("تم تحميل قالب Excel. املأ الأعمدة الثلاثة ثم ارفعه هنا.");
    setError("");
  };

  const handleFile = async event => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    setBusy("file"); setError(""); setMessage(""); setDriveFile(null);
    try {
      const parsed = await readExamPaperCoursesFile(file);
      setCourses(parsed);
      setEvaluations({});
      setSourceName(file.name);
      const departments = [...new Set(parsed.map(course => course.department))];
      setDepartment(departments[0] || "");
      const unknown = departments.filter(item => !EXAM_PAPER_DEPARTMENTS.includes(item));
      setMessage(unknown.length
        ? `تم تحميل ${parsed.length} مقرر. راجع أسماء الأقسام غير القياسية: ${unknown.join("، ")}.`
        : `تم تحميل ${parsed.length} مقرر موزعة على ${departments.length} قسم.`);
    } catch (uploadError) {
      setError(uploadError.message);
    } finally {
      setBusy("");
    }
  };

  const buildReportPayload = () => {
    if (!courses.length) throw new Error("حمّل ملف المقررات أولًا.");
    if (!department || !visibleCourses.length) throw new Error("اختر قسمًا يحتوي على مقررات.");
    if (!semester.trim() || !year.trim()) throw new Error("أدخل الفصل الدراسي والسنة.");
    for (const course of evaluatedCourses) {
      if (!course.status) throw new Error(`حدّد حالة الاستيفاء للمقرر ${course.name}.`);
      const outcome = Number(course.learningOutcomes);
      if (course.learningOutcomes === "" || !Number.isFinite(outcome) || outcome < 0 || outcome > 100) {
        throw new Error(`أدخل نسبة مخرجات تعلم صحيحة من 0 إلى 100 للمقرر ${course.name}.`);
      }
      if (course.status === "incomplete" && !course.missingItems.trim()) {
        throw new Error(`اكتب العناصر غير المستوفاة للمقرر ${course.name}.`);
      }
    }
    const cleanEvaluators = evaluators.map(evaluator => ({ name: evaluator.name.trim(), role: evaluator.role.trim() })).filter(evaluator => evaluator.name || evaluator.role);
    if (!cleanEvaluators.length || cleanEvaluators.some(evaluator => !evaluator.name || !evaluator.role)) {
      throw new Error("أكمل اسم ووظيفة كل قائم بالتقييم.");
    }
    return {
      department,
      semester: semester.trim(),
      year: year.trim(),
      courses: evaluatedCourses.map(course => ({ ...course, learningOutcomes: Number(course.learningOutcomes) })),
      evaluators: cleanEvaluators,
      committeeHead: committeeHead.trim(),
    };
  };

  const handleDownload = async () => {
    setBusy("download"); setError(""); setMessage("");
    try {
      const report = buildReportPayload();
      const blob = await buildExamPaperReportDocx(report, reportSettings);
      downloadBlob(blob, examPaperReportFilename(report));
      setMessage("تم إنشاء وتحميل تقرير Word الجاهز للطباعة.");
    } catch (actionError) {
      setError(actionError.message);
    } finally { setBusy(""); }
  };

  const handlePrint = () => {
    setError(""); setMessage("");
    try {
      const report = buildReportPayload();
      const printWindow = window.open("", "_blank", "noopener,noreferrer");
      if (!printWindow) throw new Error("اسمح بالنوافذ المنبثقة لفتح شاشة الطباعة.");
      printWindow.document.open();
      printWindow.document.write(reportPrintHtml(report));
      printWindow.document.close();
    } catch (actionError) { setError(actionError.message); }
  };

  const handleDriveUpload = async () => {
    if (!googleAuth.token) {
      setMessage("اربط حساب Google Drive أولًا، ثم اضغط زر الرفع مرة أخرى.");
      googleAuth.connect?.();
      return;
    }
    setBusy("drive"); setError(""); setMessage(""); setDriveFile(null);
    try {
      const report = buildReportPayload();
      const blob = await buildExamPaperReportDocx(report, reportSettings);
      const yearFolder = await getOrCreateFolder(googleAuth.token, ROOT_SURVEYS_FOLDER_ID, yearFolderName(report.year));
      const semesterFolder = await getOrCreateFolder(googleAuth.token, yearFolder.id, report.semester);
      const departmentFolder = await getOrCreateFolder(googleAuth.token, semesterFolder.id, report.department);
      const reportFolder = await getOrCreateFolder(
        googleAuth.token,
        departmentFolder.id,
        `تقرير تقييم الورقة الامتحانية من حيث الشكل - ${report.semester} ${report.year}`,
      );
      const uploaded = await uploadBlobToFolder(googleAuth.token, reportFolder.id, blob, examPaperReportFilename(report));
      setDriveFile(uploaded);
      setMessage("تم رفع التقرير داخل مجلد القسم على Google Drive.");
    } catch (actionError) {
      setError(actionError.message);
    } finally { setBusy(""); }
  };

  return <section className="exam-paper-tool">
    <style>{EXAM_PAPER_CSS}</style>
    <header className="exam-paper-hero">
      <div><span>أداة تقارير لجنة القياس والتقويم</span><h2>تقييم الورقة الامتحانية من حيث الشكل</h2><p>حمّل قالب المقررات، ثم قيّم مقررات كل قسم وأنشئ تقريرًا بنفس تفاصيل النموذج المعتمد.</p></div>
      <div className="exam-paper-progress"><b>{progress}%</b><small>اكتمال تقييم القسم</small></div>
    </header>

    {error && <InlineNotice style={{ marginBottom: 14 }}>{error}</InlineNotice>}
    {message && <div className="exam-paper-success" role="status">{message}{driveFile?.webViewLink && <> <a href={driveFile.webViewLink} target="_blank" rel="noreferrer">فتح التقرير على Drive</a></>}</div>}

    <div className="exam-paper-steps">
      <article className="exam-paper-panel">
        <div className="exam-paper-panel-head"><span>1</span><div><h3>قالب المقررات</h3><p>الملف الرئيسي يحتوي على ثلاثة أعمدة فقط.</p></div></div>
        <div className="exam-paper-template-actions">
          <button className="btn btn-blue btn-sm" type="button" onClick={handleTemplateDownload}><QualityIcon name="download" size={15}/> تحميل قالب Excel</button>
          <button className="btn btn-primary btn-sm" type="button" disabled={busy === "file"} onClick={() => inputRef.current?.click()}><QualityIcon name="upload" size={15}/> {busy === "file" ? "جاري القراءة…" : "رفع القالب بعد تعبئته"}</button>
          <input ref={inputRef} type="file" accept=".xlsx,.xls" hidden onChange={handleFile}/>
        </div>
        {sourceName && <div className="exam-paper-source"><QualityIcon name="sheet" size={15}/><span>{sourceName}</span><b>{courses.length} مقرر</b></div>}
      </article>

      <article className="exam-paper-panel">
        <div className="exam-paper-panel-head"><span>2</span><div><h3>بيانات التقرير والقسم</h3><p>اختر القسم ليظهر فقط ما يخصه من مقررات.</p></div></div>
        <div className="exam-paper-meta-grid">
          <label><span>القسم</span><select className="input" value={department} onChange={event => { setDepartment(event.target.value); setDriveFile(null); }} disabled={!courses.length}><option value="">اختر القسم</option>{fileDepartments.map(item => <option key={item} value={item}>{item}</option>)}</select></label>
          <label><span>الفصل الدراسي</span><select className="input" value={semester} onChange={event => setSemester(event.target.value)}><option>خريف</option><option>ربيع</option><option>صيف</option></select></label>
          <label><span>السنة الدراسية</span><input className="input" value={year} onChange={event => setYear(event.target.value)} placeholder="2025-2026"/></label>
        </div>
      </article>
    </div>

    <article className="exam-paper-panel exam-paper-courses-panel">
      <div className="exam-paper-panel-head"><span>3</span><div><h3>تقييم مقررات القسم</h3><p>{department ? `${department} — ${visibleCourses.length} مقرر` : "اختر القسم لبدء التقييم"}</p></div></div>
      {!courses.length ? <div className="exam-paper-empty"><QualityIcon name="sheet" size={27}/><b>ابدأ برفع قالب المقررات</b><span>ستظهر المقررات هنا تلقائيًا بعد قراءة الملف.</span></div>
        : !department ? <div className="exam-paper-empty"><QualityIcon name="folder" size={27}/><b>اختر قسمًا</b><span>يتم عرض مقررات القسم المختار فقط.</span></div>
        : <div className="exam-paper-course-list">{evaluatedCourses.map((course, index) => {
          const value = evaluations[examPaperCourseKey(course)] || {};
          return <div className={`exam-paper-course ${value.status || "pending"}`} key={examPaperCourseKey(course)}>
            <div className="exam-paper-course-id"><span>{index + 1}</span><div><b>{course.name}</b><small>{course.code}</small></div></div>
            <div className="exam-paper-status" role="group" aria-label={`حالة ${course.name}`}>
              <button type="button" className={value.status === "complete" ? "selected complete" : ""} onClick={() => updateEvaluation(course, { status: "complete", missingItems: "" })}>مستوفٍ لجميع العناصر</button>
              <button type="button" className={value.status === "incomplete" ? "selected incomplete" : ""} onClick={() => updateEvaluation(course, { status: "incomplete" })}>غير مستوفٍ لبعض العناصر</button>
            </div>
            <label className="exam-paper-outcome"><span>تقييم مخرجات التعلم</span><div><input className="input" type="number" min="0" max="100" step="0.1" value={value.learningOutcomes ?? ""} onChange={event => updateEvaluation(course, { learningOutcomes: event.target.value })}/><i>%</i></div></label>
            {value.status === "incomplete" && <label className="exam-paper-missing"><span>العناصر غير المستوفاة</span><textarea className="textarea" value={value.missingItems || ""} onChange={event => updateEvaluation(course, { missingItems: event.target.value })} placeholder="مثال: اسم البرنامج - الفرقة - الفصل الدراسي"/><div className="exam-paper-issue-options">{ISSUE_SUGGESTIONS.map(issue => <button type="button" key={issue} onClick={() => { const current = (value.missingItems || "").trim(); if (!current.includes(issue)) updateEvaluation(course, { missingItems: current ? `${current} - ${issue}` : issue }); }}>{issue}</button>)}</div></label>}
          </div>;
        })}</div>}
    </article>

    <div className="exam-paper-steps bottom">
      <article className="exam-paper-panel">
        <div className="exam-paper-panel-head"><span>4</span><div><h3>القائمون بالتقييم</h3><p>تظهر الأسماء والوظائف في نهاية التقرير.</p></div></div>
        <div className="exam-paper-evaluators">{evaluators.map((evaluator, index) => <div key={index} className="exam-paper-evaluator-row"><input className="input" value={evaluator.name} onChange={event => setEvaluators(current => current.map((item, i) => i === index ? { ...item, name: event.target.value } : item))} placeholder="الاسم"/><input className="input" value={evaluator.role} onChange={event => setEvaluators(current => current.map((item, i) => i === index ? { ...item, role: event.target.value } : item))} placeholder="الوظيفة"/>{evaluators.length > 1 && <button type="button" onClick={() => setEvaluators(current => current.filter((_, i) => i !== index))} aria-label="حذف المقيم">×</button>}</div>)}</div>
        <button className="exam-paper-add" type="button" onClick={() => setEvaluators(current => [...current, { name: "", role: "عضو لجنة القياس والتقويم" }])}>+ إضافة قائم بالتقييم</button>
        <label className="exam-paper-head-name"><span>اسم رئيس لجنة القياس والتقويم للتوقيع</span><input className="input" value={committeeHead} onChange={event => setCommitteeHead(event.target.value)} placeholder="يُستخدم اسم أول مقيم تلقائيًا إذا تركته فارغًا"/></label>
      </article>

      <article className="exam-paper-panel exam-paper-summary">
        <div className="exam-paper-panel-head"><span>5</span><div><h3>التقرير النهائي</h3><p>راجع الملخص ثم اختر طريقة التسليم.</p></div></div>
        <div className="exam-paper-stats"><div><b>{summary.total}</b><span>إجمالي المقررات</span></div><div className="ok"><b>{summary.completed.length}</b><span>مستوفاة</span></div><div className="warn"><b>{summary.incomplete.length}</b><span>غير مستوفاة</span></div></div>
        <div className="exam-paper-delivery">
          <button className="btn btn-primary" type="button" disabled={Boolean(busy)} onClick={handleDownload}><QualityIcon name="download" size={17}/>{busy === "download" ? "جاري إنشاء Word…" : "تحميل التقرير Word"}</button>
          <button className="btn btn-ghost" type="button" disabled={Boolean(busy)} onClick={handlePrint}><QualityIcon name="document" size={17}/> طباعة أو حفظ PDF</button>
          <button className="btn btn-blue" type="button" disabled={Boolean(busy) || googleAuth.connecting} onClick={handleDriveUpload}><GoogleDriveIcon size={18}/>{busy === "drive" ? "جاري الرفع…" : googleAuth.token ? "رفع على Google Drive" : "ربط Drive ثم الرفع"}</button>
        </div>
        <p className="exam-paper-folder-note">مسار Drive: السنة / الفصل / القسم / تقرير تقييم الورقة الامتحانية من حيث الشكل للفصل الدراسي والسنة.</p>
      </article>
    </div>
  </section>;
}

const EXAM_PAPER_CSS = `
.exam-paper-tool{direction:rtl;color:#edf7ff}.exam-paper-hero{display:flex;align-items:center;justify-content:space-between;gap:20px;padding:25px 28px;margin-bottom:14px;border-radius:21px;border:1px solid rgba(94,234,212,.17);background:linear-gradient(120deg,rgba(15,74,84,.66),rgba(22,51,83,.64))}.exam-paper-hero span{color:#74e8d2;font-size:10px;font-weight:900}.exam-paper-hero h2{font-size:25px;margin:7px 0 4px}.exam-paper-hero p{margin:0;color:rgba(235,247,255,.53);font-size:11.5px}.exam-paper-progress{flex:0 0 120px;display:grid;place-items:center;padding:15px;border-radius:17px;background:rgba(3,18,31,.32);border:1px solid rgba(255,255,255,.09)}.exam-paper-progress b{font-size:27px;color:#7cead5}.exam-paper-progress small{color:rgba(255,255,255,.42);font-size:8.5px}.exam-paper-success{margin-bottom:14px;padding:11px 14px;border-radius:11px;color:#baf7df;background:rgba(34,197,94,.1);border:1px solid rgba(74,222,128,.18);font-size:11px}.exam-paper-success a{color:#78d8ff;font-weight:800;margin-right:8px}.exam-paper-steps{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}.exam-paper-steps.bottom{align-items:stretch}.exam-paper-panel{padding:19px;border-radius:18px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.085)}.exam-paper-panel-head{display:flex;align-items:center;gap:10px;margin-bottom:15px}.exam-paper-panel-head>span{width:34px;height:34px;display:grid;place-items:center;flex:0 0 34px;border-radius:10px;color:#eafffa;background:linear-gradient(135deg,#1abc9c,#2374a5);font-weight:900;font-size:12px}.exam-paper-panel-head h3{font-size:14px;margin:0}.exam-paper-panel-head p{font-size:9.5px;color:rgba(255,255,255,.42);margin:2px 0 0}.exam-paper-template-actions{display:flex;gap:8px;flex-wrap:wrap}.exam-paper-source{display:flex;align-items:center;gap:8px;margin-top:12px;padding:9px 11px;border-radius:10px;background:rgba(2,15,27,.23);color:rgba(255,255,255,.62);font-size:10px}.exam-paper-source span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.exam-paper-source b{margin-right:auto;color:#76e5d0}.exam-paper-meta-grid{display:grid;grid-template-columns:1.4fr .8fr .9fr;gap:10px}.exam-paper-meta-grid label>span,.exam-paper-outcome>span,.exam-paper-missing>span,.exam-paper-head-name>span{display:block;margin-bottom:5px;color:rgba(221,238,250,.64);font-size:9.5px;font-weight:800}.exam-paper-courses-panel{margin-bottom:12px}.exam-paper-empty{min-height:160px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;border:1px dashed rgba(255,255,255,.14);border-radius:13px;color:rgba(255,255,255,.34)}.exam-paper-empty b{color:rgba(255,255,255,.62);font-size:12px}.exam-paper-empty span{font-size:9.5px}.exam-paper-course-list{display:grid;gap:9px}.exam-paper-course{display:grid;grid-template-columns:minmax(210px,1.2fr) minmax(250px,1.3fr) 170px;gap:11px;align-items:center;padding:12px;border-radius:13px;background:rgba(2,16,29,.27);border:1px solid rgba(255,255,255,.075);transition:.18s}.exam-paper-course.complete{border-color:rgba(52,211,153,.2)}.exam-paper-course.incomplete{border-color:rgba(251,191,36,.23)}.exam-paper-course-id{display:flex;align-items:center;gap:9px;min-width:0}.exam-paper-course-id>span{width:27px;height:27px;display:grid;place-items:center;flex:0 0 27px;border-radius:8px;color:#72e6d0;background:rgba(26,188,156,.1);font-size:9px;font-weight:900}.exam-paper-course-id div{min-width:0}.exam-paper-course-id b,.exam-paper-course-id small{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.exam-paper-course-id b{font-size:10.5px}.exam-paper-course-id small{margin-top:2px;color:#7fc8ee;font-size:9px}.exam-paper-status{display:grid;grid-template-columns:1fr 1fr;gap:6px}.exam-paper-status button{padding:8px 7px;border-radius:8px;border:1px solid rgba(255,255,255,.11);background:rgba(255,255,255,.035);color:rgba(255,255,255,.46);font:700 9px 'Cairo',sans-serif;cursor:pointer}.exam-paper-status button.selected.complete{color:#c5fae5;background:rgba(34,197,94,.13);border-color:rgba(74,222,128,.32)}.exam-paper-status button.selected.incomplete{color:#ffebad;background:rgba(245,158,11,.13);border-color:rgba(251,191,36,.32)}.exam-paper-outcome div{display:flex;align-items:center;gap:6px}.exam-paper-outcome .input{direction:ltr;text-align:center;padding:8px}.exam-paper-outcome i{font-style:normal;color:#77e6d1;font-size:11px}.exam-paper-missing{grid-column:1/-1}.exam-paper-missing .textarea{min-height:62px}.exam-paper-issue-options{display:flex;flex-wrap:wrap;gap:5px;margin-top:6px}.exam-paper-issue-options button{padding:4px 8px;border:1px solid rgba(255,255,255,.1);border-radius:999px;background:rgba(255,255,255,.04);color:rgba(255,255,255,.48);font:700 8px 'Cairo';cursor:pointer}.exam-paper-issue-options button:hover{color:#d7fff7;border-color:rgba(94,234,212,.25);background:rgba(26,188,156,.08)}.exam-paper-evaluators{display:grid;gap:7px}.exam-paper-evaluator-row{display:grid;grid-template-columns:1fr 1.3fr 28px;gap:7px}.exam-paper-evaluator-row .input{padding:9px 11px;font-size:11px}.exam-paper-evaluator-row button{border:0;border-radius:8px;background:rgba(239,68,68,.11);color:#fca5a5;cursor:pointer;font-size:17px}.exam-paper-add{margin-top:9px;border:0;background:transparent;color:#71e5d0;font:800 9.5px 'Cairo';cursor:pointer}.exam-paper-head-name{display:block;margin-top:13px}.exam-paper-head-name .input{font-size:11px}.exam-paper-summary{display:flex;flex-direction:column}.exam-paper-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:7px}.exam-paper-stats div{padding:11px;text-align:center;border-radius:10px;background:rgba(255,255,255,.045);border:1px solid rgba(255,255,255,.07)}.exam-paper-stats b,.exam-paper-stats span{display:block}.exam-paper-stats b{font-size:20px}.exam-paper-stats span{color:rgba(255,255,255,.4);font-size:8.5px}.exam-paper-stats .ok b{color:#6ee7b7}.exam-paper-stats .warn b{color:#fcd34d}.exam-paper-delivery{display:grid;grid-template-columns:1fr;gap:7px;margin-top:13px}.exam-paper-delivery .btn{justify-content:center;padding:9px 14px;font-size:10px}.exam-paper-folder-note{margin:auto 0 0;padding-top:11px;color:rgba(255,255,255,.3);font-size:8.5px;line-height:1.7}@media(max-width:900px){.exam-paper-course{grid-template-columns:1fr 1fr}.exam-paper-outcome{grid-column:1/-1}.exam-paper-meta-grid{grid-template-columns:1fr 1fr}.exam-paper-meta-grid label:first-child{grid-column:1/-1}}@media(max-width:650px){.exam-paper-hero{padding:19px}.exam-paper-progress{display:none}.exam-paper-steps{grid-template-columns:1fr}.exam-paper-meta-grid,.exam-paper-course{grid-template-columns:1fr}.exam-paper-meta-grid label:first-child,.exam-paper-outcome,.exam-paper-missing{grid-column:auto}.exam-paper-status{grid-template-columns:1fr}.exam-paper-evaluator-row{grid-template-columns:1fr}.exam-paper-evaluator-row button{min-height:30px}}
`;
