import { useMemo, useRef, useState } from "react";
import JSZip from "jszip";
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
  examPaperCourseIssues,
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
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

function ExamPaperLoading({ operation }) {
  if (!operation) return null;
  const total = Math.max(operation.total || 1, 1);
  const current = Math.min(operation.current || 0, total);
  const percent = Math.round(current * 100 / total);
  return <div className="exam-paper-loading" role="status" aria-live="polite" aria-busy="true">
    <div className="exam-paper-loading-card">
      <div className="exam-paper-loading-spinner"><i/><i/></div>
      <h3>{operation.title}</h3>
      <p>{operation.detail || "يرجى الانتظار…"}</p>
      <div className="exam-paper-loading-track"><span style={{ width: `${percent}%` }}/></div>
      <div className="exam-paper-loading-count"><b>{percent}%</b><span>{current} من {total}</span></div>
    </div>
  </div>;
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
    <table><thead><tr><th>م</th><th>كود المقرر</th><th>المقرر</th><th>تقييم مخرجات التعلم</th></tr></thead><tbody>${tableRows(summary.completed, [(c,i)=>i+1,c=>c.code,c=>c.name,c=>c.learningOutcomes === "" ? "" : `${c.learningOutcomes}%`])}</tbody></table>
    <h2>ثالثًا: المقررات غير المستوفاة لبعض عناصر تقييم الورقة الامتحانية من حيث الشكل</h2>
    <table><thead><tr><th>م</th><th>كود المقرر</th><th>المقرر</th><th>العناصر غير المستوفاة</th><th>تقييم مخرجات التعلم</th></tr></thead><tbody>${tableRows(summary.incomplete, [(c,i)=>i+1,c=>c.code,c=>c.name,c=>c.missingItems,c=>c.learningOutcomes === "" ? "" : `${c.learningOutcomes}%`])}</tbody></table>
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
  const [operation, setOperation] = useState(null);
  const [blankOutcomePrompt, setBlankOutcomePrompt] = useState(null);
  const [driveFile, setDriveFile] = useState(null);

  const fileDepartments = useMemo(() => [...new Set(courses.map(course => course.department))], [courses]);
  const visibleCourses = useMemo(() => courses.filter(course => course.department === department), [courses, department]);
  const evaluatedCourses = useMemo(() => visibleCourses.map(course => ({
    ...course,
    status: evaluations[examPaperCourseKey(course)]?.status ?? course.status ?? "",
    learningOutcomes: evaluations[examPaperCourseKey(course)]?.learningOutcomes ?? course.learningOutcomes ?? "",
    missingItems: evaluations[examPaperCourseKey(course)]?.missingItems ?? course.missingItems ?? "",
  })), [visibleCourses, evaluations]);
  const allEvaluatedCourses = useMemo(() => courses.map(course => ({
    ...course,
    status: evaluations[examPaperCourseKey(course)]?.status ?? course.status ?? "",
    learningOutcomes: evaluations[examPaperCourseKey(course)]?.learningOutcomes ?? course.learningOutcomes ?? "",
    missingItems: evaluations[examPaperCourseKey(course)]?.missingItems ?? course.missingItems ?? "",
  })), [courses, evaluations]);
  const departmentSummaries = useMemo(() => fileDepartments.map(name => {
    const departmentCourses = allEvaluatedCourses.filter(course => course.department === name);
    const issueCount = departmentCourses.reduce((total, course) => total + (examPaperCourseIssues(course).length ? 1 : 0), 0);
    return { name, total: departmentCourses.length, issueCount, ready: issueCount === 0 };
  }), [allEvaluatedCourses, fileDepartments]);
  const progress = useMemo(() => {
    if (!evaluatedCourses.length) return 0;
    const ready = evaluatedCourses.filter(course => !examPaperCourseIssues(course).length).length;
    return Math.round(ready * 100 / evaluatedCourses.length);
  }, [evaluatedCourses]);
  const overallProgress = useMemo(() => {
    if (!allEvaluatedCourses.length) return 0;
    const ready = allEvaluatedCourses.filter(course => !examPaperCourseIssues(course).length).length;
    return Math.round(ready * 100 / allEvaluatedCourses.length);
  }, [allEvaluatedCourses]);
  const summary = useMemo(() => summarizeExamPaperEvaluations(evaluatedCourses), [evaluatedCourses]);

  const updateEvaluation = (course, patch) => {
    const key = examPaperCourseKey(course);
    setEvaluations(current => ({ ...current, [key]: { ...current[key], ...patch } }));
    setDriveFile(null);
  };

  const handleTemplateDownload = async () => {
    setBusy("template");
    setOperation({ title: "تجهيز قالب Excel", detail: "تحميل القالب المعتمد…", current: 1, total: 2 });
    try {
      const response = await fetch("/course_codes_and_names.xlsx");
      if (!response.ok) throw new Error();
      setOperation({ title: "تجهيز قالب Excel", detail: "بدء تنزيل القالب…", current: 2, total: 2 });
      downloadBlob(await response.blob(), "قالب تقييم الورقة الامتحانية.xlsx");
    } catch {
      setOperation({ title: "تجهيز قالب Excel", detail: "إنشاء نسخة بديلة من القالب…", current: 2, total: 2 });
      downloadBlob(buildExamPaperCoursesTemplate(), "قالب تقييم الورقة الامتحانية.xlsx");
    }
    setMessage("تم إرسال قالب Excel إلى التنزيلات في المتصفح.");
    setError("");
    setBusy(""); setOperation(null);
  };

  const handleFile = async event => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    setBusy("file"); setError(""); setMessage(""); setDriveFile(null);
    setOperation({ title: "قراءة ملف التقييم", detail: `فتح ${file.name}…`, current: 1, total: 3 });
    try {
      const parsed = await readExamPaperCoursesFile(file);
      setOperation({ title: "قراءة ملف التقييم", detail: "تقسيم المقررات حسب الأقسام…", current: 2, total: 3 });
      setCourses(parsed);
      setEvaluations(Object.fromEntries(parsed.map(course => [examPaperCourseKey(course), {
        status: course.status,
        learningOutcomes: course.learningOutcomes,
        missingItems: course.missingItems,
      }])));
      setSourceName(file.name);
      const departments = [...new Set(parsed.map(course => course.department))];
      setDepartment(departments[0] || "");
      const unknown = departments.filter(item => !EXAM_PAPER_DEPARTMENTS.includes(item));
      const rowsNeedingReview = parsed.filter(course => examPaperCourseIssues(course).length).length;
      const blankOutcomes = parsed.filter(course => course.learningOutcomes === "").length;
      setOperation({ title: "قراءة ملف التقييم", detail: "مراجعة جاهزية التقارير…", current: 3, total: 3 });
      setMessage(unknown.length
        ? `تم تحميل ${parsed.length} مقرر. راجع أسماء الأقسام غير القياسية: ${unknown.join("، ")}.`
        : `تم تقسيم ${parsed.length} مقرر تلقائيًا على ${departments.length} قسم${rowsNeedingReview ? `، ويوجد ${rowsNeedingReview} مقرر يحتاج استكمالًا` : ""}${blankOutcomes ? `، و${blankOutcomes} مقرر بدون تقييم مخرجات تعلم` : ""}.`);
    } catch (uploadError) {
      setError(uploadError.message);
    } finally {
      setBusy(""); setOperation(null);
    }
  };

  const buildReportPayload = (departmentName = department) => {
    if (!courses.length) throw new Error("حمّل ملف المقررات أولًا.");
    const departmentCourses = allEvaluatedCourses.filter(course => course.department === departmentName);
    if (!departmentName || !departmentCourses.length) throw new Error("اختر قسمًا يحتوي على مقررات.");
    if (!semester.trim() || !year.trim()) throw new Error("أدخل الفصل الدراسي والسنة.");
    for (const course of departmentCourses) {
      const issues = examPaperCourseIssues(course);
      if (issues.length) throw new Error(`استكمل ${issues.join(" و")} للمقرر ${course.name} (${course.code}).`);
    }
    const cleanEvaluators = evaluators
      .map(evaluator => ({ name: evaluator.name.trim(), role: evaluator.role.trim() }))
      .filter(evaluator => evaluator.name)
      .map(evaluator => ({ ...evaluator, role: evaluator.role || "عضو لجنة القياس والتقويم" }));
    return {
      department: departmentName,
      semester: semester.trim(),
      year: year.trim(),
      courses: departmentCourses.map(course => ({ ...course, learningOutcomes: course.learningOutcomes === "" ? "" : Number(course.learningOutcomes) })),
      evaluators: cleanEvaluators.length ? cleanEvaluators : [{ name: "—", role: "—" }],
      committeeHead: committeeHead.trim(),
    };
  };

  const requestBlankOutcomeApproval = (type, departmentNames) => {
    const names = new Set(departmentNames);
    const missing = allEvaluatedCourses.filter(course => names.has(course.department) && course.learningOutcomes === "");
    if (!missing.length) return false;
    setBlankOutcomePrompt({ type, courses: missing });
    return true;
  };

  const handleDownloadAll = async (allowBlankOutcomes = false) => {
    if (!allowBlankOutcomes && requestBlankOutcomeApproval("download-all", fileDepartments)) return;
    setBusy("download-all"); setError(""); setMessage("");
    try {
      const zip = new JSZip();
      const totalSteps = fileDepartments.length + 1;
      setOperation({ title: "إنشاء تقارير الأقسام", detail: "التحقق من بيانات الملف…", current: 0, total: totalSteps });
      for (const [index, departmentName] of fileDepartments.entries()) {
        setOperation({ title: "إنشاء تقارير الأقسام", detail: `إنشاء تقرير قسم ${departmentName}…`, current: index, total: totalSteps });
        const report = buildReportPayload(departmentName);
        const blob = await buildExamPaperReportDocx(report, reportSettings);
        zip.file(examPaperReportFilename(report), blob);
        setOperation({ title: "إنشاء تقارير الأقسام", detail: `اكتمل تقرير قسم ${departmentName}`, current: index + 1, total: totalSteps });
      }
      setOperation({ title: "إنشاء تقارير الأقسام", detail: "ضغط التقارير في ملف ZIP…", current: fileDepartments.length, total: totalSteps });
      const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
      setOperation({ title: "إنشاء تقارير الأقسام", detail: "بدء التنزيل…", current: totalSteps, total: totalSteps });
      downloadBlob(blob, `تقارير تقييم الورقة الامتحانية - ${semester} ${year}.zip`);
      setMessage(`تم إنشاء ${fileDepartments.length} تقرير Word، تقرير مستقل لكل قسم، وتجميعها في ملف ZIP واحد.`);
    } catch (actionError) {
      setError(actionError.message);
    } finally { setBusy(""); setOperation(null); }
  };

  const handleDownload = async (allowBlankOutcomes = false) => {
    if (!allowBlankOutcomes && requestBlankOutcomeApproval("download", [department])) return;
    setBusy("download"); setError(""); setMessage("");
    try {
      setOperation({ title: "إنشاء تقرير Word", detail: "التحقق من بيانات القسم…", current: 1, total: 3 });
      const report = buildReportPayload();
      setOperation({ title: "إنشاء تقرير Word", detail: `إنشاء تقرير قسم ${report.department}…`, current: 2, total: 3 });
      const blob = await buildExamPaperReportDocx(report, reportSettings);
      setOperation({ title: "إنشاء تقرير Word", detail: "بدء تنزيل التقرير…", current: 3, total: 3 });
      downloadBlob(blob, examPaperReportFilename(report));
      setMessage("تم إنشاء التقرير وإرساله إلى التنزيلات في المتصفح.");
    } catch (actionError) {
      setError(actionError.message);
    } finally { setBusy(""); setOperation(null); }
  };

  const handlePrint = (allowBlankOutcomes = false) => {
    if (!allowBlankOutcomes && requestBlankOutcomeApproval("print", [department])) return;
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

  const handleDriveUpload = async (allowBlankOutcomes = false) => {
    if (!allowBlankOutcomes && requestBlankOutcomeApproval("drive", [department])) return;
    if (!googleAuth.token) {
      setMessage("اربط حساب Google Drive أولًا، ثم اضغط زر الرفع مرة أخرى.");
      googleAuth.connect?.();
      return;
    }
    setBusy("drive"); setError(""); setMessage(""); setDriveFile(null);
    try {
      setOperation({ title: "رفع التقرير إلى Drive", detail: "إنشاء ملف Word…", current: 1, total: 5 });
      const report = buildReportPayload();
      const blob = await buildExamPaperReportDocx(report, reportSettings);
      setOperation({ title: "رفع التقرير إلى Drive", detail: "تجهيز مجلد السنة…", current: 2, total: 5 });
      const yearFolder = await getOrCreateFolder(googleAuth.token, ROOT_SURVEYS_FOLDER_ID, yearFolderName(report.year));
      setOperation({ title: "رفع التقرير إلى Drive", detail: "تجهيز مجلد الفصل والقسم…", current: 3, total: 5 });
      const semesterFolder = await getOrCreateFolder(googleAuth.token, yearFolder.id, report.semester);
      const departmentFolder = await getOrCreateFolder(googleAuth.token, semesterFolder.id, report.department);
      const reportFolder = await getOrCreateFolder(
        googleAuth.token,
        departmentFolder.id,
        `تقرير تقييم الورقة الامتحانية من حيث الشكل - ${report.semester} ${report.year}`,
      );
      setOperation({ title: "رفع التقرير إلى Drive", detail: "رفع التقرير…", current: 4, total: 5 });
      const uploaded = await uploadBlobToFolder(googleAuth.token, reportFolder.id, blob, examPaperReportFilename(report));
      setOperation({ title: "رفع التقرير إلى Drive", detail: "اكتمل الرفع", current: 5, total: 5 });
      setDriveFile(uploaded);
      setMessage("تم رفع التقرير داخل مجلد القسم على Google Drive.");
    } catch (actionError) {
      setError(actionError.message);
    } finally { setBusy(""); setOperation(null); }
  };

  const continueWithBlankOutcomes = () => {
    const type = blankOutcomePrompt?.type;
    setBlankOutcomePrompt(null);
    if (type === "download-all") handleDownloadAll(true);
    else if (type === "download") handleDownload(true);
    else if (type === "print") handlePrint(true);
    else if (type === "drive") handleDriveUpload(true);
  };

  return <section className="exam-paper-tool">
    <style>{EXAM_PAPER_CSS}</style>
    <ExamPaperLoading operation={operation}/>
    {blankOutcomePrompt && <div className="exam-paper-confirm" role="dialog" aria-modal="true" aria-labelledby="blank-outcome-title">
      <div className="exam-paper-confirm-card">
        <div className="exam-paper-confirm-icon">!</div>
        <h3 id="blank-outcome-title">توجد تقييمات مخرجات تعلم فارغة</h3>
        <p>يوجد {blankOutcomePrompt.courses.length} مقرر بدون تقييم مخرجات تعلم. يمكنك الرجوع واستكمالها، أو المتابعة الآن وستظل الخانة فارغة في التقرير.</p>
        <div className="exam-paper-confirm-list">{blankOutcomePrompt.courses.slice(0, 6).map(course => <span key={examPaperCourseKey(course)}>{course.code} — {course.name}</span>)}{blankOutcomePrompt.courses.length > 6 && <b>و{blankOutcomePrompt.courses.length - 6} مقررات أخرى</b>}</div>
        <div className="exam-paper-confirm-actions">
          <button type="button" className="btn btn-primary" onClick={continueWithBlankOutcomes}>متابعة وتركها فارغة</button>
          <button type="button" className="btn btn-ghost" onClick={() => setBlankOutcomePrompt(null)}>الرجوع للتعديل</button>
        </div>
      </div>
    </div>}
    <header className="exam-paper-hero">
      <div><span>أداة تقارير لجنة القياس والتقويم</span><h2>تقييم الورقة الامتحانية من حيث الشكل</h2><p>حمّل القالب المعتمد، املأه في Excel، ثم ارفعه ليتم تقسيمه حسب الأقسام وإنشاء التقارير.</p></div>
      <div className="exam-paper-progress"><b>{overallProgress}%</b><small>جاهزية الملف بالكامل</small></div>
    </header>

    {error && <InlineNotice style={{ marginBottom: 14 }}>{error}</InlineNotice>}
    {message && <div className="exam-paper-success" role="status">{message}{driveFile?.webViewLink && <> <a href={driveFile.webViewLink} target="_blank" rel="noreferrer">فتح التقرير على Drive</a></>}</div>}

    <div className="exam-paper-steps">
      <article className="exam-paper-panel">
        <div className="exam-paper-panel-head"><span>1</span><div><h3>قالب تقييم الورقة الامتحانية</h3><p>نفس تنسيق الملف المرفق بجميع أعمدة التقييم العشرة.</p></div></div>
        <div className="exam-paper-template-actions">
          <button className="btn btn-blue btn-sm" type="button" onClick={handleTemplateDownload}><QualityIcon name="download" size={15}/> تحميل قالب Excel</button>
          <button className="btn btn-primary btn-sm" type="button" disabled={busy === "file"} onClick={() => inputRef.current?.click()}><QualityIcon name="upload" size={15}/> {busy === "file" ? "جاري القراءة…" : "رفع القالب بعد تعبئته"}</button>
          <input ref={inputRef} type="file" accept=".xlsx,.xls" hidden onChange={handleFile}/>
        </div>
        {sourceName && <div className="exam-paper-source"><QualityIcon name="sheet" size={15}/><span>{sourceName}</span><b>{courses.length} مقرر</b></div>}
      </article>

      <article className="exam-paper-panel">
        <div className="exam-paper-panel-head"><span>2</span><div><h3>بيانات التقرير والأقسام</h3><p>تم فصل الأقسام تلقائيًا. اختر قسمًا للمراجعة أو أنشئ تقاريرها كلها.</p></div></div>
        <div className="exam-paper-meta-grid">
          <label><span>القسم</span><select className="input" value={department} onChange={event => { setDepartment(event.target.value); setDriveFile(null); }} disabled={!courses.length}><option value="">اختر القسم</option>{fileDepartments.map(item => <option key={item} value={item}>{item}</option>)}</select></label>
          <label><span>الفصل الدراسي</span><select className="input" value={semester} onChange={event => setSemester(event.target.value)}><option>خريف</option><option>ربيع</option><option>صيف</option></select></label>
          <label><span>السنة الدراسية</span><input className="input" value={year} onChange={event => setYear(event.target.value)} placeholder="2025-2026"/></label>
        </div>
        {departmentSummaries.length > 0 && <div className="exam-paper-departments" aria-label="الأقسام المستخرجة من الملف">
          {departmentSummaries.map(item => <button type="button" key={item.name} className={`${department === item.name ? "active" : ""} ${item.ready ? "ready" : "needs-review"}`} onClick={() => setDepartment(item.name)}>
            <span>{item.name}</span><small>{item.total} مقرر</small><b>{item.ready ? "جاهز" : `${item.issueCount} يحتاج مراجعة`}</b>
          </button>)}
        </div>}
      </article>
    </div>

    <article className="exam-paper-panel exam-paper-courses-panel">
      <div className="exam-paper-panel-head"><span>3</span><div><h3>مراجعة بيانات القسم</h3><p>{department ? `${department} — ${visibleCourses.length} مقرر — جاهزية ${progress}%` : "اختر القسم لمراجعة بياناته"}</p></div></div>
      {!courses.length ? <div className="exam-paper-empty"><QualityIcon name="sheet" size={27}/><b>ابدأ برفع قالب المقررات</b><span>ستظهر المقررات هنا تلقائيًا بعد قراءة الملف.</span></div>
        : !department ? <div className="exam-paper-empty"><QualityIcon name="folder" size={27}/><b>اختر قسمًا</b><span>يتم عرض مقررات القسم المختار فقط.</span></div>
        : <div className="exam-paper-course-list">{evaluatedCourses.map((course, index) => {
          const value = evaluations[examPaperCourseKey(course)] || {};
          return <div className={`exam-paper-course ${value.status || "pending"}`} key={examPaperCourseKey(course)}>
            <div className="exam-paper-course-id"><span>{index + 1}</span><div><b>{course.name}</b><small>{course.code}</small><em>{[course.exam && `Exam: ${course.exam}`, course.blueprint && `Blueprint: ${course.blueprint}`, course.evaluation && `التقييم: ${course.evaluation}`, course.signature && `التوقيع: ${course.signature}`].filter(Boolean).join(" · ")}</em></div></div>
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
          <button className="btn btn-blue" type="button" disabled={Boolean(busy) || !fileDepartments.length} onClick={() => handleDownloadAll()}><QualityIcon name="download" size={17}/>{busy === "download-all" ? "جاري إنشاء تقارير الأقسام…" : `تحميل كل تقارير الأقسام (${fileDepartments.length})`}</button>
          <button className="btn btn-primary" type="button" disabled={Boolean(busy)} onClick={() => handleDownload()}><QualityIcon name="download" size={17}/>{busy === "download" ? "جاري إنشاء Word…" : "تحميل التقرير Word"}</button>
          <button className="btn btn-ghost" type="button" disabled={Boolean(busy)} onClick={() => handlePrint()}><QualityIcon name="document" size={17}/> طباعة أو حفظ PDF</button>
          <button className="btn btn-blue" type="button" disabled={Boolean(busy) || googleAuth.connecting} onClick={() => handleDriveUpload()}><GoogleDriveIcon size={18}/>{busy === "drive" ? "جاري الرفع…" : googleAuth.token ? "رفع على Google Drive" : "ربط Drive ثم الرفع"}</button>
        </div>
        <p className="exam-paper-folder-note">مسار Drive: السنة / الفصل / القسم / تقرير تقييم الورقة الامتحانية من حيث الشكل للفصل الدراسي والسنة.</p>
      </article>
    </div>
  </section>;
}

const EXAM_PAPER_CSS = `
.exam-paper-tool{direction:rtl;color:#edf7ff}.exam-paper-hero{display:flex;align-items:center;justify-content:space-between;gap:20px;padding:25px 28px;margin-bottom:14px;border-radius:21px;border:1px solid rgba(94,234,212,.17);background:linear-gradient(120deg,rgba(15,74,84,.66),rgba(22,51,83,.64))}.exam-paper-hero span{color:#74e8d2;font-size:10px;font-weight:900}.exam-paper-hero h2{font-size:25px;margin:7px 0 4px}.exam-paper-hero p{margin:0;color:rgba(235,247,255,.53);font-size:11.5px}.exam-paper-progress{flex:0 0 120px;display:grid;place-items:center;padding:15px;border-radius:17px;background:rgba(3,18,31,.32);border:1px solid rgba(255,255,255,.09)}.exam-paper-progress b{font-size:27px;color:#7cead5}.exam-paper-progress small{color:rgba(255,255,255,.42);font-size:8.5px}.exam-paper-success{margin-bottom:14px;padding:11px 14px;border-radius:11px;color:#baf7df;background:rgba(34,197,94,.1);border:1px solid rgba(74,222,128,.18);font-size:11px}.exam-paper-success a{color:#78d8ff;font-weight:800;margin-right:8px}.exam-paper-steps{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}.exam-paper-steps.bottom{align-items:stretch}.exam-paper-panel{padding:19px;border-radius:18px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.085)}.exam-paper-panel-head{display:flex;align-items:center;gap:10px;margin-bottom:15px}.exam-paper-panel-head>span{width:34px;height:34px;display:grid;place-items:center;flex:0 0 34px;border-radius:10px;color:#eafffa;background:linear-gradient(135deg,#1abc9c,#2374a5);font-weight:900;font-size:12px}.exam-paper-panel-head h3{font-size:14px;margin:0}.exam-paper-panel-head p{font-size:9.5px;color:rgba(255,255,255,.42);margin:2px 0 0}.exam-paper-template-actions{display:flex;gap:8px;flex-wrap:wrap}.exam-paper-source{display:flex;align-items:center;gap:8px;margin-top:12px;padding:9px 11px;border-radius:10px;background:rgba(2,15,27,.23);color:rgba(255,255,255,.62);font-size:10px}.exam-paper-source span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.exam-paper-source b{margin-right:auto;color:#76e5d0}.exam-paper-meta-grid{display:grid;grid-template-columns:1.4fr .8fr .9fr;gap:10px}.exam-paper-meta-grid label>span,.exam-paper-outcome>span,.exam-paper-missing>span,.exam-paper-head-name>span{display:block;margin-bottom:5px;color:rgba(221,238,250,.64);font-size:9.5px;font-weight:800}.exam-paper-departments{display:flex;gap:7px;overflow:auto;margin-top:12px;padding-bottom:2px}.exam-paper-departments button{min-width:135px;padding:9px 10px;text-align:right;border-radius:10px;border:1px solid rgba(255,255,255,.09);background:rgba(2,15,27,.25);color:#e8f5ff;cursor:pointer}.exam-paper-departments button.active{border-color:#5eead4;box-shadow:0 0 0 1px rgba(94,234,212,.15)}.exam-paper-departments span,.exam-paper-departments small,.exam-paper-departments b{display:block}.exam-paper-departments span{font:800 9.5px 'Cairo'}.exam-paper-departments small{margin-top:2px;color:rgba(255,255,255,.42);font:600 8px 'Cairo'}.exam-paper-departments b{margin-top:5px;color:#6ee7b7;font:800 8px 'Cairo'}.exam-paper-departments .needs-review b{color:#fcd34d}.exam-paper-courses-panel{margin-bottom:12px}.exam-paper-empty{min-height:160px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;border:1px dashed rgba(255,255,255,.14);border-radius:13px;color:rgba(255,255,255,.34)}.exam-paper-empty b{color:rgba(255,255,255,.62);font-size:12px}.exam-paper-empty span{font-size:9.5px}.exam-paper-course-list{display:grid;gap:9px}.exam-paper-course{display:grid;grid-template-columns:minmax(210px,1.2fr) minmax(250px,1.3fr) 170px;gap:11px;align-items:center;padding:12px;border-radius:13px;background:rgba(2,16,29,.27);border:1px solid rgba(255,255,255,.075);transition:.18s}.exam-paper-course.complete{border-color:rgba(52,211,153,.2)}.exam-paper-course.incomplete{border-color:rgba(251,191,36,.23)}.exam-paper-course-id{display:flex;align-items:center;gap:9px;min-width:0}.exam-paper-course-id>span{width:27px;height:27px;display:grid;place-items:center;flex:0 0 27px;border-radius:8px;color:#72e6d0;background:rgba(26,188,156,.1);font-size:9px;font-weight:900}.exam-paper-course-id div{min-width:0}.exam-paper-course-id b,.exam-paper-course-id small,.exam-paper-course-id em{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.exam-paper-course-id b{font-size:10.5px}.exam-paper-course-id small{margin-top:2px;color:#7fc8ee;font-size:9px}.exam-paper-course-id em{margin-top:3px;color:rgba(255,255,255,.32);font-size:7.5px;font-style:normal}.exam-paper-status{display:grid;grid-template-columns:1fr 1fr;gap:6px}.exam-paper-status button{padding:8px 7px;border-radius:8px;border:1px solid rgba(255,255,255,.11);background:rgba(255,255,255,.035);color:rgba(255,255,255,.46);font:700 9px 'Cairo',sans-serif;cursor:pointer}.exam-paper-status button.selected.complete{color:#c5fae5;background:rgba(34,197,94,.13);border-color:rgba(74,222,128,.32)}.exam-paper-status button.selected.incomplete{color:#ffebad;background:rgba(245,158,11,.13);border-color:rgba(251,191,36,.32)}.exam-paper-outcome div{display:flex;align-items:center;gap:6px}.exam-paper-outcome .input{direction:ltr;text-align:center;padding:8px}.exam-paper-outcome i{font-style:normal;color:#77e6d1;font-size:11px}.exam-paper-missing{grid-column:1/-1}.exam-paper-missing .textarea{min-height:62px}.exam-paper-issue-options{display:flex;flex-wrap:wrap;gap:5px;margin-top:6px}.exam-paper-issue-options button{padding:4px 8px;border:1px solid rgba(255,255,255,.1);border-radius:999px;background:rgba(255,255,255,.04);color:rgba(255,255,255,.48);font:700 8px 'Cairo';cursor:pointer}.exam-paper-issue-options button:hover{color:#d7fff7;border-color:rgba(94,234,212,.25);background:rgba(26,188,156,.08)}.exam-paper-evaluators{display:grid;gap:7px}.exam-paper-evaluator-row{display:grid;grid-template-columns:1fr 1.3fr 28px;gap:7px}.exam-paper-evaluator-row .input{padding:9px 11px;font-size:11px}.exam-paper-evaluator-row button{border:0;border-radius:8px;background:rgba(239,68,68,.11);color:#fca5a5;cursor:pointer;font-size:17px}.exam-paper-add{margin-top:9px;border:0;background:transparent;color:#71e5d0;font:800 9.5px 'Cairo';cursor:pointer}.exam-paper-head-name{display:block;margin-top:13px}.exam-paper-head-name .input{font-size:11px}.exam-paper-summary{display:flex;flex-direction:column}.exam-paper-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:7px}.exam-paper-stats div{padding:11px;text-align:center;border-radius:10px;background:rgba(255,255,255,.045);border:1px solid rgba(255,255,255,.07)}.exam-paper-stats b,.exam-paper-stats span{display:block}.exam-paper-stats b{font-size:20px}.exam-paper-stats span{color:rgba(255,255,255,.4);font-size:8.5px}.exam-paper-stats .ok b{color:#6ee7b7}.exam-paper-stats .warn b{color:#fcd34d}.exam-paper-delivery{display:grid;grid-template-columns:1fr;gap:7px;margin-top:13px}.exam-paper-delivery .btn{justify-content:center;padding:9px 14px;font-size:10px}.exam-paper-folder-note{margin:auto 0 0;padding-top:11px;color:rgba(255,255,255,.3);font-size:8.5px;line-height:1.7}.exam-paper-loading,.exam-paper-confirm{position:fixed;inset:0;z-index:9200;display:grid;place-items:center;padding:18px;background:rgba(4,13,25,.91);backdrop-filter:blur(12px)}.exam-paper-loading-card,.exam-paper-confirm-card{width:min(430px,100%);padding:28px;border-radius:22px;text-align:center;background:linear-gradient(150deg,#112a43,#0b1b2e);border:1px solid rgba(255,255,255,.13);box-shadow:0 30px 80px rgba(0,0,0,.5)}.exam-paper-loading-spinner{position:relative;width:68px;height:68px;margin:0 auto 18px}.exam-paper-loading-spinner i{position:absolute;inset:0;border:3px solid transparent;border-top-color:#5eead4;border-radius:50%;animation:examSpin 1s linear infinite}.exam-paper-loading-spinner i+ i{inset:9px;border-top-color:#60a5fa;animation-direction:reverse;animation-duration:.75s}.exam-paper-loading-card h3,.exam-paper-confirm-card h3{margin:0;color:#fff;font-size:18px}.exam-paper-loading-card p,.exam-paper-confirm-card p{margin:8px 0 18px;color:rgba(255,255,255,.58);font-size:11px;line-height:1.8}.exam-paper-loading-track{height:9px;border-radius:99px;overflow:hidden;background:rgba(255,255,255,.1)}.exam-paper-loading-track span{display:block;height:100%;border-radius:inherit;background:linear-gradient(90deg,#1abc9c,#3b82f6);transition:width .25s}.exam-paper-loading-count{display:flex;justify-content:space-between;margin-top:9px;color:rgba(255,255,255,.48);font-size:10px}.exam-paper-loading-count b{color:#78ead5}.exam-paper-confirm{z-index:9300}.exam-paper-confirm-icon{width:48px;height:48px;margin:0 auto 14px;display:grid;place-items:center;border-radius:50%;color:#ffefb0;background:rgba(245,158,11,.16);border:1px solid rgba(251,191,36,.3);font-size:24px;font-weight:900}.exam-paper-confirm-list{display:grid;gap:5px;max-height:150px;overflow:auto;margin-bottom:18px;padding:10px;border-radius:10px;text-align:right;background:rgba(0,0,0,.17);font-size:9px;color:rgba(255,255,255,.65)}.exam-paper-confirm-list b{color:#fcd34d}.exam-paper-confirm-actions{display:flex;gap:8px;justify-content:center}.exam-paper-confirm-actions .btn{font-size:10px}@keyframes examSpin{to{transform:rotate(360deg)}}@media(max-width:900px){.exam-paper-course{grid-template-columns:1fr 1fr}.exam-paper-outcome{grid-column:1/-1}.exam-paper-meta-grid{grid-template-columns:1fr 1fr}.exam-paper-meta-grid label:first-child{grid-column:1/-1}}@media(max-width:650px){.exam-paper-hero{padding:19px}.exam-paper-progress{display:none}.exam-paper-steps{grid-template-columns:1fr}.exam-paper-meta-grid,.exam-paper-course{grid-template-columns:1fr}.exam-paper-meta-grid label:first-child,.exam-paper-outcome,.exam-paper-missing{grid-column:auto}.exam-paper-status{grid-template-columns:1fr}.exam-paper-evaluator-row{grid-template-columns:1fr}.exam-paper-evaluator-row button{min-height:30px}.exam-paper-confirm-actions{flex-direction:column}}
`;
