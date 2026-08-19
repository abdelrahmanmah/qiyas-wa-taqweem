import { useState, useRef, useEffect, useCallback } from "react";
import { readExcel, detectSurveyType, analyze, SCHEMAS } from "./engine/analyze.js";
import { buildAnnualDocx } from "./engine/buildDocx.js";
import { getAllAnalysisSchemas, detectAnySurveyType } from "./engine/customSurveyModel.js";
import {
  loadGisScript, initSemesterTokenClient, SEMESTER_TOKEN_KEY, saveStoredToken, getStoredToken,
  TEMPLATE_FOLDER_ID, ROOT_SURVEYS_FOLDER_ID, SEMESTERS, loadDepartments,
  listFormsInFolder, findYearSemesterFolder, buildGenerationJobs, runGenerationJobs,
  listSemesterSurveysWithStats, getForm, listAllResponses, responsesToRows,
  departmentFromSurveyName, isValidAcademicYear,
} from "./engine/semesterSurveyModel.js";

// ── Provider catalog ──────────────────────────────────────────────────────────
export const PROVIDERS = {
  groq: {
    label: "Groq", icon: "⚡", color: "#FF6B35",
    models: [
      { id: "llama-3.1-8b-instant",    label: "Llama 8B — سريع جداً" },
      { id: "llama-3.3-70b-versatile", label: "Llama 70B — متوازن ★" },
      { id: "gemma2-9b-it",            label: "Gemma 9B" },
    ],
    default: "llama-3.3-70b-versatile",
  },
  gemini: {
    label: "Gemini", icon: "✨", color: "#4285F4",
    models: [
      { id: "gemini-2.0-flash",   label: "Flash 2.0 — سريع ★" },
      { id: "gemini-1.5-flash",   label: "Flash 1.5" },
      { id: "gemini-1.5-pro",     label: "Pro 1.5 — متقدم" },
    ],
    default: "gemini-2.0-flash",
  },
  claude: {
    label: "Claude", icon: "🤖", color: "#CC785C",
    models: [
      { id: "claude-haiku-4-5-20251001", label: "Haiku — سريع" },
      { id: "claude-sonnet-4-6",         label: "Sonnet — متوازن ★" },
      { id: "claude-opus-4-7",           label: "Opus — متقدم" },
    ],
    default: "claude-sonnet-4-6",
  },
};

export const DEFAULT_AI_SETTINGS = {
  provider:      "groq",
  model:         "llama-3.3-70b-versatile",
  key:           "",
  surveysFolder: "",
};

// ── Tool definitions (OpenAI format) ─────────────────────────────────────────
const TOOLS = [
  {
    type: "function",
    function: {
      name: "generate_report",
      description: "ينشئ تقرير Word سنوي للاستبيان المحلل ويحمّله تلقائياً. استخدمها فقط عندما يطلب المستخدم صراحةً إنشاء أو تحميل تقرير.",
      parameters: {
        type: "object",
        properties: {
          year:       { type: "string", description: "العام الدراسي مثال: 2024-2025" },
          program:    { type: "string", description: "اسم البرنامج أو القسم" },
          preparedBy: { type: "string", description: "اسم معد التقرير" },
          reviewer:   { type: "string", description: "اسم المراجع" },
        },
        required: ["year"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "request_file_upload",
      description: "يطلب من المستخدم رفع ملف Excel لتحليله. استخدمها فقط عندما لا يوجد بيانات محملة ويريد المستخدم تحليل استبيان.",
      parameters: {
        type: "object",
        properties: {
          message: { type: "string", description: "رسالة توجيهية للمستخدم" },
          type:    { type: "string", description: "نوع الاستبيان: student أو faculty أو assistant" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_survey_files",
      description: "يسرد ملفات Excel الموجودة في مجلد الاستبيانات المُعيّن. استخدمها عندما يريد المستخدم معرفة الملفات المتاحة أو يطلب قراءة ملف من المجلد.",
      parameters: {
        type: "object",
        properties: {
          subfolder: { type: "string", description: "اسم المجلد الفرعي للبحث فيه (اختياري)" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "analyze_file_from_folder",
      description: "يقرأ ويحلل ملف Excel مباشرة من مجلد الاستبيانات دون الحاجة لرفعه يدوياً.",
      parameters: {
        type: "object",
        properties: {
          file_path: { type: "string", description: "المسار الكامل للملف كما أعادته list_survey_files" },
          type:      { type: "string", description: "نوع الاستبيان: student أو faculty أو assistant" },
        },
        required: ["file_path"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_semester_survey_templates",
      description: "يسرد أنواع الاستبيانات (القوالب) المتاحة لإنشاء نسخ منها لفصل دراسي معين، من مجلد القوالب على Google Drive.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "generate_semester_surveys",
      description: "ينشئ نسخاً فعلية من استبيان(ات) لفصل دراسي وسنة معينة على Google Drive/Forms وينشرها لاستقبال الردود. عملية حقيقية غير سهلة التراجع — ستُعرض للمستخدم بطاقة تأكيد قبل التنفيذ الفعلي، فقط استدع الأداة عندما تكون السنة والفصل واضحين من كلام المستخدم.",
      parameters: {
        type: "object",
        properties: {
          templateNames: { type: "array", items: { type: "string" }, description: "أسماء أو أجزاء من أسماء القوالب المطلوب نسخها (استخدم list_semester_survey_templates أولاً). اتركه فارغاً لتحديد كل القوالب." },
          year: { type: "string", description: "السنة الدراسية بصيغة 2026/2027" },
          semester: { type: "string", description: `الفصل الدراسي: ${SEMESTERS.join(" أو ")}` },
          mode: { type: "string", enum: ["general", "departments", "programs"], description: "general = نسخة عامة واحدة (الافتراضي)، departments = نسخة لكل قسم، programs = نسخة لكل قسم وبرنامج" },
        },
        required: ["year", "semester"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_semester_surveys",
      description: "يسرد الاستبيانات التي تم إنشاؤها بالفعل لفصل دراسي وسنة معينة على Google Drive مع عدد الردود لكل استبيان.",
      parameters: {
        type: "object",
        properties: {
          year: { type: "string", description: "السنة الدراسية بصيغة 2026/2027" },
          semester: { type: "string", description: `الفصل الدراسي: ${SEMESTERS.join(" أو ")}` },
        },
        required: ["year", "semester"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "analyze_semester_surveys",
      description: "يحلل استبيانات فصل دراسي محدد (كلها أو استبيانات معينة بالاسم) وينشئ تقرير Word لكل واحد ويحمّله تلقائياً. ستُعرض للمستخدم بطاقة تأكيد قبل التنفيذ الفعلي.",
      parameters: {
        type: "object",
        properties: {
          year: { type: "string", description: "السنة الدراسية بصيغة 2026/2027" },
          semester: { type: "string", description: `الفصل الدراسي: ${SEMESTERS.join(" أو ")}` },
          surveyNames: { type: "array", items: { type: "string" }, description: "أسماء أو أجزاء من أسماء الاستبيانات المطلوب تحليلها فقط (اختياري — استخدم list_semester_surveys أولاً). اتركه فارغاً لتحليل كل استبيانات هذا الفصل." },
          preparedBy: { type: "string", description: "اسم معد التقرير" },
          reviewer: { type: "string", description: "اسم المراجع" },
        },
        required: ["year", "semester"],
      },
    },
  },
];

// ── System prompt ─────────────────────────────────────────────────────────────
function buildSystemPrompt(result, surveysFolder) {
  const dataSection = result
    ? `## البيانات المحملة\n- النوع: ${result.schemaLabel}\n- المستجيبون: ${result.n}\n- نسبة الموافقة: ${result.overallAgreePct}%\n- الاتجاه: ${result.overallDirection}\n\n### المحاور:\n${result.axes.map((ax, i) => `${i + 1}. ${ax.name}: ${ax.axisAgreePct}% (${ax.direction})`).join("\n")}`
    : `## الحالة\nلا يوجد استبيان محلل. استخدم list_survey_files أو request_file_upload عندما يريد المستخدم تحليل ملف.`;

  const folderSection = surveysFolder
    ? `\n## مجلد الاستبيانات\nالمسار: ${surveysFolder}\nيمكنك استخدام list_survey_files لاستعراض الملفات وanalyze_file_from_folder لتحليل ملف مباشرة دون رفع.`
    : "";

  const semesterSection = `
## استبيانات الفصل الدراسي (Google Drive/Forms)
تُنشأ وتُدار في تبويب «📆 استبيانات الفصل الدراسي» بالتطبيق. الأدوات المتاحة لك:
- list_semester_survey_templates: عرض القوالب المتاحة للنسخ
- generate_semester_surveys: إنشاء نسخ فعلية من استبيان(ات) لسنة وفصل دراسي معينين — عملية حقيقية على Google Drive، ستظهر للمستخدم تلقائياً بطاقة تأكيد (نعم/إلغاء) قبل التنفيذ، فلا داعي أن تسأله في الرد النصي، فقط نفّذ الأداة مباشرة متى كانت السنة والفصل والنوع المطلوب واضحين
- list_semester_surveys: عرض الاستبيانات الموجودة بالفعل لفصل دراسي مع عدد الردود
- analyze_semester_surveys: تحليل استبيان أو أكثر من فصل دراسي معين وتحميل تقرير Word لكل واحد — تظهر بطاقة تأكيد أيضاً قبل التنفيذ
هذه الأدوات تحتاج اتصال المستخدم بحساب Google من نفس التبويب مسبقاً؛ إن رجعت رسالة بضرورة الاتصال، أخبر المستخدم بذلك بوضوح بدل إعادة المحاولة.`;

  return `أنت مساعد ذكي متخصص في تحليل استبيانات ضمان الجودة الأكاديمية — وحدة ضمان الجودة، كلية الإدارة والاقتصاد، الجامعة المصرية الروسية.

## قدراتك:
- شرح نتائج الاستبيانات وتحديد المحاور الأقوى والأضعف
- تقديم توصيات تحسينية مبنية على الأرقام
- إنشاء تقارير Word بأمر واحد
- تحليل ملفات Excel من المجلد المحلي أو المرفوعة مباشرة
- إنشاء وتحليل استبيانات الفصل الدراسي على Google Drive/Forms
${folderSection}
${semesterSection}

${dataSection}

## تعليمات:
- تحدث بالعربية بشكل رئيسي، مختصر ومهني
- استخدم الأرقام الفعلية من البيانات في إجاباتك
- إذا طلب تقرير بدون عام دراسي، اسأل عنه أولاً
- إذا كان مجلد الاستبيانات متاحاً، فضّل list_survey_files ثم analyze_file_from_folder على request_file_upload
- لا تستخدم request_file_upload إذا كانت هناك بيانات محملة بالفعل`;
}

// ── Small helpers ─────────────────────────────────────────────────────────────
function MsgText({ text }) {
  return (
    <span style={{ whiteSpace: "pre-wrap", lineHeight: 1.7, fontSize: 13.5, direction: "rtl" }}>
      {text}
    </span>
  );
}

function TypingDots() {
  return (
    <div style={{ display: "flex", gap: 5, alignItems: "center", padding: "6px 2px" }}>
      {[0, 1, 2].map(i => (
        <div key={i} style={{
          width: 7, height: 7, borderRadius: "50%", background: "#1abc9c",
          animation: `aichat-bounce 1.2s ${i * 0.2}s ease-in-out infinite`,
        }} />
      ))}
    </div>
  );
}

// ── Main component ────────────────────────────────────────────────────────────
export default function AiChat({ currentResult, aiSettings, docSettings, onAnalysisComplete }) {
  const [isOpen, setIsOpen]   = useState(false);
  const [msgs, setMsgs]       = useState([]);
  const [input, setInput]     = useState("");
  const [loading, setLoading] = useState(false);

  const rawHistoryRef          = useRef([]);
  const currentResRef          = useRef(currentResult);
  const freshResRef            = useRef(null);
  const pendingUploadRef       = useRef(null);
  const pendingConfirmRef      = useRef(null);
  const semesterClientRef      = useRef(null);
  const aiSettingsRef          = useRef(aiSettings);
  const onAnalysisCompleteRef  = useRef(onAnalysisComplete);
  const fileInputRef           = useRef();
  const bottomRef              = useRef();
  const textareaRef            = useRef();

  useEffect(() => { currentResRef.current = currentResult; }, [currentResult]);
  useEffect(() => { aiSettingsRef.current = aiSettings; }, [aiSettings]);
  useEffect(() => { onAnalysisCompleteRef.current = onAnalysisComplete; }, [onAnalysisComplete]);
  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior: "smooth" }); }, [msgs, loading]);

  const addMsg = useCallback(msg =>
    setMsgs(prev => [...prev, { id: Date.now() + Math.random(), ...msg }]), []);

  // ── Semester-survey Google token (reuses the stored token from the "📆 استبيانات
  // الفصل الدراسي" tab; only tries a silent, no-popup refresh here — an explicit
  // account-picker popup from inside a chat bubble would be surprising) ──────────
  const ensureSemesterToken = useCallback(() => {
    const cached = getStoredToken(SEMESTER_TOKEN_KEY);
    if (cached) return Promise.resolve(cached);
    if (!TEMPLATE_FOLDER_ID || !ROOT_SURVEYS_FOLDER_ID) return Promise.resolve(null);
    return new Promise(resolve => {
      let settled = false;
      const finish = v => { if (!settled) { settled = true; resolve(v); } };
      loadGisScript().then(() => {
        if (!semesterClientRef.current) {
          semesterClientRef.current = initSemesterTokenClient({
            onToken: (t, expiresIn) => { saveStoredToken(SEMESTER_TOKEN_KEY, t, expiresIn); finish(t); },
            onError: () => finish(null),
          });
        }
        semesterClientRef.current.requestAccessToken({ prompt: "" });
        setTimeout(() => finish(null), 6000);
      }).catch(() => finish(null));
    });
  }, []);

  // ── Confirmation card (used before any real, hard-to-reverse Drive/Forms action) ─
  const requestConfirm = useCallback(payload => {
    const id = Date.now() + Math.random();
    setMsgs(prev => [...prev, { id, type: "confirm_action", ...payload }]);
    return new Promise(resolve => { pendingConfirmRef.current = { id, resolve }; });
  }, []);

  const handleConfirmChoice = useCallback((msgId, ok) => {
    const pending = pendingConfirmRef.current;
    if (!pending || pending.id !== msgId) return;
    pendingConfirmRef.current = null;
    setMsgs(prev => prev.map(m => m.id === msgId ? { ...m, resolved: ok ? "confirmed" : "cancelled" } : m));
    pending.resolve(ok);
  }, []);

  // ── Tool executor ─────────────────────────────────────────────────────────
  const executeTool = useCallback(async (name, args, toolCallId) => {

    if (name === "generate_report") {
      const result = freshResRef.current || currentResRef.current;
      if (!result) return "خطأ: لا يوجد استبيان محلل. ارفع ملفاً أولاً.";
      addMsg({ type: "action", text: "⚙ جاري إنشاء التقرير..." });
      try {
        const meta = {
          year: args.year || "", program: args.program || "",
          preparedBy: args.preparedBy || "", reviewer: args.reviewer || "",
        };
        const blob = await buildAnnualDocx(result, meta, docSettings);
        const dept  = meta.program ? `_${meta.program}` : "";
        const fname = `تقرير_${result.schemaLabel}${dept}_${meta.year}.docx`;
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = fname;
        a.click();
        addMsg({ type: "success", text: `✅ تم تحميل: ${fname}` });
        return `تم إنشاء التقرير وتحميله: ${fname}`;
      } catch (e) {
        addMsg({ type: "error_inline", text: `❌ فشل: ${e.message}` });
        return `فشل إنشاء التقرير: ${e.message}`;
      }
    }

    if (name === "request_file_upload") {
      const prompt = args.message || "يرجى رفع ملف Excel للاستبيان";
      addMsg({ type: "upload_prompt", text: prompt, toolCallId });
      return new Promise((resolve, reject) => {
        pendingUploadRef.current = { resolve, reject, toolCallId, surveyType: args.type };
      });
    }

    if (name === "list_survey_files") {
      if (import.meta.env.PROD) return "استعراض مجلد محلي غير متاح في النسخة المنشورة. من فضلك ارفع الملف مباشرة.";
      const folder = aiSettingsRef.current?.surveysFolder;
      if (!folder) return "مجلد الاستبيانات غير مُعيَّن. أضفه من إعدادات المساعد الذكي.";
      addMsg({ type: "action", text: "📂 جاري استعراض الملفات..." });
      try {
        const resp = await fetch("/api/list-files", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ surveysFolder: folder, subfolder: args.subfolder || "" }),
        });
        const data = await resp.json();
        if (data.error) return `خطأ: ${data.error}`;
        if (!data.files?.length) return "لا توجد ملفات Excel في المجلد المحدد.";
        const list = data.files.map(f => `- ${f.name} (${f.path})`).join("\n");
        return `الملفات المتاحة:\n${list}`;
      } catch (e) {
        return `فشل استعراض الملفات: ${e.message}`;
      }
    }

    if (name === "analyze_file_from_folder") {
      if (import.meta.env.PROD) return "استعراض مجلد محلي غير متاح في النسخة المنشورة. من فضلك ارفع الملف مباشرة.";
      const folder = aiSettingsRef.current?.surveysFolder;
      if (!folder) return "مجلد الاستبيانات غير مُعيَّن.";
      if (!args.file_path) return "يجب تحديد مسار الملف.";
      addMsg({ type: "action", text: "⏳ جاري قراءة وتحليل الملف من المجلد..." });
      try {
        const resp = await fetch("/api/read-file", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ filePath: args.file_path, surveysFolder: folder }),
        });
        const data = await resp.json();
        if (data.error) return `خطأ: ${data.error}`;

        const binary = atob(data.data);
        const bytes  = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        const buf  = bytes.buffer;
        const rows = readExcel(buf);
        const type = detectSurveyType(data.name, rows[0]) || args.type || "faculty";
        const res  = analyze(rows, SCHEMAS[type] || SCHEMAS.faculty);
        freshResRef.current = res;
        onAnalysisCompleteRef.current?.(res, type);
        const summary = `تم تحليل "${data.name}" بنجاح\n- النوع: ${res.schemaLabel}\n- المستجيبون: ${res.n}\n- نسبة الموافقة: ${res.overallAgreePct}% (${res.overallDirection})`;
        addMsg({ type: "success", text: `✅ ${summary}` });
        return summary;
      } catch (e) {
        addMsg({ type: "error_inline", text: `❌ فشل: ${e.message}` });
        return `فشل تحليل الملف: ${e.message}`;
      }
    }

    if (name === "list_semester_survey_templates") {
      const token = await ensureSemesterToken();
      if (!token) return "يجب الاتصال بحساب Google أولاً من تبويب «📆 استبيانات الفصل الدراسي» بالتطبيق، ثم إعادة المحاولة.";
      addMsg({ type: "action", text: "📂 جاري جلب القوالب المتاحة..." });
      try {
        const files = await listFormsInFolder(token, TEMPLATE_FOLDER_ID);
        if (!files.length) return "لا توجد قوالب استبيانات داخل مجلد القوالب.";
        return `القوالب المتاحة:\n${files.map(f => `- ${f.name}`).join("\n")}`;
      } catch (e) {
        return `فشل جلب القوالب: ${e.message}`;
      }
    }

    if (name === "list_semester_surveys") {
      const token = await ensureSemesterToken();
      if (!token) return "يجب الاتصال بحساب Google أولاً من تبويب «📆 استبيانات الفصل الدراسي» بالتطبيق، ثم إعادة المحاولة.";
      if (!isValidAcademicYear(args.year)) return "صيغة السنة الدراسية غير صحيحة، يجب أن تكون مثل 2026/2027.";
      if (!SEMESTERS.includes(args.semester)) return `الفصل الدراسي يجب أن يكون واحداً من: ${SEMESTERS.join("، ")}`;
      addMsg({ type: "action", text: "📂 جاري استعراض استبيانات الفصل الدراسي..." });
      try {
        const semFolder = await findYearSemesterFolder(token, args.year, args.semester);
        if (!semFolder) return `لا يوجد مجلد استبيانات لسنة ${args.year} / ${args.semester}.`;
        const surveys = await listSemesterSurveysWithStats(token, semFolder.id);
        if (!surveys.length) return "لا توجد استبيانات لهذا الفصل بعد.";
        return `استبيانات ${args.semester} ${args.year}:\n${surveys.map(s => `- ${s.name} (${s.surveyType}) — ${s.responses} رد`).join("\n")}`;
      } catch (e) {
        return `فشل استعراض الاستبيانات: ${e.message}`;
      }
    }

    if (name === "generate_semester_surveys") {
      const token = await ensureSemesterToken();
      if (!token) return "يجب الاتصال بحساب Google أولاً من تبويب «📆 استبيانات الفصل الدراسي» بالتطبيق، ثم إعادة المحاولة.";
      if (!isValidAcademicYear(args.year)) return "صيغة السنة الدراسية غير صحيحة، يجب أن تكون مثل 2026/2027.";
      if (!SEMESTERS.includes(args.semester)) return `الفصل الدراسي يجب أن يكون واحداً من: ${SEMESTERS.join("، ")}`;

      let matched, jobList;
      try {
        const allTemplates = await listFormsInFolder(token, TEMPLATE_FOLDER_ID);
        const wantedNames = (args.templateNames || []).filter(Boolean);
        matched = wantedNames.length
          ? allTemplates.filter(t => wantedNames.some(n => t.name.includes(n)))
          : allTemplates;
        if (!matched.length) return "لم يتم العثور على أي قالب مطابق. استخدم list_semester_survey_templates لمعرفة الأسماء الصحيحة.";

        const mode = ["general", "departments", "programs"].includes(args.mode) ? args.mode : "general";
        const templates = matched.map(t => ({ ...t, mode }));
        jobList = buildGenerationJobs(templates, loadDepartments());
      } catch (e) {
        return `فشل تجهيز القوالب: ${e.message}`;
      }

      const ok = await requestConfirm({
        title: "تأكيد إنشاء استبيانات فصل دراسي",
        lines: [
          `السنة: ${args.year} — الفصل: ${args.semester}`,
          `عدد النسخ التي سيتم إنشاؤها: ${jobList.length}`,
          ...matched.map(t => `• ${t.name}`),
        ],
        confirmLabel: "✅ إنشاء الآن", cancelLabel: "✕ إلغاء",
      });
      if (!ok) return "تم إلغاء إنشاء الاستبيانات بناءً على طلب المستخدم.";

      addMsg({ type: "action", text: `⚙ جاري إنشاء ${jobList.length} استبيان على Google Drive...` });
      try {
        const { successRows, failCount } = await runGenerationJobs(token, jobList, { year: args.year, semester: args.semester });
        addMsg({ type: "success", text: `✅ تم إنشاء ${successRows.length} استبيان${failCount ? `، وفشل ${failCount}` : ""}.` });
        const lines = successRows.map(r => `- ${r.name}: ${r.formUrl}`).join("\n");
        return `تم إنشاء ${successRows.length} استبيان بنجاح${failCount ? ` (فشل ${failCount})` : ""}:\n${lines}`;
      } catch (e) {
        addMsg({ type: "error_inline", text: `❌ فشل: ${e.message}` });
        return `فشل إنشاء الاستبيانات: ${e.message}`;
      }
    }

    if (name === "analyze_semester_surveys") {
      const token = await ensureSemesterToken();
      if (!token) return "يجب الاتصال بحساب Google أولاً من تبويب «📆 استبيانات الفصل الدراسي» بالتطبيق، ثم إعادة المحاولة.";
      if (!isValidAcademicYear(args.year)) return "صيغة السنة الدراسية غير صحيحة، يجب أن تكون مثل 2026/2027.";
      if (!SEMESTERS.includes(args.semester)) return `الفصل الدراسي يجب أن يكون واحداً من: ${SEMESTERS.join("، ")}`;

      let surveys;
      try {
        const semFolder = await findYearSemesterFolder(token, args.year, args.semester);
        if (!semFolder) return `لا يوجد مجلد استبيانات لسنة ${args.year} / ${args.semester}.`;
        const allSurveys = await listSemesterSurveysWithStats(token, semFolder.id);
        const wantedNames = (args.surveyNames || []).filter(Boolean);
        surveys = wantedNames.length
          ? allSurveys.filter(s => wantedNames.some(n => s.name.includes(n)))
          : allSurveys;
        if (!surveys.length) return "لم يتم العثور على أي استبيان مطابق. استخدم list_semester_surveys لمعرفة الأسماء المتاحة.";
      } catch (e) {
        return `فشل جلب الاستبيانات: ${e.message}`;
      }

      const ok = await requestConfirm({
        title: "تأكيد تحليل وتحميل التقارير",
        lines: [
          `السنة: ${args.year} — الفصل: ${args.semester}`,
          `عدد الاستبيانات: ${surveys.length}`,
          ...surveys.map(s => `• ${s.name} (${s.responses} رد)`),
        ],
        confirmLabel: "✅ تحليل وتحميل", cancelLabel: "✕ إلغاء",
      });
      if (!ok) return "تم إلغاء التحليل بناءً على طلب المستخدم.";

      addMsg({ type: "action", text: `⏳ جاري تحليل ${surveys.length} استبيان وتحميل التقارير...` });
      const schemas = getAllAnalysisSchemas();
      let successCount = 0, skipCount = 0, errorCount = 0;
      const summaries = [];
      for (const survey of surveys) {
        try {
          const [form, responses] = await Promise.all([getForm(token, survey.id), listAllResponses(token, survey.id)]);
          const rows = responsesToRows(form, responses);
          const schemaId = detectAnySurveyType(survey.name, rows[0]);
          const schema = schemaId ? schemas[schemaId] : null;
          if (!schema) { skipCount++; summaries.push(`- ${survey.name}: تعذّر التعرف على النوع`); continue; }
          const result = analyze(rows, schema);
          const meta = { year: args.year, program: departmentFromSurveyName(survey.name), preparedBy: args.preparedBy || "", reviewer: args.reviewer || "" };
          const blob = await buildAnnualDocx(result, meta, docSettings);
          const fname = `تقرير_${schema.label}_${survey.name}.docx`;
          const a = document.createElement("a");
          a.href = URL.createObjectURL(blob);
          a.download = fname;
          a.click();
          successCount++;
          summaries.push(`- ${survey.name}: ✅ ${fname}`);
          await new Promise(r => setTimeout(r, 450)); // let the browser process each download separately
        } catch (e) {
          errorCount++;
          summaries.push(`- ${survey.name}: ❌ ${e.message}`);
        }
      }
      addMsg({ type: "success", text: `✅ تم تحليل وتحميل ${successCount} تقرير${skipCount ? `، وتخطي ${skipCount}` : ""}${errorCount ? `، وفشل ${errorCount}` : ""}.` });
      return `النتيجة:\n${summaries.join("\n")}`;
    }

    return `أداة غير معروفة: ${name}`;
  }, [addMsg, docSettings, ensureSemesterToken, requestConfirm]);

  // ── File upload handler (for request_file_upload tool) ────────────────────
  const handleUploadedFile = useCallback(async file => {
    const pending = pendingUploadRef.current;
    if (!pending) return;
    pendingUploadRef.current = null;
    addMsg({ type: "action", text: "⏳ جاري تحليل الملف..." });
    try {
      const buf  = await file.arrayBuffer();
      const rows = readExcel(buf);
      const type = detectSurveyType(file.name, rows[0]) || pending.surveyType || "faculty";
      const res  = analyze(rows, SCHEMAS[type] || SCHEMAS.faculty);
      freshResRef.current = res;
      onAnalysisComplete?.(res, type);
      const summary = `تم تحليل "${file.name}" بنجاح\n- النوع: ${res.schemaLabel}\n- المستجيبون: ${res.n}\n- نسبة الموافقة: ${res.overallAgreePct}% (${res.overallDirection})`;
      addMsg({ type: "success", text: `✅ ${summary}` });
      pending.resolve(summary);
    } catch (e) {
      addMsg({ type: "error_inline", text: `❌ فشل تحليل الملف: ${e.message}` });
      pending.reject(e);
    }
  }, [addMsg, onAnalysisComplete]);

  // ── API call ──────────────────────────────────────────────────────────────
  const callApi = useCallback(async messages => {
    const resp = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        provider:     aiSettings.provider,
        model:        aiSettings.model,
        apiKey:       aiSettings.key,
        messages,
        tools:        TOOLS,
        systemPrompt: buildSystemPrompt(currentResRef.current, aiSettings.surveysFolder),
      }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || "خطأ في الاتصال");
    return data;
  }, [aiSettings]);

  // ── Conversation loop (handles tool_calls recursion) ─────────────────────
  const runLoop = useCallback(async history => {
    let raw = [...history];
    while (true) {
      const data   = await callApi(raw);
      const choice = data.choices?.[0];
      if (!choice) throw new Error("استجابة غير صالحة من النموذج");

      const msg = choice.message;
      const fr  = (choice.finish_reason || "").toLowerCase();

      if ((fr === "tool_calls" || fr === "function_call") && msg.tool_calls?.length) {
        raw.push({ role: "assistant", content: msg.content || null, tool_calls: msg.tool_calls });
        for (const tc of msg.tool_calls) {
          const args   = JSON.parse(tc.function.arguments || "{}");
          const result = await executeTool(tc.function.name, args, tc.id);
          raw.push({ role: "tool", tool_call_id: tc.id, content: String(result) });
        }
      } else {
        const text = msg.content || "";
        raw.push({ role: "assistant", content: text });
        addMsg({ type: "assistant", text });
        break;
      }
    }
    return raw;
  }, [callApi, executeTool, addMsg]);

  // ── Send message ──────────────────────────────────────────────────────────
  const send = useCallback(async (text) => {
    const t = text.trim();
    if (!t || loading) return;
    setInput("");
    setLoading(true);
    addMsg({ type: "user", text: t });
    const history = [...rawHistoryRef.current, { role: "user", content: t }];
    try {
      rawHistoryRef.current = await runLoop(history);
    } catch (e) {
      addMsg({ type: "error_inline", text: `❌ ${e.message}` });
      rawHistoryRef.current = history;
    } finally {
      setLoading(false);
      setTimeout(() => textareaRef.current?.focus(), 60);
    }
  }, [loading, addMsg, runLoop]);

  const handleKey = e => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(input); }
  };

  const prov = PROVIDERS[aiSettings.provider] || PROVIDERS.groq;

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <>
      <style>{`
        @keyframes aichat-bounce {
          0%,60%,100%{transform:translateY(0);opacity:.5}
          30%{transform:translateY(-6px);opacity:1}
        }
        .aichat-textarea:focus{border-color:#1abc9c!important;background:rgba(255,255,255,.12)!important}
      `}</style>

      {/* Hidden file input */}
      <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
        onChange={e => { if (e.target.files[0]) handleUploadedFile(e.target.files[0]); e.target.value = ""; }} />

      {/* Floating button */}
      <button
        title="المساعد الذكي"
        onClick={() => setIsOpen(v => !v)}
        style={{
          position: "fixed", bottom: "max(20px, env(safe-area-inset-bottom))", left: 20, zIndex: 1100,
          width: 54, height: 54, borderRadius: "50%", border: "none", cursor: "pointer",
          background: `linear-gradient(135deg,${prov.color},#1a3a5c)`,
          boxShadow: "0 4px 20px rgba(0,0,0,.45)",
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 22, transition: "transform .25s, box-shadow .25s",
          transform: isOpen ? "scale(1.08)" : "scale(1)",
        }}
      >
        {isOpen ? "✕" : "💬"}
      </button>

      {/* Chat panel */}
      {isOpen && (
        <div style={{
          position: "fixed", bottom: "max(86px, calc(env(safe-area-inset-bottom) + 86px))", left: 20, right: 20, zIndex: 1099,
          width: "auto", maxWidth: 420, height: "min(565px, calc(100dvh - 120px))",
          margin: "0 auto",
          background: "linear-gradient(180deg,#0e1f32 0%,#0b1a28 100%)",
          border: "1px solid rgba(255,255,255,.13)",
          borderRadius: 20, boxShadow: "0 12px 48px rgba(0,0,0,.55)",
          display: "flex", flexDirection: "column", overflow: "hidden",
          direction: "rtl",
        }}>

          {/* Header */}
          <div style={{
            padding: "12px 16px", borderBottom: "1px solid rgba(255,255,255,.09)",
            background: "rgba(255,255,255,.04)",
            display: "flex", alignItems: "center", justifyContent: "space-between",
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{
                width: 32, height: 32, borderRadius: "50%", fontSize: 16,
                background: `linear-gradient(135deg,${prov.color},#1a3a5c)`,
                display: "flex", alignItems: "center", justifyContent: "center",
              }}>{prov.icon}</div>
              <div>
                <div style={{ color: "#fff", fontWeight: 700, fontSize: 13.5 }}>المساعد الذكي</div>
                <div style={{ color: "rgba(255,255,255,.38)", fontSize: 11 }}>
                  {prov.label} · {aiSettings.model.split("-").slice(0, 2).join("-")}
                </div>
              </div>
            </div>
            {msgs.length > 0 && (
              <button
                onClick={() => { setMsgs([]); rawHistoryRef.current = []; freshResRef.current = null; }}
                style={{ background: "none", border: "none", color: "rgba(255,255,255,.3)", cursor: "pointer", fontSize: 11, fontFamily: "'Cairo',sans-serif", padding: "4px 8px", borderRadius: 6 }}
              >مسح ✕</button>
            )}
          </div>

          {/* Messages */}
          <div style={{ flex: 1, overflowY: "auto", padding: "14px 12px", display: "flex", flexDirection: "column", gap: 10 }}>

            {/* Empty state */}
            {msgs.length === 0 && (
              <div style={{ textAlign: "center", color: "rgba(255,255,255,.28)", fontSize: 12.5, marginTop: 36, lineHeight: 2.2 }}>
                <div style={{ fontSize: 32, marginBottom: 10 }}>💬</div>
                <div>اسألني عن نتائج الاستبيان</div>
                <div>أو قل <span style={{ color: "#1abc9c" }}>"اعمل تقرير 2024-2025"</span></div>
                {!currentResult && <div>أو <span style={{ color: "#1abc9c" }}>ارفع لي ملف استبيان</span></div>}
              </div>
            )}

            {msgs.map(m => (
              <div key={m.id} style={{ display: "flex", direction: "ltr",
                justifyContent:
                  m.type === "user" ? "flex-end" :
                  (m.type === "upload_prompt" || m.type === "confirm_action") ? "center" : "flex-start",
              }}>

                {/* User bubble */}
                {m.type === "user" && (
                  <div style={{
                    maxWidth: "78%", background: "rgba(26,188,156,.15)", border: "1px solid rgba(26,188,156,.3)",
                    borderRadius: "16px 4px 16px 16px", padding: "9px 13px",
                    color: "#e8f0fe", direction: "rtl", textAlign: "right",
                  }}><MsgText text={m.text} /></div>
                )}

                {/* Assistant bubble */}
                {m.type === "assistant" && (
                  <div style={{
                    maxWidth: "86%", background: "rgba(255,255,255,.07)", border: "1px solid rgba(255,255,255,.1)",
                    borderRadius: "4px 16px 16px 16px", padding: "9px 13px",
                    color: "#e8f0fe", direction: "rtl", textAlign: "right",
                  }}><MsgText text={m.text} /></div>
                )}

                {/* Action / success / error inline */}
                {(m.type === "action" || m.type === "success" || m.type === "error_inline") && (
                  <div style={{
                    fontSize: 12, borderRadius: 10, padding: "6px 12px", direction: "rtl", textAlign: "right",
                    background: "rgba(255,255,255,.04)", border: "1px solid rgba(255,255,255,.07)",
                    color: m.type === "error_inline" ? "#e74c3c" : m.type === "success" ? "#1abc9c" : "rgba(255,255,255,.5)",
                    maxWidth: "90%",
                  }}><MsgText text={m.text} /></div>
                )}

                {/* File upload prompt */}
                {m.type === "upload_prompt" && (
                  <div style={{
                    background: "rgba(26,188,156,.09)", border: "1px solid rgba(26,188,156,.28)",
                    borderRadius: 14, padding: "14px 18px", textAlign: "center", maxWidth: "88%", direction: "rtl",
                  }}>
                    <div style={{ color: "#e8f0fe", fontSize: 13, marginBottom: 10 }}>{m.text}</div>
                    <button
                      onClick={() => fileInputRef.current?.click()}
                      style={{
                        background: "linear-gradient(135deg,#1abc9c,#16a085)", color: "#fff", border: "none",
                        borderRadius: 20, padding: "8px 22px", cursor: "pointer",
                        fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 13,
                      }}
                    >📂 اختر الملف</button>
                  </div>
                )}

                {/* Confirmation card — required before any real Drive/Forms side effect */}
                {m.type === "confirm_action" && (
                  <div style={{
                    background: "rgba(255,193,7,.08)", border: "1px solid rgba(255,193,7,.3)",
                    borderRadius: 14, padding: "14px 16px", maxWidth: "92%", direction: "rtl", textAlign: "right",
                  }}>
                    <div style={{ color: "#ffd166", fontWeight: 700, fontSize: 13, marginBottom: 8 }}>{m.title}</div>
                    <div style={{ color: "#e8f0fe", fontSize: 12.5, lineHeight: 1.9, whiteSpace: "pre-wrap", marginBottom: 10 }}>
                      {m.lines?.join("\n")}
                    </div>
                    {m.resolved ? (
                      <div style={{ fontSize: 12, color: m.resolved === "confirmed" ? "#1abc9c" : "rgba(255,255,255,.5)" }}>
                        {m.resolved === "confirmed" ? "✅ تم التأكيد" : "✕ تم الإلغاء"}
                      </div>
                    ) : (
                      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                        <button
                          onClick={() => handleConfirmChoice(m.id, true)}
                          style={{
                            background: "linear-gradient(135deg,#1abc9c,#16a085)", color: "#fff", border: "none",
                            borderRadius: 16, padding: "7px 16px", cursor: "pointer",
                            fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 12.5,
                          }}
                        >{m.confirmLabel || "تأكيد"}</button>
                        <button
                          onClick={() => handleConfirmChoice(m.id, false)}
                          style={{
                            background: "rgba(255,255,255,.08)", color: "#e8f0fe", border: "1px solid rgba(255,255,255,.18)",
                            borderRadius: 16, padding: "7px 16px", cursor: "pointer",
                            fontFamily: "'Cairo',sans-serif", fontSize: 12.5,
                          }}
                        >{m.cancelLabel || "إلغاء"}</button>
                      </div>
                    )}
                  </div>
                )}

              </div>
            ))}

            {/* Typing indicator */}
            {loading && (
              <div style={{ display: "flex", direction: "ltr", justifyContent: "flex-start" }}>
                <div style={{
                  background: "rgba(255,255,255,.07)", border: "1px solid rgba(255,255,255,.1)",
                  borderRadius: "4px 16px 16px 16px", padding: "8px 14px",
                }}>
                  <TypingDots />
                </div>
              </div>
            )}

            <div ref={bottomRef} />
          </div>

          {/* Input */}
          <div style={{
            padding: "10px 12px", borderTop: "1px solid rgba(255,255,255,.07)",
            background: "rgba(255,255,255,.03)", display: "flex", gap: 8, alignItems: "flex-end",
          }}>
            <textarea
              ref={textareaRef}
              className="aichat-textarea"
              value={input}
              onChange={e => setInput(e.target.value)}
              onKeyDown={handleKey}
              placeholder="اكتب رسالة… (Enter للإرسال)"
              rows={1}
              style={{
                flex: 1, background: "rgba(255,255,255,.07)", border: "1px solid rgba(255,255,255,.14)",
                borderRadius: 12, padding: "9px 13px", color: "#fff",
                fontFamily: "'Cairo',sans-serif", fontSize: 13, outline: "none",
                resize: "none", direction: "rtl", maxHeight: 100, lineHeight: 1.55,
                transition: "border-color .2s, background .2s",
              }}
              onInput={e => {
                e.target.style.height = "auto";
                e.target.style.height = Math.min(e.target.scrollHeight, 100) + "px";
              }}
            />
            <button
              onClick={() => send(input)}
              disabled={!input.trim() || loading}
              style={{
                width: 38, height: 38, borderRadius: "50%", border: "none", cursor: "pointer",
                background: input.trim() && !loading ? "linear-gradient(135deg,#1abc9c,#16a085)" : "rgba(255,255,255,.1)",
                color: "#fff", fontSize: 17, display: "flex", alignItems: "center", justifyContent: "center",
                flexShrink: 0, transition: "background .2s",
              }}
            >↑</button>
          </div>

        </div>
      )}
    </>
  );
}
