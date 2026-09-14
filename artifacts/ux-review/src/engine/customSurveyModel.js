// Data model + storage for the Dynamic Survey Management system.
//
// This module does NOT modify analyze.js, buildDocx.js, or the existing
// src/schemas/*.yaml definitions — it only reads SCHEMAS (read-only) to
// merge custom surveys alongside them. toAnalysisSchema() converts a custom
// survey into the exact shape analyze.js's analyze()/detectSurveyType-style
// logic expects (see src/schemas/compileSchema.js for the equivalent
// conversion the existing YAML schemas go through at build time), so the
// generic, schema-driven analyze() function works unchanged for custom
// surveys too.

import * as XLSX from "xlsx";
import { SCHEMAS as BUILTIN_SCHEMAS } from "../schemas/index.js";

export const STORAGE_KEY = "eruQA_customSurveys_v1";
export const BUILTIN_REPORT_TEXTS_KEY = "eruQA_builtinReportTexts_v1";

export const SCALE_TYPES = {
  "likert-3": {
    label: "مقياس ثلاثي (3 درجات)",
    agreementCodes: ["agree"],
    values: [
      { code: "agree", label: "أوافق", score: 3 },
      { code: "neutral", label: "محايد", score: 2 },
      { code: "disagree", label: "لا أوافق", score: 1 },
    ],
  },
  "likert-5": {
    label: "مقياس خماسي (5 درجات)",
    agreementCodes: ["4", "5"],
    values: [
      { code: "5", label: "أوافق بشدة", score: 5 },
      { code: "4", label: "أوافق", score: 4 },
      { code: "3", label: "محايد", score: 3 },
      { code: "2", label: "لا أوافق", score: 2 },
      { code: "1", label: "لا أوافق بشدة", score: 1 },
    ],
  },
};

// Same default interpretation tiers compileSchema.js assigns to YAML schemas
// that don't declare their own — kept in sync manually since that file
// doesn't export them. analyze.js's interpret() requires this shape.
const DEFAULT_INTERPRETATION = {
  "likert-5": [
    { min: 4.5, label: "أوافق بشدة", tier: "excellent", color: "0d6e3a" },
    { min: 3.5, label: "أوافق", tier: "good", color: "1a5276" },
    { min: 2.5, label: "محايد", tier: "neutral", color: "784212" },
    { min: 1.0, label: "لا أوافق", tier: "low", color: "922b21" },
  ],
  "likert-3": [
    { min: 85, label: "موافقة قوية", tier: "excellent", color: "0d6e3a" },
    { min: 70, label: "موافقة", tier: "good", color: "1a5276" },
    { min: 50, label: "محايد", tier: "neutral", color: "784212" },
    { min: 0, label: "عدم الموافقة", tier: "low", color: "922b21" },
  ],
};

export const QUESTION_TYPES = [
  { id: "likert", label: "مقياس (يتبع مقياس الاستبيان)" },
  { id: "text", label: "نص مفتوح" },
  { id: "numeric", label: "رقمي (دعم مستقبلي)" },
];

export const SURVEY_STATUSES = [
  { id: "draft", label: "مسودة" },
  { id: "active", label: "نشط" },
];

function genId(prefix) {
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
}

export function createQuestion(overrides = {}) {
  return {
    id: genId("q"),
    text: "",
    excelColumn: "",
    type: "likert",
    required: true,
    category: "",
    weight: null,
    ...overrides,
  };
}

export function createSection(overrides = {}) {
  return {
    id: genId("sec"),
    name: "",
    description: "",
    questions: [],
    ...overrides,
  };
}

export function createMetadata(overrides = {}) {
  return {
    timestampCol: [],
    emailCol: [],
    nameCol: [],
    degreeCol: [],
    departmentCol: [],
    freeTextCols: [],
    ...overrides,
  };
}

export function createReportTexts(overrides = {}) {
  return {
    reportTitle: "",
    reportSubtitle: "",
    introduction: "",
    variablesText: "",
    methodologyText: "",
    resultsHeading: "",
    summaryHeading: "",
    recommendationsHeading: "",
    noRecommendationsText: "",
    recommendationTemplate: "",
    ...overrides,
  };
}

export function createSurvey(overrides = {}) {
  const now = new Date().toISOString();
  return {
    id: genId("survey"),
    name: "",
    description: "",
    version: "1.0.0",
    status: "draft",
    surveyType: "",
    scaleType: "likert-5",
    sections: [],
    metadata: createMetadata(),
    reportTexts: createReportTexts(),
    createdAt: now,
    updatedAt: now,
    ...overrides,
  };
}

function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

// ── Storage (localStorage, separate from the existing eruQA_settings_v1 / eruQA_ai_v1 keys) ──

export function loadCustomSurveys() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const list = JSON.parse(raw);
      return list.map(s => ({ ...s, metadata: createMetadata(s.metadata), reportTexts: createReportTexts(s.reportTexts) }));
    }
  } catch { /* ignore */ }
  return [];
}

function persist(list) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(list)); } catch { /* ignore */ }
}

// If a question's Excel column name was left blank, default it to the
// question text itself so the survey is still usable without that field.
function normalizeQuestionColumns(survey) {
  return {
    ...survey,
    sections: (survey.sections ?? []).map(sec => ({
      ...sec,
      questions: (sec.questions ?? []).map(q => ({
        ...q,
        excelColumn: (q.excelColumn && q.excelColumn.trim()) || q.text,
      })),
    })),
  };
}

export function saveCustomSurvey(survey) {
  const list = loadCustomSurveys();
  const idx = list.findIndex(s => s.id === survey.id);
  const updated = { ...normalizeQuestionColumns(survey), updatedAt: new Date().toISOString() };
  if (idx >= 0) list[idx] = updated;
  else list.push(updated);
  persist(list);
  return updated;
}

export function deleteCustomSurvey(id) {
  persist(loadCustomSurveys().filter(s => s.id !== id));
}

export function duplicateCustomSurvey(id) {
  const list = loadCustomSurveys();
  const src = list.find(s => s.id === id);
  if (!src) return null;
  const now = new Date().toISOString();
  const copy = {
    ...deepClone(src),
    id: genId("survey"),
    name: `${src.name} (نسخة)`,
    status: "draft",
    createdAt: now,
    updatedAt: now,
  };
  list.push(copy);
  persist(list);
  return copy;
}

export function setCustomSurveyStatus(id, status) {
  const list = loadCustomSurveys();
  const idx = list.findIndex(s => s.id === id);
  if (idx < 0) return null;
  list[idx] = { ...list[idx], status, updatedAt: new Date().toISOString() };
  persist(list);
  return list[idx];
}

// ── JSON backup (export/import) ───────────────────────────────────────────────
// localStorage is scoped per browser origin — including the port the dev
// server happens to be running on. Running `vite --port X` with a different
// port each time makes saved surveys "disappear" because it's really a
// different storage bucket, not data loss. This backup lets surveys be
// restored regardless of origin/port/browser.

export function buildSurveysBackupBlob() {
  const list = loadCustomSurveys();
  return new Blob([JSON.stringify(list, null, 2)], { type: "application/json" });
}

// Merges incoming surveys into the existing list (by id) and persists.
// Returns the number of surveys in the resulting list.
export function importSurveysBackup(jsonText) {
  let incoming;
  try { incoming = JSON.parse(jsonText); } catch { throw new Error("الملف ليس JSON صالحاً."); }
  if (!Array.isArray(incoming)) throw new Error("تنسيق الملف غير صالح — يجب أن يكون مصفوفة استبيانات.");
  const byId = new Map(loadCustomSurveys().map(s => [s.id, s]));
  for (const s of incoming) byId.set(s.id, { ...s, metadata: createMetadata(s.metadata), reportTexts: createReportTexts(s.reportTexts) });
  const merged = [...byId.values()];
  persist(merged);
  return merged.length;
}

// ── Import from Excel: auto-detect questions vs. general-info columns ────────
// Works on the raw rows produced by analyze.js's readExcel() (2D array, row 0
// = headers). This is a standalone heuristic — it does not call into
// analyze.js's detection logic — so it has no effect on the existing engine.

function normalizeHeader(v) {
  if (v == null) return "";
  return String(v)
    .trim()
    .replace(/[أإآ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .replace(/[\s_ـ]+/g, " ")
    .toLowerCase();
}

const META_KEYWORDS = {
  timestamp: ["timestamp"],
  email: ["email", "البريد", "بريد"],
  name: ["الاسم", "اسمك", "name"],
  degree: ["الوظيفه", "الدرجه", "degree", "الدرجه العلميه"],
  department: ["القسم", "التخصص", "department"],
  freetext: ["مقترح", "ملاحظ", "اقتراح"],
};

// Metadata columns (name/email/degree/department/free-text) are short labels.
// Real survey questions are full sentences. Matching keywords as a substring
// of the whole header (without this guard) misclassifies long questions that
// merely mention a keyword mid-sentence — e.g. "يوجد توصيف وظيفي به تحديد
// للسلطات ومسئوليات الوظيفة" contains "الوظيفة" but is a Likert question, not
// a job-title column. Real metadata headers observed in sample exports
// ("القسم", "الدرجة العلمية", "Email Address") are all <= 5 words.
function isShortLabel(header) {
  const words = String(header).trim().split(/\s+/).filter(Boolean);
  return words.length > 0 && words.length <= 5;
}

function classifyHeader(header, colIdx, isLikertColumn) {
  const n = normalizeHeader(header);
  if (!n) return null;
  if (colIdx === 0 && n.includes("timestamp")) return "timestamp";
  // A column whose answers are actually Likert-scale values is always a real
  // question, regardless of what its header text happens to contain.
  if (isLikertColumn) return "question";
  if (!isShortLabel(header)) return "question";
  for (const [key, words] of Object.entries(META_KEYWORDS)) {
    if (key === "timestamp") continue;
    if (words.some(w => n.includes(normalizeHeader(w)))) return key;
  }
  return "question";
}

const LIKERT3_WORDS = ["اوافق", "موافق", "محايد", "لا اوافق"];

function detectAnswerType(samples) {
  const vals = samples.filter(v => v != null && String(v).trim() !== "");
  if (!vals.length) return { type: "likert", scale: null };
  if (vals.every(v => /^\(\d\)/.test(String(v).trim()))) return { type: "likert", scale: "likert-5" };
  if (vals.every(v => LIKERT3_WORDS.some(w => normalizeHeader(v).includes(w)))) return { type: "likert", scale: "likert-3" };
  if (vals.every(v => !Number.isNaN(Number(v)))) return { type: "numeric", scale: null };
  return { type: "text", scale: null };
}

// Returns { metadata, scaleType, questions } or null if rows are empty.
// `questions` are all returned flat — the caller groups them into a single
// section, since a flat Excel header row carries no axis/section grouping
// information (verified against real survey exports: one header row, no
// merged section markers). The user splits them into multiple sections
// afterward using the existing add/move-question UI.
export function importSurveyStructureFromRows(rows) {
  if (!rows || !rows.length) return null;
  const headers = rows[0] ?? [];
  const sampleRows = rows.slice(1, 6);
  const metadata = createMetadata();
  const questions = [];
  const scaleVotes = { "likert-5": 0, "likert-3": 0 };

  headers.forEach((header, colIdx) => {
    if (header == null || String(header).trim() === "") return;
    const text = String(header).trim();
    const samples = sampleRows.map(r => r?.[colIdx]);
    const { type, scale } = detectAnswerType(samples);
    const isLikertColumn = type === "likert" && scale != null;

    const cls = classifyHeader(text, colIdx, isLikertColumn);
    if (cls === "timestamp") return metadata.timestampCol.push(text);
    if (cls === "email") return metadata.emailCol.push(text);
    if (cls === "name") return metadata.nameCol.push(text);
    if (cls === "degree") return metadata.degreeCol.push(text);
    if (cls === "department") return metadata.departmentCol.push(text);
    if (cls === "freetext") return metadata.freeTextCols.push(text);

    if (scale) scaleVotes[scale] = (scaleVotes[scale] ?? 0) + 1;
    questions.push(createQuestion({ text, excelColumn: text, type, required: type === "likert" }));
  });

  const scaleType = scaleVotes["likert-3"] > scaleVotes["likert-5"] ? "likert-3" : "likert-5";
  return { metadata, scaleType, questions };
}

// ── Survey-definition Excel template (download + import) ─────────────────────
// Unlike importSurveyStructureFromRows() above (which reads a real *response*
// export and has no axis information), this template format lets the admin
// define the full survey — name, scale, sections/axes, and their questions —
// offline in Excel, then import it in one shot. A blank "المحور" column on a
// question row means that survey has no axes; all such rows are grouped into
// one fallback section so the data model (which is section-based) still
// works for flat surveys like the existing `graduates` schema.

const TEMPLATE_INFO_SHEET = "معلومات الاستبيان";
const TEMPLATE_QUESTIONS_SHEET = "المحاور والأسئلة";

export function buildSurveyTemplateBlob() {
  const infoRows = [
    ["الحقل", "القيمة"],
    ["اسم الاستبيان", ""],
    ["وصف الاستبيان", ""],
    ["الإصدار", "1.0.0"],
    ["نوع الاستبيان", ""],
    ["نوع المقياس (اكتب: مقياس ثلاثي أو مقياس خماسي)", "مقياس خماسي"],
    ["عنوان التقرير (اختياري)", ""],
    ["العنوان الفرعي للتقرير (اختياري)", ""],
    ["مقدمة التقرير (اختياري)", ""],
  ];
  const questionRows = [
    ["المحور (اتركه فارغاً إن لم يوجد محاور لهذا الاستبيان)", "السؤال *", "اسم العمود في Excel (اختياري)", "نوع السؤال (مقياس / نص / رقمي)", "إلزامي (نعم/لا)", "الفئة (اختياري)", "الوزن (اختياري)"],
    ["المحور الأول", "مثال: هل أنت راضٍ عن الخدمة؟ — احذف صفوف المثال وأضف أسئلتك", "", "مقياس", "نعم", "", ""],
    ["المحور الأول", "مثال: هل توصي بالبرنامج لغيرك؟", "", "مقياس", "نعم", "", ""],
    ["المحور الثاني", "مثال: سؤال في محور مختلف", "", "مقياس", "نعم", "", ""],
    ["", "مثال: سؤال بدون محور — لاستبيان لا يحتوي على محاور اترك عمود المحور فارغاً لكل الأسئلة", "", "نص", "لا", "", ""],
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(infoRows), TEMPLATE_INFO_SHEET);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(questionRows), TEMPLATE_QUESTIONS_SHEET);
  const arr = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return new Blob([arr], { type: "application/octet-stream" });
}

const INFO_LABEL_MAP = [
  { key: "name", words: ["اسم الاستبيان"] },
  { key: "description", words: ["وصف الاستبيان", "وصف"] },
  { key: "version", words: ["الاصدار"] },
  { key: "surveyType", words: ["نوع الاستبيان"] },
  { key: "scaleType", words: ["نوع المقياس", "المقياس"] },
  { key: "reportTitle", words: ["عنوان التقرير"] },
  { key: "reportSubtitle", words: ["العنوان الفرعي للتقرير", "العنوان الفرعي"] },
  { key: "introduction", words: ["مقدمة التقرير", "المقدمة"] },
];

function parseTemplateInfoRows(rows) {
  const info = { name: "", description: "", version: "1.0.0", surveyType: "", scaleType: "likert-5", reportTitle: "", reportSubtitle: "", introduction: "" };
  for (const row of rows ?? []) {
    const label = normalizeHeader(row?.[0]);
    const rawValue = row?.[1];
    if (!label || rawValue == null || String(rawValue).trim() === "") continue;
    const value = String(rawValue).trim();
    const field = INFO_LABEL_MAP.find(f => f.words.some(w => label.includes(normalizeHeader(w))));
    if (!field) continue;
    if (field.key === "scaleType") info.scaleType = normalizeHeader(value).includes("ثلاثي") ? "likert-3" : "likert-5";
    else info[field.key] = value;
  }
  return info;
}

function parseTemplateRequired(v) {
  const n = normalizeHeader(v);
  if (!n) return true;
  return !["لا", "no", "false", "0"].includes(n);
}

function parseTemplateQuestionType(v) {
  const n = normalizeHeader(v);
  if (n.includes("نص")) return "text";
  if (n.includes("رقم")) return "numeric";
  return "likert";
}

const NO_AXIS_KEY = "__no_axis__";

function parseTemplateQuestionRows(rows) {
  const sectionsByKey = new Map();
  const order = [];
  for (const row of (rows ?? []).slice(1)) {
    const text = row?.[1] != null ? String(row[1]).trim() : "";
    if (!text) continue;
    const axisName = row?.[0] != null ? String(row[0]).trim() : "";
    const excelColumn = row?.[2] != null ? String(row[2]).trim() : "";
    const type = parseTemplateQuestionType(row?.[3]);
    const required = parseTemplateRequired(row?.[4]);
    const category = row?.[5] != null ? String(row[5]).trim() : "";
    const weightRaw = row?.[6];
    const weight = weightRaw != null && String(weightRaw).trim() !== "" && !Number.isNaN(Number(weightRaw)) ? Number(weightRaw) : null;

    const key = axisName || NO_AXIS_KEY;
    if (!sectionsByKey.has(key)) {
      sectionsByKey.set(key, createSection({ name: axisName || "الأسئلة" }));
      order.push(key);
    }
    sectionsByKey.get(key).questions.push(createQuestion({ text, excelColumn, type, required, category, weight }));
  }
  return order.map(k => sectionsByKey.get(k));
}

export function buildSurveyFromTemplateRows(infoRows, questionRows) {
  const info = parseTemplateInfoRows(infoRows);
  const sections = parseTemplateQuestionRows(questionRows);
  return createSurvey({
    name: info.name,
    description: info.description,
    version: info.version || "1.0.0",
    surveyType: info.surveyType,
    scaleType: info.scaleType,
    reportTexts: createReportTexts({ reportTitle: info.reportTitle, reportSubtitle: info.reportSubtitle, introduction: info.introduction }),
    sections,
  });
}

// Reads a full template workbook (both sheets) from an uploaded file's
// arrayBuffer and returns a ready-to-save survey object.
export function importSurveyFromTemplateArrayBuffer(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  const infoSheetName = wb.SheetNames.includes(TEMPLATE_INFO_SHEET) ? TEMPLATE_INFO_SHEET : wb.SheetNames[0];
  const qSheetName = wb.SheetNames.includes(TEMPLATE_QUESTIONS_SHEET) ? TEMPLATE_QUESTIONS_SHEET : (wb.SheetNames[1] ?? wb.SheetNames[0]);
  const infoRows = XLSX.utils.sheet_to_json(wb.Sheets[infoSheetName], { header: 1, defval: "" });
  const questionRows = XLSX.utils.sheet_to_json(wb.Sheets[qSheetName], { header: 1, defval: "" });
  return buildSurveyFromTemplateRows(infoRows, questionRows);
}

// ── Wiring into the analysis engine ────────────────────────────────────────────
// Converts a custom survey into a schema object shaped exactly like what
// src/schemas/compileSchema.js produces for the hardcoded YAML surveys, so
// analyze.js's analyze()/analyzeRows() — completely schema-driven, no
// knowledge of where a schema came from — works on it unmodified.
//
// Important boundary: analyze.js only knows how to statistically score
// Likert-scale columns (mean/%/agreement). It has no concept of "text" or
// "numeric" question types. Those question types are therefore excluded
// from the schema handed to analyze() — they simply don't appear in the
// generated statistics/report, the same way open-ended "freeTextCols" in
// the existing YAML schemas are excluded already. This is a real limitation
// of reusing analyze() as-is (which we were asked not to modify), not a bug.
//
// Likert-5 caveat: analyze.js maps likert-5 questions to file columns purely
// by position (colIndex into the columns left after removing metadata
// columns) — never by column name — exactly like the hardcoded YAML schemas
// already do. So colIndex is assigned here in current section/question
// order. Reordering questions/sections in the editor changes which file
// column each question reads from for likert-5 surveys. Likert-3 surveys
// are unaffected by reordering: analyze.js fuzzy-matches likert-3 questions
// against the actual header text, so matching follows the question's text,
// not its position.
export function toAnalysisSchema(survey) {
  const scaleType = survey.scaleType === "likert-3" ? "likert-3" : "likert-5";
  const scaleDef = SCALE_TYPES[scaleType];
  const metadata = survey.metadata ?? createMetadata();

  let autoColIndex = 0;
  let globalSeq = 1;
  const axes = (survey.sections ?? [])
    .map(sec => {
      const questions = (sec.questions ?? [])
        .filter(q => q.type === "likert")
        .map(q => {
          const out = { id: q.id, seq: globalSeq++, text: q.text };
          if (scaleType === "likert-5") out.colIndex = autoColIndex++;
          return out;
        });
      return questions.length ? { id: sec.id, name: sec.name || "محور", recommendation: null, questions } : null;
    })
    .filter(Boolean);

  const schema = {
    id: survey.id,
    label: survey.name || "(بدون اسم)",
    fileHints: [survey.name, survey.surveyType].filter(Boolean),
    scale: { type: scaleType, values: scaleDef.values, agreementCodes: scaleDef.agreementCodes },
    metadata,
    reportTexts: createReportTexts(survey.reportTexts),
    interpretation: DEFAULT_INTERPRETATION[scaleType],
    axes,
  };

  if (scaleType === "likert-3") {
    // Mirrors compileSchema.js's questionStartIndex heuristic: count the
    // leading metadata columns (timestamp/name/email/degree/department) that
    // come before the question columns. Free-text columns are excluded since
    // they sit at the end of the row, not the start.
    schema.questionStartIndex =
      (metadata.timestampCol?.length ? 1 : 0) +
      (metadata.nameCol?.length ? 1 : 0) +
      (metadata.emailCol?.length ? 1 : 0) +
      (metadata.degreeCol?.length ? 1 : 0) +
      (metadata.departmentCol?.length ? 1 : 0);
  }

  return schema;
}

// All active custom surveys, converted to analysis-ready schemas, keyed by
// id — surveys with no Likert questions at all are skipped (nothing for
// analyze() to compute).
export function getActiveCustomAnalysisSchemas() {
  return Object.fromEntries(
    loadCustomSurveys()
      .filter(s => s.status === "active" && (s.sections ?? []).some(sec => (sec.questions ?? []).some(q => q.type === "likert")))
      .map(s => [s.id, toAnalysisSchema(s)])
  );
}

// The 5 built-in YAML schemas plus active custom surveys. Safe to use
// everywhere the app currently uses the static SCHEMAS import — it's a
// superset, so behavior for the 5 built-in ids is unchanged.
export function getAllAnalysisSchemas() {
  const textOverrides = loadBuiltInReportTexts();
  const builtIns = Object.fromEntries(Object.entries(BUILTIN_SCHEMAS).map(([id, schema]) => [id, {
    ...schema,
    reportTexts: createReportTexts(textOverrides[id]),
  }]));
  return { ...builtIns, ...getActiveCustomAnalysisSchemas() };
}

export function loadBuiltInReportTexts() {
  try { return JSON.parse(localStorage.getItem(BUILTIN_REPORT_TEXTS_KEY) || "{}"); }
  catch { return {}; }
}

export function saveBuiltInReportTexts(id, reportTexts) {
  const all = loadBuiltInReportTexts();
  all[id] = createReportTexts(reportTexts);
  try { localStorage.setItem(BUILTIN_REPORT_TEXTS_KEY, JSON.stringify(all)); } catch { /* ignore */ }
  return all[id];
}

export function getBuiltInSurveyCatalog() {
  const textOverrides = loadBuiltInReportTexts();
  return Object.values(BUILTIN_SCHEMAS).map(schema => ({
    id: schema.id,
    name: schema.label,
    description: schema.desc || schema.description || "استبيان مدمج وجاهز للتحليل وإصدار التقارير.",
    status: "active",
    scaleType: schema.scale?.type,
    sections: schema.axes ?? [],
    version: schema.version || "مدمج",
    updatedAt: null,
    builtIn: true,
    isFlat: schema.isFlat === true,
    reportTexts: createReportTexts(textOverrides[schema.id]),
    programs: schema.programs || [],
  }));
}

export function createCustomSurveyFromBuiltIn(id) {
  const schema = BUILTIN_SCHEMAS[id];
  if (!schema) return null;
  return createSurvey({
    name: `${schema.label} — نسخة مخصصة`,
    description: "نسخة قابلة للتعديل مبنية على الاستبيان المدمج.",
    surveyType: schema.label,
    scaleType: schema.scale?.type || "likert-5",
    metadata: createMetadata(schema.metadata),
    sections: (schema.axes ?? []).map(axis => createSection({
      name: axis.name,
      description: axis.description || "",
      questions: (axis.questions ?? []).map(question => createQuestion({
        text: question.text,
        excelColumn: question.text,
        type: "likert",
        required: true,
      })),
    })),
  });
}

// Same hint-match-then-fuzzy-similarity algorithm as analyze.js's
// detectSurveyType(), duplicated here (rather than imported) so this module
// has no dependency on analyze.js, parameterized over the merged schema set
// instead of analyze.js's closed-over static SCHEMAS.
function similarity(a, b) {
  const na = normalizeHeader(a), nb = normalizeHeader(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  const longer = na.length > nb.length ? na : nb;
  const shorter = na.length <= nb.length ? na : nb;
  let matches = 0;
  for (const ch of shorter) if (longer.includes(ch)) matches++;
  return (matches / longer.length) * 0.7 + (na.includes(nb) || nb.includes(na) ? 0.3 : 0);
}

export function detectAnySurveyType(filename, headers) {
  const schemas = getAllAnalysisSchemas();
  const name = normalizeHeader(filename);
  for (const schema of Object.values(schemas)) {
    if ((schema.fileHints ?? []).some(h => name.includes(normalizeHeader(h)))) return schema.id;
  }
  const scores = Object.values(schemas).map(schema => {
    const allQTexts = schema.axes.flatMap(ax => ax.questions.map(q => q.text));
    const score = (headers ?? []).reduce((sum, h) => {
      const best = allQTexts.length ? Math.max(...allQTexts.map(t => similarity(h, t))) : 0;
      return sum + best;
    }, 0);
    return { id: schema.id, score };
  });
  scores.sort((a, b) => b.score - a.score);
  return scores[0] && scores[0].score > 5 ? scores[0].id : null;
}

// ── Future-integration preview helper ───────────────────────────────────────
// Mirrors the shape produced by src/schemas/compileSchema.js, WITHOUT touching
// that file or analyze.js. Used today only by the Survey Management preview
// tab.
export function toRuntimeSchemaShape(survey) {
  const scale = SCALE_TYPES[survey.scaleType] ?? SCALE_TYPES["likert-5"];
  let globalSeq = 1;
  return {
    id: survey.id,
    label: survey.name || "(بدون اسم)",
    version: survey.version,
    status: survey.status,
    scale: {
      type: survey.scaleType,
      values: scale.values,
    },
    metadata: survey.metadata ?? createMetadata(),
    axes: (survey.sections ?? []).map((sec, ai) => ({
      id: sec.id,
      name: sec.name || `المحور ${ai + 1}`,
      questions: (sec.questions ?? []).map(q => ({
        id: q.id,
        seq: globalSeq++,
        text: q.text,
        excelColumn: q.excelColumn,
        type: q.type,
        required: q.required,
        category: q.category,
        weight: q.weight,
      })),
    })),
  };
}
