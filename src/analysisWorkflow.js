export const ANALYSIS_STAGES = Object.freeze({
  SELECTING: "selecting",
  REVIEWING: "reviewing",
  CLEANING: "cleaning",
  PROCESSING: "processing",
  RESULTS: "results",
});

export const ANALYSIS_SOURCES = Object.freeze({
  LOCAL: "local",
  DRIVE: "drive",
  FORMS: "forms",
  SUMMARY: "summary",
  SEMESTER: "semester",
});

export const ANALYSIS_MODES = Object.freeze({
  SINGLE: "single",
  BATCH: "batch",
  COMPARISON: "comparison",
});

export function createAnalysisRequest(overrides = {}) {
  const items = Array.isArray(overrides.items) ? overrides.items.filter(Boolean) : [];
  const inferredMode = overrides.mode || (items.length > 1 ? ANALYSIS_MODES.BATCH : ANALYSIS_MODES.SINGLE);
  return {
    id: overrides.id || `analysis-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    source: overrides.source || ANALYSIS_SOURCES.LOCAL,
    mode: inferredMode,
    items,
    detected: {
      type: "",
      year: "",
      program: "",
      semester: "",
      ...(overrides.detected || {}),
    },
    output: {
      word: inferredMode !== ANALYSIS_MODES.BATCH,
      pdf: true,
      destination: "download",
      ...(overrides.output || {}),
    },
    returnView: overrides.returnView || "analytics",
    context: overrides.context || {},
  };
}

export function analysisSourceLabel(source) {
  return {
    [ANALYSIS_SOURCES.LOCAL]: "الجهاز",
    [ANALYSIS_SOURCES.DRIVE]: "Google Drive",
    [ANALYSIS_SOURCES.FORMS]: "Google Forms",
    [ANALYSIS_SOURCES.SUMMARY]: "ملخص الاستبيانات",
    [ANALYSIS_SOURCES.SEMESTER]: "استبيانات الفصل",
  }[source] || "مصدر البيانات";
}

export function analysisModeLabel(mode, count = 0) {
  if (mode === ANALYSIS_MODES.COMPARISON) return "مقارنة سنوات";
  if (mode === ANALYSIS_MODES.BATCH || count > 1) return `تحليل جماعي (${count})`;
  return "تحليل استبيان واحد";
}

export function validateAnalysisRequest(request, meta = {}) {
  const issues = [];
  if (!request?.items?.length) issues.push("اختر استبيانًا واحدًا على الأقل.");
  if (!String(meta.preparedBy || "").trim()) issues.push("أدخل اسم مُعدّ التقرير.");
  if (!String(meta.reviewer || "").trim()) issues.push("أدخل اسم مراجع التقرير.");
  if (request?.mode === ANALYSIS_MODES.COMPARISON) {
    const years = request.items.map(item => String(item.year || "").trim()).filter(Boolean);
    if (years.length !== request.items.length) issues.push("حدد السنة الدراسية لكل ملف في المقارنة.");
    if (new Set(years).size !== years.length) issues.push("يجب أن تكون سنوات المقارنة مختلفة.");
    const types = request.items.map(item => item.type).filter(Boolean);
    if (types.length === request.items.length && new Set(types).size > 1) issues.push("يجب أن تكون ملفات المقارنة من نوع الاستبيان نفسه.");
  }
  return issues;
}
