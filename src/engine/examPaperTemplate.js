import * as XLSX from "xlsx";

export const EXAM_PAPER_DEPARTMENTS = [
  "Business Adminstration",
  "Business Technology",
  "Accounting",
  "Economics",
  "Political Science",
  "General",
  "Languages",
];

const HEADER_ALIASES = {
  name: ["اسم المقرر", "المقرر", "course name", "course"],
  code: ["كود المقرر", "رمز المقرر", "course code", "code"],
  department: ["القسم", "اسم القسم", "department"],
  exam: ["الامتحان", "ورقة الامتحان", "exam"],
  blueprint: ["blueprint", "مخطط الاختبار", "جدول المواصفات"],
  evaluation: ["التقييم", "evaluation"],
  signature: ["التوقيع", "signature"],
  learningOutcomes: ["تقييم مخرجات التعلم", "مخرجات التعلم", "learning outcomes", "learning outcome"],
  status: ["مستوفي", "مستوفى", "حالة الاستيفاء", "status", "compliant"],
  missingItems: ["العناصر غير المستوفاة", "عناصر غير مستوفاة", "missing items", "missing requirements"],
};

const DEPARTMENT_ALIASES = new Map([
  ["business administration", "Business Adminstration"],
  ["business adminstration", "Business Adminstration"],
  ["ادارة الاعمال", "Business Adminstration"],
  ["إدارة الأعمال", "Business Adminstration"],
  ["business technology", "Business Technology"],
  ["تكنولوجيا الاعمال", "Business Technology"],
  ["تكنولوجيا الأعمال", "Business Technology"],
  ["accounting", "Accounting"],
  ["accoounting", "Accounting"],
  ["المحاسبة", "Accounting"],
  ["محاسبة", "Accounting"],
  ["economics", "Economics"],
  ["الاقتصاد", "Economics"],
  ["اقتصاد", "Economics"],
  ["political science", "Political Science"],
  ["العلوم السياسية", "Political Science"],
  ["علوم سياسية", "Political Science"],
  ["general", "General"],
  ["عام", "General"],
  ["languages", "Languages"],
  ["اللغات", "Languages"],
  ["لغات", "Languages"],
]);

function normalizeText(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function normalizedKey(value) {
  return normalizeText(value).toLowerCase().replace(/[أإآ]/g, "ا");
}

function canonicalDepartment(value) {
  const normalized = normalizedKey(value);
  return DEPARTMENT_ALIASES.get(normalized)
    || EXAM_PAPER_DEPARTMENTS.find(item => item.toLowerCase() === normalized)
    || normalizeText(value);
}

function findColumn(headers, aliases) {
  const normalizedHeaders = headers.map(normalizedKey);
  return normalizedHeaders.findIndex(header => aliases.some(alias => normalizedKey(alias) === header));
}

function parseYesNo(value) {
  const normalized = normalizedKey(value).replace(/[.،]/g, "");
  if (["yes", "y", "true", "1", "نعم", "مستوفي", "مستوفى", "complete", "completed", "done"].includes(normalized)) return true;
  if (["no", "n", "false", "0", "لا", "غير مستوفي", "غير مستوفى", "incomplete", "not complete"].includes(normalized)) return false;
  return null;
}

function parsePercentage(value) {
  const normalized = normalizeText(value).replace("%", "").replace(",", ".");
  if (!normalized) return "";
  const number = Number(normalized);
  if (!Number.isFinite(number)) return normalized;
  return number >= 0 && number <= 1 && normalized.includes(".") ? number * 100 : number;
}

function headerIndexMap(headers) {
  return Object.fromEntries(Object.entries(HEADER_ALIASES).map(([key, aliases]) => [key, findColumn(headers, aliases)]));
}

export function examPaperCourseKey(course) {
  return `${course.department}::${course.code}::${course.name}`;
}

export function buildExamPaperCoursesTemplate() {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet([[
    "كود المقرر",
    "اسم المقرر",
    "القسم",
    "الامتحان",
    "Blueprint",
    "التقييم",
    "التوقيع",
    "تقييم مخرجات التعلم",
    "مستوفي",
    "العناصر غير المستوفاة",
  ]]);
  worksheet["!cols"] = [
    { wch: 20 }, { wch: 38 }, { wch: 26 }, { wch: 14 }, { wch: 16 },
    { wch: 16 }, { wch: 16 }, { wch: 24 }, { wch: 14 }, { wch: 42 },
  ];
  worksheet["!autofilter"] = { ref: "A1:J1" };
  XLSX.utils.book_append_sheet(workbook, worksheet, "Courses");
  workbook.Props = {
    Title: "قالب تقييم الورقة الامتحانية",
    Subject: "بيانات المقررات ونتائج تقييم الورقة الامتحانية",
    Author: "لجنة القياس والتقويم",
  };
  return new Blob([XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true })], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export async function readExamPaperCoursesFile(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });
  if (!rows.length) throw new Error("الملف فارغ.");
  const headers = rows[0].map(normalizeText);
  const columns = headerIndexMap(headers);
  const { name: nameIndex, code: codeIndex, department: departmentIndex } = columns;
  if ([nameIndex, codeIndex, departmentIndex].some(index => index < 0)) {
    throw new Error("يجب أن يحتوي الملف على الأعمدة: اسم المقرر، كود المقرر، القسم.");
  }
  const parsed = [];
  const incompleteRows = [];
  rows.slice(1).forEach((row, index) => {
    const course = {
      name: normalizeText(row[nameIndex]),
      code: normalizeText(row[codeIndex]),
      department: canonicalDepartment(row[departmentIndex]),
      exam: columns.exam >= 0 ? normalizeText(row[columns.exam]) : "",
      blueprint: columns.blueprint >= 0 ? normalizeText(row[columns.blueprint]) : "",
      evaluation: columns.evaluation >= 0 ? normalizeText(row[columns.evaluation]) : "",
      signature: columns.signature >= 0 ? normalizeText(row[columns.signature]) : "",
      learningOutcomes: columns.learningOutcomes >= 0 ? parsePercentage(row[columns.learningOutcomes]) : "",
      missingItems: columns.missingItems >= 0 ? normalizeText(row[columns.missingItems]) : "",
    };
    const statusValue = columns.status >= 0 ? parseYesNo(row[columns.status]) : null;
    course.status = statusValue === true ? "complete" : statusValue === false ? "incomplete" : "";
    if (!course.name && !course.code && !course.department) return;
    if (!course.name || !course.code || !course.department) incompleteRows.push(index + 2);
    else parsed.push(course);
  });
  if (incompleteRows.length) throw new Error(`توجد بيانات ناقصة في الصفوف: ${incompleteRows.slice(0, 8).join("، ")}.`);
  if (!parsed.length) throw new Error("لم يتم العثور على مقررات داخل الملف.");
  return [...new Map(parsed.map(course => [examPaperCourseKey(course), course])).values()];
}

export function examPaperCourseIssues(course) {
  const issues = [];
  if (!course.status) issues.push("حالة الاستيفاء");
  const outcome = Number(course.learningOutcomes);
  if (course.learningOutcomes !== "" && (!Number.isFinite(outcome) || outcome < 0 || outcome > 100)) issues.push("تقييم مخرجات التعلم");
  if (course.status === "incomplete" && !normalizeText(course.missingItems)) issues.push("العناصر غير المستوفاة");
  return issues;
}
