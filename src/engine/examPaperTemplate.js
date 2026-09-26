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

export function examPaperCourseKey(course) {
  return `${course.department}::${course.code}::${course.name}`;
}

export function buildExamPaperCoursesTemplate() {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet([["اسم المقرر", "كود المقرر", "القسم"]]);
  worksheet["!cols"] = [{ wch: 38 }, { wch: 20 }, { wch: 28 }];
  worksheet["!autofilter"] = { ref: "A1:C1" };
  XLSX.utils.book_append_sheet(workbook, worksheet, "المقررات");
  const departmentsSheet = XLSX.utils.aoa_to_sheet([
    ["الأقسام المعتمدة"],
    ...EXAM_PAPER_DEPARTMENTS.map(department => [department]),
  ]);
  departmentsSheet["!cols"] = [{ wch: 32 }];
  XLSX.utils.book_append_sheet(workbook, departmentsSheet, "الأقسام");
  workbook.Props = {
    Title: "قالب مقررات تقييم الورقة الامتحانية",
    Subject: "اسم المقرر وكود المقرر والقسم",
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
  const nameIndex = findColumn(headers, HEADER_ALIASES.name);
  const codeIndex = findColumn(headers, HEADER_ALIASES.code);
  const departmentIndex = findColumn(headers, HEADER_ALIASES.department);
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
    };
    if (!course.name && !course.code && !course.department) return;
    if (!course.name || !course.code || !course.department) incompleteRows.push(index + 2);
    else parsed.push(course);
  });
  if (incompleteRows.length) throw new Error(`توجد بيانات ناقصة في الصفوف: ${incompleteRows.slice(0, 8).join("، ")}.`);
  if (!parsed.length) throw new Error("لم يتم العثور على مقررات داخل الملف.");
  return [...new Map(parsed.map(course => [examPaperCourseKey(course), course])).values()];
}
