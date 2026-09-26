import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  ImageRun,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";

const FONT = "Arial";
const PRIMARY = "1F4E78";
const PALE = "F4F8FC";
const BORDER = "B8C6D1";
const PAGE_WIDTH = 11906;
const MARGINS = { top: 1040, bottom: 1040, left: 850, right: 850 };
const CONTENT_WIDTH = PAGE_WIDTH - MARGINS.left - MARGINS.right;
const singleBorder = { style: BorderStyle.SINGLE, size: 4, color: BORDER };
const borders = { top: singleBorder, bottom: singleBorder, left: singleBorder, right: singleBorder };

const textRun = (text, options = {}) => new TextRun({
  text: String(text ?? ""),
  font: FONT,
  size: options.size ?? 20,
  bold: options.bold ?? false,
  color: options.color ?? "000000",
  rightToLeft: options.rightToLeft ?? true,
});

const paragraph = (text, options = {}) => new Paragraph({
  bidirectional: options.bidirectional ?? true,
  alignment: options.alignment ?? AlignmentType.RIGHT,
  spacing: { before: options.before ?? 0, after: options.after ?? 80, line: options.line ?? 280 },
  keepNext: options.keepNext,
  children: [textRun(text, options)],
});

const cell = (text, options = {}) => new TableCell({
  width: options.width ? { size: options.width, type: WidthType.DXA } : undefined,
  borders,
  shading: options.fill ? { fill: options.fill, type: ShadingType.CLEAR } : undefined,
  margins: { top: 100, bottom: 100, left: 110, right: 110 },
  verticalAlign: VerticalAlign.CENTER,
  children: [paragraph(text, {
    size: options.size ?? 18,
    bold: options.bold,
    color: options.color,
    alignment: options.alignment ?? AlignmentType.CENTER,
    after: 0,
    line: 250,
  })],
});

const headerRow = (labels, widths) => new TableRow({
  tableHeader: true,
  cantSplit: true,
  children: labels.map((label, index) => cell(label, {
    width: widths[index], fill: PRIMARY, color: "FFFFFF", bold: true, size: 17,
  })),
});

const dataRow = (values, widths, index, narrativeColumns = []) => new TableRow({
  cantSplit: true,
  children: values.map((value, columnIndex) => cell(value, {
    width: widths[columnIndex],
    fill: index % 2 ? PALE : "FFFFFF",
    alignment: narrativeColumns.includes(columnIndex) ? AlignmentType.RIGHT : AlignmentType.CENTER,
    size: 17,
  })),
});

const reportTable = (headers, widths, rows, narrativeColumns = []) => new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  visuallyRightToLeft: true,
  rows: [
    headerRow(headers, widths),
    ...rows.map((row, index) => dataRow(row, widths, index, narrativeColumns)),
  ],
});

const sectionTitle = (text) => new Paragraph({
  bidirectional: true,
  alignment: AlignmentType.RIGHT,
  spacing: { before: 180, after: 90 },
  keepNext: true,
  children: [textRun(text, { size: 23, bold: true, color: "000000" })],
});

const spacer = (after = 70) => new Paragraph({ spacing: { after }, children: [] });

async function loadLogo(settings) {
  const source = settings?.logoDataUrl || "/logo.png";
  try {
    const response = await fetch(source);
    if (!response.ok) return null;
    return response.arrayBuffer();
  } catch {
    return null;
  }
}

function buildHeader(logo, settings) {
  const children = [];
  if (logo) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 20 },
      children: [new ImageRun({ data: logo, type: "png", transformation: { width: 58, height: 58 } })],
    }));
  }
  children.push(
    paragraph(settings?.facultyNameArabic || "كلية الإدارة والاقتصاد وتكنولوجيا الأعمال", { alignment: AlignmentType.CENTER, bold: true, size: 19, after: 0 }),
    paragraph("لجنة القياس والتقويم", { alignment: AlignmentType.CENTER, bold: true, size: 18, after: 30 }),
  );
  return new Header({ children });
}

function buildFooter() {
  return new Footer({ children: [
    paragraph("رؤية الكلية: أن تكون مؤسسة تعليم عالٍ متميزة محليًا وإقليميًا، تدعم البحث والتطوير وتشارك بفاعلية في تشكيل مستقبل الأعمال والمجتمع.", { size: 13, color: "555555", after: 15, line: 210 }),
    paragraph("رسالة الكلية: تقديم برامج تعليمية متميزة تواكب التطورات المعرفية والتقنيات الحديثة، وتطوير البحث العلمي والخدمات المجتمعية في إطار القيم الأخلاقية والمسؤولية المهنية.", { size: 13, color: "555555", after: 0, line: 210 }),
  ] });
}

export function summarizeExamPaperEvaluations(courses) {
  const completed = courses.filter(course => course.status === "complete");
  const incomplete = courses.filter(course => course.status === "incomplete");
  const total = courses.length;
  const pct = count => total ? `${(count * 100 / total).toFixed(1).replace(".0", "")}%` : "0%";
  return { completed, incomplete, total, completedPct: pct(completed.length), incompletePct: pct(incomplete.length) };
}

export function examPaperReportFilename({ department, semester, year }) {
  const clean = value => String(value || "").replace(/[\\/:*?"<>|]/g, "-").trim();
  return `تقرير تقييم الورقة الامتحانية من حيث الشكل - ${clean(department)} - ${clean(semester)} ${clean(year)}.docx`;
}

const learningOutcomeLabel = value => value === "" || value === null || value === undefined ? "" : `${value}%`;

export async function buildExamPaperReportDocx(report, settings = {}) {
  const logo = await loadLogo(settings);
  const summary = summarizeExamPaperEvaluations(report.courses);
  const title = `تقرير تقييم الورقة الامتحانية من حيث الشكل لقسم ${report.department} للفصل الدراسي ${report.semester} ${report.year}`;
  const children = [
    paragraph(title, { alignment: AlignmentType.CENTER, bold: true, size: 28, before: 100, after: 210, line: 330 }),
    sectionTitle("أولًا: بيان بعدد ونسبة مقررات القسم التي تم تقييم الورقة الامتحانية لها من حيث الشكل"),
    reportTable(
      ["مقررات القسم", "العدد", "النسبة"],
      [6100, 1900, 1900],
      [
        ["المقررات المستوفاة لجميع العناصر", summary.completed.length, summary.completedPct],
        ["المقررات غير المستوفاة لبعض العناصر", summary.incomplete.length, summary.incompletePct],
        ["الإجمالي", summary.total, "100%"],
      ],
      [0],
    ),
    spacer(),
    sectionTitle("ثانيًا: المقررات المستوفاة لجميع عناصر تقييم الورقة الامتحانية من حيث الشكل"),
    reportTable(
      ["م", "كود المقرر", "المقرر", "تقييم مخرجات التعلم"],
      [650, 1900, 4850, 2500],
      summary.completed.length
        ? summary.completed.map((course, index) => [index + 1, course.code, course.name, learningOutcomeLabel(course.learningOutcomes)])
        : [["—", "—", "لا توجد مقررات مستوفاة", "—"]],
      [2],
    ),
    spacer(),
    sectionTitle("ثالثًا: المقررات غير المستوفاة لبعض عناصر تقييم الورقة الامتحانية من حيث الشكل"),
    reportTable(
      ["م", "كود المقرر", "المقرر", "العناصر غير المستوفاة", "تقييم مخرجات التعلم"],
      [550, 1450, 2850, 3300, 1750],
      summary.incomplete.length
        ? summary.incomplete.map((course, index) => [index + 1, course.code, course.name, course.missingItems, learningOutcomeLabel(course.learningOutcomes)])
        : [["—", "—", "لا توجد مقررات غير مستوفاة", "—", "—"]],
      [2, 3],
    ),
    spacer(100),
    sectionTitle("القائمون بالتقييم"),
    reportTable(
      ["م", "الاسم", "الوظيفة"],
      [700, 4300, 4900],
      report.evaluators.map((evaluator, index) => [index + 1, evaluator.name, evaluator.role]),
      [1, 2],
    ),
    spacer(180),
    paragraph("رئيس لجنة القياس والتقويم", { alignment: AlignmentType.CENTER, bold: true, size: 20, after: 110 }),
    paragraph(report.committeeHead || report.evaluators[0]?.name || "", { alignment: AlignmentType.CENTER, bold: true, size: 20, after: 0 }),
  ];

  const doc = new Document({
    creator: "لجنة القياس والتقويم",
    title,
    description: "تقرير تقييم الورقة الامتحانية من حيث الشكل",
    styles: {
      default: {
        document: { run: { font: FONT, size: 20, color: "000000" }, paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT } },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_WIDTH, height: 16838 },
          margin: MARGINS,
        },
      },
      headers: { default: buildHeader(logo, settings) },
      footers: { default: buildFooter() },
      children,
    }],
  });
  return Packer.toBlob(doc);
}
