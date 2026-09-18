const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak, Footer, Header,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

const data = JSON.parse(fs.readFileSync('/home/claude/analysis_result.json', 'utf8'));

// ===================================================================
// Helpers
// ===================================================================
const border = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const darkBorder = { style: BorderStyle.SINGLE, size: 6, color: "1F3864" };
const borders = { top: border, bottom: border, left: border, right: border };
const darkBorders = { top: darkBorder, bottom: darkBorder, left: darkBorder, right: darkBorder };

const ARABIC_FONT = "Arial";
const PAGE_W = 11906; // A4
const MARGINS = { top: 1000, bottom: 1000, left: 1200, right: 1200 };
const CONTENT_W = PAGE_W - MARGINS.left - MARGINS.right; // 9506 DXA

const COLORS = {
  primary: "1F3864",
  secondary: "2E75B6",
  light: "D6E4F0",
  lighter: "EBF3FB",
  header: "1F3864",
  axis: "2E75B6",
  q_alt: "F5F9FD",
  white: "FFFFFF",
  accent: "C00000",
  green: "375623",
  greenLight: "E2EFDA",
  yellow: "FFF2CC",
  orange: "FCE4D6",
  red: "FCE4D6",
};

const rtlPara = (children, opts = {}) => new Paragraph({
  bidirectional: true,
  alignment: opts.alignment || AlignmentType.RIGHT,
  ...opts,
  children,
});

const cell = (children, opts = {}) => new TableCell({
  borders: opts.borders || borders,
  shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
  width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
  margins: { top: 80, bottom: 80, left: 120, right: 120 },
  verticalAlign: VerticalAlign.CENTER,
  children: Array.isArray(children) ? children : [children],
});

const txtCell = (text, opts = {}) => cell(
  [rtlPara([new TextRun({
    text,
    font: ARABIC_FONT,
    size: opts.size || 18,
    bold: opts.bold || false,
    color: opts.color || "000000",
  })], { alignment: opts.align || AlignmentType.CENTER })],
  opts
);

const heading = (text, level = 1) => rtlPara([
  new TextRun({ text, font: ARABIC_FONT, bold: true, size: level === 1 ? 36 : level === 2 ? 28 : 24, color: COLORS.primary })
], {
  alignment: AlignmentType.CENTER,
  spacing: { before: level === 1 ? 400 : 240, after: level === 1 ? 300 : 180 },
  border: level === 1 ? { bottom: { style: BorderStyle.SINGLE, size: 8, color: COLORS.secondary, space: 4 } } : undefined,
});

const sectionTitle = (text) => rtlPara([
  new TextRun({ text, font: ARABIC_FONT, bold: true, size: 26, color: COLORS.white })
], {
  alignment: AlignmentType.CENTER,
  spacing: { before: 200, after: 120 },
  shading: { fill: COLORS.primary, type: ShadingType.CLEAR },
  indent: { left: 0, right: 0 },
});

const para = (text, opts = {}) => rtlPara([
  new TextRun({ text, font: ARABIC_FONT, size: opts.size || 20, bold: opts.bold || false, color: opts.color || "000000" })
], {
  alignment: opts.align || AlignmentType.RIGHT,
  spacing: { before: opts.before || 80, after: opts.after || 80 },
  indent: { right: opts.indent || 0 },
});

const bullet = (text) => rtlPara([
  new TextRun({ text: `● ${text}`, font: ARABIC_FONT, size: 20 })
], { alignment: AlignmentType.RIGHT, spacing: { before: 60, after: 60 }, indent: { right: 360 } });

// ===================================================================
// Build questionnaire variables section
// ===================================================================
function buildVariablesSection() {
  const items = [
    rtlPara([], { spacing: { before: 0, after: 100 } }),
    sectionTitle("أولاً: متغيرات الاستبيان"),
    rtlPara([], { spacing: { before: 0, after: 80 } }),
    para(`يتكون الاستبيان من المحاور والعناصر التالية:`, { bold: true, size: 22 }),
    rtlPara([], { spacing: { before: 0, after: 60 } }),
  ];

  // Table: محور, عدد الأسئلة
  const colW = [1400, 5900, 2200];
  const tableRows = [
    new TableRow({
      tableHeader: true,
      children: [
        txtCell("عدد الأسئلة", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: colW[2], size: 20 }),
        txtCell("اسم المحور", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: colW[1], size: 20 }),
        txtCell("م", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: colW[0], size: 20 }),
      ]
    }),
    ...data.axes.map((ax, i) => new TableRow({
      children: [
        txtCell(String(ax.questions.length), { fill: i % 2 === 0 ? COLORS.lighter : COLORS.white, width: colW[2] }),
        cell([rtlPara([new TextRun({ text: ax.name, font: ARABIC_FONT, size: 18 })], { alignment: AlignmentType.RIGHT })],
          { fill: i % 2 === 0 ? COLORS.lighter : COLORS.white, width: colW[1] }),
        txtCell(String(i + 1), { fill: i % 2 === 0 ? COLORS.lighter : COLORS.white, width: colW[0] }),
      ]
    })),
    new TableRow({
      children: [
        txtCell(String(data.n_questions), { fill: COLORS.light, bold: true, width: colW[2] }),
        txtCell("الإجمالي", { fill: COLORS.light, bold: true, width: colW[1] }),
        txtCell("", { fill: COLORS.light, width: colW[0] }),
      ]
    }),
  ];

  items.push(new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: colW,
    rows: tableRows,
  }));

  items.push(rtlPara([], { spacing: { before: 0, after: 120 } }));
  items.push(para(`إجمالي عدد الأسئلة: ${data.n_questions} سؤالاً موزعة على ${data.n_axes} محوراً رئيسياً`, { bold: true, size: 21 }));
  items.push(para(`عدد المستجيبين: ${data.n_responses} طالباً وطالبة`, { bold: true, size: 21 }));

  return items;
}

// ===================================================================
// Statistical method section
// ===================================================================
function buildMethodSection() {
  return [
    rtlPara([], { spacing: { before: 0, after: 100 } }),
    sectionTitle("ثانياً: المعالجة الإحصائية المستخدمة"),
    rtlPara([], { spacing: { before: 0, after: 80 } }),
    para("اعتمد الباحثون في تحليل بيانات الاستبيان على المعالجة الإحصائية التالية:", { bold: true, size: 21 }),
    rtlPara([], { spacing: { before: 0, after: 60 } }),
    bullet("التكرارات والنسب المئوية: لحساب توزيع استجابات الطلاب على كل عبارة وفق مقياس ليكرت الخماسي."),
    bullet("المتوسط الحسابي: للحكم على مستوى الرضا في كل عبارة ومحور ولتحديد الاتجاه العام للاستجابات."),
    rtlPara([], { spacing: { before: 40, after: 40 } }),
    para("مقياس الحكم على المتوسط الحسابي (مقياس ليكرت الخماسي):", { bold: true, size: 20 }),
    rtlPara([], { spacing: { before: 0, after: 60 } }),
    (() => {
      const colW2 = [1900, 1900, 1900, 1900, 1901];
      return new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: colW2,
        rows: [
          new TableRow({ tableHeader: true, children: [
            txtCell("أوافق بشدة", { fill: COLORS.greenLight, bold: true, width: colW2[4], size: 19 }),
            txtCell("أوافق", { fill: COLORS.lighter, bold: true, width: colW2[3], size: 19 }),
            txtCell("محايد", { fill: COLORS.yellow, bold: true, width: colW2[2], size: 19 }),
            txtCell("لا أوافق", { fill: COLORS.orange, bold: true, width: colW2[1], size: 19 }),
            txtCell("لا أوافق بشدة", { fill: COLORS.red, bold: true, width: colW2[0], size: 19 }),
          ]}),
          new TableRow({ children: [
            txtCell("4.50 – 5.00", { width: colW2[4], size: 18 }),
            txtCell("3.50 – 4.49", { width: colW2[3], size: 18 }),
            txtCell("2.50 – 3.49", { width: colW2[2], size: 18 }),
            txtCell("1.50 – 2.49", { width: colW2[1], size: 18 }),
            txtCell("1.00 – 1.49", { width: colW2[0], size: 18 }),
          ]}),
          new TableRow({ children: [
            txtCell("5", { fill: COLORS.greenLight, width: colW2[4], size: 18 }),
            txtCell("4", { fill: COLORS.lighter, width: colW2[3], size: 18 }),
            txtCell("3", { fill: COLORS.yellow, width: colW2[2], size: 18 }),
            txtCell("2", { fill: COLORS.orange, width: colW2[1], size: 18 }),
            txtCell("1", { fill: COLORS.red, width: colW2[0], size: 18 }),
          ]}),
        ]
      });
    })(),
  ];
}

// ===================================================================
// Results section: per-axis tables
// ===================================================================
function buildResultsSection() {
  const items = [
    rtlPara([], { spacing: { before: 0, after: 100 } }),
    sectionTitle("ثالثاً: عرض النتائج وتحليلها ومناقشتها"),
    rtlPara([], { spacing: { before: 0, after: 80 } }),
    para("يعرض هذا القسم نتائج استجابات الطلاب لكل محور من محاور الاستبيان، موضحاً النسب المئوية للاستجابات والمتوسط الحسابي لكل عبارة واتجاهها العام.", { size: 20 }),
    rtlPara([], { spacing: { before: 0, after: 100 } }),
  ];

  const pctCols = [
    "لا أوافق بشدة", "لا أوافق", "محايد", "أوافق", "أوافق بشدة"
  ];

  data.axes.forEach((ax, axIdx) => {
    // Axis header
    items.push(rtlPara([
      new TextRun({ text: `محور ${ax.number}: ${ax.name}`, font: ARABIC_FONT, bold: true, size: 24, color: COLORS.white })
    ], {
      alignment: AlignmentType.RIGHT,
      spacing: { before: 200, after: 100 },
      shading: { fill: COLORS.secondary, type: ShadingType.CLEAR },
      indent: { left: 0, right: 0 },
    }));

    // Column widths for results table
    // م | العبارة | لا أوافق بشدة | لا أوافق | محايد | أوافق | أوافق بشدة | المتوسط | الاتجاه
    const cw = [700, 3200, 700, 700, 700, 700, 700, 800, 806];

    const headerRow = new TableRow({
      tableHeader: true,
      children: [
        txtCell("الاتجاه", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[8], size: 17 }),
        txtCell("المتوسط", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[7], size: 17 }),
        txtCell("أوافق بشدة (5)", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[6], size: 16 }),
        txtCell("أوافق (4)", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[5], size: 16 }),
        txtCell("محايد (3)", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[4], size: 16 }),
        txtCell("لا أوافق (2)", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[3], size: 16 }),
        txtCell("لا أوافق بشدة (1)", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[2], size: 15 }),
        cell([rtlPara([new TextRun({ text: "العبارة", font: ARABIC_FONT, size: 17, bold: true, color: "FFFFFF" })], { alignment: AlignmentType.RIGHT })],
          { fill: COLORS.primary, borders: darkBorders, width: cw[1] }),
        txtCell("م", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[0], size: 17 }),
      ]
    });

    const dataRows = ax.questions.map((q, qi) => {
      const bg = qi % 2 === 0 ? COLORS.white : COLORS.q_alt;
      return new TableRow({ children: [
        txtCell(q.direction, { fill: bg, width: cw[8], size: 15 }),
        txtCell(String(q.mean), { fill: bg, width: cw[7], size: 16, bold: true }),
        txtCell(`${q.percentages["أوافق بشدة"]}%`, { fill: bg, width: cw[6], size: 16 }),
        txtCell(`${q.percentages["أوافق"]}%`, { fill: bg, width: cw[5], size: 16 }),
        txtCell(`${q.percentages["محايد"]}%`, { fill: bg, width: cw[4], size: 16 }),
        txtCell(`${q.percentages["لا أوافق"]}%`, { fill: bg, width: cw[3], size: 16 }),
        txtCell(`${q.percentages["لا أوافق بشدة"]}%`, { fill: bg, width: cw[2], size: 16 }),
        cell([rtlPara([new TextRun({ text: q.text.trim(), font: ARABIC_FONT, size: 17 })], { alignment: AlignmentType.RIGHT })],
          { fill: bg, width: cw[1] }),
        txtCell(String(q.number), { fill: bg, width: cw[0], size: 16 }),
      ]});
    });

    // Axis average row
    const avgRow = new TableRow({ children: [
      txtCell(ax.direction, { fill: COLORS.light, bold: true, width: cw[8], size: 16 }),
      txtCell(String(ax.mean), { fill: COLORS.light, bold: true, width: cw[7], size: 17 }),
      txtCell("", { fill: COLORS.light, width: cw[6] }),
      txtCell("", { fill: COLORS.light, width: cw[5] }),
      txtCell("", { fill: COLORS.light, width: cw[4] }),
      txtCell("", { fill: COLORS.light, width: cw[3] }),
      txtCell("", { fill: COLORS.light, width: cw[2] }),
      cell([rtlPara([new TextRun({ text: "متوسط المحور", font: ARABIC_FONT, size: 17, bold: true })], { alignment: AlignmentType.RIGHT })],
        { fill: COLORS.light, width: cw[1] }),
      txtCell("", { fill: COLORS.light, width: cw[0] }),
    ]});

    items.push(new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: cw,
      rows: [headerRow, ...dataRows, avgRow],
    }));

    items.push(rtlPara([], { spacing: { before: 60, after: 60 } }));
  });

  return items;
}

// ===================================================================
// Summary section
// ===================================================================
function buildSummarySection() {
  const items = [
    new Paragraph({ children: [new PageBreak()] }),
    sectionTitle("رابعاً: ملخص النتائج"),
    rtlPara([], { spacing: { before: 0, after: 80 } }),
    para("يوضح الجدول التالي ملخص نتائج جميع محاور الاستبيان مع المتوسط الحسابي والاتجاه العام لكل محور:", { size: 20 }),
    rtlPara([], { spacing: { before: 0, after: 80 } }),
  ];

  const cw = [1200, 4200, 1600, 1700, 806];

  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      txtCell("النسبة المئوية", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[0], size: 17 }),
      txtCell("الاتجاه العام", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[1], size: 17 }),
      txtCell("المتوسط الحسابي", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[2], size: 17 }),
      cell([rtlPara([new TextRun({ text: "المحور", font: ARABIC_FONT, size: 17, bold: true, color: "FFFFFF" })], { alignment: AlignmentType.RIGHT })],
        { fill: COLORS.primary, borders: darkBorders, width: cw[3] }),
      txtCell("م", { fill: COLORS.primary, bold: true, color: "FFFFFF", borders: darkBorders, width: cw[4], size: 17 }),
    ]
  });

  const dataRows = data.axes.map((ax, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.lighter;
    return new TableRow({ children: [
      txtCell(`${ax.agree_pct}%`, { fill: bg, width: cw[0], size: 17 }),
      txtCell(ax.direction, { fill: bg, width: cw[1], size: 17 }),
      txtCell(String(ax.mean), { fill: bg, bold: true, width: cw[2], size: 17 }),
      cell([rtlPara([new TextRun({ text: ax.name, font: ARABIC_FONT, size: 17 })], { alignment: AlignmentType.RIGHT })],
        { fill: bg, width: cw[3] }),
      txtCell(String(i + 1), { fill: bg, width: cw[4], size: 17 }),
    ]});
  });

  // Overall row
  const totalRow = new TableRow({ children: [
    txtCell(`${data.overall_agree_pct}%`, { fill: COLORS.light, bold: true, width: cw[0], size: 18 }),
    txtCell(data.overall_direction, { fill: COLORS.light, bold: true, width: cw[1], size: 18 }),
    txtCell(String(data.overall_mean), { fill: COLORS.light, bold: true, width: cw[2], size: 18 }),
    txtCell("نسبة الاستقصاء الكلية / المتوسط العام", { fill: COLORS.light, bold: true, width: cw[3], size: 17 }),
    txtCell("", { fill: COLORS.light, width: cw[4] }),
  ]});

  items.push(new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: cw,
    rows: [headerRow, ...dataRows, totalRow],
  }));

  items.push(rtlPara([], { spacing: { before: 120, after: 80 } }));
  items.push(para(`نسبة الاستقصاء الكلية (الموافقة): ${data.overall_agree_pct}%`, { bold: true, size: 22 }));
  items.push(para(`المتوسط العام للاستبيان: ${data.overall_mean} من 5`, { bold: true, size: 22 }));
  items.push(para(`الاتجاه العام: ${data.overall_direction}`, { bold: true, size: 22, color: COLORS.secondary }));
  items.push(rtlPara([], { spacing: { before: 0, after: 100 } }));
  items.push(para(`عدد المستجيبين: ${data.n_responses} طالباً وطالبة`, { size: 20 }));
  items.push(para(`عدد الأسئلة: ${data.n_questions} سؤالاً موزعة على ${data.n_axes} محاور`, { size: 20 }));

  return items;
}

// ===================================================================
// Title page
// ===================================================================
function buildTitlePage() {
  return [
    rtlPara([], { spacing: { before: 600, after: 100 } }),
    rtlPara([
      new TextRun({ text: "نتائج وتحليل", font: ARABIC_FONT, size: 48, bold: true, color: COLORS.primary }),
    ], { alignment: AlignmentType.CENTER }),
    rtlPara([
      new TextRun({ text: "استبيان رضا الطلاب", font: ARABIC_FONT, size: 52, bold: true, color: COLORS.secondary }),
    ], { alignment: AlignmentType.CENTER }),
    rtlPara([], { spacing: { before: 60, after: 60 } }),
    rtlPara([
      new TextRun({ text: "ACC 2024 - 2025", font: "Arial", size: 36, bold: true, color: COLORS.primary }),
    ], { alignment: AlignmentType.CENTER }),
    rtlPara([], { spacing: { before: 100, after: 100 } }),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
      rows: [
        new TableRow({ children: [
          txtCell(`عدد المحاور: ${data.n_axes}`, { fill: COLORS.lighter, bold: true, width: CONTENT_W / 2, size: 22 }),
          txtCell(`عدد الأسئلة: ${data.n_questions}`, { fill: COLORS.lighter, bold: true, width: CONTENT_W / 2, size: 22 }),
        ]}),
        new TableRow({ children: [
          txtCell(`المتوسط العام: ${data.overall_mean} / 5`, { fill: COLORS.light, bold: true, width: CONTENT_W / 2, size: 22 }),
          txtCell(`عدد المستجيبين: ${data.n_responses}`, { fill: COLORS.light, bold: true, width: CONTENT_W / 2, size: 22 }),
        ]}),
        new TableRow({ children: [
          txtCell(`الاتجاه العام: ${data.overall_direction}`, { fill: COLORS.lighter, bold: true, width: CONTENT_W / 2, size: 22 }),
          txtCell(`نسبة الموافقة: ${data.overall_agree_pct}%`, { fill: COLORS.lighter, bold: true, width: CONTENT_W / 2, size: 22 }),
        ]}),
      ]
    }),
    rtlPara([], { spacing: { before: 100, after: 100 } }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ===================================================================
// Assemble document
// ===================================================================
const allChildren = [
  ...buildTitlePage(),
  ...buildVariablesSection(),
  ...buildMethodSection(),
  ...buildResultsSection(),
  ...buildSummarySection(),
];

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: ARABIC_FONT, size: 20 } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: 16838 },
        margin: MARGINS,
      },
    },
    children: allChildren,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/survey_analysis.docx', buf);
  console.log('Done!');
});
