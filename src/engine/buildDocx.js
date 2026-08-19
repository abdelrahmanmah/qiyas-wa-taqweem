import {
  Document, Header, Footer, Packer,
  Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageBreak,
} from "docx";

const FONT = "Arial";

// ── Mutable design state (set by applyDesign before each build) ───────────────
let _font      = FONT;
let _isRTL     = true;
let _textAlign = AlignmentType.RIGHT;
let _numAlign  = AlignmentType.CENTER;

const COLORS = {
  primary:   "1F3864",
  secondary: "2E75B6",
  lightBlue: "BDD7EE",
  light:     "D6E4F0",
  lighter:   "EBF3FB",
  white:     "FFFFFF",
  gray:      "F2F2F2",
  darkGray:  "404040",
  red:       "FFCCCC",
  yellow:    "FFF2CC",
};

const COLOR_THEMES = {
  default: COLORS,
  green:  { ...COLORS, primary: "1a4731", secondary: "2d6a4f", lightBlue: "a8d5b5", light: "d4edda", lighter: "eaf6ee" },
  purple: { ...COLORS, primary: "4a235a", secondary: "7d3c98", lightBlue: "d7bde2", light: "e8daef", lighter: "f5eef8" },
  dark:   { ...COLORS, primary: "1c1c1c", secondary: "444444", lightBlue: "c0c0c0", light: "e0e0e0", lighter: "f5f5f5" },
};
let _C = COLORS;

const ACADEMIC = {
  headRowFill: "D9D9D9",
  zebra:       "F5F5F5",
  total:       "BFBFBF",
  border:      "808080",
};

const PAGE_W  = 11906;
const MARGINS = { top: 1134, bottom: 1134, left: 1134, right: 1134 };
const CW      = PAGE_W - MARGINS.left - MARGINS.right;
const bdr     = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
const BORDERS = { top: bdr, bottom: bdr, left: bdr, right: bdr };
const NO_BORDERS = {
  top:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

// ── Default settings (exported so App can import and use as initial state) ────
export const DEFAULT_SETTINGS = {
  uniName:       "Egyptian Russian University",
  facultyName:   "Faculty of Management, Economics and Business Technology",
  unitName:      "Quality Assurance unit (QAU)",
  committeeName: "Measurement and Evaluation Committee",
  qmName:        "د/ رويدا علي",
  quName:        "د/ هند الجميل",
  vision:
    "رؤية الوحدة: أن تكون وحدة ضمان الجودة بكلية الإدارة والاقتصاد وتكنولوجيا الأعمال بالجامعة المصرية الروسية " +
    "وحدة متميزة فى نظام إدارة الجودة الداخلية بما يؤهل الكلية للحصول على الإعتماد المؤسسي لها والأكاديمي " +
    "لبرامجها فى ضوء معايير الجودة بما يعمل على وضع الكلية فى مصاف الكليات الرائدة والمتميزة على المستوى " +
    "الوطنى و الدولي، وبما يلبي احتياجات سوق العمل وكسب ثقة المجتمع في خريجي الكلية",
  mission:
    "رسالة الوحدة: تسعى وحدة ضمان الجودة بكلية الإدارة والاقتصاد وتكنولوجيا الأعمال بالجامعة المصرية الروسية " +
    "فى مصر لنشر وتعميق فكر جودة التعليم الجامعى من خلال تنفيذ وتفعيل أنشطة معايير جودة التعليم العالى " +
    "والمراجعة الداخلية والخارجية الدورية مع العمل على التطوير والتحسين المستمر للأرتقاء بالعملية التعليمية " +
    "والبحثية والخدمة المجتمعية بما يحقق رسالة الكلية ، ويتسق مع رسالة الجامعة لكسب ثقة المجتمع وتحقيق " +
    "التنمية المستدامة ويؤهل الكلية للحصول علي الإعتماد المؤسسي والأكاديمي .",
  email: "E-mail: qa-mebt@eru.edu.eg    |    E-mail: meb-maec@eru.edu.eg",
  includeRecommendations: true,
  includeEvaluatorsTable: true,
  includeParticipants:    true,
  recommendationsCount:   5,
  includePdfCharts:          true,
  recommendationsThreshold:  70,
  logoDataUrl: null,
  reportFont:      "Arial",
  reportDirection: "rtl",
  colorTheme:      "default",
  textAlign:       "right",
  numAlign:        "center",
};

// ── Primitive builders ────────────────────────────────────────────────────────
const rp = (children, opts = {}) =>
  new Paragraph({ bidirectional: _isRTL, alignment: cellAlign(_isRTL ? AlignmentType.RIGHT : AlignmentType.LEFT), ...opts, children });

const mk = (text, opts = {}) =>
  new TextRun({
    text,
    font: _font,
    size: opts.size ?? 20,
    bold: opts.bold ?? false,
    color: opts.color ?? "000000",
    rightToLeft: opts.rightToLeft ?? _isRTL,
  });

const cell = (children, opts = {}) =>
  new TableCell({
    borders: opts.noBorder ? NO_BORDERS : BORDERS,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    width: opts.w ? { size: opts.w, type: WidthType.DXA } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    ...(opts.columnSpan ? { columnSpan: opts.columnSpan } : {}),
    children: Array.isArray(children) ? children : [children],
  });

const tc = (text, opts = {}) =>
  cell(
    [rp([mk(text, { size: opts.size ?? 18, bold: opts.bold, color: opts.color ?? "000000" })],
      { alignment: opts.align ?? _textAlign })],
    opts
  );

const tcNum = (text, opts = {}) =>
  tc(text, { align: _numAlign, ...opts });

const hCell = (text, w, size = 17) =>
  tc(text, { fill: _C.primary, bold: true, color: "FFFFFF", w, size });

const textCell = (text, fill, w, rightAlign = false) =>
  cell(
    [rp([mk(text, { size: 17 })], { alignment: rightAlign ? _textAlign : AlignmentType.CENTER })],
    { fill, w }
  );

const sectionHeading = (text) =>
  new Paragraph({
    bidirectional: _isRTL,
    alignment: cellAlign(_isRTL ? AlignmentType.RIGHT : AlignmentType.LEFT),
    spacing: { before: 160, after: 80 },
    shading: { fill: _C.primary, type: ShadingType.CLEAR },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: _C.secondary } },
    children: [mk(text, { size: 26, bold: true, color: _C.white })],
  });

const axisHeading = (text) =>
  new Paragraph({
    bidirectional: _isRTL,
    alignment: cellAlign(_isRTL ? AlignmentType.RIGHT : AlignmentType.LEFT),
    spacing: { before: 160, after: 80 },
    shading: { fill: _C.secondary, type: ShadingType.CLEAR },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: _C.lightBlue } },
    children: [mk(text, { size: 22, bold: true, color: _C.white })],
  });

const bodyPara = (text, opts = {}) =>
  rp(
    [mk(text, { size: opts.size ?? 20, bold: opts.bold, color: opts.color ?? "000000" })],
    {
      alignment: cellAlign(opts.align ?? (_isRTL ? AlignmentType.RIGHT : AlignmentType.LEFT)),
      spacing: { before: opts.before ?? 80, after: opts.after ?? 80 },
      ...(opts.shading ? { shading: { fill: opts.shading, type: ShadingType.CLEAR } } : {}),
    }
  );

const spacer = () => rp([], { spacing: { before: 60, after: 60 } });

function axisFill(agreePct) {
  if (agreePct >= 85) return _C.lightBlue;
  if (agreePct >= 70) return _C.yellow;
  return _C.red;
}

// ── Logo loader ───────────────────────────────────────────────────────────────
async function getLogoData(settings) {
  if (settings.logoDataUrl) {
    try {
      const resp = await fetch(settings.logoDataUrl);
      return resp.arrayBuffer();
    } catch { /* fall through */ }
  }
  try {
    const resp = await fetch("/logo.png");
    if (!resp.ok) return null;
    return resp.arrayBuffer();
  } catch {
    return null;
  }
}

// ── Page header ───────────────────────────────────────────────────────────────
function buildHeader(logoData, s) {
  const lines = [
    { text: s.uniName,       bold: true,  size: 19 },
    { text: s.facultyName,   bold: false, size: 17 },
    { text: s.unitName,      bold: false, size: 17 },
    { text: s.committeeName, bold: false, size: 17 },
  ];
  const children = [];

  if (logoData) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 20 },
        children: [new ImageRun({ data: logoData, transformation: { width: 60, height: 60 }, type: "png" })],
      })
    );
  }

  for (const line of lines) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: line.text, font: _font, size: line.size, bold: line.bold, color: _C.primary })],
      })
    );
  }

  children.push(
    new Paragraph({
      spacing: { before: 40, after: 0 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: _C.primary } },
      children: [],
    })
  );

  return new Header({ children });
}

// ── Page footer ───────────────────────────────────────────────────────────────
function buildFooter(s) {
  return new Footer({
    children: [
      new Paragraph({
        spacing: { before: 0, after: 0 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: _C.lightBlue } },
        children: [],
      }),
      new Paragraph({
        bidirectional: true,
        alignment: cellAlign(AlignmentType.RIGHT),
        spacing: { before: 20, after: 0 },
        children: [new TextRun({ text: s.vision, font: _font, size: 13, color: _C.darkGray, rightToLeft: true })],
      }),
      new Paragraph({
        bidirectional: true,
        alignment: cellAlign(AlignmentType.RIGHT),
        spacing: { before: 20, after: 20 },
        children: [new TextRun({ text: s.mission, font: _font, size: 13, color: _C.darkGray, rightToLeft: true })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: s.email, font: _font, size: 14, color: _C.primary })],
      }),
    ],
  });
}

// ── Cover page ────────────────────────────────────────────────────────────────
function buildCoverPage(result, meta) {
  const { n, schemaLabel } = result;
  const nodes = [
    spacer(),
    bodyPara("نتائج تحليل استبيان",
      { bold: true, size: 36, color: _C.primary, align: AlignmentType.CENTER, before: 400, after: 60 }),
    bodyPara(`قياس آراء ورضا ${schemaLabel}`,
      { bold: true, size: 28, color: _C.darkGray, align: AlignmentType.CENTER, before: 40, after: 40 }),
    bodyPara(`للعام الدراسي ${meta.year}`,
      { bold: true, size: 22, color: _C.secondary, align: AlignmentType.CENTER, before: 40, after: 60 }),
    spacer(), spacer(),
    bodyPara(`عدد المشاركين بعد حذف التكرارات: ${n}`, { size: 22, bold: true, align: AlignmentType.CENTER }),
  ];

  if (meta.program) {
    nodes.push(bodyPara(`البرنامج / القسم: ${meta.program}`, { bold: true, size: 22, color: _C.secondary, align: AlignmentType.CENTER }));
  }

  nodes.push(
    bodyPara("موجه إلى: مدير وحدة الجودة بالكلية", { bold: true, size: 22, color: _C.primary, align: AlignmentType.CENTER }),
    spacer(),
    new Paragraph({ children: [new PageBreak()] }),
  );
  return nodes;
}

// ── Evaluators table ──────────────────────────────────────────────────────────
function buildEvaluatorsSection(meta, s) {
  const cw = [Math.floor(CW * 0.33), Math.floor(CW * 0.33), CW - 2 * Math.floor(CW * 0.33)];
  return [
    bodyPara("موجه إلى: مدير وحدة الجودة بالكلية", { bold: true, size: 22, color: _C.primary, align: AlignmentType.CENTER }),
    spacer(),
    sectionHeading("  جدول القائم بالتقييم  "),
    spacer(),
    new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: cw,
      rows: [
        new TableRow({ tableHeader: true, children: [
          hCell("التوقيع", cw[0]), hCell("الوظيفة", cw[1]), hCell("الاسم", cw[2]),
        ]}),
        ...[
          { name: meta.preparedBy || "", role: "أعد التقرير" },
          { name: meta.reviewer   || "", role: "راجع التقرير" },
          { name: "",                    role: "" },
          { name: "",                    role: "" },
        ].map((row, i) => {
          const bg = i % 2 === 0 ? _C.lighter : _C.white;
          return new TableRow({ children: [
            tc("",             { fill: bg, w: cw[0] }),
            textCell(row.role, bg, cw[1]),
            textCell(row.name, bg, cw[2]),
          ]});
        }),
      ],
    }),
    spacer(),
  ];
}

// ── Survey variables (أولاً) ──────────────────────────────────────────────────
function buildVariablesSection(result) {
  return [
    sectionHeading("  أولاً: متغيرات الاستبيان  "),
    spacer(),
    bodyPara(
      `اشتمل الاستبيان على (${result.totalQuestions}) عبارة تتمثل في (${result.axes.length}) محور رئيسي.`,
      { size: 22 }
    ),
    spacer(),
  ];
}

// ── Statistical methodology (ثانياً) ──────────────────────────────────────────
function buildMethodologySection(result) {
  const text = result.scaleType === "likert-5" ? "النسب والمتوسط الحسابي." : "النسب.";
  return [
    sectionHeading("  ثانياً: المعالجة الإحصائية المستخدمة  "),
    spacer(),
    bodyPara(text, { size: 22, bold: true }),
    spacer(),
  ];
}

// ── Participants ──────────────────────────────────────────────────────────────
function buildCrosstabTable(cross, rowHeaderLabel = "المسمى الوظيفي / القسم") {
  const { rows, cols, matrix, rowTotals, colTotals, grandTotal } = cross;
  const firstColW = Math.floor(CW * 0.22);
  const totalColW = Math.floor(CW * 0.10);
  const dataColW = Math.floor((CW - firstColW - totalColW) / cols.length);
  const colWidths = [firstColW, ...cols.map(() => dataColW), totalColW];

  const headerRow = new TableRow({ tableHeader: true, children: [
    hCell("الإجمالي", totalColW),
    ...cols.slice().reverse().map(c => hCell(c, dataColW, 15)),
    cell([rp([mk(rowHeaderLabel, { size: 17, bold: true, color: "FFFFFF" })],
      { alignment: _textAlign })], { fill: _C.primary, w: firstColW }),
  ]});

  const dataRows = rows.map((r, i) => {
    const bg = i % 2 === 0 ? _C.white : _C.lighter;
    return new TableRow({ children: [
      tcNum(String(rowTotals[r]), { fill: _C.light, bold: true, w: totalColW, size: 17 }),
      ...cols.slice().reverse().map(c =>
        tcNum(String(matrix[r]?.[c] ?? 0), { fill: bg, w: dataColW, size: 16 })),
      textCell(r, bg, firstColW, true),
    ]});
  });

  const totalRow = new TableRow({ children: [
    tcNum(String(grandTotal), { fill: _C.light, bold: true, w: totalColW, size: 18 }),
    ...cols.slice().reverse().map(c =>
      tcNum(String(colTotals[c]), { fill: _C.light, bold: true, w: dataColW, size: 17 })),
    tc("الإجمالي", { fill: _C.light, bold: true, w: firstColW, size: 18 }),
  ]});

  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: colWidths,
    rows: [headerRow, ...dataRows, totalRow],
  });
}

// Fixed-category headcount table (degree × department) shown near the top of
// the report, before أولاً, for faculty/assistant/coordinator surveys — see
// FIXED_PARTICIPANT_CONFIG in analyze.js for the category lists. Unlike
// buildParticipantsSection below (dynamic, whatever values appear in the
// file), every category always shows here, in a fixed order, even with 0.
const FIXED_PARTICIPANTS_TITLE = {
  faculty:     "بيان بعدد أعضاء هيئة التدريس المشاركين بالاستبيان",
  assistant:   "بيان بعدد أعضاء الهيئة المعاونة المشاركين بالاستبيان",
  coordinator: "بيان بعدد المشاركين بالاستبيان",
};
const FIXED_PARTICIPANTS_ROW_LABEL = {
  faculty:     "الدرجة العلمية / القسم",
  assistant:   "الدرجة العلمية / القسم",
  coordinator: "الوظيفة / القسم",
};

function buildFixedParticipantsSection(result) {
  const { fixedParticipants, schemaId, n } = result;
  if (!fixedParticipants) return [];
  return [
    sectionHeading(`  ${FIXED_PARTICIPANTS_TITLE[schemaId] ?? "بيان بعدد المشاركين بالاستبيان"}  `),
    spacer(),
    bodyPara(`بلغ إجمالي المشاركين في الاستبيان (${n}) مشاركاً.`, { bold: true }),
    spacer(),
    buildCrosstabTable(fixedParticipants, FIXED_PARTICIPANTS_ROW_LABEL[schemaId] ?? "الدرجة العلمية / القسم"),
    spacer(),
  ];
}

function buildParticipantsSection(result) {
  const { byDegree, byDepartment, crossDegreeByDept, fixedParticipants, n } = result;
  const nodes = [];

  // Covered earlier in the report by buildFixedParticipantsSection — skip
  // the duplicate dynamic table for these schemas.
  if (fixedParticipants) return nodes;

  if (crossDegreeByDept) {
    const isCoord = result.schemaId === "coordinator";
    const crossTitle = isCoord ? "توزيع المشاركين حسب الوظيفة والتخصص:" : "توزيع المشاركين حسب المسمى الوظيفي والقسم:";
    const rowLabel  = isCoord ? "الوظيفة / التخصص" : "المسمى الوظيفي / القسم";
    nodes.push(sectionHeading("  ثالثاً: بيان بعدد المشاركين بالاستبيان  "));
    nodes.push(bodyPara(`بلغ إجمالي المشاركين في الاستبيان (${n}) مشاركاً.`, { bold: true }));
    nodes.push(spacer(), bodyPara(crossTitle, { bold: true }));
    nodes.push(buildCrosstabTable(crossDegreeByDept, rowLabel));
    return nodes;
  }

  const makeTable = (counts, header) => {
    const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    if (!entries.length) return null;
    const total = entries.reduce((s, [, c]) => s + c, 0);
    const cw = [Math.floor(CW * 0.2), Math.floor(CW * 0.2), Math.floor(CW * 0.6)];
    return new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: cw,
      rows: [
        new TableRow({ tableHeader: true, children: [
          hCell("النسبة %", cw[0]), hCell("العدد", cw[1]),
          cell([rp([mk(header, { size: 17, bold: true, color: "FFFFFF" })], { alignment: _textAlign })],
            { fill: _C.primary, w: cw[2] }),
        ]}),
        ...entries.map(([label, count], i) => {
          const bg = i % 2 === 0 ? _C.lighter : _C.white;
          return new TableRow({ children: [
            tcNum(`${(count / total * 100).toFixed(1)}%`, { fill: bg, w: cw[0] }),
            tcNum(String(count), { fill: bg, w: cw[1] }),
            textCell(label, bg, cw[2], true),
          ]});
        }),
        new TableRow({ children: [
          tcNum("100%",        { fill: _C.light, bold: true, w: cw[0], size: 18 }),
          tcNum(String(total), { fill: _C.light, bold: true, w: cw[1], size: 18 }),
          tc("الإجمالي",      { fill: _C.light, bold: true, w: cw[2], size: 18 }),
        ]}),
      ],
    });
  };

  if (Object.keys(byDegree).length || Object.keys(byDepartment).length) {
    nodes.push(sectionHeading("  ثالثاً: بيان بعدد المشاركين بالاستبيان  "));
    nodes.push(bodyPara(`بلغ إجمالي المشاركين في الاستبيان (${n}) مشاركاً.`, { bold: true }));
  }
  if (Object.keys(byDepartment).length) {
    nodes.push(spacer(), bodyPara("توزيع المشاركين حسب القسم:", { bold: true }));
    const t = makeTable(byDepartment, "القسم"); if (t) nodes.push(t);
  }
  if (Object.keys(byDegree).length) {
    nodes.push(spacer(), bodyPara("توزيع المشاركين حسب المسمى الوظيفي:", { bold: true }));
    const t = makeTable(byDegree, "المسمى الوظيفي"); if (t) nodes.push(t);
  }
  return nodes;
}

// ── Detailed results (رابعاً) ─────────────────────────────────────────────────
// One big table; axis names appear as merged gray header rows separating their question rows.
function buildDetailedSection(result) {
  const { axes, scaleType } = result;
  const is5 = scaleType === "likert-5";
  const fmtPct = v => (v === 0 || v == null) ? "-" : `${v}%`;

  // Column widths (declared right-to-left: index 0 = rightmost visible column)
  const cw = is5
    ? (() => {
        // [seq, العبارات, لا أوافق بشدة, لا أوافق, محايد, أوافق, أوافق بشدة] — RTL declaration order
        const seqW = Math.floor(CW * 0.05);
        const textW = Math.floor(CW * 0.45);
        const pctW = Math.floor((CW - seqW - textW) / 5);
        return [pctW, pctW, pctW, pctW, pctW, textW, seqW];
      })()
    : (() => {
        // [أوافق %, محايد %, لا أوافق %, العبارات, م]
        const seqW = Math.floor(CW * 0.06);
        const textW = Math.floor(CW * 0.55);
        const pctW = Math.floor((CW - seqW - textW) / 3);
        return [pctW, pctW, pctW, textW, seqW];
      })();

  const colCount = cw.length;

  // Header labels in RTL declaration order (index 0 = rightmost cell)
  const headerLabels = is5
    ? ["أوافق بشدة", "أوافق", "محايد", "لا أوافق", "لا أوافق بشدة", "العبارات", "م"]
    : ["أوافق %", "محايد %", "لا أوافق %", "العبارات", "م"];

  const headerRow = new TableRow({
    tableHeader: true,
    children: headerLabels.map((label, i) => {
      const w = cw[i];
      const isTextCol = i === colCount - 2;
      return isTextCol
        ? cell(
            [rp([mk(label, { size: 17, bold: true, color: "FFFFFF" })], { alignment: _textAlign })],
            { fill: _C.primary, w }
          )
        : hCell(label, w, 16);
    }),
  });

  const allRows = [headerRow];

  axes.forEach(ax => {
    // Axis-name row: single cell merged across all columns
    allRows.push(new TableRow({
      children: [cell(
        [rp([mk(ax.name, { size: 18, bold: true, color: _C.primary })], { alignment: _textAlign })],
        { fill: ACADEMIC.headRowFill, w: CW, columnSpan: colCount }
      )],
    }));

    ax.questions.forEach((q, qi) => {
      const bg = qi % 2 === 0 ? _C.white : ACADEMIC.zebra;

      const pctCells = is5
        ? [
            tcNum(fmtPct(q.pcts["5"]), { fill: bg, w: cw[0], size: 16 }),
            tcNum(fmtPct(q.pcts["4"]), { fill: bg, w: cw[1], size: 16 }),
            tcNum(fmtPct(q.pcts["3"]), { fill: bg, w: cw[2], size: 16 }),
            tcNum(fmtPct(q.pcts["2"]), { fill: bg, w: cw[3], size: 16 }),
            tcNum(fmtPct(q.pcts["1"]), { fill: bg, w: cw[4], size: 16 }),
          ]
        : [
            tcNum(fmtPct(q.pcts["agree"]),    { fill: bg, w: cw[0], size: 16 }),
            tcNum(fmtPct(q.pcts["neutral"]),  { fill: bg, w: cw[1], size: 16 }),
            tcNum(fmtPct(q.pcts["disagree"]), { fill: bg, w: cw[2], size: 16 }),
          ];

      allRows.push(new TableRow({
        children: [
          ...pctCells,
          cell(
            [rp([mk(q.text, { size: 17 })], { alignment: _textAlign })],
            { fill: bg, w: cw[colCount - 2] }
          ),
          tcNum(String(q.seq), { fill: bg, w: cw[colCount - 1], size: 16 }),
        ],
      }));
    });
  });

  return [
    new Paragraph({ children: [new PageBreak()] }),
    sectionHeading("  رابعاً: عرض النتائج وتحليلها ومناقشتها  "),
    spacer(),
    new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: cw,
      rows: allRows,
    }),
    spacer(),
  ];
}

// ── Summary table (خامساً) ────────────────────────────────────────────────────
function buildSummarySection(result) {
  const { axes, overallAgreePct, overallMean, overallDirection, scaleType } = result;
  const is5 = scaleType === "likert-5";
  const cw = is5
    ? [Math.floor(CW * 0.15), Math.floor(CW * 0.15), Math.floor(CW * 0.15), Math.floor(CW * 0.47), Math.floor(CW * 0.08)]
    : [Math.floor(CW * 0.20), Math.floor(CW * 0.15),                         Math.floor(CW * 0.57), Math.floor(CW * 0.08)];

  const axisCol = cell(
    [rp([mk("المحور", { size: 17, bold: true, color: "FFFFFF" })], { alignment: _textAlign })],
    { fill: _C.primary, w: is5 ? cw[3] : cw[2] }
  );

  const headerCells = is5
    ? [hCell("الاتجاه العام", cw[0]), hCell("النسبة", cw[1]), hCell("المتوسط الحسابي", cw[2]), axisCol, hCell("م", cw[4])]
    : [hCell("الاتجاه العام", cw[0]), hCell("النسبة", cw[1]), axisCol, hCell("م", cw[3])];

  const dataRows = axes.map((ax, i) => {
    const bg = i % 2 === 0 ? _C.white : ACADEMIC.zebra;
    return new TableRow({ children: is5
      ? [tc(ax.direction,                 { fill: bg, w: cw[0], size: 17 }),
         tcNum(`${ax.axisAgreePct}%`,     { fill: bg, w: cw[1], size: 17, bold: true }),
         tcNum(String(ax.axisMean ?? ""), { fill: bg, bold: true, w: cw[2], size: 17 }),
         textCell(ax.name, bg, cw[3], true),
         tcNum(`${i + 1}-`,               { fill: bg, w: cw[4], size: 17 })]
      : [tc(ax.direction,                 { fill: bg, w: cw[0], size: 17 }),
         tcNum(`${ax.axisAgreePct}%`,     { fill: bg, w: cw[1], size: 17, bold: true }),
         textCell(ax.name, bg, cw[2], true),
         tcNum(`${i + 1}-`,               { fill: bg, w: cw[3], size: 17 })]
    });
  });

  const totalRow = new TableRow({ children: is5
    ? [tc(overallDirection,             { fill: ACADEMIC.total, bold: true, w: cw[0], size: 18 }),
       tcNum(`${overallAgreePct}%`,     { fill: ACADEMIC.total, bold: true, w: cw[1], size: 18 }),
       tcNum(String(overallMean ?? ""), { fill: ACADEMIC.total, bold: true, w: cw[2], size: 18 }),
       tc("متوسط الرضا",                { fill: ACADEMIC.total, bold: true, w: cw[3], size: 18 }),
       tcNum("",                         { fill: ACADEMIC.total, w: cw[4] })]
    : [tc(overallDirection,             { fill: ACADEMIC.total, bold: true, w: cw[0], size: 18 }),
       tcNum(`${overallAgreePct}%`,     { fill: ACADEMIC.total, bold: true, w: cw[1], size: 18 }),
       tc("متوسط الرضا",                { fill: ACADEMIC.total, bold: true, w: cw[2], size: 18 }),
       tcNum("",                         { fill: ACADEMIC.total, w: cw[3] })]
  });

  return [
    new Paragraph({ children: [new PageBreak()] }),
    sectionHeading("  خامساً: ملخص النتائج  "),
    spacer(),
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: cw,
      rows: [new TableRow({ tableHeader: true, children: headerCells }), ...dataRows, totalRow] }),
    spacer(),
  ];
}

// ── Recommendations (أخيراً) ──────────────────────────────────────────────────
function buildRecommendationsSection(result, s) {
  const RECOMMEND_THRESHOLD = 70;
  const candidates = result.axes.filter(a => a.axisAgreePct < RECOMMEND_THRESHOLD);

  if (candidates.length === 0) {
    return [
      spacer(),
      sectionHeading("  أخيراً: التوصيات  "),
      spacer(),
      bodyPara("لا يوجد توصيات.", { bold: true, size: 22 }),
    ];
  }

  const count = Math.max(1, Math.min(s.recommendationsCount, candidates.length));
  const lowAxes = [...candidates].sort((a, b) => a.axisAgreePct - b.axisAgreePct).slice(0, count);

  return [
    spacer(),
    sectionHeading("  أخيراً: التوصيات  "),
    spacer(),
    ...lowAxes.flatMap(ax => {
      const text = ax.recommendation
        ?? `مراجعة محور "${ax.name}" لأنه سجل نسبة موافقة ${ax.axisAgreePct}%.`;
      return [
        bodyPara(`• ${ax.name} (${ax.axisAgreePct}%)`, { bold: true, size: 21, color: _C.primary }),
        bodyPara(text, { size: 20 }),
        spacer(),
      ];
    }),
  ];
}

// ── Signature block ───────────────────────────────────────────────────────────
function buildSignatureBlock(s) {
  const hw = Math.floor(CW / 2);
  return [
    spacer(), spacer(),
    new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: [hw, hw],
      rows: [
        new TableRow({ children: [
          tc("رئيس وحدة القياس والتقويم", { fill: _C.white, bold: true, color: _C.primary, w: hw, size: 20, noBorder: true }),
          tc("رئيس وحدة الجودة",          { fill: _C.white, bold: true, color: _C.primary, w: hw, size: 20, noBorder: true }),
        ]}),
        new TableRow({ children: [
          tc(s.qmName, { fill: _C.white, color: _C.darkGray, w: hw, size: 18, noBorder: true }),
          tc(s.quName, { fill: _C.white, color: _C.darkGray, w: hw, size: 18, noBorder: true }),
        ]}),
      ],
    }),
  ];
}

// ── Comparison section ────────────────────────────────────────────────────────
function buildComparisonSection(comparison, meta) {
  const { axes, overall } = comparison;
  const years = overall.map(o => o.year);
  const yCount = years.length;
  const nameW = Math.floor(CW * 0.4);
  const dataW = Math.floor((CW - nameW - 700 - 700) / yCount);
  const colWidths = [700, nameW, ...Array(yCount).fill(dataW), 700];

  const headerChildren = [
    hCell("الاتجاه", colWidths[colWidths.length - 1]),
    ...years.map((y, i) => hCell(y, colWidths[2 + i], 15)).reverse(),
    cell([rp([mk("المحور", { size: 17, bold: true, color: "FFFFFF" })], { alignment: _textAlign })],
      { fill: _C.primary, w: colWidths[1] }),
    hCell("م", colWidths[0]),
  ];

  const dataRows = axes.map((ax, i) => {
    const bg = i % 2 === 0 ? _C.white : _C.lighter;
    const trendFill = ax.trend === "تحسن" ? "EBF5EB" : ax.trend === "تراجع" ? _C.red : bg;
    return new TableRow({ children: [
      tc(ax.trend ?? "—", { fill: trendFill, bold: true, w: colWidths[colWidths.length - 1], size: 15 }),
      ...ax.years.map((y, yi) =>
        tc(y.agreePct !== null ? `${y.agreePct}%` : "—", { fill: bg, w: colWidths[2 + yi], bold: true, size: 16 })
      ).reverse(),
      textCell(ax.name, bg, colWidths[1], true),
      tc(String(i + 1), { fill: bg, w: colWidths[0], size: 16 }),
    ]});
  });

  const overallRow = new TableRow({ children: [
    tc("", { fill: _C.light, w: colWidths[colWidths.length - 1] }),
    ...overall.map((o, oi) =>
      tc(`${o.agreePct}%`, { fill: _C.light, bold: true, w: colWidths[2 + oi], size: 17 })
    ).reverse(),
    tc("الإجمالي", { fill: _C.light, bold: true, w: colWidths[1], size: 17 }),
    tc("", { fill: _C.light, w: colWidths[0] }),
  ]});

  return [
    new Paragraph({ children: [new PageBreak()] }),
    sectionHeading(`  سادساً: مقارنة نتائج ${yCount} سنوات  `),
    spacer(),
    ...(meta.program ? [bodyPara(`البرنامج / القسم: ${meta.program}`, { bold: true, size: 21 })] : []),
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: colWidths,
      rows: [new TableRow({ tableHeader: true, children: headerChildren }), ...dataRows, overallRow] }),
    spacer(),
    bodyPara("الأخضر = تحسن  |  الأحمر = تراجع  |  بدون لون = استقرار", { size: 18, color: "666666" }),
  ];
}

// ── Coordinator detailed results (رابعاً) ─────────────────────────────────────
function buildCoordinatorDetailedSection(result) {
  const { axes } = result;
  const fmtPct = v => (v === 0 || v == null) ? "-" : `${v}%`;

  const seqW  = Math.floor(CW * 0.05);
  const textW = Math.floor(CW * 0.37);
  const meanW = Math.floor(CW * 0.08);
  const aptW  = Math.floor(CW * 0.08);
  const pctW  = Math.floor((CW - seqW - textW - meanW - aptW) / 5);
  // Declaration order: leftmost → rightmost (RTL document: first child = leftmost visual)
  const cw    = [aptW, meanW, pctW, pctW, pctW, pctW, pctW, textW, seqW];

  const headerLabels = ["نسبة الرضا", "المتوسط", "أوافق بشدة", "أوافق", "محايد", "لا أوافق", "لا أوافق بشدة", "العبارة", "م"];
  const headerRow = new TableRow({
    tableHeader: true,
    children: headerLabels.map((label, i) =>
      i === 7
        ? cell([rp([mk(label, { size: 16, bold: true, color: "FFFFFF" })], { alignment: _textAlign })],
            { fill: _C.primary, w: cw[i] })
        : hCell(label, cw[i], 15)
    ),
  });

  const allRows = [headerRow];

  axes.forEach(ax => {
    allRows.push(new TableRow({
      children: [cell(
        [rp([mk(ax.name, { size: 18, bold: true, color: _C.primary })], { alignment: _textAlign })],
        { fill: ACADEMIC.headRowFill, w: CW, columnSpan: 9 }
      )],
    }));

    ax.questions.forEach((q, qi) => {
      const bg = qi % 2 === 0 ? _C.white : ACADEMIC.zebra;
      allRows.push(new TableRow({
        children: [
          tcNum(fmtPct(q.agreePct),              { fill: bg, w: cw[0], size: 16 }),
          tcNum(q.mean != null ? String(q.mean) : "-", { fill: bg, w: cw[1], size: 16 }),
          tcNum(fmtPct(q.pcts["5"]),             { fill: bg, w: cw[2], size: 16 }),
          tcNum(fmtPct(q.pcts["4"]),             { fill: bg, w: cw[3], size: 16 }),
          tcNum(fmtPct(q.pcts["3"]),             { fill: bg, w: cw[4], size: 16 }),
          tcNum(fmtPct(q.pcts["2"]),             { fill: bg, w: cw[5], size: 16 }),
          tcNum(fmtPct(q.pcts["1"]),             { fill: bg, w: cw[6], size: 16 }),
          cell([rp([mk(q.text, { size: 16 })], { alignment: _textAlign })], { fill: bg, w: cw[7] }),
          tcNum(String(q.seq),                   { fill: bg, w: cw[8], size: 16 }),
        ],
      }));
    });

    // Axis satisfaction summary row
    allRows.push(new TableRow({
      children: [
        tcNum(`${ax.axisAgreePct}%`,          { fill: _C.lightBlue, bold: true, w: cw[0], size: 17 }),
        tcNum(String(ax.axisMean ?? ""),      { fill: _C.lightBlue, bold: true, w: cw[1], size: 17 }),
        cell(
          [rp([mk("متوسط الرضا للمحور", { size: 17, bold: true, color: _C.primary })], { alignment: AlignmentType.CENTER })],
          { fill: _C.lightBlue, w: 5 * pctW, columnSpan: 5 }
        ),
        tc("", { fill: _C.lightBlue, w: cw[7] }),
        tc("", { fill: _C.lightBlue, w: cw[8] }),
      ],
    }));
  });

  return [
    new Paragraph({ children: [new PageBreak()] }),
    sectionHeading("  رابعاً: عرض النتائج وتحليلها  "),
    spacer(),
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: cw, rows: allRows }),
    spacer(),
  ];
}

// ── Coordinator summary (خامساً) ──────────────────────────────────────────────
function buildCoordinatorSummarySection(result) {
  const { axes, overallAgreePct, overallMean } = result;
  const mW  = Math.floor(CW * 0.07);
  const vW  = Math.floor(CW * 0.12);
  const axW = CW - vW * 2 - mW;
  const cw  = [vW, vW, axW, mW]; // [نسبة الرضا, متوسط الرضا, المحور, م]

  const headerRow = new TableRow({ tableHeader: true, children: [
    hCell("نسبة الرضا",  cw[0], 16),
    hCell("متوسط الرضا", cw[1], 16),
    cell([rp([mk("المحور", { size: 17, bold: true, color: "FFFFFF" })], { alignment: _textAlign })],
      { fill: _C.primary, w: cw[2] }),
    hCell("م", cw[3], 16),
  ]});

  const dataRows = axes.map((ax, i) => {
    const bg = i % 2 === 0 ? _C.white : ACADEMIC.zebra;
    return new TableRow({ children: [
      tcNum(`${ax.axisAgreePct}%`,        { fill: bg, w: cw[0], size: 17, bold: true }),
      tcNum(String(ax.axisMean ?? ""),    { fill: bg, w: cw[1], size: 17, bold: true }),
      textCell(ax.name, bg, cw[2], true),
      tcNum(String(i + 1),               { fill: bg, w: cw[3], size: 17 }),
    ]});
  });

  const totalRow = new TableRow({ children: [
    tcNum(`${overallAgreePct}%`,        { fill: ACADEMIC.total, bold: true, w: cw[0], size: 18 }),
    tcNum(String(overallMean ?? ""),    { fill: ACADEMIC.total, bold: true, w: cw[1], size: 18 }),
    tc("متوسط الرضا العام",             { fill: ACADEMIC.total, bold: true, w: cw[2], size: 18 }),
    tcNum("",                           { fill: ACADEMIC.total, w: cw[3] }),
  ]});

  return [
    new Paragraph({ children: [new PageBreak()] }),
    sectionHeading("  خامساً: ملخص النتائج  "),
    spacer(),
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: cw,
      rows: [headerRow, ...dataRows, totalRow] }),
    spacer(),
  ];
}

// ── Document factory ──────────────────────────────────────────────────────────
function makeDoc(children, header, footer) {
  return new Document({
    styles: {
      default: {
        document: { run: { font: _font, size: 20 } },
      },
      paragraphStyles: [{
        id: "Normal",
        name: "Normal",
        basedOn: "Normal",
        next: "Normal",
        run: { font: _font, size: 20 },
        paragraph: { bidirectional: _isRTL, alignment: _isRTL ? AlignmentType.RIGHT : AlignmentType.LEFT },
      }],
    },
    sections: [{
      properties: {
        page: { size: { width: PAGE_W, height: 16838 }, margin: MARGINS },
        bidi: _isRTL,
      },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });
}

// ── Design helpers ────────────────────────────────────────────────────────────
function toAlignType(val) {
  if (val === "center") return AlignmentType.CENTER;
  if (val === "left")   return AlignmentType.LEFT;
  return AlignmentType.RIGHT;
}

// docx 9.x has NO section-level RTL option (it ignores section `bidi`), and our
// tables are not marked bidiVisual — so the ENTIRE document renders in an LTR
// frame. Word then resolves every bidi (RTL) paragraph's physical left/right
// against that LTR frame, flipping them: `jc=right` ends up hugging the visual
// LEFT. To make Arabic text (body headings, paragraphs, footer AND table cells)
// sit on the physical RIGHT (reading start) we emit the opposite alignment.
// Every RTL paragraph alignment in this file is routed through cellAlign().
function cellAlign(a) {
  if (!_isRTL) return a;
  if (a === AlignmentType.RIGHT) return AlignmentType.LEFT;
  if (a === AlignmentType.LEFT)  return AlignmentType.RIGHT;
  return a; // center / both unaffected
}

function applyDesign(s) {
  _C         = COLOR_THEMES[s.colorTheme] ?? COLORS;
  _font      = s.reportFont      || FONT;
  _isRTL     = (s.reportDirection || "rtl") === "rtl";
  _textAlign = cellAlign(toAlignType(s.textAlign || (_isRTL ? "right" : "left")));
  _numAlign  = cellAlign(toAlignType(s.numAlign  || "center"));
}

// ── Exported builders ─────────────────────────────────────────────────────────
export async function buildAnnualDocx(result, meta, settings = {}) {
  const s = { ...DEFAULT_SETTINGS, ...settings };
  applyDesign(s);
  const logoData = await getLogoData(s);
  const isCoord = result.schemaId === "coordinator";
  const children = [
    ...buildCoverPage(result, meta),
    ...(s.includeEvaluatorsTable ? buildEvaluatorsSection(meta, s) : []),
    ...(s.includeParticipants ? buildFixedParticipantsSection(result) : []), // قبل أولاً
    ...buildVariablesSection(result),                                       // أولاً
    ...buildMethodologySection(result),                                     // ثانياً
    ...(s.includeParticipants ? buildParticipantsSection(result) : []),     // ثالثاً
    ...(isCoord ? buildCoordinatorDetailedSection(result) : buildDetailedSection(result)),   // رابعاً
    ...(isCoord ? buildCoordinatorSummarySection(result) : buildSummarySection(result)),     // خامساً
    ...(!isCoord && s.includeRecommendations ? buildRecommendationsSection(result, s) : []), // أخيراً
    ...buildSignatureBlock(s),
  ];
  return Packer.toBlob(makeDoc(children, buildHeader(logoData, s), buildFooter(s)));
}

export async function buildComparisonDocx(comparison, meta, settings = {}) {
  const s = { ...DEFAULT_SETTINGS, ...settings };
  applyDesign(s);
  const logoData = await getLogoData(s);
  const children = [
    ...buildCoverPage(comparison.slots[0].result, meta),
    ...buildComparisonSection(comparison, meta),
    ...buildSignatureBlock(s),
  ];
  return Packer.toBlob(makeDoc(children, buildHeader(logoData, s), buildFooter(s)));
}
