/**
 * buildReportPdf.js
 * ──────────────────────────────────────────────────────────────────────────────
 * Generates a branded, print-styled PDF version of an annual survey report
 * (cover page, vision/mission, methodology, results summary + chart, one
 * detailed page per axis with per-question bars, recommendations) directly
 * from the same `result` / `meta` / `settings` objects buildAnnualDocx() and
 * EnhancedReportView already consume — no new data model, no schema changes.
 *
 * Rendered off-screen as HTML, one `.pdf-page-outer` section at a time, and
 * rasterized via html2canvas + jsPDF directly (not html2pdf.js's one-shot
 * whole-document capture — a report with many axes easily produces a canvas
 * taller than the browser's max canvas height, which silently rasterizes as
 * blank with no thrown error; per-section capture keeps every canvas well
 * under that limit, and any section taller than one physical page is sliced
 * across multiple PDF pages instead of being cropped).
 */

// ── design tokens (matches the pink/coral ERU report template) ────────────────
const PINK       = "#f6d7d7";
const PINK_SOFT  = "#fbeaea";
const PINK_LINE  = "#e6a9a9";
const ACCENT     = "#b3373a";
const TEXT_DARK  = "#22262b";
const TEXT_MUTED = "#6b7280";
const BAR_COLORS = ["#e74c3c", "#f39c12", "#27ae60", "#17a2b8", "#9b59b6", "#2980b9", "#e67e22", "#16a085"];
const PDF_THEMES = {
  default: { accent: "#1f3864", soft: "#ebf3fb", heading: "#d6e4f0", line: "#9fbad0" },
  green:   { accent: "#1a4731", soft: "#eaf6ee", heading: "#d4edda", line: "#a8d5b5" },
  purple:  { accent: "#4a235a", soft: "#f5eef8", heading: "#e8daef", line: "#d7bde2" },
  dark:    { accent: "#1c1c1c", soft: "#f5f5f5", heading: "#e0e0e0", line: "#c0c0c0" },
  red:     { accent: ACCENT, soft: PINK_SOFT, heading: PINK, line: PINK_LINE },
};

function esc(v) {
  return String(v ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function reportText(result, key, fallback, meta = {}) {
  const raw = result.reportTexts?.[key]?.trim() || fallback;
  const values = {
    "{اسم_الاستبيان}": result.schemaLabel ?? "",
    "{عدد_المشاركين}": result.n ?? "",
    "{عدد_الأسئلة}": result.totalQuestions ?? "",
    "{عدد_المحاور}": result.axes?.length ?? "",
    "{العام}": meta.year ?? "",
    "{البرنامج}": meta.program ?? "",
  };
  return Object.entries(values).reduce((text, [token, value]) => text.split(token).join(String(value)), raw);
}

function barColor(i) {
  return BAR_COLORS[i % BAR_COLORS.length];
}

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function balancedChunks(arr, maxSize) {
  if (!arr.length) return [];
  const pageCount = Math.ceil(arr.length / maxSize);
  return chunk(arr, Math.ceil(arr.length / pageCount));
}

// Rows/bars per page are picked so a page never has to be split mid-row —
// html2canvas captures one whole .pdf-page-outer per page, so any content
// that overflows the physical page height gets sliced at an arbitrary pixel
// boundary (mid-table-row, mid-bar) further down in downloadBrandedReportPdf.
// Chunking the summary table, the axis chart, and the recommendations table
// up front avoids that for the sections most likely to be long.
const SUMMARY_ROWS_PER_PAGE = 16;
const RECS_ROWS_PER_PAGE = 14;

// ── small building blocks ──────────────────────────────────────────────────────
function pageOpen(extraClass = "") {
  return `<div class="pdf-page-outer"><div class="pdf-page ${extraClass}">`;
}

function pageClose(s) {
  const qualityEmail = String(s.qualityEmail || "qa-mebt@eru.edu.eg").trim();
  const measurementEmail = String(s.measurementEmail || "meb-maec@eru.edu.eg").trim();
  return `
    <div class="pdf-page-footer">
      <i class="pdf-capture-bottom-marker" aria-hidden="true"></i>
      ${qualityEmail ? `<span>Quality Assurance: ${esc(qualityEmail)}</span>` : ""}
      ${measurementEmail ? `<span>Measurement &amp; Evaluation: ${esc(measurementEmail)}</span>` : ""}
    </div>
  </div></div>`;
}

function pageHeader(s, logoSrc) {
  return `
    <div class="pdf-header">
      ${logoSrc ? `<img class="pdf-logo" src="${logoSrc}" alt="" />` : ""}
      <div class="pdf-header-text">
        <i class="pdf-capture-top-marker" aria-hidden="true"></i>
        <div class="pdf-uni-name">${esc(s.uniName)}</div>
        <div class="pdf-faculty-name">${esc(s.facultyName)}</div>
      </div>
    </div>`;
}

function sectionHeading(text) {
  return `<div class="pdf-section-heading"><span>${esc(text)}</span></div>`;
}

function infoRow(label, value, i) {
  return `
    <div class="pdf-info-row ${i % 2 === 0 ? "alt" : ""}">
      <div class="pdf-info-label">${esc(label)}</div>
      <div class="pdf-info-val">${esc(value)}</div>
    </div>`;
}

function axesOverviewChart(items) {
  return `
    <div class="pdf-axes-chart ${items.length > 13 ? "two-columns" : ""}">
      ${items.map((it, i) => `
        <div class="pdf-axis-chart-row">
          <div class="pdf-axis-chart-head">
            <span class="pdf-axis-chart-label">${esc(it.label)}</span>
            <span class="pdf-axis-chart-pct">${it.pct}%</span>
          </div>
          <div class="pdf-axis-chart-track">
            <div class="pdf-axis-chart-fill" style="width:${Math.max(2, it.pct)}%;background:${barColor(i)}"></div>
          </div>
        </div>
      `).join("")}
    </div>`;
}

function horizontalBar(label, pct, i) {
  return `
    <div class="pdf-hbar-row">
      <div class="pdf-hbar-track">
        <div class="pdf-hbar-fill" style="width:${Math.max(3, pct)}%;background:${barColor(i)}">
          <span class="pdf-hbar-pct">${pct}%</span>
        </div>
      </div>
      <div class="pdf-hbar-label">${esc(label)}</div>
    </div>`;
}

// ── page builders ───────────────────────────────────────────────────────────────
function buildCoverPage(result, meta, s, logoSrc) {
  const { schemaLabel, n } = result;
  const rows = [
    ["اسم الاستبيان", schemaLabel],
    ...(meta.program ? [["البرنامج / القسم", meta.program]] : []),
    ["عدد المشاركين في التحليل", n],
    ["العام الدراسي", meta.year || "—"],
    ["موجه إلى", "مدير وحدة الجودة بالكلية"],
  ];

  const signers = s.includeEvaluatorsTable === false ? [] : [
    ...(s.includeCommitteeHead !== false ? [{ name: s.qmName || "", role: "رئيس لجنة القياس والتقويم" }] : []),
    ...(s.includeEvaluator !== false ? [{ name: meta.preparedBy || "", role: "القائم بالتقييم" }] : []),
    ...(s.includeReviewer !== false ? [{ name: meta.reviewer || "", role: "القائم بالمراجعة" }] : []),
  ];

  return `
    ${pageOpen("pdf-cover")}
      ${pageHeader(s, logoSrc)}
      <div class="pdf-cover-title-wrap">
        <div class="pdf-cover-title">${esc(reportText(result, "reportTitle", `استبيان ${schemaLabel}`, meta))}</div>
        <div class="pdf-cover-subtitle">${esc(reportText(result, "reportSubtitle", "تقرير نتائج تحليل الاستبيانات", meta))}</div>
        <div class="pdf-cover-committee">${esc(s.committeeName)}</div>
      </div>
      ${sectionHeading("البيانات الأساسية")}
      <div class="pdf-info-table">
        ${rows.map(([l, v], i) => infoRow(l, v, i)).join("")}
      </div>
      ${signers.length ? `
        ${sectionHeading("القائم بالتقييم")}
        <div class="pdf-sig-row">
          ${signers.map(({ name, role }) => `
            <div class="pdf-sig-box">
              <div class="pdf-sig-name">${esc(name || "—")}</div>
              <div class="pdf-sig-role">${esc(role)}</div>
              <div class="pdf-sig-line">التوقيع: ....................</div>
            </div>`).join("")}
        </div>` : ""}
    ${pageClose(s)}`;
}

function buildContextPage(result, meta, s, logoSrc) {
  const section = result.reportSections?.afterEvaluators;
  const introduction = result.reportTexts?.introduction?.trim()
    ? reportText(result, "introduction", "", meta)
    : "";
  const includeVisionMission = s.includeVisionMission !== false;
  if (!section && !introduction && !includeVisionMission) return "";
  return `
    ${pageOpen("pdf-context-page")}
      ${pageHeader(s, logoSrc)}
      ${introduction ? `
        ${sectionHeading("مقدمة")}
        <p class="pdf-body-text">${esc(introduction)}</p>` : ""}
      ${includeVisionMission ? `
        ${sectionHeading("رؤية ورسالة وحدة ضمان الجودة")}
        <div class="pdf-vm-block">
          <div class="pdf-vm-title">الرؤية</div>
          <div class="pdf-vm-text">${esc(s.vision)}</div>
        </div>
        <div class="pdf-vm-block">
          <div class="pdf-vm-title">الرسالة</div>
          <div class="pdf-vm-text">${esc(s.mission)}</div>
        </div>` : ""}
      ${section ? `
        ${sectionHeading(section.title)}
        ${section.intro ? `<p class="pdf-body-text">${esc(section.intro)}</p>` : ""}
        <ol class="pdf-procedure-list">
          ${(section.items ?? []).map(item => `<li>${esc(item)}</li>`).join("")}
        </ol>` : ""}
    ${pageClose(s)}`;
}

// Same fixed degree×department headcount table as buildFixedParticipantsSection
// in buildDocx.js (see analyze.js's FIXED_PARTICIPANT_CONFIG / fixedCrossTab) —
// only rendered when result.fixedParticipants exists (faculty/assistant/
// coordinator schemas), placed right before the methodology page (أولاً).
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

function buildFixedParticipantsBlock(result) {
  const { fixedParticipants, schemaId, n } = result;
  if (!fixedParticipants) return "";
  const { rows, cols, matrix, rowTotals, colTotals, grandTotal } = fixedParticipants;
  const rowLabel = FIXED_PARTICIPANTS_ROW_LABEL[schemaId] ?? "الدرجة العلمية / القسم";

  return `
    ${sectionHeading(FIXED_PARTICIPANTS_TITLE[schemaId] ?? "بيان بعدد المشاركين بالاستبيان")}
    <p class="pdf-body-text">بلغ إجمالي المشاركين في الاستبيان (${esc(String(n))}) مشاركاً.</p>
    <table class="pdf-table">
      <thead>
        <tr>
          <th class="pdf-th-wide">${esc(rowLabel)}</th>
          ${cols.map(c => `<th>${esc(c)}</th>`).join("")}
          <th>الإجمالي</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map((r, i) => `
          <tr class="${i % 2 === 0 ? "alt" : ""}">
            <td class="pdf-td-right">${esc(r)}</td>
            ${cols.map(c => `<td>${esc(String(matrix[r][c]))}</td>`).join("")}
            <td class="pdf-td-strong">${esc(String(rowTotals[r]))}</td>
          </tr>`).join("")}
        <tr class="pdf-total-row">
          <td class="pdf-td-right">الإجمالي</td>
          ${cols.map(c => `<td>${esc(String(colTotals[c]))}</td>`).join("")}
          <td>${esc(String(grandTotal))}</td>
        </tr>
      </tbody>
    </table>`;
}

function buildMethodologyPage(result, s, logoSrc) {
  const { totalQuestions, axes, scaleType } = result;
  const is5 = scaleType === "likert-5";
  return `
    ${pageOpen()}
      ${pageHeader(s, logoSrc)}
      ${sectionHeading(`نتائج تحليل ${esc(result.schemaLabel)}`)}
      <p class="pdf-body-text">
        ${esc(reportText(result, "methodologyText", `يتم إجراء التحليل الإحصائي لتقييم الاستبيان وفقاً لعدد من الخطوات، بدءاً بقيام الطلاب/المشاركين
        بتعبئة الاستبيان، ثم تجميع البيانات ومعالجتها إحصائياً باستخدام المعايير والمؤشرات
        المناسبة، وذلك بهدف تقييم جودة العملية التي يقيسها الاستبيان وتحديد نقاط القوة والجوانب
        التي تتطلب التحسين أو التطوير.`))}
      </p>

      ${sectionHeading("متغيرات الاستبيان")}
      <p class="pdf-body-text">
        ${esc(reportText(result, "variablesText", `اشتمل الاستبيان على (${totalQuestions}) عبارة تتمثل في (${axes.length}) محور رئيسي.`))}
      </p>

      ${sectionHeading("المعالجة الإحصائية المستخدمة")}
      <div class="pdf-formula-block">
        <div class="pdf-formula-title">1- حساب النسبة المئوية لكل فئة استجابة:</div>
        <div class="pdf-formula">p<sub>i</sub> = ( f<sub>i</sub> / N ) &times; 100</div>
        <div class="pdf-formula-note">p<sub>i</sub>: النسبة المئوية لكل فئة &nbsp; | &nbsp; f<sub>i</sub>: عدد الاستجابات لكل فئة &nbsp; | &nbsp; N: إجمالي عدد الاستجابات</div>
      </div>
      ${is5 ? `
      <div class="pdf-formula-block">
        <div class="pdf-formula-title">2- حساب الوسط الحسابي لكل محور:</div>
        <div class="pdf-formula">x&#772; = ( &Sigma; x<sub>i</sub> ) / n</div>
        <div class="pdf-formula-note">x&#772;: الوسط الحسابي للمحور &nbsp; | &nbsp; x<sub>i</sub>: درجة كل عبارة &nbsp; | &nbsp; n: عدد العبارات الخاصة بالمحور</div>
      </div>` : ""}

      ${buildFixedParticipantsBlock(result)}
    ${pageClose(s)}`;
}

function buildSummaryTablePages(result, s, logoSrc) {
  const { axes, n, overallAgreePct, overallDirection } = result;
  // Fill each summary page before opening the next one. The previous balanced
  // split (for example 9 + 8 rows) left a large unused area on the first page.
  const pages = chunk(axes, SUMMARY_ROWS_PER_PAGE);

  return pages.map((pageAxes, pageIdx) => {
    const isLast = pageIdx === pages.length - 1;
    return `
      ${pageOpen()}
        ${pageHeader(s, logoSrc)}
        ${sectionHeading(pageIdx === 0 ? reportText(result, "summaryHeading", "ملخص النتائج") : `${reportText(result, "summaryHeading", "ملخص النتائج")} (تابع)`)}
        <table class="pdf-table">
          <colgroup><col style="width:58%"><col style="width:12%"><col style="width:13%"><col style="width:17%"></colgroup>
          <thead>
            <tr><th class="pdf-th-wide">المحور</th><th>العدد</th><th>النسبة</th><th>الاتجاه العام</th></tr>
          </thead>
          <tbody>
            ${pageAxes.map((ax, i) => `
              <tr class="${i % 2 === 0 ? "alt" : ""}">
                <td class="pdf-td-right">${esc(ax.name)}</td>
                <td>${n}</td>
                <td class="pdf-td-strong">${ax.axisAgreePct}%</td>
                <td>${esc(ax.direction)}</td>
              </tr>`).join("")}
            ${isLast ? `
              <tr class="pdf-total-row">
                <td class="pdf-td-right">المتوسط العام</td>
                <td>${n}</td>
                <td>${overallAgreePct}%</td>
                <td>${esc(overallDirection)}</td>
              </tr>` : ""}
          </tbody>
        </table>
        ${isLast ? `
          <div class="pdf-summary-chart">
            ${sectionHeading("التمثيل البياني لنسب تحقق المحاور")}
            ${axesOverviewChart(axes.map(ax => ({ label: ax.name, pct: ax.axisAgreePct })))}
          </div>` : ""}
      ${pageClose(s)}`;
  }).join("");
}

function buildCompactFlatSummary(result) {
  const { axes, n, overallAgreePct, overallDirection } = result;
  const ax = axes[0];
  if (!ax) return "";
  return `
    <div class="pdf-compact-summary">
      ${sectionHeading(reportText(result, "summaryHeading", "ملخص النتائج"))}
      <table class="pdf-table">
        <colgroup><col style="width:58%"><col style="width:12%"><col style="width:13%"><col style="width:17%"></colgroup>
        <thead>
          <tr><th class="pdf-th-wide">المحور</th><th>العدد</th><th>النسبة</th><th>الاتجاه العام</th></tr>
        </thead>
        <tbody>
          <tr class="alt">
            <td class="pdf-td-right">${esc(ax.name)}</td>
            <td>${n}</td>
            <td class="pdf-td-strong">${ax.axisAgreePct}%</td>
            <td>${esc(ax.direction)}</td>
          </tr>
          <tr class="pdf-total-row">
            <td class="pdf-td-right">المتوسط العام</td>
            <td>${n}</td>
            <td>${overallAgreePct}%</td>
            <td>${esc(overallDirection)}</td>
          </tr>
        </tbody>
      </table>
      <div class="pdf-summary-chart">
        ${sectionHeading("التمثيل البياني لنسب تحقق المحاور")}
        ${axesOverviewChart(axes.map(item => ({ label: item.name, pct: item.axisAgreePct })))}
      </div>
    </div>`;
}

function buildAxisBlock(ax, result, includeCharts, options = {}) {
  const continued = options.continued === true;
  const showTotal = options.showTotal !== false;
  const is5 = result.scaleType === "likert-5";
  const scale5Labels = result.scaleValues?.length === 5
    ? result.scaleValues.map(v => v.label)
    : ["لا أوافق بشدة", "لا أوافق", "محايد", "أوافق", "أوافق بشدة"];
  const headerLabels = is5
    ? ["م", "العبارات", ...scale5Labels]
    : ["م", "العبارات", "لا أوافق %", "محايد %", "أوافق %"];

  const fmtPct = v => (v === 0 || v == null) ? "-" : `${v}%`;

  return `
    ${result.isFlat ? "" : sectionHeading(`${ax.name}${continued ? " (تابع)" : ""}`)}
    <table class="pdf-table pdf-table-detail">
      <colgroup>
        <col style="width:6%"><col style="width:${is5 ? 54 : 64}%">
        ${headerLabels.slice(2).map(() => `<col style="width:${is5 ? 8 : 10}%">`).join("")}
      </colgroup>
      <thead>
        <tr>${headerLabels.map(l => `<th>${esc(l)}</th>`).join("")}</tr>
      </thead>
      <tbody>
        ${ax.questions.map((q, qi) => `
          <tr class="${qi % 2 === 0 ? "alt" : ""}">
            <td>${q.seq}</td>
            <td class="pdf-td-right">${esc(q.text)}</td>
            ${is5 ? `
              <td>${fmtPct(q.pcts["1"])}</td>
              <td>${fmtPct(q.pcts["2"])}</td>
              <td>${fmtPct(q.pcts["3"])}</td>
              <td>${fmtPct(q.pcts["4"])}</td>
              <td>${fmtPct(q.pcts["5"])}</td>
            ` : `
              <td>${fmtPct(q.pcts["disagree"])}</td>
              <td>${fmtPct(q.pcts["neutral"])}</td>
              <td>${fmtPct(q.pcts["agree"])}</td>
            `}
          </tr>`).join("")}
        ${showTotal ? `<tr class="pdf-total-row">
          <td colspan="${is5 ? 6 : 4}" class="pdf-td-right">إجمالي درجات المحور</td>
          <td>${ax.axisAgreePct}%</td>
        </tr>` : ""}
      </tbody>
    </table>
    ${includeCharts ? `
    <div class="pdf-hbars">
      ${ax.questions.map((q, i) => horizontalBar(q.text, q.agreePct ?? 0, i)).join("")}
    </div>` : ""}
    <div class="pdf-axis-spacer"></div>`;
}

// Estimates each axis block's rendered height in CSS px (header row, one row
// per question, the total row, and — when charts are shown — one hbar row
// per question) so several short axes can be packed onto one physical page
// instead of always burning a full page per axis. Measured against a real,
// attached-but-hidden `.pdf-page` probe rather than estimated — a fixed
// per-row px guess undercounts headers like "لا أوافق بشدة" that wrap onto
// two lines in a narrow column, which caused pages to overflow anyway.
async function measureBlockHeightsPx(htmlBlocks, shouldCancel) {
  const probe = document.createElement("div");
  probe.className = "pdf-page";
  probe.style.cssText = "position:absolute;visibility:hidden;left:-9999px;top:0;min-height:0;padding:0;border:0;";
  document.body.appendChild(probe);
  const heights = [];
  try {
    for (let i = 0; i < htmlBlocks.length; i++) {
      throwIfCancelled(shouldCancel);
      probe.innerHTML = htmlBlocks[i];
      heights.push(probe.getBoundingClientRect().height);
      if (i % 3 === 2) await yieldToBrowser();
    }
    return heights;
  } finally {
    probe.remove();
  }
}

// Available content height per page: .pdf-page min-height (1080) minus the
// header block (~110) and top/bottom padding (~68), with a safety margin.
const AXIS_PAGE_BUDGET_PX = 850;

async function buildAxisDetailPages(result, s, logoSrc, includeCharts, shouldCancel, firstPagePrefix = "") {
  const maxQuestionsPerBlock = includeCharts ? 5 : 12;
  const blocks = result.axes.flatMap(ax => {
    const questionGroups = chunk(ax.questions, maxQuestionsPerBlock);
    return questionGroups.map((questions, index) => buildAxisBlock(
      { ...ax, questions }, result, includeCharts,
      { continued: index > 0, showTotal: index === questionGroups.length - 1 }
    ));
  });
  const heights = await measureBlockHeightsPx(blocks, shouldCancel);
  const prefixHeight = firstPagePrefix
    ? (await measureBlockHeightsPx([firstPagePrefix], shouldCancel))[0]
    : 0;

  const groups = [];
  let current = [];
  let currentHeight = prefixHeight;

  blocks.forEach((blockHtml, i) => {
    const h = heights[i];
    if ((current.length > 0 || (groups.length === 0 && firstPagePrefix)) && currentHeight + h > AXIS_PAGE_BUDGET_PX) {
      groups.push(current);
      current = [];
      currentHeight = 0;
    }
    current.push(blockHtml);
    currentHeight += h;
  });
  if (current.length > 0) groups.push(current);

  return groups.map((group, groupIndex) => `
    ${pageOpen()}
      ${pageHeader(s, logoSrc)}
      ${groupIndex === 0 ? firstPagePrefix : ""}
      ${group.length > 0 && (groupIndex === 0 || (groupIndex === 1 && firstPagePrefix && groups[0].length === 0)) && (result.isFlat || result.reportTexts?.resultsHeading?.trim()) ? sectionHeading(reportText(result, "resultsHeading", "عرض النتائج وتحليلها")) : ""}
      ${group.join("")}
    ${pageClose(s)}`).join("");
}

function buildSignaturesBlock(s) {
  return `
    <div class="pdf-sig-row" style="margin-top:32px">
      <div class="pdf-sig-box">
        <div class="pdf-sig-name">${esc(s.quName)}</div>
        <div class="pdf-sig-role">رئيس وحدة الجودة</div>
        <div class="pdf-sig-line">التوقيع: ....................</div>
      </div>
      <div class="pdf-sig-box">
        <div class="pdf-sig-name">${esc(s.qmName)}</div>
        <div class="pdf-sig-role">رئيس لجنة القياس والتقويم</div>
        <div class="pdf-sig-line">التوقيع: ....................</div>
      </div>
    </div>`;
}

// Any individual statement (not just whole axes) that scores below
// s.recommendationsThreshold (default 70%) is listed as a row — axis /
// statement / percentage — so a low-scoring axis with one strong statement
// doesn't hide the specific statement that actually needs attention.
function buildRecommendationsPages(result, s, logoSrc) {
  const threshold = s.recommendationsThreshold ?? 70;
  const rows = [];
  result.axes.forEach(ax => {
    ax.questions.forEach(q => {
      const pct = q.agreePct ?? 0;
      if (pct < threshold) rows.push({ axisName: ax.name, text: q.text, pct });
    });
  });
  rows.sort((a, b) => a.pct - b.pct);

  if (rows.length === 0) {
    return `
      ${pageOpen()}
        ${pageHeader(s, logoSrc)}
        ${sectionHeading(reportText(result, "recommendationsHeading", "التوصيات"))}
        <div class="pdf-no-recs">${esc(reportText(result, "noRecommendationsText", "لا توجد توصيات"))}</div>
        ${buildSignaturesBlock(s)}
      ${pageClose(s)}`;
  }

  const pages = balancedChunks(rows, RECS_ROWS_PER_PAGE);
  return pages.map((pageRows, pageIdx) => {
    const isLast = pageIdx === pages.length - 1;
    return `
      ${pageOpen()}
        ${pageHeader(s, logoSrc)}
        ${sectionHeading(pageIdx === 0 ? `${reportText(result, "recommendationsHeading", "التوصيات")} (أقل من ${threshold}%)` : `${reportText(result, "recommendationsHeading", "التوصيات")} (تابع)`)}
        <table class="pdf-table">
          <colgroup><col style="width:13%"><col style="width:61%"><col style="width:26%"></colgroup>
          <thead>
            <tr><th>النسبة</th><th class="pdf-th-wide">العبارة</th><th>المحور</th></tr>
          </thead>
          <tbody>
            ${pageRows.map((r, i) => `
              <tr class="${i % 2 === 0 ? "alt" : ""}">
                <td class="pdf-td-strong">${r.pct}%</td>
                <td class="pdf-td-right">${esc(r.text)}</td>
                <td class="pdf-td-right">${esc(r.axisName)}</td>
              </tr>`).join("")}
          </tbody>
        </table>
        ${isLast ? buildSignaturesBlock(s) : ""}
      ${pageClose(s)}`;
  }).join("");
}

function buildComparisonPages(comparison, meta, s, logoSrc) {
  const firstResult = comparison.slots?.find(slot => slot?.result)?.result ?? {};
  const years = comparison.overall.map(item => item.year);
  const title = `مقارنة نتائج ${years.length} ${years.length === 2 ? "عامين" : "أعوام"}`;
  const rows = [
    ["نوع الاستبيان", firstResult.schemaLabel || "—"],
    ...(meta.program ? [["البرنامج / القسم", meta.program]] : []),
    ["الأعوام محل المقارنة", years.join(" / ") || "—"],
    ["عدد المحاور", comparison.axes.length],
    ["موجه إلى", "مدير وحدة الجودة بالكلية"],
  ];
  const signers = s.includeEvaluatorsTable === false ? [] : [
    ...(s.includeCommitteeHead !== false ? [{ name: s.qmName || "", role: "رئيس لجنة القياس والتقويم" }] : []),
    ...(s.includeEvaluator !== false ? [{ name: meta.preparedBy || "", role: "القائم بالتقييم" }] : []),
    ...(s.includeReviewer !== false ? [{ name: meta.reviewer || "", role: "القائم بالمراجعة" }] : []),
  ];
  const cover = `
    ${pageOpen("pdf-cover")}
      ${pageHeader(s, logoSrc)}
      <div class="pdf-cover-title-wrap">
        <div class="pdf-cover-title">${esc(title)}</div>
        <div class="pdf-cover-subtitle">تقرير مقارنة نتائج الاستبيانات</div>
        <div class="pdf-cover-committee">${esc(s.committeeName)}</div>
      </div>
      ${sectionHeading("البيانات الأساسية")}
      <div class="pdf-info-table">${rows.map(([l, v], i) => infoRow(l, v, i)).join("")}</div>
      ${signers.length ? `
        ${sectionHeading("القائم بالتقييم")}
        <div class="pdf-sig-row">
          ${signers.map(({ name, role }) => `
            <div class="pdf-sig-box">
              <div class="pdf-sig-name">${esc(name || "—")}</div>
              <div class="pdf-sig-role">${esc(role)}</div>
              <div class="pdf-sig-line">التوقيع: ....................</div>
            </div>`).join("")}
        </div>` : ""}
    ${pageClose(s)}`;

  const axesPerPage = years.length === 3 ? 13 : 14;
  const axisChunks = chunk(comparison.axes, axesPerPage);
  const tablePages = axisChunks.map((pageAxes, pageIdx) => {
    const isLast = pageIdx === axisChunks.length - 1;
    return `
      ${pageOpen()}
        ${pageHeader(s, logoSrc)}
        ${sectionHeading(pageIdx === 0 ? title : `${title} (تابع)`)}
        <table class="pdf-table pdf-comparison-table">
          <colgroup>
            <col style="width:5%"><col style="width:${years.length === 3 ? 43 : 49}%">
            ${years.map(() => `<col style="width:${years.length === 3 ? 13 : 17}%">`).join("")}
            <col style="width:${years.length === 3 ? 13 : 12}%">
          </colgroup>
          <thead><tr><th>م</th><th class="pdf-th-wide">المحور</th>${years.map(y => `<th>${esc(y)}</th>`).join("")}<th>الاتجاه</th></tr></thead>
          <tbody>
            ${pageAxes.map((axis, index) => `
              <tr class="${index % 2 === 0 ? "alt" : ""}">
                <td>${pageIdx * axesPerPage + index + 1}</td>
                <td class="pdf-td-right">${esc(axis.name)}</td>
                ${axis.years.map(year => `<td class="pdf-td-strong">${year.agreePct == null ? "—" : `${year.agreePct}%`}</td>`).join("")}
                <td class="pdf-trend ${axis.trend === "تحسن" ? "up" : axis.trend === "تراجع" ? "down" : "steady"}">${esc(axis.trend || "—")}</td>
              </tr>`).join("")}
            ${isLast ? `
              <tr class="pdf-total-row"><td></td><td class="pdf-td-right">الإجمالي</td>
                ${comparison.overall.map(item => `<td>${item.agreePct == null ? "—" : `${item.agreePct}%`}</td>`).join("")}<td></td>
              </tr>` : ""}
          </tbody>
        </table>
        ${isLast ? `<div class="pdf-comparison-legend"><span class="up">تحسن</span><span class="down">تراجع</span><span class="steady">استقرار</span></div>` : ""}
      ${pageClose(s)}`;
  }).join("");

  return cover + tablePages;
}

function examPaperSummary(report) {
  const completed = report.courses.filter(course => course.status === "complete");
  const incomplete = report.courses.filter(course => course.status === "incomplete");
  const total = report.courses.length;
  const percentage = count => total ? Math.round(count * 1000 / total) / 10 : 0;
  return { completed, incomplete, total, completedPct: percentage(completed.length), incompletePct: percentage(incomplete.length) };
}

function buildExamPaperCover(report, summary, s, logoSrc) {
  const rows = [
    ["القسم", report.department],
    ["الفصل الدراسي", report.semester],
    ["العام الدراسي", report.year],
    ["عدد المقررات التي تم تقييمها", summary.total],
    ["الجهة القائمة بالتقييم", s.committeeName || "لجنة القياس والتقويم"],
  ];
  return `
    ${pageOpen("pdf-cover pdf-exam-cover")}
      ${pageHeader(s, logoSrc)}
      <div class="pdf-cover-title-wrap">
        <div class="pdf-cover-title">تقييم الورقة الامتحانية من حيث الشكل</div>
        <div class="pdf-cover-subtitle">تقرير مراجعة استيفاء عناصر الورقة الامتحانية</div>
        <div class="pdf-cover-committee">${esc(report.department)}</div>
      </div>
      ${sectionHeading("البيانات الأساسية")}
      <div class="pdf-info-table">${rows.map(([label, value], index) => infoRow(label, value, index)).join("")}</div>
      ${sectionHeading("ملخص حالة المقررات")}
      <div class="pdf-exam-stats">
        <div><b>${summary.total}</b><span>إجمالي المقررات</span></div>
        <div class="complete"><b>${summary.completed.length}</b><span>مستوفاة لجميع العناصر</span></div>
        <div class="incomplete"><b>${summary.incomplete.length}</b><span>غير مستوفاة لبعض العناصر</span></div>
      </div>
    ${pageClose(s)}`;
}

function buildExamPaperSummaryPage(report, summary, s, logoSrc) {
  return `
    ${pageOpen("pdf-exam-summary")}
      ${pageHeader(s, logoSrc)}
      ${sectionHeading("أولًا: الملخص التنفيذي لتقييم أوراق القسم")}
      <p class="pdf-body-text">تم تقييم أوراق ${summary.total} مقررًا بقسم ${esc(report.department)} خلال الفصل الدراسي ${esc(report.semester)} للعام ${esc(report.year)}.</p>
      <table class="pdf-table pdf-exam-summary-table">
        <thead><tr><th class="pdf-th-wide">حالة المقررات</th><th>العدد</th><th>النسبة</th></tr></thead>
        <tbody>
          <tr class="alt"><td class="pdf-td-right">المقررات المستوفاة لجميع عناصر التقييم</td><td class="pdf-td-strong">${summary.completed.length}</td><td class="pdf-td-strong">${summary.completedPct}%</td></tr>
          <tr><td class="pdf-td-right">المقررات غير المستوفاة لبعض عناصر التقييم</td><td class="pdf-td-strong">${summary.incomplete.length}</td><td class="pdf-td-strong">${summary.incompletePct}%</td></tr>
          <tr class="pdf-total-row"><td class="pdf-td-right">الإجمالي</td><td>${summary.total}</td><td>100%</td></tr>
        </tbody>
      </table>
      ${sectionHeading("التوزيع النسبي")}
      <div class="pdf-exam-bars">
        <div><span>المقررات المستوفاة</span><b>${summary.completedPct}%</b><i><u style="width:${summary.completedPct}%"></u></i></div>
        <div class="incomplete"><span>المقررات غير المستوفاة</span><b>${summary.incompletePct}%</b><i><u style="width:${summary.incompletePct}%"></u></i></div>
      </div>
    ${pageClose(s)}`;
}

function buildExamPaperCoursePages(rows, kind, s, logoSrc) {
  const completed = kind === "completed";
  const pageSize = completed ? 14 : 9;
  const pages = chunk(rows, pageSize);
  if (!pages.length) pages.push([]);
  const heading = completed
    ? "ثانيًا: المقررات المستوفاة لجميع عناصر التقييم"
    : "ثالثًا: المقررات غير المستوفاة لبعض عناصر التقييم";
  return pages.map((pageRows, pageIndex) => `
    ${pageOpen("pdf-exam-courses")}
      ${pageHeader(s, logoSrc)}
      ${sectionHeading(`${heading}${pages.length > 1 ? ` - صفحة ${pageIndex + 1} من ${pages.length}` : ""}`)}
      <table class="pdf-table pdf-table-detail">
        <colgroup>${completed
          ? '<col style="width:8%"><col style="width:20%"><col style="width:48%"><col style="width:24%">'
          : '<col style="width:7%"><col style="width:16%"><col style="width:27%"><col style="width:32%"><col style="width:18%">'}</colgroup>
        <thead><tr><th>م</th><th>كود المقرر</th><th class="pdf-th-wide">اسم المقرر</th>${completed ? "" : '<th class="pdf-th-wide">العناصر غير المستوفاة</th>'}<th>مخرجات التعلم</th></tr></thead>
        <tbody>${pageRows.length ? pageRows.map((course, index) => `
          <tr class="${index % 2 === 0 ? "alt" : ""}">
            <td>${pageIndex * pageSize + index + 1}</td><td>${esc(course.code)}</td><td class="pdf-td-right">${esc(course.name)}</td>
            ${completed ? "" : `<td class="pdf-td-right">${esc(course.missingItems || "—")}</td>`}
            <td class="pdf-td-strong">${course.learningOutcomes === "" || course.learningOutcomes == null ? "—" : `${esc(course.learningOutcomes)}%`}</td>
          </tr>`).join("") : `<tr><td colspan="${completed ? 4 : 5}" class="pdf-exam-empty">لا توجد مقررات ضمن هذه الفئة</td></tr>`}</tbody>
      </table>
    ${pageClose(s)}`).join("");
}

function buildExamPaperEvaluatorsPage(report, s, logoSrc) {
  const evaluators = report.evaluators?.length ? report.evaluators : [{ name: "—", role: "—" }];
  return `
    ${pageOpen("pdf-exam-evaluators")}
      ${pageHeader(s, logoSrc)}
      ${sectionHeading("القائمون بالتقييم والاعتماد")}
      <table class="pdf-table">
        <thead><tr><th style="width:10%">م</th><th class="pdf-th-wide">الاسم</th><th class="pdf-th-wide">الوظيفة</th></tr></thead>
        <tbody>${evaluators.map((evaluator, index) => `<tr class="${index % 2 === 0 ? "alt" : ""}"><td>${index + 1}</td><td class="pdf-td-right pdf-td-strong">${esc(evaluator.name)}</td><td class="pdf-td-right">${esc(evaluator.role)}</td></tr>`).join("")}</tbody>
      </table>
      <div class="pdf-exam-approval">
        <span>رئيس لجنة القياس والتقويم</span>
        <b>${esc(report.committeeHead || evaluators[0]?.name || "—")}</b>
        <i>التوقيع: ................................................</i>
      </div>
    ${pageClose(s)}`;
}

function buildExamPaperPages(report, s, logoSrc) {
  const summary = examPaperSummary(report);
  return [
    buildExamPaperCover(report, summary, s, logoSrc),
    buildExamPaperSummaryPage(report, summary, s, logoSrc),
    buildExamPaperCoursePages(summary.completed, "completed", s, logoSrc),
    buildExamPaperCoursePages(summary.incomplete, "incomplete", s, logoSrc),
    buildExamPaperEvaluatorsPage(report, s, logoSrc),
  ].join("");
}

// ── CSS ─────────────────────────────────────────────────────────────────────────
const PDF_CSS = `
  * { box-sizing: border-box; margin: 0; padding: 0; }
  .pdf-root { direction: rtl; font-family: 'Cairo', Arial, sans-serif; color: ${TEXT_DARK}; }
  .pdf-page-outer {
    width: 808px; background: ${PINK_SOFT}; margin: 0 auto 0; padding: 0 14px;
  }
  .pdf-page {
    width: 780px; min-height: 1080px; background: #fff; margin: 0 auto;
    padding: 34px 44px 76px; position: relative;
    border-left: 5px solid ${PINK_LINE}; border-right: 5px solid ${PINK_LINE};
  }
  .pdf-header { position: relative; display: flex; align-items: center; justify-content: center; gap: 14px; margin-bottom: 22px; padding-bottom: 14px; border-bottom: 2px solid ${PINK}; }
  .pdf-logo { width: 52px; height: 52px; object-fit: contain; }
  .pdf-header-text { text-align: center; }
  .pdf-uni-name { font-size: 19px; font-weight: 900; color: ${TEXT_DARK}; }
  .pdf-faculty-name { font-size: 12px; color: ${TEXT_MUTED}; margin-top: 2px; }
  .pdf-page-footer { position: absolute; right: 44px; bottom: 20px; left: 44px; display: flex; justify-content: center; gap: 22px; padding-top: 8px; border-top: 1px solid ${PINK_LINE}; color: ${TEXT_MUTED}; font-size: 10.5px; direction: ltr; }
  .pdf-capture-top-marker, .pdf-capture-bottom-marker { display: block; flex: 0 0 4px; width: 4px; height: 4px; }
  .pdf-capture-top-marker { margin: 0 auto; background: rgb(0, 255, 255); }
  .pdf-capture-bottom-marker { align-self: center; background: rgb(255, 0, 255); }

  .pdf-cover-title-wrap { text-align: center; margin: 30px 0 26px; }
  .pdf-cover-title { display: inline-block; font-size: 26px; font-weight: 900; padding: 6px 22px; background: ${PINK}; border-radius: 6px; color: ${TEXT_DARK}; }
  .pdf-cover-intro { margin: -8px auto 22px; max-width: 620px; color: ${TEXT_MUTED}; font-size: 12px; line-height: 1.9; text-align: center; }
  .pdf-cover-subtitle { font-size: 15px; color: ${TEXT_MUTED}; margin-top: 12px; }
  .pdf-cover-committee { font-size: 17px; font-weight: 800; margin-top: 6px; color: ${ACCENT}; }

  .pdf-section-heading { display: flex; align-items: center; gap: 8px; font-size: 15px; font-weight: 900; color: ${TEXT_DARK}; background: ${PINK}; padding: 8px 14px; border-radius: 6px; margin: 22px 0 12px; }

  .pdf-info-table { border: 1px solid #eee; border-radius: 8px; overflow: hidden; }
  .pdf-info-row { display: flex; justify-content: space-between; padding: 9px 16px; font-size: 13px; }
  .pdf-info-row.alt { background: ${PINK_SOFT}; }
  .pdf-info-label { color: ${TEXT_MUTED}; font-weight: 700; }
  .pdf-info-val { font-weight: 800; color: ${TEXT_DARK}; }

  .pdf-sig-row { display: flex; gap: 18px; margin-top: 8px; }
  .pdf-sig-box { flex: 1; text-align: center; padding: 14px 10px; border: 1px dashed #ddd; border-radius: 8px; }
  .pdf-sig-name { font-weight: 900; font-size: 14px; }
  .pdf-sig-role { font-size: 12px; color: ${TEXT_MUTED}; margin: 4px 0; }
  .pdf-sig-line { font-size: 11px; color: #999; margin-top: 10px; }

  .pdf-body-text { font-size: 13.5px; line-height: 1.9; color: ${TEXT_DARK}; margin-bottom: 6px; text-align: justify; }
  .pdf-vm-block { margin-bottom: 22px; }
  .pdf-vm-title { font-weight: 900; font-size: 14px; color: ${ACCENT}; margin-bottom: 6px; }
  .pdf-vm-text { font-size: 13px; line-height: 1.9; text-align: justify; }
  .pdf-context-page { direction: rtl; text-align: right; }
  .pdf-context-page .pdf-section-heading { margin-top: 14px; margin-bottom: 8px; direction: rtl; text-align: right; justify-content: flex-start; }
  .pdf-context-page .pdf-vm-block { margin-bottom: 10px; }
  .pdf-context-page .pdf-vm-title { direction: rtl; text-align: right; }
  .pdf-context-page .pdf-vm-text { font-size: 12px; line-height: 1.65; direction: rtl; text-align: justify; }
  .pdf-context-page .pdf-body-text { font-size: 12px; line-height: 1.7; direction: rtl; text-align: justify; }
  .pdf-context-page .pdf-procedure-list { margin-top: 7px; font-size: 12px; line-height: 1.7; }
  .pdf-context-page .pdf-procedure-list li { margin-bottom: 4px; }
  .pdf-formula-block { background: ${PINK_SOFT}; border-radius: 8px; padding: 12px 16px; margin-bottom: 14px; }
  .pdf-formula-title { font-weight: 800; font-size: 13px; margin-bottom: 6px; }
  .pdf-formula { font-size: 16px; font-weight: 700; direction: ltr; text-align: center; margin: 6px 0; }
  .pdf-formula-note { font-size: 11px; color: ${TEXT_MUTED}; text-align: center; }

  .pdf-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  .pdf-table { table-layout: fixed; border: 1px solid #d8d8d8; }
  .pdf-table th { background: ${ACCENT}; color: #fff; padding: 9px 7px; font-weight: 800; border: 1px solid rgba(255,255,255,.28); }
  .pdf-table td { padding: 8px 7px; text-align: center; border: 1px solid #e1e1e1; }
  .pdf-table tr { break-inside: avoid; page-break-inside: avoid; }
  .pdf-table tr.alt td { background: ${PINK_SOFT}; }
  .pdf-table .pdf-th-wide { text-align: right; padding-right: 12px; }
  .pdf-table .pdf-td-right { text-align: right; padding-right: 12px; }
  .pdf-table .pdf-td-strong { font-weight: 800; }
  .pdf-total-row td { background: ${PINK}; font-weight: 900; }
  .pdf-table-detail td { font-size: 11.5px; }
  .pdf-comparison-table td { font-size: 12px; }
  .pdf-trend { font-weight: 900; }
  .pdf-trend.up, .pdf-comparison-legend .up { color: #168447; }
  .pdf-trend.down, .pdf-comparison-legend .down { color: #b4232d; }
  .pdf-trend.steady, .pdf-comparison-legend .steady { color: #8a6116; }
  .pdf-comparison-legend { display: flex; justify-content: center; gap: 26px; margin-top: 18px; font-size: 12px; font-weight: 900; }
  .pdf-compact-summary { margin-bottom: 14px; }
  .pdf-summary-chart .pdf-section-heading { margin-top: 18px; }

  .pdf-axes-chart { display: grid; grid-template-columns: 1fr; gap: 10px; margin-top: 16px; direction: rtl; }
  .pdf-axes-chart.two-columns { grid-template-columns: repeat(2, minmax(0, 1fr)); column-gap: 22px; row-gap: 8px; }
  .pdf-axis-chart-row { min-width: 0; break-inside: avoid; }
  .pdf-axis-chart-head { display: flex; align-items: flex-end; justify-content: space-between; gap: 10px; min-height: 27px; margin-bottom: 4px; }
  .pdf-axis-chart-label { min-width: 0; font-size: 10px; color: ${TEXT_DARK}; line-height: 1.35; font-weight: 700; }
  .pdf-axis-chart-pct { flex: 0 0 auto; font-size: 10.5px; font-weight: 900; color: ${ACCENT}; }
  .pdf-axis-chart-track { height: 11px; overflow: hidden; border-radius: 3px; background: #f0f1f3; direction: rtl; }
  .pdf-axis-chart-fill { height: 100%; border-radius: 3px; }

  .pdf-hbars { margin-top: 18px; display: flex; flex-direction: column; gap: 10px; }
  .pdf-hbar-row { display: flex; align-items: center; gap: 10px; }
  .pdf-hbar-label { flex: 0 0 220px; font-size: 11px; text-align: right; line-height: 1.4; }
  .pdf-hbar-track { flex: 1; background: #f4f4f4; border-radius: 10px; height: 20px; overflow: hidden; }
  .pdf-hbar-fill { height: 100%; border-radius: 10px; display: flex; align-items: center; justify-content: center; }
  .pdf-hbar-pct { font-size: 10.5px; color: #fff; font-weight: 800; padding: 0 6px; }

  .pdf-axis-spacer { height: 22px; }

  .pdf-no-recs { text-align: center; font-weight: 800; color: ${ACCENT}; font-size: 15px; margin-top: 30px; }
  .pdf-exam-stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 8px; }
  .pdf-exam-stats > div { padding: 18px 10px; text-align: center; border: 1px solid #dde4ea; border-radius: 9px; background: #f8fafc; }
  .pdf-exam-stats b, .pdf-exam-stats span { display: block; }
  .pdf-exam-stats b { color: ${ACCENT}; font-size: 25px; }
  .pdf-exam-stats span { margin-top: 5px; color: ${TEXT_MUTED}; font-size: 10.5px; line-height: 1.5; }
  .pdf-exam-stats .complete { background: #eef9f2; border-color: #b9e1c6; }
  .pdf-exam-stats .complete b { color: #168447; }
  .pdf-exam-stats .incomplete { background: #fff8e8; border-color: #eed69d; }
  .pdf-exam-stats .incomplete b { color: #a66a09; }
  .pdf-exam-summary-table { margin-bottom: 28px; }
  .pdf-exam-bars { display: grid; gap: 20px; margin-top: 20px; }
  .pdf-exam-bars > div { display: grid; grid-template-columns: 1fr 52px; gap: 8px 12px; align-items: center; }
  .pdf-exam-bars span { font-size: 12.5px; font-weight: 800; }
  .pdf-exam-bars b { direction: ltr; color: #168447; text-align: left; font-size: 13px; }
  .pdf-exam-bars i { grid-column: 1 / -1; height: 18px; overflow: hidden; border-radius: 9px; background: #edf1f4; }
  .pdf-exam-bars u { display: block; height: 100%; border-radius: inherit; background: #27ae60; text-decoration: none; }
  .pdf-exam-bars .incomplete b { color: #b7790a; }
  .pdf-exam-bars .incomplete u { background: #f0a725; }
  .pdf-exam-empty { padding: 30px !important; color: ${TEXT_MUTED}; font-weight: 800; }
  .pdf-exam-approval { width: 58%; margin: 80px auto 0; padding: 22px; text-align: center; border: 1px dashed ${PINK_LINE}; border-radius: 10px; }
  .pdf-exam-approval span, .pdf-exam-approval b, .pdf-exam-approval i { display: block; }
  .pdf-exam-approval span { color: ${TEXT_MUTED}; font-size: 12px; }
  .pdf-exam-approval b { margin-top: 9px; color: ${ACCENT}; font-size: 16px; }
  .pdf-exam-approval i { margin-top: 28px; color: #888; font-size: 11px; font-style: normal; }
  .pdf-procedure-list { direction: rtl; list-style: none; counter-reset: procedure-item; margin: 18px 0 0; padding: 0; color: ${TEXT_DARK}; font-size: 13px; line-height: 2.05; text-align: right; }
  .pdf-procedure-list li { direction: rtl; display: flex; flex-direction: row; align-items: flex-start; gap: 8px; counter-increment: procedure-item; margin-bottom: 8px; text-align: right; }
  .pdf-procedure-list li::before { content: counter(procedure-item) "."; flex: 0 0 22px; direction: ltr; text-align: right; font-weight: 700; color: ${ACCENT}; }
`;

function themedPdfCss(themeId) {
  const t = PDF_THEMES[themeId] ?? PDF_THEMES.default;
  return PDF_CSS
    .replaceAll(ACCENT, t.accent)
    .replaceAll(PINK_SOFT, t.soft)
    .replaceAll(PINK_LINE, t.line)
    .replaceAll(PINK, t.heading);
}

function cancellationError() {
  const error = new Error("تم إلغاء إنشاء ملف PDF");
  error.name = "AbortError";
  return error;
}

function throwIfCancelled(shouldCancel) {
  if (shouldCancel?.()) throw cancellationError();
}

// html2canvas and jsPDF both do substantial work on the main thread. Yielding
// between pages/slices lets React paint progress and lets cancel clicks run.
function yieldToBrowser() {
  return new Promise(resolve => setTimeout(resolve, 0));
}

function waitForLayout() {
  return new Promise(resolve => {
    let finished = false;
    const done = () => {
      if (finished) return;
      finished = true;
      resolve();
    };
    setTimeout(done, 120);
    requestAnimationFrame(() => requestAnimationFrame(done));
  });
}

function consumeCaptureMarkers(canvas) {
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const image = ctx.getImageData(0, 0, canvas.width, canvas.height);
  let topMarker = null;
  let bottomMarker = null;
  const topLimit = canvas.height * 0.08;
  const bottomLimit = canvas.height * 0.92;
  for (let pixel = 0; pixel < image.data.length; pixel += 4) {
    const index = pixel / 4;
    const x = index % canvas.width;
    const y = Math.floor(index / canvas.width);
    const red = image.data[pixel];
    const green = image.data[pixel + 1];
    const blue = image.data[pixel + 2];
    if (!topMarker && y < topLimit && red < 20 && green > 245 && blue > 245) {
      topMarker = { x, y };
    } else if (!bottomMarker && y > bottomLimit && red > 245 && green < 20 && blue > 245) {
      bottomMarker = { x, y };
    }
    if (topMarker && bottomMarker) break;
  }
  if (!topMarker || !bottomMarker) return false;
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(topMarker.x - 2, topMarker.y - 2, 10, 10);
  ctx.fillRect(bottomMarker.x - 2, bottomMarker.y - 2, 10, 10);
  ctx.restore();
  return true;
}

function hasCompletePageFrame(canvas) {
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const image = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
  const band = Math.max(2, Math.floor(canvas.width * 0.05));
  let sampledRows = 0;
  let framedRows = 0;
  const isFramePixel = offset => image[offset] < 248 || image[offset + 1] < 248 || image[offset + 2] < 248;
  for (let y = 0; y < canvas.height; y += 4) {
    sampledRows++;
    let left = false;
    let right = false;
    for (let x = 0; x < band && (!left || !right); x++) {
      const leftOffset = (y * canvas.width + x) * 4;
      const rightOffset = (y * canvas.width + (canvas.width - 1 - x)) * 4;
      if (!left && isFramePixel(leftOffset)) left = true;
      if (!right && isFramePixel(rightOffset)) right = true;
    }
    if (left && right) framedRows++;
  }
  return framedRows / Math.max(1, sampledRows) > 0.85;
}

async function renderRootToPdf({ root, container, settings, onProgress, shouldCancel, filename, html2canvas, jsPDF }) {
  const savedScrollX = window.scrollX;
  const savedScrollY = window.scrollY;
  try {
    const pageEls = Array.from(root.querySelectorAll(".pdf-page-outer"));
    root.remove();
    // html2canvas's first clone can inherit the live page scroll offset even
    // when scrollY is supplied in its options. Normalise the real viewport
    // once before capturing so page 1 cannot start below its header.
    window.scrollTo(0, 0);
    await waitForLayout();
    const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
    const pageWidthMm = pdf.internal.pageSize.getWidth();
    const pageHeightMm = pdf.internal.pageSize.getHeight();
    let firstImage = true;
    let pageIndex = 0;

    for (const el of pageEls) {
      throwIfCancelled(shouldCancel);
      pageIndex++;
      onProgress?.(pageIndex, pageEls.length);
      const captureRoot = document.createElement("div");
      captureRoot.className = "pdf-root";
      const captureEl = el.cloneNode(true);
      captureRoot.appendChild(captureEl);
      container.appendChild(captureRoot);
      captureRoot.style.width = `${captureEl.offsetWidth}px`;
      captureRoot.style.height = `${captureEl.offsetHeight}px`;
      captureRoot.style.overflow = "hidden";
      await waitForLayout();
      await new Promise(resolve => setTimeout(resolve, 50));
      throwIfCancelled(shouldCancel);
      let canvas = null;
      for (let attempt = 0; attempt < 3; attempt++) {
        window.scrollTo(0, 0);
        await waitForLayout();
        canvas = await html2canvas(captureRoot, {
          scale: Math.min(2, Math.max(1, Number(settings.pdfRenderScale) || 1.5)),
          useCORS: true, logging: false, backgroundColor: "#ffffff",
          x: 0, y: 0, width: captureRoot.offsetWidth, height: captureRoot.offsetHeight,
          scrollX: 0, scrollY: 0, imageTimeout: 4000,
          onclone: clonedDocument => {
            clonedDocument.documentElement.scrollTop = 0;
            clonedDocument.body.scrollTop = 0;
          },
        });
        if (consumeCaptureMarkers(canvas) && hasCompletePageFrame(canvas)) break;
        canvas.width = 1;
        canvas.height = 1;
        canvas = null;
        await waitForLayout();
        await new Promise(resolve => setTimeout(resolve, 100));
      }
      if (!canvas) {
        throw new Error(`تعذّر التقاط الصفحة ${pageIndex} كاملة بالترويسة والتذييل`);
      }
      captureRoot.remove();
      await yieldToBrowser();
      throwIfCancelled(shouldCancel);
      const fullHeightMm = (canvas.height * pageWidthMm) / canvas.width;
      const sliceCount = Math.max(1, Math.ceil(fullHeightMm / pageHeightMm));
      const sliceHeightPx = Math.ceil(canvas.height / sliceCount);

      for (let sliceIndex = 0; sliceIndex < sliceCount; sliceIndex++) {
        throwIfCancelled(shouldCancel);
        await yieldToBrowser();
        const sliceCanvas = document.createElement("canvas");
        sliceCanvas.width = canvas.width;
        sliceCanvas.height = sliceHeightPx;
        const ctx = sliceCanvas.getContext("2d");
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, sliceCanvas.width, sliceCanvas.height);
        ctx.drawImage(
          canvas, 0, sliceIndex * sliceHeightPx, canvas.width, sliceHeightPx,
          0, 0, canvas.width, sliceHeightPx
        );
        const sliceImgData = sliceCanvas.toDataURL("image/jpeg", 0.9);
        const sliceHeightMm = (sliceHeightPx * pageWidthMm) / canvas.width;
        if (!firstImage) pdf.addPage();
        firstImage = false;
        pdf.addImage(
          sliceImgData, "JPEG", 0, 0, pageWidthMm, sliceHeightMm,
          `pdf-page-${pageIndex}-${sliceIndex}`, "FAST"
        );
        sliceCanvas.width = 1;
        sliceCanvas.height = 1;
      }
      canvas.width = 1;
      canvas.height = 1;
    }

    await yieldToBrowser();
    throwIfCancelled(shouldCancel);
    return { blob: pdf.output("blob"), filename };
  } finally {
    container.remove();
    window.scrollTo(savedScrollX, savedScrollY);
  }
}

// Comparison reports are short enough to capture as one canvas. Cropping that
// canvas at the measured page boundaries avoids Chromium shifting alternate
// per-page captures while keeping the same page design as regular reports.
async function renderComparisonRootToPdf({ root, container, settings, onProgress, shouldCancel, filename, html2canvas, jsPDF }) {
  const savedScrollX = window.scrollX;
  const savedScrollY = window.scrollY;
  try {
    window.scrollTo(0, 0);
    await waitForLayout();
    const pageEls = Array.from(root.querySelectorAll(".pdf-page-outer"));
    if (!pageEls.length) throw new Error("لا توجد صفحات مقارنة قابلة للتصدير");
    root.style.width = `${pageEls[0].offsetWidth}px`;
    await waitForLayout();
    const pages = pageEls.map(el => ({ top: el.offsetTop, height: el.offsetHeight }));
    const scale = Math.min(2, Math.max(1, Number(settings.pdfRenderScale) || 1.5));
    const canvas = await html2canvas(root, {
      scale, useCORS: true, logging: false, backgroundColor: "#ffffff",
      x: 0, y: 0, width: root.offsetWidth, height: root.scrollHeight,
      scrollX: 0, scrollY: 0, imageTimeout: 4000,
      onclone: clonedDocument => {
        clonedDocument.documentElement.scrollTop = 0;
        clonedDocument.body.scrollTop = 0;
        clonedDocument.querySelectorAll(".pdf-capture-top-marker,.pdf-capture-bottom-marker")
          .forEach(marker => { marker.style.visibility = "hidden"; });
      },
    });
    const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
    const pageWidthMm = pdf.internal.pageSize.getWidth();
    for (let index = 0; index < pages.length; index++) {
      throwIfCancelled(shouldCancel);
      onProgress?.(index + 1, pages.length);
      const topPx = Math.round(pages[index].top * scale);
      const heightPx = Math.round(pages[index].height * scale);
      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = canvas.width;
      pageCanvas.height = heightPx;
      const ctx = pageCanvas.getContext("2d");
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
      ctx.drawImage(canvas, 0, topPx, canvas.width, heightPx, 0, 0, canvas.width, heightPx);
      if (index > 0) pdf.addPage();
      const pageHeightMm = (heightPx * pageWidthMm) / pageCanvas.width;
      pdf.addImage(pageCanvas.toDataURL("image/jpeg", 0.9), "JPEG", 0, 0, pageWidthMm, pageHeightMm);
      pageCanvas.width = 1;
      pageCanvas.height = 1;
      await yieldToBrowser();
    }
    canvas.width = 1;
    canvas.height = 1;
    return { blob: pdf.output("blob"), filename };
  } finally {
    container.remove();
    window.scrollTo(savedScrollX, savedScrollY);
  }
}

export async function buildExamPaperReportPdf(report, settings = {}, onProgress, options = {}) {
  const shouldCancel = options.shouldCancel;
  throwIfCancelled(shouldCancel);
  const [{ default: html2canvas }, { jsPDF }] = await Promise.all([
    import("html2canvas"),
    import("jspdf"),
  ]);
  throwIfCancelled(shouldCancel);

  const s = {
    uniName: "Egyptian Russian University",
    facultyName: "Faculty of Management, Economics and Business Technology",
    committeeName: "لجنة القياس والتقويم",
    qualityEmail: "qa-mebt@eru.edu.eg",
    measurementEmail: "meb-maec@eru.edu.eg",
    ...settings,
  };
  const logoSrc = s.logoDataUrl || "/logo.png";
  const container = document.createElement("div");
  container.style.cssText = "position:absolute;top:0;left:0;width:850px;z-index:-2147483647;pointer-events:none;";
  const styleEl = document.createElement("style");
  styleEl.textContent = themedPdfCss(s.colorTheme);
  container.appendChild(styleEl);
  const root = document.createElement("div");
  root.className = "pdf-root";
  root.innerHTML = buildExamPaperPages(report, s, logoSrc);
  container.appendChild(root);
  document.body.appendChild(container);
  await waitForLayout();
  throwIfCancelled(shouldCancel);

  const clean = value => String(value || "").replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, " ").trim();
  const filename = `تقرير تقييم الورقة الامتحانية من حيث الشكل - ${clean(report.department)} - ${clean(report.semester)} ${clean(report.year)}.pdf`;
  return renderRootToPdf({
    root, container, settings: s, onProgress, shouldCancel, filename, html2canvas, jsPDF,
  });
}

// ── main export ─────────────────────────────────────────────────────────────────
// Renders one .pdf-page at a time (rather than one giant html2canvas capture of
// the whole multi-page report) because a long report — e.g. the 25-axis
// `student` schema produces 30 pages — easily exceeds the browser's maximum
// canvas height (~65535px in Chromium) at 2x scale, which silently yields a
// fully transparent canvas with no thrown error. Per-page capture keeps each
// canvas well under that limit regardless of how many axes a schema has.
export async function buildBrandedReportPdf(result, meta, settings, onProgress, options = {}) {
  const shouldCancel = options.shouldCancel;
  throwIfCancelled(shouldCancel);
  const [{ default: html2canvas }, { jsPDF }] = await Promise.all([
    import("html2canvas"),
    import("jspdf"),
  ]);
  throwIfCancelled(shouldCancel);

  const perSurvey = settings.surveyReportOptions?.[result.schemaId] ?? {};
  settings = { ...settings, ...perSurvey };
  if (settings.reportTexts) {
    result = { ...result, reportTexts: { ...(result.reportTexts ?? {}), ...settings.reportTexts } };
  }
  let logoSrc = settings.logoDataUrl || "/logo.png";

  const container = document.createElement("div");
 container.style.cssText = "position:absolute;top:0;left:0;width:850px;z-index:-2147483647;pointer-events:none;";

  const styleEl = document.createElement("style");
  styleEl.textContent = themedPdfCss(settings.colorTheme);
  container.appendChild(styleEl);

  const root = document.createElement("div");
  root.className = "pdf-root";
  container.appendChild(root);

  // The stylesheet and container must already be attached to the document
  // before buildAxisDetailPages runs — it measures real block heights
  // against a `.pdf-page`-classed probe, which only picks up the right CSS
  // once styleEl is live in the DOM.
  document.body.appendChild(container);

  const includeCharts = settings.includePdfCharts !== false;

  // The summary overview chart is mandatory and is rendered directly under
  // the final summary-table chunk. includeCharts only controls axis details.
  const leadingPages = [
    buildCoverPage(result, meta, settings, logoSrc),
    buildContextPage(result, meta, settings, logoSrc),
    result.reportSections?.skipStandardSections === true ? "" : buildMethodologyPage(result, settings, logoSrc),
  ];
  const compactFlatSummary = result.isFlat && result.axes.length === 1;
  const axisPages = await buildAxisDetailPages(
    result, settings, logoSrc, includeCharts, shouldCancel,
    compactFlatSummary ? buildCompactFlatSummary(result) : ""
  );
  await yieldToBrowser();
  throwIfCancelled(shouldCancel);
  root.innerHTML = [
    ...leadingPages,
    compactFlatSummary ? "" : buildSummaryTablePages(result, settings, logoSrc),
    axisPages,
    buildRecommendationsPages(result, settings, logoSrc),
  ].join("");

  await yieldToBrowser();
  throwIfCancelled(shouldCancel);

  const sourceBase = String(meta.sourceName || "").replace(/\.(xlsx|xls|csv)$/i, "").trim();
  const fallbackBase = [result.schemaLabel || "استبيان", meta.program, meta.year].filter(Boolean).join(" - ");
  const filenameBase = (sourceBase || fallbackBase || "استبيان").replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, " ").trim();
  const filename = `تقرير تحليل - ${filenameBase}.pdf`;

  return renderRootToPdf({
    root, container, settings, onProgress, shouldCancel, filename, html2canvas, jsPDF,
  });
}

export async function buildComparisonReportPdf(comparison, meta, settings, onProgress, options = {}) {
  const shouldCancel = options.shouldCancel;
  throwIfCancelled(shouldCancel);
  const [{ default: html2canvas }, { jsPDF }] = await Promise.all([
    import("html2canvas"),
    import("jspdf"),
  ]);
  throwIfCancelled(shouldCancel);

  const firstResult = comparison?.slots?.find(slot => slot?.result)?.result;
  if (!firstResult || !comparison?.overall?.length) throw new Error("بيانات المقارنة غير مكتملة");
  const perSurvey = settings.surveyReportOptions?.[firstResult.schemaId] ?? {};
  settings = { ...settings, ...perSurvey };
  const logoSrc = settings.logoDataUrl || "/logo.png";
  const container = document.createElement("div");
  container.style.cssText = "position:absolute;top:0;left:0;width:850px;z-index:-2147483647;pointer-events:none;";
  const styleEl = document.createElement("style");
  styleEl.textContent = themedPdfCss(settings.colorTheme);
  container.appendChild(styleEl);
  const root = document.createElement("div");
  root.className = "pdf-root";
  root.innerHTML = buildComparisonPages(comparison, meta, settings, logoSrc);
  container.appendChild(root);
  document.body.appendChild(container);
  await waitForLayout();
  throwIfCancelled(shouldCancel);

  const years = comparison.overall.map(item => item.year).filter(Boolean).join(" و ");
  const filenameBase = ["مقارنة", firstResult.schemaLabel, meta.program, years]
    .filter(Boolean).join(" - ").replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, " ").trim();
  const filename = `${filenameBase || "تقرير مقارنة"}.pdf`;
  return renderComparisonRootToPdf({
    root, container, settings, onProgress, shouldCancel, filename, html2canvas, jsPDF,
  });
}

export async function downloadBrandedReportPdf(result, meta, settings, onProgress, options) {
  const built = await buildBrandedReportPdf(result, meta, settings, onProgress, options);
  const url = URL.createObjectURL(built.blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = built.filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  return built;
}
