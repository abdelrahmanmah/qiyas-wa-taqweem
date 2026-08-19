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

function esc(v) {
  return String(v ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function barColor(i) {
  return BAR_COLORS[i % BAR_COLORS.length];
}

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

// Rows/bars per page are picked so a page never has to be split mid-row —
// html2canvas captures one whole .pdf-page-outer per page, so any content
// that overflows the physical page height gets sliced at an arbitrary pixel
// boundary (mid-table-row, mid-bar) further down in downloadBrandedReportPdf.
// Chunking the summary table, the axis chart, and the recommendations table
// up front avoids that for the sections most likely to be long.
const SUMMARY_ROWS_PER_PAGE = 18;
const CHART_BARS_PER_PAGE = 8;
const RECS_ROWS_PER_PAGE = 20;

// ── small building blocks ──────────────────────────────────────────────────────
function pageOpen(extraClass = "") {
  return `<div class="pdf-page-outer"><div class="pdf-page ${extraClass}">`;
}
const pageClose = "</div></div>";

function pageHeader(s, logoSrc) {
  return `
    <div class="pdf-header">
      ${logoSrc ? `<img class="pdf-logo" src="${logoSrc}" alt="" />` : ""}
      <div class="pdf-header-text">
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

function verticalBarChart(items) {
  // items: [{ label, pct }]
  const max = 100;
  return `
    <div class="pdf-vchart">
      ${items.map((it, i) => `
        <div class="pdf-vchart-col">
          <div class="pdf-vchart-pct">${it.pct}%</div>
          <div class="pdf-vchart-track">
            <div class="pdf-vchart-fill" style="height:${Math.max(4, (it.pct / max) * 100)}%;background:${barColor(i)}"></div>
          </div>
          <div class="pdf-vchart-label">${esc(it.label)}</div>
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
    ...(meta.preparedBy ? [["أعدّ التقرير", meta.preparedBy]] : []),
    ...(meta.reviewer ? [["راجع التقرير", meta.reviewer]] : []),
    ["موجه إلى", "مدير وحدة الجودة بالكلية"],
  ];

  return `
    ${pageOpen("pdf-cover")}
      ${pageHeader(s, logoSrc)}
      <div class="pdf-cover-title-wrap">
        <div class="pdf-cover-title">استبيان ${esc(schemaLabel)}</div>
        <div class="pdf-cover-subtitle">تقرير نتائج تحليل الاستبيانات</div>
        <div class="pdf-cover-committee">${esc(s.committeeName)}</div>
      </div>
      ${sectionHeading("البيانات الأساسية")}
      <div class="pdf-info-table">
        ${rows.map(([l, v], i) => infoRow(l, v, i)).join("")}
      </div>
      ${sectionHeading("القائم بالتقييم")}
      <div class="pdf-sig-row">
        <div class="pdf-sig-box">
          <div class="pdf-sig-name">${esc(meta.preparedBy || "—")}</div>
          <div class="pdf-sig-role">أعد التقرير</div>
          <div class="pdf-sig-line">التوقيع: ....................</div>
        </div>
        <div class="pdf-sig-box">
          <div class="pdf-sig-name">${esc(meta.reviewer || "—")}</div>
          <div class="pdf-sig-role">راجع التقرير</div>
          <div class="pdf-sig-line">التوقيع: ....................</div>
        </div>
      </div>
    ${pageClose}`;
}

function buildVisionPage(s, logoSrc) {
  return `
    ${pageOpen()}
      ${pageHeader(s, logoSrc)}
      ${sectionHeading(`رؤية ورسالة ${s.unitName}`)}
      <div class="pdf-vm-block">
        <div class="pdf-vm-title">الرؤية</div>
        <div class="pdf-vm-text">${esc(s.vision)}</div>
      </div>
      <div class="pdf-vm-block">
        <div class="pdf-vm-title">الرسالة</div>
        <div class="pdf-vm-text">${esc(s.mission)}</div>
      </div>
      <div class="pdf-vm-footer">${esc(s.email)}</div>
    ${pageClose}`;
}

function buildMethodologyPage(result, s, logoSrc) {
  const { totalQuestions, axes, scaleType } = result;
  const is5 = scaleType === "likert-5";
  return `
    ${pageOpen()}
      ${pageHeader(s, logoSrc)}
      ${sectionHeading(`نتائج تحليل ${esc(result.schemaLabel)}`)}
      <p class="pdf-body-text">
        يتم إجراء التحليل الإحصائي لتقييم الاستبيان وفقاً لعدد من الخطوات، بدءاً بقيام الطلاب/المشاركين
        بتعبئة الاستبيان، ثم تجميع البيانات ومعالجتها إحصائياً باستخدام المعايير والمؤشرات
        المناسبة، وذلك بهدف تقييم جودة العملية التي يقيسها الاستبيان وتحديد نقاط القوة والجوانب
        التي تتطلب التحسين أو التطوير.
      </p>

      ${sectionHeading("متغيرات الاستبيان")}
      <p class="pdf-body-text">
        اشتمل الاستبيان على (${totalQuestions}) عبارة تتمثل في (${axes.length}) محور رئيسي.
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
    ${pageClose}`;
}

function buildSummaryTablePages(result, s, logoSrc) {
  const { axes, n, overallAgreePct, overallDirection } = result;
  const pages = chunk(axes, SUMMARY_ROWS_PER_PAGE);

  return pages.map((pageAxes, pageIdx) => {
    const isLast = pageIdx === pages.length - 1;
    return `
      ${pageOpen()}
        ${pageHeader(s, logoSrc)}
        ${sectionHeading(pageIdx === 0 ? "ملخص النتائج" : "ملخص النتائج (تابع)")}
        <table class="pdf-table">
          <thead>
            <tr><th>الاتجاه العام</th><th>النسبة</th><th>العدد</th><th class="pdf-th-wide">المحاور</th></tr>
          </thead>
          <tbody>
            ${pageAxes.map((ax, i) => `
              <tr class="${i % 2 === 0 ? "alt" : ""}">
                <td>${esc(ax.direction)}</td>
                <td class="pdf-td-strong">${ax.axisAgreePct}%</td>
                <td>${n}</td>
                <td class="pdf-td-right">${esc(ax.name)}</td>
              </tr>`).join("")}
            ${isLast ? `
              <tr class="pdf-total-row">
                <td>${esc(overallDirection)}</td>
                <td>${overallAgreePct}%</td>
                <td>${n}</td>
                <td class="pdf-td-right">المتوسط العام</td>
              </tr>` : ""}
          </tbody>
        </table>
      ${pageClose}`;
  }).join("");
}

function buildChartPages(result, s, logoSrc) {
  const pages = chunk(result.axes, CHART_BARS_PER_PAGE);
  return pages.map((pageAxes, pageIdx) => `
    ${pageOpen()}
      ${pageHeader(s, logoSrc)}
      ${sectionHeading(pageIdx === 0 ? "التمثيل البياني لنسب تحقق المحاور" : "التمثيل البياني لنسب تحقق المحاور (تابع)")}
      ${verticalBarChart(pageAxes.map(ax => ({ label: ax.name, pct: ax.axisAgreePct })))}
    ${pageClose}`).join("");
}

function buildAxisBlock(ax, result, includeCharts) {
  const is5 = result.scaleType === "likert-5";
  const headerLabels = is5
    ? ["م", "العبارات", "لا أوافق بشدة", "لا أوافق", "محايد", "أوافق", "أوافق بشدة"]
    : ["م", "العبارات", "لا أوافق %", "محايد %", "أوافق %"];

  const fmtPct = v => (v === 0 || v == null) ? "-" : `${v}%`;

  return `
    ${sectionHeading(ax.name)}
    <table class="pdf-table pdf-table-detail">
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
        <tr class="pdf-total-row">
          <td colspan="${is5 ? 6 : 4}" class="pdf-td-right">إجمالي درجات المحور</td>
          <td>${ax.axisAgreePct}%</td>
        </tr>
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
function measureBlockHeightsPx(htmlBlocks) {
  const probe = document.createElement("div");
  probe.className = "pdf-page";
  probe.style.cssText = "position:absolute;visibility:hidden;left:-9999px;top:0;min-height:0;padding:0;border:0;";
  document.body.appendChild(probe);
  const heights = htmlBlocks.map(html => {
    probe.innerHTML = html;
    return probe.getBoundingClientRect().height;
  });
  document.body.removeChild(probe);
  return heights;
}

// Available content height per page: .pdf-page min-height (1080) minus the
// header block (~110) and top/bottom padding (~68), with a safety margin.
const AXIS_PAGE_BUDGET_PX = 850;

function buildAxisDetailPages(result, s, logoSrc, includeCharts) {
  const blocks = result.axes.map(ax => buildAxisBlock(ax, result, includeCharts));
  const heights = measureBlockHeightsPx(blocks);

  const groups = [];
  let current = [];
  let currentHeight = 0;

  result.axes.forEach((ax, i) => {
    const h = heights[i];
    if (current.length > 0 && currentHeight + h > AXIS_PAGE_BUDGET_PX) {
      groups.push(current);
      current = [];
      currentHeight = 0;
    }
    current.push(ax);
    currentHeight += h;
  });
  if (current.length > 0) groups.push(current);

  return groups.map(group => `
    ${pageOpen()}
      ${pageHeader(s, logoSrc)}
      ${group.map(ax => buildAxisBlock(ax, result, includeCharts)).join("")}
    ${pageClose}`).join("");
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
        ${sectionHeading("التوصيات")}
        <div class="pdf-no-recs">لا توجد توصيات</div>
        ${buildSignaturesBlock(s)}
      ${pageClose}`;
  }

  const pages = chunk(rows, RECS_ROWS_PER_PAGE);
  return pages.map((pageRows, pageIdx) => {
    const isLast = pageIdx === pages.length - 1;
    return `
      ${pageOpen()}
        ${pageHeader(s, logoSrc)}
        ${sectionHeading(pageIdx === 0 ? `التوصيات (أقل من ${threshold}%)` : "التوصيات (تابع)")}
        <table class="pdf-table">
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
      ${pageClose}`;
  }).join("");
}

// ── CSS ─────────────────────────────────────────────────────────────────────────
const PDF_CSS = `
  * { box-sizing: border-box; margin: 0; padding: 0; }
  .pdf-root { direction: rtl; font-family: 'Cairo', Arial, sans-serif; color: ${TEXT_DARK}; }
  .pdf-page-outer {
    width: 808px; background: ${PINK_SOFT}; margin: 0 auto 0; padding: 0 14px;
    page-break-after: always; break-after: page;
  }
  .pdf-page {
    width: 780px; min-height: 1080px; background: #fff; margin: 0 auto;
    padding: 34px 44px; position: relative;
    border-left: 5px solid ${PINK_LINE}; border-right: 5px solid ${PINK_LINE};
  }
  .pdf-header { display: flex; align-items: center; justify-content: center; gap: 14px; margin-bottom: 22px; padding-bottom: 14px; border-bottom: 2px solid ${PINK}; }
  .pdf-logo { width: 52px; height: 52px; object-fit: contain; }
  .pdf-header-text { text-align: center; }
  .pdf-uni-name { font-size: 19px; font-weight: 900; color: ${TEXT_DARK}; }
  .pdf-faculty-name { font-size: 12px; color: ${TEXT_MUTED}; margin-top: 2px; }

  .pdf-cover-title-wrap { text-align: center; margin: 30px 0 26px; }
  .pdf-cover-title { display: inline-block; font-size: 26px; font-weight: 900; padding: 6px 22px; background: ${PINK}; border-radius: 6px; color: ${TEXT_DARK}; }
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
  .pdf-vm-footer { margin-top: 30px; font-size: 12px; color: ${TEXT_MUTED}; text-align: center; }

  .pdf-formula-block { background: ${PINK_SOFT}; border-radius: 8px; padding: 12px 16px; margin-bottom: 14px; }
  .pdf-formula-title { font-weight: 800; font-size: 13px; margin-bottom: 6px; }
  .pdf-formula { font-size: 16px; font-weight: 700; direction: ltr; text-align: center; margin: 6px 0; }
  .pdf-formula-note { font-size: 11px; color: ${TEXT_MUTED}; text-align: center; }

  .pdf-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  .pdf-table th { background: ${ACCENT}; color: #fff; padding: 8px 6px; font-weight: 800; }
  .pdf-table td { padding: 7px 6px; text-align: center; border-bottom: 1px solid #f0f0f0; }
  .pdf-table tr.alt td { background: ${PINK_SOFT}; }
  .pdf-table .pdf-th-wide { text-align: right; padding-right: 12px; }
  .pdf-table .pdf-td-right { text-align: right; padding-right: 12px; }
  .pdf-table .pdf-td-strong { font-weight: 800; }
  .pdf-total-row td { background: ${PINK}; font-weight: 900; }
  .pdf-table-detail td { font-size: 11.5px; }

  .pdf-vchart { display: flex; align-items: flex-end; justify-content: space-around; gap: 10px; height: 220px; margin-top: 14px; padding: 0 6px; }
  .pdf-vchart-col { flex: 1; display: flex; flex-direction: column; align-items: center; height: 100%; }
  .pdf-vchart-pct { font-size: 11px; font-weight: 800; margin-bottom: 4px; }
  .pdf-vchart-track { flex: 1; width: 26px; display: flex; align-items: flex-end; background: #f4f4f4; border-radius: 6px; overflow: hidden; }
  .pdf-vchart-fill { width: 100%; border-radius: 6px 6px 0 0; }
  .pdf-vchart-label { font-size: 9.5px; color: ${TEXT_MUTED}; margin-top: 6px; text-align: center; line-height: 1.3; max-width: 70px; }

  .pdf-hbars { margin-top: 18px; display: flex; flex-direction: column; gap: 10px; }
  .pdf-hbar-row { display: flex; align-items: center; gap: 10px; }
  .pdf-hbar-label { flex: 0 0 220px; font-size: 11px; text-align: right; line-height: 1.4; }
  .pdf-hbar-track { flex: 1; background: #f4f4f4; border-radius: 10px; height: 20px; overflow: hidden; }
  .pdf-hbar-fill { height: 100%; border-radius: 10px; display: flex; align-items: center; justify-content: center; }
  .pdf-hbar-pct { font-size: 10.5px; color: #fff; font-weight: 800; padding: 0 6px; }

  .pdf-axis-spacer { height: 22px; }

  .pdf-no-recs { text-align: center; font-weight: 800; color: ${ACCENT}; font-size: 15px; margin-top: 30px; }
`;

// ── main export ─────────────────────────────────────────────────────────────────
// Renders one .pdf-page at a time (rather than one giant html2canvas capture of
// the whole multi-page report) because a long report — e.g. the 25-axis
// `student` schema produces 30 pages — easily exceeds the browser's maximum
// canvas height (~65535px in Chromium) at 2x scale, which silently yields a
// fully transparent canvas with no thrown error. Per-page capture keeps each
// canvas well under that limit regardless of how many axes a schema has.
export async function downloadBrandedReportPdf(result, meta, settings, onProgress) {
  const [{ default: html2canvas }, { jsPDF }] = await Promise.all([
    import("html2canvas"),
    import("jspdf"),
  ]);

  let logoSrc = settings.logoDataUrl || "/logo.png";

  const container = document.createElement("div");
  container.style.cssText = "position:absolute;top:0;left:-3000px;width:850px;";

  const styleEl = document.createElement("style");
  styleEl.textContent = PDF_CSS;
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

  // Section order mirrors buildAnnualDocx() in buildDocx.js: detailed
  // per-axis breakdown (رابعاً) first, then the results summary (خامساً)
  // right before the recommendations — not summary-then-detail.
  root.innerHTML = [
    buildCoverPage(result, meta, settings, logoSrc),
    buildVisionPage(settings, logoSrc),
    buildMethodologyPage(result, settings, logoSrc),
    buildAxisDetailPages(result, settings, logoSrc, includeCharts),
    buildSummaryTablePages(result, settings, logoSrc),
    includeCharts ? buildChartPages(result, settings, logoSrc) : "",
    buildRecommendationsPages(result, settings, logoSrc),
  ].join("");

  const filename = `تقرير_${result.schemaLabel || "استبيان"}_${meta.program || ""}_${meta.year || ""}.pdf`;

  try {
    const pageEls = Array.from(root.querySelectorAll(".pdf-page-outer"));
    const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
    const pageWidthMm = pdf.internal.pageSize.getWidth();
    const pageHeightMm = pdf.internal.pageSize.getHeight();
    let firstImage = true;
    let pageIndex = 0;

    for (const el of pageEls) {
      pageIndex++;
      onProgress?.(pageIndex, pageEls.length);
      const canvas = await html2canvas(el, {
        scale: 2, useCORS: true, logging: false, backgroundColor: "#ffffff", scrollX: 0, scrollY: 0,
      });
      const fullHeightMm = (canvas.height * pageWidthMm) / canvas.width;
      // A section can render taller than one physical A4 page (e.g. a 25-axis
      // summary table) — slice the canvas across as many PDF pages as needed
      // instead of silently cropping it to a single page's height.
      const sliceCount = Math.max(1, Math.ceil(fullHeightMm / pageHeightMm));
      const sliceHeightPx = Math.ceil(canvas.height / sliceCount);

      for (let s = 0; s < sliceCount; s++) {
        const sliceCanvas = document.createElement("canvas");
        sliceCanvas.width = canvas.width;
        sliceCanvas.height = sliceHeightPx;
        const ctx = sliceCanvas.getContext("2d");
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, sliceCanvas.width, sliceCanvas.height);
        ctx.drawImage(
          canvas, 0, s * sliceHeightPx, canvas.width, sliceHeightPx,
          0, 0, canvas.width, sliceHeightPx
        );
        const sliceImgData = sliceCanvas.toDataURL("image/jpeg", 0.95);
        const sliceHeightMm = (sliceHeightPx * pageWidthMm) / canvas.width;
        if (!firstImage) pdf.addPage();
        firstImage = false;
        pdf.addImage(sliceImgData, "JPEG", 0, 0, pageWidthMm, sliceHeightMm);
      }
    }

    pdf.save(filename);
  } finally {
    document.body.removeChild(container);
  }
}
