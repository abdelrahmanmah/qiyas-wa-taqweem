/**
 * EnhancedReportView.jsx
 * ──────────────────────────────────────────────────────────────────────────────
 * ISOLATED component — does NOT alter any existing code.
 * Accepts the exact same `result`, `meta`, and `settings` objects produced by
 * the existing analyze.js / App.jsx pipeline.
 *
 * Usage:
 *   import EnhancedReportView from "./EnhancedReportView";
 *   <EnhancedReportView result={singleResult} meta={meta} settings={settings} />
 */

import { useRef, useState } from "react";

// ── Design tokens ─────────────────────────────────────────────────────────────
const T = {
  navy:     "#1e3a8a",
  navyDark: "#172554",
  navyMid:  "#1d4ed8",
  accent:   "#2563eb",
  accentSoft:"#dbeafe",
  teal:     "#0d9488",
  tealSoft: "#ccfbf1",
  gold:     "#b45309",
  goldSoft: "#fef3c7",
  gray50:   "#f8fafc",
  gray100:  "#f1f5f9",
  gray200:  "#e2e8f0",
  gray300:  "#cbd5e1",
  gray400:  "#94a3b8",
  gray600:  "#475569",
  gray700:  "#334155",
  gray900:  "#0f172a",
  white:    "#ffffff",
  excellent:"#166534",
  excellentBg:"#dcfce7",
  good:     "#1e40af",
  goodBg:   "#dbeafe",
  neutral:  "#92400e",
  neutralBg:"#fef3c7",
  low:      "#991b1b",
  lowBg:    "#fee2e2",
};

// ── Tier color helpers ────────────────────────────────────────────────────────
function tierColor(tier) {
  return tier === "excellent" ? T.excellent
       : tier === "good"      ? T.good
       : tier === "neutral"   ? T.neutral
       :                        T.low;
}
function tierBg(tier) {
  return tier === "excellent" ? T.excellentBg
       : tier === "good"      ? T.goodBg
       : tier === "neutral"   ? T.neutralBg
       :                        T.lowBg;
}

// ── Scale labels ──────────────────────────────────────────────────────────────
const L5_LABELS = {
  "5": "أوافق بشدة",
  "4": "أوافق",
  "3": "محايد",
  "2": "لا أوافق",
  "1": "لا أوافق بشدة",
};
// likert-3 uses string codes: "agree" / "neutral" / "disagree"
const L3_LABELS = {
  "agree":    "أوافق",
  "neutral":  "محايد",
  "disagree": "لا أوافق",
};
// ── Grouped 3-column display (always: موافقة / محايد / غير موافق) ─────────────
const DISPLAY_GROUPS = [
  { key: "agree",    label: "موافقة",    color: "#166534" },
  { key: "neutral",  label: "محايد",     color: "#ca8a04" },
  { key: "disagree", label: "غير موافق", color: "#dc2626" },
];

function groupedCounts(counts, scaleType) {
  if (scaleType === "likert-5") {
    return {
      agree:    (counts["5"] ?? 0) + (counts["4"] ?? 0),
      neutral:   counts["3"] ?? 0,
      disagree: (counts["2"] ?? 0) + (counts["1"] ?? 0),
    };
  }
  return {
    agree:    counts.agree    ?? 0,
    neutral:  counts.neutral  ?? 0,
    disagree: counts.disagree ?? 0,
  };
}

function groupedPcts(pcts, scaleType) {
  if (scaleType === "likert-5") {
    return {
      agree:    (pcts["5"] ?? 0) + (pcts["4"] ?? 0),
      neutral:   pcts["3"] ?? 0,
      disagree: (pcts["2"] ?? 0) + (pcts["1"] ?? 0),
    };
  }
  return {
    agree:    pcts.agree    ?? 0,
    neutral:  pcts.neutral  ?? 0,
    disagree: pcts.disagree ?? 0,
  };
}

function getScaleKeys(scaleType) {
  return scaleType === "likert-5"
    ? ["5","4","3","2","1"]
    : ["agree","neutral","disagree"];
}

// ── Tiny helper: format number ────────────────────────────────────────────────
const fmt = (v) => v == null ? "—" : Number(v).toFixed(2);
const pct  = (v) => v == null ? "—" : `${Number(v).toFixed(1)}%`;

// ── CSS (injected once) ───────────────────────────────────────────────────────
const PRINT_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;500;600;700;900&display=swap');

.erv-root {
  font-family: 'Cairo', 'Segoe UI', Arial, sans-serif;
  direction: rtl;
  color: ${T.gray900};
  background: ${T.gray50};
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}

/* ── Screen-only wrapper ── */
.erv-screen-wrap {
  max-width: 900px;
  margin: 0 auto;
  padding: 32px 20px 64px;
}

/* ── Print button ── */
.erv-print-btn {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  background: ${T.navy};
  color: #fff;
  border: none;
  border-radius: 10px;
  padding: 11px 28px;
  font-family: 'Cairo', sans-serif;
  font-size: 15px;
  font-weight: 700;
  cursor: pointer;
  transition: background .2s, box-shadow .2s;
  margin-bottom: 28px;
}
.erv-print-btn:hover {
  background: ${T.navyDark};
  box-shadow: 0 4px 18px rgba(30,58,138,.35);
}

/* ── Page ── */
.erv-page {
  background: ${T.white};
  box-shadow: 0 2px 24px rgba(0,0,0,.10);
  border-radius: 14px;
  margin-bottom: 24px;
  overflow: hidden;
}

/* ── Cover page ── */
.erv-cover {
  padding: 52px 52px 40px;
  background: linear-gradient(160deg, ${T.navyDark} 0%, ${T.navy} 55%, ${T.accent} 100%);
  color: ${T.white};
  text-align: center;
  position: relative;
  min-height: 320px;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 10px;
}
.erv-cover::after {
  content: '';
  position: absolute;
  bottom: 0; left: 0; right: 0;
  height: 5px;
  background: linear-gradient(90deg, ${T.teal}, ${T.gold}, ${T.accent});
}
.erv-cover-logo {
  width: 80px;
  height: 80px;
  object-fit: contain;
  border-radius: 12px;
  background: rgba(255,255,255,.12);
  padding: 6px;
  margin-bottom: 8px;
}
.erv-cover-logo-placeholder {
  width: 80px;
  height: 80px;
  border-radius: 12px;
  background: rgba(255,255,255,.12);
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 36px;
  margin-bottom: 8px;
}
.erv-cover-uni   { font-size: 14px; font-weight: 600; opacity: .85; letter-spacing: .3px; }
.erv-cover-fac   { font-size: 13px; font-weight: 500; opacity: .70; }
.erv-cover-unit  { font-size: 12px; font-weight: 500; opacity: .60; margin-bottom: 16px; }
.erv-cover-title {
  font-size: 28px;
  font-weight: 900;
  line-height: 1.35;
  margin: 8px 0 4px;
  text-shadow: 0 2px 8px rgba(0,0,0,.25);
}
.erv-cover-sub   { font-size: 16px; font-weight: 700; opacity: .85; }
.erv-cover-prog  { font-size: 13px; font-weight: 600; opacity: .75; background: rgba(255,255,255,.12); padding: 4px 18px; border-radius: 20px; }
.erv-cover-meta  { font-size: 12px; opacity: .60; margin-top: 14px; }

/* ── Section card ── */
.erv-section {
  padding: 32px 36px;
  border-bottom: 1px solid ${T.gray100};
}
.erv-section:last-child { border-bottom: none; }

/* ── Section heading ── */
.erv-section-heading {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 22px;
}
.erv-section-heading-bar {
  width: 5px;
  border-radius: 4px;
  background: linear-gradient(180deg, ${T.navy}, ${T.accent});
  align-self: stretch;
  min-height: 28px;
}
.erv-section-heading h2 {
  font-size: 18px;
  font-weight: 900;
  color: ${T.navyDark};
  margin: 0;
}
.erv-section-heading p {
  font-size: 12px;
  color: ${T.gray400};
  margin: 2px 0 0;
}

/* ── Stat cards row ── */
.erv-stats-row {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 14px;
  margin-bottom: 4px;
}
.erv-stat-card {
  background: ${T.gray50};
  border: 1px solid ${T.gray200};
  border-radius: 12px;
  padding: 18px 14px;
  text-align: center;
  transition: box-shadow .2s;
}
.erv-stat-card:hover { box-shadow: 0 4px 16px rgba(30,58,138,.09); }
.erv-stat-icon  { font-size: 22px; margin-bottom: 6px; }
.erv-stat-value { font-size: 30px; font-weight: 900; color: ${T.navy}; line-height: 1; }
.erv-stat-label { font-size: 11.5px; color: ${T.gray400}; margin-top: 4px; font-weight: 600; }

/* ── Overall badge ── */
.erv-overall-row {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 14px;
  background: ${T.accentSoft};
  border: 1px solid #bfdbfe;
  border-radius: 12px;
  padding: 14px 22px;
  margin-top: 18px;
  flex-wrap: wrap;
}
.erv-overall-label { font-size: 14px; font-weight: 700; color: ${T.navy}; }
.erv-overall-value { font-size: 22px; font-weight: 900; color: ${T.accent}; }
.erv-overall-badge {
  padding: 5px 18px;
  border-radius: 20px;
  font-size: 13px;
  font-weight: 800;
}

/* ── Generic table ── */
.erv-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
}
.erv-table thead tr { background: ${T.navy}; }
.erv-table thead th {
  color: ${T.white};
  padding: 10px 12px;
  font-weight: 700;
  font-size: 12px;
  text-align: center;
  white-space: nowrap;
}
.erv-table thead th.erv-th-right { text-align: right; white-space: normal; word-break: keep-all; }
.erv-table tbody tr:nth-child(even) td { background: ${T.gray50}; }
.erv-table tbody tr:hover td { background: ${T.accentSoft}; }
.erv-table td {
  padding: 9px 12px;
  border-bottom: 1px solid ${T.gray200};
  text-align: center;
  vertical-align: middle;
  line-height: 1.65;
}
.erv-table td.erv-td-right { text-align: right; line-height: 1.65; word-break: keep-all; overflow-wrap: break-word; }
.erv-table .erv-total-row td {
  background: ${T.navyDark} !important;
  color: ${T.white} !important;
  font-weight: 800;
  font-size: 13px;
}

/* ── Direction badge ── */
.erv-badge {
  display: inline-block;
  padding: 3px 12px;
  border-radius: 20px;
  font-size: 11.5px;
  font-weight: 800;
  white-space: nowrap;
}

/* ── Progress bar ── */
.erv-bar-wrap {
  background: ${T.gray200};
  border-radius: 4px;
  height: 7px;
  flex: 1;
  min-width: 60px;
  overflow: hidden;
}
.erv-bar-fill {
  height: 100%;
  border-radius: 4px;
  transition: width .4s ease;
}

/* ── Axis block ── */
.erv-axis-block {
  margin-bottom: 32px;
  border: 1px solid ${T.gray200};
  border-radius: 12px;
  overflow: hidden;
  page-break-inside: avoid;
}
.erv-axis-header {
  background: linear-gradient(135deg, ${T.navy} 0%, ${T.navyMid} 100%);
  color: ${T.white};
  padding: 14px 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  flex-wrap: wrap;
}
.erv-axis-header-title {
  font-size: 14px;
  font-weight: 800;
  line-height: 1.4;
  flex: 1;
}
.erv-axis-header-meta {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  align-items: center;
}
.erv-axis-pill {
  background: rgba(255,255,255,.18);
  border: 1px solid rgba(255,255,255,.25);
  border-radius: 8px;
  padding: 4px 12px;
  font-size: 12px;
  font-weight: 700;
  white-space: nowrap;
}

/* ── Question rows inside axis ── */
.erv-q-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12.5px;
}
.erv-q-table thead tr { background: ${T.gray100}; }
.erv-q-table thead th {
  padding: 9px 10px;
  color: ${T.gray700};
  font-weight: 700;
  font-size: 11.5px;
  text-align: center;
  border-bottom: 2px solid ${T.gray200};
  white-space: nowrap;
}
.erv-q-table thead th.erv-th-right { text-align: right; }
.erv-q-table tbody tr { transition: background .15s; }
.erv-q-table tbody tr:nth-child(even) { background: ${T.gray50}; }
.erv-q-table tbody tr:hover { background: ${T.accentSoft}; }
.erv-q-table td {
  padding: 9px 10px;
  border-bottom: 1px solid ${T.gray100};
  text-align: center;
  vertical-align: middle;
  line-height: 1.55;
}
.erv-q-table td.erv-td-right { text-align: right; font-size: 12px; }
.erv-q-table .erv-q-total td {
  background: #eff6ff !important;
  font-weight: 800;
  color: ${T.navy};
  font-size: 12.5px;
  border-top: 2px solid ${T.gray200};
}

/* ── Stacked mini-bar (distribution) ── */
.erv-dist-bar {
  display: flex;
  height: 8px;
  border-radius: 5px;
  overflow: hidden;
  min-width: 80px;
}
.erv-dist-seg { height: 100%; }

/* ── Scrollable table wrappers (screen only — overridden in PDF) ── */
.erv-q-table-scroll { overflow-x: auto; }

/* ── Cross-tab ── */
.erv-crosstab-wrap {
  overflow-x: auto;
  margin-top: 12px;
}

/* ══════════════════════════════════════════════════════════════════════════════
   PRINT STYLES
   ══════════════════════════════════════════════════════════════════════════════ */
@media print {
  @page {
    size: A4 portrait;
    margin: 14mm 15mm 14mm 15mm;
  }

  html, body {
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
    background: #fff !important;
    margin: 0 !important;
    padding: 0 !important;
  }

  .erv-root { background: #fff !important; }

  /* Remove screen-only padding */
  .erv-screen-wrap {
    max-width: none !important;
    padding: 0 !important;
    margin: 0 !important;
  }

  /* Hide print button */
  .erv-print-btn { display: none !important; }

  /* Remove card shadows / border-radius for print */
  .erv-page {
    box-shadow: none !important;
    border-radius: 0 !important;
    margin-bottom: 0 !important;
    border: none !important;
  }

  /* Cover: full page, then forced break */
  .erv-cover-page-wrap {
    page-break-after: always;
    break-after: page;
  }

  /* Axes summary table: keep together */
  .erv-axis-block {
    page-break-inside: avoid;
    break-inside: avoid;
    border-radius: 0 !important;
  }

  /* Repeat table headers across pages */
  thead { display: table-header-group; }

  /* Avoid row breaks */
  tr { page-break-inside: avoid; break-inside: avoid; }

  /* Section spacing */
  .erv-section { padding: 18px 22px !important; }

  /* Force each section to not orphan */
  .erv-section-heading { page-break-after: avoid; break-after: avoid; }

  /* Stats grid: 4 cols on A4 */
  .erv-stats-row {
    grid-template-columns: repeat(4, 1fr) !important;
  }

  /* Stat card borders stay on print */
  .erv-stat-card {
    border: 1px solid ${T.gray200} !important;
    box-shadow: none !important;
  }

  /* Table header backgrounds */
  .erv-table thead tr,
  .erv-q-table thead tr {
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }

  /* Bar fills */
  .erv-bar-fill,
  .erv-dist-seg {
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }

  /* Overall row */
  .erv-overall-row {
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }

  /* Axis header gradient */
  .erv-axis-header {
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }

  /* Keep axis question tables together where possible */
  .erv-q-table { page-break-inside: auto; }
  .erv-q-table tr { page-break-inside: avoid; break-inside: avoid; }

  /* Section divider */
  .erv-section-page-break {
    page-break-before: always;
    break-before: page;
  }
}
`;

// Extra CSS injected only during PDF rendering (appended to PRINT_CSS in the off-screen container)
const PDF_EXTRA_CSS = `
  /* Remove all overflow clipping so nothing gets cut */
  * { overflow: visible !important; overflow-x: visible !important; overflow-y: visible !important; }

  /* Hide screen-only chrome */
  .erv-print-btn { display: none !important; }
  .erv-screen-wrap { max-width: none !important; padding: 0 !important; margin: 0 !important; }

  /* Clean card look */
  .erv-page { box-shadow: none !important; border-radius: 4px !important; margin-bottom: 12px !important; }
  .erv-axis-block { border-radius: 4px !important; }
  .erv-stat-card  { box-shadow: none !important; }

  /* Tighter question table to fit A4 comfortably */
  .erv-q-table { font-size: 11.5px !important; }
  .erv-q-table thead th { font-size: 10.5px !important; padding: 7px 7px !important; white-space: normal !important; }
  .erv-q-table td { padding: 7px 7px !important; }
  .erv-q-table td.erv-td-right { font-size: 11px !important; }

  /* Let question text wrap freely */
  .erv-th-right, .erv-td-right { white-space: normal !important; word-break: break-word !important; }

  /* Cover gradient must print */
  .erv-cover, .erv-axis-header {
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
`;

// ── Sub-components ────────────────────────────────────────────────────────────

/** Stacked distribution bar — always 3 groups: موافقة / محايد / غير موافق */
function DistBar({ counts, pcts, scaleType }) {
  const gp = groupedPcts(pcts, scaleType);
  return (
    <div className="erv-dist-bar" title={DISPLAY_GROUPS.map(g => `${g.label}: ${gp[g.key].toFixed(1)}%`).join(" | ")}>
      {DISPLAY_GROUPS.map((g) => (
        <div
          key={g.key}
          className="erv-dist-seg"
          style={{
            width: `${gp[g.key]}%`,
            background: g.color,
            minWidth: gp[g.key] > 2 ? undefined : 0,
          }}
        />
      ))}
    </div>
  );
}

/** Single axis detailed block */
function AxisBlock({ ax, index, scaleType, result }) {
  const is5 = scaleType === "likert-5";

  return (
    <div className="erv-axis-block">
      {/* Axis header */}
      <div className="erv-axis-header">
        <div className="erv-axis-header-title">
          <span style={{ opacity: .7, fontSize: 12, marginLeft: 8 }}>المحور {index + 1}</span>
          {ax.name}
        </div>
        <div className="erv-axis-header-meta">
          {is5 && ax.axisMean != null && (
            <div className="erv-axis-pill">
              المتوسط: {fmt(ax.axisMean)}
            </div>
          )}
          <div className="erv-axis-pill">
            الموافقة: {pct(ax.axisAgreePct)}
          </div>
          <div
            className="erv-badge"
            style={{ background: tierBg(ax.tier), color: tierColor(ax.tier) }}
          >
            {ax.direction}
          </div>
        </div>
      </div>

      {/* Questions table */}
      <div className="erv-q-table-scroll">
        <table className="erv-q-table">
          <thead>
            <tr>
              <th style={{ width: 32 }}>م</th>
              <th className="erv-th-right" style={{ minWidth: 220 }}>العبارة</th>
              {DISPLAY_GROUPS.map((g) => (
                <th key={g.key} style={{ color: g.color, minWidth: 60 }}>{g.label}</th>
              ))}
              <th style={{ minWidth: 70 }}>التوزيع</th>
              {is5 && <th style={{ minWidth: 60 }}>المتوسط</th>}
              <th style={{ minWidth: 68 }}>الموافقة</th>
            </tr>
          </thead>
          <tbody>
            {ax.questions.map((q, qi) => {
              const gc = groupedCounts(q.counts, scaleType);
              const gp = groupedPcts(q.pcts, scaleType);
              return (
              <tr key={q.id}>
                <td style={{ color: T.navy, fontWeight: 700 }}>{q.seq ?? qi + 1}</td>
                <td className="erv-td-right">{q.text}</td>
                {DISPLAY_GROUPS.map((g) => (
                  <td key={g.key}>
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 2 }}>
                      <span style={{ fontWeight: 600 }}>{gc[g.key]}</span>
                      <span style={{ fontSize: 10, color: T.gray400 }}>
                        {gp[g.key].toFixed(0)}%
                      </span>
                    </div>
                  </td>
                ))}
                <td>
                  <DistBar counts={q.counts} pcts={q.pcts} scaleType={scaleType} />
                </td>
                {is5 && (
                  <td style={{ fontWeight: 700, color: T.navy }}>{fmt(q.mean)}</td>
                )}
                <td>
                  <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                    <div className="erv-bar-wrap">
                      <div
                        className="erv-bar-fill"
                        style={{
                          width: `${q.agreePct ?? 0}%`,
                          background: `linear-gradient(90deg, ${T.teal}, ${T.accent})`,
                        }}
                      />
                    </div>
                    <span style={{ fontWeight: 700, color: T.accent, fontSize: 12, minWidth: 36, textAlign: "left" }}>
                      {pct(q.agreePct)}
                    </span>
                  </div>
                </td>
              </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/** Section heading */
function SectionHeading({ number, title, subtitle }) {
  return (
    <div className="erv-section-heading">
      <div className="erv-section-heading-bar" />
      <div>
        {number && (
          <div style={{ fontSize: 11, fontWeight: 700, color: T.accent, marginBottom: 2, letterSpacing: .5 }}>
            {number}
          </div>
        )}
        <h2>{title}</h2>
        {subtitle && <p>{subtitle}</p>}
      </div>
    </div>
  );
}

// ── Main component ────────────────────────────────────────────────────────────
export default function EnhancedReportView({ result, meta = {}, settings = {} }) {
  const rootRef  = useRef(null);
  const [pdfBusy, setPdfBusy] = useState(false);

  if (!result) {
    return (
      <div className="erv-root" style={{ padding: 40, textAlign: "center", color: T.gray400 }}>
        لا توجد بيانات للعرض. قم برفع ملف استبيان أولاً.
      </div>
    );
  }

  const {
    schemaLabel, scaleType, n, axes,
    overallMean, overallAgreePct, overallDirection,
    byDegree, byDepartment, crossDegreeByDept,
    totalQuestions,
  } = result;

  const is5          = scaleType === "likert-5";
  const hasDegree    = byDegree    && Object.keys(byDegree).length    > 0;
  const hasDept      = byDepartment && Object.keys(byDepartment).length > 0;
  const hasCrossTab  = crossDegreeByDept && Array.isArray(crossDegreeByDept.rows) && crossDegreeByDept.rows.length > 0;
  // Use the tier already computed by analyze.js (respects schema interpretation thresholds)
  const overallTier  = result.overallTier ?? "good";

  const handlePrint = async () => {
    if (pdfBusy) return;
    setPdfBusy(true);
    try {
      const { default: html2pdf } = await import("html2pdf.js");

      // Build an off-screen container that mirrors only the report content
      const container = document.createElement("div");
      container.style.cssText = "position:fixed;top:-99999px;left:-99999px;width:900px;direction:rtl;";

      // Inject combined CSS (screen + PDF overrides)
      const styleEl = document.createElement("style");
      styleEl.textContent = PRINT_CSS + PDF_EXTRA_CSS;

      const root = document.createElement("div");
      root.className = "erv-root";
      root.appendChild(styleEl);

      // Clone the visible content (screen-wrap)
      const wrap = rootRef.current?.querySelector(".erv-screen-wrap");
      if (!wrap) return;
      const clone = wrap.cloneNode(true);

      // Remove the download button from the clone
      clone.querySelector(".erv-print-btn")?.remove();

      root.appendChild(clone);
      container.appendChild(root);
      document.body.appendChild(container);

      const filename = `تقرير-${schemaLabel || "استبيان"}.pdf`;

      await html2pdf()
        .set({
          margin: [12, 12, 12, 12],          // mm — top right bottom left
          filename,
          image:      { type: "jpeg", quality: 0.97 },
          html2canvas: {
            scale:       2,                   // high-DPI
            useCORS:     true,
            logging:     false,
            windowWidth: 900,                 // same as container width
          },
          jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
          pagebreak: {
            mode:  ["css", "legacy"],
            avoid: [".erv-axis-block", ".erv-stat-card", ".erv-section-heading", ".erv-overall-row"],
          },
        })
        .from(root)
        .save();

      document.body.removeChild(container);
    } catch (err) {
      console.error("PDF generation failed:", err);
      alert("حدث خطأ أثناء إنشاء PDF. حاول مرة أخرى.");
    } finally {
      setPdfBusy(false);
    }
  };

  return (
    <div className="erv-root" ref={rootRef}>
      {/* Inject CSS */}
      <style>{PRINT_CSS}</style>

      <div className="erv-screen-wrap">

        {/* Download PDF button */}
        <button
          className="erv-print-btn"
          onClick={handlePrint}
          disabled={pdfBusy}
          style={pdfBusy ? { opacity: 0.65, cursor: "wait" } : undefined}
        >
          {pdfBusy ? "⏳ جاري إنشاء PDF..." : "⬇ تحميل PDF"}
        </button>

        {/* ══════════════════════════════════════════════════════════════════
            COVER PAGE
           ══════════════════════════════════════════════════════════════════ */}
        <div className="erv-cover-page-wrap">
          <div className="erv-page">
            <div className="erv-cover">
              {/* Logo */}
              {settings.logoDataUrl ? (
                <img src={settings.logoDataUrl} className="erv-cover-logo" alt="university logo" />
              ) : (
                <div className="erv-cover-logo-placeholder">🎓</div>
              )}

              <div className="erv-cover-uni">{settings.uniName || "Egyptian Russian University"}</div>
              <div className="erv-cover-fac">{settings.facultyName || "كلية الإدارة والاقتصاد وتكنولوجيا الأعمال"}</div>
              <div className="erv-cover-unit">{settings.unitName || "وحدة ضمان الجودة"}</div>

              <div style={{ width: "60%", height: 1, background: "rgba(255,255,255,.25)", margin: "4px 0 12px" }} />

              <div className="erv-cover-title">
                تقرير تحليل استبيان
                <br />
                {schemaLabel}
              </div>

              {meta.program && (
                <div className="erv-cover-prog">برنامج: {meta.program}</div>
              )}
              {meta.year && (
                <div className="erv-cover-sub">العام الأكاديمي: {meta.year}</div>
              )}

              <div style={{ width: "40%", height: 1, background: "rgba(255,255,255,.18)", margin: "12px 0 6px" }} />

              <div className="erv-cover-meta">
                {meta.preparedBy && <span>أعده: {meta.preparedBy}</span>}
                {meta.preparedBy && meta.reviewer && <span style={{ margin: "0 10px" }}>|</span>}
                {meta.reviewer && <span>راجعه: {meta.reviewer}</span>}
                {settings.committeeName && (
                  <>
                    {(meta.preparedBy || meta.reviewer) && <br />}
                    <span>{settings.committeeName}</span>
                  </>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* ══════════════════════════════════════════════════════════════════
            SECTION 1 — ملخص تنفيذي
           ══════════════════════════════════════════════════════════════════ */}
        <div className="erv-page">
          <div className="erv-section">
            <SectionHeading number="أولاً" title="الملخص التنفيذي" subtitle="إحصائيات عامة للاستبيان" />

            {/* Stat cards */}
            <div className="erv-stats-row">
              {[
                { icon: "👥", value: n,              label: "إجمالي المستجيبين" },
                { icon: "📋", value: axes.length,    label: "عدد المحاور"       },
                { icon: "❓", value: totalQuestions,  label: "عدد الأسئلة"       },
                { icon: "📊", value: `${overallAgreePct}%`, label: "نسبة الموافقة الكلية" },
              ].map((s, i) => (
                <div key={i} className="erv-stat-card">
                  <div className="erv-stat-icon">{s.icon}</div>
                  <div className="erv-stat-value">{s.value}</div>
                  <div className="erv-stat-label">{s.label}</div>
                </div>
              ))}
            </div>

            {/* Overall direction pill */}
            <div className="erv-overall-row">
              <span className="erv-overall-label">الاتجاه الكلي للاستجابة:</span>
              {is5 && overallMean != null && (
                <>
                  <span className="erv-overall-label">المتوسط العام:</span>
                  <span className="erv-overall-value">{fmt(overallMean)}</span>
                  <span style={{ color: T.gray400 }}>|</span>
                </>
              )}
              <span className="erv-overall-label">نسبة الموافقة:</span>
              <span className="erv-overall-value">{pct(overallAgreePct)}</span>
              <span
                className="erv-badge erv-overall-badge"
                style={{ background: tierBg(overallTier), color: tierColor(overallTier) }}
              >
                {overallDirection}
              </span>
            </div>
          </div>
        </div>

        {/* ══════════════════════════════════════════════════════════════════
            SECTION 2 — المشاركون
           ══════════════════════════════════════════════════════════════════ */}
        {(hasDegree || hasDept || hasCrossTab) && (
          <div className="erv-page">
            <div className="erv-section">
              <SectionHeading number="ثانياً" title="المشاركون" subtitle="توزيع المستجيبين حسب المتغيرات الديموغرافية" />

              <div style={{ display: "grid", gridTemplateColumns: hasDegree && hasDept ? "1fr 1fr" : "1fr", gap: 20 }}>
                {/* By degree */}
                {hasDegree && (
                  <div>
                    <div style={{ fontWeight: 700, color: T.navy, fontSize: 13, marginBottom: 10 }}>
                      التوزيع حسب الدرجة / المسمى الوظيفي
                    </div>
                    <table className="erv-table">
                      <thead>
                        <tr>
                          <th>العدد</th>
                          <th>النسبة</th>
                          <th className="erv-th-right">المسمى</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Object.entries(byDegree)
                          .sort((a, b) => b[1] - a[1])
                          .map(([k, v], i) => (
                            <tr key={i}>
                              <td style={{ fontWeight: 700 }}>{v}</td>
                              <td style={{ color: T.accent, fontWeight: 700 }}>
                                {((v / n) * 100).toFixed(1)}%
                              </td>
                              <td className="erv-td-right">{k}</td>
                            </tr>
                          ))}
                        <tr className="erv-total-row">
                          <td>{n}</td>
                          <td>100%</td>
                          <td className="erv-td-right">الإجمالي</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                )}

                {/* By department */}
                {hasDept && (
                  <div>
                    <div style={{ fontWeight: 700, color: T.navy, fontSize: 13, marginBottom: 10 }}>
                      التوزيع حسب التخصص / القسم
                    </div>
                    <table className="erv-table">
                      <thead>
                        <tr>
                          <th>العدد</th>
                          <th>النسبة</th>
                          <th className="erv-th-right">القسم</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Object.entries(byDepartment)
                          .sort((a, b) => b[1] - a[1])
                          .map(([k, v], i) => (
                            <tr key={i}>
                              <td style={{ fontWeight: 700 }}>{v}</td>
                              <td style={{ color: T.accent, fontWeight: 700 }}>
                                {((v / n) * 100).toFixed(1)}%
                              </td>
                              <td className="erv-td-right">{k}</td>
                            </tr>
                          ))}
                        <tr className="erv-total-row">
                          <td>{n}</td>
                          <td>100%</td>
                          <td className="erv-td-right">الإجمالي</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              {/* Cross-tab table */}
              {hasCrossTab && (
                <div style={{ marginTop: 24 }}>
                  <div style={{ fontWeight: 700, color: T.navy, fontSize: 13, marginBottom: 10 }}>
                    التوزيع المتقاطع (الدرجة × التخصص)
                  </div>
                  <div className="erv-crosstab-wrap">
                    <table className="erv-table">
                      <thead>
                        <tr>
                          <th className="erv-th-right">الدرجة \ التخصص</th>
                          {crossDegreeByDept.cols.map((c, i) => (
                            <th key={i}>{c}</th>
                          ))}
                          <th>الإجمالي</th>
                        </tr>
                      </thead>
                      <tbody>
                        {crossDegreeByDept.rows.map((row, ri) => (
                          <tr key={ri}>
                            <td className="erv-td-right" style={{ fontWeight: 600 }}>{row}</td>
                            {crossDegreeByDept.cols.map((col, ci) => (
                              <td key={ci}>
                                {/* matrix is keyed by label strings, not numeric indices */}
                                {crossDegreeByDept.matrix[row]?.[col] ?? 0}
                              </td>
                            ))}
                            <td style={{ fontWeight: 700, color: T.navy }}>
                              {crossDegreeByDept.rowTotals[row] ?? 0}
                            </td>
                          </tr>
                        ))}
                        <tr className="erv-total-row">
                          <td className="erv-td-right">الإجمالي</td>
                          {/* colTotals is an object keyed by label strings, not an array */}
                          {crossDegreeByDept.cols.map((col, ci) => (
                            <td key={ci}>{crossDegreeByDept.colTotals[col] ?? 0}</td>
                          ))}
                          <td>{crossDegreeByDept.grandTotal ?? 0}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ══════════════════════════════════════════════════════════════════
            SECTION 3 — ملخص المحاور
           ══════════════════════════════════════════════════════════════════ */}
        <div className="erv-page">
          <div className="erv-section">
            <SectionHeading number="ثالثاً" title="ملخص المحاور" subtitle="نتائج كل محور بالمتوسط ونسبة الموافقة" />

            <div style={{ overflowX: "auto", direction: "ltr" }}>
            <table className="erv-table" style={{ minWidth: 560, direction: "rtl" }}>
              <thead>
                <tr>
                  <th style={{ width: 38, whiteSpace: "nowrap" }}>م</th>
                  <th className="erv-th-right" style={{ minWidth: 180 }}>المحور</th>
                  <th style={{ whiteSpace: "nowrap" }}>الأسئلة</th>
                  {is5 && <th style={{ whiteSpace: "nowrap" }}>المتوسط</th>}
                  <th style={{ whiteSpace: "nowrap" }}>نسبة الموافقة</th>
                  <th style={{ minWidth: 90, whiteSpace: "nowrap" }}>مؤشر الموافقة</th>
                  <th style={{ whiteSpace: "nowrap" }}>الاتجاه</th>
                </tr>
              </thead>
              <tbody>
                {axes.map((ax, i) => (
                  <tr key={ax.id}>
                    <td style={{ color: T.navy, fontWeight: 800, whiteSpace: "nowrap" }}>{i + 1}</td>
                    <td className="erv-td-right" style={{ fontWeight: 600 }}>{ax.name}</td>
                    <td style={{ whiteSpace: "nowrap" }}>{ax.questions.length}</td>
                    {is5 && (
                      <td style={{ fontWeight: 700, color: T.navyMid, whiteSpace: "nowrap" }}>
                        {fmt(ax.axisMean)}
                      </td>
                    )}
                    <td style={{ fontWeight: 700, color: T.accent, whiteSpace: "nowrap" }}>
                      {pct(ax.axisAgreePct)}
                    </td>
                    <td>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <div className="erv-bar-wrap">
                          <div
                            className="erv-bar-fill"
                            style={{
                              width: `${ax.axisAgreePct ?? 0}%`,
                              background: `linear-gradient(90deg, ${tierColor(ax.tier)}, ${tierBg(ax.tier) === T.excellentBg ? T.teal : tierColor(ax.tier)})`,
                            }}
                          />
                        </div>
                      </div>
                    </td>
                    <td>
                      <span
                        className="erv-badge"
                        style={{ background: tierBg(ax.tier), color: tierColor(ax.tier) }}
                      >
                        {ax.direction}
                      </span>
                    </td>
                  </tr>
                ))}

                {/* Total row */}
                <tr className="erv-total-row">
                  <td colSpan={is5 ? 3 : 3} className="erv-td-right" style={{ whiteSpace: "nowrap" }}>
                    الإجمالي الكلي
                  </td>
                  {is5 && (
                    <td style={{ whiteSpace: "nowrap" }}>{overallMean != null ? fmt(overallMean) : "—"}</td>
                  )}
                  <td style={{ whiteSpace: "nowrap" }}>{pct(overallAgreePct)}</td>
                  <td></td>
                  <td>
                    <span
                      className="erv-badge"
                      style={{ background: tierBg(overallTier), color: tierColor(overallTier) }}
                    >
                      {overallDirection}
                    </span>
                  </td>
                </tr>
              </tbody>
            </table>
            </div>{/* end overflow wrapper ثالثاً */}
          </div>
        </div>

        {/* ══════════════════════════════════════════════════════════════════
            SECTION 4 — النتائج التفصيلية (axis by axis)
           ══════════════════════════════════════════════════════════════════ */}
        <div className="erv-page">
          <div className="erv-section erv-section-page-break">
            <SectionHeading number="رابعاً" title="النتائج التفصيلية" subtitle="تحليل كل محور وعباراته بصورة تفصيلية" />

            {/* Scale legend */}
            <div style={{
              display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 22,
              padding: "10px 14px", background: T.gray50,
              borderRadius: 10, border: `1px solid ${T.gray200}`,
            }}>
              <span style={{ fontSize: 11.5, color: T.gray600, fontWeight: 700, marginLeft: 6 }}>مفتاح المقياس:</span>
              {DISPLAY_GROUPS.map((g) => (
                <span key={g.key} style={{
                  display: "inline-flex", alignItems: "center", gap: 5,
                  fontSize: 11.5, color: T.gray700,
                }}>
                  <span style={{ width: 10, height: 10, borderRadius: 3, background: g.color, display: "inline-block" }} />
                  {g.label}
                </span>
              ))}
            </div>

            {axes.map((ax, i) => (
              <AxisBlock
                key={ax.id}
                ax={ax}
                index={i}
                scaleType={scaleType}
                result={result}
              />
            ))}
          </div>
        </div>

        {/* ══════════════════════════════════════════════════════════════════
            SECTION 5 — التوصيات (if any axis has recommendation text)
           ══════════════════════════════════════════════════════════════════ */}
        {axes.some(ax => ax.recommendation && (ax.axisAgreePct ?? 100) < 70) && (
          <div className="erv-page">
            <div className="erv-section">
              <SectionHeading number="خامساً" title="التوصيات" subtitle="التوصيات المستخلصة من نتائج التحليل" />
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                {axes
                  .filter(ax => ax.recommendation && (ax.axisAgreePct ?? 100) < 70)
                  .map((ax, i) => (
                    <div
                      key={ax.id}
                      style={{
                        display: "flex", gap: 14, alignItems: "flex-start",
                        padding: "14px 16px",
                        background: T.gray50,
                        border: `1px solid ${T.gray200}`,
                        borderRadius: 10,
                        borderRight: `4px solid ${T.navy}`,
                      }}
                    >
                      <span style={{
                        width: 28, height: 28, borderRadius: "50%",
                        background: T.navy, color: "#fff",
                        display: "flex", alignItems: "center", justifyContent: "center",
                        fontWeight: 900, fontSize: 13, flexShrink: 0,
                      }}>{i + 1}</span>
                      <div>
                        <div style={{ fontWeight: 700, color: T.navy, fontSize: 13, marginBottom: 4 }}>
                          {ax.name}
                        </div>
                        <div style={{ color: T.gray700, fontSize: 13, lineHeight: 1.7 }}>
                          {ax.recommendation}
                        </div>
                      </div>
                    </div>
                  ))}
              </div>
            </div>
          </div>
        )}

        {/* ══════════════════════════════════════════════════════════════════
            SECTION — جدول ملخص النتائج
           ══════════════════════════════════════════════════════════════════ */}
        <div className="erv-page">
          <div className="erv-section">
            <SectionHeading
              number={axes.some(ax => ax.recommendation && (ax.axisAgreePct ?? 100) < 70) ? "سادساً" : "خامساً"}
              title="جدول ملخص النتائج"
              subtitle="نظرة شاملة على أداء جميع المحاور"
            />

            <div style={{ overflowX: "auto", direction: "ltr" }}>
            <table className="erv-table" style={{ minWidth: 680, direction: "rtl" }}>
              <thead>
                <tr>
                  <th style={{ width: 36, whiteSpace: "nowrap" }}>م</th>
                  <th className="erv-th-right" style={{ minWidth: 180 }}>المحور</th>
                  <th style={{ minWidth: 52, whiteSpace: "nowrap" }}>الأسئلة</th>
                  {is5 && <th style={{ minWidth: 62, whiteSpace: "nowrap" }}>المتوسط</th>}
                  <th style={{ minWidth: 64, whiteSpace: "nowrap", color: DISPLAY_GROUPS[0].color }}>موافقة</th>
                  <th style={{ minWidth: 58, whiteSpace: "nowrap", color: DISPLAY_GROUPS[1].color }}>محايد</th>
                  <th style={{ minWidth: 70, whiteSpace: "nowrap", color: DISPLAY_GROUPS[2].color }}>غير موافق</th>
                  <th style={{ minWidth: 100, whiteSpace: "nowrap" }}>مؤشر الموافقة</th>
                  <th style={{ minWidth: 78, whiteSpace: "nowrap" }}>نسبة الموافقة</th>
                  <th style={{ minWidth: 64, whiteSpace: "nowrap" }}>الاتجاه</th>
                </tr>
              </thead>
              <tbody>
                {axes.map((ax, i) => {
                  const agreeCount    = ax.questions.reduce((s, q) => s + groupedCounts(q.counts, scaleType).agree,    0);
                  const neutralCount  = ax.questions.reduce((s, q) => s + groupedCounts(q.counts, scaleType).neutral,  0);
                  const disagreeCount = ax.questions.reduce((s, q) => s + groupedCounts(q.counts, scaleType).disagree, 0);
                  const total = agreeCount + neutralCount + disagreeCount || 1;
                  const agreePct    = agreeCount    / total * 100;
                  const neutralPct  = neutralCount  / total * 100;
                  const disagreePct = disagreeCount / total * 100;
                  return (
                    <tr key={ax.id}>
                      <td style={{ color: T.navy, fontWeight: 800, whiteSpace: "nowrap" }}>{i + 1}</td>
                      <td className="erv-td-right" style={{ fontWeight: 600 }}>{ax.name}</td>
                      <td style={{ whiteSpace: "nowrap" }}>{ax.questions.length}</td>
                      {is5 && (
                        <td style={{ fontWeight: 700, color: T.navyMid, whiteSpace: "nowrap" }}>{fmt(ax.axisMean)}</td>
                      )}
                      <td style={{ fontWeight: 700, color: DISPLAY_GROUPS[0].color, whiteSpace: "nowrap" }}>{agreePct.toFixed(1)}%</td>
                      <td style={{ fontWeight: 700, color: DISPLAY_GROUPS[1].color, whiteSpace: "nowrap" }}>{neutralPct.toFixed(1)}%</td>
                      <td style={{ fontWeight: 700, color: DISPLAY_GROUPS[2].color, whiteSpace: "nowrap" }}>{disagreePct.toFixed(1)}%</td>
                      <td>
                        <div style={{ display: "flex", height: 9, borderRadius: 5, overflow: "hidden", minWidth: 90 }}>
                          <div style={{ width: `${agreePct}%`,    background: DISPLAY_GROUPS[0].color }} />
                          <div style={{ width: `${neutralPct}%`,  background: DISPLAY_GROUPS[1].color }} />
                          <div style={{ width: `${disagreePct}%`, background: DISPLAY_GROUPS[2].color }} />
                        </div>
                      </td>
                      <td style={{ fontWeight: 700, color: T.accent, whiteSpace: "nowrap" }}>{pct(ax.axisAgreePct)}</td>
                      <td>
                        <span className="erv-badge" style={{ background: tierBg(ax.tier), color: tierColor(ax.tier), whiteSpace: "nowrap" }}>
                          {ax.direction}
                        </span>
                      </td>
                    </tr>
                  );
                })}

                {/* Overall total row */}
                <tr className="erv-total-row">
                  <td colSpan={is5 ? 3 : 3} className="erv-td-right" style={{ whiteSpace: "nowrap" }}>الإجمالي الكلي</td>
                  {is5 && <td style={{ whiteSpace: "nowrap" }}>{overallMean != null ? fmt(overallMean) : "—"}</td>}
                  {(() => {
                    const allQ = axes.flatMap(ax => ax.questions);
                    const ag  = allQ.reduce((s, q) => s + groupedCounts(q.counts, scaleType).agree,    0);
                    const ne  = allQ.reduce((s, q) => s + groupedCounts(q.counts, scaleType).neutral,  0);
                    const di  = allQ.reduce((s, q) => s + groupedCounts(q.counts, scaleType).disagree, 0);
                    const tot = ag + ne + di || 1;
                    return (
                      <>
                        <td style={{ whiteSpace: "nowrap" }}>{(ag / tot * 100).toFixed(1)}%</td>
                        <td style={{ whiteSpace: "nowrap" }}>{(ne / tot * 100).toFixed(1)}%</td>
                        <td style={{ whiteSpace: "nowrap" }}>{(di / tot * 100).toFixed(1)}%</td>
                      </>
                    );
                  })()}
                  <td />
                  <td style={{ whiteSpace: "nowrap" }}>{pct(overallAgreePct)}</td>
                  <td>
                    <span className="erv-badge" style={{ background: tierBg(overallTier), color: tierColor(overallTier), whiteSpace: "nowrap" }}>
                      {overallDirection}
                    </span>
                  </td>
                </tr>
              </tbody>
            </table>
            </div>{/* end overflow wrapper */}
          </div>
        </div>

        {/* ══════════════════════════════════════════════════════════════════
            FOOTER
           ══════════════════════════════════════════════════════════════════ */}
        <div style={{
          textAlign: "center", padding: "16px 0",
          color: T.gray400, fontSize: 11.5,
          borderTop: `1px solid ${T.gray200}`,
          marginTop: 8,
        }}>
          {settings.uniName} — {settings.unitName}
          {meta.year && ` — العام الأكاديمي ${meta.year}`}
          {settings.email && (
            <div style={{ marginTop: 3, direction: "ltr" }}>{settings.email}</div>
          )}
        </div>

      </div>
    </div>
  );
}
