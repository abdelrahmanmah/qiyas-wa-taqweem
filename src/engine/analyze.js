import * as XLSX from "xlsx";
import { SCHEMAS } from "../schemas/index.js";
export { SCHEMAS };

// ── Arabic normalization ──────────────────────────────────────────────────────
function normalize(v) {
  if (v == null) return "";
  return String(v)
    .trim()
    .replace(/[أإآ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .replace(/[\s_ـ]+/g, " ")
    .toLowerCase();
}

function similarity(a, b) {
  const na = normalize(a), nb = normalize(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  // Jaro-like: shared chars / max length
  const longer = na.length > nb.length ? na : nb;
  const shorter = na.length <= nb.length ? na : nb;
  let matches = 0;
  for (const ch of shorter) if (longer.includes(ch)) matches++;
  return (matches / longer.length) * 0.7 + (na.includes(nb) || nb.includes(na) ? 0.3 : 0);
}

// ── Survey type auto-detection ────────────────────────────────────────────────
export function detectSurveyType(filename, headers) {
  const name = normalize(filename);
  for (const schema of Object.values(SCHEMAS)) {
    if (schema.fileHints.some(h => name.includes(normalize(h)))) {
      return schema.id;
    }
  }
  // Score by header overlap
  const scores = Object.values(SCHEMAS).map(schema => {
    const allQTexts = schema.axes.flatMap(ax => ax.questions.map(q => q.text));
    const score = headers.reduce((sum, h) => {
      const best = Math.max(...allQTexts.map(t => similarity(h, t)));
      return sum + best;
    }, 0);
    return { id: schema.id, score };
  });
  scores.sort((a, b) => b.score - a.score);
  return scores[0].score > 5 ? scores[0].id : null;
}

// ── Column mapping ────────────────────────────────────────────────────────────
function mapColumns(headers, schema) {
  if (schema.scale.type === "likert-5") {
    // Student survey: colIndex is 0-based into the question column array
    // (headers after filtering out Timestamp, level, free-text cols)
    const metaKeywords = [
      ...(schema.metadata?.timestampCol ?? ["Timestamp"]),
      ...(schema.metadata?.levelCol ?? ["المستوى"]),
      ...(schema.metadata?.freeTextCols ?? ["مقترحات", "اية مقترحات"]),
      ...(schema.metadata?.degreeCol ?? []),
      ...(schema.metadata?.departmentCol ?? []),
    ];
    const qColIdx = headers.reduce((acc, h, i) => {
      const hn = normalize(String(h ?? ""));
      const isMeta = metaKeywords.some(k => hn.includes(normalize(k)));
      if (!isMeta) acc.push(i);
      return acc;
    }, []);

    const map = {};
    schema.axes.forEach(ax =>
      ax.questions.forEach(q => {
        map[q.id] = qColIdx[q.colIndex] ?? q.colIndex;
      })
    );
    return map;
  }

  // Faculty/assistant: fuzzy-match headers against question texts
  const startIdx = schema.questionStartIndex ?? 5;
  const questionHeaders = headers.slice(startIdx);
  const used = new Set();
  const map = {};

  schema.axes.forEach(ax =>
    ax.questions.forEach(q => {
      let bestIdx = null, bestScore = 0;
      questionHeaders.forEach((h, i) => {
        if (used.has(i)) return;
        const score = similarity(h, q.text);
        if (score > bestScore) { bestScore = score; bestIdx = i; }
      });
      if (bestIdx !== null && bestScore >= 0.35) {
        map[q.id] = startIdx + bestIdx;
        used.add(bestIdx);
      } else {
        // Positional fallback
        const globalSeq = q.seq - 1;
        map[q.id] = startIdx + globalSeq;
      }
    })
  );
  return map;
}

// ── Response value parser ─────────────────────────────────────────────────────
function parseResponse5(v) {
  if (v == null) return null;
  const s = String(v).trim();
  const m = s.match(/\((\d)\)/);
  if (m) return String(m[1]);
  if (/^[1-5]$/.test(s)) return s;
  // Fallback: plain Arabic 5-point labels with no leading "(digit)" code —
  // some Google Forms exports (and custom surveys built from them) write
  // "أوافق بشدة" instead of "(5) أوافق بشدة". Order matters: check the
  // "بشدة" (strongly) and "لا" (negation) variants before the bare ones.
  const t = normalize(s);
  if (!t) return null;
  if (t.includes("لا") && t.includes("بشده")) return "1";
  if (t.includes("لا") && (t.includes("اوافق") || t.includes("agree"))) return "2";
  if (t.includes("محايد") || t.includes("neutral")) return "3";
  if (t.includes("بشده") && (t.includes("اوافق") || t.includes("موافق"))) return "5";
  if (t.includes("اوافق") || t.includes("موافق") || t === "agree") return "4";
  return null;
}

function parseResponse3(v) {
  if (v == null) return null;
  const t = normalize(v);
  if (!t) return null;
  if (t.includes("لا") && (t.includes("اوافق") || t.includes("agree"))) return "disagree";
  if (t === "disagree" || t.includes("disagree")) return "disagree";
  if (t.includes("محايد") || t.includes("neutral")) return "neutral";
  if (t.includes("اوافق") || t.includes("موافق") || t === "agree") return "agree";
  return null;
}

function parseResponse(v, schema) {
  return schema.scale.type === "likert-5" ? parseResponse5(v) : parseResponse3(v);
}

// ── Interpretation ────────────────────────────────────────────────────────────
function interpret(pct, schema) {
  const rules = schema.interpretation;
  for (const rule of rules) {
    if (pct >= rule.min) return { label: rule.label, tier: rule.tier, color: rule.color };
  }
  return rules[rules.length - 1];
}

// ── Metadata column detection ─────────────────────────────────────────────────
function findCol(headers, keywords) {
  for (let i = 0; i < headers.length; i++) {
    const h = normalize(String(headers[i] ?? ""));
    if (keywords.some(k => h.includes(normalize(k)))) return i;
  }
  return null;
}

// Collapse Arabic spelling/whitespace variants into one bucket while preserving
// the first-seen original label for display.
function bucketKey(v) {
  return normalize(String(v ?? "")).replace(/\s+/g, " ").trim();
}

function countByCol(rows, colIdx) {
  const counts = {};
  const labels = {};
  if (colIdx == null) return counts;
  for (const row of rows) {
    const v = row[colIdx];
    if (v == null) continue;
    const orig = String(v).trim();
    if (!orig) continue;
    const key = bucketKey(orig);
    if (!key) continue;
    if (!(key in labels)) labels[key] = orig;
    counts[key] = (counts[key] ?? 0) + 1;
  }
  // Re-emit with the human-readable label as the key
  const out = {};
  for (const key of Object.keys(counts)) out[labels[key]] = counts[key];
  return out;
}

function crossTab(rows, rowColIdx, colColIdx) {
  if (rowColIdx == null || colColIdx == null) return null;
  const matrix = {};
  const rowLabel = {};
  const colLabel = {};
  for (const row of rows) {
    const r = row[rowColIdx];
    const c = row[colColIdx];
    if (r == null || c == null) continue;
    const rOrig = String(r).trim();
    const cOrig = String(c).trim();
    if (!rOrig || !cOrig) continue;
    const rk = bucketKey(rOrig);
    const ck = bucketKey(cOrig);
    if (!rk || !ck) continue;
    if (!(rk in rowLabel)) rowLabel[rk] = rOrig;
    if (!(ck in colLabel)) colLabel[ck] = cOrig;
    matrix[rk] ??= {};
    matrix[rk][ck] = (matrix[rk][ck] ?? 0) + 1;
  }
  const rKeys = Object.keys(rowLabel);
  const cKeys = Object.keys(colLabel);
  if (!rKeys.length || !cKeys.length) return null;
  rKeys.sort((a, b) => rowLabel[a].localeCompare(rowLabel[b], "ar"));
  cKeys.sort((a, b) => colLabel[a].localeCompare(colLabel[b], "ar"));
  const rows_ = rKeys.map(k => rowLabel[k]);
  const cols_ = cKeys.map(k => colLabel[k]);
  const displayMatrix = {};
  rKeys.forEach(rk => {
    displayMatrix[rowLabel[rk]] = {};
    cKeys.forEach(ck => {
      displayMatrix[rowLabel[rk]][colLabel[ck]] = matrix[rk]?.[ck] ?? 0;
    });
  });
  const rowTotals = Object.fromEntries(rows_.map(r =>
    [r, cols_.reduce((s, c) => s + (displayMatrix[r]?.[c] ?? 0), 0)]));
  const colTotals = Object.fromEntries(cols_.map(c =>
    [c, rows_.reduce((s, r) => s + (displayMatrix[r]?.[c] ?? 0), 0)]));
  const grandTotal = rows_.reduce((s, r) => s + rowTotals[r], 0);
  return { rows: rows_, cols: cols_, matrix: displayMatrix, rowTotals, colTotals, grandTotal };
}

// ── Deduplication ────────────────────────────────────────────────────────────
function dedup(rows, emailIdx) {
  if (emailIdx == null) return rows;
  const seen = new Map();
  const noEmail = [];
  for (const row of rows) {
    const email = String(row[emailIdx] ?? "").trim().toLowerCase();
    if (email) seen.set(email, row);
    else noEmail.push(row);
  }
  return [...noEmail, ...seen.values()];
}

// ── Data preparation (extracted so the UI can hold deduped rows) ─────────────
export function prepareData(rows, schema) {
  const headers = rows[0];
  const meta = schema.metadata ?? {};
  const emailIdx   = findCol(headers, meta.emailCol ?? []);
  const deptIdx    = findCol(headers, meta.departmentCol ?? []);
  const degreeIdx  = findCol(headers, meta.degreeCol ?? []);

  const dataRows = dedup(rows.slice(1), emailIdx);
  return { headers, dataRows, metaCols: { emailIdx, deptIdx, degreeIdx } };
}

// ── Main analysis function ────────────────────────────────────────────────────
export function analyze(rows, schema) {
  const { headers, dataRows, metaCols } = prepareData(rows, schema);
  return analyzeRows(headers, dataRows, schema, metaCols);
}

export function analyzeRows(headers, dataRows, schema, metaCols) {
  if (!metaCols) {
    const meta = schema.metadata ?? {};
    metaCols = {
      emailIdx:  findCol(headers, meta.emailCol      ?? []),
      deptIdx:   findCol(headers, meta.departmentCol ?? []),
      degreeIdx: findCol(headers, meta.degreeCol     ?? []),
    };
  }
  const { deptIdx, degreeIdx } = metaCols;
  const n = dataRows.length;
  const colMap = mapColumns(headers, schema);
  const scaleCodes = schema.scale.values.map(v => v.code);
  const agreementCodes = new Set(schema.scale.agreementCodes);

  let globalSeq = 1;
  const axes = schema.axes.map(axDef => {
    const questions = axDef.questions.map(qDef => {
      const colIdx = colMap[qDef.id];
      const counts = Object.fromEntries(scaleCodes.map(c => [c, 0]));
      let total = 0;

      for (const row of dataRows) {
        const code = parseResponse(colIdx != null ? row[colIdx] : null, schema);
        if (code && code in counts) { counts[code]++; total++; }
      }

      const pcts = Object.fromEntries(
        scaleCodes.map(c => [c, total ? +(counts[c] / total * 100).toFixed(1) : 0])
      );

      // Mean only for 5-point (numeric scores)
      let mean = null;
      if (schema.scale.type === "likert-5") {
        const sum = scaleCodes.reduce((s, c) => s + counts[c] * schema.scale.values.find(v => v.code === c).score, 0);
        mean = total ? +(sum / total).toFixed(2) : 0;
      }

      const agreePct = total
        ? +([...agreementCodes].reduce((s, c) => s + (counts[c] ?? 0), 0) / total * 100).toFixed(1)
        : 0;

      const interp = schema.scale.type === "likert-5"
        ? interpret(mean ?? 0, schema)
        : interpret(agreePct, schema);

      return {
        seq: globalSeq++,
        id: qDef.id,
        text: qDef.text,
        counts,
        pcts,
        mean,
        agreePct,
        direction: interp.label,
        tier: interp.tier,
        color: interp.color,
        total,
        colIdx,
      };
    });

    // Axis aggregate
    const totalResponses = questions.reduce((s, q) => s + q.total, 0);
    const totalAgree = questions.reduce(
      (s, q) => s + [...agreementCodes].reduce((a, c) => a + (q.counts[c] ?? 0), 0), 0
    );
    const axisAgreePct = totalResponses ? +(totalAgree / totalResponses * 100).toFixed(1) : 0;

    let axisMean = null;
    if (schema.scale.type === "likert-5") {
      const means = questions.filter(q => q.mean !== null && q.total > 0).map(q => q.mean);
      axisMean = means.length ? +(means.reduce((s, m) => s + m, 0) / means.length).toFixed(2) : 0;
    }

    const axisInterp = schema.scale.type === "likert-5"
      ? interpret(axisMean ?? 0, schema)
      : interpret(axisAgreePct, schema);

    return {
      id: axDef.id,
      name: axDef.name,
      recommendation: axDef.recommendation ?? null,
      questions,
      axisMean,
      axisAgreePct,
      direction: axisInterp.label,
      tier: axisInterp.tier,
      color: axisInterp.color,
    };
  });

  // Overall
  const allAxesAgreePct = axes.filter(a => a.questions.some(q => q.total > 0));
  const overallAgreePct = allAxesAgreePct.length
    ? +(allAxesAgreePct.reduce((s, a) => s + a.axisAgreePct, 0) / allAxesAgreePct.length).toFixed(1)
    : 0;

  let overallMean = null;
  if (schema.scale.type === "likert-5") {
    const means = axes.filter(a => a.axisMean !== null && a.axisMean > 0).map(a => a.axisMean);
    overallMean = means.length ? +(means.reduce((s, m) => s + m, 0) / means.length).toFixed(2) : 0;
  }

  const overallVal = schema.scale.type === "likert-5" ? (overallMean ?? 0) : overallAgreePct;
  const overallInterp = interpret(overallVal, schema);

  return {
    schemaId: schema.id,
    schemaLabel: schema.label,
    scaleType: schema.scale.type,
    n,
    axes,
    overallMean,
    overallAgreePct,
    overallDirection: overallInterp.label,
    overallTier: overallInterp.tier,
    byDegree: countByCol(dataRows, degreeIdx),
    byDepartment: countByCol(dataRows, deptIdx),
    crossDegreeByDept: crossTab(dataRows, degreeIdx, deptIdx),
    totalQuestions: axes.reduce((s, a) => s + a.questions.length, 0),
  };
}

// ── Excel loader ──────────────────────────────────────────────────────────────
export function readExcel(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
}

// ── Comparison engine ─────────────────────────────────────────────────────────
export function buildComparison(slots) {
  // slots: [{ year, result }]  — result from analyze()
  if (!slots.length) return null;
  const baseAxes = slots[0].result.axes;

  const axes = baseAxes.map(ax => ({
    id: ax.id,
    name: ax.name,
    years: slots.map(s => {
      const found = s.result.axes.find(a => a.id === ax.id);
      return {
        year: s.year,
        agreePct: found?.axisAgreePct ?? null,
        axisMean: found?.axisMean ?? null,
        direction: found?.direction ?? "—",
      };
    }),
  }));

  // Trend per axis (first→last delta)
  axes.forEach(ax => {
    const vals = ax.years.map(y => y.agreePct).filter(v => v !== null);
    if (vals.length >= 2) {
      ax.delta = +(vals[vals.length - 1] - vals[0]).toFixed(1);
      ax.trend = ax.delta >= 2 ? "تحسن" : ax.delta <= -2 ? "تراجع" : "استقرار";
    } else {
      ax.delta = null;
      ax.trend = "—";
    }
  });

  const overall = slots.map(s => ({
    year: s.year,
    agreePct: s.result.overallAgreePct,
    mean: s.result.overallMean,
    direction: s.result.overallDirection,
  }));

  return { axes, overall, scaleType: slots[0].result.scaleType };
}
