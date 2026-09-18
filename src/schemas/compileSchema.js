// Compiles the minimal YAML schema format into the full runtime schema shape
// that analyze.js and buildDocx.js already consume.

const SCALE5_VALUES = [
  { code: "1", label: "لا أوافق بشدة", score: 1 },
  { code: "2", label: "لا أوافق",      score: 2 },
  { code: "3", label: "محايد",          score: 3 },
  { code: "4", label: "أوافق",          score: 4 },
  { code: "5", label: "أوافق بشدة",     score: 5 },
];

const SCALE3_VALUES = [
  { code: "agree",    label: "أوافق",    score: 3 },
  { code: "neutral",  label: "محايد",    score: 2 },
  { code: "disagree", label: "لا أوافق", score: 1 },
];

const INTERP5 = [
  { min: 4.5, label: "أوافق بشدة",  tier: "excellent", color: "0d6e3a" },
  { min: 3.5, label: "أوافق",       tier: "good",      color: "1a5276" },
  { min: 2.5, label: "محايد",       tier: "neutral",   color: "784212" },
  { min: 1.0, label: "لا أوافق",   tier: "low",       color: "922b21" },
];

const INTERP3 = [
  { min: 85, label: "موافقة قوية",  tier: "excellent", color: "0d6e3a" },
  { min: 70, label: "موافقة",       tier: "good",      color: "1a5276" },
  { min: 50, label: "محايد",        tier: "neutral",   color: "784212" },
  { min: 0,  label: "عدم الموافقة", tier: "low",       color: "922b21" },
];

function buildScale(n, customValues) {
  if (n === 5) {
    return {
      type: "likert-5",
      values: Array.isArray(customValues) && customValues.length === 5 ? customValues : SCALE5_VALUES,
      agreementCodes: ["4", "5"],
      tokenPattern: "\\((\\d)\\)",
    };
  }
  return {
    type: "likert-3",
    values: SCALE3_VALUES,
    agreementCodes: ["agree"],
    detectFn: "arabic3point",
  };
}

function buildMeta(raw) {
  const m = raw || {};
  return {
    timestampCol:  m.timestamp  || ["Timestamp"],
    ...(m.level   ? { levelCol:      Array.isArray(m.level)   ? m.level   : [m.level]   } : {}),
    ...(m.email   ? { emailCol:      Array.isArray(m.email)   ? m.email   : [m.email]   } : {}),
    ...(m.degree  ? { degreeCol:     Array.isArray(m.degree)  ? m.degree  : [m.degree]  } : {}),
    ...(m.dept    ? { departmentCol: Array.isArray(m.dept)    ? m.dept    : [m.dept]    } : {}),
    ...(m.program ? { programCol:    Array.isArray(m.program) ? m.program : [m.program] } : {}),
    freeTextCols:  m.freetext || [],
  };
}

function buildAxes(raw, isLikert5) {
  // Determine axes source — flat `questions:` wraps into a single anonymous axis
  const axesSrc = raw.axes
    ? raw.axes
    : [{ name: raw.name, questions: raw.questions, recommendation: raw.recommendation }];

  let globalSeq = 1;
  let autoColIndex = 0;

  return axesSrc.map((ax, ai) => {
    const axId = `ax${String(ai + 1).padStart(2, "0")}`;

    const questions = ax.questions.map((q) => {
      const text = typeof q === "string" ? q : q.text;
      const qId  = `q${String(globalSeq).padStart(2, "0")}`;

      let colIndex;
      if (isLikert5) {
        if (typeof q === "object" && q.col != null) {
          // Explicit override: use it and advance the counter past it
          colIndex = q.col;
          autoColIndex = q.col + 1;
        } else {
          colIndex = autoColIndex++;
        }
      }

      const out = {
        id: qId,
        seq: globalSeq,
        text,
        ...(isLikert5 ? { colIndex } : {}),
      };
      globalSeq++;
      return out;
    });

    return {
      id: axId,
      name: ax.name,
      ...(ax.recommendation ? { recommendation: ax.recommendation } : {}),
      questions,
    };
  });
}

export function compileSchema(raw) {
  const isLikert5 = raw.scale === 5;
  const scale     = buildScale(raw.scale, raw.scaleValues);
  const meta      = buildMeta(raw.meta);
  const axes      = buildAxes(raw, isLikert5);

  // For likert-3: compute questionStartIndex from declared meta columns
  let questionStartIndex = raw.questionStartIndex;
  if (!isLikert5 && questionStartIndex == null) {
    const m = raw.meta || {};
    let count = 1; // Timestamp is always present
    if (m.email)   count++;
    if (m.degree)  count++;
    if (m.dept)    count++;
    if (m.program) count++;
    questionStartIndex = count;
  }

  return {
    id:          raw.id,
    label:       raw.name,
    labelEn:     raw.nameEn  || "",
    fileHints:   raw.hints   || [],
    icon:        raw.icon    || null,
    desc:        raw.desc    || null,
    programs:    raw.programs || [],
    isFlat:      raw.flat === true || !raw.axes,
    reportTexts: raw.reportTexts || null,
    reportSections: raw.reportSections || null,
    scale,
    metadata:    meta,
    ...(!isLikert5 ? { questionStartIndex } : {}),
    interpretation: raw.interpretation || (isLikert5 ? INTERP5 : INTERP3),
    axes,
  };
}
