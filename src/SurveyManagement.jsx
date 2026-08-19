import { useState, useRef } from "react";
import {
  SCALE_TYPES, QUESTION_TYPES, SURVEY_STATUSES,
  createSurvey, createSection, createQuestion, createMetadata,
  loadCustomSurveys, saveCustomSurvey, deleteCustomSurvey,
  duplicateCustomSurvey, setCustomSurveyStatus, toRuntimeSchemaShape,
  importSurveyStructureFromRows,
  buildSurveyTemplateBlob, importSurveyFromTemplateArrayBuffer,
  buildSurveysBackupBlob, importSurveysBackup,
} from "./engine/customSurveyModel.js";
import { readExcel } from "./engine/analyze.js";

// ── small helpers ─────────────────────────────────────────────────────────────
function moveItem(arr, idx, dir) {
  const next = [...arr];
  const target = idx + dir;
  if (target < 0 || target >= next.length) return arr;
  [next[idx], next[target]] = [next[target], next[idx]];
  return next;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function countQuestions(survey) {
  return (survey.sections ?? []).reduce((sum, s) => sum + (s.questions?.length ?? 0), 0);
}

function fmtDate(iso) {
  if (!iso) return "—";
  try { return new Date(iso).toLocaleDateString("ar-EG", { year: "numeric", month: "short", day: "numeric" }); }
  catch { return iso; }
}

// ── Status badge ──────────────────────────────────────────────────────────────
function StatusBadge({ status }) {
  const isActive = status === "active";
  return (
    <span style={{
      display: "inline-block", padding: "4px 12px", borderRadius: 20, fontSize: 11, fontWeight: 700,
      background: isActive ? "rgba(26,188,156,.18)" : "rgba(255,193,7,.15)",
      color: isActive ? "#1abc9c" : "#ffd54f",
      border: `1px solid ${isActive ? "rgba(26,188,156,.4)" : "rgba(255,193,7,.35)"}`,
    }}>
      {isActive ? "● نشط" : "● مسودة"}
    </span>
  );
}

// ── Confirm dialog (lightweight inline modal) ─────────────────────────────────
function ConfirmDialog({ title, message, onConfirm, onCancel }) {
  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(5,12,22,.72)", zIndex: 500,
      display: "flex", alignItems: "center", justifyContent: "center", padding: 20,
    }} onClick={e => { if (e.target === e.currentTarget) onCancel(); }}>
      <div className="card" style={{ maxWidth: 420, padding: 28, textAlign: "center" }}>
        <div style={{ fontSize: 32, marginBottom: 10 }}>⚠️</div>
        <div style={{ color: "#fff", fontWeight: 900, fontSize: 17, marginBottom: 8 }}>{title}</div>
        <div style={{ color: "rgba(255,255,255,.6)", fontSize: 13, marginBottom: 22 }}>{message}</div>
        <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
          <button className="btn btn-ghost btn-sm" onClick={onCancel}>إلغاء</button>
          <button className="btn btn-danger btn-sm" onClick={onConfirm}>تأكيد الحذف</button>
        </div>
      </div>
    </div>
  );
}

// ── List view ─────────────────────────────────────────────────────────────────
function SurveyListView({
  surveys, onCreate, onEdit, onDelete, onDuplicate, onToggleStatus,
  onDownloadTemplate, onImportTemplate, onExportBackup, onImportBackup,
}) {
  const [confirmId, setConfirmId] = useState(null);
  const [toolbarMsg, setToolbarMsg] = useState("");
  const templateFileRef = useRef();
  const backupFileRef = useRef();

  const handleTemplateFile = async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      await onImportTemplate(file);
      setToolbarMsg(`✓ تم إنشاء استبيان جديد من القالب «${file.name}» — افتحه من القائمة لمراجعته.`);
    } catch (err) {
      setToolbarMsg("✗ تعذّر استيراد القالب: " + err.message);
    }
  };

  const handleBackupFile = async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      const text = await file.text();
      const count = onImportBackup(text);
      setToolbarMsg(`✓ تم استرجاع النسخة الاحتياطية — العدد الإجمالي الآن: ${count} استبيان.`);
    } catch (err) {
      setToolbarMsg("✗ تعذّر استيراد النسخة الاحتياطية: " + err.message);
    }
  };

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
        <div>
          <div style={{ color: "#fff", fontWeight: 900, fontSize: 20 }}>🗂️ إدارة الاستبيانات</div>
          <div style={{ color: "rgba(255,255,255,.45)", fontSize: 12.5, marginTop: 4 }}>
            إنشاء وتعديل تعريفات الاستبيانات دون كتابة كود — هذا النظام يعمل بشكل مستقل عن الاستبيانات الحالية ولا يؤثر عليها.
          </div>
        </div>
        <button className="btn btn-primary" onClick={onCreate}>+ استبيان جديد</button>
      </div>

      <div style={{
        display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center",
        background: "rgba(255,255,255,.04)", border: "1px solid rgba(255,255,255,.08)",
        borderRadius: 10, padding: 12, marginBottom: 20,
      }}>
        <span style={{ color: "rgba(255,255,255,.45)", fontSize: 11.5, fontWeight: 700 }}>قالب Excel:</span>
        <button className="btn btn-ghost btn-sm" onClick={onDownloadTemplate}>⬇ تحميل القالب</button>
        <input ref={templateFileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleTemplateFile} />
        <button className="btn btn-blue btn-sm" onClick={() => templateFileRef.current.click()}>📤 استيراد استبيان من قالب</button>

        <span style={{ width: 1, height: 20, background: "rgba(255,255,255,.12)", margin: "0 6px" }} />

        <span style={{ color: "rgba(255,255,255,.45)", fontSize: 11.5, fontWeight: 700 }}>نسخة احتياطية:</span>
        <button className="btn btn-ghost btn-sm" onClick={onExportBackup}>⬇ تصدير نسخة احتياطية (JSON)</button>
        <input ref={backupFileRef} type="file" accept=".json" style={{ display: "none" }} onChange={handleBackupFile} />
        <button className="btn btn-ghost btn-sm" onClick={() => backupFileRef.current.click()}>⬆ استيراد نسخة احتياطية</button>
      </div>

      {toolbarMsg && (
        <div style={{ color: toolbarMsg.startsWith("✗") ? "#ff6b5b" : "#1abc9c", fontSize: 12.5, marginBottom: 14 }}>{toolbarMsg}</div>
      )}

      {surveys.length === 0 ? (
        <div className="card" style={{ padding: 48, textAlign: "center" }}>
          <div style={{ fontSize: 40, marginBottom: 10 }}>📭</div>
          <div style={{ color: "rgba(255,255,255,.6)", fontSize: 14 }}>
            لا توجد استبيانات مُعرّفة بعد. اضغط "+ استبيان جديد" للبدء.
          </div>
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {surveys.map(s => (
            <div key={s.id} className="card" style={{ padding: 18, display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
              <div style={{ minWidth: 220, flex: 1 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
                  <span style={{ color: "#fff", fontWeight: 800, fontSize: 15.5 }}>{s.name || "(بدون اسم)"}</span>
                  <StatusBadge status={s.status} />
                  <span style={{ color: "rgba(255,255,255,.4)", fontSize: 11.5 }}>v{s.version}</span>
                </div>
                <div style={{ color: "rgba(255,255,255,.5)", fontSize: 12, display: "flex", gap: 14, flexWrap: "wrap" }}>
                  <span>{SCALE_TYPES[s.scaleType]?.label ?? s.scaleType}</span>
                  <span>{(s.sections ?? []).length} محاور</span>
                  <span>{countQuestions(s)} سؤال</span>
                  {s.surveyType && <span>النوع: {s.surveyType}</span>}
                  <span>آخر تحديث: {fmtDate(s.updatedAt)}</span>
                </div>
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button className="btn btn-ghost btn-sm" onClick={() => onEdit(s)}>✎ تعديل</button>
                <button className="btn btn-ghost btn-sm" onClick={() => onDuplicate(s.id)}>⧉ نسخ</button>
                <button className="btn btn-ghost btn-sm" onClick={() => onToggleStatus(s.id, s.status === "active" ? "draft" : "active")}>
                  {s.status === "active" ? "⏸ تعطيل" : "▶ تفعيل"}
                </button>
                <button className="btn btn-danger btn-sm" onClick={() => setConfirmId(s.id)}>🗑 حذف</button>
              </div>
            </div>
          ))}
        </div>
      )}

      {confirmId && (
        <ConfirmDialog
          title="حذف الاستبيان؟"
          message="سيتم حذف تعريف هذا الاستبيان بشكل نهائي. لا يمكن التراجع عن هذا الإجراء."
          onCancel={() => setConfirmId(null)}
          onConfirm={() => { onDelete(confirmId); setConfirmId(null); }}
        />
      )}
    </div>
  );
}

// ── Question row editor ───────────────────────────────────────────────────────
function QuestionRow({ q, index, total, onChange, onDelete, onMove, sectionId, allSections, onMoveToSection }) {
  const [moveTarget, setMoveTarget] = useState("");
  const otherSections = allSections.filter(s => s.id !== sectionId);

  return (
    <div style={{
      background: "rgba(255,255,255,.04)", border: "1px solid rgba(255,255,255,.08)",
      borderRadius: 10, padding: 12, marginBottom: 8,
    }}>
      <div style={{ display: "flex", gap: 8, alignItems: "flex-start" }}>
        <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 8 }}>
          <input className="input" placeholder="نص السؤال" value={q.text}
            onChange={e => onChange({ ...q, text: e.target.value })} />
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <input className="input" style={{ flex: "1 1 160px" }} placeholder="اسم العمود في Excel"
              value={q.excelColumn} onChange={e => onChange({ ...q, excelColumn: e.target.value })} />
            <select className="input" style={{ flex: "1 1 140px" }} value={q.type}
              onChange={e => onChange({ ...q, type: e.target.value })}>
              {QUESTION_TYPES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
            </select>
            <input className="input" style={{ flex: "1 1 120px" }} placeholder="الفئة (اختياري)"
              value={q.category ?? ""} onChange={e => onChange({ ...q, category: e.target.value })} />
            <input className="input" type="number" style={{ flex: "1 1 90px" }} placeholder="الوزن"
              value={q.weight ?? ""} onChange={e => onChange({ ...q, weight: e.target.value === "" ? null : Number(e.target.value) })} />
            <button
              className={`btn btn-sm ${q.required ? "btn-primary" : "btn-ghost"}`}
              onClick={() => onChange({ ...q, required: !q.required })}
              title="إلزامي؟"
            >
              {q.required ? "✓ إلزامي" : "اختياري"}
            </button>
          </div>
          {otherSections.length > 0 && (
            <div style={{ display: "flex", gap: 8 }}>
              <select className="input" style={{ flex: 1 }} value={moveTarget} onChange={e => setMoveTarget(e.target.value)}>
                <option value="">نقل هذا السؤال إلى محور آخر...</option>
                {otherSections.map(s => <option key={s.id} value={s.id}>{s.name || "(بدون اسم)"}</option>)}
              </select>
              <button
                className="btn btn-ghost btn-sm"
                disabled={!moveTarget}
                onClick={() => { onMoveToSection(moveTarget); setMoveTarget(""); }}
              >
                ➡ نقل
              </button>
            </div>
          )}
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <button className="btn btn-ghost btn-sm" disabled={index === 0} onClick={() => onMove(-1)} title="تحريك للأعلى" aria-label="تحريك السؤال للأعلى">↑</button>
          <button className="btn btn-ghost btn-sm" disabled={index === total - 1} onClick={() => onMove(1)} title="تحريك للأسفل" aria-label="تحريك السؤال للأسفل">↓</button>
          <button className="btn btn-danger btn-sm" onClick={onDelete} title="حذف السؤال" aria-label="حذف السؤال">✕</button>
        </div>
      </div>
    </div>
  );
}

// ── Section editor ────────────────────────────────────────────────────────────
function SectionBlock({ section, index, total, onChange, onDelete, onMove, allSections, onMoveQuestionToSection }) {
  const updateQuestion = (qIdx, q) => {
    const questions = [...section.questions];
    questions[qIdx] = q;
    onChange({ ...section, questions });
  };
  const deleteQuestion = (qIdx) => {
    onChange({ ...section, questions: section.questions.filter((_, i) => i !== qIdx) });
  };
  const moveQuestion = (qIdx, dir) => {
    onChange({ ...section, questions: moveItem(section.questions, qIdx, dir) });
  };
  const addQuestion = () => {
    onChange({ ...section, questions: [...section.questions, createQuestion()] });
  };

  return (
    <div className="card" style={{ padding: 18, marginBottom: 14 }}>
      <div style={{ display: "flex", gap: 8, alignItems: "flex-start", marginBottom: 12 }}>
        <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 8 }}>
          <input className="input" placeholder={`اسم المحور ${index + 1}`} value={section.name}
            onChange={e => onChange({ ...section, name: e.target.value })}
            style={{ fontWeight: 800 }} />
          <input className="input" placeholder="وصف المحور (اختياري)" value={section.description ?? ""}
            onChange={e => onChange({ ...section, description: e.target.value })} />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <button className="btn btn-ghost btn-sm" disabled={index === 0} onClick={() => onMove(-1)} title="تحريك المحور للأعلى" aria-label="تحريك المحور للأعلى">↑</button>
          <button className="btn btn-ghost btn-sm" disabled={index === total - 1} onClick={() => onMove(1)} title="تحريك المحور للأسفل" aria-label="تحريك المحور للأسفل">↓</button>
          <button className="btn btn-danger btn-sm" onClick={onDelete} title="حذف المحور" aria-label="حذف المحور">✕</button>
        </div>
      </div>

      <div style={{ color: "rgba(255,255,255,.45)", fontSize: 11.5, marginBottom: 8, fontWeight: 700 }}>
        الأسئلة ({section.questions.length})
      </div>
      {section.questions.map((q, qi) => (
        <QuestionRow
          key={q.id} q={q} index={qi} total={section.questions.length}
          onChange={nq => updateQuestion(qi, nq)}
          onDelete={() => deleteQuestion(qi)}
          onMove={dir => moveQuestion(qi, dir)}
          sectionId={section.id}
          allSections={allSections}
          onMoveToSection={toSectionId => onMoveQuestionToSection(qi, toSectionId)}
        />
      ))}
      <button className="btn btn-ghost btn-sm" onClick={addQuestion}>+ إضافة سؤال</button>
    </div>
  );
}

// ── Preview tab ───────────────────────────────────────────────────────────────
function PreviewTab({ survey }) {
  const compiled = toRuntimeSchemaShape(survey);
  return (
    <div className="card" style={{ padding: 24 }}>
      <div style={{ color: "#fff", fontWeight: 900, fontSize: 19, marginBottom: 4 }}>{compiled.label}</div>
      <div style={{ color: "rgba(255,255,255,.5)", fontSize: 12.5, marginBottom: 4 }}>{survey.description}</div>
      <div style={{ display: "flex", gap: 14, flexWrap: "wrap", color: "rgba(255,255,255,.45)", fontSize: 12, marginBottom: 18 }}>
        <span>الإصدار: {compiled.version}</span>
        <span><StatusBadge status={compiled.status} /></span>
        <span>{SCALE_TYPES[survey.scaleType]?.label}</span>
        <span>{compiled.axes.length} محاور · {countQuestions(survey)} سؤال</span>
      </div>

      {compiled.axes.length === 0 && (
        <div style={{ color: "rgba(255,255,255,.4)", fontSize: 13 }}>لا توجد محاور بعد. أضف محوراً من تبويب "المحاور".</div>
      )}

      {compiled.axes.map((ax, ai) => (
        <div key={ax.id} style={{ marginBottom: 18 }}>
          <div style={{ color: "#1abc9c", fontWeight: 800, fontSize: 14.5, marginBottom: 8, borderBottom: "1px solid rgba(26,188,156,.25)", paddingBottom: 6 }}>
            {ai + 1}. {ax.name}
          </div>
          {ax.questions.length === 0 ? (
            <div style={{ color: "rgba(255,255,255,.35)", fontSize: 12.5 }}>لا توجد أسئلة في هذا المحور.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {ax.questions.map(q => (
                <div key={q.id} style={{ display: "flex", gap: 10, alignItems: "center", fontSize: 13, color: "rgba(255,255,255,.8)" }}>
                  <span style={{ color: "rgba(255,255,255,.35)", minWidth: 26 }}>{q.seq}.</span>
                  <span style={{ flex: 1 }}>{q.text || "(بدون نص)"}</span>
                  {q.excelColumn && <span style={{ fontSize: 10.5, color: "rgba(255,255,255,.4)" }}>[{q.excelColumn}]</span>}
                  <span style={{ fontSize: 10.5, color: "#1abc9c" }}>{QUESTION_TYPES.find(t => t.id === q.type)?.label ?? q.type}</span>
                  {q.required && <span style={{ fontSize: 10.5, color: "#ffd54f" }}>إلزامي</span>}
                </div>
              ))}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

// ── comma-separated-list helpers for metadata column fields ───────────────────
function listToText(arr) { return (arr ?? []).join("، "); }
function textToList(text) { return text.split(/[،,]/).map(s => s.trim()).filter(Boolean); }

// ── Import-from-Excel panel (General step) ────────────────────────────────────
function ImportFromExcel({ onImported }) {
  const fileRef = useRef();
  const [summary, setSummary] = useState(null);
  const [error, setError] = useState("");

  const handleFile = async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setError("");
    try {
      const buf = await file.arrayBuffer();
      const rows = readExcel(buf);
      const imported = importSurveyStructureFromRows(rows);
      if (!imported || imported.questions.length === 0) {
        setError("لم يتم العثور على أعمدة أسئلة صالحة في هذا الملف.");
        return;
      }
      onImported(imported);
      const metaCount = Object.values(imported.metadata).reduce((s, a) => s + a.length, 0);
      setSummary({ fileName: file.name, questionCount: imported.questions.length, metaCount, scaleType: imported.scaleType });
    } catch (err) {
      setError("تعذّر قراءة الملف: " + err.message);
    }
  };

  return (
    <div style={{
      background: "rgba(26,188,156,.06)", border: "1px dashed rgba(26,188,156,.35)",
      borderRadius: 12, padding: 16, marginBottom: 18,
    }}>
      <div style={{ color: "#fff", fontWeight: 800, fontSize: 13.5, marginBottom: 4 }}>📥 استيراد تلقائي من ملف Excel (اختياري)</div>
      <div style={{ color: "rgba(255,255,255,.5)", fontSize: 12, marginBottom: 10 }}>
        يكتشف النظام الأسئلة من صفوف العناوين تلقائياً، ويستثني أعمدة المعلومات العامة (الاسم، البريد، الوظيفة/الدرجة، القسم/التخصص، الملاحظات).
        كل الأسئلة المكتشفة توضع في محور واحد جديد — يمكنك بعد ذلك تقسيمها على عدة محاور من تبويب "المحاور".
      </div>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={handleFile} />
      <button className="btn btn-blue btn-sm" onClick={() => fileRef.current.click()}>اختر ملف Excel…</button>
      {summary && (
        <div style={{ color: "#1abc9c", fontSize: 12, marginTop: 10 }}>
          ✓ تم استيراد {summary.questionCount} سؤال و {summary.metaCount} عمود معلومات عامة من «{summary.fileName}» (مقياس مكتشف: {SCALE_TYPES[summary.scaleType]?.label}).
        </div>
      )}
      {error && <div style={{ color: "#ff6b5b", fontSize: 12, marginTop: 10 }}>{error}</div>}
    </div>
  );
}

const EDITOR_STEPS = [
  { id: "general", label: "عام" },
  { id: "sections", label: "المحاور" },
  { id: "preview", label: "معاينة" },
];

// ── Editor view (General → Sections → Preview, step-by-step) ──────────────────
function SurveyEditorView({ survey: initial, onSave, onCancel }) {
  const [survey, setSurvey] = useState({ metadata: createMetadata(), ...initial });
  const [tab, setTab] = useState("general");

  const updateSection = (idx, sec) => {
    const sections = [...survey.sections];
    sections[idx] = sec;
    setSurvey({ ...survey, sections });
  };
  const deleteSection = (idx) => setSurvey({ ...survey, sections: survey.sections.filter((_, i) => i !== idx) });
  const moveSection = (idx, dir) => setSurvey({ ...survey, sections: moveItem(survey.sections, idx, dir) });
  const addSection = () => setSurvey({ ...survey, sections: [...survey.sections, createSection()] });

  const moveQuestionToSection = (fromIdx, qIdx, toSectionId) => {
    setSurvey(s => {
      const sections = s.sections.map(sec => ({ ...sec, questions: [...sec.questions] }));
      const toIdx = sections.findIndex(sec => sec.id === toSectionId);
      if (toIdx === -1 || toIdx === fromIdx) return s;
      const [moved] = sections[fromIdx].questions.splice(qIdx, 1);
      sections[toIdx].questions.push(moved);
      return { ...s, sections };
    });
  };

  const updateMetaList = (field, text) => setSurvey({ ...survey, metadata: { ...survey.metadata, [field]: textToList(text) } });

  const handleImported = (imported) => {
    setSurvey(s => ({
      ...s,
      scaleType: imported.scaleType,
      metadata: imported.metadata,
      sections: [...s.sections, createSection({ name: `محور مستورد ${s.sections.length + 1}`, questions: imported.questions })],
    }));
  };

  const canSave = survey.name.trim().length > 0;
  const stepIdx = EDITOR_STEPS.findIndex(s => s.id === tab);
  const goTo = (id) => setTab(id);
  const goNext = () => goTo(EDITOR_STEPS[Math.min(stepIdx + 1, EDITOR_STEPS.length - 1)].id);
  const goBack = () => goTo(EDITOR_STEPS[Math.max(stepIdx - 1, 0)].id);

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
        <div style={{ color: "#fff", fontWeight: 900, fontSize: 19 }}>
          {survey.name || "استبيان جديد"} <span style={{ color: "rgba(255,255,255,.4)", fontWeight: 600, fontSize: 13 }}>— محرر الاستبيان</span>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="btn btn-ghost btn-sm" onClick={onCancel}>إلغاء</button>
          <button className="btn btn-primary btn-sm" disabled={!canSave} onClick={() => onSave(survey)} title="حفظ سريع من أي خطوة">💾 حفظ</button>
        </div>
      </div>

      <div style={{ display: "flex", gap: 6, marginBottom: 18, borderBottom: "1px solid rgba(255,255,255,.1)" }}>
        {EDITOR_STEPS.map((t, i) => (
          <button key={t.id} onClick={() => goTo(t.id)} style={{
            padding: "10px 20px", border: "none", cursor: "pointer", background: "transparent",
            fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 13.5,
            color: tab === t.id ? "#1abc9c" : "rgba(255,255,255,.45)",
            borderBottom: tab === t.id ? "2px solid #1abc9c" : "2px solid transparent",
          }}>
            {i + 1}. {t.label}
          </button>
        ))}
      </div>

      {tab === "general" && (
        <div>
          <ImportFromExcel onImported={handleImported} />
          <div className="card" style={{ padding: 22, display: "flex", flexDirection: "column", gap: 16, maxWidth: 640 }}>
            <div>
              <label className="label">اسم الاستبيان *</label>
              <input className="input" value={survey.name} onChange={e => setSurvey({ ...survey, name: e.target.value })} />
            </div>
            <div>
              <label className="label">وصف الاستبيان</label>
              <textarea className="textarea" rows={3} value={survey.description}
                onChange={e => setSurvey({ ...survey, description: e.target.value })} />
            </div>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <div style={{ flex: "1 1 140px" }}>
                <label className="label">الإصدار</label>
                <input className="input" value={survey.version} onChange={e => setSurvey({ ...survey, version: e.target.value })} />
              </div>
              <div style={{ flex: "1 1 140px" }}>
                <label className="label">الحالة</label>
                <select className="input" value={survey.status} onChange={e => setSurvey({ ...survey, status: e.target.value })}>
                  {SURVEY_STATUSES.map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                </select>
              </div>
            </div>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <div style={{ flex: "1 1 200px" }}>
                <label className="label">نوع الاستبيان</label>
                <input className="input" placeholder="مثال: رضا الطلاب، خريجين..." value={survey.surveyType}
                  onChange={e => setSurvey({ ...survey, surveyType: e.target.value })} />
              </div>
              <div style={{ flex: "1 1 200px" }}>
                <label className="label">نوع المقياس</label>
                <select className="input" value={survey.scaleType} onChange={e => setSurvey({ ...survey, scaleType: e.target.value })}>
                  {Object.entries(SCALE_TYPES).map(([id, sc]) => <option key={id} value={id}>{sc.label}</option>)}
                </select>
              </div>
            </div>

            <div style={{ borderTop: "1px solid rgba(255,255,255,.08)", paddingTop: 14 }}>
              <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12, fontWeight: 700, marginBottom: 10 }}>
                معلومات عامة (تُستثنى تلقائياً من الأسئلة عند التحليل)
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                <div>
                  <label className="label">أعمدة الاسم</label>
                  <input className="input" placeholder="مثال: الاسم" value={listToText(survey.metadata.nameCol)}
                    onChange={e => updateMetaList("nameCol", e.target.value)} />
                </div>
                <div>
                  <label className="label">أعمدة البريد الإلكتروني</label>
                  <input className="input" placeholder="مثال: Email" value={listToText(survey.metadata.emailCol)}
                    onChange={e => updateMetaList("emailCol", e.target.value)} />
                </div>
                <div>
                  <label className="label">أعمدة الوظيفة / الدرجة</label>
                  <input className="input" placeholder="مثال: الوظيفة، الدرجة العلمية" value={listToText(survey.metadata.degreeCol)}
                    onChange={e => updateMetaList("degreeCol", e.target.value)} />
                </div>
                <div>
                  <label className="label">أعمدة القسم / التخصص</label>
                  <input className="input" placeholder="مثال: القسم" value={listToText(survey.metadata.departmentCol)}
                    onChange={e => updateMetaList("departmentCol", e.target.value)} />
                </div>
                <div>
                  <label className="label">أعمدة نصية مفتوحة (ملاحظات/مقترحات)</label>
                  <input className="input" placeholder="مثال: مقترحات أخرى" value={listToText(survey.metadata.freeTextCols)}
                    onChange={e => updateMetaList("freeTextCols", e.target.value)} />
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {tab === "sections" && (
        <div>
          {survey.sections.map((sec, i) => (
            <SectionBlock
              key={sec.id} section={sec} index={i} total={survey.sections.length}
              onChange={s => updateSection(i, s)}
              onDelete={() => deleteSection(i)}
              onMove={dir => moveSection(i, dir)}
              allSections={survey.sections.map(s => ({ id: s.id, name: s.name }))}
              onMoveQuestionToSection={(qIdx, toSectionId) => moveQuestionToSection(i, qIdx, toSectionId)}
            />
          ))}
          <button className="btn btn-blue btn-sm" onClick={addSection}>+ إضافة محور</button>
        </div>
      )}

      {tab === "preview" && <PreviewTab survey={survey} />}

      {/* Step navigation: Next after each step, Save (distinct from the top one) on the last */}
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 22, paddingTop: 18, borderTop: "1px solid rgba(255,255,255,.08)" }}>
        <div>
          {stepIdx > 0 && <button className="btn btn-ghost btn-sm" onClick={goBack}>← رجوع</button>}
        </div>
        <div>
          {stepIdx < EDITOR_STEPS.length - 1 ? (
            <button
              className="btn btn-blue btn-sm"
              disabled={stepIdx === 0 && !canSave}
              title={stepIdx === 0 && !canSave ? "أدخل اسم الاستبيان للمتابعة" : undefined}
              onClick={goNext}
            >
              التالي ←
            </button>
          ) : (
            <button className="btn btn-primary" disabled={!canSave} onClick={() => onSave(survey)}>
              💾 حفظ الاستبيان
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// ── Root component ────────────────────────────────────────────────────────────
export default function SurveyManagement() {
  const [surveys, setSurveys] = useState(loadCustomSurveys);
  const [editing, setEditing] = useState(null); // survey object being edited, or null = list view

  const refresh = () => setSurveys(loadCustomSurveys());

  const handleCreate = () => setEditing(createSurvey());
  const handleEdit = (survey) => setEditing(survey);
  const handleSave = (survey) => { saveCustomSurvey(survey); refresh(); setEditing(null); };
  const handleCancel = () => setEditing(null);
  const handleDelete = (id) => { deleteCustomSurvey(id); refresh(); };
  const handleDuplicate = (id) => { duplicateCustomSurvey(id); refresh(); };
  const handleToggleStatus = (id, status) => { setCustomSurveyStatus(id, status); refresh(); };

  const handleDownloadTemplate = () => downloadBlob(buildSurveyTemplateBlob(), "قالب_استبيان.xlsx");

  const handleImportTemplate = async (file) => {
    const buf = await file.arrayBuffer();
    const survey = importSurveyFromTemplateArrayBuffer(buf);
    if (!survey.sections.length) throw new Error("لم يتم العثور على أسئلة في تبويب \"المحاور والأسئلة\" بالملف.");
    saveCustomSurvey(survey);
    refresh();
  };

  const handleExportBackup = () => downloadBlob(buildSurveysBackupBlob(), `نسخة_احتياطية_الاستبيانات_${new Date().toISOString().slice(0, 10)}.json`);

  const handleImportBackup = (text) => {
    const count = importSurveysBackup(text);
    refresh();
    return count;
  };

  if (editing) {
    return <SurveyEditorView survey={editing} onSave={handleSave} onCancel={handleCancel} />;
  }

  return (
    <SurveyListView
      surveys={surveys}
      onCreate={handleCreate}
      onEdit={handleEdit}
      onDelete={handleDelete}
      onDuplicate={handleDuplicate}
      onToggleStatus={handleToggleStatus}
      onDownloadTemplate={handleDownloadTemplate}
      onImportTemplate={handleImportTemplate}
      onExportBackup={handleExportBackup}
      onImportBackup={handleImportBackup}
    />
  );
}
