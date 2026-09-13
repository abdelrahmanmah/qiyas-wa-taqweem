// Semester Survey Generator — Google Drive + Forms REST integration.
// Google Drive is the single source of truth: no database, no localStorage cache.
// OAuth is client-side only (Google Identity Services token client), same pattern
// as the existing Drive integration in App.jsx, just with broader scopes.

const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID ?? "";
export const TEMPLATE_FOLDER_ID = import.meta.env.VITE_GOOGLE_TEMPLATE_FOLDER_ID ?? "";
export const ROOT_SURVEYS_FOLDER_ID = import.meta.env.VITE_GOOGLE_ROOT_SURVEYS_FOLDER_ID ?? "";

export const SEMESTER_SCOPE = [
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/forms.body",
  "https://www.googleapis.com/auth/forms.responses.readonly",
].join(" ");

export const SEMESTERS = [
  "خريف",
  "ربيع",
  "صيف",
];

// Departments (and, where relevant, the programs inside them) that "نسخة لكل قسم"
// expands into. Editable at runtime via the Departments tab in SemesterSurveys.jsx —
// persisted to localStorage (not Drive: this is app configuration, not survey data).
export const DEFAULT_DEPARTMENTS = [
  { id: "accounting", name: "محاسبة", programs: [] },
  { id: "economics", name: "اقتصاد", programs: [] },
  { id: "political-science", name: "علوم سياسية", programs: [] },
  { id: "business-admin", name: "إدارة أعمال", programs: ["مالية", "تسويق"] },
  { id: "business-tech", name: "تكنولوجيا الأعمال", programs: ["BA", "MIS", "Fintech", "MKI"] },
];

const DEPARTMENTS_KEY = "eruQA_departments_v1";

export function loadDepartments() {
  try {
    const raw = localStorage.getItem(DEPARTMENTS_KEY);
    if (raw) return JSON.parse(raw);
  } catch { /* ignore */ }
  return DEFAULT_DEPARTMENTS;
}

export function saveDepartments(departments) {
  try { localStorage.setItem(DEPARTMENTS_KEY, JSON.stringify(departments)); } catch { /* ignore */ }
}

// Expands the department list into generation units, per the chosen granularity:
// - "departments": always one unit per department, regardless of whether it has
//   programs (a department with programs is NOT further split — used when the user
//   wants a per-department copy only, e.g. "الأقسام").
// - "programs" (default): a department with no programs becomes one unit (the
//   department itself); a department WITH programs becomes one unit per program
//   instead (used for "الأقسام والبرامج").
export function expandDepartmentUnits(departments, granularity = "programs") {
  const units = [];
  for (const d of departments) {
    if (granularity === "departments" || !d.programs?.length) {
      units.push({ label: d.name, department: d.name, program: null });
    } else {
      for (const p of d.programs) {
        units.push({ label: `${d.name} - ${p}`, department: d.name, program: p });
      }
    }
  }
  return units;
}

const FORM_MIME = "application/vnd.google-apps.form";
const FOLDER_MIME = "application/vnd.google-apps.folder";
const DRIVE_API = "https://www.googleapis.com/drive/v3";
const FORMS_API = "https://forms.googleapis.com/v1";

export function loadGisScript() {
  return new Promise(resolve => {
    if (window.google?.accounts?.oauth2) { resolve(); return; }
    const s = document.createElement("script");
    s.src = "https://accounts.google.com/gsi/client";
    s.onload = resolve;
    document.head.appendChild(s);
  });
}

// GIS access tokens are short-lived (~1hr) and there's no refresh token in this
// implicit, backend-less flow — but re-prompting for an account/consent on every
// page load is a bad experience. Two mitigations, both client-side only:
// 1) cache the still-valid token in localStorage so a reload doesn't need Google at all;
// 2) once it expires, try a *silent* requestAccessToken({prompt:""}) first (no popup,
//    works if the browser still has an active Google session + prior consent) before
//    ever falling back to the visible account-picker button.
// Shared by the entire app (Drive uploads, Forms creation/responses and dashboards).
// Versioned separately from the former per-screen tokens so an older, narrower token
// is never mistaken for a fully-authorized global connection.
export const SEMESTER_TOKEN_KEY = "eruQA_googleWorkspaceToken_v1";

export function saveStoredToken(key, accessToken, expiresInSec) {
  try {
    const expiresAt = Date.now() + (expiresInSec ?? 3600) * 1000 - 60000; // 60s safety margin
    localStorage.setItem(key, JSON.stringify({ accessToken, expiresAt }));
  } catch { /* ignore */ }
}

export function getStoredToken(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const { accessToken, expiresAt } = JSON.parse(raw);
    if (!accessToken || !expiresAt || Date.now() >= expiresAt) return null;
    return accessToken;
  } catch { return null; }
}

export function clearStoredToken(key) {
  try { localStorage.removeItem(key); } catch { /* ignore */ }
}

export function initSemesterTokenClient({ onToken, onError }) {
  if (!GOOGLE_CLIENT_ID) throw new Error("VITE_GOOGLE_CLIENT_ID غير مضبوط.");
  if (!window.google?.accounts?.oauth2) throw new Error("لم يتم تحميل Google Identity Services بعد.");
  return window.google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: SEMESTER_SCOPE,
    prompt: "select_account",
    callback: (resp) => {
      if (resp.error) { onError?.(resp.error); return; }
      onToken?.(resp.access_token, resp.expires_in);
    },
  });
}

async function apiFetch(base, token, path, options = {}) {
  const res = await fetch(`${base}${path}`, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`${res.status} ${res.statusText}${body ? " — " + body.slice(0, 300) : ""}`);
  }
  return res.status === 204 ? null : res.json();
}

const driveFetch = (token, path, options) => apiFetch(DRIVE_API, token, path, options);
const formsFetch = (token, path, options) => apiFetch(FORMS_API, token, path, options);

function escapeDriveQueryValue(v) {
  return String(v).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

export async function listFormsInFolder(token, folderId) {
  if (!folderId) return [];
  const q = `'${escapeDriveQueryValue(folderId)}' in parents and mimeType='${FORM_MIME}' and trashed=false`;
  const data = await driveFetch(token, `/files?q=${encodeURIComponent(q)}&fields=files(id,name,modifiedTime)&pageSize=200&orderBy=name`);
  return data?.files ?? [];
}

export async function listSubfolders(token, parentId) {
  if (!parentId) return [];
  const q = `'${escapeDriveQueryValue(parentId)}' in parents and mimeType='${FOLDER_MIME}' and trashed=false`;
  const data = await driveFetch(token, `/files?q=${encodeURIComponent(q)}&fields=files(id,name)&pageSize=200&orderBy=name`);
  return data?.files ?? [];
}

export async function findFolderByName(token, parentId, name) {
  const q = `'${escapeDriveQueryValue(parentId)}' in parents and mimeType='${FOLDER_MIME}' and name='${escapeDriveQueryValue(name)}' and trashed=false`;
  const data = await driveFetch(token, `/files?q=${encodeURIComponent(q)}&fields=files(id,name)&pageSize=1`);
  return data?.files?.[0] ?? null;
}

export async function createFolder(token, parentId, name) {
  return driveFetch(token, `/files?fields=id,name`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, mimeType: FOLDER_MIME, parents: [parentId] }),
  });
}

export async function getOrCreateFolder(token, parentId, name) {
  const existing = await findFolderByName(token, parentId, name);
  if (existing) return existing;
  return createFolder(token, parentId, name);
}

export async function copyFile(token, fileId, newName) {
  return driveFetch(token, `/files/${fileId}/copy?fields=id,name,parents`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: newName }),
  });
}

export async function moveFile(token, fileId, newParentId, oldParentIds) {
  const params = new URLSearchParams({ addParents: newParentId, fields: "id,parents" });
  if (oldParentIds?.length) params.set("removeParents", oldParentIds.join(","));
  return driveFetch(token, `/files/${fileId}?${params.toString()}`, { method: "PATCH" });
}

export async function getForm(token, formId) {
  return formsFetch(token, `/forms/${formId}`);
}

// Copying a form via Drive renames the Drive *file*, not the form's own title (what
// respondents see when they open it). Set both so they stay in sync.
export async function updateFormTitle(token, formId, title) {
  return formsFetch(token, `/forms/${formId}:batchUpdate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      requests: [{ updateFormInfo: { info: { title }, updateMask: "title" } }],
    }),
  });
}

// As of 2026-06-30, forms created via the API start unpublished and don't accept
// responses until explicitly published — publish + open for responses right away so
// a generated survey is immediately usable without anyone opening it in the UI first.
export async function publishForm(token, formId) {
  return formsFetch(token, `/forms/${formId}:setPublishSettings`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      publishSettings: { publishState: { isPublished: true, isAcceptingResponses: true } },
    }),
  });
}

export async function listAllResponses(token, formId) {
  let all = [];
  let pageToken;
  do {
    const qs = pageToken ? `?pageToken=${encodeURIComponent(pageToken)}` : "";
    const data = await formsFetch(token, `/forms/${formId}/responses${qs}`);
    all = all.concat(data?.responses ?? []);
    pageToken = data?.nextPageToken;
  } while (pageToken);
  return all;
}

export function computeResponseStats(responses) {
  let last = null;
  for (const r of responses) {
    if (r.lastSubmittedTime && (!last || r.lastSubmittedTime > last)) last = r.lastSubmittedTime;
  }
  return { count: responses.length, lastSubmittedTime: last };
}

export async function runBatched(items, worker, batchSize = 4) {
  const results = [];
  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize);
    const batchResults = await Promise.all(batch.map((item, j) => worker(item, i + j)));
    results.push(...batchResults);
  }
  return results;
}

export function buildSurveyName(templateName, semester, academicYear, department, program = null) {
  const unit = [department, program].filter(Boolean).join(" - ");
  return unit
    ? `${templateName} - ${unit} - ${semester} - ${academicYear}`
    : `${templateName} - ${semester} - ${academicYear}`;
}

export function yearFolderName(academicYear) {
  return (academicYear || "").trim().replace("/", "-");
}

export function isValidAcademicYear(v) {
  return /^\d{4}\/\d{4}$/.test((v || "").trim());
}

// Converts a Forms API `form` + its `responses` into the [header, ...rows] shape
// analyze.js expects from a real Google Forms→Sheets Excel export, so the existing
// analysis engine (detectAnySurveyType/analyze) works completely unmodified —
// answer values come back as plain option text (e.g. "أوافق بشدة"), which the
// existing parseResponse5/parseResponse3 fallbacks already handle (see CLAUDE.md
// "Response Value Parsing").
export function responsesToRows(form, responses) {
  const items = (form.items ?? [])
    .filter(it => it.questionItem?.question)
    .map(it => ({ title: it.title ?? "", questionId: it.questionItem.question.questionId }));
  const collectsEmail = !!form.settings?.emailCollectionType && form.settings.emailCollectionType !== "DO_NOT_COLLECT";

  const header = ["Timestamp", ...(collectsEmail ? ["Email Address"] : []), ...items.map(i => i.title)];
  const rows = responses.map(r => {
    const ts = r.lastSubmittedTime || r.createTime || "";
    const answers = items.map(i => {
      const vals = r.answers?.[i.questionId]?.textAnswers?.answers;
      return vals?.length ? vals.map(v => v.value).join(", ") : "";
    });
    return [ts, ...(collectsEmail ? [r.respondentEmail ?? ""] : []), ...answers];
  });
  return [header, ...rows];
}

// A survey name built by buildSurveyName() is always "{template} - [...unit] - {semester} -
// {year}", where the department/program unit in the middle is 0, 1, or 2 " - "-joined
// segments (e.g. "تكنولوجيا الأعمال - Fintech" for a program-level copy). Semester and year
// are always the last two segments and the template is always the first, so whatever's left
// in between (if anything) is the unit — parse it back out for report metadata
// (buildAnnualDocx's meta.program).
export function departmentFromSurveyName(name) {
  const parts = (name || "").split(" - ");
  if (parts.length <= 3) return "";
  return parts.slice(1, parts.length - 2).join(" - ");
}

export async function findYearSemesterFolder(token, academicYear, semester) {
  const yFolder = await findFolderByName(token, ROOT_SURVEYS_FOLDER_ID, yearFolderName(academicYear));
  if (!yFolder) return null;
  return findFolderByName(token, yFolder.id, semester);
}

// Expands selected templates (each carrying its own .mode — "general"/"departments"/
// "programs", set per-row in GenerateSurveysView, or uniformly by a chatbot tool call)
// into the flat job list runGenerationJobs() executes. Pure — no Drive calls — so it can
// seed job-progress UI state before the batch starts.
export function buildGenerationJobs(templates, departments) {
  const jobList = [];
  for (const t of templates) {
    const mode = t.mode || "general";
    if (mode === "departments" || mode === "programs") {
      const units = expandDepartmentUnits(departments, mode);
      for (const u of units) jobList.push({
        key: `${t.id}:${u.label}`,
        template: t,
        department: u.department,
        program: u.program,
        unitLabel: u.label,
      });
    } else {
      jobList.push({ key: t.id, template: t, department: null, program: null, unitLabel: null });
    }
  }
  return jobList;
}

// Creates {year}/{semester}/{department}/{survey folder}/{form}. Program-level
// forms live together inside the survey folder for their department; a program is
// represented in the form name and never gets its own folder. General surveys use
// {year}/{semester}/{survey folder}/{form} because they have no department.
// onJobUpdate(key, status) is optional — lets callers drive a live progress UI.
export async function runGenerationJobs(token, jobList, { year, semester }, onJobUpdate) {
  const yFolder = await getOrCreateFolder(token, ROOT_SURVEYS_FOLDER_ID, yearFolderName(year));
  const semFolder = await getOrCreateFolder(token, yFolder.id, semester);

  const departmentNames = [...new Set(jobList.map(j => j.department).filter(Boolean))];
  const departmentFolders = {};
  await Promise.all(departmentNames.map(async department => {
    departmentFolders[department] = await getOrCreateFolder(token, semFolder.id, department);
  }));

  const destinationFolders = {};
  const uniqueDestinations = [...new Map(jobList.map(job => [
    JSON.stringify([job.department || null, job.template.id]),
    job,
  ])).entries()];
  await Promise.all(uniqueDestinations.map(async ([key, job]) => {
    const parent = job.department ? departmentFolders[job.department] : semFolder;
    destinationFolders[key] = await getOrCreateFolder(token, parent.id, job.template.name);
  }));

  const done = await runBatched(jobList, async job => {
    onJobUpdate?.(job.key, "active");
    try {
      const newName = buildSurveyName(job.template.name, semester, year, job.department, job.program);
      const copy = await copyFile(token, job.template.id, newName);
      const destinationKey = JSON.stringify([job.department || null, job.template.id]);
      const destinationFolder = destinationFolders[destinationKey];
      await moveFile(token, copy.id, destinationFolder.id, copy.parents);
      await updateFormTitle(token, copy.id, newName);
      await publishForm(token, copy.id);
      const form = await getForm(token, copy.id);
      onJobUpdate?.(job.key, "done");
      return { ok: true, row: {
        key: job.key,
        name: newName,
        department: job.department,
        program: job.program,
        surveyFolder: job.template.name,
        formId: copy.id,
        formUrl: form.responderUri || formViewUrl(copy.id),
      } };
    } catch (e) {
      onJobUpdate?.(job.key, "error");
      return { ok: false, key: job.key, error: e.message };
    }
  }, 4);

  const successRows = done.filter(d => d.ok).map(d => d.row);
  const failCount = done.length - successRows.length;
  return { successRows, failCount };
}

// Reads the current two-level department/survey layout, plus direct forms and the
// previous one-level layout so already-generated surveys remain visible.
export async function listSemesterSurveysWithStats(token, semesterFolderId) {
  const [directForms, childFolders] = await Promise.all([
    listFormsInFolder(token, semesterFolderId),
    listSubfolders(token, semesterFolderId),
  ]);
  const perFolder = await runBatched(childFolders, async topFolder => {
    const [files, surveyFolders] = await Promise.all([
      listFormsInFolder(token, topFolder.id),
      listSubfolders(token, topFolder.id),
    ]);
    const directInFolder = files.map(f => ({
      ...f,
      department: null,
      surveyType: topFolder.name,
      parentId: topFolder.id,
    }));
    const nested = await runBatched(surveyFolders, async surveyFolder => {
      const nestedFiles = await listFormsInFolder(token, surveyFolder.id);
      return nestedFiles.map(f => ({
        ...f,
        department: topFolder.name,
        surveyType: surveyFolder.name,
        parentId: surveyFolder.id,
      }));
    }, 4);
    return [...directInFolder, ...nested.flat()];
  }, 4);
  const allFiles = [
    ...directForms.map(f => ({
      ...f,
      department: null,
      surveyType: "استبيان عام",
      parentId: semesterFolderId,
    })),
    ...perFolder.flat(),
  ];
  return runBatched(allFiles, async f => {
    try {
      const [form, responses] = await Promise.all([getForm(token, f.id), listAllResponses(token, f.id)]);
      const stats = computeResponseStats(responses);
      return { id: f.id, name: f.name, surveyType: f.surveyType, parentId: f.parentId, responses: stats.count, last: stats.lastSubmittedTime, formUrl: form.responderUri || formViewUrl(f.id), error: null };
    } catch (e) {
      return { id: f.id, name: f.name, surveyType: f.surveyType, parentId: f.parentId, responses: 0, last: null, formUrl: formViewUrl(f.id), error: e.message };
    }
  }, 4);
}

export function formViewUrl(formId) {
  return `https://docs.google.com/forms/d/${formId}/viewform`;
}

export function editorResponsesUrl(formId) {
  return `https://docs.google.com/forms/d/${formId}/edit#responses`;
}
