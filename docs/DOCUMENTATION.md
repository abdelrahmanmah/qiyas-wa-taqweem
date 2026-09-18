# Redaa2 — Technical Documentation

Architecture, data model, and processing-logic reference for the Redaa2 codebase. This
document intentionally excludes UI/UX details (layout, styling, wizard step visuals,
animations, colors) — for those, see the code directly or `CLAUDE.md`.

---

## 1. Overview

Redaa2 is a browser-only (no backend database, no server-side persistence) tool built for
ERU's Quality Assurance unit that:

- Parses Google Forms / UMIS Excel exports.
- Runs a schema-driven statistical analysis engine (Likert-5 / Likert-3).
- Generates RTL Arabic Word (`.docx`) reports.
- Generates and tracks per-semester Google Forms copies directly on Google Drive.
- Splits, matches, and classifies UMIS course-evaluation exports by department.
- Optionally uses an LLM (Groq / Gemini / Claude) as a tool-calling assistant over already-
  loaded survey data.

Everything runs client-side. The only server-side code is three Vite dev-middleware routes
(`/api/chat`, `/api/list-files`, `/api/read-file`) that exist purely to support local
development of the AI assistant feature.

**Entry point:** `src/main.jsx` → `src/App.jsx`
**Dev server:** `npx vite` (port 3000, must stay fixed — see §6)
**Build:** `npx vite build`

---

## 2. Repository Structure

```
src/
├── main.jsx                       # ReactDOM root
├── App.jsx                        # Wizard state machine + all step logic
├── AiChat.jsx                     # AI assistant: providers, tool-calling loop
├── EnhancedReportView.jsx         # On-screen report renderer + PDF export
├── SurveyManagement.jsx           # Custom-survey CRUD/editor logic
├── SemesterSurveys.jsx            # Semester Survey Generator logic
├── SemesterFormPicker.jsx         # Drive semester-survey browse/select bridge
├── CourseEvaluationHub.jsx        # Container routing between the 5 course-eval tools
├── CourseSplitter.jsx             # UMIS merged-export splitter/matcher/classifier
├── CourseTemplateTool.jsx         # Course-assignment sheet import/edit/export
├── SurveyParticipationTool.jsx    # Voting/participation classifier
├── PdfRecommendationReviewer.jsx  # PDF last-page thumbnail + recommendation flag export
├── engine/
│   ├── analyze.js                 # Excel parsing, column mapping, statistics
│   ├── buildDocx.js                # Word document generation (docx library)
│   ├── customSurveyModel.js        # Custom-survey data model + analysis-engine bridge
│   └── semesterSurveyModel.js      # Drive/Forms REST helpers, OAuth, naming utilities
└── schemas/
    ├── index.js                    # Loads + compiles every *.yaml at build time
    ├── compileSchema.js             # Expands minimal YAML → full runtime schema shape
    ├── student.yaml                 # likert-5, 25 axes
    ├── faculty.yaml                 # likert-3, 17 axes
    ├── assistant.yaml               # likert-3, 17 axes
    ├── graduates.yaml               # likert-3, 11 questions (flat)
    └── coordinator.yaml             # likert-5, 2 axes, 23 questions
```

`vite.config.js` also defines the yaml-loader Vite plugin and the three dev-only AI API
routes.

---

## 3. Technology Stack

| Layer | Library | Purpose |
|---|---|---|
| UI framework | React 19 | Component tree, all state local to components (no external state library) |
| Build/dev server | Vite 8 | Bundling, dev server, custom YAML-loader plugin, dev API middleware |
| Excel parsing/writing | `xlsx` (SheetJS) | Read `.xlsx/.xls/.csv`; write `.xlsx` (links export, course-eval report/template exports) |
| Word generation | `docx` | Build `.docx` binaries entirely client-side |
| Client-side PDF | `html2pdf.js` (html2canvas) | Enhanced-report PDF export |
| PDF rendering | `pdfjs-dist` | Render last page of uploaded PDFs to canvas thumbnails |
| Zipping | `jszip` | Bundle split course files, department-foldered |
| YAML (build-time) | `js-yaml` | Compile schema source files |
| Auth | Google Identity Services (GIS), OAuth2 implicit flow | Two independent token clients/scopes (see §7) |

No ORM, no database, no server framework. `localStorage` is the only persistence layer (see
§8).

---

## 4. Schema System (`src/schemas/*.yaml`)

Every built-in survey type is declared in a compact YAML file and compiled at build time by
`compileSchema.js` into the runtime shape consumed by `analyze.js`. `src/schemas/index.js`
uses `import.meta.glob("./*.yaml", { eager: true })` to load and compile every file
automatically — no manual registration step.

### Source format

```yaml
id: mysurvey
name: Arabic name
scale: 5            # or 3
hints: [keyword1, keyword2]     # filename auto-detect substrings

meta:
  email:   [Email, البريد]
  degree:  [الوظيفة, الدرجة]
  dept:    [القسم]
  program: [الإدارة]
  level:   [المستوى]            # likert-5 only
  freetext: [مقترحات]

axes:
  - name: "المحور الأول"
    recommendation: "text used in the recommendations section"
    questions:
      - "نص العبارة"
      - { text: "عبارة بعمود محدد", col: 4 }   # explicit colIndex override (likert-5)
```

A flat survey (no axes) uses a top-level `questions:` list instead; `compileSchema.js` wraps
it into one anonymous axis.

### Compiled shape

```jsonc
{
  "id": "mysurvey",
  "label": "...",
  "fileHints": [...],
  "scale": { "type": "likert-5"|"likert-3", "values": [...], "agreementCodes": [...] },
  "metadata": { "timestampCol": [...], "levelCol": [...], "degreeCol": [...], "departmentCol": [...], "freeTextCols": [...] },
  "questionStartIndex": 5,     // likert-3 only
  "interpretation": [{ "min": 4.5, "label": "...", "tier": "excellent", "color": "0d6e3a" }, ...],
  "axes": [{ "id": "ax01", "name": "...", "questions": [{ "id": "q01", "seq": 1, "text": "...", "colIndex": 0 }] }]
}
```

### Column detection strategy

| Scale | Method | Field |
|---|---|---|
| likert-5 | Filter out metadata columns → positional `colIndex` (axis/question order) | `colIndex` |
| likert-3 | Fuzzy text similarity between header and question text | `questionStartIndex` (skip N leading cols) |

### Current built-in survey types

| Schema id | Scale | Axes | Notes |
|---|---|---|---|
| `student` | likert-5 | 25 | has `levelCol` metadata |
| `faculty` | likert-3 | 17 | degree × department cross-tab |
| `assistant` | likert-3 | 17 | degree × department cross-tab |
| `graduates` | likert-3 | 1 (flat, 11 questions) | fixed recommendation text |
| `coordinator` | likert-5 | 2 (23 questions) | custom Word layout, degree × department cross-tab |

---

## 5. Statistical Analysis Engine (`src/engine/analyze.js`)

### Response value parsing

`parseResponse(v, schema)` dispatches on `schema.scale.type`.

**likert-5** recognizes, in priority order:
1. Parenthesized digit, e.g. `"(5) أوافق بشدة"` → `"5"` (real Google Forms export format).
2. Bare digit `"1"`–`"5"`.
3. Plain Arabic 5-point labels with no code (`"أوافق بشدة"` → `"5"`, down to `"لا أوافق بشدة"`
   → `"1"`) — added for custom-survey exports where some Google Forms configurations write the
   label only. This branch only runs if neither of the first two patterns matched, so built-in
   survey parsing is unaffected.

**likert-3** matches plain Arabic/English keywords (`أوافق`/`موافق`/`agree`,
`محايد`/`neutral`, `لا أوافق`/`disagree`) directly — no numeric-code dependency.

A cell that matches nothing returns `null` and is excluded from all counts for that question.

### Pipeline

```
Raw Excel rows (readExcel())
   → detectSurveyType(filename, headers) — filename hints, then fuzzy header match
   → schema selection
   → identify metadata columns (timestamp/email/degree/dept/level/freetext)
   → remaining columns mapped to questions (colIndex or fuzzy header match)
   → per-user row filtering (removed rows, dept/degree filters)
   → analyzeRows()/analyze():
        - per-question counts, percentages, mean (likert-5), agreement %
        - per-axis aggregation + interpretation tier from schema.interpretation
        - overall mean/agreement %/direction
        - byDegree / byDepartment counts
        - crossDegreeByDept matrix (when both columns detected)
   → result object
```

### Result object shape

```javascript
{
  schemaId, schemaLabel, scaleType, n,
  axes: [{
    id, name, axisMean, axisAgreePct, direction, tier,
    questions: [{ seq, id, text, counts, pcts, mean, agreePct, total }]
  }],
  overallMean, overallAgreePct, overallDirection,
  byDegree, byDepartment,
  crossDegreeByDept: { rows, cols, matrix, rowTotals, colTotals, grandTotal },
  totalQuestions,
}
```

### Comparison mode

For 2- or 3-year comparisons, the pipeline above runs once per year/slot; `buildComparison()`
merges the per-year `result` objects into a single `comparison` object (with a `.slots` array
carrying per-year metadata) consumed by `buildComparisonDocx()`.

---

## 6. Custom Survey System (`src/engine/customSurveyModel.js`)

A second, independent survey-definition mechanism alongside the 5 built-in YAML schemas —
lets a survey be defined at runtime (no code change, no rebuild) and become immediately
usable by the analysis engine.

### Data model

```javascript
{
  id, name, description, version, status: "draft" | "active",
  surveyType, scaleType: "likert-3" | "likert-5",
  metadata: { timestampCol, emailCol, nameCol, degreeCol, departmentCol, freeTextCols },
  sections: [{
    id, name, description,
    questions: [{ id, text, excelColumn, type: "likert" | "text" | "numeric", required, category, weight }]
  }],
  createdAt, updatedAt,
}
```

`excelColumn` defaults to the question's own `text` if left blank, applied once in
`saveCustomSurvey()` regardless of how the survey was built.

### Storage

`localStorage` key `eruQA_customSurveys_v1`, scoped per browser origin **including port** —
running the dev server on a different port makes previously saved surveys appear to vanish.
A JSON backup export/import mechanism (`buildSurveysBackupBlob` / `importSurveysBackup`,
merged by survey `id`) exists as a portability layer independent of browser storage.

### Building a survey (three input paths)

1. **Manual** — step-by-step model construction via the exposed factories
   (`createSurvey`/`createSection`/`createQuestion`/`createMetadata`).
2. **Import from a real response-data file** (`importSurveyStructureFromRows`) — auto-detects
   which columns are questions vs. general-info metadata via `classifyHeader()`. All detected
   questions land in one section since a flat header row carries no axis information.
   Classification guards against misclassifying long Likert questions that happen to mention a
   metadata keyword mid-sentence: a column is only treated as metadata if its header is a short
   label (≤5 words) **and** its sample answers aren't Likert-scaled.
3. **Import a structured Excel template** (`buildSurveyTemplateBlob` /
   `importSurveyFromTemplateArrayBuffer`) — a 2-sheet workbook (info sheet + محور/سؤال/عمود/
   نوع/إلزامي/فئة/وزن rows), grouped into sections by المحور value, with a fallback single
   section for surveys with no axes at all.

### Analysis-engine bridge

`toAnalysisSchema(survey)` converts a custom survey into the exact schema shape
`compileSchema.js` produces for built-in surveys, so `analyze()`/`analyzeRows()` run on it
completely unmodified:

- `getActiveCustomAnalysisSchemas()` — all `status: "active"` surveys, converted (surveys with
  zero Likert questions are skipped).
- `getAllAnalysisSchemas()` — `{ ...5 built-in schemas, ...active custom surveys }`.
- `detectAnySurveyType(filename, headers)` — the same hint-match-then-fuzzy algorithm as
  `analyze.js`, parameterized over the merged schema set.

**Known limitations** (both a direct consequence of reusing `analyze()` unmodified by
design):
1. Only `type: "likert"` questions are analyzed — `text`/`numeric` questions are excluded from
   quantitative output entirely.
2. For likert-5 custom surveys, column mapping is **positional** (`colIndex` assigned in
   current section/question order) — reordering questions changes which file column each
   question reads from. likert-3 custom surveys are unaffected since matching is by header
   text, not position.

---

## 7. Word Report Generation (`src/engine/buildDocx.js`)

All reports share a common header/footer/cover structure; section content varies by schema.

| Section | Standard surveys | `coordinator` schema |
|---|---|---|
| Cover | `buildCoverPage` | same |
| Evaluators table | `buildEvaluatorsSection` | same (if enabled in settings) |
| Variables | `buildVariablesSection` | same |
| Methodology | `buildMethodologySection` | same |
| Participants | `buildParticipantsSection` | same (auto cross-tab from degreeCol × departmentCol) |
| Detailed results | `buildDetailedSection` | `buildCoordinatorDetailedSection` (9-column layout) |
| Summary | `buildSummarySection` | `buildCoordinatorSummarySection` (4-column layout) |
| Recommendations | `buildRecommendationsSection` | skipped |

`buildComparisonDocx(comparison, meta, settings)` generates the multi-year comparison variant.

### RTL rendering constraint

`docx` 9.x has no working section-level RTL — `properties: { bidi }` is silently ignored, so
the entire document renders in an LTR frame. Word then resolves every bidi (RTL) paragraph's
`left`/`right` against that LTR frame and flips them (`AlignmentType.RIGHT` ends up hugging the
visual left). The `cellAlign(a)` helper compensates by swapping `RIGHT↔LEFT` whenever
`_isRTL` is set (center/both untouched); every RTL paragraph alignment in the file is routed
through it. This is a documented, intentional workaround — not a bug to "simplify" away.

---

## 8. Settings & Persistence (`localStorage` keys)

| Key | Contents |
|---|---|
| `eruQA_settings_v1` | Institution name, fonts, colors, signature blocks, evaluator-table toggle (`DEFAULT_SETTINGS` in `buildDocx.js`) |
| `eruQA_ai_v1` | AI assistant provider/model/API key/surveys-folder path |
| `eruQA_customSurveys_v1` | Custom survey definitions (see §6) |
| Drive/Semester token keys | Persisted OAuth access tokens with expiry timestamps (see §7 of `CLAUDE.md` for the silent-refresh mechanism) |
| Department/program list (Semester Survey Generator) | Editable local config, seeded with a default department/program list |

No server-side database exists anywhere in the system.

---

## 9. Google Integrations

### 9.1 Drive (read-only, main wizard)

Pure client-side OAuth2 via Google Identity Services — `initTokenClient`, popup-based, no
backend. Scope: `drive.readonly`. Used to list/browse/download survey response files and
auto-export Google Sheets to `.xlsx` via the Drive `/export` endpoint.

### 9.2 Drive + Forms (read/write, Semester Survey Generator)

A **separate** token client/scope (`drive`, `forms.body`, `forms.responses.readonly`) since
this feature needs to copy, rename, move, and create folders, plus read Forms data — broader
access than the read-only Drive scope above. Kept fully independent so the main wizard's Drive
flow is unaffected.

Core operations (`src/engine/semesterSurveyModel.js`, no React):
- `listFormsInFolder` / `listSubfolders` — enumerate templates and year/semester folders,
  nothing hardcoded.
- `copyFile` → `moveFile` → `updateFormTitle` → `publishForm` → `getForm` — the per-survey
  generation pipeline. `publishForm` is required because forms created via the API start
  unpublished and reject responses until explicitly published (`forms.{id}:setPublishSettings`).
- `listAllResponses` — paginated via `nextPageToken`.
- `responsesToRows(form, responses)` — converts a Forms API `form`/`responses[]` pair into the
  exact `[header, ...rows]` shape `analyze.js` expects from a parsed Excel export, letting
  Forms-sourced data run through the completely unmodified analysis engine.
- `buildSurveyName(...)` / `departmentFromSurveyName(...)` — deterministic naming convention
  (`"{template} - [{department}[ - {program}]] - {semester} - {year}"`) used both to name the
  generated Drive file/form title and to recover department/program for report metadata later.
- `expandDepartmentUnits(departments, granularity)` — expands one template selection into
  per-department or per-department-and-program generation jobs based on the selected mode
  (`"general" | "departments" | "programs"`).

No linked-response-Sheet mechanism exists — the Forms REST API cannot read/attach the classic
Sheet-link feature (Apps-Script-only), and copying a form via the Drive API drops any existing
link. Response counts come exclusively from `forms.responses.list`.

### 9.3 Token persistence

Both scopes use the same client-side persistence pattern: `saveStoredToken`/`getStoredToken`/
`clearStoredToken` wrap `localStorage` with an expiry timestamp (60s safety margin), and a
`requestAccessToken({ prompt: "" })` silent-refresh attempt runs on mount before ever falling
back to a visible consent popup.

---

## 10. Course Evaluation Tooling

Five independent, standalone modules (no shared state, no props from `App`) grouped behind
`CourseEvaluationHub.jsx`. Each reads/writes only local files — no network calls.

### 10.1 Course Splitter (`CourseSplitter.jsx`)

Splits a single UMIS RDLC-format merged Excel export (all courses' results in one sheet, one
repeating block-header marker per course) into one `.xlsx` workbook per course.

- **Block detection**: scans every row for a configurable marker string (default `"بنود
  الاستبيان"`) and a configurable row offset (default `10`, the distance between a course's
  data start and its marker row) — both user-editable since different UMIS report
  templates/versions may vary. If the configured marker matches zero rows,
  `findMarkerCandidates(wb)` scans for other strings that repeat a plausible number of times
  (2–200) and surfaces them as fallback suggestions.
- **Course identification**: regex-matches a `"... (CODE)"` pattern in the rows just above each
  marker to recover the course code and title.
- **Splitting**: slices the sheet into per-course row ranges (fixed header rows + block rows),
  remapping cell merges (`!merges`) into the new row-index space for each split-out workbook.
- **Matching**: an optional reference course list (`COURSE_CODE`/`COURSE_DESCR_EN`/
  `COURSE_DESCR_AR` columns) determines which expected courses have **no** split-out file at
  all (missing).
- **Department classification**: an optional department-distribution list, matched by course
  **name first** (exact match, then Levenshtein-similarity fuzzy match ≥0.8), falling back to
  exact code match — name-first because course codes are observed to vary between UMIS and
  other systems while names don't (`findDepartment()`).
- **Duplicate detection**: the same course code appearing in more than one uploaded merged
  file is flagged.
- **Outputs**: per-course `.xlsx` download, a department-foldered ZIP of every split course
  (unmatched courses land in a `"غير محدد - يحتاج مراجعة"` folder), and a 3-sheet Excel report
  (موجودة / ناقصة / مكررة) summarizing the full run.

### 10.2 Course Template Tool (`CourseTemplateTool.jsx`)

Manages the pre-cycle course-assignment sheet (اسم المقرر / كود المقرر / عضو هيئة التدريس /
عضو الهيئة المعاونة / القسم العلمي / هل يوجد لاب / القائم بالمراجعة). Header matching on
import is normalized-text-based (parenthetical suffixes stripped before comparison) with a
positional fallback for unrecognized headers. Computes: distribution by department (course +
lab counts), workload per instructor/assistant/reviewer, and missing-data/duplicate-code
detection — all derived, no persistence.

### 10.3 Survey Participation Tool (`SurveyParticipationTool.jsx`)

Reads a system export with `NoOfVotes` as a `"voted/total"` fraction string (`parseVotes()`
falls back to treating the whole value as `voted` with `total: null` if no `/` is present).
Classification (`classify()`):

| Condition | Category |
|---|---|
| `total === 0` | لا يوجد طلاب مسجلين |
| `voted === 0 && total > 0` | لم يتم التقييم |
| `0 < voted < threshold` (default 10, configurable) | مشكوك في انتظامها |
| otherwise | طبيعية |

Exports per-category and combined Excel reports.

### 10.4 PDF Recommendation Reviewer (`PdfRecommendationReviewer.jsx`)

Renders the last page of every PDF in a user-picked folder (`webkitdirectory`) as a thumbnail
via a 4-way concurrent render pool (`pdfjs-dist`, worker resolved through Vite's `?url`
import — no external CDN dependency at runtime). Exports the checked (has-recommendation)
subset as Excel.

---

## 11. Export Summary

| Export | Format | Mechanism |
|---|---|---|
| Annual survey report | `.docx` | `buildAnnualDocx()` → `docx` → Blob |
| Comparison report | `.docx` | `buildComparisonDocx()` → `docx` → Blob |
| Enhanced report | `.pdf` | `html2pdf.js` (html2canvas) client-side render |
| Semester survey report | `.docx` | `buildAnnualDocx()` fed by `responsesToRows()`-converted Forms data |
| Survey links list | `.xlsx` / clipboard text | `XLSX.utils.json_to_sheet` + `XLSX.write`, or `navigator.clipboard.writeText` |
| Split course workbook(s) | `.xlsx` / `.zip` | `XLSX.write` per course, `jszip` for the bundled ZIP |
| Course-eval status reports | `.xlsx` | `XLSX.utils.json_to_sheet` + `XLSX.write` |

All exports are generated and downloaded entirely client-side; no file is ever uploaded to a
server as part of report generation.

---

## 12. Known Limitations

- No automated test suite — correctness relies on manual verification and `npx vite build`.
- The AI assistant is non-functional without a configured provider API key.
- The Drive/Forms integrations require Google Cloud OAuth setup
  (`VITE_GOOGLE_CLIENT_ID`, `VITE_GOOGLE_TEMPLATE_FOLDER_ID`,
  `VITE_GOOGLE_ROOT_SURVEYS_FOLDER_ID`) and, while the consent screen is in "Testing" mode,
  every user account must be added as a test user with the broader Forms scopes enabled.
- Semester Survey Generator's "Analyze All" depends on schema auto-detection succeeding —
  surveys whose wording doesn't reasonably match a known schema are skipped, not force-matched.
- Custom likert-5 surveys use positional column mapping — reordering questions after real
  response data has been collected against the old order will misalign analysis.
- `docx` 9.x's RTL limitation requires the `cellAlign` workaround throughout `buildDocx.js` —
  a known, permanent constraint of the library, not a defect to fix.
- The Course Splitter's default block marker/offset were reconstructed against a real,
  verified-working reference implementation (`legacy/course_eval_splitte.html`) but different
  UMIS report template versions may still require adjusting them via the exposed
  marker/offset settings.
