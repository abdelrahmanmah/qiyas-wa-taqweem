# Redaa2 — محلل الاستبيانات الأكاديمية

## Project Overview

A React (Vite) single-page application for ERU's Quality Assurance unit. It parses Google Forms Excel exports, runs statistical analysis, and generates RTL Arabic Word (.docx) reports.

**Entry point:** `src/main.jsx` → `src/App.jsx`  
**Dev server:** `npx vite` (port 3000) or `npx vite --port 5174`  
**Build:** `npx vite build`

---

## Architecture

```
src/
├── main.jsx                  # ReactDOM root
├── App.jsx                   # Full UI (step wizard, settings, survey picker)
├── AiChat.jsx                # Floating AI assistant panel
├── EnhancedReportView.jsx    # Standalone enhanced report view overlay
├── SurveyManagement.jsx      # "إدارة الاستبيانات" tab — custom survey CRUD/editor (see dedicated section below)
├── engine/
│   ├── analyze.js            # Excel parsing, column mapping, statistics
│   ├── buildDocx.js          # Word document generation (docx library)
│   └── customSurveyModel.js  # Custom-survey data model/storage + analysis-engine bridge (see below)
└── schemas/
    ├── index.js               # Loads + compiles all *.yaml files at build time into SCHEMAS
    ├── compileSchema.js        # Expands the minimal YAML format into the full runtime schema shape
    ├── student.yaml           # رضا الطلاب         — likert-5, 25 axes
    ├── faculty.yaml           # هيئة التدريس        — likert-3, 17 axes
    ├── assistant.yaml         # الهيئة المعاونة     — likert-3, 17 axes
    ├── graduates.yaml         # آراء الخريجين       — likert-3, 11 questions (flat)
    └── coordinator.yaml       # تقييم أداء منسقي البرامج — likert-5, 2 axes, 23 questions
```

> The schema system was migrated from hand-written `*.json` files to compact `*.yaml` files
> compiled at build time (`src/schemas/index.js` + `compileSchema.js`). The legacy `*.json`
> files have been deleted — they were unused duplicates. See "Schema System" below for the
> YAML format and "How to Add a New Survey Type" for the no-code path via the Survey
> Management UI.

`vite.config.js` also exposes three dev-only API routes (`/api/chat`, `/api/list-files`, `/api/read-file`) for the AI assistant feature.

### Components map (`App.jsx`)

All components live in one file. Listed in source order:

| Component / function | Line (approx) | Purpose |
|---|---|---|
| `CSS` (const string) | 8 | All inline styles injected via `<style>` |
| `loadGisScript()` | 186 | Lazily loads Google Identity Services SDK |
| `PROGRAMS` | 196 | Static list of faculty programs for Drive filter |
| `detectProgramFromFilename()` | 342 | Extracts program name from filename |
| `detectYearFromFilename()` | 357 | Extracts academic year (handles `2024_25` → `2024-2025`) |
| `detectTypeHintFromFilename()` | 365 | Matches survey type from filename vs `schema.fileHints` |
| `computeFilteredRows()` | 386 | Applies removed-set + dept/degree filters to data rows |
| `SettingsPanel` | 422 | Full settings UI (institution, fonts, colors, signatures, AI) |
| `BatchItem` | 785 | Single file row in batch processing mode |
| `BatchProcessor` | 871 | Multi-file batch processing UI |
| `DriveDashboard` | 1013 | Drive dashboard view (stats by type/program/year, per-file response counts) |
| `TUTORIAL_SLIDES` | 1283 | Array of 6 slide objects for the tutorial |
| `TutorialOverlay` | 1361 | 6-slide how-to guide with nav arrows + dots |
| `ProcessingOverlay` | 1512 | Step-by-step animation shown while reading/parsing a file |
| `LoadingOverlay` | 1629 | Generic spinner overlay (used for Word generation) |
| `StepBar` | 1650 | Wizard step progress bar |
| `SurveyTypePicker` | 1663 | Grid of survey type cards |
| `ReportModePicker` | 1697 | Annual / compare-2 / compare-3 selector |
| `FileSlot` | 1720 | Year+file upload row for comparison mode |
| `MetadataForm` | 1753 | Year, program, preparedBy, reviewer inputs |
| `ProgramPicker` | 1779 | Program dropdown |
| `FilterDropdown` | 1812 | Reusable labeled select |
| `DataPreviewTable` | 1834 | Row-level data preview with checkboxes + filters |
| `ResultsPreview` | 1905 | Inline result summary on step 3 |
| `ComparisonPreview` | 1994 | Comparison result summary |
| `LOCAL_STEPS` / `DRIVE_STEPS` | 2044 | Step label arrays for `ProcessingOverlay` |
| `App` (default export) | 2048 | Root component — all wizard state lives here |

### Key module-level constants (`App.jsx`)

```js
const STEPS        = ["رفع الملف", "التحقق", "نوع التقرير", "معاينة البيانات", "بيانات التقرير", "النتائج"];
const SETTINGS_KEY = "eruQA_settings_v1";
const AI_KEY       = "eruQA_ai_v1";
const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID ?? "";
const DRIVE_SCOPE  = "https://www.googleapis.com/auth/drive.readonly";
const GSHEETS_MIME = "application/vnd.google-apps.spreadsheet";
const LOCAL_STEPS  = ["قراءة البيانات", "تنظيف البيانات", "كشف نوع الاستبيان", "حساب المؤشرات", "تجهيز العرض"];
const DRIVE_STEPS  = ["تحميل الملف من Drive", "قراءة البيانات", "تنظيف البيانات", "كشف نوع الاستبيان", "تجهيز العرض"];
const PROGRAMS     = ["محاسبة", "اقتصاد", "إدارة", "علوم سياسية", "تكنولوجيا أعمال"];
```

### `SCHEMAS` is no longer a static import in `App.jsx`

```javascript
import { getAllAnalysisSchemas as allSchemas, detectAnySurveyType as detectSurveyType }
  from "./engine/customSurveyModel.js";
```

Every place that used to read the static `SCHEMAS` object now calls `allSchemas()` instead —
a function (from `customSurveyModel.js`) that returns `{ ...the 5 built-in YAML schemas,
...active custom surveys from Survey Management }`, recomputed fresh on every call (cheap
localStorage read, no caching/memoization needed). `detectSurveyType(...)` calls in `App.jsx`
are unchanged textually — the import alias quietly swaps the implementation for one that also
checks active custom surveys; behavior for the 5 built-in schemas is identical to before. See
"Survey Management System → Wiring into the analysis engine" below for the full mechanism.

---

## Wizard Step Flow

| Step | What the user sees | How entered |
|---|---|---|
| 0 | Mode selection (annual / compare-2 / compare-3) | Initial view |
| 1 | File upload zone (local or Drive tab) | "التالي ←" from step 0 |
| 2 | Type validation — detected schema card, manual override buttons, column preview | After `handleFileSelected` / `handleDriveFileSelect` animate → `setStep(2)` |
| 3 | Data preview — row table with remove checkboxes + live result preview | After "✓ تأكيد ومتابعة ←" in step 2 |
| 4 | Metadata form (year, program, preparedBy, reviewer) | After "متابعة ←" in step 3 |
| 5 | Results & download | After "عرض النتائج ←" in step 4 |

**Comparison mode** skips step 2 / step 3 — from step 1 the user fills the slots and clicks "تحليل ←" → goes straight to step 4 (comparison results).

---

## Data Flow

1. User picks mode (step 0) → clicks التالي → step 1 (upload)
2. File selected → `handleFileSelected()` reads with `FileReader`, parses with `readExcel()`, calls `detectSurveyType(filename, headers)` → animates steps → `setStep(2)`
3. Step 2: user confirms or overrides type → `handleValidationConfirm()` → `prepareData()` → `setStep(3)`
4. Step 3: user removes bad rows → `analyzeRows()` live preview → "متابعة" → `setStep(4)`
5. Step 4: metadata form → `downloadAnnual()` → `buildAnnualDocx()` → blob download
6. For comparison mode: `handleProcessMulti()` reads all slots, validates, calls `analyze()` per file then `buildComparison()` → `setStep(4)` (comparison results)
7. `buildComparisonDocx(comparison, meta, settings)` generates the Word blob

---

## Schema System (`src/schemas/*.yaml`)

Every built-in survey is defined by a compact YAML file, compiled at build time into the full
runtime schema shape consumed by `analyze.js`. `src/schemas/index.js` uses
`import.meta.glob("./*.yaml", { eager: true })` (via the `yaml-loader` Vite plugin in
`vite.config.js`) to load every `*.yaml` file and run it through `compileSchema.js`, producing
the `SCHEMAS` object keyed by `id`.

### YAML source format

```yaml
id: mysurvey
name: اسم الاستبيان بالعربي
nameEn: English Name
icon: "📋"
scale: 5            # or 3 — selects likert-5 vs likert-3
hints: [keyword1, keyword2]   # substrings matched against filename for auto-detection

meta:
  email:   [Email, البريد]      # → metadata.emailCol
  degree:  [الوظيفة, الدرجة]    # → metadata.degreeCol
  dept:    [القسم]              # → metadata.departmentCol
  program: [الإدارة]            # → metadata.programCol (likert-3 questionStartIndex count only)
  level:   [المستوى]            # → metadata.levelCol (likert-5 only)
  freetext: [مقترحات]           # → metadata.freeTextCols

axes:
  - name: "المحور الأول: ..."
    recommendation: "optional text used in التوصيات section"
    questions:
      - "نص العبارة الأولى"
      - { text: "عبارة بعمود محدد", col: 4 }   # optional explicit colIndex override (likert-5)
```

A flat (no-axes) survey like `graduates.yaml` omits `axes:` and uses a top-level `questions:`
list instead — `compileSchema.js` wraps it into a single anonymous axis.

### What `compileSchema.js` expands this into

```jsonc
{
  "id": "mysurvey",
  "label": "اسم الاستبيان بالعربي",
  "fileHints": ["keyword1", "keyword2"],
  "scale": {
    "type": "likert-5" | "likert-3",
    "values": [{ "code": "1", "label": "...", "score": 1 }, ...],
    "agreementCodes": ["4", "5"]   // codes counted as "agreement"
  },
  "metadata": {
    "timestampCol":  ["Timestamp"],
    "levelCol":      ["المستوى"],        // likert-5 only
    "degreeCol":     ["الوظيفة"],
    "departmentCol": ["التخصص"],
    "freeTextCols":  ["مقترحات"]
  },
  "questionStartIndex": 5,   // likert-3 only — computed from how many meta.* fields are set
  "interpretation": [
    { "min": 4.5, "label": "أوافق بشدة", "tier": "excellent", "color": "0d6e3a" },
    ...
  ],
  "axes": [
    {
      "id": "ax01",
      "name": "اسم المحور",
      "recommendation": "optional text used in توصيات section",
      "questions": [
        { "id": "q01", "seq": 1, "text": "نص العبارة", "colIndex": 0 }
        // colIndex: 0-based index into the FILTERED question columns
        // (after removing all metadata columns). Only used for likert-5.
      ]
    }
  ]
}
```

### likert-5 vs likert-3 column detection

| Scale | Detection method | Schema field |
|---|---|---|
| `likert-5` | Filter out metadata cols → use `colIndex` (positional, in axis/question order) | `colIndex` per question |
| `likert-3` | Fuzzy text similarity between header and `q.text` | `questionStartIndex` (skip first N cols) |

This same shape (and the same positional-vs-fuzzy distinction) is what
`customSurveyModel.js`'s `toAnalysisSchema()` produces for custom surveys built through the
Survey Management UI — see that section below.

---

## Response Value Parsing (`analyze.js`)

`parseResponse(v, schema)` dispatches to `parseResponse5` or `parseResponse3` based on
`schema.scale.type`. Both must return a code matching one of `schema.scale.values[].code` (or
`null` if the cell doesn't match anything, which excludes that response from all counts).

**`parseResponse5`** (likert-5) recognizes, in order:
1. A parenthesized digit, e.g. `"(5) أوافق بشدة"` → `"5"` — the format used by all 5 built-in
   YAML surveys' real Google Forms exports.
2. A bare digit `"1"`–`"5"`.
3. **(Added for custom surveys)** Plain Arabic 5-point labels with no leading code at all —
   `"أوافق بشدة"`, `"أوافق"`, `"محايد"`, `"لا أوافق"`, `"لا أوافق بشدة"` — mapped to codes
   `"5"`–`"1"` respectively. Some Google Forms exports write the label only, no `(digit)`
   prefix; this came up with a real custom survey (likert-5) created through Survey
   Management whose responses were all plain text, which made every count/percentage compute
   to `0` until this fallback was added. **This branch only runs if neither of the first two
   patterns matched**, so the 5 built-in surveys' parsing is byte-for-byte unchanged.

**`parseResponse3`** (likert-3) already worked purely off plain Arabic/English keywords
(`أوافق`/`موافق`/`agree`, `محايد`/`neutral`, `لا أوافق`/`disagree`) — no change needed there.

---

## Analysis Output (`result` object)

```javascript
{
  schemaId:         "coordinator",
  schemaLabel:      "تقييم أداء منسقي البرامج",
  scaleType:        "likert-5",
  n:                42,                 // total respondents
  axes: [{
    id:             "ax01",
    name:           "المحور الأول: ...",
    axisMean:       4.12,               // null for likert-3
    axisAgreePct:   78.5,               // % of agree/strongly-agree responses
    direction:      "أوافق",
    tier:           "good",
    questions: [{
      seq: 1, id: "q01", text: "...",
      counts:   { "1": 0, "2": 2, "3": 5, "4": 20, "5": 15 },
      pcts:     { "1": 0, "2": 4.8, "3": 11.9, "4": 47.6, "5": 35.7 },
      mean:     4.14,                   // null for likert-3
      agreePct: 83.3,
      total:    42,
    }]
  }],
  overallMean:       4.05,
  overallAgreePct:   76.2,
  overallDirection:  "أوافق",
  byDegree:          { "أستاذ": 10, "أستاذ مساعد": 32 },   // from degreeCol
  byDepartment:      { "محاسبة": 20, "اقتصاد": 22 },        // from departmentCol
  crossDegreeByDept: { rows, cols, matrix, rowTotals, colTotals, grandTotal },
  totalQuestions:    23,
}
```

---

## Word Report Structure (`buildDocx.js`)

All reports share the same header/footer/cover structure. Section content varies by schema:

| Section | Standard surveys | Coordinator only |
|---|---|---|
| Cover | `buildCoverPage` | same |
| Evaluators table | `buildEvaluatorsSection` | same (if setting enabled) |
| بيان بعدد المشاركين (fixed) | `buildFixedParticipantsSection` — faculty/assistant/coordinator only, see below | same |
| أولاً: متغيرات | `buildVariablesSection` | same |
| ثانياً: معالجة إحصائية | `buildMethodologySection` | same |
| ثالثاً: المشاركون | `buildParticipantsSection` | same (cross-tab auto from degreeCol×departmentCol) — skipped when the fixed table above already covered it |
| رابعاً: النتائج | `buildDetailedSection` | `buildCoordinatorDetailedSection` |
| خامساً: الملخص | `buildSummarySection` | `buildCoordinatorSummarySection` |
| التوصيات | `buildRecommendationsSection` | skipped |

### Fixed-category participants headcount table (faculty / assistant / coordinator)

Placed right after the evaluators table and **before** أولاً — unlike `buildParticipantsSection`'s
cross-tab (ثالثاً, whatever degree/department values happen to appear in the uploaded file, sorted
alphabetically), this table's rows and columns are a **fixed, known list per survey type**, always
shown in the same order, with `0` for any category that had no respondents in this particular file:

| Schema | Row categories (الدرجة/الوظيفة) | Columns (القسم) |
|---|---|---|
| `faculty` | مدرس، أستاذ مساعد | إدارة، اقتصاد، محاسبة، علوم سياسية، تكنولوجيا الأعمال |
| `assistant` | معيد، مدرس مساعد | إدارة، اقتصاد، محاسبة، علوم سياسية، تكنولوجيا الأعمال |
| `coordinator` | أستاذ، أستاذ مساعد، مدرس، مدرس مساعد، معيد، طالب | إدارة، اقتصاد، محاسبة، علوم سياسية، تكنولوجيا الأعمال |

`FIXED_PARTICIPANT_CONFIG` (in `analyze.js`) holds these lists; `fixedCrossTab()` buckets each raw
degree/department cell value against them via the same `bucketKey`/`normalize` Arabic-spelling-
variant matching the rest of the file already uses (so `"أستاذ  مساعد"` / `"استاذ مساعد"` etc. all
land in the same category) — a value that doesn't match any configured category is silently
excluded from this table (but still counted in the survey's overall `n`). The result field is
`result.fixedParticipants` (same `{ rows, cols, matrix, rowTotals, colTotals, grandTotal }` shape
`crossTab()` produces, rendered through the same `buildCrosstabTable()` helper). Both this table
and the old ثالثاً one are gated by the same `includeParticipants` setting.

The switch in `buildAnnualDocx`:
```javascript
const isCoord = result.schemaId === "coordinator";
```

### Coordinator رابعاً columns (9 columns, left→right physical = right→left visual in RTL)
`نسبة الرضا | المتوسط | أوافق بشدة | أوافق | محايد | لا أوافق | لا أوافق بشدة | العبارة | م`

After each axis: a merged summary row showing `متوسط الرضا للمحور` with the axis mean and agreement %.

### Coordinator خامساً columns (4 columns)
`نسبة الرضا | متوسط الرضا | المحور | م` + total row

---

## Branded PDF Report (`src/engine/buildReportPdf.js`)

A second, independent report generator — `downloadBrandedReportPdf(result, meta, settings,
onProgress)` — produces a print-styled `.pdf` (cover, vision/mission, methodology, per-axis
detail pages, summary table, bar-chart pages, recommendations) from the exact same `result` /
`meta` / `settings` objects `buildAnnualDocx()` consumes. It works by rendering plain HTML/CSS
off-screen (one `.pdf-page-outer` at a time, matching `.pdf-page` classed template strings + the
`PDF_CSS` string), then rasterizing each page with `html2canvas` and assembling the pages into a
PDF with `jsPDF` — **not** `html2pdf.js`'s one-shot capture, and **not** `docx`'s RTL rendering
path, so none of the `cellAlign`/bidi flipping described above applies here: this is a normal
browser-rendered `direction: rtl` HTML table, so table cells are written in plain right-to-left
reading order in the template strings with no left/right flip needed.

`buildFixedParticipantsBlock(result)` renders the same fixed degree×department headcount table
described above (`result.fixedParticipants`) as an HTML fragment (heading + intro line + table,
no `pageOpen`/`pageClose` of its own) appended at the bottom of `buildMethodologyPage`, directly
under its "المعالجة الإحصائية المستخدمة" section — sharing that page rather than getting a page
of its own. If the combined content overflows one physical page, the existing slice-across-
multiple-PDF-pages logic in `downloadBrandedReportPdf` (see the comment on `sliceCount` there)
already handles it — `buildMethodologyPage`'s DOM node has no fixed height, so it grows to fit.
Returns `""` (nothing appended) for schemas without a fixed-participants config (student,
graduates).

---

## How to Add a New Survey Type

There are now **two** ways to add a survey type. Which one to use depends on whether it needs
to be a permanent, built-in survey or just usable without touching code.

### Option A — No-code, via the Survey Management UI (recommended for most new surveys)

Use the **🗂️ إدارة الاستبيانات** tab (`SurveyManagement.jsx`) to define the survey through the
UI (or by importing an Excel template/real response file — see "Survey Management System"
below), set its status to **نشط**, and it becomes immediately selectable/auto-detectable in
the normal upload wizard — no file changes, no rebuild. This is the path for one-off or
frequently-changing surveys. See the dedicated section below for the full data model,
storage, and analysis-engine integration details, and its current limitations (open-ended/
numeric questions aren't analyzed; likert-5 column mapping is positional).

### Option B — Built-in, via a new YAML schema file (for permanent additions to the codebase)

#### Step 1 — Create the schema file

```bash
# Create src/schemas/<id>.yaml
```

Minimum required fields: `id`, `name`, `scale`, `axes` (or flat `questions:`). See "Schema
System" above for the full YAML format. `compileSchema.js` derives `colIndex` (likert-5) /
`questionStartIndex` (likert-3) and default `interpretation` automatically — no need to write
those by hand.

If the survey has demographic cross-tabulation (job × specialization), set under `meta:`:
```yaml
meta:
  degree: [column header keyword for rows]
  dept:   [column header keyword for cols]
```
The cross-tab appears automatically in ثالثاً when both columns are detected.

#### Step 2 — Nothing to register

`src/schemas/index.js` picks up every `*.yaml` file automatically via `import.meta.glob` —
there is no manual import/registration step (unlike the old `*.json` system this replaced).

#### Step 3 — Add to the survey picker in `App.jsx`

The actual manual-override button list at step 2 of the wizard (`Object.values(allSchemas())`,
see `App.jsx` around the "النوع غير صحيح؟" block) already iterates every schema automatically
— a new YAML file appears there with no further change needed. (`SurveyTypePicker`, the
component literally named for this, is currently unused/dead code — left as-is since it's
pre-existing and out of scope.)

### Step 4 — Word report customization (optional)

If the new survey needs **non-standard رابعاً/خامساً sections**, add dedicated builder functions in `buildDocx.js` and switch on `result.schemaId` inside `buildAnnualDocx`:

```javascript
// src/engine/buildDocx.js
function buildMyDetailedSection(result) { /* ... */ }
function buildMySummarySection(result)  { /* ... */ }

// Inside buildAnnualDocx:
const isMyType = result.schemaId === "mysurvey";
const children = [
  ...
  ...(isMyType ? buildMyDetailedSection(result) : buildDetailedSection(result)),
  ...(isMyType ? buildMySummarySection(result)  : buildSummarySection(result)),
  ...
];
```

Standard surveys (likert-3 or likert-5 with no special layout) require **no changes** to `buildDocx.js`.

### Step 5 — Verify

```bash
npx vite build   # must complete with no errors
```

---

## Survey Management System (Dynamic Custom Surveys)

A second, fully independent survey-definition system, added alongside the 5 built-in YAML
surveys. Accessed via the **🗂️ إدارة الاستبيانات** header button (`showSurveyManagement`
state in `App`), which renders `SurveyManagement.jsx` in place of the wizard.

### Files

| File | Purpose |
|---|---|
| `src/engine/customSurveyModel.js` | Data model, factories (`createSurvey`/`createSection`/`createQuestion`/`createMetadata`), `localStorage` CRUD, Excel data-import detection, Excel template build/import, JSON backup export/import, and the bridge into the real analysis engine (`toAnalysisSchema`, `getAllAnalysisSchemas`, `detectAnySurveyType`) |
| `src/SurveyManagement.jsx` | All UI: survey list (CRUD/duplicate/activate-deactivate), step-by-step editor (عام → المحاور → معاينة), section/question add/edit/delete/reorder, cross-section question move, Excel-template & data-import panels, backup import/export buttons |

### Data model

```javascript
// customSurveyModel.js
{
  id, name, description, version, status: "draft" | "active",
  surveyType, scaleType: "likert-3" | "likert-5",
  metadata: { timestampCol, emailCol, nameCol, degreeCol, departmentCol, freeTextCols },  // each: string[]
  sections: [{
    id, name, description,
    questions: [{ id, text, excelColumn, type: "likert" | "text" | "numeric", required, category, weight }]
  }],
  createdAt, updatedAt,
}
```

`excelColumn` defaults to the question's own `text` if left blank — applied once, in
`saveCustomSurvey()`, so it applies regardless of whether the survey was built by hand,
imported from a data file, or imported from a template.

### Storage

`localStorage` key **`eruQA_customSurveys_v1`** (separate from `eruQA_settings_v1` /
`eruQA_ai_v1`). `localStorage` is scoped per browser **origin including port** — running the
dev server on a different `--port` each time makes saved surveys look like they vanished
(they're really just in a different bucket). Always use the same port (`npm run dev` /
`npx vite`, default port 3000 per `vite.config.js`) for surveys to persist across reopens.

As a backup/portability mechanism independent of browser storage: **⬇ تصدير نسخة احتياطية
(JSON)** downloads all custom surveys as one file; **⬆ استيراد نسخة احتياطية** merges a
backup file back in by survey `id` (`buildSurveysBackupBlob` / `importSurveysBackup`).

### Building a survey — three ways

1. **Manual, step-by-step**: the editor is a 3-step wizard (عام → المحاور → معاينة) with a
   "التالي ←" button after each step (step 1 is gated on a non-empty survey name) and a
   distinct "💾 حفظ الاستبيان" button on the final step, plus a quick "💾 حفظ" in the header
   usable from any step. Sections and questions each support add/edit/delete/reorder
   (↑/↓ buttons); questions can also be **moved to a different section** via a per-question
   "نقل هذا السؤال إلى محور آخر..." dropdown + "➡ نقل" button (`moveQuestionToSection` in
   `SurveyEditorView`).

2. **Import a real response-data file** (📥 panel inside step 1 of the editor,
   `importSurveyStructureFromRows`): reads an actual Excel export and auto-detects which
   columns are questions vs. general info (name/email/degree/department/free-text), via
   `classifyHeader()`. Detected questions all land in **one new section** — a flat header row
   carries no axis/section information (verified against real survey exports: single header
   row, no merged section markers), so multi-axis surveys need manual splitting afterward
   using the section/question tools above. Classification is guarded against misclassifying
   real long Likert questions that happen to mention a metadata keyword mid-sentence: a column
   is only treated as metadata if its header is a short label (≤ 5 words) **and** its actual
   sample answers aren't Likert-scaled (`isShortLabel` + `isLikertColumn` check in
   `classifyHeader`/`importSurveyStructureFromRows`).

3. **Import a structured Excel template** (list-view toolbar, ⬇ تحميل القالب / 📤 استيراد
   استبيان من قالب, `buildSurveyTemplateBlob` / `importSurveyFromTemplateArrayBuffer`): a
   2-sheet workbook — **معلومات الاستبيان** (name/description/version/survey type/scale, as
   label/value rows) and **المحاور والأسئلة** (محور | السؤال | اسم العمود | النوع | إلزامي |
   الفئة | الوزن). Rows are grouped into sections by their المحور value, preserving order;
   rows with a **blank المحور are grouped into one fallback section** (named "الأسئلة"), so
   surveys with no axes at all (e.g. like `graduates.yaml`) are supported directly. This is
   the fastest path for defining a brand-new survey with multiple real axes without using the
   UI's add-section/add-question buttons one at a time.

### Wiring into the analysis engine

`customSurveyModel.js`'s `toAnalysisSchema(survey)` converts a custom survey into the exact
schema shape `compileSchema.js` produces for the built-in YAML surveys (see "Schema System"
above), so `analyze()`/`analyzeRows()` — fully schema-driven, no knowledge of where a schema
came from — works on it completely unmodified:

- `getActiveCustomAnalysisSchemas()` — all `status: "active"` custom surveys, converted, keyed
  by id (surveys with zero Likert questions are skipped — nothing for `analyze()` to compute).
- `getAllAnalysisSchemas()` — `{ ...the 5 built-in YAML schemas, ...active custom surveys }`.
  Used in `App.jsx` as `allSchemas()` (see "`SCHEMAS` is no longer a static import" above) —
  a drop-in superset, so every existing built-in-schema lookup is unaffected.
- `detectAnySurveyType(filename, headers)` — duplicates `analyze.js`'s own
  hint-match-then-fuzzy-similarity algorithm, parameterized over the merged schema set instead
  of `analyze.js`'s closed-over static `SCHEMAS`. Imported in `App.jsx` as `detectSurveyType`
  (import alias) — every existing call site is unchanged textually.

**Two limitations, both consequences of reusing `analyze()` unmodified (by design — it was
never changed for this feature):**

1. **Only `type: "likert"` questions are analyzed.** `analyze.js` has no concept of "text" or
   "numeric" questions — they're filtered out of the schema handed to `analyze()` entirely, the
   same way open-ended `freeTextCols` in the built-in surveys are already excluded from
   quantitative analysis. A custom survey's text/numeric questions simply won't appear in the
   generated statistics or Word report.
2. **For `likert-5` custom surveys, column mapping is positional**, exactly like the built-in
   YAML surveys: `colIndex` is assigned in current section/question order inside
   `toAnalysisSchema`. Reordering questions/sections in the editor changes which file column
   each question reads from. **`likert-3` custom surveys are unaffected by reordering** —
   `analyze.js` fuzzy-matches likert-3 questions against the actual header text, not position.

### Excel template format (reference)

| Sheet | Columns |
|---|---|
| معلومات الاستبيان | الحقل \| القيمة — rows: اسم الاستبيان, وصف الاستبيان, الإصدار, نوع الاستبيان, نوع المقياس (`"...ثلاثي..."` → `likert-3`, anything else → `likert-5`) |
| المحاور والأسئلة | المحور \| السؤال \| اسم العمود في Excel \| نوع السؤال (مقياس/نص/رقمي) \| إلزامي (نعم/لا) \| الفئة \| الوزن |

---

## Settings Persistence

User settings (institution name, fonts, colors, signatures) are stored in `localStorage` under key `eruQA_settings_v1`. Default values are exported from `buildDocx.js` as `DEFAULT_SETTINGS`.

AI assistant settings (provider, model, API key) are stored under `eruQA_ai_v1`.

---

## Google Drive Integration (Step 0)

Step 0 of the wizard has two tabs: **📁 رفع ملف** (local upload) and **☁️ Google Drive**.

### How it works
- Pure browser-side OAuth2 via **Google Identity Services (GIS)** — no backend needed.
- GIS script loaded lazily from `https://accounts.google.com/gsi/client`.
- `google.accounts.oauth2.initTokenClient` obtains an access token in a popup.
- Token is used directly with **Google Drive REST API v3** via `fetch` (no GAPI library).
- Google Sheets files are exported as `.xlsx` via the `/export` endpoint.

### Key constants in `App.jsx` (top of file)
```js
const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID ?? "";
const DRIVE_SCOPE      = "https://www.googleapis.com/auth/drive.readonly";
```

### Environment variable
Stored in `.env.local` (not committed):
```
VITE_GOOGLE_CLIENT_ID=<your-client-id>.apps.googleusercontent.com
```
Restart the dev server after changing this file.

### Google Cloud Console setup (one-time)
1. Enable **Google Drive API** on the project.
2. Create an **OAuth 2.0 Client ID** (Web application type).
3. Add `http://localhost:3000` to **Authorized JavaScript Origins**.
4. For production, also add the production domain.

### Account picker
`requestAccessToken({ prompt: "select_account" })` is called every time **`connectDrive()` runs** (an explicit, user-clicked connect) so the user always sees the account picker there — Google never auto-selects a cached account for that path. This is no longer the *only* way a token gets set, though — see "Staying signed in" below.

### Staying signed in (token persistence + silent refresh)

Originally, every GIS token client in this app (`connectDrive()` here, and the Semester Survey
Generator's own separate token client) forced the account-picker popup on every single visit,
since a plain OAuth2 implicit flow with no backend has no refresh token — the ~1hr access token
just evaporated on reload with nothing to fall back on. Two client-side-only mitigations, added
to both `App.jsx`'s Drive flow and `semesterSurveyModel.js`'s Semester flow identically:

1. **Persist the still-valid token.** `saveStoredToken(key, accessToken, expiresInSec)` /
   `getStoredToken(key)` / `clearStoredToken(key)` (all in `semesterSurveyModel.js`, imported
   into `App.jsx` too since the Drive flow needed the exact same behavior) wrap
   `localStorage` with an expiry timestamp (60s safety margin). `driveToken`/the Semester
   `useSemesterAuth` hook's `token` both now **initialize from `getStoredToken(...)`** instead
   of `null` — a reload within the token's lifetime needs zero Google calls at all.
2. **Silent refresh before ever showing a popup.** Once on mount, if there's no valid stored
   token, a `requestAccessToken({ prompt: "" })` (empty prompt = silent, no visible UI) is
   tried in the background — this succeeds without any popup if the browser still has an
   active Google session and the user previously consented to the scope. Only if *that* fails
   (or on an explicit user click) does the visible "الاتصال بـ Google" / "ربط Google Drive"
   button/popup ever appear. A `explicitRef`/`explicitDriveRef` ref (flipped right before each
   `requestAccessToken` call) tells the shared `onError` callback whether to surface an error
   banner — a silent attempt failing is expected/normal (no prior session yet) and must stay
   invisible, while an explicit connect failing should tell the user.

`App.jsx`'s "قطع" (disconnect) button now also calls `clearStoredToken(DRIVE_TOKEN_KEY)`, and a
token restored straight from `localStorage` (bypassing `connectDrive()`'s callback, which
normally triggers the file fetch) is picked up by a small effect that fetches the Drive file
list once the Drive tab is actually opened. The Semester flow's `SemesterSurveys.jsx` and
`SemesterFormPicker.jsx` both read/write the **same** `SEMESTER_TOKEN_KEY` — connecting in
either one covers the other for the rest of that token's lifetime, since they share scope and
client ID.

---

## Semester Survey Generator (`src/SemesterSurveys.jsx`)

A third top-level tab alongside **🗂️ إدارة الاستبيانات** and **⚙ الإعدادات** (`showSemesterSurveys`
state in `App`, header button **📆 استبيانات الفصل الدراسي**, mutually exclusive with the other
two — same toggle-and-clear-the-others pattern). Generates per-semester copies of Google Forms
templates directly on Google Drive and reports response stats — **no database, no localStorage
cache**; Drive (via the Drive + Forms REST APIs) is the only source of truth, read fresh on every
load/filter-change/refresh.

### Files

| File | Purpose |
|---|---|
| `src/engine/semesterSurveyModel.js` | OAuth (own GIS token client + scope), Drive REST helpers (list/copy/move/create-folder), Forms REST helpers (get form, list responses), naming/util helpers. No React. |
| `src/SemesterSurveys.jsx` | UI: tab bar (إنشاء استبيانات / لوحة المتابعة), auth gate, `GenerateSurveysView`, `DashboardView`, a local toast stack, and CSS-div bar charts (no chart library, mirrors `DriveDashboard`'s bar-chart technique). |

### Why a separate OAuth scope/token client

The existing Drive integration (see above) requests `drive.readonly` — enough to browse/download
survey response files, but not to copy, rename, move, or create folders. This feature needs
broader access plus Forms access, so `semesterSurveyModel.js` runs its **own** `initTokenClient`
with scope:
```
https://www.googleapis.com/auth/drive
https://www.googleapis.com/auth/forms.body
https://www.googleapis.com/auth/forms.responses.readonly
```
kept fully separate from `App.jsx`'s `DRIVE_SCOPE`/`tokenClientRef` so the existing upload/Drive-
dashboard flow is untouched. OAuth is **client-side only** (Google Identity Services token-client
popup, same as the existing Drive integration) — no `GOOGLE_CLIENT_SECRET`/server-side code
exchange. A secret baked into a static SPA bundle isn't actually secret, and this app has no
server-side token storage to make a refresh-token flow worthwhile.

### Why there's no "linked response Sheet"

The Google Forms REST API has no way to read or attach the classic "responses saved to this
Google Sheet" link — that's an Apps-Script-only capability (`FormApp.getDestinationId()`), and a
form copied via `Drive.files.copy` loses any existing Sheet link entirely. So this feature never
touches the Sheets API: response counts and "آخر رد" come straight from the Forms API's
`forms.responses.list` (paginated via `nextPageToken` in `listAllResponses`), and every "فتح
الردود" action deep-links to the form's own built-in Responses tab —
`editorResponsesUrl(formId)` → `https://docs.google.com/forms/d/{formId}/edit#responses` — not a
spreadsheet.

### Environment variables

Two new `VITE_*` vars (same module-scope-const pattern as `GOOGLE_CLIENT_ID`, read in
`semesterSurveyModel.js`), added to `.env.local` (gitignored) — see the new `.env.example` at the
repo root for the full list with `GOOGLE_CLIENT_ID` reused:
```
VITE_GOOGLE_TEMPLATE_FOLDER_ID=       # Drive folder containing the Form templates
VITE_GOOGLE_ROOT_SURVEYS_FOLDER_ID=   # Root folder; {year}/{semester} subfolders are created under it
```
If either is unset, `SemesterSurveys` shows a warning card instead of the auth gate — it never
silently no-ops.

### Generate Surveys flow

1. On mount (after Drive connect), `listFormsInFolder` lists every `application/vnd.google-apps.form`
   file directly inside `VITE_GOOGLE_TEMPLATE_FOLDER_ID` — never hardcoded, so any form dropped
   into that folder appears automatically. Each gets a "تحديد" checkbox + **Select All**.
2. Each selected template also has a **generation-mode `<select>`**: **نسخة عامة** (default —
   one general copy), **الأقسام** (one copy per *department*, ignoring any programs it has),
   or **الأقسام والبرامج** (one copy per *program* for departments that have programs, one
   copy per department for those that don't) — `t.mode` is `"general" | "departments" |
   "programs"`, and `"departments"`/`"programs"` map directly onto
   `expandDepartmentUnits(departments, granularity)`'s second argument in
   `semesterSurveyModel.js`. The department list itself comes from `loadDepartments()`
   (localStorage-backed, default seed `DEFAULT_DEPARTMENTS`: محاسبة، اقتصاد، علوم سياسية have
   no programs; إدارة أعمال has مالية/تسويق; تكنولوجيا الأعمال has BA/MIS/Fintech/MKI) — editable
   at runtime via the **⚙ الأقسام والبرامج** tab (`DepartmentsView`, works without a Google
   connection since it's pure local config, not Drive data). Each generated copy is named
   `"{template} - [{department}[ - {program}]] - {semester} - {year}"` via
   `buildSurveyName(...)`; `departmentFromSurveyName(...)` parses that middle unit back out
   generically (whatever's between the first and last-two `" - "`-joined segments), used for
   report metadata in Analyze All (see below).
3. Academic Year — a free-text input backed by a `<datalist>` (`ssg-year-options`) populated
   from existing year folders under `VITE_GOOGLE_ROOT_SURVEYS_FOLDER_ID` (fetched once via
   `listSubfolders` on mount) — pick an existing year or type a new one; validated against
   `/^\d{4}\/\d{4}$/` (e.g. `2026/2027`) either way. Semester dropdown (`SEMESTERS`: خريف /
   ربيع / صيف).
4. On **إنشاء**: expands each selected template — via `expandDepartmentUnits` per its own
   generation-mode select, or as a single general job if left on "نسخة عامة" — into a flat job list,
   creates/reuses the `{year}` folder, then the `{semester}` subfolder inside it, then **one
   subfolder per selected survey type** inside the semester folder (named after the template —
   created once per distinct template via a single `Promise.all` pass over the *unique* selected
   templates, before the job batch starts, so per-department jobs from the same template share
   one folder instead of racing to create duplicates). Final structure:
   `{root}/{year}/{semester}/{templateName}/{copy...}`. Jobs then run through `runBatched`
   (batches of 4 via `Promise.all` — the exact same batching idiom `DriveDashboard.loadStats`/
   `BatchProcessor` already use elsewhere in `App.jsx`, reused rather than reinvented). Each job:
   `copyFile` → `moveFile` (into its template's subfolder) → `updateFormTitle` → `publishForm` →
   `getForm` (to get the public `responderUri`). `updateFormTitle` matters because
   `Drive.files.copy` only renames the Drive **file**, not the form's own internal title (what
   respondents actually see when they open it) — it's set separately via `forms.{id}:batchUpdate`
   / `updateFormInfo{info:{title},updateMask:"title"}` so both stay in sync with the same
   `buildSurveyName(...)` string. `publishForm` calls `forms.{id}:setPublishSettings` with
   `{isPublished:true,isAcceptingResponses:true}` — **as of 2026-06-30, forms created via the
   API start unpublished and reject responses until explicitly published**, so every generated
   copy is published immediately; without this step the survey link would look fine but silently
   refuse to accept any responses until someone opened it in the Forms UI and published it by
   hand. Per-job status renders live (○ → ⏳ → ✔/✖).
5. Successful jobs feed a results table: Open Form / Open Form Responses / Copy Form Link / Copy
   Responses Link, plus a success/failure toast.

### Dashboard flow

Academic Year / Semester filters are populated by listing Drive **subfolders** (`listSubfolders`)
of `VITE_GOOGLE_ROOT_SURVEYS_FOLDER_ID`, then of the selected year folder — never hardcoded.
Changing a filter (or clicking **↻ تحديث**) walks one level deeper than the filters: it lists the
survey-type subfolders inside the resolved semester folder, then the form(s) inside each, and
tags every row with its `surveyType` (shown as a "النوع" column in the table) — mirroring the
`{year}/{semester}/{templateName}/{copy}` structure created by Generate Surveys. Then, in
batches of 4, fetches `getForm` + `listAllResponses` per form to compute stats client-side. No
caching anywhere — every load/filter-change/refresh is a fresh set of API calls, per the "Drive
only, no DB, no cache" constraint. Stat cards (total surveys, total responses, avg/survey, last
response received), a CSS-div bar chart of responses-per-survey (`BarChart` in
`SemesterSurveys.jsx` — same width-percentage/gradient-bar technique as `DriveDashboard`, no
chart library dependency), and a table with a per-row **↻ تحديث** action that re-fetches just
that one survey's stats.

### Progress feedback

Both long batch operations (Generate's job loop, Analyze All's per-survey loop) render a shared
`ProgressBar` (done/total count + animated striped fill) plus a rotating status line
(`useRotatingTip`, cycling through a short list of Arabic phrases every ~2.6s — `GENERATE_TIPS`/
`ANALYZE_TIPS`) so a multi-step wait doesn't look frozen. Both sit above the existing per-item
○/⏳/✔/✖ status list rather than replacing it — the tip line is deliberately generic/non-literal
("جاري نسخ النماذج...") since the two loops don't expose a truly step-by-step public API to
narrate against.

### Exporting links (`📋 نسخ الروابط كرسالة` / `⬇ تنزيل Excel`)

Two shared bulk-export buttons (`LinksExportButtons` in `SemesterSurveys.jsx`) appear wherever a
list of generated surveys is shown — the results table at the end of a Generate run, and the
Dashboard's survey table — each acting on whatever rows are currently visible there (a fresh
generate batch vs. every survey in the selected year/semester, respectively):
- **نسخ الروابط كرسالة**: copies a plain-text numbered list (name + form link + responses link
  per survey) to the clipboard via the same `copyToClipboard` helper the per-row "نسخ رابط..."
  buttons already use — ready to paste into an email/chat.
- **تنزيل Excel**: builds a 3-column `.xlsx` (الاستبيان / رابط النموذج / رابط الردود) with
  `XLSX.utils.json_to_sheet` + `XLSX.write` (the `xlsx` package the app already depends on for
  reading uploads — this is the first place in the codebase that *writes* one) and downloads it
  via the same `downloadBlob` helper.

### Analyze All (`🔍 تحليل الكل`)

Runs every currently-listed survey through the app's **existing, unmodified** analysis engine
and generates a Word report per survey — no separate analysis path was written for Forms data.
The bridge is `responsesToRows(form, responses)` in `semesterSurveyModel.js`: it converts a
Forms API `form` (its `items[]`) + `responses[]` into the exact `[header, ...rows]` shape
`analyze.js` already expects from a parsed Excel export (`Timestamp`, optionally `Email Address`
if the form collects it, then one column per question in form order). This works because Google
Forms API answers come back as plain option text (e.g. `"أوافق بشدة"`), which is exactly the
plain-label fallback `parseResponse5`/`parseResponse3` already handle for custom surveys (see
"Response Value Parsing" above) — so `detectAnySurveyType(survey.name, rows[0])` +
`analyze(rows, schema)` (both imported straight from `analyze.js`/`customSurveyModel.js`) run
unchanged. Surveys whose type can't be auto-detected are skipped with a toast, not force-matched.

Clicking the button shows a small shared-fields form (اسم المعد / المراجع — **not** persisted
anywhere, entered fresh each run, matching how the rest of the app already collects these two
fields per report) before running. `meta.program` is recovered from the survey's own Drive file
name via `departmentFromSurveyName()` (parses the `" - "`-joined `buildSurveyName()` format —
4 parts means a per-department copy, 3 means general). Report branding (institution name, logo,
signatures, etc.) comes from the same `eruQA_settings_v1` localStorage key `SettingsPanel`
already writes — read via a small local `loadReportSettings()` duplicate of `App.jsx`'s
`loadSettings()`, since `SemesterSurveys.jsx` has no prop access into `App`'s state. Each
successful report downloads immediately via a local `downloadBlob()` (same tiny helper duplicated
in `SurveyManagement.jsx`/`App.jsx`) with a **450ms delay between downloads** — the same delay
`BatchProcessor.downloadAll` already uses in `App.jsx`, needed because browsers throttle/block
several near-simultaneous auto-downloads.

### Google Cloud Console setup (in addition to the Drive integration's setup above)
1. Enable the **Google Forms API** on the same project (`redaa2`).
2. If the OAuth consent screen is in "Testing" mode, add the `forms.body` and
   `forms.responses.readonly` scopes to it (and add test users) — otherwise the consent popup
   will reject the broader scope request.
3. The signed-in Google account must have **edit access** to both the configured template folder
   and the root surveys folder.

---

## Course Evaluation Hub (`src/CourseEvaluationHub.jsx`)

A fourth top-level tab (`showCourseEval` state in `App`, header button **📚 تقييم المقررات**,
mutually exclusive with **🗂️ إدارة الاستبيانات** / **📆 استبيانات الفصل الدراسي** /
**⚙ الإعدادات** — same toggle-and-clear-the-others pattern). Groups every tool used to run a
full course-evaluation cycle behind its own internal sub-tab bar (`activeTab` state local to
the hub, not lifted into `App`): **📖 دليل الاستخدام** (a static ordered guide tying the other
four together, with a "افتح تاب..." jump button per step that just calls `setActiveTab`),
**📋 قالب بيانات المقررات**, **📊 أداء الاستبيانات**, **🧩 تقسيم التقييم**, and
**📄 مراجعة التوصيات**. Each sub-tool is its own file, rendered conditionally by the hub —
none of them know they're inside a hub (no shared state, no props from `App`), same
standalone-file pattern as `SurveyManagement.jsx`/`SemesterSurveys.jsx`.

This replaced a previous standalone top-level tab for just the splitter (`showCourseSplitter`,
button **🧩 تقسيم تقييم المقررات**) — the splitter's own code/logic is untouched, only its
entry point moved from the header into the hub's sub-tab bar.

### Course Splitter (`src/CourseSplitter.jsx`)

Fully independent of the survey wizard/analysis engine — it's a self-contained utility ported
from a standalone HTML tool (`course_eval_splitte V3r.html`) that a UMIS export needs run
through *before* any of those files can be analyzed as individual surveys, since a single UMIS
"course evaluation" report export bundles every course's results into **one Excel sheet with
repeating block headers**, not one file per course.

### Marker text is user-editable, not hardcoded

The block marker (`DEFAULT_MARKER`, default `"بنود الاستبيان"`) and `BLOCK_OFFSET` (rows
between a course's data start and its marker row, default `10`) now match the verified-working
standalone tool (`course_eval_splitte V3r.html`) byte-for-byte — an earlier port had
`DEFAULT_MARKER` as `"بيانات الاستبيان"` (reconstructed from a garbled copy-paste), which matched
zero rows in real UMIS exports and made the splitter silently produce no courses; fixed by
copying the exact string from the standalone tool's source. Different UMIS report
templates/versions may still use different text or row offsets, so both remain exposed in an
"⚙ إعدادات متقدمة" collapsible in step 3, not hardcoded constants a user would need a code change
to fix. If the configured marker matches zero rows, `findMarkerCandidates(wb)` scans the sheet
for other strings that repeat a plausible number of
times (2–200×) and surfaces them as clickable suggestions, so a user can find the right marker
without opening the file in Excel to hunt for it manually.

### What it does

1. **Splits** a merged UMIS RDLC-format Excel export into one `.xlsx` workbook per course, by
   scanning every row for a repeating marker string (`"بيانات الاستبيان"`) that starts each
   course's block, then slicing the sheet (including remapped cell merges) into per-course
   ranges. The course code/title is recovered by regex-matching a `"... (CODE)"` pattern in the
   few rows just above each marker.
2. **Matches** each split-out course against an optional **reference course list** (columns
   `COURSE_CODE` / `COURSE_DESCR_EN` / `COURSE_DESCR_AR`) — this is what lets it report which
   expected courses are **missing** (no evaluation file was found for them at all), not just
   split what's present.
3. **Classifies by department**, from an optional **department-distribution list** (columns
   اسم المقرر / كود المقرر / القسم العلمي), matched **by course name first** (exact, then
   Levenshtein-similarity fuzzy ≥0.8), falling back to exact code match — chosen in that order
   because course codes are observed to vary between UMIS and other systems while names don't.
4. **Flags duplicates** — the same course code appearing in more than one uploaded merged file.

All three inputs (reference list, department list, one-or-more merged files) are independent
optional/required uploads — only the merged file(s) are required to run a split at all; the
other two only add matching/classification on top.

### Output

- Per-course "⬇ تحميل" button — downloads that one course's split-out `.xlsx`.
- "⬇ تحميل كل المواد (ZIP)" — zips every split course via `jszip`, foldered by department name
  when a department list was supplied (courses with no confident department match land in a
  `"غير محدد - يحتاج مراجعة"` folder so they're never silently dropped), flat otherwise.
- A results table + filter row (الكل / تم تقسيمها / لم تُرفع بعد / مكررة) and a summary stat
  bar (matched/total against the reference list, missing count, duplicate count, no-department
  count, fuzzy-match count) — same visual language (`.mini-table`, `.card`, stat tiles) as the
  rest of the app, not a re-skinned copy of the standalone tool's own dark-panel CSS.

### Why it's a separate file, not folded into `analyze.js`

The splitter's "rows" are a raw UMIS RDLC layout (merged cells, a fixed header block, blocks of
arbitrary length keyed off a marker string) — nothing like the `[header, ...dataRows]` shape
`analyze.js`/`readExcel()` expect from a Google Forms export. It also writes new `.xlsx`
workbooks (via `XLSX.write`) rather than only reading them, and is the second place in the
codebase (after the Semester Survey Generator's Excel-link export) to do so. Kept fully
standalone — no shared state, no props from `App` — exactly like `SurveyManagement.jsx` and
`SemesterSurveys.jsx`.

### Course Template Tool (`src/CourseTemplateTool.jsx`)

Manages the per-course assignment sheet handed out before an evaluation cycle: اسم المقرر /
كود المقرر / عضو هيئة التدريس / عضو الهيئة المعاونة / القسم العلمي / هل يوجد لاب (نعم/لا) /
القائم بالمراجعة. **⬇ تحميل قالب فارغ** downloads just the header row; **📤 رفع ملف موجود**
re-imports a filled sheet (`mapHeaders()` matches uploaded headers by normalized text — parens
like `"(نعم-لا)"` stripped before comparing — falling back to plain column position for
unrecognized headers) into an editable `<table>` where every cell is a live `<input>` (the
"هل يوجد لاب" column is a tri-state toggle button instead, cycling نعم → لا → empty, so the
insights below can rely on an exact value rather than free text). **⬇ تنزيل نسخة معدّلة**
re-exports the current in-memory rows. Insights (`useMemo` over `rows`, no persistence) cover
exactly what was asked for: توزيع حسب القسم (course + lab count per department, CSS-div bar),
عبء العمل لكل عضو (per-instructor/assistant/reviewer load, same bar technique), and بيانات
ناقصة / تكرار (rows missing اسم المقرر/كود المقرر/القسم العلمي, and rows whose كود المقرر
repeats).

### Survey Participation Tool (`src/SurveyParticipationTool.jsx`)

Reads the system's survey-performance export — columns `COURSE_CODE`, `COURSE_DESCR_EN`,
`NoOfVotes` — where `NoOfVotes` is a **"voted/total" fraction string** (e.g. `"0/25"`, not a
plain count). `parseVotes()` splits on `/`; if no `/` is present it falls back to treating the
whole value as `voted` with `total: null` rather than failing outright. Each course is
classified into one of four states (`classify()`): `total === 0` → **لا يوجد طلاب مسجلين**
(the `0/0` case — no one was ever enrolled to evaluate this course), `voted === 0 && total > 0`
→ **لم يتم التقييم**, `0 < voted < threshold` → **مشكوك في انتظامها** (needs a manual check for
whether the course is actually a regular one — `threshold` defaults to 10 but is a plain
number input, not hardcoded), otherwise **طبيعية**. Stat tiles + a status filter row +
per-category "⬇ تنزيل Excel" buttons (تم التقييم / لم يتم التقييم / مشكوك فيها / تقرير كامل)
cover the three exportable reports asked for, plus a combined one with a status column.

### PDF Recommendation Reviewer (`src/PdfRecommendationReviewer.jsx`)

Ported from the standalone `pdf_rec_reviewer.html`. Renders the **last page** of every PDF in a
user-picked folder (`webkitdirectory` input, Chrome/Edge-only — same constraint the standalone
tool had) as a thumbnail — that's where a reviewer's recommendation section typically lives —
via a 4-way concurrent render pool (`CONCURRENCY = 4`, identical batching idea to the
`Promise.all`-batches-of-4 pattern already used in `DriveDashboard`/`SemesterSurveys`/
`BatchProcessor`). Each card is a checkbox toggle for "has a recommendation"; **📊 تصدير Excel**
writes the checked subset (`#`, اسم المقرر, اسم الملف) via `XLSX`. Unlike the original tool
(which loaded pdf.js off `cdnjs.cloudflare.com`), this uses the npm `pdfjs-dist` package with
its worker resolved through Vite's `?url` import (`import pdfjsWorker from
"pdfjs-dist/build/pdf.worker.min.mjs?url"`) so it works in a production build with no external
CDN dependency. Visual language switched from the standalone tool's own dark-panel CSS to the
app's shared `.card`/`.btn` classes and accent color, same principle already applied when
`CourseSplitter.jsx` was ported.

---

## FileSlot — Comparison Mode Upload

`FileSlot` is the per-year upload widget used in comparison mode (step 1). Each slot holds `{ year, fileName, result, _file, _type, _program }`.

### Auto-detection on upload
When a file is chosen (local or Drive), the slot immediately reads the file content:
- **Local:** `f.arrayBuffer()` → `readExcel()` → `detectSurveyType(name, rows[0])` — result stored in `slot._type`
- **Drive:** `res.arrayBuffer()` → same detection chain — also creates a `File` object for later processing

`detectProgramFromFilename(name)` also runs and stores result in `slot._program`.  
`detectYearFromFilename(name)` runs and auto-fills `slot.year` if detected.

### Per-slot type badge + manual override
After a file is loaded, the slot shows:
- A green badge with the detected type label (or "⚠ نوع غير محدد" in amber if unknown)
- The program name badge (if detected)
- A row of small type-selector buttons ("النوع غير صحيح؟ اختر:") — clicking sets `slot._type` and updates the badge immediately

The override buttons appear in **both** local upload tab and Drive tab.

### Inline validation warnings (live, before clicking تحليل)
A computed warning block renders in the comparison step 1 JSX on every render:
- Duplicate academic years across slots → yellow warning
- Different survey types across slots → yellow warning  
- Different programs detected across slots → yellow warning (soft, not a block — detection from filenames can be imprecise)

### `handleProcessMulti` validation (hard block)
Before reading any files:
1. **All slots must have a file** — else error "ارفع ملفاً لكل سنة."
2. **Unique years** — if any two slots share the same year string → error
3. **Same survey type** — if all slots have a `_type` and they differ → error with type names listed

After processing: calls `setSurveyType(firstDetectedType)` so the schema label is correct in results. Also auto-fills `meta.program` from the first slot that has a program detected.

### Comparison output filename
`downloadComparison()` builds: `مقارنة_{schema.label}_{program}_{y1}_و_{y2}.docx`  
Years come from `comparison.slots.map(s => s.year).join("_و_")` — not from `meta.year`.

---

### Drive filters (client-side, no extra API calls)
Three `<select>` dropdowns filter the `driveFiles` array already in memory:
- **نوع الاستبيان** — populated from `Object.values(allSchemas())` labels (built-in + active custom surveys)
- **السنة الدراسية** — dynamically built from filenames via `detectYearFromFilename()`
- **البرنامج** — static list in `PROGRAMS` constant: محاسبة، اقتصاد، إدارة، علوم سياسية، تكنولوجيا أعمال

Helper functions (defined after `detectProgramFromFilename`, before `uniqueCol`):
- `detectYearFromFilename(filename)` — regex `/(\d{4}[-_]\d{4}|\d{4}[-_]\d{2})/`, normalises `2024_25` → `2024-2025`
- `detectTypeHintFromFilename(filename)` — matches against `schema.fileHints` with Arabic normalisation

### Drive Dashboard (`DriveDashboard` component)
Three-way toggle in the Drive tab's connected-bar — **📁 قائمة** (flat list with filters),
**📊 لوحة** (dashboard), **🗓️ الفصل الدراسي** (semester-survey picker, see below) — via
`driveViewMode` state (`"list" | "dashboard" | "semester"`; was a `driveDashboard` boolean
before the semester picker was added).

`DriveDashboard` props: `{ files, token, onSelectFile }`. Internal state: `stats[]`, `loading`, `done`.

On-demand analysis: clicking "📊 تحليل الملفات (N)" downloads files in batches of 4 via `Promise.all`, parses with `readExcel()`, counts `rows.length - 1`. UI shows:
1. Summary cards by survey type (file count + response count)
2. CSS bar charts — distribution by program and by academic year
3. Detailed table with per-file type/year/program/responses and "تحليل" button

### Semester-survey picker (`SemesterFormPicker.jsx`)

The **🗓️ الفصل الدراسي** view mode above lets the single-file wizard's Drive tab analyze a
survey generated by the Semester Survey Generator (see that section) directly — browsing
`{root}/{year}/{semester}` (year/semester `<select>`s populated via `listSubfolders`, exactly
like the Generator's own Dashboard) and listing every form across that semester's survey-type
subfolders, tagged with its type. It needs its **own** OAuth token (a local
`useSemesterAuth()` duplicate, same as `SemesterSurveys.jsx`'s) because the wizard's existing
`driveToken` is `drive.readonly` only and can't read Forms responses.

Selecting a form fetches `getForm` + `listAllResponses`, converts them via
`responsesToRows()` (the same Forms→rows bridge Analyze All uses), and calls
`onFormSelected(rows, name, department)` — wired in `App.jsx` to a new
`handleDriveFormSelect`, which is `handleDriveFileSelect`'s exact tail (detect type → animate
`DRIVE_STEPS` → `setStep(2)`) minus the download step (already "downloaded" via the Forms API
by the time it's called, so step 0 renders `"done"` immediately instead of animating). From
step 2 onward the survey flows through the completely unmodified normal wizard (preview,
metadata, Word report) — no special-casing needed past this handoff.

### State variables (inside App component)

**Wizard core:**
| State | Purpose |
|---|---|
| `step` | Current wizard step 0–5 |
| `mode` | `"annual"` / `"compare2"` / `"compare3"` |
| `surveyType` | Active schema key (e.g. `"faculty"`) |
| `meta` | `{ year, program, preparedBy, reviewer }` |
| `processing` | Spinner for Word generation / data preparation |
| `error` | Error message string |
| `showSettings` | Settings panel visible |
| `showSurveyManagement` | Survey Management tab visible (mutually exclusive with `showSettings`; see Survey Management System section) |
| `showCourseEval` | Course Evaluation Hub visible (mutually exclusive with the above; see Course Evaluation Hub section) |
| `showTutorial` | Tutorial overlay visible |

**File & detection:**
| State | Purpose |
|---|---|
| `singleFile` | Uploaded `File` object (annual mode) |
| `singleResult` | Analysis result for annual report |
| `rawRows` | 2D array from `readExcel()` before analysis |
| `detectedAutoType` | Schema key auto-detected from filename/headers |
| `procSteps` | `null` or `{label, status}[]` driving `ProcessingOverlay` |
| `procFile` | Filename shown in `ProcessingOverlay` |
| `dragging` | Drag-over state for upload zone |

**Batch mode:**
| State | Purpose |
|---|---|
| `batchMode` | Multi-file batch processing active |
| `batchFiles` | Array of `File` objects for batch |
| `batchYear` | Academic year for batch reports |

**Data preview (Step 3):**
| State | Purpose |
|---|---|
| `singleHeaders` | Column headers after `prepareData()` |
| `singleAllRows` | All data rows |
| `singleSchema` | Active schema object |
| `singleMetaCols` | Metadata column indices to exclude |
| `singleRemoved` | `Set` of row indices manually excluded |
| `singleFilters` | `{ dept, degree }` filter values |

**Drive:**
| State | Purpose |
|---|---|
| `uploadTab` | `"upload"` or `"drive"` — active tab in Step 0 |
| `driveToken` | Current OAuth access token |
| `driveFiles` | File list from Drive REST API |
| `driveSearch` | Search string for Drive file list |
| `driveLoading` | Spinner while fetching Drive file list |
| `driveSelected` | Currently highlighted Drive file |
| `driveProcessing` | True while downloading + parsing |
| `driveFilterType` | Survey type filter |
| `driveFilterYear` | Academic year filter |
| `driveFilterProgram` | Program filter |
| `driveViewMode` | `"list"` / `"dashboard"` / `"semester"` — which of the three Drive tab views is active |
| `tokenClientRef` | `useRef` holding GIS token client |

**Comparison mode:**
| State | Purpose |
|---|---|
| `slots` | Array of `{ year, fileName, result, _file, _type, _program }` — `_type`/`_program` auto-detected on upload |
| `slotCount` | 2 for compare2, 3 for compare3 |
| `comparison` | Built comparison object from `buildComparison()` — also has `.slots` for year list |

### Key functions in `App.jsx`
- `connectDrive()` — loads GIS script, initializes token client, triggers account picker popup
- `fetchDriveFiles(token)` — fetches spreadsheet files from Drive REST API
- `handleDriveFileSelect(file)` — shows processing overlay step 0 active during real network fetch, then animates remaining steps, then advances to **step 2**
- `handleFileSelected(file)` — reads file, then animates all 5 LOCAL_STEPS, then advances to **step 2**
- `handleValidationConfirm()` — called from step 2; runs `prepareData()` and advances to step 3
- `handleProcessMulti()` — called in comparison mode from step 1; validates unique years + same survey type across all slots, then reads all files and calls `buildComparison()`, then `setStep(4)`
- `runStepAnim(labels, startAt, msPerStep, onDone)` — drives the `ProcessingOverlay` by advancing `procSteps` on a setTimeout chain
- `downloadAnnual()` — generates annual Word doc; filename: `تقرير_{schema.label}_{program}_{year}.docx`
- `downloadComparison()` — generates comparison Word doc; filename: `مقارنة_{schema.label}_{program}_{year1}_و_{year2}.docx` (years from `comparison.slots`)

---

## Tutorial Overlay (`TutorialOverlay` component)

A `?` circular button in the header (right side, before ⚙) opens a 6-slide how-to guide. Controlled by `showTutorial` state in `App`.

### Slide data (`TUTORIAL_SLIDES` constant — module-level array)
Each object: `{ icon, color, title, subtitle, body, tips[] }`. Six slides:

| # | Title |
|---|---|
| 1 | مرحباً — overview of app capabilities |
| 2 | رفع الملف — upload from device or Google Drive |
| 3 | التحقق من النوع — auto-detection, manual override |
| 4 | نوع التقرير — annual / 2-year / 3-year comparison |
| 5 | معاينة البيانات — row-level review and deduplication |
| 6 | البيانات والنتائج — metadata entry + Word download |

### Navigation
- Left/right `‹`/`›` arrow buttons (`.tut-nav` class)
- Clickable progress dots (`.tut-dot` class) — active dot scales up and takes `slide.color`
- Last slide shows "ابدأ الآن ✓" button instead of `›`
- Clicking the backdrop (`e.target === e.currentTarget`) closes the overlay
- `✕` close button (top-left of card)

### CSS keyframes (in `CSS` string)
- `tutSlideIn` — entrance animation for the whole card
- `tutSlideNext` / `tutSlidePrev` — directional slide for slide content (applied via `.tut-slide-next` / `.tut-slide-prev` classes toggled on `idx` change using `requestAnimationFrame`)

### Accent bar
A 4px colored `div` at the top of the card uses `slide.color` (each slide has its own color).

---

## Processing Overlay (`ProcessingOverlay` component)

Replaces the old `LoadingOverlay` for file detection. Shows a full-screen backdrop with a step-by-step checklist animation.

**Step constants** (defined outside `App`, above the component):
```js
const LOCAL_STEPS = ["قراءة البيانات", "تنظيف البيانات", "كشف نوع الاستبيان", "حساب المؤشرات", "تجهيز العرض"];
const DRIVE_STEPS = ["تحميل الملف من Drive", "قراءة البيانات", "تنظيف البيانات", "كشف نوع الاستبيان", "تجهيز العرض"];
```

**Step statuses:** `"pending"` (gray ○) → `"active"` (spinning ⟳, pulsing border) → `"done"` (green ✓ + "تم" badge)

**Animation flow:**
- Local file: all 5 steps animate at 520 ms/step (~2.6 s total), then `setStep(2)`
- Drive file: step 0 stays `"active"` during the real network fetch; after fetch, `runStepAnim(DRIVE_STEPS, 1, 400, ...)` animates remaining 4 steps at 400 ms/step, then `setStep(2)`

**CSS keyframes added to the `CSS` string:**
- `fadeInUp` — card entrance
- `stepPop` — brief scale bounce when a step turns done
- `activePulse` — glow pulse on the active step icon
- `progressFlow` — animated gradient on the progress bar

The overlay also shows a progress bar with percentage (`doneCount / total * 100`) and the filename.

---

## Layout

- **Outer container:** `maxWidth: 1400`, `padding: "28px 32px"` (was 980 / 32px 24px)
- **Step 0 card:** When `uploadTab === "drive" && driveToken` is set, the folder emoji + title + description are **hidden** to save vertical space, card padding shrinks to `"24px 32px"`, and the source tabs stretch to full width
- **Drive tab content wrapper:** `maxWidth` is `"100%"` when connected, `520` when not connected
- **Upload zone:** `maxWidth: 640` (was 520)

---

## Key Dependencies

| Package | Purpose |
|---|---|
| `xlsx` | Parse `.xlsx`/`.xls`/`.csv` files |
| `docx` | Build `.docx` Word files |
| `react` / `vite` | UI framework & bundler |

RTL layout is handled by setting `bidi: true` on the DOCX section and `bidirectional: true` on paragraphs. Table cells are declared **left-to-right** in code but appear **right-to-left** visually in RTL Word documents (first declared cell = leftmost visual = rightmost reading position in Arabic).

### ⚠️ RTL alignment everywhere (the `cellAlign` flip)

**docx 9.x has no working section-level RTL** — it ignores `properties: { bidi }` (verified: nothing is emitted in `<w:sectPr>`), and tables are not marked `bidiVisual`. So the **entire document renders in an LTR frame**. Word then resolves every bidi (RTL) paragraph's physical `left`/`right` against that LTR frame and **flips them**: `AlignmentType.RIGHT` ends up hugging the visual **left**. This affects **body headings, body paragraphs, the footer, AND table cells** — everything, not just tables.

**Fix:** the `cellAlign(a)` helper in `buildDocx.js` swaps `RIGHT↔LEFT` when `_isRTL` (center/both untouched). **Every** RTL paragraph alignment is routed through it:
- `rp` (base paragraph helper) default alignment → `cellAlign(...)`
- `sectionHeading`, `axisHeading`, `bodyPara` → `cellAlign(...)`
- Footer vision/mission paragraphs → `cellAlign(AlignmentType.RIGHT)`
- `_textAlign` / `_numAlign` set via `cellAlign(toAlignType(...))` in `applyDesign`
- `tc` / `hCell` / `textCell` + the 11 direct header/axis-name/statement cells use `_textAlign`

Net result in the generated XML: Arabic text emits `<w:bidi/>` + `<w:jc w:val="left"/>` → renders on the physical **right**. Verified counts on a real report: `bidi+jc=right = 0`, `bidi+jc=left = 26`, plus `jc=center` for centered lines.

**Centered cover lines:** "عدد المشاركين بعد حذف التكرارات", "البرنامج / القسم", and "موجه إلى ..." are passed `align: AlignmentType.CENTER` (in `buildCoverPage` and `buildEvaluatorsSection`). `cellAlign(CENTER)` stays center.

**Do NOT** "simplify" any of these back to `AlignmentType.RIGHT` — that reintroduces the left-alignment bug across the whole document. The proper long-term alternative is to add `visuallyRightToLeft: true` to every `Table` **and** reverse all column arrays (since `bidiVisual` makes index 0 the rightmost column), plus find a real section-RTL mechanism for the body; only do that as a complete, tested change.
