# Redaa2 — محلل الاستبيانات الأكاديمية

## 1. Project Overview

Redaa2 is a Arabic-first, RTL, browser-based survey analysis tool built for ERU's Quality
Assurance unit. It is a single-page React application (no backend database, no server-side
persistence) that:

- Imports Google Forms / Excel survey exports (`.xlsx`, `.xls`, `.csv`).
- Auto-detects the survey type from the filename and/or column headers.
- Computes per-question and per-axis descriptive statistics (means, percentages, agreement
  rates) according to a 5-point or 3-point Likert configuration.
- Produces a fully formatted, RTL, Arabic Word (`.docx`) report, including cover page,
  methodology, participant breakdowns, detailed results tables, summary tables, and
  recommendations.
- Supports single-year ("annual") reports as well as 2-year and 3-year comparison reports.
- Optionally pulls files directly from Google Drive (OAuth2, browser-only, no backend).
- Offers an on-screen "enhanced report" preview with CSS-based charts and a client-side PDF
  export.
- Includes an optional AI chat assistant (Groq / Gemini / Claude) that can read survey data
  already loaded in the app and trigger report generation via tool calls.
- Includes a **Semester Survey Generator**: copies Google Forms templates per academic
  year/semester (optionally one copy per department), organizes them on Google Drive, and
  reports live response stats — with Google Drive itself as the only source of truth (see
  §4.12).

The app runs entirely in the browser at build/runtime, with the exception of three small
dev-only Vite middleware routes (`/api/chat`, `/api/list-files`, `/api/read-file`) used by the
AI assistant during local development.

---

## 2. Folder Structure

```
Redaa2/
├── index.html                  # Vite HTML entry
├── main.jsx                    # (legacy/duplicate root-level entry — see Limitations)
├── vite.config.js              # Vite config: React plugin, YAML loader plugin, dev API routes
├── package.json                # Dependencies & scripts
├── public/
│   └── logo.png
├── src/
│   ├── main.jsx                 # ReactDOM root — actual entry used by index.html
│   ├── App.jsx                  # Full UI: step wizard, settings, survey/Drive pickers (~3400 lines)
│   ├── AiChat.jsx                # Floating AI assistant panel + tool-calling logic
│   ├── EnhancedReportView.jsx     # Standalone enhanced report view with CSS charts + PDF export
│   ├── SurveyManagement.jsx      # No-code custom survey builder/editor tab
│   ├── SemesterSurveys.jsx       # Semester Survey Generator tab (Generate + Dashboard pages)
│   ├── engine/
│   │   ├── analyze.js            # Excel parsing, column mapping, statistics engine
│   │   ├── buildDocx.js          # Word (.docx) document generation
│   │   ├── customSurveyModel.js  # Custom-survey storage + analysis-engine bridge
│   │   └── semesterSurveyModel.js # Drive/Forms REST helpers + OAuth for the Semester Survey Generator
│   └── schemas/
│       ├── index.js              # Loads all *.yaml schemas at build time, compiles them
│       ├── compileSchema.js       # Converts minimal YAML schema format → full runtime schema
│       ├── student.yaml          # رضا الطلاب          — likert-5, 25 axes
│       ├── faculty.yaml          # هيئة التدريس         — likert-3, 17 axes
│       ├── assistant.yaml        # الهيئة المعاونة      — likert-3, 17 axes
│       ├── graduates.yaml        # آراء الخريجين        — likert-3, 11 questions (flat)
│       └── coordinator.yaml      # تقييم أداء منسقي البرامج — likert-5, 2 axes, 23 questions
├── dist/                        # Last production build output (gitignored)
├── .gitignore                   # Excludes node_modules/, dist/, .env.local, credentials.json,
                                  # apis.txt, and real survey sample data
└── (loose root-level files)     # analyze.py + requirements.txt (standalone Python CLI mirror),
                                  # generate_report.js (standalone docx utility), surveys_guide.md,
                                  # real sample survey .xlsx/.docx files — kept locally, gitignored
```

### Note on schema system (differs from older internal notes)

The schema system was migrated from hand-written JSON to a compact YAML format compiled at
build time:

- Each `src/schemas/*.yaml` file declares `id`, `name`, `nameEn`, `icon`, `scale` (3 or 5),
  `hints` (filename auto-detect keywords), `meta` (which columns are degree/department/email/
  freetext/etc.), and `axes` (or a flat `questions` list for single-axis surveys like
  `graduates`).
- `src/schemas/compileSchema.js` expands this into the full runtime schema shape (adds
  `colIndex` per question for likert-5, `questionStartIndex` for likert-3, default
  interpretation tiers, default scale value/score tables).
- `src/schemas/index.js` uses `import.meta.glob("./*.yaml", { eager: true })` (via a custom
  Vite YAML-loader plugin in `vite.config.js`) to load and compile every YAML file into the
  exported `SCHEMAS` object, keyed by `id`.
- The legacy `src/schemas/*.json` files were confirmed unused by any current code path
  (`analyze.js` imports `SCHEMAS` exclusively from `./schemas/index.js`) and have been
  removed as part of v1.0.0 cleanup.

---

## 3. Technology Stack

| Layer | Technology | Version (from package.json) |
|---|---|---|
| UI framework | React | ^19.2.5 |
| Build tool / dev server | Vite | ^8.0.10 |
| React Vite plugin | @vitejs/plugin-react | ^4.7.0 |
| Excel parsing & writing | xlsx (SheetJS) | ^0.18.5 |
| Word document generation | docx | ^9.6.1 |
| Client-side PDF export | html2pdf.js | ^0.14.0 |
| YAML parsing (build-time only) | js-yaml | ^4.1.0 |
| Language | Arabic (RTL primary UI/output), English schema metadata |
| Auth (Drive feature) | Google Identity Services (GIS), OAuth2 implicit flow, browser-only |
| Auth (Semester Survey Generator) | Google Identity Services (GIS), separate token client/scope (`drive` + `forms.body` + `forms.responses.readonly`), browser-only |
| Google Forms API | `forms.googleapis.com/v1` REST — read/copy/rename/move/publish forms, list responses | used only by the Semester Survey Generator |
| AI providers (optional assistant) | Groq, Google Gemini, Anthropic Claude — via OpenAI-compatible / native REST calls |

No server framework, ORM, or database is used. The only server-side code is three small Vite
dev-middleware routes that exist purely to support local development of the AI chat feature
(see Limitations). `xlsx` was previously read-only in this app (parsing uploaded Excel files);
the Semester Survey Generator's Excel-links export is the first place the app *writes* an
`.xlsx` file client-side.

---

## 4. Existing Features

### 4.1 Survey ingestion
- Local file upload (`.xlsx`, `.xls`, `.csv`) via drag-and-drop or file picker.
- Google Drive browsing and selection (OAuth2 popup, Drive REST API v3, Google Sheets
  auto-exported to `.xlsx`).
- Drag-over visual state, processing/step animation overlay during read + parse.
- Batch mode: process multiple files at once (`BatchProcessor` / `BatchItem`).
- Drive Dashboard: aggregate stats across all Drive survey files (by type, program, year),
  computed on-demand by downloading and parsing files in batches of 4.

### 4.2 Survey type detection & validation
- Filename-based detection (`detectTypeHintFromFilename`, schema `hints`/`fileHints`).
- Header-based fallback detection by fuzzy text similarity against known question banks
  (`detectSurveyType` in `analyze.js`).
- Manual override UI if detection is wrong or ambiguous.
- Per-slot detection and override in comparison mode (each year's file is detected
  independently).
- Academic year and program auto-detection from filename (`detectYearFromFilename`,
  `detectProgramFromFilename`).

### 4.3 Data preview & cleaning
- Row-level preview table with per-row removal checkboxes (manual exclusion of bad/duplicate
  responses).
- Department/degree filter dropdowns over the preview table.
- Live recomputation of summary statistics as rows are removed/filtered
  (`computeFilteredRows`).

### 4.4 Statistical analysis engine (`engine/analyze.js`)
- Reads raw Excel rows, strips metadata columns (timestamp, email, degree, department, free
  text), maps remaining columns to schema-defined questions.
- Two column-mapping strategies:
  - **likert-5**: explicit `colIndex` per question (after excluding metadata columns).
  - **likert-3**: fuzzy header-text matching against `questionStartIndex`.
- Computes, per question: response counts per scale value, percentages, mean (likert-5 only),
  agreement percentage.
- Aggregates per axis: axis mean, axis agreement %, qualitative direction/tier label from the
  schema's `interpretation` table.
- Computes overall survey-level mean, agreement %, and direction.
- Computes demographic breakdowns: `byDegree`, `byDepartment` counts, and a
  degree × department cross-tabulation matrix when both columns are present in the schema.
- Arabic text normalization (handles Arabic letter variants, diacritics, spacing) used both
  for detection and for fuzzy header matching.

### 4.5 Reporting modes
- **Annual**: single file, single year report.
- **Compare-2 / Compare-3**: 2 or 3 academic years analyzed together and merged into one
  comparison report, with validation that all years are unique and all files are the same
  survey type before processing.

### 4.6 Word (.docx) report generation (`engine/buildDocx.js`)
- Fully RTL, Arabic-formatted Word document via the `docx` library.
- Shared structure for every report: cover page, evaluators table (optional), variables
  section, statistical methodology section, participants section (with demographic
  cross-tab), detailed results section, summary section, recommendations section.
- Schema-specific detailed/summary table layouts for the `coordinator` schema (different
  column sets) vs. all other (standard likert-3/likert-5) schemas.
- Custom RTL-correction logic (`cellAlign` helper) that compensates for a `docx` 9.x
  limitation where the library does not emit a working section-level RTL frame (see the
  detailed note in `CLAUDE.md` — preserved as-is, not modified).
- Comparison-mode-specific document builder (`buildComparisonDocx`) producing a multi-year
  comparison report.
- User-configurable branding: institution name, fonts, colors, signature blocks (persisted in
  `localStorage`).

### 4.7 On-screen enhanced report view (`EnhancedReportView.jsx`)
- Standalone, scrollable on-screen report rendering (separate from the Word document),
  including CSS-based distribution bar charts per question/axis, tier-colored badges, and
  section headings mirroring the Word report structure.
- One-click client-side PDF export of this view via `html2pdf.js` (uses `html2canvas`
  internally) — produces a downloadable PDF without any server involvement.

### 4.8 AI Assistant panel (`AiChat.jsx`)
- Floating chat panel, togglable from the main UI.
- Supports three LLM providers, each with multiple model choices:
  - Groq (Llama 3.1/3.3, Gemma2)
  - Google Gemini (2.0 Flash, 1.5 Flash/Pro)
  - Anthropic Claude (Haiku, Sonnet, Opus)
- Provider-agnostic message format internally converted to/from Anthropic's native API shape
  when Claude is selected (`toAnthropicMessages`, `fromAnthropicToOpenAI` in
  `vite.config.js`).
- Tool-calling (function calling) support with four tools:
  - `generate_report` — triggers Word report generation/download for the currently analyzed
    survey.
  - `request_file_upload` — asks the user to upload a file of a given survey type.
  - `list_survey_files` — lists Excel files in a configured local "surveys folder" (dev-only,
    via `/api/list-files`).
  - `analyze_file_from_folder` — reads + analyzes a file directly from that folder
    (dev-only, via `/api/read-file`).
- System prompt is dynamically built from the currently loaded analysis result (schema label,
  respondent count, agreement %, per-axis breakdown) so the assistant can answer questions
  about the loaded survey.
- API keys/provider/model/surveys-folder settings persisted in `localStorage`
  (`eruQA_ai_v1`).

### 4.9 Settings & customization
- Persistent app settings (`eruQA_settings_v1` in `localStorage`): institution name, fonts,
  colors, signature blocks, evaluator-table toggle, etc.
- Settings panel UI (`SettingsPanel`) covering institution, document styling, and AI
  assistant configuration.

### 4.10 Onboarding / UX aids
- 6-slide interactive tutorial overlay (`TutorialOverlay`) explaining each wizard step, with
  slide navigation, dots, and per-slide accent colors/animations.
- Step-by-step processing overlay (`ProcessingOverlay`) with animated checklist during file
  read/parse/detect/compute phases (different step lists for local vs. Drive sources).
- Generic loading overlay (`LoadingOverlay`) shown during Word document generation.
- Step progress bar (`StepBar`) across the 6-step wizard.

### 4.11 Google Drive integration
- Pure browser OAuth2 (no backend) via Google Identity Services.
- Drive REST API v3 file listing, search, and client-side filtering by survey type, academic
  year (parsed from filename), and program (static list).
- List view and aggregate Dashboard view, toggle between the two.
- Google Sheets files auto-exported to `.xlsx` via the Drive `/export` endpoint before
  parsing.

### 4.12 Semester Survey Generator (`SemesterSurveys.jsx` + `engine/semesterSurveyModel.js`)

A third top-level tab (📆 استبيانات الفصل الدراسي), fully independent of the main analysis
wizard, that treats Google Drive as the single source of truth for a per-semester survey
lifecycle — no database, no localStorage cache; every load/filter-change/refresh is a fresh
Drive/Forms API call.

- **Generate Surveys page**: lists every Google Form inside a configured template folder
  (`VITE_GOOGLE_TEMPLATE_FOLDER_ID`), each with a select checkbox and a "نسخة لكل قسم"
  (per-department) toggle. On Generate, it copies the selected template(s) — one copy per
  entry in the app's department list if per-department is checked — sets the copy's Drive
  file name **and** the form's own internal title (`اسم الاستبيان - [القسم -] الفصل - السنة`),
  publishes the form so it can accept responses (required since forms created via the API are
  unpublished by default as of 2026-06-30), and files it under
  `{root}/{year}/{semester}/{templateName}/` (`VITE_GOOGLE_ROOT_SURVEYS_FOLDER_ID`), creating
  folders as needed. A live progress bar + rotating status text tracks the batch.
- **Dashboard page**: year/semester filters (populated from Drive subfolders, not hardcoded),
  stat cards (total surveys, total responses, avg/survey, last response), a CSS bar chart of
  responses per survey, and a table with per-survey/per-row refresh.
- **Analyze All**: runs every listed survey through the app's *existing, unmodified* analysis
  engine (`analyze.js`) by converting each form's Forms-API responses into the same
  `[header, ...rows]` shape a real Excel export would produce, auto-detecting the matching
  schema exactly like an uploaded file — then generates and downloads a Word report per survey
  via the same `buildAnnualDocx()` used everywhere else in the app.
- **Links export**: "نسخ الروابط كرسالة" (clipboard) and "تنزيل Excel" (a written `.xlsx`)
  bulk-export the form/responses links for a Generate batch or the Dashboard's current list.
- Uses its own OAuth token client/scope, separate from the read-only Drive integration in
  §4.11, since it needs write access (copy/rename/move/create-folder) plus Forms access.

---

## 5. Current Survey Types

| Schema id | Arabic label | English label | Scale | Axes | Notes |
|---|---|---|---|---|---|
| `student` | رضا الطلاب | Student Satisfaction | likert-5 | 25 | Has `levelCol` (المستوى) metadata column |
| `faculty` | أعضاء هيئة التدريس | Faculty Members | likert-3 | 17 | degree + department cross-tab supported |
| `assistant` | أعضاء الهيئة المعاونة | Teaching Assistants | likert-3 | 17 | degree + department cross-tab supported |
| `graduates` | آراء الخريجين عن البرنامج | Graduate Opinions on the Program | likert-3 | 1 (flat, 11 questions) | Has a fixed `recommendation` text baked into the schema |
| `coordinator` | تقييم أداء منسقي البرامج | Program Coordinator Performance | likert-5 | 2 (23 questions total) | Custom Word report layout (different columns), degree + department cross-tab |

All five are defined in `src/schemas/*.yaml` and compiled at build time by
`compileSchema.js`. New survey types are added by creating a new YAML file — no app-code
changes are required unless a non-standard report layout is needed (see
`CLAUDE.md → How to Add a New Survey Type`).

---

## 6. Current Analysis Flow

1. **Mode selection** (wizard step 0): annual / compare-2 / compare-3.
2. **File upload** (step 1): local file or Google Drive file, per-slot for comparison modes.
3. **Type validation** (step 2, annual/skipped in comparison mode): auto-detected schema is
   shown, user may override; column preview is shown.
4. **Data preview** (step 3, annual only): row table with per-row removal + dept/degree
   filters; live result preview recomputes on every change.
5. **Metadata entry** (step 4): academic year, program, prepared-by, reviewer.
6. **Results & download** (step 5): final statistics view + Word document download button.
   - Comparison mode skips steps 2–3 and jumps from step 1 directly to step 4 after
     `handleProcessMulti()` validates and merges all slots.

---

## 7. Data Processing Pipeline

```
Raw Excel file (.xlsx/.xls/.csv)
        │  XLSX.read() — src/engine/analyze.js: readExcel()
        ▼
2D array of rows (rawRows)
        │  detectSurveyType(filename, headers) — filename hints, then fuzzy header match
        ▼
Schema selection (SCHEMAS[id])
        │  prepareData() — identify metadata columns (timestamp/email/degree/dept/freetext)
        ▼
Filtered question columns + metadata columns
        │  computeFilteredRows() — apply user row-removal + dept/degree filters
        ▼
analyzeRows() / analyze()
        │  per-question counts/pcts/mean/agreePct
        │  per-axis aggregation + interpretation tier
        │  overall mean/agreement/direction
        │  byDegree / byDepartment counts
        │  crossDegreeByDept matrix (if both columns present)
        ▼
result object (see CLAUDE.md → Analysis Output for exact shape)
        │
        ├──► EnhancedReportView.jsx — on-screen charts + PDF export
        └──► buildAnnualDocx() / buildComparisonDocx() — Word (.docx) generation
```

For comparison modes, this pipeline runs once per slot/year, then `buildComparison()` merges
the per-year `result` objects into a single `comparison` object consumed by
`buildComparisonDocx()`.

---

## 8. Charts and Reports

- **In-app charts**: CSS-based horizontal distribution bars (`DistBar` in
  `EnhancedReportView.jsx`) showing the percentage breakdown of responses per question/axis,
  color-coded by interpretation tier. No charting library (e.g. Chart.js/Recharts) is used —
  charts are hand-built with styled `div`s.
- **Drive Dashboard charts**: similarly CSS-bar-based, showing distribution by survey type,
  program, and academic year across all analyzed Drive files.
- **Word report**: tables only (no embedded charts/images) — detailed per-question tables,
  axis summary rows, demographic cross-tab tables, all RTL-formatted per the `cellAlign`
  convention documented in `CLAUDE.md`.

---

## 9. Export Features

| Export | Format | Mechanism | Trigger |
|---|---|---|---|
| Annual report | `.docx` | `buildAnnualDocx()` → `docx` library → Blob → browser download | Step 5 "تحميل" button, or AI assistant `generate_report` tool call |
| Comparison report | `.docx` | `buildComparisonDocx()` → `docx` library → Blob → browser download | Step 4 results screen (comparison mode) |
| Enhanced report | `.pdf` | `EnhancedReportView` → `html2pdf.js` (html2canvas under the hood) → client-side PDF | Export button inside `EnhancedReportView` |
| Semester survey report | `.docx` | `buildAnnualDocx()` fed by Forms-API responses converted to rows | "🔍 تحليل الكل" on the Semester Surveys Dashboard, once per survey |
| Survey links list | `.xlsx` | `XLSX.utils.json_to_sheet` + `XLSX.write` → Blob → browser download | "⬇ تنزيل Excel" on Generate results / Dashboard |
| Survey links list | plain text (clipboard) | `navigator.clipboard.writeText` | "📋 نسخ الروابط كرسالة" on Generate results / Dashboard |

All exports happen entirely client-side; no file is ever uploaded to a server for report
generation. Filenames are auto-generated from schema label, program, and year(s)
(see `CLAUDE.md` for exact naming patterns).

---

## 10. Limitations

### Resolved as part of the v1.0.0 release prep

- ~~Duplicate/legacy entry points~~ — root-level `main.jsx` and `survey_analyzer.jsx`
  (an earlier, unreferenced prototype of the app, importable only by the now-deleted root
  `main.jsx`) were verified unreferenced by `index.html` (which only ever pointed at
  `/src/main.jsx`) and **removed**.
- ~~Legacy schema JSON files~~ — `src/schemas/*.json` were verified unreferenced (confirmed
  via grep — `analyze.js` only imports `SCHEMAS` from `./schemas/index.js`, which compiles
  the `*.yaml` files) and **removed**.
- ~~Exposed secrets at repo root~~ — `apis.txt` (a plaintext live Groq API key) and
  `credentials.json` (a Google OAuth client config) were **removed from disk**. The Groq key
  should be treated as compromised and rotated/revoked, since it was sitting unguarded in a
  plaintext file.
- ~~No `.gitignore`~~ — a `.gitignore` was added at the repo root excluding `node_modules/`,
  `dist/`, `.env.local`/`.env`, `credentials.json`, `apis.txt`, and the real survey sample
  data (`*.xlsx`, `*.xls`, `*.docx`, `استبيانات/`) so none of these are committed once the
  repo is git-initialized.
- Build and dev-server behavior were **re-verified after all removals**: `npx vite build`
  completes cleanly, and the dev server renders the full step wizard with no console/runtime
  errors (verified via a Playwright screenshot of the running app).

### Intentionally kept (verified as legitimate, separate from the app runtime)

- `analyze.py` / `requirements.txt` — a standalone Python CLI mirror of the analysis engine
  (reads a YAML schema + Excel file, outputs the same result JSON the React app produces).
  Not used by the app at runtime; kept as a useful scripting/batch tool.
- `generate_report.js` — a standalone Node script that builds a `.docx` report from a JSON
  result file. Not used by the app at runtime; kept as a separate utility.
- `surveys_guide.md` — reference documentation, unrelated to runtime behavior.
- Real sample survey `.xlsx`/`.docx` files and the `استبيانات/` folder — left on disk (useful
  for local testing) but now excluded from git via `.gitignore`, since they may contain real
  respondent data.

### Still open (by design, or out of scope for this release)

- **No automated tests**: there is no test runner, test files, or CI configuration in the
  repository. Correctness currently relies entirely on manual verification. Adding a test
  suite is a feature, not a documentation/cleanup fix, and was out of scope here.
- **AI assistant requires API keys**: the Groq/Gemini/Claude features are non-functional
  without a valid key (set via Settings or server-side env var); this is expected/by-design,
  not a defect.
- **Drive integration requires Google Cloud OAuth setup**: `VITE_GOOGLE_CLIENT_ID` must be
  configured in `.env.local`; without it the Drive tab will not authenticate. Expected/by
  design.
- **Semester Survey Generator requires additional setup**: `VITE_GOOGLE_TEMPLATE_FOLDER_ID` and
  `VITE_GOOGLE_ROOT_SURVEYS_FOLDER_ID` in `.env.local`; the Google Forms API enabled on the
  same Cloud project; and — while the OAuth consent screen is in "Testing" mode — every user's
  account added as a test user, plus the `forms.body`/`forms.responses.readonly` scopes added
  to the consent screen, or Google rejects the broader scope request with `Error 403:
  access_denied`. Expected/by design, but easy to hit on first setup.
- **No native "linked response Sheet" for generated surveys**: the Forms REST API has no way
  to read/attach the classic Sheet-linked-to-a-form feature (Apps-Script-only), and copying a
  form via the Drive API drops any such link anyway. Response counts and "Open Responses" are
  read straight from the Forms API's `responses.list` / a deep link to the form's own
  Responses tab instead — by design, not a bug to fix.
- **Semester Survey Generator's "Analyze All" depends on auto-detection**: a generated
  survey's responses are matched against one of the app's 5 built-in or custom schemas via the
  same fuzzy detection used for uploaded files; a template whose question wording doesn't
  reasonably match an existing schema will be skipped rather than force-analyzed.
- **`docx` 9.x RTL limitation**: documented in `CLAUDE.md` — the library does not provide a
  working section-level RTL frame, requiring the manual `cellAlign` left/right-swap
  workaround throughout `buildDocx.js`. This is a known, already-mitigated constraint, not a
  bug to fix in this release.
This document originally described the application as observed in the code at the v1.0.0
cleanup pass (dead files, legacy duplicates, and exposed secrets removed; no application logic
changed). It has since been updated to describe the Semester Survey Generator (§4.12), added on
the `feature/semester-survey-generator` branch — that work did add new application logic/UI,
unlike the v1.0.0 pass above.
