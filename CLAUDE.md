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
├── engine/
│   ├── analyze.js            # Excel parsing, column mapping, statistics
│   └── buildDocx.js          # Word document generation (docx library)
└── schemas/
    ├── student.json          # رضا الطلاب         — likert-5, 25 axes
    ├── faculty.json          # هيئة التدريس        — likert-3, 17 axes
    ├── assistant.json        # الهيئة المعاونة     — likert-3, 17 axes
    ├── graduates.json        # آراء الخريجين       — likert-3, 11 questions
    └── coordinator.json      # تقييم أداء منسقي البرامج — likert-5, 2 axes, 23 questions
```

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

## Schema System (`src/schemas/*.json`)

Every survey is defined by a JSON schema. Key fields:

```jsonc
{
  "id": "unique_key",           // must match key in SCHEMAS object in analyze.js
  "label": "Arabic display name",
  "labelEn": "English name",
  "fileHints": ["keyword1"],    // substrings matched against filename for auto-detection
  "scale": {
    "type": "likert-5" | "likert-3",
    "values": [{ "code": "1", "label": "...", "score": 1 }, ...],
    "agreementCodes": ["4", "5"],   // codes counted as "agreement"
    "tokenPattern": "\\((\\d)\\)"  // regex for likert-5 cell values like "(4)"
  },
  "metadata": {
    "timestampCol":  ["Timestamp"],      // columns to exclude from question mapping
    "levelCol":      ["المستوى"],        // excluded for likert-5 colIndex mapping
    "degreeCol":     ["الوظيفة"],        // → byDegree in result + excluded from qs
    "departmentCol": ["التخصص"],         // → byDepartment in result + excluded from qs
    "freeTextCols":  ["مقترحات"]         // excluded from question mapping
  },
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
        // (after removing all metadata columns)
        // Only used for likert-5 schemas.
        // likert-3 schemas use fuzzy header matching instead (set questionStartIndex).
      ]
    }
  ]
}
```

### likert-5 vs likert-3 column detection

| Scale | Detection method | Schema field |
|---|---|---|
| `likert-5` | Filter out metadata cols → use `colIndex` | `colIndex` per question |
| `likert-3` | Fuzzy text similarity between header and `q.text` | `questionStartIndex` (skip first N cols) |

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
| أولاً: متغيرات | `buildVariablesSection` | same |
| ثانياً: معالجة إحصائية | `buildMethodologySection` | same |
| ثالثاً: المشاركون | `buildParticipantsSection` | same (cross-tab auto from degreeCol×departmentCol) |
| رابعاً: النتائج | `buildDetailedSection` | `buildCoordinatorDetailedSection` |
| خامساً: الملخص | `buildSummarySection` | `buildCoordinatorSummarySection` |
| التوصيات | `buildRecommendationsSection` | skipped |

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

## How to Add a New Survey Type

### Step 1 — Create the schema file

```bash
# Create src/schemas/<id>.json
```

Minimum required fields: `id`, `label`, `fileHints`, `scale`, `metadata`, `interpretation`, `axes`.

**For likert-5:** add `colIndex` to every question (0-based after excluding metadata columns).  
**For likert-3:** omit `colIndex`, set `questionStartIndex` (number of non-question columns at the start).

If the survey has demographic cross-tabulation (job × specialization), set:
```json
"degreeCol":     ["column header keyword for rows"],
"departmentCol": ["column header keyword for cols"]
```
The cross-tab appears automatically in ثالثاً when both columns are detected.

### Step 2 — Register the schema in `analyze.js`

```javascript
// src/engine/analyze.js
import mySchema from "../schemas/mysurvey.json";

export const SCHEMAS = {
  student: studentSchema,
  // ... existing ...
  mysurvey: mySchema,   // ← add here, key must match schema "id"
};
```

### Step 3 — Add to the survey picker in `App.jsx`

```javascript
// src/App.jsx — inside SurveyTypePicker component
const TYPES = [
  // ... existing entries ...
  { id: "mysurvey", icon: "📋", label: "اسم الاستبيان", desc: "X محاور – مقياس Y درجات" },
];
```

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
`requestAccessToken({ prompt: "select_account" })` is called every time so the user always sees the account picker — Google never auto-selects a cached account.

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
- **نوع الاستبيان** — populated from `Object.values(SCHEMAS)` labels
- **السنة الدراسية** — dynamically built from filenames via `detectYearFromFilename()`
- **البرنامج** — static list in `PROGRAMS` constant: محاسبة، اقتصاد، إدارة، علوم سياسية، تكنولوجيا أعمال

Helper functions (defined after `detectProgramFromFilename`, before `uniqueCol`):
- `detectYearFromFilename(filename)` — regex `/(\d{4}[-_]\d{4}|\d{4}[-_]\d{2})/`, normalises `2024_25` → `2024-2025`
- `detectTypeHintFromFilename(filename)` — matches against `schema.fileHints` with Arabic normalisation

### Drive Dashboard (`DriveDashboard` component)
Toggle between **📁 قائمة** (list with filters) and **📊 لوحة** (dashboard) via `driveDashboard` boolean state.

`DriveDashboard` props: `{ files, token, onSelectFile }`. Internal state: `stats[]`, `loading`, `done`.

On-demand analysis: clicking "📊 تحليل الملفات (N)" downloads files in batches of 4 via `Promise.all`, parses with `readExcel()`, counts `rows.length - 1`. UI shows:
1. Summary cards by survey type (file count + response count)
2. CSS bar charts — distribution by program and by academic year
3. Detailed table with per-file type/year/program/responses and "تحليل" button

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
| `driveDashboard` | `true` = dashboard, `false` = file list |
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
