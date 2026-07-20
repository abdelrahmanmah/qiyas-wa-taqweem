# Changelog

All notable releases of Redaa2 (محلل الاستبيانات الأكاديمية) are documented here.
This project follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added — Dynamic Survey Management system (new "🗂️ إدارة الاستبيانات" tab)
A second, fully independent survey-definition system, built alongside the 5 existing
hardcoded YAML surveys without modifying them. New files: `src/SurveyManagement.jsx`,
`src/engine/customSurveyModel.js`. See `CLAUDE.md → Survey Management System` for full
details. Highlights:
- Survey list view: create, edit, duplicate, delete (with confirmation), activate/deactivate.
- Step-by-step editor (عام → المحاور → معاينة) with a "التالي ←" button after each step and a
  distinct "💾 حفظ الاستبيان" button on the final step.
- Sections (axes) and questions: add/edit/delete/reorder, plus moving a question **across**
  sections via a per-question "نقل إلى محور آخر" control.
- Auto-import questions from a real Excel response file, auto-detecting which columns are
  questions vs. general info (name/email/degree/department/free-text) — including a guard
  against misclassifying long Likert questions that merely mention a metadata keyword
  mid-sentence.
- Excel **template** download/import (`⬇ تحميل القالب` / `📤 استيراد استبيان من قالب`) — a
  2-sheet workbook (معلومات الاستبيان + المحاور والأسئلة) that defines a survey's full
  name/scale/sections/questions in one import, including surveys with no axes at all (rows
  with a blank محور column fall back to a single auto-named section).
- JSON backup export/import (`⬇ تصدير نسخة احتياطية` / `⬆ استيراد نسخة احتياطية`) so custom
  surveys survive even if `localStorage` is cleared or the dev server is run on a different
  port (which otherwise looks like data loss, since `localStorage` is scoped per origin+port).
- If a question's Excel column name is left blank, it now defaults to the question's own text
  on save.

### Added — Custom surveys are now connected to the real analysis engine
- `customSurveyModel.js`: `toAnalysisSchema()` converts a custom survey into the exact schema
  shape `analyze.js` already expects (same shape `compileSchema.js` produces for the YAML
  surveys); `getAllAnalysisSchemas()` merges the 5 built-in schemas with active
  (`status: "active"`) custom surveys; `detectAnySurveyType()` mirrors `analyze.js`'s own
  filename/header detection across the merged set.
- `App.jsx`: every previous use of the static `SCHEMAS` import and `detectSurveyType` now
  resolves through these merged, live functions instead (`allSchemas()` / `detectSurveyType`
  via import alias) — a behavior-preserving superset; verified via regression test that the 5
  built-in surveys still detect/analyze identically.
- **Known, intentional limitations** of this integration (consequences of reusing `analyze()`
  unmodified): only `type: "likert"` questions are statistically analyzed (text/numeric
  question types are excluded, same as existing open-ended columns); for `likert-5` custom
  surveys, column-to-question mapping is positional (same as the built-in surveys), so
  reordering questions/sections changes which file column each one reads from — `likert-3`
  custom surveys are unaffected by reordering since matching is by fuzzy text, not position.

### Fixed
- `analyze.js`'s `parseResponse5` (5-point Likert response parser) only recognized
  `"(5) أوافق بشدة"`-style coded answers or a bare digit. Some Google Forms exports (and a
  real custom survey built from one) write the plain label only — `"أوافق بشدة"` with no
  leading code — which silently failed to parse, making every count/percentage/mean compute to
  `0`. Added a fallback that recognizes the 5 plain Arabic labels directly; only runs when the
  coded/digit formats don't match, so the 5 built-in surveys' parsing is unchanged (verified
  via regression test against the real student-survey sample file).

### Changed (documentation)
- `CLAUDE.md` updated throughout to describe the current `*.yaml`-based schema system
  (replacing stale documentation that still described the deleted `*.json` format), the new
  Survey Management system, the `SCHEMAS` → `allSchemas()` live-merge mechanism, and the
  `parseResponse5` fallback.

## [1.0.0] — 2026-06-27

No prior tagged releases exist. This is the first documented baseline.

### Added (baseline feature set — first stable snapshot)
- Local file upload and Google Drive import of Excel survey exports (`.xlsx`, `.xls`, `.csv`).
- Automatic survey-type detection from filename and column headers, with manual override.
- Five built-in survey schemas: `student` (likert-5, 25 axes), `faculty` (likert-3, 17 axes),
  `assistant` (likert-3, 17 axes), `graduates` (likert-3, flat 11 questions), `coordinator`
  (likert-5, 2 axes / 23 questions), defined in `src/schemas/*.yaml` and compiled at build
  time via `src/schemas/compileSchema.js`.
- Statistical analysis engine (`src/engine/analyze.js`): per-question and per-axis means,
  percentages, agreement rates, qualitative interpretation tiers, demographic breakdowns
  (`byDegree`, `byDepartment`), and degree×department cross-tabulation.
- Row-level data preview with manual row exclusion and department/degree filtering.
- Annual report mode and 2-year / 3-year comparison report mode.
- RTL Arabic Word (`.docx`) report generation (`src/engine/buildDocx.js`) with cover page,
  evaluators table, methodology, participants, detailed results, summary, and
  recommendations sections; custom schema-specific layout for the `coordinator` survey.
- On-screen enhanced report view with CSS-based distribution charts and one-click client-side
  PDF export (`html2pdf.js`).
- Google Drive OAuth2 integration (browser-only, Google Identity Services) with file
  list/search/filter and an aggregate Drive Dashboard view.
- Batch processing mode for analyzing multiple files at once.
- Floating AI assistant panel supporting Groq, Google Gemini, and Anthropic Claude, with
  tool-calling for report generation, file-upload prompts, and reading files from a
  configured local surveys folder (dev-only API routes).
- Persistent app settings (institution branding, document styling, signatures) and AI
  assistant settings via `localStorage`.
- 6-slide interactive tutorial overlay and animated step-by-step processing overlays.

### Notes
- This release is a **documentation-only snapshot** of the existing, already-working
  application. No application logic, UI behavior, or statistical computation was changed to
  produce this release.
