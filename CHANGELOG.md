# Changelog

All notable releases of Redaa2 (محلل الاستبيانات الأكاديمية) are documented here.
This project follows [Semantic Versioning](https://semver.org/).

No prior tagged releases exist. This is the first documented baseline.

## [1.0.0] — 2026-06-27

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
