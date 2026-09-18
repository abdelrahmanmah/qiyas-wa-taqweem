# Release Notes — v1.0.0

**Release type:** First stable baseline (documentation-only release — no code or behavior
changes).

**Date:** 2026-06-27

## What this release is

This release captures the application **exactly as it currently works**, with no
refactoring, no logic changes, and no new features. Its purpose is to give the team a known,
named, reviewable baseline (`v1.0.0`) before any further feature work begins.

See [PROJECT_OVERVIEW.md](docs/PROJECT_OVERVIEW.md) for the full technical breakdown and
[CHANGELOG.md](CHANGELOG.md) for the itemized feature list.

## Headline capabilities at v1.0.0

- Import Excel survey exports (local upload or Google Drive) and auto-detect one of 5 survey
  types: رضا الطلاب، أعضاء هيئة التدريس، الهيئة المعاونة، آراء الخريجين،
  تقييم أداء منسقي البرامج.
- Clean and review data row-by-row before analysis.
- Generate full statistical breakdowns (per-question, per-axis, overall, demographic).
- Produce polished, fully RTL Arabic Word reports — single-year or multi-year comparison
  (2 or 3 years).
- Preview results on-screen with charts and export directly to PDF.
- Optional AI assistant (Groq / Gemini / Claude) that can read the loaded survey and trigger
  report downloads conversationally.

## Explicitly out of scope for this release

- No code refactoring.
- No logic or statistical-computation changes.
- No new survey types or report formats.
- No fixes to known limitations (see "Limitations" in `docs/PROJECT_OVERVIEW.md`) — they are
  documented, not resolved, here.

## Cleanup performed before this release

The following findings from the initial audit were fixed (file removals only — no app logic
changed; verified with `npx vite build` and a live dev-server render afterward):

- [x] Removed dead/unreferenced entry-point files: root `main.jsx`, `survey_analyzer.jsx`.
- [x] Removed legacy, unused schema files: `src/schemas/*.json` (superseded by `*.yaml`).
- [x] Removed exposed secrets from disk: `apis.txt` (plaintext Groq API key — **rotate this
      key**, it was unguarded) and `credentials.json` (Google OAuth client config).
- [x] Added `.gitignore` covering `node_modules/`, `dist/`, `.env.local`/`.env`,
      `credentials.json`, `apis.txt`, and real survey sample data.
- [x] Re-verified the app builds and renders correctly after all removals.

## Pre-tag checklist (for the user to action before `git tag v1.0.0`)

- [ ] Confirm this is acceptable as a git repository root, and run `git init` if not already
      a repo (environment shows it is not currently tracked by git).
- [ ] Rotate/revoke the Groq API key that was found in `apis.txt`.
- [ ] Confirm the feature summary below matches your understanding before tagging.

## Suggested semantic version

**v1.0.0** — appropriate because:
- The application is feature-complete for its current intended use case (no breaking gaps
  observed).
- `package.json` already declares `"version": "1.0.0"`, so this aligns the documented
  release with the existing package metadata.
- No prior tags exist, so this is correctly the first stable release rather than a
  pre-release (`0.x`) or patch.
