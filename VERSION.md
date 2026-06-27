# Version

**Current version: 1.0.0**

This matches the `version` field already declared in [package.json](package.json).

## Versioning policy

This project follows [Semantic Versioning 2.0.0](https://semver.org/) (`MAJOR.MINOR.PATCH`):

- **MAJOR** — incompatible changes to report output structure, schema format, or data shapes
  consumed by `analyze.js` / `buildDocx.js` that would break existing settings or generated
  reports' expected structure.
- **MINOR** — backward-compatible additions: new survey schemas, new wizard features, new
  export formats, new AI providers/tools.
- **PATCH** — backward-compatible bug fixes that don't change documented behavior.

## Release history

| Version | Date | Summary |
|---|---|---|
| 1.0.0 | 2026-06-27 | First stable, documented baseline of the existing application. No behavior changes — see [RELEASE_NOTES_v1.0.0.md](RELEASE_NOTES_v1.0.0.md) and [CHANGELOG.md](CHANGELOG.md). |

## Recommended git tag

```
git tag -a v1.0.0 -m "First stable release: full feature baseline, documentation only"
```
