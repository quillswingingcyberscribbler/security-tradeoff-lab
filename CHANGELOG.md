# Changelog

## Web-publication preparation — 2026-10-26

Repository-only update before the initial GitHub import. The unpublished package version and functional-review baseline remain unchanged.

- Replaced the command-line import guide with browser-only repository, upload, Actions, visibility, and release instructions.
- Expanded CI to run the complete v6.2 scenario, arithmetic, and structural checker.
- Added an Office publication-metadata record, a native Excel acceptance template, and explicit evidence checks in the stable-release gate.
- Added the author ORCID to the primary citation metadata and clarified post-import repository status documentation.

No source workbook or manuscript bytes were changed. This preparation does not establish native Microsoft Excel acceptance, resolve licensing decisions, or close functional blockers.

## 0.1.0-review.2 — 2026-10-23

Adopts the user-supplied Opportunity-Cost Security Model v6.2 and reviews its redesigned workflow.

- Replaced the current v6.1 workbook with byte-identical contents of `Opportunity-Cost-Security-Model-v6.2(1).xlsx`, stored under the portable name `Opportunity-Cost-Security-Model-v6.2.xlsx`.
- Reviewed the new 14-sheet structure, 50 keyed case slots, 20 demand rows and 2,947 formulas. Documented its new example baseline and earlier-source provenance; no numerical continuity with v6.1 is implied.
- Closed OC-01 through OC-05 in the scoped tests: readiness, spare-row keys, missing-versus-zero inputs, Example/Actuals eligibility, and coherent supply/overlap inclusion.
- Confirmed positive challenge/approval gates, dates, measurement provenance, usable shares, duplicate keys and reordered input records.
- Added 136 scenarios with 544 output assertions, 115 independent arithmetic assertions and 10 structural assertions. All pass. All 2,947 formula caches match a disposable LibreOffice recalculation and the independent interpreter baseline.
- Extended the read-only interpreter for Excel TRIM and replaced the v6.1 scenario harness with v6.2-specific checks.
- Added OC-06 for the secondary print-layout issue. Native Microsoft Excel acceptance and final licensing remain outstanding; the package stays a reviewed prerelease with no new calculation blocker found in scope.
- Archived the previous review, manifest, release state, scenario results and prior checker/evaluator. Added instructions for updating an existing repository without replacing `.git` or unrelated user content.

This package does not alter the supplied spreadsheet. It adopts a new author-provided source and updates the surrounding review and repository files.

## 0.1.0-review.1 — 2026-10-22

Initial repository preparation and independent review package for Opportunity-Cost Security Model v6.1.

- Preserved the original v6.1 source bytes.
- Added orientation, methodology, quick start, cell-specific issues and reproducible read-only checks.
- Added mascot, citation metadata, contribution/issue templates, source checksum manifest and package-integrity workflow.
- Recorded five calculation/workflow blockers and proposed corrections; those findings describe the earlier source and are retained in the historical review.

Package versions describe repository preparation and review. Workbook versions identify the supplied source. See [release status](validation/release-status.json) and the [current review](docs/review.md).
