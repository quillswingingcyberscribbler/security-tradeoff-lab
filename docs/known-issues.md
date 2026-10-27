# Issue status — v6.2

Review date: 23 October 2026. Package: `0.1.0-review.2`. Source: `Opportunity-Cost-Security-Model-v6.2.xlsx`, preserved unchanged from the uploaded file.

## Previous blockers resolved

| ID | Previous problem | v6.2 disposition |
| --- | --- | --- |
| OC-01 | Missing readiness formulas and misleading realized-data warning | Resolved in the tested readiness, missing-output and measurement-warning scenarios. |
| OC-02 | Spare-row key overwritten by a status message | Resolved by separate keyed tables and identity checks; spare slots, duplicates and reordered records tested. |
| OC-03 | Unknown current inputs treated as zero downstream | Resolved by complete profile gates; missing, invalid and explicit-zero inputs tested. |
| OC-04 | Examples and unsupported realized figures entered planning | Resolved by explicit Live filtering, complete actual records, date rules and documented annual-equivalent inputs. |
| OC-05 | Demand excluded before its case supplied planning resources | Resolved by one selected-plan eligibility definition for supply and linked overlap. |

The [full review](review.md) identifies cells and evidence. The 136 scenarios, 544 output assertions, 115 independent arithmetic assertions and 10 structural assertions all pass. This is scoped verification, not exhaustive validation or native Excel acceptance.

## OC-06 — First-use sheets need print setup

**Priority:** secondary publication/layout improvement. **Affected:** Start, Case View and Guide page setup; Start's PDF output was inspected.

The source lacks explicit print areas and page setup on these sheets. The reviewed LibreOffice output splits Start across four pages, with horizontal continuation clipping text and separating labels from values. Interactive calculations are unaffected.

**Current use:** prefer the interactive XLSX. Export selected report sheets and inspect every page before sharing.

**Correction:** define report-specific print areas, orientation, margins and one-page-width scaling; confirm long text, values and labels remain together. Record the supported Excel version and inspect both PDF and physical-page layout. Do not mistake a previous application's export for native Excel acceptance.

## Remaining release requirements

Formal native Microsoft Excel acceptance has not been documented, including the complete edit/recalculate/save/close/reopen and interaction checklist. The workbook has been opened and used in Microsoft Excel, and the metadata cleanup save/reopen cycle is complete. The MIT/CC BY 4.0 license allocation and the publication-metadata review are complete. The public copy retains the Public sensitivity label and has been cleaned of document properties and its absolute local path. Formal native Excel acceptance and the secondary print-layout issue keep the package a reviewed prerelease; they are not calculation failures.

## Interpretation boundaries

- v6.2 is a new table structure and example baseline. Do not bulk-paste v6.1 rows or formulas into it.
- Actuals are signed annual equivalents entered by the assessor. Annualization and its justification are documented inputs, not an automatic calculation or statistical validation.
- Approval/evidence fields record human decisions; they do not authenticate reviewers or verify the referenced evidence.
- All sheets are unprotected. Preserve a baseline and edit designated inputs; Start detects missing readiness outputs and key faults, but is not a general tamper detector.
- Capacity is 50 prepared case IDs and 20 demand rows. Appended rows are unsupported.

See [draft issue text](issues-to-open.md) for the remaining work and comments suitable for closing the old issues.

## PUB-01 — Office sensitivity-label metadata

## PUB-01 — Office sensitivity-label metadata — resolved

The publication-metadata review was completed on 27 October 2026. The visible Public classification was deliberately retained. Document properties and the absolute local workbook path were removed in Microsoft Excel, and the `lastModifiedBy` property is empty. The public copy was saved, closed and reopened.

All 14 worksheets, 2,947 formulas and saved cell values match the original source. The complete model checker passes. The current public-copy hash is recorded in [validation/manifest.json](../validation/manifest.json), and the review decision is documented in [validation/publication-metadata.json](../validation/publication-metadata.json). Full native Excel acceptance remains a separate release requirement.
