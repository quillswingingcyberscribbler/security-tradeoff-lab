# Contributing

Start with the [known issues](docs/known-issues.md) and the [review](docs/review.md). Use synthetic examples in public reports.

## Report a calculation problem

Identify the workbook version, sheet and cell. State the input values, expected result, observed result and calculation application/version. Describe a small reproduction. Distinguish a saved display value, a fresh recalculation and an independently calculated expectation.

## Propose a change

Explain the decision or user problem first. Keep formula corrections, changes to methodology and visual changes distinct in the description. A changed risk threshold, eligibility rule or scope definition needs a reason and a visible change note.

Preserve stable identifiers, formulas, dropdowns, extended validations, tables, conditional formatting, charts, protection settings and hidden working ranges unless the change deliberately affects them. For a binary workbook change, include a text change list with affected sheet/cell ranges and before/after evidence. Do not replace formulas with saved values.

For a functional correction, test the affected ordinary case, a boundary case and incomplete input. Check that a valid zero remains different from missing information. Recalculate in the intended Excel version, save, close and reopen. Record that native check separately from script results.

## Repository checks

Run `python3 tools/validate_repository.py` from the repository root. This checks package structure, links, hashes and saved errors; it does not recalculate Excel. Run the project-specific checks documented in the quick start or review.

Update the change log and validation evidence when behavior changes. If a source workbook is intentionally revised, update the corresponding hash only after reviewing the change. Retain accurate provenance and use a new source version.

## Conduct

Be specific, constructive and respectful. Discuss the reasoning and evidence. Credit others' work and make uncertainty visible. Do not include confidential assessments, credentials, personal data or employer/customer material without permission. Licensing status is stated in the repository's license documentation; proposals do not silently change it.
