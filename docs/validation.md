# Validation scope — v6.2

Review date: **23 October 2026**. Package: **`0.1.0-review.2`**.

## Completed checks

| Layer | Recorded result | Boundary |
| --- | --- | --- |
| Source/package structure | 14 visible sheets, 8 tables, 26 validation rules, 2,947 formula cells; no saved errors or external workbook dependencies found | Does not recalculate formulas or prove business logic. |
| LibreOffice recalculation | All 2,947 source caches match the disposable recalculated copy; zero formula errors | LibreOfficeDev 26.8.0.0.alpha0, not Microsoft Excel. |
| Interpreter baseline | All 2,947 source-formula results match the fresh calculation | Purpose-limited supported-formula interpreter, not a general Excel implementation. |
| Scenarios | 136 scenarios / 544 output assertions pass | In-memory synthetic overrides, not native input editing or exhaustive testing. |
| Independent arithmetic | 115 assertions pass | Separate calculations for eleven example profiles and five aggregate measures. |
| Structure assertions | 10 pass | Prepared IDs in eight tables/ledger ranges and two 50-row readiness ranges. |
| Visual output | Start print issue documented | First two exported pages inspected; native Excel printing remains pending. |

The eight identity ranges in the structural assertions are the seven input tables plus the calculated Checks ledger. The eight actual Excel tables are a separate package-structure count.

The five earlier blockers are resolved in the scoped tests. No new calculation blocker was found. The remaining print issue and unperformed native Excel acceptance are kept visible in [known issues](known-issues.md) and [release status](../validation/release-status.json).

## Evidence files

- [manifest.json](../validation/manifest.json): original uploaded filename, portable repository path, size and SHA-256.
- [structure-results.json](../validation/structure-results.json): current worksheet, table and validation inventory.
- [recalculation-results.json](../validation/recalculation-results.json): source-cache comparison and LibreOffice version/build.
- [scenario-results.json](../validation/scenario-results.json): interpreter baseline, independent arithmetic, structural checks and every scenario assertion.
- [Historical review](history/0.1.0-review.1.md) and `validation/history/0.1.0-review.1/`: earlier findings and results for v6.1.

The packaged workbook has the exact uploaded bytes. Disposable recalculated files and audit PDFs are not shipped as replacements for it.

## Reproduce locally

From the repository root, Python 3.10 or later:

```bash
# Uses only the standard library; checks ZIP/XML, relationships, hashes and local links.
python3 tools/validate_repository.py

# Optional formula checks use the pinned read-only openpyxl dependency.
python3 -m pip install -r requirements-test.txt
python3 tools/check_opportunity.py --output validation/local-scenarios.json
```

The checker defaults to the packaged v6.2 file relative to its own location, so it also works from another working directory. To name a different source or compare against a separately recalculated copy:

```bash
python3 tools/check_opportunity.py /path/to/source.xlsx   --recalculated /path/to/recalculated-copy.xlsx   --output /path/to/results.json
```

Without `--recalculated`, the baseline compares against source caches; the independent review established that these match fresh LibreOffice calculation. The command does not launch LibreOffice or Microsoft Excel. It exits nonzero for mismatches, evaluation errors or failed assertions. It never saves a workbook.

The interpreter adds TRIM support for v6.2 and uses read-only workbook loading. Reordering scenarios move complete input records and their IDs in memory, preserving the calculation rows. They validate joins, not the behavior of Excel's Sort UI. Actuals fixtures contain fixed synthetic annual-equivalent values, not claims about real measurements.

## CI and stable release

The default GitHub workflow is named Repository checks. It runs package integrity and the complete v6.2 scenario, independent-arithmetic, and structural checker. A green run confirms the recorded scoped checks; it does not certify every decision, publish a release, or run native Excel. Local commands were executed during preparation. After import, the repository's Actions tab is the authoritative record of hosted runs.

`python3 tools/release_gate.py` remains an explicit stable-release check. It intentionally fails while formal native Excel acceptance is unconfirmed, even though the scoped model checks pass. The MIT/CC BY 4.0 license allocation and the publication-metadata review are complete. The public workbook was saved, closed and reopened after metadata cleanup; its Public sensitivity label was retained, while document properties and the absolute local path were removed. All worksheets, formulas and saved cell values match the original source. The complete edit/recalculate/save/close/reopen and interaction acceptance has not been documented.

For native acceptance, use the [acceptance record](../validation/native-excel-acceptance.md), record Excel build and platform, use a working copy, exercise the documented first-case and boundary paths, then recalculate/save/close/reopen. Inspect output values, tables, dropdowns, navigation links, sorting and print views. Complete the [publication-metadata review](../validation/publication-metadata.json). Update source hashes only if a newly reviewed source is intentionally adopted. Preserve previous dated assessments and evidence externally.
