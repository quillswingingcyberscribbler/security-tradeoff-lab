# Maintainer issue updates

These are prepared issue/closure texts. Copy the items you want to track into GitHub Issues; the issue tracker becomes authoritative once an item has been created there.

## Close OC-01 through OC-05 after reviewing the v6.2 evidence

**Suggested closure comment:**

The supplied v6.2 redesign resolves this finding in the recorded regression scenarios. Review package `0.1.0-review.2`, dated 23 October 2026, preserves the new workbook's source bytes and documents its changed table structure and example baseline. All 2,947 source-formula results agree with fresh LibreOffice calculation; the independent interpreter matches that baseline. The 136 scenarios, 544 output assertions, 115 independent arithmetic assertions and 10 structural assertions pass. Native Microsoft Excel acceptance remains separately outstanding.

Use the finding-specific explanation below with that comment:

| Issue | Finding-specific closure evidence |
| --- | --- |
| OC-01 | `Checks!J7:J56` and `Review & Decision!M7:M56` contain readiness formulas; missing Actuals triggers explicit row and Start warnings and no planning supply. |
| OC-02 | All seven input tables retain unique C001–C050 keys. Blocked spare cases and independently reordered review records remain joined to the correct case. |
| OC-03 | Current/Proposed completeness gates require numeric inputs, groups, sources and rates; blanks/invalid inputs hold eligibility while confirmed zeros calculate. |
| OC-04 | Live filtering excludes completed Examples. Actuals require period, signed values, evidence, verifier/date and normalization; Realized requires Implemented. |
| OC-05 | Linked demand exclusion and credited resources both consume selected-plan eligibility, including Measured only, rejection and blocking transitions. |

The historical issue reproductions remain under `validation/history/0.1.0-review.1/`. They describe v6.1 and should not be rerun against the redesigned v6.2 cell layout.

## Open OC-06: Define and verify the operational print views

**Suggested labels:** `usability`, `documentation`, `workbook`.

### Problem

The first-use sheets have no explicit print area/page setup. LibreOffice's reviewed Start export spans four pages and clips text in horizontal continuations, separating example labels from values. This does not affect the tested calculations but limits shareable report quality.

### Proposed work

Define the intended report sheets and print areas. Set orientation, margins, scaling and page breaks deliberately. Keep long explanations readable and labels with their values. Use the XLSX as the primary interactive artifact until the report layout is accepted.

### Acceptance

- Start and Case View print without clipped headings, explanation text or separated comparison labels/values.
- The guide prints in a readable deliberate page sequence.
- Selected operational report sheets show complete values and readiness status.
- The supported Excel build/platform and inspected export are recorded.

## Complete native Excel acceptance

**Suggested labels:** `release`, `validation`.

Open a disposable v6.2 copy in the intended Excel build. Exercise the first Live case workflow, dropdowns, missing-versus-zero inputs, positive approvals, period/evidence boundaries, full-table sorting, duplicate IDs, planning modes and overlap transitions. Recalculate, save, close and reopen; check values, retained formulas, tables, links and validation. Check the print work above. Record the environment and results separately from the Python interpreter and LibreOffice records.

## Record confirmed license allocation

The author confirmed the MIT/CC BY 4.0 allocation on 26 October 2026. The workbook and Python scripts use MIT; explanatory documentation, figures, the project mascot, and illustrative examples use CC BY 4.0. The root `LICENSE`, `LICENSES/CC-BY-4.0.txt`, `LICENSE-STATUS.md`, `CITATION.cff`, `README.md`, and the release record have been updated accordingly. No separate licensing issue needs to be opened. Original source provenance and applicable external attributions must be retained.

## Review Office publication metadata

**Problem:** Every supplied Office source contains a Microsoft Information Protection label record. The earlier package review did not give this portfolio-wide metadata finding a separate publication task.

**Change:** Review the Sensitivity setting and applicable policy in the intended Office applications. Decide whether public release copies retain or remove the label. Preserve original sources separately and update hashes and validation evidence if any artifact changes. Do not copy label or site identifiers into public explanatory documents.

**Acceptance:** `validation/publication-metadata.json` records the reviewer, date, decision, and evidence. `validation/release-status.json` changes `publication_metadata_reviewed` only after that decision. Any changed release artifact has a reviewed hash and passing repository checks.
