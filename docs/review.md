# Security Tradeoff Lab — v6.2 review

Review date: **23 October 2026**. Package: **`0.1.0-review.2`**. Reviewed upload: `Opportunity-Cost-Security-Model-v6.2(1).xlsx`, packaged unchanged as `workbooks/Opportunity-Cost-Security-Model-v6.2.xlsx`.

## Outcome

**All five blockers from the v6.1 review are resolved in the tested v6.2 workflows.** Readiness is populated and actionable; input tables use stable keys; incomplete profiles cannot enter planning; Examples are isolated; measured records require period and evidence; and overlap exclusions follow the same eligibility as the credited resources. The earlier open question about manual review is also resolved: positive challenge clearance and a complete positive human decision are now required.

No new calculation blocker was found in this scoped review. The project remains a **reviewed prerelease** because native Microsoft Excel acceptance and final licensing are not recorded. A secondary print-layout issue remains: the interactive entry flow is much clearer, but the Start export needs deliberate page setup.

The source bytes were not modified. This review adopts the user's new workbook rather than applying formula patches.

## The new baseline

v6.2 is a redesign with different tables, input fields, gating and example values. `Sources!B24:F26` describes its earlier source material and distinguishes it from the 636-formula workbook reviewed previously. It should not be described as a numerically identical v6.1 repair or used as a destination for a bulk paste of old rows/formulas.

| Measure | Earlier v6.1 baseline | Reviewed v6.2 baseline |
| --- | ---: | ---: |
| Worksheets | 9 | 14 |
| Formula cells | 636 | 2,947 |
| Prepared case slots | 15 | 50 |
| Illustrative cases | 11 | 11 |
| Current annual example cost | €1,393,050 | €1,423,050 |
| Net business/staff hours | 11,753.33 | 14,313.33 |
| Net specialist hours | 938 | 1,370 |
| Recurring cash change | −€70,000 | −€97,000 |
| Annual time-valued benefit | €738,250 | €899,050 |

The business-hour label and full-profile method have also changed, so the table is a provenance comparison rather than a claim of like-for-like improvement. These figures are example assumptions. The positive time-valued benefit includes increased recurring cash spending and is not a budget saving.

## Verification performed

The source ZIP and XML are intact. Its 14 worksheets are visible, with eight Excel tables, 26 data-validation rules and 2,947 formula cells. No saved formula errors, VBA project, external workbook links, data connections or comment parts were found. No sheet is protected. The guide accurately says that the workbook is not access-controlled and that calculated cells should not be overwritten. The reviewed content consists of examples and unused slots; no populated organizational assessment or embedded overall license grant was found.

A disposable copy was opened and exported by **LibreOfficeDev 26.8.0.0.alpha0**, build `2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`. All **2,947 saved formula values matched** the fresh calculation, with no recalculated errors. The recalculated copy is evidence for the review and is not the packaged source.

The read-only source-formula interpreter was extended to support Excel's ASCII-space TRIM behavior. It then matched all **2,947 baseline results**, with zero mismatches or evaluation errors. This agreement is evidence for the formulas exercised here, not a general Excel-engine claim.

The final harness passes **136 scenarios with 544 output assertions**. It also passes **115 independent arithmetic assertions** across the eleven examples and their five aggregate measures, plus **10 structural assertions** for the prepared keys and readiness ranges. The checks use synthetic input overrides in memory and never save a workbook. They cover the original regressions and the surrounding approval, date, measurement, usable-capacity and demand workflows.

A LibreOffice PDF check of the first-use flow identified the print issue below. No native Microsoft Excel UI automation, Windows/Mac acceptance, browser Excel acceptance or real-world evidence verification was performed. Local commands were executed during the review. After import, the repository Actions tab records the hosted workflow runs.

## Disposition of the five prior blockers

### OC-01 — Readiness and realized warnings: resolved

The missing v6.1 readiness column has been replaced by `Checks!J7:J56`, displayed through `Review & Decision!M7:M56` and `Case View!A7`. The Start integrity banner at `A12` checks for missing readiness outputs and key problems. Measurement warnings use dedicated `Checks!AE:AF` states.

In the synthetic C012 fixture, Realized plus Implemented without actuals produces `Add measurements, period and evidence`, no planning contribution, and the Start warning that a Live Realized case lacks complete measurement evidence. Clearing either displayed or ledger readiness triggers the integrity warning. A complete measured record instead produces `Measured result eligible`. The 39 unused slots have explicit `Unused slot` states rather than an absent formula.

### OC-02 — Spare-row review-key collision: resolved

Each of the seven case input tables has the same 50 stable IDs in `A7:A56`; no portfolio status message occupies a case key. `Checks!AG7:AG56` requires one occurrence in every input table, and lookups join by ID.

Blocked challenge fixtures for C012, C013, C014, C015, C030 and C050 all reach the matching case and prevent eligibility. Duplicate IDs were tested independently in every input table; affected cases become invalid and contribute nothing. A missing review ID also holds the case.

Additional in-memory reordering moved complete input records, including IDs, within Current, Review & Decision and Challenge Review. Approval and quantities still followed C012, and a blocker entered in its moved challenge record still blocked C012. This establishes the keyed calculation behavior; native Excel's actual table-sort interaction remains an acceptance item.

### OC-03 — Missing current values treated as zero: resolved

`Current!G7:G56` requires a recognized execution group, three numeric nonnegative inputs, a source and valid rates. `Proposed!H7:H56` applies the equivalent rule to four numeric inputs, including one-time cost. `Checks!E:F` and `AH` consume these states.

Individually clearing execution group, hours, extra specialist effort, cash or the source blocks planning. Clearing one-time proposed cost also blocks it. Confirmed zero in each applicable numeric field remains valid. Negative quantities, numbers pasted as text, unknown groups and missing/invalid rates were tested and do not supply planning. The full proposed profile also avoids the earlier assumption that scope reduction could scale every fixed cost uniformly.

### OC-04 — Example leakage and incomplete actuals: resolved

`Checks!K:L` explicitly require Live type. An Example cannot enter planning even after its approval, evidence and measured fields are completed. The Modeled potential view is kept separate from selected supply.

`Actuals!M7:M56` requires a recognized basis, numeric start/end and signed result fields, evidence, verifier/date and a normalization description. Full-year actual requires at least 365 days; Annualized measurement permits a shorter period. `Checks!I` additionally checks the measured period against the effective date, and measured eligibility requires Realized plus Implemented.

Tests individually removed every required Actuals input, tested whitespace provenance, short/full-year periods, reversed or future chronology, verification before period end, and period end before effective date. Incomplete records remain excluded. Explicit zero actuals qualify; signed negative actuals count as added work/spend. Valid forecast-rate changes do not rewrite stored actual quantities or labor-value snapshots.

Annualization is now explicit in the method: **the assessor enters annual-equivalent results and documents normalization**. The workbook does not convert raw short-period totals automatically or determine whether the observation is representative.

### OC-05 — Premature overlap exclusion: resolved

`Reinvestment!L31:L50` uses the linked case's **selected-plan eligibility** in `Checks!M7:M56`, the same state used for resource supply. The mapping is a stable case ID rather than a fixed case row.

A modeled case leaves linked demand counted in full. A complete approved forecast both contributes supply and activates its documented overlap under Approved + measured. Switching to Measured only removes that forecast supply and restores the demand. A qualified measured case works in Measured only. Blocking or rejecting the case restores independent demand, and an Example cannot exclude Live demand.

Explicit No overlap remains additional demand. Unknown links count demand conservatively. Duplicate claims against the same eligible case are counted and flagged, and an incomplete approved demand prevents the clean coverage statement.

## Additional workflow findings

**Positive manual approval is enforced.** Challenge review requires six Yes answers, a reviewer and date, without unresolved concern. The forecast gate requires Evidence-backed confidence, an appropriate protection outcome, evidence reference, Approver, decision/effective dates, usable shares and Approved or Implemented. Open, Pilot approved, Rejected, Hypothesis, Plausible and incomplete records do not qualify. Accepted reduction requires its risk-acceptance reference.

**Timing and usable capacity are explicit.** Future dates do not contribute before the planning as-of date. Positive time savings use the specified usable fractions; extra work is never reduced by those fractions. Cash is kept separate. Specialist execution is allocated to specialist capacity. Retain with a changed annual profile is incomplete; an unchanged Retain yields no release.

**The helper is a documented estimate, not a live driver.** `Settings!B22` responds to the entered per-person frequency, duration and follow-up. It does not change stored Current or Proposed hours. This differs from the old workbook's directly linked headcount driver and is stated in the new guide.

**Evidence remains a human responsibility.** The model verifies presence, values, dates and state transitions. It does not authenticate reviewers, inspect evidence, confirm actual cost overlap, prove protective equivalence or assign released hours to a suitable person. Keep fixed measured inputs and external evidence/version records. No workbook rule establishes legal compliance.

## OC-06 — Print setup needs a separate pass

The interactive Start/Case View/Guide flow is clearer and the edit links support one-case work. The source has no explicit print area or page setup for those sheets. In the reviewed LibreOffice PDF export, Start spans four pages, with horizontal continuation separating example labels from their values and clipping long headings or explanation text. The first two pages were visually inspected.

This is a secondary publication issue, not a calculation blocker. Use the interactive XLSX as the primary tool. Define print areas, orientation and one-page-width scaling for intended report sheets, then inspect the actual pages in the supported Excel version. A selected-sheet export still needs visual acceptance; the audit PDF is not packaged as a finished report.

## Release status and provenance

The workbook is packaged under a portable name; the [manifest](../validation/manifest.json) records the exact uploaded name, byte size and SHA-256. The package does not modify the spreadsheet. The old workbook is no longer the current file under `workbooks/`; the [prior review](history/0.1.0-review.1.md) and earlier results remain archived for traceability.

The [release record](../validation/release-status.json) marks the five original issues resolved, the scoped model checks passed and no known calculation blockers in this review. It retains the outstanding native Excel and licensing requirements. The [validation guide](validation.md) provides commands and exact evidence files.

## Publication preparation addendum — 26 October 2026

A subsequent repository-only inspection found Microsoft Information Protection label records in every supplied Office source in this repository. The identifiers are omitted from the [publication-metadata record](../validation/publication-metadata.json). Their presence alone is not a confidentiality determination. No source Office artifact was modified during this follow-up.

The GitHub workflow now runs the complete v6.2 scenario, arithmetic, and structural checker in addition to package integrity. The [browser publishing guide](../PUBLISHING.md), [native Excel acceptance record](../validation/native-excel-acceptance.md), and release-evidence checks were added for publication preparation. These changes affect repository scaffolding and documentation; they do not change the functional findings or numerical/workbook baselines recorded above.
