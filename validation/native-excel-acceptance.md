# Native Microsoft Excel acceptance

Status: **PENDING**. No native Microsoft Excel acceptance has been recorded for Security Tradeoff Lab.

Package: `0.1.0-review.2`. Sources: Opportunity-Cost Security Model v6.2. This record was prepared on 2026-10-26; it is an acceptance template, not evidence that the checks have occurred.

## Environment and artifact identity

- Excel edition and full build:
- Operating system and version:
- Reviewer:
- Test date:
- Working-copy filenames:
- Working-copy SHA-256 hashes, if different from the manifest:
- Other applications used for secondary rendering, if any:

## Required application cycle

- [ ] Open a disposable working copy in the intended desktop Microsoft Excel build.
- [ ] Confirm automatic calculation and exercise the intended input controls.
- [ ] Force a full recalculation and inspect relevant outputs for errors or stale values.
- [ ] Save, close, and reopen the file; repeat the affected checks against the reopened file.
- [ ] Confirm formulas, validations, tables, conditional formatting, navigation, protection, charts, and print settings survived as intended.
- [ ] Inspect the intended report/print views rather than inferring them from the interactive display.

## Project-specific acceptance cases

- [ ] Open C001 as an example and C012 as a first live case; exercise the keyed edit links and complete the first-case workflow.
- [ ] Test missing inputs versus confirmed zero, incomplete profiles, invalid review states, future-effective changes, and duplicate keys.
- [ ] Verify both planning modes, measured-actual prerequisites, linked-demand overlap, and record reordering by whole-table sort.
- [ ] Recalculate, save, close, and reopen; confirm all case joins, readiness outputs, and baseline results remain correct.
- [ ] Inspect Start, Case View, and Guide print/PDF output after defining the intended report areas and page setup.

## Results and evidence

Record each result, unexpected behavior, screenshot/PDF evidence location, and follow-up issue here. Keep real organizational assessments and personal data out of the public repository.

## Completion

- [ ] All required cases pass. Any exception remains recorded as a release blocker and `stable_release_ready` remains false.
- [ ] The reviewer and test date are recorded above.
- [ ] `validation/release-status.json` has been updated only after reviewing this evidence.
- [ ] Source hashes, citation/version metadata, and the change log have been updated if an artifact intentionally changed.
