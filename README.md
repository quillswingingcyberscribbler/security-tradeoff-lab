<p align="left"><img src="assets/logo.png" width="160" alt="Squirrel weighing two acorns, the Security Tradeoff Lab mascot"></p>

# Security Tradeoff Lab

**Compare security-control changes through employee time, specialist capacity, cash and evidence.**

Security work consumes scarce resources. This Excel model compares a current activity with its full proposed operating profile, then distinguishes estimated potential, approved forecasts and measured results. Business hours, specialist hours, recurring cash and one-time funding remain separate.

**Status: reviewed prerelease `0.1.0-review.2`, reviewed 23 October 2026.** The supplied **v6.2 resolves all five calculation/workflow blockers identified in the previous review**, within the tested scenarios. No new calculation blocker was found in that scope. Native Microsoft Excel acceptance and final licensing remain outstanding; a secondary print-layout issue is documented. The uploaded workbook bytes are preserved unchanged.

## Start here

- [Download Opportunity-Cost-Security-Model-v6.2.xlsx](workbooks/Opportunity-Cost-Security-Model-v6.2.xlsx)
- [Quick start](docs/quickstart.md)
- [Model and definitions](docs/methodology.md)
- [Review, resolved findings and evidence](docs/review.md)
- [Known issues and remaining acceptance work](docs/known-issues.md)

Open a working copy in desktop Excel. Read `Start`, select C001 on `Case View` to explore an example, or C012 to begin a live case. Follow the edit links through scope, Current, Proposed, challenge review, approval and Actuals. Python is needed only for optional reproducibility checks.

The redesigned workbook has 14 sheets, 50 prepared case IDs and 20 demand rows. Eleven cases are illustrative examples; 39 are unused slots. Every input table retains its case IDs, so whole-record reordering can be joined by key. Sort the entire table, keep the IDs, and do not append unsupported rows.

## What changed in this revision

Readiness messages and measurement warnings are present. Missing current or proposed inputs block planning while confirmed zero remains valid. Challenge review and human approval must be positively complete. Example cases never supply the live plan. Realized results require measured values, an observation period, evidence, verification and normalization. Linked demand is excluded only when its case is eligible under the same selected planning basis.

v6.2 is a substantial redesign with a new example baseline, rather than a cell-for-cell update to v6.1. Its example portfolio shows **€899,050 of annual time-valued benefit and a €97,000 increase in recurring cash spend**. These are illustrative assumptions, not observed organizational savings. See the [version comparison](docs/review.md#the-new-baseline) before migrating an assessment; do not bulk-paste old rows or formulas.

## Verification included

| Verification | Result |
| --- | --- |
| Fresh disposable LibreOffice recalculation | All 2,947 saved formula results agree; no formula errors. |
| Read-only formula interpreter | All 2,947 baseline results agree with the fresh calculation. |
| Regression and boundary scenarios | 136 scenarios, 544 output assertions, all matched. |
| Independent example arithmetic | 115 assertions passed. |
| Prepared keys and readiness structure | 10 assertions passed. |
| Native Excel interaction and print acceptance | Still pending. |

The [validation guide](docs/validation.md) explains the limits. These checks establish the recorded behavior, not every possible input or the quality of real evidence.

```bash
# Standard-library package integrity checks
python3 tools/validate_repository.py

# Optional source-formula regression checks; never saves the workbook
python3 -m pip install -r requirements-test.txt
python3 tools/check_opportunity.py --output validation/local-scenarios.json
```

[Recorded results](validation/scenario-results.json) include each expected and observed value. The checker accepts an optional workbook path and `--recalculated` comparison file. It defaults to the packaged v6.2 workbook from any working directory.

## Repository contents

`workbooks/` contains the current source. `docs/` contains first-use guidance, methodology, review and issue status. `tools/` holds read-only checks; `validation/` holds provenance and results. The prior review and test records are archived under `docs/history/` and `validation/history/`; the earlier workbook remains in the prior release package. The squirrel mascot is under `assets/`.

Use [PUBLISHING.md](PUBLISHING.md) to import a new repository or update an existing one. The related IMS Compass, Agent Assurance Atlas and GL-Agent projects answer different questions; their scores are not combined here.

## Licensing and contribution

No overall reuse license was found in the supplied workbook. See [LICENSE-STATUS.md](LICENSE-STATUS.md) for the remaining author decision. [CONTRIBUTING.md](CONTRIBUTING.md) explains how to report a reproducible issue, and [CITATION.cff](CITATION.cff) identifies this package version. Keep completed organizational assessments outside the public repository.

## Prepare publication

The [browser-only publishing guide](PUBLISHING.md) covers repository creation, file upload, Actions, visibility, and prerelease creation through GitHub's web interface. Complete the [publication-metadata review](validation/publication-metadata.json) and, for a stable release, the [native Excel acceptance record](validation/native-excel-acceptance.md). These records do not replace the functional issues described above.
