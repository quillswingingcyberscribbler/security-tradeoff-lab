# Model and definitions — v6.2

## Unit of analysis

One case compares one security activity in one process with one explicitly chosen proposal. Its scope, owner, protection purpose, retained work and assumptions must be recorded. v6.2 uses full current and proposed annual profiles; it does not select a proposal from a relevance score or automatically scale every cost when scope changes.

The model estimates resource consequences. It does not convert a qualitative control rating into an incident probability, loss estimate or proof of equivalent protection.

## Current and proposed profiles

Each profile specifies the execution group, annual execution hours, additional specialist hours, recurring non-labor cash and a source. Proposed also requires one-time implementation cost. Blank numeric inputs remain unknown; confirmed zero is valid. Invalid groups, negative quantities and missing rates hold the estimate.

Business hours comprise Staff and Management execution. Specialist execution joins additional specialist hours as specialist-team capacity. The same work must not be counted both as execution and as additional effort.

```text
Annual time cost = execution hours × selected group rate
                 + additional specialist hours × specialist rate
Annual total cost = annual time cost + recurring non-labor cash
Annual change = current profile − proposed profile
```

Positive change indicates lower work or spend; negative change indicates added work or spend. The time-valued benefit is the difference between the two annual total costs. It may be positive while recurring cash change is negative. One-time funding remains separate.

The rates on `Settings!B9:B11` are forecast assumptions. Replace and document them for the local analysis. The optional workload helper uses people × events per person × (minutes per event + follow-up share × extra minutes) / 60. Its result must be copied as a documented annual estimate; it is not linked to case rows.

## Proposal semantics

Simplify, Replace, Reduce scope and Retire all require an explicit full proposed profile. Include fixed licenses, retained work and transition cost. The label alone does not remove costs. Retain requires the same execution group, annual hours, additional specialist hours and recurring cash as Current; it is excluded from resource-release planning.

The optional Classification diagnostic highlights questions about dominant labels, label-dependent protections and defaults. Its ratios do not determine eligibility or prove that rare sensitive cases can be ignored.

## Four stages of evidence and use

| View or state | Meaning and inclusion |
| --- | --- |
| Example | Teaching material, never Live planning supply. |
| Modeled potential | Complete Live estimates, visible before positive review; not allocatable supply. |
| Approved forecast | Live, complete estimates with Clear challenge review, complete Evidence-backed approval, Validated stage and an eligible proposal. |
| Measured result | The same scope and approval foundation, with Realized stage, Implemented decision and complete measurement records. |

A forecast needs local evidence, a valid protection outcome, approval identity, decision and effective dates, and explicit usable shares. Accepted reduction needs a risk-acceptance reference. The challenge review requires six Yes answers, reviewer and date; No or an unresolved concern blocks it. Open, Pilot approved, Rejected and Unchecked never qualify as positive approval.

Dates are compared with the fixed planning as-of date in `Settings!B7`. Decision and effective dates must be positive numeric dates not later than the as-of date, and effective date cannot precede decision date. Challenge review must also be dated by the as-of date. Formula checks protect the tested gates even when ordinary dropdown guidance is bypassed.

These records establish input completeness. They do not authenticate an approver, inspect the evidence or make the workbook an immutable audit trail.

## Measured annual equivalents

`Actuals` stores signed annual-equivalent changes for business hours, specialist hours, recurring cash and labor value excluding cash. It also requires a recognized measurement basis, period start/end, evidence, verifier, verification date and normalization/comparison description.

Full-year actual requires at least 365 days. Annualized measurement allows a shorter observed period, but the annual-equivalent numbers and normalization must be supplied by the assessor. The workbook does not multiply raw period totals by 365/days or test whether the observed period is representative.

The observation period must be ordered and completed by the as-of date. Verification must be on or after period end and no later than the as-of date. The measured period must end on or after the case's effective date. The record may cover a comparison involving baseline observations; the normalization must explain the like-for-like basis.

Labor-value change is recorded separately using the period's fixed rates and mix. It is not derived by multiplying net hours by today's global rate, which could misvalue a shift between groups. Stored actual inputs remain unchanged when valid forecast rates change. Actual figures still require the complete case/estimate/approval foundation to be planning-eligible.

## Usable capacity and selected planning basis

Positive business and specialist savings are multiplied by their respective usable-time shares from 0 to 1. Negative hours count in full; the model does not discount extra work just because saved hours are hard to redeploy. Recurring cash is not multiplied by a time share.

Approved + measured selects eligible forecasts and measured cases. Measured only selects measured cases. A case cannot be counted as both: Validated and Realized are distinct stages. Forecast one-time implementation cost appears in future funding; measured cases do not re-add that implementation cost as future funding.

## Reinvestment and overlap

Live approved demands require a name, business hours, specialist hours, recurring cash, one-time cost, an explicit overlap answer and a cost/overlap/resource-fit record. Unknown figures make the approved demand incomplete and prevent a reassuring coverage statement.

An overlap answer Yes additionally identifies the linked case. It is excluded only when that case has selected-plan eligibility in `Checks!M7:M56`, the same inclusion state used for credited supply. A Modeled, Example, blocked, rejected or out-of-basis forecast cannot remove independent demand. Duplicate overlap claims against one eligible case are counted and flagged, rather than silently excluded twice.

The workbook checks the linkage and presence of a reference. A reviewer must verify that the demand is actually included in the case profile. Aggregated hours still require a practical skill, team and timing match. Cash funding is independent of time valuation.

## Identity, capacity and limits

Seven input tables each contain C001–C050 exactly once. The calculated ledger joins them by ID and checks uniqueness; it has 50 prepared case slots. Reinvestment has 20 prepared demand rows. Whole-record input-table reordering was checked in memory; arbitrary appended rows are unsupported. The workbook has no sheet protection, so preserve a baseline and edit only designated inputs.

The examples are assumptions, not observed savings or guaranteed protection outcomes. Sources are carried over and explicitly marked as not freshly verified in the workbook. The [review](review.md) checks calculations and the specified workflows, not every cited standard, legal obligation or empirical claim. Native Excel interaction and print acceptance remain separate tasks.
