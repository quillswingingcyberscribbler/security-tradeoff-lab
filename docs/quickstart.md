# Quick start for v6.2

Download the [workbook](../workbooks/Opportunity-Cost-Security-Model-v6.2.xlsx) and save a working copy. This revision's scoped calculation review passed; native Excel acceptance is still pending. Read the [remaining limitations](known-issues.md).

## Explore one example

Open `Start`, then `Case View`. Select **C001** in `Case View!C5`. The summary shows the current and proposed annual profiles, their differences, evidence state, decision and measurement status. C001–C011 are examples and are excluded from live planning even if you experiment with their approval or benefit-stage fields.

The example comparison demonstrates why time and cash must stay separate: it can show a positive time-valued benefit while requiring additional cash. A euro valuation of employee time is not spendable money.

## Prepare your first live case

1. **Select C012.** Use `Case View` and its edit links. C012 is initially row 18 of each keyed table, but the ID is the durable link. C012–C050 are the 39 prepared unused slots.
2. **Define the decision in Cases.** Set type to Live. Enter one activity, its process/scope, owner, protection purpose, explicit proposal, retained work and assumptions. Keep the prepared case ID unchanged.
3. **Enter the current annual profile.** In `Current`, choose Staff, Management or Specialist, then enter execution hours, additional specialist hours, recurring non-labor cash and a source. All numeric fields are required: blank means unknown, while 0 means confirmed zero.
4. **Enter the complete proposed profile.** In `Proposed`, include retained work, new operating work, recurring cash, one-time implementation cost and a source. Scope reduction does not automatically scale every cost. Retain requires the same annual profile and releases no resource.
5. **Clear the challenge review.** All six checks must be Yes, with reviewer and date. Any No or unresolved concern blocks the case. Unchecked is not approval. Inspect protection, full alternative cost, binding requirements, rare failures, overlap, and skills/timing.
6. **Record a positive human decision.** In `Review & Decision`, complete local Evidence-backed confidence, the protection outcome, comparative evidence, approver, decision date, effective date and usable-time shares. Accepted reduction needs a risk-acceptance reference. Only Approved or Implemented can support a Validated forecast; Open, Pilot approved and Rejected do not qualify.
7. **Check readiness and timing.** The calculated Next action identifies what remains. Set the planning as-of date on `Settings`; future decisions, review dates and effective dates are excluded. Modeled potential may be visible before approval, but it does not fund the plan.

The optional `Settings!B17:B22` workload helper calculates annual hours from people, events per person, minutes and follow-up. It is intentionally separate from case inputs. Copy its result into the appropriate annual-hours field and document the assumptions; changing the helper later does not update the stored case.

## Add measured results

After implementation, choose Realized and set the decision to Implemented. In `Actuals`, record the measurement basis, observation dates, signed annual-equivalent business hours, specialist hours, recurring cash change, and labor-value change excluding cash. Add evidence, verifier, verification date and normalization/comparison basis.

**Enter annual equivalents, not raw totals for a short observation period.** A short period may use Annualized measurement with a documented normalization. Full-year actual requires at least 365 observation days. The workbook checks completeness and chronology; it does not perform the annualization or judge whether it is statistically justified.

Keep actuals as fixed measured values supported by the evidence. Forecast rate changes do not rewrite stored measurements. Positive hour savings are scaled by the usable-time shares; negative changes count in full. Confirmed zero actuals remain valid measured results.

## Build a resource plan

Choose **Approved + measured** or **Measured only** in `Settings!B8`. In `Reinvestment`, compare business hours, specialist hours and recurring cash separately. One-time funding is displayed separately and must have its own funding source.

Enter Live approved demands in the prepared demand table. Supply all four numeric requirements, including zeros, and the cost/overlap/resource-fit reference. If a demand is already included in a case estimate, record its case ID and Yes. The exclusion activates only while that case is eligible under the selected planning basis; otherwise the demand is counted independently. Duplicate overlap claims are counted and flagged for resolution.

An approved but incomplete demand suppresses the clean coverage statement. Even a balanced aggregate needs matching people, skills and availability; the workbook does not schedule or assign capacity.

## Preserve identities and history

Sort complete Excel tables, never isolated columns. Keep all 50 IDs exactly once in every case input table; duplicate or missing keys hold affected cases. The in-memory reordering tests establish lookup behavior, while native Excel sorting remains part of application acceptance.

v6.2 has a different structure and example baseline from v6.1. Re-enter each live case through the new workflow and reconnect its evidence. Do not paste entire legacy rows or formulas into the new tables. Keep the previous dated workbook and its evidence outside the public repository.

Use the XLSX as the primary interactive artifact. Before sharing a PDF, choose the required sheets and inspect each page; the reviewed Start export needs print-layout work.

## Optional checks

```bash
python3 tools/validate_repository.py
python3 -m pip install -r requirements-test.txt
python3 tools/check_opportunity.py --output validation/local-scenarios.json
```

The [validation guide](validation.md) distinguishes package integrity, independent calculations, in-memory scenarios and native Excel acceptance.
