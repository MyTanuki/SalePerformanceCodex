# Revenue Recognition

Revenue logic is finance-sensitive. Prefer changing and testing `scripts/revenue-logic.js` before touching rendering code.

## Source Deals

Revenue Analysis builds schedules from counted Won / Pre-WON deals. In code, Pre-WON is normalized to Won before revenue analysis runs.

Revenue schedule source:

- counted deals only
- status `won`
- scope from sale group, sale owner, New/Renew, and selected period rules

## Revenue Start Date

Preferred source:

1. Contract Start Date

Fallback behavior:

2. Expected Close Date may be used when the fallback option is enabled and Contract Start Date is missing.
3. Other fallback dates may include start date, stage change date, expected date, or created date depending on current app settings.

Expected Close Date is not conceptually the same as Contract Start Date. It is only a planning fallback when contract data is incomplete.

## One-time / OTC

One-time and OTC revenue is recognized once in the start month.

Example:

- Contract period: 2025-02-01 to 2026-01-31
- Amount: 12,000
- Billing Type: OTC
- Recognition: 12,000 in 2025-02 only

It should not be divided by 12 months.

## Recurring

Recurring revenue is distributed straight-line across service months.

Period source order:

1. Contract Start Date and Contract End Date.
2. Contract Period when end date is missing.
3. MRR inference when amount and MRR imply a month count.
4. One-month fallback with a data check flag if period cannot be determined.

The allocation routine works in cents to keep monthly splits summing back to the original contract amount.

## Recurring Target

Recurring Target is a planning baseline, not a sale target file value.

If the selected year is 2026:

- source year is 2025
- only recurring billing is used
- each source-year monthly revenue amount is shifted 12 months forward
- January 2025 recurring revenue becomes January 2026 target

Recurring Target should follow:

- selected year
- sale group
- sale owner
- New/Renew type

Recurring Target should not be changed by:

- keyword search
- Revenue Detail AND/NOT search

## Revenue Distribution

Revenue Distribution shows recognized revenue for the selected year/month scope. It includes old contracts when their schedule still has revenue entries inside the selected period.

This is important: selecting 2026 should include a contract that started in 2025 if it still recognizes revenue in 2026.

## Revenue Performance

Revenue Performance compares actual recurring revenue against Recurring Target.

- Actual recurring revenue: recurring entries in the selected period.
- Target: Recurring Target from prior-year recurring revenue.
- Gap: Actual - Target.
- Positive gap means above target.
- Negative gap means below target.

One-time revenue may appear in total revenue views, but it should not inflate recurring achievement.

## Invoice Analysis

Invoice Analysis is separate from revenue recognition. It distributes invoice amounts either by:

- Posting Date: accounting/invoice timing view.
- Service Period: service-month view.

Invoice Analysis groups multiple invoice rows for the same company/project into one monthly distribution row.

## Revenue Invoice Reconciliation

Reconciliation compares Revenue Distribution and Invoice Analysis.

Current reliable grain:

- Customer-level Compare Key

Not yet reliable:

- Project-level reconciliation

Reason: CRM deal names and invoice descriptions do not match consistently. Project-level comparison needs a future mapping layer.

