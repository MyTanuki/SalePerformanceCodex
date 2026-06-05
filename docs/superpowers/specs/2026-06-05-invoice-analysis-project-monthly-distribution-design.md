# Invoice Analysis Project Monthly Distribution Design

Date: 2026-06-05

## Goal

Improve `Revenue Analysis > Invoice Analysis` so invoice rows for the same company/project are combined into a single detail row and distributed across monthly columns, matching the monthly distribution style used by Revenue Analysis.

## Business Rule

For Invoice Analysis, "same company/project" means:

- Same normalized company name.
- Same normalized deal/project/description text.

The grouping must not include invoice document number, service period, posting date, or amount. Those fields explain invoice entries, but they should not split the displayed company/project row.

## Current Problem

Invoice Analysis currently groups detail rows too narrowly. The detail key includes fields such as document number, contract/service period, and amount. As a result, one company/project can appear as multiple detail rows even when the intent is to review the monthly distribution of the same customer/project in one row.

## Desired Behavior

Invoice Analysis should keep the existing parent grouping by:

- Sale owner.
- Sale group.
- Last deal type.

Within each parent row, expanded detail rows should be grouped by:

- Company.
- Deal/project/description.

For each grouped detail row:

- Invoice amounts are summed into `monthly[analysisMonthKey]`.
- Multiple invoices in the same month are summed in that month.
- Invoices in different months are distributed across the month columns in the same row.
- `total` is the sum of all grouped invoice amounts.
- `postingDate` used for sorting/export is the earliest posting date in that grouped row.
- `contractRange`/service period displayed for the row should remain readable. If multiple service periods exist, use the earliest-to-latest service period range when possible; otherwise keep the first available range.
- `billingType` should remain based on the invoice sale type and billing type.

## Filters

Invoice Analysis must continue to honor:

- Sale group filter.
- Sale owner filter.
- New/Renew filter.
- Global Search.
- Invoice mode:
  - Posting Date mode uses `Posting Date`.
  - Service Period mode uses service period start, with Posting Date fallback.

## Non-Goals

- Do not change Revenue Distribution.
- Do not change Revenue Target.
- Do not change Revenue Detail.
- Do not change invoice CSV parsing fields.
- Do not add a new UI control.
- Do not attempt fuzzy matching beyond normalized company and normalized project/description.

## CSV Export Behavior

Invoice Analysis export must use the same grouped detail rows shown in the table.

CSV rows should remain sorted by:

1. `Customer/Project`
2. `Posting Date`
3. `Amount`

Date columns must remain `DD/MM/YYYY`.

The `Posting Date` column should use the earliest posting date in the grouped company/project row.

## Testing

Add or update tests to cover:

- Two invoice rows with the same company and project/description in different months produce one detail row with monthly values in both months.
- Two invoice rows with the same company and project/description in the same month are summed in that month.
- Different project/description under the same company remains a separate detail row.
- Earliest posting date is preserved for sorting/export.
- Existing revenue logic tests continue to pass.

## Documentation

Update README or context documentation to state that Invoice Analysis detail rows are grouped by company and project/description, then distributed across monthly columns.

## Self-Review

- No placeholder requirements remain.
- Scope is limited to Invoice Analysis grouping, display dataset, and export dataset.
- The grouping key is explicit.
- Non-goals protect Revenue Distribution and Revenue Target from accidental changes.
