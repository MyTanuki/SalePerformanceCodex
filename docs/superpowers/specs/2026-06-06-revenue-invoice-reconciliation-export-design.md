# Revenue vs Invoice Reconciliation Export Design

Date: 2026-06-06

## Goal

Make exported data from `Revenue Distribution` and `Invoice Analysis` usable for tracking recognized sales revenue against invoice/collection progress.

The current exports have compatible month columns, but `Customer/Project` text does not match between CRM revenue data and invoice descriptions. The first reliable reconciliation level is therefore `Customer`.

## Source Evidence

Audit file:

- `docs/revenue-invoice-export-comparison-audit-2026-06-06.md`

Observed from exported files:

- Exact `Customer/Project` overlap is `0` for 2025 and 2026.
- Company-level overlap exists but is partial.
- Current files can support high-level monthly total comparison, but not reliable project-level reconciliation.

## Scope

Implement a first reconciliation phase with:

1. Explicit comparison columns in Revenue Distribution export.
2. Explicit comparison columns in Invoice Analysis export.
3. A new reconciliation CSV export that joins Revenue Distribution and Invoice Analysis at customer level.

## Non-Goals

- Do not claim reliable project-level reconciliation yet.
- Do not add fuzzy project matching.
- Do not add project mapping maintenance UI.
- Do not change revenue recognition rules.
- Do not change invoice parsing fields.
- Do not change existing dashboard totals.

## Comparison Columns

Add these columns to both `Revenue Distribution` and `Invoice Analysis` CSV exports:

- `Customer`
- `Project/Description`
- `Compare Key`

Definitions:

- `Customer`: text before the first `|` in `Customer/Project`, trimmed.
- `Project/Description`: text after the first `|` in `Customer/Project`, trimmed. Empty if no `|` exists.
- `Compare Key`: normalized `Customer` only for this phase.

Rationale:

- Audit showed project text is not reliable across CRM and invoice sources.
- Customer-level comparison is the safest useful baseline.

## New Reconciliation Export

Add a new CSV export action in `Revenue Analysis > Invoice Analysis` or the Revenue Analysis area.

Suggested label:

- `Export Reconciliation CSV`

The export should build from the currently filtered `Revenue Distribution` and `Invoice Analysis` datasets.

Rows are grouped by:

- `Compare Key`
- `Customer`
- Sale owner when available

If multiple sales owners exist for the same customer, create separate rows by sale owner to keep responsibility traceable.

## Reconciliation CSV Columns

Core columns:

- `Compare Key`
- `Customer`
- `Sale`
- `Revenue Projects`
- `Invoice Projects`
- `Revenue Total`
- `Invoice Total`
- `Gap Total`
- `Status`

Month columns:

For every selected year month, include:

- `<MMM YYYY> Revenue`
- `<MMM YYYY> Invoice`
- `<MMM YYYY> Gap`

Definitions:

- `Revenue`: recognized revenue from Revenue Distribution for that customer/month.
- `Invoice`: invoice amount from Invoice Analysis for that customer/month.
- `Gap`: `Revenue - Invoice`.
- `Gap Total`: `Revenue Total - Invoice Total`.

Status:

- `Matched`: `Gap Total` is `0`.
- `Under-invoiced`: `Gap Total` is positive.
- `Over-invoiced`: `Gap Total` is negative.
- `Revenue only`: revenue exists but invoice total is `0`.
- `Invoice only`: invoice exists but revenue total is `0`.

Use exact numeric values, not formatted currency strings, so the CSV can be used in spreadsheets.

## Filter Behavior

The reconciliation export should use the same currently filtered datasets:

- Sale group
- Sale owner
- New/Renew
- Global Search
- Selected Year / Quarter / Month / Range
- Invoice mode (`Posting Date` or `Service Period`) for Invoice Analysis

The export should not use detail AND/NOT filters from Revenue Detail.

## Sorting

Sort reconciliation rows by:

1. `Customer`
2. `Sale`
3. `Compare Key`

## Testing

Add tests for pure reconciliation helpers:

- `splitCustomerProject("A | B")` returns customer `A`, project `B`.
- `Compare Key` normalizes customer consistently.
- Revenue and invoice rows for the same customer produce one reconciliation row.
- `Gap` is `Revenue - Invoice`.
- `Revenue only`, `Invoice only`, `Under-invoiced`, `Over-invoiced`, and `Matched` statuses are assigned correctly.

## Documentation

Update documentation to explain:

- Existing exports now include comparison columns.
- Reconciliation export is customer-level in this phase.
- Project-level reconciliation requires a future mapping between CRM Deal Name and invoice Description.

## Self-Review

- No placeholder requirements remain.
- The design does not overstate project-level reliability.
- Existing dashboard calculations are protected.
- The new reconciliation export directly addresses the audit finding.
