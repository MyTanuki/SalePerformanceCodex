# Export And Reconciliation

Exports are used outside the web app in spreadsheets and by finance/accounting users. Keep exported files stable and spreadsheet-friendly.

## General CSV Rules

- Export numeric amounts as plain numbers, not formatted currency strings.
- Export dates as `DD/MM/YYYY` where dates are explicit columns.
- Avoid ambiguous values such as `2/12` in one column when spreadsheets may auto-convert them to dates.
- For month position, use separate columns such as current month index and contract duration.
- Keep headers stable unless the user explicitly approves a schema change.

## Revenue Distribution Export

Purpose:

- Export recognized revenue distribution for the selected year/month scope.

Important columns include:

- Compare Key
- Customer
- Project/Description
- Sale
- Customer/Project
- Posting Date
- Sale Group / Contract
- Last Deal Type / Billing
- month columns
- Total

Posting Date in Revenue Distribution is generated from the revenue schedule. For recurring schedules, each recognized month should have a posting date equivalent to the contract start date plus the month offset.

## Invoice Analysis Export

Purpose:

- Export invoice distribution by Posting Date or Service Period mode.

Important behavior:

- Multiple invoice rows for the same company/project are grouped.
- Posting Date mode reflects invoice/accounting timing.
- Service Period mode reflects service-month timing.

## Reconciliation Export

Purpose:

- Compare Revenue Distribution and Invoice Analysis at customer level.

Important columns:

- Compare Key
- Customer
- Sale
- Revenue Projects
- Invoice Projects
- monthly Revenue / Invoice / Gap columns
- Revenue Total
- Invoice Total
- Gap Total
- Status

Status meanings:

- Matched: total gap is zero.
- Under-invoiced: revenue is greater than invoice.
- Over-invoiced: invoice is greater than revenue.
- Revenue only: revenue exists but invoice total is zero.
- Invoice only: invoice exists but revenue total is zero.

Gap formula:

```text
Gap = Revenue - Invoice
```

## Current Reconciliation Limitation

The current safe Compare Key is customer-level.

Project-level reconciliation is not reliable yet because:

- CRM deal names and invoice descriptions are different text sources.
- Historical export audit found exact Customer/Project overlap of zero.

Do not claim project-level reconciliation until a project mapping layer exists.

