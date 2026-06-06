# Revenue Distribution vs Invoice Analysis Export Audit

Date: 2026-06-06

## Objective

Check whether exported CSV files from `Revenue Distribution` and `Invoice Analysis` can be compared to track sales revenue recognition and invoice/collection progress.

## Files Checked

- `C:\Users\noppadol.s\Downloads\2025-revenue-distribution.csv`
- `C:\Users\noppadol.s\Downloads\2025-invoice-analysis-posting.csv`
- `C:\Users\noppadol.s\Downloads\2026-revenue-distribution.csv`
- `C:\Users\noppadol.s\Downloads\2026-invoice-analysis-posting.csv`

## Schema Check

All four files have the same column structure:

- `Sale`
- `Customer/Project`
- `Posting Date`
- `Sale Group / Contract`
- `Last Deal Type / Billing`
- 12 month columns for the selected year
- `Total`

The date columns are exported as `DD/MM/YYYY`.

## Data Check Summary

| Year | Revenue Rows | Invoice Rows | Revenue Unique Customer/Project | Invoice Unique Customer/Project | Exact Customer/Project Overlap | Revenue Unique Company | Invoice Unique Company | Company Overlap |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 2025 | 9,617 | 1,624 | 1,862 | 1,525 | 0 | 1,219 | 660 | 217 |
| 2026 | 13,716 | 861 | 1,910 | 829 | 0 | 1,219 | 440 | 131 |

## Findings

1. **The CSV schemas are aligned.**
   Both Revenue Distribution and Invoice Analysis export the same high-level columns and month columns. This makes side-by-side monthly comparison technically possible.

2. **Exact `Customer/Project` matching does not work.**
   Exact overlap is `0` for both 2025 and 2026. The Revenue file uses CRM deal names, while the Invoice file uses invoice description/service text. These are not the same project key.

3. **Company-level matching is possible but partial.**
   Matching by company name only gives partial overlap:
   - 2025: 217 overlapping companies
   - 2026: 131 overlapping companies

4. **Project-level reconciliation is not reliable yet.**
   Because project names come from different source fields, a user cannot safely compare sales revenue recognition and invoice/collection progress at `Customer/Project` grain using the current exported CSVs.

5. **Revenue Distribution still has many repeated Customer/Project groups.**
   The checked exports show repeated `Customer/Project` values:
   - 2025 Revenue: 1,361 duplicate `Customer/Project` groups
   - 2026 Revenue: 1,661 duplicate `Customer/Project` groups
   This means the export is not yet fully normalized to one row per comparable customer/project.

## Conclusion

The current exported CSV files are **partially usable**:

- Suitable for high-level monthly total comparison.
- Suitable for limited company-level comparison after manual grouping.
- Not yet suitable for reliable project-level tracking of sales revenue recognition versus invoice/collection.

## Recommended Fixes

1. Add explicit comparison columns to both exports:
   - `Customer`
   - `Project/Description`
   - `Compare Key`

2. Make `Compare Key` stable and intentionally defined:
   - minimum viable version: normalized `Customer`
   - better version: normalized `Customer + mapped project`

3. Complete Revenue Distribution detail grouping so export is one row per company/project, matching the display and Invoice Analysis behavior.

4. Consider adding a separate reconciliation export that joins Revenue Distribution and Invoice Analysis into one file with:
   - Revenue by month
   - Invoice by month
   - Gap by month
   - Total Revenue
   - Total Invoice
   - Remaining / Over-collected amount

