# Revenue Distribution Grouped Detail Design

Date: 2026-06-06

## Goal

Adjust `Revenue Analysis > Revenue Distribution` so expanded detail rows display in the same grouped style as `Invoice Analysis`: one row per company/project with revenue distributed across month columns.

## Scope

This change applies only to `Revenue Analysis > Revenue Distribution` expanded detail rows and CSV export rows derived from those details.

Do not change:

- Revenue Distribution parent rows.
- Revenue Distribution summary cards.
- Revenue Distribution monthly totals.
- Revenue Target.
- Revenue Detail.
- Invoice Analysis.

## Business Rule

For Revenue Distribution detail display, "same company/project" means:

- Same normalized company name.
- Same normalized deal/project name.

Expanded details should not be split by:

- Contract range.
- Revenue month.
- Monthly amount.
- Internal deal ID.

If multiple won deals or schedules map to the same company/project under the same Sale owner, Sale group, and Last Deal Type parent row, their monthly revenue should be combined into one detail row.

## Parent Row Behavior

Parent rows remain grouped by:

- Sale owner.
- Sale group.
- Last Deal Type.

Parent monthly totals and total amount must remain unchanged.

## Detail Row Behavior

For each grouped detail row:

- `Sale` column shows the Sale owner from the parent row.
- `Customer/Project` shows company and deal/project name.
- `Sale Group / Contract` shows a readable contract range.
- `Last Deal Type / Billing` shows billing type.
- Month columns show recognized revenue by month.
- `Total` shows the sum of the grouped row.

If multiple contract ranges exist in one grouped detail row:

- Sort ranges by parsed date, not by raw string.
- Display an earliest-to-latest readable range when possible.
- Fall back to the first available range if dates cannot be parsed.

## CSV Export Behavior

Revenue Distribution export must use the same grouped detail rows as the expanded table.

Rows remain sorted by:

1. `Customer/Project`
2. `Posting Date`
3. `Amount`

Dates must export as `DD/MM/YYYY`.

`Posting Date` for a grouped row should be the earliest scheduled revenue date in that grouped detail row.

## Filters

Revenue Distribution continues to honor:

- Sale group filter.
- Sale owner filter.
- New/Renew filter.
- Global Search.
- Selected Year / Quarter / Month / Range through revenue month scoping.

The grouping must happen after these filters and revenue schedule scoping are applied.

## Non-Goals

- Do not add UI controls.
- Do not change invoice grouping.
- Do not change revenue recognition rules.
- Do not change Recurring Target logic.
- Do not fuzzy match beyond normalized company and normalized deal/project name.

## Testing

Add or update tests to cover:

- Two revenue schedules under the same company/project combine into one Revenue Distribution detail row.
- Monthly values from combined schedules are summed correctly.
- Same company with different project remains separate.
- Contract ranges in a grouped detail row are ordered chronologically.
- Parent totals remain equal to the sum of grouped detail rows.
- Export still uses grouped detail rows and sorted `Customer/Project > Posting Date > Amount` behavior.

## Documentation

Update documentation to state that Revenue Distribution expanded detail rows are grouped by company/project, while parent rows and totals remain at Sale owner / Sale group / Last Deal Type level.

## Self-Review

- No placeholder requirements remain.
- Scope is limited to Revenue Distribution expanded detail and export details.
- Parent totals and other Revenue Analysis sections are explicitly protected.
- Grouping key and non-goals are explicit.
