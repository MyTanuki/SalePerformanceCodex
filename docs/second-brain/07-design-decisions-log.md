# Design Decisions Log

This is a curated memory of important decisions from the project history. It is not a full transcript.

## Data Loading

- `dashboard-data.js` should be an empty skeleton.
- The web app should start without deal data.
- `Stage Mapping.csv` and `Sale Target.csv` may auto-load from `data/`.
- Deal CSV and Invoice CSV are uploaded manually through Setting.

## Mapping

- `Stage Mapping Button.csv` references were removed from the system.
- Pipeline Matching is maintained in the Setting UI rather than requiring upload of a separate button mapping file.
- Pipeline Matching controls whether sales amount is counted for each sale group and New/Renew type.

## Stage Logic

- Pre-WON and Won are one successful status for reporting.
- Pre-LOST and Lost are one unsuccessful status for reporting.
- Stage source values must map from source column to display stage. Example: source `Commit` should remain `Commit`, not become Open.

## Filtering

- A central filtered-view approach was introduced because table changes and KPI/chart refreshes could drift apart.
- New Deals should use Created Date and its own ISO Week filters.
- Global Week filtering was removed.
- Global period filtering uses Year, Quarter, Month, and Range.

## Revenue

- Revenue Analysis uses contract/service timing, not sales tracking timing.
- One-time/OTC revenue is recognized once at contract start.
- Recurring revenue is distributed across contract months.
- Expected Close Date is only a fallback for revenue start when contract start is missing.
- Recurring Target is calculated from previous-year recurring revenue and shifted to the selected year.
- Recurring Target must not be distorted by keyword search or detail searches.

## Invoice And Reconciliation

- Invoice Analysis can view invoice amounts by Posting Date or Service Period.
- Invoice rows are grouped by company and project/description.
- Revenue Distribution and Invoice Analysis exports were aligned with comparison columns.
- Reconciliation is currently customer-level because project names do not reliably match between CRM and invoice data.

## UI And Visual Standards

- Tabs use a highlighted selected state.
- Total sales uses blue.
- Renew and recurring revenue use green.
- New sales uses purple.
- Target vs Actual bars use a transparent target outline and solid actual bar.
- Actual must be able to exceed target visually; do not cap actual at target width.
- Target outline must not block hover over the actual bar.

## Documentation

- `CONTEXT.md` is the business glossary.
- `docs/CHART_STANDARDS.md` is the visual standard for chart behavior.
- `AGENT_CONTEXT.md` and `docs/second-brain/` are the cross-agent memory layer.

