# Agent Context: Sale Performance Dashboard

Last updated: 2026-06-15

This file is the first document an AI agent should read before changing this project. It summarizes the business intent, core rules, file structure, and safe working practices. Detailed context lives in `docs/second-brain/`.

## Project Purpose

Sale Performance Dashboard is a static CSV-driven web app for tracking:

- Sales achievement against company targets.
- New and Renew performance by sale owner and sale group.
- Counted pipeline and excluded/unmapped pipeline logic.
- Revenue recognition from Won / Pre-WON deals.
- Invoice distribution and revenue-vs-invoice reconciliation.
- New lead creation by ISO Week.

The app is used by Sales, Finance, Accounting, and Management. Do not treat it as only a sales dashboard; several views carry finance/accounting assumptions.

## Important Files

- `index.html`: dashboard layout, tabs, setting modal, upload inputs.
- `app.js`: app state, CSV parsing, filters, rendering, export functions.
- `styles.css`: light/dark theme, cards, tables, charts.
- `scripts/revenue-logic.js`: pure revenue recognition, recurring target, invoice distribution, reconciliation helpers.
- `scripts/test-revenue-logic.js`: Node tests for revenue logic.
- `CONTEXT.md`: business glossary.
- `README.md`: user-facing setup and behavior summary.
- `docs/CHART_STANDARDS.md`: chart drawing standards.
- `docs/second-brain/`: agent handoff and deeper project context.

## Run And Verify

Offline preview:

```powershell
node .\scripts\preview-server.js
```

Open:

```text
http://127.0.0.1:4174/
```

Revenue logic tests:

```powershell
node .\scripts\test-revenue-logic.js
```

Syntax checks:

```powershell
node --check .\app.js
node --check .\scripts\revenue-logic.js
```

## Core Business Rules

- `Pre-WON` and `Won` are treated as Won.
- `Pre-LOST` and `Lost` are treated as Lost.
- Sales Performance uses Expected Close Date for sales period tracking.
- New Deals uses Created Date and ISO Week.
- Revenue Analysis uses Contract Start Date as the primary revenue start date.
- Expected Close Date may be used as a revenue fallback only when contract start data is missing.
- One-time / OTC revenue is recognized once at the contract/service start month.
- Recurring revenue is distributed across service months.
- Recurring Target is estimated from previous-year recurring revenue, shifted into the selected year.
- Recurring Target should follow year, sale group, sale owner, and New/Renew scope, but should not be changed by keyword search or detail AND/NOT search.
- Revenue Distribution and Invoice Analysis exports include comparison columns, but the current reliable reconciliation grain is customer level, not project level.

## Data Loading

The app starts from `data/dashboard-data.js`, which is intentionally an empty skeleton.

Auto-loaded setup files when present:

- `data/Stage Mapping.csv`
- `data/Sale Target.csv`

Deal and Invoice data are uploaded from the Setting modal. Uploaded data may be retained in browser local storage.

## Before Changing Code

1. Read `CONTEXT.md`.
2. Read the relevant second-brain document for the area you are touching.
3. For revenue/invoice changes, edit pure helpers in `scripts/revenue-logic.js` first when possible.
4. Add or update `scripts/test-revenue-logic.js` for finance/accounting logic.
5. Keep exports spreadsheet-friendly: stable columns, numeric amounts, DD/MM/YYYY dates where date strings are exported.
6. Avoid broad refactors in `app.js` unless the task explicitly asks for architecture cleanup.

## Known Design Boundaries

- Sales target achievement and revenue recognition are different concepts.
- Renew sale type is not the same as Recurring billing.
- New sale type is not the same as One-time billing.
- Revenue Performance measures recurring revenue performance, not sales booking.
- Invoice Analysis can be viewed by Posting Date or Service Period.
- Revenue-vs-invoice reconciliation is customer-level until a reliable CRM-to-invoice project mapping exists.

## More Context

Start here:

- `docs/second-brain/README.md`

