# Agent Handoff

Use this checklist before another AI agent changes the project.

## Current Project Path

```text
C:\CodexWS\sale-performance-dashboard
```

## Minimum Reading

1. `AGENT_CONTEXT.md`
2. `CONTEXT.md`
3. `README.md`
4. The matching `docs/second-brain/` topic file.

For chart changes also read:

- `docs/CHART_STANDARDS.md`

For revenue/invoice changes also read:

- `scripts/revenue-logic.js`
- `scripts/test-revenue-logic.js`
- `docs/revenue-invoice-export-comparison-audit-2026-06-06.md`

## Safe Change Pattern

1. Identify the business concept first.
2. Find the existing helper/render/export function.
3. Prefer pure helper changes for finance logic.
4. Add or update tests for revenue, invoice, target, or reconciliation logic.
5. Run syntax checks and relevant tests.
6. Verify UI behavior in offline preview when layout or interaction changes.
7. Do not rewrite unrelated app sections.

## Verification Commands

```powershell
node --check .\app.js
node --check .\scripts\revenue-logic.js
node .\scripts\test-revenue-logic.js
```

Preview:

```powershell
node .\scripts\preview-server.js
```

Then open:

```text
http://127.0.0.1:4174/
```

## Risk Areas

- `app.js` is large and contains many coupled render/filter/export functions.
- Filter changes can make KPI, chart, and table totals disagree.
- Revenue Target must not be affected by search-only filters.
- Invoice and revenue exports must remain spreadsheet-friendly.
- Date fields differ by tab; do not use one date rule everywhere.
- New/Renew sale type and Recurring/One-time billing type are different axes.

## When To Ask The User

Ask before changing:

- revenue recognition assumptions
- project-level reconciliation rules
- export schemas used by finance/accounting
- pipeline matching meaning
- whether Expected Close Date should replace Contract Start Date

Do not ask for routine implementation details when the existing code pattern is clear.

