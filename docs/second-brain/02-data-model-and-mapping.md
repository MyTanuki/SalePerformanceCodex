# Data Model And Mapping

The app is CSV-driven and runs in the browser. The source data quality varies, so mapping and fallback rules are part of the business logic.

## Input Files

### Sale Target CSV

Defines company targets by sale owner, sale group, sale type, and period. It is used as the denominator for Sales Performance achievement.

Auto-load file when present:

- `data/Sale Target.csv`

### Deal CSV

CRM deal export. It drives Sales Performance, All Deals, Pipeline Risk, New Deals, Revenue Analysis, Revenue Distribution, and Revenue Performance.

Important fields include:

- Deal ID
- Sale owner / responsible person
- Company
- Deal name
- Pipeline
- Stage
- Deal Type
- Amount
- Expected Close Date
- Created Date
- Stage Change Date
- Contract Start Date
- Contract End Date
- Contract Period
- Billing Type
- MRR

### Mapping CSV

Maps raw CRM values to dashboard reporting values.

Auto-load file when present:

- `data/Stage Mapping.csv`

Supported mapping areas:

- Pipeline Matching by sale group and sale type.
- Stage source to dashboard stage.
- Deal Type source to New/Renew.
- Sale owner to Sale Group.

### Invoice CSV

Invoice export used for Invoice Analysis and revenue-vs-invoice comparison.

Important fields include:

- Sale owner
- Company
- Description / project text
- Posting Date
- Service Period / contract range
- Amount before VAT
- Document number
- Deal type / billing type where available

## Mapping Principles

- Stage mapping normalizes raw CRM stages before status logic runs.
- `Pre-WON` and `Won` both become Won.
- `Pre-LOST` and `Lost` both become Lost.
- Pipeline Matching decides whether a deal is counted for sales performance.
- An unmapped item means the dashboard cannot confidently classify it.
- An excluded item means it is mapped but intentionally outside counting scope.

## Counting Scope

A deal counts for sales performance only when:

- Its sale owner and sale group are known enough for reporting.
- Its sale type is known as New or Renew.
- Pipeline Matching includes that pipeline for the sale group and sale type.
- Its stage and date place it inside the selected period.

All Deals may show deals that are included, excluded, or unmapped so users can audit why totals differ from raw CRM exports.

