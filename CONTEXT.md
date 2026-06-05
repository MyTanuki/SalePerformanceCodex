# Sale Performance Dashboard

This context defines the business language for the Sale Performance Dashboard. It exists to keep sales tracking, revenue planning, and lead analysis terms consistent when the dashboard is changed.

## Sales Performance

**Deal**:
A sales opportunity from the CRM that may become revenue for the company. A deal has an owner, customer, sale type, stage, value, and expected timeline.
_Avoid_: Transaction, row, project

**Deal ที่นับยอดขาย**:
A deal that is eligible to count toward sales performance because it belongs to the relevant sale group, sale type, and pipeline counting scope. It is the only deal set used for target achievement and performance tracking.
_Avoid_: All Deal, included row, valid row

**All Deals**:
The full deal view used for audit and follow-up, including deals that count toward sales performance and deals that do not count. It exists to explain why a deal is included, excluded, or unmapped.
_Avoid_: Deal ที่นับยอดขาย, Transaction Detail

**Transaction Detail**:
The sales performance detail view for counted deals in the selected period. It supports checking the amount behind performance charts and tracing individual counted deals.
_Avoid_: All Deals, Revenue Detail

**Pipeline Risk**:
Open counted deals that show follow-up risk, such as overdue expected close date, missing expected date, stale stage, or not contacted status. It represents sales execution risk before the deal is won or lost.
_Avoid_: Data Quality issue, Revenue risk

**Sale Type**:
The business category of a deal as either New or Renew. It is used to compare performance and target achievement by sales motion.
_Avoid_: Deal Type, Product Type

**New**:
A sale type for new business, upsell, cross-sell, or other acquisition-oriented sales activity. It represents growth beyond existing recurring renewal motion.
_Avoid_: New Lead

**Renew**:
A sale type for renewal or continuation of existing customer revenue. It represents retained or continuing business.
_Avoid_: Recurring

**Won**:
A completed successful deal that counts as achieved sales when it is in the counted sales scope. Pre-Won is treated as Won for dashboard reporting.
_Avoid_: Closed, success row

**Lost**:
A completed unsuccessful deal that did not become achieved sales. Pre-Lost is treated as Lost for dashboard reporting.
_Avoid_: Failed row, excluded deal

**Open Deal**:
A deal that is still in progress and has not yet become Won or Lost. It is used for pipeline, forecast, and risk visibility.
_Avoid_: Lead, Revenue Schedule

## Targets

**Sale Target**:
The company-defined sales goal assigned to a sale owner or sale group for a period. It is used as the denominator for achievement measurement.
_Avoid_: Revenue Target, quota file

**Achievement**:
The percentage of counted Won sales compared with the relevant Sale Target. It indicates performance against company expectations.
_Avoid_: Forecast, conversion rate

**Forecast**:
The expected value of open opportunities after applying business confidence from their current stage. It is not the same as achieved sales.
_Avoid_: Won amount, Revenue Target

## Revenue Analysis

**Revenue Schedule**:
The monthly distribution of revenue expected from Won or Pre-Won deals. It is filtered by revenue month, so contracts that started before the selected period still contribute when they recognize revenue during the selected period. It is used for finance planning and management visibility.
_Avoid_: Sales performance, pipeline forecast

**Revenue Detail**:
The monthly line-item view behind the Revenue Schedule. It is intended for reconciliation and review by finance, accounting, and management.
_Avoid_: Transaction Detail, All Deals

**Revenue Performance**:
The management view that compares actual recurring revenue recognized from Won or Pre-Won deals against the Recurring Target for the selected period. It measures revenue achievement, not sales booking achievement.
_Avoid_: Sales Performance, Revenue Detail

**Revenue Gap**:
The difference between actual recurring revenue and Recurring Target for the selected period. A positive gap means actual recurring revenue is above target; a negative gap means it is below target.
_Avoid_: Forecast gap, pipeline gap

**Target Contribution Detail**:
The recurring-revenue detail view used to explain which deals contribute to Revenue Performance against the Recurring Target. It is a performance attribution view, not an accounting reconciliation view.
_Avoid_: Revenue Detail, Recurring Target Detail

**Recurring**:
A billing pattern where contract value is recognized across multiple service months. It describes revenue recognition behavior, not sale type.
_Avoid_: Renew

**One-time**:
A billing pattern where revenue is recognized once at the beginning of the contract or service period. OTC is treated as One-time.
_Avoid_: New

**Recurring Target**:
A planning baseline for recurring revenue in the selected year, estimated from recurring revenue in the previous year. It should follow year, sale group, sale owner, and New/Renew scope, but should not change because of keyword search or detail-only search.
_Avoid_: Sale Target, search result total

**Recurring Target Detail**:
The per-deal audit view behind a Recurring Target row for a sale owner and sale type. It exists to reconcile the target baseline with Sales without changing the target calculation.
_Avoid_: Revenue Detail, Transaction Detail

**Contract Period**:
The service period over which recurring revenue should be recognized. It is defined by contract start and end dates, or by an agreed contract month count when dates are incomplete.
_Avoid_: Expected Close Date, sales period

**Expected Close Date**:
The sales tracking date for when a deal is expected to close. It is not a contract start date, but may be used as a fallback planning date when contract dates are missing.
_Avoid_: Contract Start Date

## Lead Analysis

**New Deals**:
The lead creation analysis view based on Created Date and ISO Week. It measures sales activity in creating opportunities, not revenue recognition or target achievement.
_Avoid_: New sale type, Revenue Schedule

**Lead**:
A newly created deal used to measure sales prospecting activity. It may later become an open opportunity, Won deal, or Lost deal.
_Avoid_: Won deal, revenue row

**ISO Week**:
A calendar week standard used to group lead creation activity consistently across months and years. It is used for New Deals analysis.
_Avoid_: Month, sales period

## Data Governance

**Mapping**:
The business translation between raw CRM values and dashboard reporting values. It defines stage names, sale types, sale groups, and pipeline counting scope.
_Avoid_: Filter, formula

**Pipeline Matching**:
The rule that states which pipelines count for each sale group and sale type. It controls whether a deal is counted for sales performance.
_Avoid_: Pipeline group, stage mapping

**Unmapped**:
A value that cannot be matched to the dashboard's business definitions. Unmapped items require review before they can be trusted in performance reporting.
_Avoid_: Excluded

**Excluded**:
A deal that has a known mapping but is outside the pipeline counting scope for its sale group or sale type. It is intentionally not counted toward sales performance.
_Avoid_: Unmapped, Lost

**Data Quality**:
The review area for missing, inconsistent, or unmapped information that affects dashboard trust. It is a business data issue, not a sales performance result.
_Avoid_: Pipeline Risk, Lost
