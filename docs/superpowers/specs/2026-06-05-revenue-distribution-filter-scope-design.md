# Revenue Distribution Filter Scope Design

Date: 2026-06-05

## Goal

Improve `Revenue Analysis > Revenue Distribution` so it shows all recognized revenue inside the selected filter period, including revenue from contracts that started or closed before the selected period but still have scheduled revenue during that period.

The table must still respond to Global Search. If the user searches for a customer or deal, Revenue Distribution should show only matching customer/deal revenue, even when the contract is older than the selected filter period.

## Business Rule

Revenue Distribution is a revenue-recognition view, not a sales-close-date view.

Included deals:

- Status is `Won` or `Pre-WON`.
- Deal passes pipeline counting rules.
- Deal matches selected sale group, sale owner, and sale type (`รวม`, `New`, `Renew`).
- Deal matches Global Search when a keyword is entered.

Period filtering:

- Do not exclude a deal because its `Expected Close Date` is outside the selected Year, Quarter, Month, or Range.
- Build the full revenue schedule for each included deal first.
- Then keep only schedule entries whose `Revenue Month` is inside the selected period.

This means a contract that starts in 2025 and continues recognizing revenue into 2026 will appear when the selected period is 2026, if its schedule has 2026 revenue months.

## Non-Goals

- Do not change `Recurring Target` behavior. It should continue to ignore Global Search and detail AND/NOT search.
- Do not change Sales Performance date filtering.
- Do not add a new UI filter mode.
- Do not change revenue allocation rules for recurring or one-time billing.

## Data Flow

1. Build revenue source deals from all deal details using:
   - counting scope,
   - status scope,
   - sale group / sale owner scope,
   - sale type scope,
   - Global Search scope.
2. Do not apply selected-period filtering at the deal level.
3. Convert each source deal into a full monthly revenue schedule.
4. Build Revenue Distribution from schedule entries that fall in the selected year used by the Revenue Analysis subtabs.
5. For period-specific views such as Monthly Revenue Schedule and Revenue Detail, keep only schedule entries inside `scope.monthSet`.

## UI Behavior

Revenue Distribution should show totals for the selected revenue year/subtab context. When Global Search is entered, the rows and totals should narrow to matching customer/deal text.

Existing display and collapse/expand behavior should remain unchanged.

## CSV Export Behavior

Revenue Distribution export must use the same filtered dataset shown on screen.

Rows should remain sorted by:

1. `Customer/Project`
2. `Posting Date`
3. `Amount`

Date columns must export as `DD/MM/YYYY`.

For Revenue Distribution, `Posting Date` is the scheduled revenue date for the month:

- Start from `Contract Start Date`.
- If missing and fallback is enabled, use `Expected Close Date` or the existing schedule fallback start.
- Add the schedule month offset.
- Example: start `01/01/2026`, 12 months:
  - month 1 posting date: `01/01/2026`
  - month 2 posting date: `01/02/2026`
  - until the final contract month.

## Error Handling

If a deal cannot create a schedule because no start date can be determined, it should not contribute revenue. It should remain visible through existing data checks where applicable.

If fallback is used instead of Contract Start Date, the existing warning behavior should continue.

## Testing

Add or update tests to cover:

- A contract starting before the selected period but recognizing revenue inside the selected period appears in Revenue Distribution.
- Global Search narrows Revenue Distribution to the matching customer/deal.
- Expected Close Date outside the selected period does not exclude a deal whose revenue month is inside the selected period.
- Recurring Target remains unaffected by Global Search.
- Revenue Distribution export formats all date columns as `DD/MM/YYYY`.
- Revenue Distribution export sorts by `Customer/Project`, then `Posting Date`, then `Amount`.

## Self-Review

- No placeholder requirements remain.
- Scope is limited to Revenue Distribution filtering and export consistency.
- Recurring Target behavior is explicitly protected.
- The distinction between deal-level filtering and schedule-entry filtering is explicit.
