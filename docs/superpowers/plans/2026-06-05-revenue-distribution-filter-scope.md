# Revenue Distribution Filter Scope Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `Revenue Analysis > Revenue Distribution` include all revenue recognized inside the selected filter period, including older contracts, while still honoring Global Search.

**Architecture:** Keep deal-level filtering for sale scope, sale type, counting scope, status, and Global Search, but do not apply selected-period filtering before creating revenue schedules. Add small testable helpers in `app.js` for revenue distribution source selection and export row sorting/formatting, without restructuring the whole dashboard.

**Tech Stack:** Plain JavaScript browser app, existing `scripts/revenue-logic.js`, Node syntax check, existing Node revenue logic tests.

---

### Task 1: Add Focused Revenue Distribution Tests

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\test-revenue-logic.js`
- Reference: `C:\CodexWS\sale-performance-dashboard\scripts\revenue-logic.js`

- [ ] **Step 1: Add a regression test for old contracts with in-period revenue**

Append this test near the existing recurring schedule tests:

```js
const oldContractStillActive = revenueScheduleForDeal(
  deal({
    id: "old-active",
    amount: 120000,
    billingType: "Recurring",
    contractStartDate: "2025-07-01",
    contractEndDate: "2026-06-30",
    expectedDate: "2025-06-15",
  }),
);
assert.equal(oldContractStillActive.entries.length, 12);
assert.equal(oldContractStillActive.entries.filter((entry) => entry.monthKey.startsWith("2026-")).length, 6);
assert.equal(
  oldContractStillActive.entries
    .filter((entry) => entry.monthKey.startsWith("2026-"))
    .reduce((total, entry) => total + entry.amount, 0),
  60000,
);
```

- [ ] **Step 2: Add a regression test for fallback expected close date schedule**

Append this after the old contract test:

```js
const fallbackSchedule = revenueScheduleForDeal(
  deal({
    id: "fallback-active",
    amount: 60000,
    billingType: "Recurring",
    contractStartDate: "",
    contractEndDate: "",
    contractPeriodMonths: "6",
    expectedDate: "2026-01-15",
  }),
);
assert.equal(fallbackSchedule.contractStart, "2026-01-15");
assert.equal(fallbackSchedule.entries[0].monthKey, "2026-01");
assert.equal(fallbackSchedule.entries[5].monthKey, "2026-06");
assert.equal(fallbackSchedule.entries[5].monthIndex, 6);
assert.equal(fallbackSchedule.entries[5].monthsCount, 6);
```

- [ ] **Step 3: Run the revenue tests**

Run:

```powershell
node .\scripts\test-revenue-logic.js
```

Expected: `Revenue logic tests passed`.

### Task 2: Add Source Selection Helper for Revenue Distribution

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`

- [ ] **Step 1: Add a helper below `invoiceRowsForAnalysis(scope)` or near `buildRevenueAnalysis(scope)`**

Add:

```js
function revenueDistributionDeals(scope) {
  return filteredDeals(scope, { period: false, countingIncluded: true, globalSearch: true })
    .filter((deal) => deal.status === "won");
}
```

This explicitly documents that Revenue Distribution ignores selected-period deal filtering but honors Global Search.

- [ ] **Step 2: Update `buildRevenueAnalysis(scope)` to use the helper**

Replace:

```js
const baseDeals = filteredDeals(scope, { period: false, countingIncluded: true })
  .filter((deal) => deal.status === "won");
```

With:

```js
const baseDeals = revenueDistributionDeals(scope);
```

Keep this block unchanged:

```js
const targetBaseDeals = filteredDeals(scope, { period: false, countingIncluded: true, globalSearch: false })
  .filter((deal) => deal.status === "won");
```

This preserves the rule that Recurring Target ignores Global Search.

- [ ] **Step 3: Run syntax check**

Run:

```powershell
node --check .\app.js
```

Expected: no output and exit code `0`.

### Task 3: Constrain Revenue Distribution Display to Selected Revenue Months

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`

- [ ] **Step 1: Add a schedule filter helper near `buildRevenueDistributionFromRevenue`**

Add:

```js
function schedulesWithMonths(schedules, monthSet) {
  return schedules
    .map((schedule) => ({
      ...schedule,
      entries: schedule.entries.filter((entry) => monthSet.has(entry.monthKey)),
      months: schedule.months.filter((month) => monthSet.has(month)),
    }))
    .filter((schedule) => schedule.entries.length);
}
```

- [ ] **Step 2: Use selected months for Revenue Distribution**

In `buildRevenueAnalysis(scope)`, replace:

```js
const revenueDistribution = buildRevenueDistributionFromRevenue(schedules, revenueYear);
```

With:

```js
const revenueDistribution = buildRevenueDistributionFromRevenue(schedulesWithMonths(schedules, scope.monthSet), revenueYear);
```

This ensures the displayed distribution and its export reflect only selected Year/Quarter/Month/Range entries, while still allowing old contracts into the source schedule.

- [ ] **Step 3: Keep annual target behavior unchanged**

Confirm this line remains:

```js
const recurringTarget = buildRecurringTargetFromRevenue(targetSchedules, revenueYear);
```

- [ ] **Step 4: Run syntax check**

Run:

```powershell
node --check .\app.js
```

Expected: no output and exit code `0`.

### Task 4: Verify CSV Export Sorting and Date Formatting

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`

- [ ] **Step 1: Confirm export sorting helper remains in use**

Confirm `exportRevenueBreakdownCsv()` still sorts before writing CSV:

```js
const rows = exportRows
  .sort(exportCustomerPostingAmountSort)
  .map((item) => item.row);
```

- [ ] **Step 2: Confirm Revenue Distribution export uses monthly posting dates**

Confirm `exportRevenueDistributionCsv()` still passes:

```js
includePostingDate: true,
splitByMonthPostingDate: true,
formatContractRange: csvDateRange,
```

- [ ] **Step 3: Confirm date formatting**

Confirm `exportRevenueBreakdownCsv()` uses `csvDate(...)` for Posting Date and `csvDateRange(...)` for contract ranges. The export row for split monthly posting date should include:

```js
csvDate(detail.postingDatesByMonth?.[month] || `${month}-01`),
formatContractRange(detail.contractRange || "-"),
```

- [ ] **Step 4: Run syntax check and revenue tests**

Run:

```powershell
node --check .\app.js
node .\scripts\test-revenue-logic.js
```

Expected:

- `node --check` exits `0`.
- `Revenue logic tests passed`.

### Task 5: Update Documentation

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\README.md`
- Modify: `C:\CodexWS\sale-performance-dashboard\CONTEXT.md`

- [ ] **Step 1: Update README core assumptions**

Replace the Revenue Analysis assumption line with:

```md
- Revenue Analysis actual/detail rows and Revenue Distribution build schedules from Won / Pre-WON counted deals that match view/search/type scope, then filter schedule entries by the selected Year, Quarter, Month, or Range. Older contracts are included when they still recognize revenue inside the selected period. Recurring Target follows year/group/sale/New-Renew counting scope but ignores keyword search and detail AND/NOT search.
```

- [ ] **Step 2: Update CONTEXT.md Revenue Schedule definition**

Replace the `Revenue Schedule` definition with:

```md
**Revenue Schedule**:
The monthly distribution of revenue expected from Won or Pre-Won deals. It is filtered by revenue month, so contracts that started before the selected period still contribute when they recognize revenue during the selected period. It is used for finance planning and management visibility.
_Avoid_: Sales performance, pipeline forecast
```

- [ ] **Step 3: Run a docs diff review**

Run:

```powershell
git diff -- README.md CONTEXT.md
```

Expected: only the intended wording changes.

### Task 6: Final Verification and Commit

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\test-revenue-logic.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\README.md`
- Modify: `C:\CodexWS\sale-performance-dashboard\CONTEXT.md`

- [ ] **Step 1: Run final checks**

Run:

```powershell
node --check .\app.js
node .\scripts\test-revenue-logic.js
```

Expected:

- `node --check` exits `0`.
- `Revenue logic tests passed`.

- [ ] **Step 2: Review changed files**

Run:

```powershell
git diff -- app.js scripts/test-revenue-logic.js README.md CONTEXT.md
```

Expected:

- Revenue Distribution source selection is explicit.
- Revenue Distribution schedule entries are filtered by `scope.monthSet`.
- Recurring Target still ignores Global Search.
- CSV date/sort behavior is preserved.
- Docs match the implemented rule.

- [ ] **Step 3: Commit only files changed for this feature**

Run:

```powershell
git add -- app.js scripts/test-revenue-logic.js README.md CONTEXT.md docs/superpowers/plans/2026-06-05-revenue-distribution-filter-scope.md
git commit -m "fix: include active contract revenue in distribution"
```

Expected: commit succeeds with only the relevant files.

---

## Self-Review

- Spec coverage: plan covers old contracts, Global Search, schedule-month filtering, Recurring Target isolation, export date formatting, and export sorting.
- Placeholder scan: no TBD/TODO/placeholders.
- Type consistency: helper names are `revenueDistributionDeals`, `schedulesWithMonths`, `exportCustomerPostingAmountSort`, `csvDate`, and `csvDateRange`; these match current code naming.
