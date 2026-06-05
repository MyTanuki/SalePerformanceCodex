# Invoice Analysis Project Monthly Distribution Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Group `Invoice Analysis` detail rows by company and project/description so all invoice amounts for the same company/project are distributed across monthly columns in one row.

**Architecture:** Keep existing parent grouping by sale owner, sale group, and last deal type. Change only the Invoice Analysis detail grouping key and grouped detail metadata, leaving Revenue Distribution, Revenue Target, and Revenue Detail unchanged.

**Tech Stack:** Plain JavaScript browser app in `app.js`, existing Node revenue logic tests in `scripts/test-revenue-logic.js`, documentation in Markdown.

---

### Task 1: Add Invoice Grouping Regression Tests

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\test-revenue-logic.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\revenue-logic.js`

- [x] **Step 1: Add pure grouping helper to revenue logic**

In `scripts/revenue-logic.js`, add pure helpers inside `createRevenueLogic(deps)` near other utility helpers:

```js
function invoiceProjectKey(invoice) {
  return `${deps.normalizeName(invoice.company)}|${deps.normalizeName(invoice.dealName || invoice.description || "")}`;
}

function earliestDateValue(currentValue, nextValue) {
  const current = deps.parseDateValue(currentValue);
  const next = deps.parseDateValue(nextValue);
  if (!current) return nextValue || currentValue || "";
  if (!next) return currentValue || nextValue || "";
  return next.ts < current.ts ? nextValue : currentValue;
}
```

Export both in the returned object:

```js
invoiceProjectKey,
earliestDateValue,
```

- [x] **Step 2: Add tests for grouping key and earliest posting date**

In `scripts/test-revenue-logic.js`, destructure the new helpers:

```js
const { revenueScheduleForDeal, buildRecurringTargetFromRevenue, schedulesWithMonths, invoiceProjectKey, earliestDateValue } = revenueLogic;
```

Add tests near the helper tests:

```js
assert.equal(
  invoiceProjectKey({ company: "บริษัท ก จำกัด", dealName: "Service A Jan" }),
  invoiceProjectKey({ company: " บริษัท ก จำกัด ", dealName: "service a jan" }),
);
assert.notEqual(
  invoiceProjectKey({ company: "บริษัท ก จำกัด", dealName: "Service A" }),
  invoiceProjectKey({ company: "บริษัท ก จำกัด", dealName: "Service B" }),
);
assert.equal(earliestDateValue("15/02/2026", "01/02/2026"), "01/02/2026");
assert.equal(earliestDateValue("", "01/03/2026"), "01/03/2026");
```

- [x] **Step 3: Run tests**

Run:

```powershell
node .\scripts\test-revenue-logic.js
```

Expected: `Revenue logic tests passed`.

### Task 2: Update Invoice Analysis Detail Grouping

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\revenue-logic.js`

- [x] **Step 1: Implement and export the invoice distribution helper in revenue logic**

In `scripts/revenue-logic.js`, implement and export `buildInvoiceDistributionFromInvoices(invoices, targetYear)` so the production grouping behavior is testable outside the browser app. The helper expects invoices to already include `analysisMonthKey`, then groups detail rows by:

```js
const detailKey = invoiceProjectKey(invoice);
```

The helper owns monthly summing, grouped metadata, earliest posting date preservation, detail normalization, row counters, and `dealCount`.

- [x] **Step 2: Delegate app Invoice Analysis building to the revenue logic helper**

In `app.js`, keep `invoiceAnalysisMonthKey(invoice)` app-side because it depends on UI state, then delegate directly to the shared helper:

```js
function buildInvoiceDistributionFromInvoices(invoices, targetYear) {
  return revenueLogic.buildInvoiceDistributionFromInvoices(
    invoices.map((invoice) => ({ ...invoice, analysisMonthKey: invoiceAnalysisMonthKey(invoice) })),
    targetYear,
  );
}
```

- [x] **Step 3: Ensure the detail grouping key excludes invoice-specific fields**

Inside the revenue logic helper, use only Company + Deal/Description via `invoiceProjectKey(invoice)`. Do not include document number, service period, posting date, or amount in the grouping key.

```js
const detailKey = invoiceProjectKey(invoice);
```

- [x] **Step 4: Track grouped invoice metadata**

When creating a detail entry, include:

```js
documentNos: new Set(),
contractRanges: new Set(),
postingDates: new Set(),
```

The detail object should still include:

```js
company: invoice.company || "-",
dealName: invoice.dealName || invoice.documentNo || "-",
postingDate: invoice.postingDate || "-",
contractRange: invoice.contractRange || invoice.postingDate || "-",
billingType: [invoice.category, invoice.billingType].filter(Boolean).join(" / ") || "-",
monthly: Object.fromEntries(months.map((month) => [month, 0])),
total: 0,
sourceMonths: new Set(),
```

- [x] **Step 5: Update grouped metadata on each invoice**

After `const detail = row.detailMap.get(detailKey);`, add:

```js
detail.postingDate = earliestDateValue(detail.postingDate, invoice.postingDate || "");
if (invoice.contractRange) detail.contractRanges.add(invoice.contractRange);
if (invoice.postingDate) detail.postingDates.add(invoice.postingDate);
if (invoice.documentNo || invoice.id) detail.documentNos.add(invoice.documentNo || invoice.id);
```

- [x] **Step 6: Preserve row counters**

Keep the existing monthly summing:

```js
row.monthly[invoice.analysisMonthKey] = roundMoney(row.monthly[invoice.analysisMonthKey] + invoice.amount);
detail.monthly[invoice.analysisMonthKey] = roundMoney(detail.monthly[invoice.analysisMonthKey] + invoice.amount);
detail.total = roundMoney(detail.total + invoice.amount);
detail.sourceMonths.add(invoice.analysisMonthKey);
```

Change row deal counting to use the grouped project key:

```js
row.deals.add(detailKey);
invoiceKeys.add(detailKey);
```

- [x] **Step 7: Normalize grouped detail output**

When converting detail map values, include arrays:

```js
sourceMonths: Array.from(detail.sourceMonths).sort(),
documentNos: Array.from(detail.documentNos).sort(),
contractRanges: Array.from(detail.contractRanges).sort(),
postingDates: Array.from(detail.postingDates).sort(),
```

If `contractRanges.length > 1`, keep `contractRange` as the earliest-to-latest readable range by using the first and last sorted values:

```js
contractRange: detail.contractRanges.size > 1
  ? `${Array.from(detail.contractRanges).sort()[0]} | ${Array.from(detail.contractRanges).sort().at(-1)}`
  : detail.contractRange,
```

- [x] **Step 8: Run syntax check**

Run:

```powershell
node --check .\app.js
```

Expected: no output and exit code `0`.

### Task 3: Verify Export Uses Grouped Detail Rows

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`

- [x] **Step 1: Confirm no export-specific regrouping exists**

Confirm `exportInvoiceAnalysisCsv()` still calls:

```js
const invoiceAnalysis = buildRevenueAnalysis(calcScope()).invoiceAnalysis;
exportRevenueBreakdownCsv(invoiceAnalysis, `${invoiceAnalysis.targetYear}-invoice-analysis-${state.invoiceDateMode}.csv`, {
  includePostingDate: true,
  formatContractRange: csvDateRange,
});
```

- [x] **Step 2: Confirm export sorting and date formatting still use grouped fields**

Confirm grouped rows flow through `exportRevenueBreakdownCsv()` and use:

```js
csvDate(postingDate)
csvDateRange(detail.contractRange || "-")
exportCustomerPostingAmountSort
```

- [x] **Step 3: Run syntax check and tests**

Run:

```powershell
node --check .\app.js
node .\scripts\test-revenue-logic.js
```

Expected:

- `node --check` exits `0`.
- `Revenue logic tests passed`.

### Task 4: Update Documentation

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\README.md`
- Modify: `C:\CodexWS\sale-performance-dashboard\CONTEXT.md`

- [x] **Step 1: Update README**

Add or update the Invoice Analysis note:

```md
- Invoice Analysis groups invoice detail rows by Company + Deal/Description, then distributes invoice amounts across monthly columns using Posting Date or Service Period mode. Multiple invoice rows for the same company/project are combined into one monthly distribution row.
```

- [x] **Step 2: Update CONTEXT.md**

Under `Revenue Analysis`, add:

```md
**Invoice Analysis**:
The invoice-based monthly distribution view. Detail rows are grouped by company and project/description, then invoice amounts are distributed across month columns using either Posting Date or Service Period mode.
_Avoid_: Revenue Detail, Invoice raw rows
```

- [x] **Step 3: Review docs diff**

Run:

```powershell
git diff -- README.md CONTEXT.md
```

Expected: only the intended documentation wording changes.

### Task 5: Final Verification and Commit

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\revenue-logic.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\test-revenue-logic.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\README.md`
- Modify: `C:\CodexWS\sale-performance-dashboard\CONTEXT.md`
- Add: `C:\CodexWS\sale-performance-dashboard\docs\superpowers\plans\2026-06-05-invoice-analysis-project-monthly-distribution.md`

- [x] **Step 1: Run final checks**

Run:

```powershell
node --check .\app.js
node .\scripts\test-revenue-logic.js
```

Expected:

- `node --check` exits `0`.
- `Revenue logic tests passed`.

- [x] **Step 2: Review changed files**

Run:

```powershell
git diff -- app.js scripts/revenue-logic.js scripts/test-revenue-logic.js README.md CONTEXT.md docs/superpowers/plans/2026-06-05-invoice-analysis-project-monthly-distribution.md
```

Expected:

- Invoice detail grouping key no longer includes document number, contract/service period, posting date, or amount.
- Same company/project invoices distribute across month columns in one detail row.
- Different project/description under the same company remains separate.
- Export uses grouped rows.
- Revenue Distribution/Target/Detail logic is unchanged except shared helper additions.

- [x] **Step 3: Commit only files changed for this feature**

Run:

```powershell
git add -- app.js scripts/revenue-logic.js scripts/test-revenue-logic.js README.md CONTEXT.md docs/superpowers/plans/2026-06-05-invoice-analysis-project-monthly-distribution.md
git commit -m "feat: group invoice analysis by project"
```

Expected: commit succeeds with only relevant files.

---

## Self-Review

- Spec coverage: plan covers grouping by Company + Deal/Description, monthly distribution in one row, earliest posting date, export reuse, docs, and tests.
- Placeholder scan: no TBD/TODO/placeholders.
- Type consistency: `invoiceProjectKey`, `earliestDateValue`, and `buildInvoiceDistributionFromInvoices` are implemented and exported from `scripts/revenue-logic.js`; `app.js` delegates Invoice Analysis distribution directly to `revenueLogic.buildInvoiceDistributionFromInvoices(...)`.
