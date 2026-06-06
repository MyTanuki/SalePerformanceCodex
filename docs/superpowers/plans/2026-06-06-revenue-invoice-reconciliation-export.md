# Revenue Invoice Reconciliation Export Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [x]`) syntax for tracking.

**Goal:** Add comparison columns to Revenue Distribution and Invoice Analysis exports, and add a customer-level reconciliation CSV export for tracking recognized revenue against invoice/collection progress.

**Architecture:** Add pure reconciliation helpers to `scripts/revenue-logic.js` so tests can verify key parsing, grouping, gap calculation, and statuses. Wire those helpers into `app.js` export functions. Add one UI export button in the Revenue Analysis / Invoice Analysis area. Keep existing dashboard totals and recognition logic unchanged.

**Tech Stack:** Plain JavaScript browser app, existing `scripts/revenue-logic.js` module, Node-based tests in `scripts/test-revenue-logic.js`, static HTML/CSS.

---

### Task 1: Add Pure Reconciliation Helpers and Tests

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\revenue-logic.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\test-revenue-logic.js`

- [x] **Step 1: Add helper functions in `scripts/revenue-logic.js`**

Inside `createRevenueLogic(deps)`, add:

```js
function splitCustomerProject(value) {
  const text = String(value ?? "").trim();
  if (!text) return { customer: "-", project: "" };
  const [customer, ...projectParts] = text.split("|");
  return {
    customer: customer.trim() || "-",
    project: projectParts.join("|").trim(),
  };
}

function compareKeyForCustomer(customer) {
  return deps.normalizeName(customer || "-");
}

function reconciliationStatus(revenueTotal, invoiceTotal) {
  const revenue = deps.roundMoney(revenueTotal || 0);
  const invoice = deps.roundMoney(invoiceTotal || 0);
  const gap = deps.roundMoney(revenue - invoice);
  if (revenue && !invoice) return "Revenue only";
  if (!revenue && invoice) return "Invoice only";
  if (!gap) return "Matched";
  return gap > 0 ? "Under-invoiced" : "Over-invoiced";
}
```

Export the helpers in the returned object:

```js
splitCustomerProject,
compareKeyForCustomer,
reconciliationStatus,
```

- [x] **Step 2: Add helper tests**

In `scripts/test-revenue-logic.js`, destructure the helpers from `revenueLogic` and add:

```js
assert.deepEqual(splitCustomerProject("Customer A | Project B"), {
  customer: "Customer A",
  project: "Project B",
});
assert.deepEqual(splitCustomerProject("Customer A"), {
  customer: "Customer A",
  project: "",
});
assert.equal(compareKeyForCustomer(" Customer A "), compareKeyForCustomer("customer a"));
assert.equal(reconciliationStatus(100, 100), "Matched");
assert.equal(reconciliationStatus(100, 80), "Under-invoiced");
assert.equal(reconciliationStatus(80, 100), "Over-invoiced");
assert.equal(reconciliationStatus(100, 0), "Revenue only");
assert.equal(reconciliationStatus(0, 100), "Invoice only");
```

- [x] **Step 3: Run tests**

Run:

```powershell
node --check .\scripts\revenue-logic.js
node .\scripts\test-revenue-logic.js
```

Expected:

- Syntax check exits `0`.
- `Revenue logic tests passed`.

### Task 2: Add Comparison Columns to Existing Exports

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`

- [x] **Step 1: Add wrappers near other revenue logic wrappers**

Add:

```js
function splitCustomerProject(value) {
  return revenueLogic.splitCustomerProject(value);
}

function compareKeyForCustomer(customer) {
  return revenueLogic.compareKeyForCustomer(customer);
}
```

- [x] **Step 2: Update `exportRevenueBreakdownCsv` headers**

Change headers from:

```js
const headers = [
  "Sale",
  "Customer/Project",
  ...(options.includePostingDate ? ["Posting Date"] : []),
  "Sale Group / Contract",
  "Last Deal Type / Billing",
  ...target.months.map((month) => monthLabel(month)),
  "Total",
];
```

To:

```js
const headers = [
  "Compare Key",
  "Customer",
  "Project/Description",
  "Sale",
  "Customer/Project",
  ...(options.includePostingDate ? ["Posting Date"] : []),
  "Sale Group / Contract",
  "Last Deal Type / Billing",
  ...target.months.map((month) => monthLabel(month)),
  "Total",
];
```

- [x] **Step 3: Update row building in both export branches**

After computing:

```js
const detailName = `${detail.company || "-"}${detail.dealName ? ` | ${detail.dealName}` : ""}`;
```

Add:

```js
const { customer, project } = splitCustomerProject(detailName);
const compareKey = compareKeyForCustomer(customer);
const comparisonColumns = [compareKey, customer, project];
```

Prepend `comparisonColumns` before `baseColumns` in both the `splitByMonthPostingDate` branch and the normal grouped branch.

- [x] **Step 4: Run syntax and tests**

Run:

```powershell
node --check .\app.js
node .\scripts\test-revenue-logic.js
```

Expected:

- Syntax check exits `0`.
- `Revenue logic tests passed`.

### Task 3: Add Reconciliation Export Builder

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`

- [x] **Step 1: Add wrapper for reconciliation status**

Near the wrappers added in Task 2:

```js
function reconciliationStatus(revenueTotal, invoiceTotal) {
  return revenueLogic.reconciliationStatus(revenueTotal, invoiceTotal);
}
```

- [x] **Step 2: Add `buildReconciliationRows` helper**

Add near export helpers:

```js
function buildReconciliationRows(revenueDistribution, invoiceAnalysis) {
  const months = revenueDistribution.months || invoiceAnalysis.months || [];
  const rowsByKey = new Map();

  function ensureRow(compareKey, customer, sale) {
    const key = `${compareKey}|${sale || "-"}`;
    if (!rowsByKey.has(key)) {
      rowsByKey.set(key, {
        compareKey,
        customer,
        sale: sale || "-",
        revenueProjects: new Set(),
        invoiceProjects: new Set(),
        revenueMonthly: Object.fromEntries(months.map((month) => [month, 0])),
        invoiceMonthly: Object.fromEntries(months.map((month) => [month, 0])),
      });
    }
    return rowsByKey.get(key);
  }

  function addDetails(source, side) {
    (source.rows || []).forEach((row) => {
      (row.details || []).forEach((detail) => {
        const detailName = `${detail.company || "-"}${detail.dealName ? ` | ${detail.dealName}` : ""}`;
        const { customer, project } = splitCustomerProject(detailName);
        const compareKey = compareKeyForCustomer(customer);
        const target = ensureRow(compareKey, customer, row.saleName || "-");
        const projectLabel = project || detailName;
        if (side === "revenue") target.revenueProjects.add(projectLabel);
        if (side === "invoice") target.invoiceProjects.add(projectLabel);
        months.forEach((month) => {
          const amount = Number(detail.monthly?.[month] || 0);
          if (side === "revenue") target.revenueMonthly[month] = roundMoney(target.revenueMonthly[month] + amount);
          if (side === "invoice") target.invoiceMonthly[month] = roundMoney(target.invoiceMonthly[month] + amount);
        });
      });
    });
  }

  addDetails(revenueDistribution, "revenue");
  addDetails(invoiceAnalysis, "invoice");

  return Array.from(rowsByKey.values())
    .map((row) => {
      const revenueTotal = roundMoney(sum(months, (month) => row.revenueMonthly[month] || 0));
      const invoiceTotal = roundMoney(sum(months, (month) => row.invoiceMonthly[month] || 0));
      const gapTotal = roundMoney(revenueTotal - invoiceTotal);
      return {
        ...row,
        revenueTotal,
        invoiceTotal,
        gapTotal,
        status: reconciliationStatus(revenueTotal, invoiceTotal),
      };
    })
    .sort(
      (a, b) =>
        a.customer.localeCompare(b.customer, "th", { numeric: true }) ||
        a.sale.localeCompare(b.sale, "th", { numeric: true }) ||
        a.compareKey.localeCompare(b.compareKey, "th", { numeric: true }),
    );
}
```

- [x] **Step 3: Add `exportRevenueInvoiceReconciliationCsv`**

Add:

```js
function exportRevenueInvoiceReconciliationCsv() {
  const analysis = buildRevenueAnalysis(calcScope());
  const months = analysis.revenueDistribution.months || [];
  const headers = [
    "Compare Key",
    "Customer",
    "Sale",
    "Revenue Projects",
    "Invoice Projects",
    ...months.flatMap((month) => [`${monthLabel(month)} Revenue`, `${monthLabel(month)} Invoice`, `${monthLabel(month)} Gap`]),
    "Revenue Total",
    "Invoice Total",
    "Gap Total",
    "Status",
  ];
  const rows = buildReconciliationRows(analysis.revenueDistribution, analysis.invoiceAnalysis).map((row) => [
    row.compareKey,
    row.customer,
    row.sale,
    Array.from(row.revenueProjects).sort().join(" | "),
    Array.from(row.invoiceProjects).sort().join(" | "),
    ...months.flatMap((month) => {
      const revenue = row.revenueMonthly[month] || 0;
      const invoice = row.invoiceMonthly[month] || 0;
      return [revenue, invoice, roundMoney(revenue - invoice)];
    }),
    row.revenueTotal,
    row.invoiceTotal,
    row.gapTotal,
    row.status,
  ]);
  const csv = [headers, ...rows].map((row) => row.map(csvCell).join(",")).join("\r\n");
  downloadTextFile(`${analysis.revenueDistribution.targetYear}-revenue-invoice-reconciliation.csv`, `\uFEFF${csv}`, "text/csv;charset=utf-8");
}
```

- [x] **Step 4: Run syntax check**

Run:

```powershell
node --check .\app.js
```

Expected: no output and exit code `0`.

### Task 4: Add UI Button and Event Binding

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\index.html`
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`

- [x] **Step 1: Add button to Invoice Analysis actions**

In `index.html`, near `exportInvoiceAnalysisCsv`, add:

```html
<button type="button" class="primary-button export-button" id="exportRevenueInvoiceReconciliationCsv">Export Reconciliation CSV</button>
```

- [x] **Step 2: Add element reference**

In `app.js` `els`, add:

```js
exportRevenueInvoiceReconciliationCsv: document.querySelector("#exportRevenueInvoiceReconciliationCsv"),
```

- [x] **Step 3: Add event listener**

Near export listeners:

```js
els.exportRevenueInvoiceReconciliationCsv?.addEventListener("click", exportRevenueInvoiceReconciliationCsv);
```

- [x] **Step 4: Run syntax check**

Run:

```powershell
node --check .\app.js
```

Expected: no output and exit code `0`.

### Task 5: Update Documentation

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\README.md`
- Modify: `C:\CodexWS\sale-performance-dashboard\CONTEXT.md`

- [x] **Step 1: Update README**

Add:

```md
- Revenue Distribution and Invoice Analysis exports include `Customer`, `Project/Description`, and `Compare Key` columns. The first reconciliation phase uses customer-level `Compare Key` because CRM deal names and invoice descriptions do not reliably match at project level.
- `Export Reconciliation CSV` produces a customer-level monthly comparison of Revenue, Invoice, and Gap for tracking invoice/collection progress against recognized revenue.
```

- [x] **Step 2: Update CONTEXT.md**

Add under Revenue Analysis:

```md
**Revenue Invoice Reconciliation**:
The customer-level export that compares recognized revenue against invoice amounts by month. It uses a customer-level Compare Key in the first phase because CRM deal names and invoice descriptions are not reliable project-level keys.
_Avoid_: Project-level reconciliation, collection guarantee
```

### Task 6: Final Verification and Commit

**Files:**
- Modify: `C:\CodexWS\sale-performance-dashboard\app.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\index.html`
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\revenue-logic.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\scripts\test-revenue-logic.js`
- Modify: `C:\CodexWS\sale-performance-dashboard\README.md`
- Modify: `C:\CodexWS\sale-performance-dashboard\CONTEXT.md`
- Add: `C:\CodexWS\sale-performance-dashboard\docs\superpowers\plans\2026-06-06-revenue-invoice-reconciliation-export.md`

- [x] **Step 1: Run final checks**

Run:

```powershell
node --check .\app.js
node --check .\scripts\revenue-logic.js
node .\scripts\test-revenue-logic.js
```

Expected:

- Syntax checks exit `0`.
- `Revenue logic tests passed`.

- [x] **Step 2: Review diff**

Run:

```powershell
git diff -- app.js index.html scripts/revenue-logic.js scripts/test-revenue-logic.js README.md CONTEXT.md docs/superpowers/plans/2026-06-06-revenue-invoice-reconciliation-export.md
```

Expected:

- Existing exports include comparison columns.
- New reconciliation export exists and is wired to the UI.
- Reconciliation is customer-level, not project-level.
- Existing dashboard calculations are unchanged.

- [x] **Step 3: Commit**

Run:

```powershell
git add -- app.js index.html scripts/revenue-logic.js scripts/test-revenue-logic.js README.md CONTEXT.md docs/superpowers/plans/2026-06-06-revenue-invoice-reconciliation-export.md
git commit -m "feat: export revenue invoice reconciliation"
```

Expected: commit succeeds with only relevant files.

---

## Self-Review

- Spec coverage: comparison columns, customer-level compare key, reconciliation export, monthly revenue/invoice/gap, statuses, sorting, docs, and tests are covered.
- Placeholder scan: no TBD/TODO/placeholders.
- Scope control: project-level matching remains a non-goal and is not implied by the export.

