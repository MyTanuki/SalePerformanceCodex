const assert = require("node:assert/strict");
const { createRevenueLogic } = require("./revenue-logic.js");

function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeName(value) {
  return cleanText(value)
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseMoneyValue(value) {
  const text = cleanText(value).replace(/[฿,\s]/g, "");
  if (!text) return 0;
  const number = Number(text);
  return Number.isFinite(number) ? number : 0;
}

function parsePositiveInt(value) {
  const number = Math.round(parseMoneyValue(value));
  return Number.isFinite(number) && number > 0 ? number : 0;
}

function roundMoney(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function parseDateValue(value) {
  const text = cleanText(value);
  if (!text) return null;
  const match = text.match(/^(\d{1,4})[-/](\d{1,2})[-/](\d{1,4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!match) return null;

  let day;
  let month;
  let year;
  if (match[1].length === 4) {
    year = Number(match[1]);
    month = Number(match[2]);
    day = Number(match[3]);
  } else {
    day = Number(match[1]);
    month = Number(match[2]);
    year = Number(match[3]);
  }
  if (year > 2200) year -= 543;
  if (!year || !month || !day || month < 1 || month > 12 || day < 1 || day > 31) return null;
  const date = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  return {
    date,
    year,
    month,
    day,
    monthKey: `${year}-${String(month).padStart(2, "0")}`,
    ts: Date.UTC(year, month - 1, day),
  };
}

function addMonths(monthKey, offset) {
  const [year, month] = monthKey.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1 + offset, 1));
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
}

function monthDiffInclusive(startMonth, endMonth) {
  const [startYear, start] = startMonth.split("-").map(Number);
  const [endYear, end] = endMonth.split("-").map(Number);
  return Math.max(0, (endYear - startYear) * 12 + end - start + 1);
}

function monthRange(startMonth, monthsCount) {
  return Array.from({ length: Math.max(0, monthsCount) }, (_, index) => addMonths(startMonth, index));
}

function monthEndDate(monthKey) {
  const [year, month] = monthKey.split("-").map(Number);
  const end = new Date(Date.UTC(year, month, 0));
  return `${end.getUTCFullYear()}-${String(end.getUTCMonth() + 1).padStart(2, "0")}-${String(end.getUTCDate()).padStart(2, "0")}`;
}

function money(value) {
  return `฿${Math.round(value || 0).toLocaleString("th-TH")}`;
}

function sum(items, selector) {
  return items.reduce((total, item) => total + selector(item), 0);
}

function dealDetailKey(deal) {
  return [deal.id, deal.saleKey, deal.trackingMonth || deal.expectedMonth, deal.amount, deal.company]
    .map((value) => cleanText(value))
    .join("||");
}

let revenueExpectedFallback = true;

const revenueLogic = createRevenueLogic({
  parseDateValue,
  parsePositiveInt,
  parseMoneyValue,
  roundMoney,
  normalizeName,
  monthEndDate,
  addMonths,
  monthDiffInclusive,
  monthRange,
  money,
  sum,
  dealDetailKey,
  getRevenueExpectedFallback: () => revenueExpectedFallback,
});

function deal(overrides = {}) {
  const value = (key, fallback) => (Object.prototype.hasOwnProperty.call(overrides, key) ? overrides[key] : fallback);
  return {
    id: value("id", "D-1"),
    saleKey: value("saleKey", "sale-a"),
    saleName: value("saleName", "Sale A"),
    group: value("group", "AM1"),
    company: value("company", "Company A"),
    dealName: value("dealName", "Deal A"),
    category: value("category", "new"),
    amount: value("amount", 120000),
    billingType: value("billingType", "Recurring"),
    contractStartDate: value("contractStartDate", "2025-01-01"),
    contractEndDate: value("contractEndDate", "2025-12-31"),
    contractPeriodMonths: value("contractPeriodMonths", 0),
    mrr: value("mrr", 0),
    expectedDate: value("expectedDate", ""),
    startDate: value("startDate", ""),
    stageChangeDate: value("stageChangeDate", ""),
    createdDate: value("createdDate", ""),
    trackingMonth: value("trackingMonth", ""),
    expectedMonth: value("expectedMonth", ""),
  };
}

function amounts(entries) {
  return entries.map((entry) => entry.amount);
}

const {
  revenueScheduleForDeal,
  schedulesWithMonths,
  buildRecurringTargetFromRevenue,
  invoiceProjectKey,
  splitCustomerProject,
  compareKeyForCustomer,
  reconciliationStatus,
  earliestDateValue,
  buildInvoiceDistributionFromInvoices,
  buildReconciliationRows,
} = revenueLogic;

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

assert.equal(
  invoiceProjectKey({ company: "บริษัท ก จำกัด", dealName: "Service A Jan", documentNo: "INV-001" }),
  invoiceProjectKey({ company: " บริษัท ก จำกัด ", dealName: "service a jan", documentNo: "INV-002" }),
);
assert.notEqual(
  invoiceProjectKey({ company: "บริษัท ก จำกัด", dealName: "Service A" }),
  invoiceProjectKey({ company: "บริษัท ก จำกัด", dealName: "Service B" }),
);
assert.equal(
  invoiceProjectKey({ company: "บริษัท ก จำกัด", description: "Service A", documentNo: "INV-001" }),
  invoiceProjectKey({ company: "บริษัท ก จำกัด", description: "Service A", documentNo: "INV-002" }),
);
assert.equal(earliestDateValue("15/02/2026", "01/02/2026"), "01/02/2026");
assert.equal(earliestDateValue("", "01/03/2026"), "01/03/2026");

const invoiceDistribution = buildInvoiceDistributionFromInvoices(
  [
    {
      id: "inv-jan-a",
      documentNo: "INV-001",
      saleKey: "sale-a",
      saleName: "Sale A",
      group: "AM1",
      category: "new",
      company: "Company A",
      dealName: "Project Alpha",
      postingDate: "15/02/2026",
      contractRange: "15/01/2026 - 31/01/2026",
      billingType: "Recurring",
      amount: 100,
      analysisMonthKey: "2026-01",
    },
    {
      id: "inv-feb-a",
      documentNo: "INV-002",
      saleKey: "sale-a",
      saleName: "Sale A",
      group: "AM1",
      category: "new",
      company: " company a ",
      dealName: "project alpha",
      postingDate: "01/02/2026",
      contractRange: "01/02/2026 - 28/02/2026",
      billingType: "Recurring",
      amount: 200,
      analysisMonthKey: "2026-02",
    },
    {
      id: "inv-feb-b",
      documentNo: "INV-003",
      saleKey: "sale-a",
      saleName: "Sale A",
      group: "AM1",
      category: "new",
      company: "Company A",
      dealName: "Project Alpha",
      postingDate: "10/02/2026",
      contractRange: "15/02/2026 - 28/02/2026",
      billingType: "Recurring",
      amount: 50,
      analysisMonthKey: "2026-02",
    },
    {
      id: "inv-beta",
      documentNo: "INV-004",
      saleKey: "sale-a",
      saleName: "Sale A",
      group: "AM1",
      category: "new",
      company: "Company A",
      dealName: "Project Beta",
      postingDate: "05/03/2026",
      contractRange: "01/03/2026 - 31/03/2026",
      billingType: "Recurring",
      amount: 75,
      analysisMonthKey: "2026-03",
    },
  ],
  2026,
);
assert.equal(invoiceDistribution.rows.length, 1);
assert.equal(invoiceDistribution.dealCount, 2);
const invoiceDetails = invoiceDistribution.rows[0].details;
assert.equal(invoiceDetails.length, 2);
const alphaDetail = invoiceDetails.find((detail) => normalizeName(detail.dealName) === "project alpha");
const betaDetail = invoiceDetails.find((detail) => normalizeName(detail.dealName) === "project beta");
assert.ok(alphaDetail);
assert.ok(betaDetail);
assert.equal(alphaDetail.monthly["2026-01"], 100);
assert.equal(alphaDetail.monthly["2026-02"], 250);
assert.equal(alphaDetail.total, 350);
assert.equal(alphaDetail.postingDate, "01/02/2026");
assert.equal(alphaDetail.contractRange, "15/01/2026 - 31/01/2026 | 15/02/2026 - 28/02/2026");
assert.deepEqual(alphaDetail.contractRanges, [
  "15/01/2026 - 31/01/2026",
  "01/02/2026 - 28/02/2026",
  "15/02/2026 - 28/02/2026",
]);
assert.deepEqual(alphaDetail.sourceMonths, ["2026-01", "2026-02"]);
assert.deepEqual(alphaDetail.documentNos, ["INV-001", "INV-002", "INV-003"]);
assert.equal(betaDetail.monthly["2026-03"], 75);
assert.equal(betaDetail.total, 75);

const reconciliationRows = buildReconciliationRows(
  {
    months: ["2026-01", "2026-02"],
    rows: [
      {
        saleName: "Sale A",
        details: [
          {
            company: "Customer A",
            dealName: "CRM Project",
            monthly: { "2026-01": 100, "2026-02": 80 },
          },
        ],
      },
      {
        saleName: "Sale A",
        details: [
          {
            company: "Revenue Only Co",
            dealName: "Revenue Project",
            monthly: { "2026-01": 50, "2026-02": 0 },
          },
        ],
      },
      {
        saleName: "Sale A",
        details: [
          {
            company: "Matched Co",
            dealName: "Matched Project",
            monthly: { "2026-01": 20, "2026-02": 0 },
          },
        ],
      },
    ],
  },
  {
    months: ["2026-01", "2026-02"],
    rows: [
      {
        saleName: "Sale A",
        details: [
          {
            company: " customer a ",
            dealName: "Invoice Service",
            monthly: { "2026-01": 60, "2026-02": 100 },
          },
        ],
      },
      {
        saleName: "Sale A",
        details: [
          {
            company: "Invoice Only Co",
            dealName: "Invoice Project",
            monthly: { "2026-01": 30, "2026-02": 0 },
          },
        ],
      },
      {
        saleName: "Sale A",
        details: [
          {
            company: "Matched Co",
            dealName: "Matched Invoice",
            monthly: { "2026-01": 20, "2026-02": 0 },
          },
        ],
      },
    ],
  },
);
const customerAReconciliation = reconciliationRows.find((row) => row.compareKey === compareKeyForCustomer("Customer A"));
const revenueOnlyReconciliation = reconciliationRows.find((row) => row.compareKey === compareKeyForCustomer("Revenue Only Co"));
const invoiceOnlyReconciliation = reconciliationRows.find((row) => row.compareKey === compareKeyForCustomer("Invoice Only Co"));
const matchedReconciliation = reconciliationRows.find((row) => row.compareKey === compareKeyForCustomer("Matched Co"));
assert.ok(customerAReconciliation);
assert.equal(customerAReconciliation.revenueMonthly["2026-01"], 100);
assert.equal(customerAReconciliation.invoiceMonthly["2026-01"], 60);
assert.equal(customerAReconciliation.revenueTotal, 180);
assert.equal(customerAReconciliation.invoiceTotal, 160);
assert.equal(customerAReconciliation.gapTotal, 20);
assert.equal(customerAReconciliation.status, "Under-invoiced");
assert.ok(customerAReconciliation.revenueProjects.has("CRM Project"));
assert.ok(customerAReconciliation.invoiceProjects.has("Invoice Service"));
assert.equal(revenueOnlyReconciliation.status, "Revenue only");
assert.equal(invoiceOnlyReconciliation.status, "Invoice only");
assert.equal(matchedReconciliation.status, "Matched");

const oneTime = revenueScheduleForDeal(
  deal({
    id: "OTC-1",
    amount: 10000,
    billingType: "OTC",
    contractStartDate: "2025-02-01",
    contractEndDate: "2026-01-31",
  }),
);
assert.deepEqual(oneTime.months, ["2025-02"], "One-time/OTC revenue should be recognized once at contract start");
assert.equal(oneTime.entries.length, 1);
assert.equal(oneTime.entries[0].amount, 10000);
assert.equal(oneTime.entries[0].monthsCount, 1);
assert.equal(oneTime.contractEnd, "2026-01-31");

const recurring = revenueScheduleForDeal(
  deal({
    id: "REC-1",
    amount: 120000,
    billingType: "Recurring",
    contractStartDate: "2025-01-15",
    contractEndDate: "2025-12-14",
  }),
);
assert.equal(recurring.revenueRule, "Contract start/end");
assert.equal(recurring.entries.length, 12);
assert.deepEqual(recurring.months.slice(0, 3), ["2025-01", "2025-02", "2025-03"]);
assert.equal(amounts(recurring.entries).reduce((total, value) => total + value, 0), 120000);
assert.ok(amounts(recurring.entries).every((value) => value === 10000));

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
const oldContractSelectedSchedules = schedulesWithMonths([oldContractStillActive], new Set(["2026-02"]));
assert.equal(oldContractSelectedSchedules.length, 1);
assert.equal(oldContractSelectedSchedules[0].entries.length, 1);
assert.deepEqual(oldContractSelectedSchedules[0].months, ["2026-02"]);
assert.equal(oldContractSelectedSchedules[0].entries[0].monthKey, "2026-02");
assert.equal(oldContractSelectedSchedules[0].entries[0].amount, 10000);

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

revenueExpectedFallback = true;
const expectedFallback = revenueScheduleForDeal(
  deal({
    id: "FALLBACK-1",
    amount: 36000,
    billingType: "Recurring",
    contractStartDate: "",
    contractEndDate: "",
    contractPeriodMonths: 3,
    expectedDate: "2026-04-20",
    stageChangeDate: "2026-03-01",
    createdDate: "2026-02-01",
  }),
);
assert.equal(expectedFallback.contractStart, "2026-04-20");
assert.deepEqual(expectedFallback.months, ["2026-04", "2026-05", "2026-06"]);
assert.ok(expectedFallback.flags.includes("Missing contract start date"));
assert.ok(expectedFallback.flags.includes("Using expected close date as contract start"));

revenueExpectedFallback = false;
const stageFallback = revenueScheduleForDeal(
  deal({
    id: "FALLBACK-2",
    amount: 36000,
    billingType: "Recurring",
    contractStartDate: "",
    contractEndDate: "",
    contractPeriodMonths: 3,
    expectedDate: "2026-04-20",
    stageChangeDate: "2026-03-01",
    createdDate: "2026-02-01",
  }),
);
assert.equal(stageFallback.contractStart, "2026-03-01");
revenueExpectedFallback = true;

const mrrMismatch = revenueScheduleForDeal(
  deal({
    id: "MRR-1",
    amount: 42000,
    billingType: "Recurring",
    contractStartDate: "2026-01-01",
    contractEndDate: "2026-12-31",
    mrr: 5000,
  }),
);
assert.equal(mrrMismatch.entries.length, 12);
assert.ok(
  mrrMismatch.flags.some((flag) => flag.startsWith("MRR x months mismatch")),
  "MRR mismatch should be flagged when MRR x months materially differs from contract amount",
);

const targetSchedules = [
  revenueScheduleForDeal(
    deal({
      id: "T-REC-A",
      saleKey: "sale-a",
      saleName: "Sale A",
      group: "AM1",
      category: "new",
      amount: 120000,
      billingType: "Recurring",
      contractStartDate: "2025-01-01",
      contractEndDate: "2025-12-31",
    }),
  ),
  revenueScheduleForDeal(
    deal({
      id: "T-REC-B",
      saleKey: "sale-a",
      saleName: "Sale A",
      group: "AM1",
      category: "renew",
      amount: 60000,
      billingType: "Recurring",
      contractStartDate: "2025-11-01",
      contractEndDate: "2026-10-31",
    }),
  ),
  revenueScheduleForDeal(
    deal({
      id: "T-OTC",
      saleKey: "sale-a",
      saleName: "Sale A",
      group: "AM1",
      category: "new",
      amount: 99999,
      billingType: "One-time",
      contractStartDate: "2025-03-01",
      contractEndDate: "2025-03-31",
    }),
  ),
];
const target = buildRecurringTargetFromRevenue(targetSchedules, 2026);
assert.equal(target.sourceYear, 2025);
assert.equal(target.targetYear, 2026);
assert.equal(target.monthlyTotals["2026-01"], 10000);
assert.equal(target.monthlyTotals["2026-10"], 10000);
assert.equal(target.monthlyTotals["2026-11"], 15000);
assert.equal(target.monthlyTotals["2026-12"], 15000);
assert.equal(target.total, 130000);
assert.equal(target.dealCount, 2);
assert.equal(target.rows.length, 2);
assert.equal(target.rows[0].details.length, 1);
assert.ok(target.rows.every((row) => row.details.every((detail) => detail.total > 0)));
assert.deepEqual(target.rows.find((row) => row.category === "renew").details[0].sourceMonths, ["2025-11", "2025-12"]);

console.log("Revenue logic tests passed");
