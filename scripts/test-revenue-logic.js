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

const { revenueScheduleForDeal, buildRecurringTargetFromRevenue } = revenueLogic;

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

console.log("Revenue logic tests passed");
