(function exposeRevenueLogic(root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.RevenueLogic = api;
})(typeof globalThis !== "undefined" ? globalThis : window, function revenueLogicFactory() {
  function createRevenueLogic(deps) {
    const required = [
      "parseDateValue",
      "parsePositiveInt",
      "parseMoneyValue",
      "roundMoney",
      "normalizeName",
      "monthEndDate",
      "addMonths",
      "monthDiffInclusive",
      "monthRange",
      "money",
      "sum",
      "dealDetailKey",
    ];
    required.forEach((key) => {
      if (typeof deps[key] !== "function") throw new Error(`Revenue logic requires dependency: ${key}`);
    });

    const getRevenueExpectedFallback =
      typeof deps.getRevenueExpectedFallback === "function" ? deps.getRevenueExpectedFallback : () => true;

    function revenueDateParts(...values) {
      for (const value of values) {
        const parsed = deps.parseDateValue(value);
        if (parsed) return parsed;
      }
      return null;
    }

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

    function revenuePeriodMonths(deal) {
      return Number(deal.contractPeriodMonths) || deps.parsePositiveInt(deal.contractPeriodMonths);
    }

    function isOneTimeBilling(value) {
      const text = deps.normalizeName(value);
      const compact = text.replace(/\s+/g, "");
      return (
        compact === "onetime" ||
        compact === "otc" ||
        text.includes("one time") ||
        text.includes("one off") ||
        text.includes("จายครงเดยว") ||
        text.includes("ครงเดยว")
      );
    }

    function isRecurringBilling(value) {
      const text = deps.normalizeName(value);
      return (
        text.includes("recurring") ||
        text.includes("subscription") ||
        text.includes("renewal") ||
        text.includes("รายเดอน") ||
        text.includes("ตออาย")
      );
    }

    function allocateRevenueAmount(amount, monthsCount) {
      if (!monthsCount) return [];
      const sign = amount < 0 ? -1 : 1;
      const cents = Math.round(Math.abs(amount || 0) * 100);
      const base = Math.floor(cents / monthsCount);
      const remainder = cents % monthsCount;
      return Array.from({ length: monthsCount }, (_, index) => (sign * (base + (index < remainder ? 1 : 0))) / 100);
    }

    function revenueScheduleForDeal(deal) {
      const amount = deps.roundMoney(Number(deal.amount) || deps.parseMoneyValue(deal.amount));
      const contractStart = deps.parseDateValue(deal.contractStartDate);
      const expectedStart = deps.parseDateValue(deal.expectedDate);
      const fallbackStart = getRevenueExpectedFallback()
        ? revenueDateParts(deal.expectedDate, deal.startDate, deal.stageChangeDate, deal.createdDate)
        : revenueDateParts(deal.startDate, deal.stageChangeDate, deal.expectedDate, deal.createdDate);
      const start = contractStart || fallbackStart;
      const end = deps.parseDateValue(deal.contractEndDate);
      const period = revenuePeriodMonths(deal);
      const mrr = Number(deal.mrr) || deps.parseMoneyValue(deal.mrr);
      const flags = [];

      if (!amount) flags.push("Zero contract amount");
      if (!contractStart) {
        flags.push("Missing contract start date");
        if (getRevenueExpectedFallback() && expectedStart && start?.date === expectedStart.date) {
          flags.push("Using expected close date as contract start");
        }
      }
      if (!start) {
        return {
          deal,
          entries: [],
          months: [],
          contractStart: "",
          contractEnd: "",
          contractRange: "-",
          revenueRule: "No schedule",
          flags: flags.length ? flags : ["Missing contract start date"],
        };
      }

      let monthsCount = 0;
      let revenueRule = "";
      const hasValidEnd = Boolean(end && end.ts >= start.ts);
      if (end && end.ts < start.ts) flags.push("Contract end before start");

      if (isOneTimeBilling(deal.billingType)) {
        const months = [start.monthKey];
        const serviceEnd = hasValidEnd
          ? end.date
          : period > 1
            ? deps.monthEndDate(deps.addMonths(start.monthKey, period - 1))
            : start.date;
        return {
          deal,
          entries: [
            {
              deal,
              monthKey: start.monthKey,
              amount,
              monthIndex: 1,
              monthsCount: 1,
            },
          ],
          months,
          contractStart: start.date,
          contractEnd: serviceEnd,
          contractRange: `${start.date} - ${serviceEnd}`,
          revenueRule: "One-time at contract start",
          flags,
        };
      }

      if (hasValidEnd) {
        monthsCount = deps.monthDiffInclusive(start.monthKey, end.monthKey);
        revenueRule = "Contract start/end";
      }

      if (!monthsCount && period > 0) {
        monthsCount = period;
        revenueRule = "Contract period";
      }

      if (!monthsCount && mrr > 0 && amount > 0) {
        monthsCount = Math.max(1, Math.round(amount / mrr));
        revenueRule = "Inferred from MRR";
        flags.push("Contract period inferred from MRR");
      }

      if (!monthsCount) {
        monthsCount = 1;
        revenueRule = "One-month fallback";
        flags.push("Missing contract period/end date");
      }

      if (monthsCount > 120) {
        monthsCount = 120;
        flags.push("Contract period capped at 120 months");
      }

      const months = deps.monthRange(start.monthKey, monthsCount);
      const endMonth = months[months.length - 1] || start.monthKey;
      const allocations = allocateRevenueAmount(amount, months.length);
      const entries = months.map((monthKey, index) => ({
        deal,
        monthKey,
        amount: deps.roundMoney(allocations[index] || 0),
        monthIndex: index + 1,
        monthsCount: months.length,
      }));

      if (mrr > 0 && months.length) {
        const mrrTotal = deps.roundMoney(mrr * months.length);
        if (Math.abs(mrrTotal - amount) > Math.max(1, Math.abs(amount) * 0.02)) {
          flags.push(`MRR x months mismatch ${deps.money(mrrTotal)}`);
        }
      }

      const contractEnd = hasValidEnd ? end.date : deps.monthEndDate(endMonth);
      return {
        deal,
        entries,
        months,
        contractStart: start.date,
        contractEnd,
        contractRange: `${start.date} - ${contractEnd}`,
        revenueRule,
        flags,
      };
    }

    function targetYearMonths(year) {
      return Array.from({ length: 12 }, (_, index) => `${year}-${String(index + 1).padStart(2, "0")}`);
    }

    function schedulesWithMonths(schedules, monthSet) {
      return schedules
        .map((schedule) => ({
          ...schedule,
          entries: schedule.entries.filter((entry) => monthSet.has(entry.monthKey)),
          months: schedule.months.filter((month) => monthSet.has(month)),
        }))
        .filter((schedule) => schedule.entries.length);
    }

    function buildRecurringTargetFromRevenue(schedules, targetYear) {
      const sourceYear = targetYear - 1;
      const targetMonths = targetYearMonths(targetYear);
      const targetMonthSet = new Set(targetMonths);
      const rowsByKey = new Map();
      const monthlyTotals = Object.fromEntries(targetMonths.map((month) => [month, 0]));
      const dealKeys = new Set();

      schedules
        .filter((schedule) => isRecurringBilling(schedule.deal.billingType))
        .forEach((schedule) => {
          schedule.entries
            .filter((entry) => entry.monthKey.startsWith(`${sourceYear}-`))
            .forEach((entry) => {
              const targetMonth = deps.addMonths(entry.monthKey, 12);
              if (!targetMonthSet.has(targetMonth)) return;
              const deal = schedule.deal;
              const key = `${deal.saleKey || deal.saleName}|${deal.category || "-"}|${deal.group || "-"}`;
              if (!rowsByKey.has(key)) {
                rowsByKey.set(key, {
                  key,
                  saleName: deal.saleName || "-",
                  group: deal.group || "-",
                  category: deal.category || "-",
                  monthly: Object.fromEntries(targetMonths.map((month) => [month, 0])),
                  deals: new Set(),
                  companies: new Set(),
                  detailMap: new Map(),
                });
              }
              const row = rowsByKey.get(key);
              const detailKey = deps.dealDetailKey(deal);
              if (!row.detailMap.has(detailKey)) {
                row.detailMap.set(detailKey, {
                  key: detailKey,
                  company: deal.company || "-",
                  dealName: deal.dealName || deal.id || "-",
                  contractRange: schedule.contractRange || "-",
                  billingType: deal.billingType || "-",
                  monthly: Object.fromEntries(targetMonths.map((month) => [month, 0])),
                  total: 0,
                  sourceMonths: new Set(),
                });
              }
              const detail = row.detailMap.get(detailKey);
              row.monthly[targetMonth] = deps.roundMoney(row.monthly[targetMonth] + entry.amount);
              detail.monthly[targetMonth] = deps.roundMoney(detail.monthly[targetMonth] + entry.amount);
              detail.total = deps.roundMoney(detail.total + entry.amount);
              detail.sourceMonths.add(entry.monthKey);
              row.deals.add(detailKey);
              row.companies.add(deal.company || deal.dealName || deal.id || "-");
              monthlyTotals[targetMonth] = deps.roundMoney(monthlyTotals[targetMonth] + entry.amount);
              dealKeys.add(detailKey);
            });
        });

      const rows = Array.from(rowsByKey.values())
        .map((row) => ({
          ...row,
          total: deps.roundMoney(deps.sum(targetMonths, (month) => row.monthly[month] || 0)),
          dealCount: row.deals.size,
          companyCount: row.companies.size,
          details: Array.from(row.detailMap.values())
            .map((detail) => ({
              ...detail,
              sourceMonths: Array.from(detail.sourceMonths).sort(),
            }))
            .sort(
              (a, b) =>
                b.total - a.total ||
                a.company.localeCompare(b.company, "th", { numeric: true }) ||
                a.dealName.localeCompare(b.dealName, "th", { numeric: true }),
            ),
        }))
        .sort(
          (a, b) =>
            a.group.localeCompare(b.group, "th", { numeric: true }) ||
            a.saleName.localeCompare(b.saleName, "th", { numeric: true }) ||
            a.category.localeCompare(b.category, "th", { numeric: true }),
        );
      const total = deps.roundMoney(deps.sum(targetMonths, (month) => monthlyTotals[month] || 0));
      return {
        sourceYear,
        targetYear,
        months: targetMonths,
        rows,
        monthlyTotals,
        total,
        dealCount: dealKeys.size,
      };
    }

    return {
      revenueDateParts,
      revenuePeriodMonths,
      isOneTimeBilling,
      isRecurringBilling,
      invoiceProjectKey,
      earliestDateValue,
      allocateRevenueAmount,
      revenueScheduleForDeal,
      targetYearMonths,
      schedulesWithMonths,
      buildRecurringTargetFromRevenue,
    };
  }

  return { createRevenueLogic };
});
