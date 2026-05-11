const dashboardData = window.DASHBOARD_DATA;

const state = {
  periodMode: "ytd",
  month: "2026-05",
  group: "all",
  sale: "all",
  category: "all",
  showUnmapped: false,
  search: "",
  tab: "overview",
};

const monthLabels = {
  "2026-01": "Jan",
  "2026-02": "Feb",
  "2026-03": "Mar",
  "2026-04": "Apr",
  "2026-05": "May",
  "2026-06": "Jun",
  "2026-07": "Jul",
  "2026-08": "Aug",
  "2026-09": "Sep",
  "2026-10": "Oct",
  "2026-11": "Nov",
  "2026-12": "Dec",
};

const els = {
  metaBox: document.querySelector("#metaBox"),
  periodMode: document.querySelector("#periodMode"),
  monthSelect: document.querySelector("#monthSelect"),
  groupSelect: document.querySelector("#groupSelect"),
  saleSelect: document.querySelector("#saleSelect"),
  searchInput: document.querySelector("#searchInput"),
  showUnmapped: document.querySelector("#showUnmapped"),
  kpiGrid: document.querySelector("#kpiGrid"),
  monthlyChart: document.querySelector("#monthlyChart"),
  teamChart: document.querySelector("#teamChart"),
  categoryMix: document.querySelector("#categoryMix"),
  statusMix: document.querySelector("#statusMix"),
  topDealsTable: document.querySelector("#topDealsTable"),
  salesTable: document.querySelector("#salesTable"),
  riskSummary: document.querySelector("#riskSummary"),
  riskByStage: document.querySelector("#riskByStage"),
  riskDealsTable: document.querySelector("#riskDealsTable"),
  assumptionList: document.querySelector("#assumptionList"),
  normalizationStats: document.querySelector("#normalizationStats"),
  unmappedTable: document.querySelector("#unmappedTable"),
};

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function normalizeSearch(value) {
  return String(value ?? "").trim().toLowerCase();
}

function compactMoney(value) {
  const abs = Math.abs(value);
  const sign = value < 0 ? "-" : "";
  if (abs >= 1_000_000_000) return `${sign}฿${(abs / 1_000_000_000).toFixed(2)}B`;
  if (abs >= 1_000_000) return `${sign}฿${(abs / 1_000_000).toFixed(1)}M`;
  if (abs >= 1_000) return `${sign}฿${(abs / 1_000).toFixed(0)}K`;
  return `${sign}฿${abs.toLocaleString("th-TH", { maximumFractionDigits: 0 })}`;
}

function money(value) {
  return `฿${Math.round(value || 0).toLocaleString("th-TH")}`;
}

function percent(value) {
  if (!Number.isFinite(value)) return "-";
  return `${(value * 100).toFixed(value >= 1 ? 0 : 1)}%`;
}

function sum(items, selector) {
  return items.reduce((total, item) => total + selector(item), 0);
}

function periodMonths() {
  if (state.periodMode === "year") return dashboardData.months;
  if (state.periodMode === "month") return [state.month];
  const endIndex = dashboardData.months.indexOf(state.month);
  return dashboardData.months.slice(0, endIndex + 1);
}

function categoryMatches(category) {
  return state.category === "all" || category === state.category;
}

function visibleSales() {
  return dashboardData.sales.filter((sale) => {
    if (!state.showUnmapped && !sale.hasPositiveTarget) return false;
    if (state.group !== "all" && sale.group !== state.group) return false;
    if (state.sale !== "all" && sale.key !== state.sale) return false;
    return true;
  });
}

function salesForOptions() {
  return dashboardData.sales.filter((sale) => {
    if (!state.showUnmapped && !sale.hasPositiveTarget) return false;
    if (state.group !== "all" && sale.group !== state.group) return false;
    return true;
  });
}

function targetForSale(sale, months, category = state.category) {
  return months.reduce((total, month) => {
    const bucket = sale.monthly[month] || { renew: 0, new: 0, total: 0 };
    if (category === "renew") return total + bucket.renew;
    if (category === "new") return total + bucket.new;
    return total + bucket.total;
  }, 0);
}

function calcScope() {
  const months = periodMonths();
  const monthSet = new Set(months);
  const sales = visibleSales();
  const saleSet = new Set(sales.map((sale) => sale.key));
  const facts = dashboardData.facts.filter(
    (fact) => saleSet.has(fact.saleKey) && monthSet.has(fact.monthKey) && categoryMatches(fact.category),
  );
  const actualFacts = facts.filter((fact) => fact.status === "won");
  const lostFacts = facts.filter((fact) => fact.status === "lost");
  const openAll = dashboardData.openDeals.filter((deal) => saleSet.has(deal.saleKey) && categoryMatches(deal.category));
  const openPeriod = openAll.filter((deal) => deal.expectedMonth && monthSet.has(deal.expectedMonth));

  const target = sum(sales, (sale) => targetForSale(sale, months));
  const actual = sum(actualFacts, (fact) => fact.amount);
  const lost = sum(lostFacts, (fact) => fact.amount);
  const openPipeline = sum(openPeriod, (deal) => deal.amount);
  const forecast = sum(openPeriod, (deal) => deal.forecastAmount);
  const risk = riskTotals(openAll);

  return {
    months,
    monthSet,
    sales,
    saleSet,
    facts,
    actualFacts,
    lostFacts,
    openAll,
    openPeriod,
    target,
    actual,
    lost,
    openPipeline,
    forecast,
    risk,
  };
}

function riskTotals(deals) {
  const totals = {
    overdue: { count: 0, amount: 0 },
    noExpected: { count: 0, amount: 0 },
    stale30: { count: 0, amount: 0 },
    notContacted: { count: 0, amount: 0 },
  };
  for (const deal of deals) {
    for (const key of Object.keys(totals)) {
      if (deal.risk[key]) {
        totals[key].count += 1;
        totals[key].amount += deal.amount;
      }
    }
  }
  return totals;
}

function initFilters() {
  els.monthSelect.innerHTML = dashboardData.months
    .map((month) => `<option value="${month}">${monthLabels[month]} ${month.slice(0, 4)}</option>`)
    .join("");
  els.monthSelect.value = state.month;
  renderGroupOptions();
  renderSaleOptions();

  els.periodMode.addEventListener("change", () => {
    state.periodMode = els.periodMode.value;
    render();
  });
  els.monthSelect.addEventListener("change", () => {
    state.month = els.monthSelect.value;
    render();
  });
  els.groupSelect.addEventListener("change", () => {
    state.group = els.groupSelect.value;
    renderSaleOptions();
    render();
  });
  els.saleSelect.addEventListener("change", () => {
    state.sale = els.saleSelect.value;
    render();
  });
  els.searchInput.addEventListener("input", () => {
    state.search = els.searchInput.value;
    renderTablesOnly();
  });
  els.showUnmapped.addEventListener("change", () => {
    state.showUnmapped = els.showUnmapped.checked;
    renderGroupOptions();
    renderSaleOptions();
    render();
  });
  document.querySelectorAll(".segment button").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".segment button").forEach((item) => item.classList.remove("active"));
      button.classList.add("active");
      state.category = button.dataset.category;
      render();
    });
  });
  document.querySelectorAll(".tabs button").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tabs button").forEach((item) => item.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.remove("active"));
      button.classList.add("active");
      state.tab = button.dataset.tab;
      document.querySelector(`#tab-${state.tab}`).classList.add("active");
      renderTablesOnly();
    });
  });
}

function renderGroupOptions() {
  const groups = Array.from(
    new Set(
      dashboardData.sales
        .filter((sale) => state.showUnmapped || sale.hasPositiveTarget)
        .map((sale) => sale.group),
    ),
  ).sort((a, b) => a.localeCompare(b, "th"));
  if (state.group !== "all" && !groups.includes(state.group)) state.group = "all";
  els.groupSelect.innerHTML = `<option value="all">ทุกทีม</option>${groups
    .map((group) => `<option value="${escapeHtml(group)}">${escapeHtml(group)}</option>`)
    .join("")}`;
  els.groupSelect.value = state.group;
}

function renderSaleOptions() {
  const sales = salesForOptions();
  if (state.sale !== "all" && !sales.some((sale) => sale.key === state.sale)) state.sale = "all";
  els.saleSelect.innerHTML = `<option value="all">ทุกคน</option>${sales
    .map((sale) => `<option value="${escapeHtml(sale.key)}">${escapeHtml(sale.name)}</option>`)
    .join("")}`;
  els.saleSelect.value = state.sale;
}

function renderMeta() {
  const meta = dashboardData.metadata;
  const generated = new Date(meta.generatedAt).toLocaleString("th-TH", {
    dateStyle: "medium",
    timeStyle: "short",
  });
  els.metaBox.innerHTML = `
    <strong>As of ${escapeHtml(meta.asOfDate)}</strong><br>
    Deals: ${dashboardData.quality.dealRows.toLocaleString("th-TH")} rows<br>
    Targets: ${dashboardData.quality.targetRows.toLocaleString("th-TH")} sales<br>
    Built: ${escapeHtml(generated)}
  `;
}

function renderKpis(scope) {
  const gap = scope.target - scope.actual;
  const achievement = scope.target ? scope.actual / scope.target : scope.actual > 0 ? 1 : 0;
  const coverage = scope.target ? (scope.actual + scope.forecast) / scope.target : 0;
  const winRate = scope.actual + scope.lost > 0 ? scope.actual / (scope.actual + scope.lost) : 0;
  const riskAmount = scope.risk.overdue.amount + scope.risk.noExpected.amount;
  const kpis = [
    {
      label: "Target",
      value: compactMoney(scope.target),
      hint: periodLabel(),
      tone: "info",
    },
    {
      label: "Actual Won",
      value: compactMoney(scope.actual),
      hint: `Achievement ${percent(achievement)}`,
      tone: achievement >= 1 ? "good" : achievement >= 0.7 ? "warn" : "danger",
    },
    {
      label: "Gap to Target",
      value: compactMoney(gap),
      hint: gap <= 0 ? "เกินเป้าแล้ว" : "ยอดที่ยังต้องปิด",
      tone: gap <= 0 ? "good" : "warn",
    },
    {
      label: "Weighted Forecast",
      value: compactMoney(scope.forecast),
      hint: `Coverage ${percent(coverage)}`,
      tone: coverage >= 1 ? "good" : "warn",
    },
    {
      label: "Open Pipeline",
      value: compactMoney(scope.openPipeline),
      hint: `${scope.openPeriod.length.toLocaleString("th-TH")} deals ในช่วง forecast`,
      tone: "info",
    },
    {
      label: "Lost Amount",
      value: compactMoney(scope.lost),
      hint: `Win rate by value ${percent(winRate)}`,
      tone: winRate >= 0.6 ? "good" : "danger",
    },
    {
      label: "Pipeline Risk",
      value: compactMoney(riskAmount),
      hint: "Overdue + no expected close",
      tone: riskAmount > 0 ? "danger" : "good",
    },
    {
      label: "No Target Sales",
      value: dashboardData.sales.filter((sale) => !sale.hasPositiveTarget).length.toLocaleString("th-TH"),
      hint: "Target = 0 หรือไม่มีชื่อใน Sale Target",
      tone: dashboardData.unmappedResponsibles.length ? "warn" : "good",
    },
  ];

  els.kpiGrid.innerHTML = kpis
    .map(
      (kpi) => `
        <article class="kpi-card ${kpi.tone}">
          <div class="label">${escapeHtml(kpi.label)}</div>
          <div class="value">${escapeHtml(kpi.value)}</div>
          <div class="hint">${escapeHtml(kpi.hint)}</div>
        </article>
      `,
    )
    .join("");
}

function periodLabel() {
  if (state.periodMode === "year") return "ทั้งปี 2026";
  if (state.periodMode === "month") return `${monthLabels[state.month]} 2026`;
  return `YTD Jan-${monthLabels[state.month]} 2026`;
}

function metricForMonth(month, sales, category = state.category) {
  const saleSet = new Set(sales.map((sale) => sale.key));
  const target = sum(sales, (sale) => targetForSale(sale, [month], category));
  const actual = sum(
    dashboardData.facts.filter(
      (fact) =>
        saleSet.has(fact.saleKey) &&
        fact.monthKey === month &&
        fact.status === "won" &&
        (category === "all" || fact.category === category),
    ),
    (fact) => fact.amount,
  );
  const forecast = sum(
    dashboardData.openDeals.filter(
      (deal) =>
        saleSet.has(deal.saleKey) &&
        deal.expectedMonth === month &&
        (category === "all" || deal.category === category),
    ),
    (deal) => deal.forecastAmount,
  );
  return { month, target, actual, forecast };
}

function renderMonthlyChart(scope) {
  const rows = dashboardData.months.map((month) => metricForMonth(month, scope.sales));
  const maxValue = Math.max(1, ...rows.flatMap((row) => [row.target, row.actual, row.forecast]));
  els.monthlyChart.innerHTML = rows
    .map((row) => {
      const current = row.month === state.month ? `<span class="badge info">selected</span>` : "";
      return `
        <div class="month-row">
          <div class="month-name">${monthLabels[row.month]} ${current}</div>
          <div class="bar-group">
            ${barLine("target", row.target, maxValue)}
            ${barLine("actual", row.actual, maxValue)}
            ${barLine("forecast", row.forecast, maxValue)}
          </div>
          <div class="bar-label">
            ${compactMoney(row.actual)} / ${compactMoney(row.target)}<br>
            <span class="muted">Forecast ${compactMoney(row.forecast)}</span>
          </div>
        </div>
      `;
    })
    .join("");
}

function barLine(className, value, maxValue) {
  const width = Math.max(0, Math.min(100, (value / maxValue) * 100));
  return `
    <div class="bar-track" title="${money(value)}">
      <div class="bar ${className}" style="width:${width}%"></div>
    </div>
  `;
}

function renderTeamChart(scope) {
  const groups = new Map();
  for (const sale of scope.sales) {
    if (!groups.has(sale.group)) groups.set(sale.group, { group: sale.group, sales: [] });
    groups.get(sale.group).sales.push(sale);
  }
  const rows = Array.from(groups.values())
    .map((group) => {
      const target = sum(group.sales, (sale) => targetForSale(sale, scope.months));
      const saleSet = new Set(group.sales.map((sale) => sale.key));
      const actual = sum(
        scope.actualFacts.filter((fact) => saleSet.has(fact.saleKey)),
        (fact) => fact.amount,
      );
      const forecast = sum(
        scope.openPeriod.filter((deal) => saleSet.has(deal.saleKey)),
        (deal) => deal.forecastAmount,
      );
      return {
        group: group.group,
        target,
        actual,
        forecast,
        achievement: target ? actual / target : actual > 0 ? 1 : 0,
      };
    })
    .sort((a, b) => b.achievement - a.achievement || b.actual - a.actual);

  els.teamChart.innerHTML = rows.length
    ? rows
        .map(
          (row) => `
            <div class="rank-row">
              <div class="rank-name">${escapeHtml(row.group)}</div>
              <div class="progress-track">
                <div class="progress-fill" style="width:${Math.min(140, row.achievement * 100)}%"></div>
              </div>
              <div class="rank-value">${percent(row.achievement)}<br><span class="muted">${compactMoney(row.actual)}</span></div>
            </div>
          `,
        )
        .join("")
    : `<div class="empty">ไม่มีข้อมูลใน filter นี้</div>`;
}

function renderCategoryMix(scope) {
  els.categoryMix.innerHTML = ["new", "renew"]
    .map((category) => {
      const target = sum(scope.sales, (sale) => targetForSale(sale, scope.months, category));
      const actual = sum(
        scope.actualFacts.filter((fact) => fact.category === category),
        (fact) => fact.amount,
      );
      const forecast = sum(
        scope.openPeriod.filter((deal) => deal.category === category),
        (deal) => deal.forecastAmount,
      );
      const achievement = target ? actual / target : actual > 0 ? 1 : 0;
      return `
        <div class="mix-card">
          <strong>${category === "new" ? "New Sales" : "Renewal"}</strong>
          <div class="progress-track">
            <div class="progress-fill" style="width:${Math.min(140, achievement * 100)}%"></div>
          </div>
          <p class="note">Actual ${compactMoney(actual)} จาก Target ${compactMoney(target)}</p>
          <span class="badge ${achievement >= 1 ? "good" : "warn"}">${percent(achievement)}</span>
          <span class="badge info">Forecast ${compactMoney(forecast)}</span>
        </div>
      `;
    })
    .join("");
}

function renderStatusMix(scope) {
  const total = Math.max(1, scope.actual + scope.lost + scope.openPipeline);
  els.statusMix.innerHTML = `
    <div class="status-bars" aria-label="Status amount mix">
      <span class="won" style="width:${(scope.actual / total) * 100}%"></span>
      <span class="lost" style="width:${(scope.lost / total) * 100}%"></span>
      <span class="open" style="width:${(scope.openPipeline / total) * 100}%"></span>
    </div>
    <div class="status-legend">
      <div><strong>Won</strong><br>${compactMoney(scope.actual)}</div>
      <div><strong>Lost</strong><br>${compactMoney(scope.lost)}</div>
      <div><strong>Open</strong><br>${compactMoney(scope.openPipeline)}</div>
    </div>
  `;
}

function calcSaleMetric(sale, months) {
  const saleFacts = dashboardData.facts.filter(
    (fact) => fact.saleKey === sale.key && months.includes(fact.monthKey) && categoryMatches(fact.category),
  );
  const actual = sum(
    saleFacts.filter((fact) => fact.status === "won"),
    (fact) => fact.amount,
  );
  const lost = sum(
    saleFacts.filter((fact) => fact.status === "lost"),
    (fact) => fact.amount,
  );
  const openAll = dashboardData.openDeals.filter((deal) => deal.saleKey === sale.key && categoryMatches(deal.category));
  const openPeriod = openAll.filter((deal) => deal.expectedMonth && months.includes(deal.expectedMonth));
  const target = targetForSale(sale, months);
  const forecast = sum(openPeriod, (deal) => deal.forecastAmount);
  const openPipeline = sum(openPeriod, (deal) => deal.amount);
  const risk = riskTotals(openAll);
  return {
    sale,
    target,
    actual,
    lost,
    gap: target - actual,
    forecast,
    openPipeline,
    achievement: target ? actual / target : actual > 0 ? 1 : 0,
    coverage: target ? (actual + forecast) / target : 0,
    winRate: actual + lost > 0 ? actual / (actual + lost) : 0,
    risk,
  };
}

function renderSalesTable(scope) {
  const q = normalizeSearch(state.search);
  const rows = scope.sales
    .map((sale) => calcSaleMetric(sale, scope.months))
    .filter((row) => {
      if (!q) return true;
      const text = `${row.sale.name} ${row.sale.group} ${row.sale.responsibleNames.join(" ")}`.toLowerCase();
      return text.includes(q);
    })
    .sort((a, b) => {
      if (a.sale.hasPositiveTarget !== b.sale.hasPositiveTarget) return a.sale.hasPositiveTarget ? -1 : 1;
      return b.gap - a.gap || b.openPipeline - a.openPipeline;
    });

  if (!rows.length) {
    els.salesTable.innerHTML = `<div class="empty">ไม่มี Sale ที่ตรงกับ filter</div>`;
    return;
  }

  els.salesTable.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Sale</th>
          <th>Team</th>
          <th class="num">Target</th>
          <th class="num">Actual</th>
          <th class="num">Ach.</th>
          <th class="num">Gap</th>
          <th class="num">Forecast</th>
          <th class="num">Coverage</th>
          <th class="num">Lost</th>
          <th class="num">Risk</th>
          <th>Next Action</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (row) => `
              <tr>
                <td><strong>${escapeHtml(row.sale.name)}</strong><br><span class="muted">${escapeHtml(row.sale.responsibleNames.join(", ") || "-")}</span></td>
                <td>${escapeHtml(row.sale.group)}</td>
                <td class="num">${money(row.target)}</td>
                <td class="num">${money(row.actual)}</td>
                <td class="num">${percent(row.achievement)}</td>
                <td class="num">${money(row.gap)}</td>
                <td class="num">${money(row.forecast)}</td>
                <td class="num">${percent(row.coverage)}</td>
                <td class="num">${money(row.lost)}</td>
                <td class="num">${row.risk.overdue.count + row.risk.noExpected.count + row.risk.stale30.count}</td>
                <td>${actionBadge(row)}</td>
              </tr>
            `,
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function actionBadge(row) {
  if (!row.sale.hasPositiveTarget && row.actual + row.openPipeline > 0) return `<span class="badge warn">เพิ่ม Target / Mapping</span>`;
  if (row.target > 0 && row.achievement >= 1) return `<span class="badge good">รักษายอดและปิดเอกสาร</span>`;
  if (row.risk.overdue.amount > 0 || row.risk.noExpected.count > 0) return `<span class="badge danger">เคลียร์ Forecast</span>`;
  if (row.target > 0 && row.coverage < 1) return `<span class="badge warn">เติม Pipeline</span>`;
  if (row.risk.notContacted.count > 0) return `<span class="badge info">Follow-up Not Contacted</span>`;
  return `<span class="badge info">เร่งปิดดีลใกล้สำเร็จ</span>`;
}

function riskScore(deal) {
  return (
    (deal.risk.overdue ? 4 : 0) +
    (deal.risk.noExpected ? 3 : 0) +
    (deal.risk.stale30 ? 2 : 0) +
    (deal.risk.notContacted ? 1 : 0)
  );
}

function dealMatchesSearch(deal) {
  const q = normalizeSearch(state.search);
  if (!q) return true;
  const text = `${deal.saleName} ${deal.responsible} ${deal.company} ${deal.dealName} ${deal.pipeline} ${deal.stage}`.toLowerCase();
  return text.includes(q);
}

function renderTopDealsTable(scope) {
  const rows = scope.openPeriod
    .filter(dealMatchesSearch)
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 30);
  els.topDealsTable.innerHTML = dealTable(rows, "ไม่มี open opportunity ในช่วงที่เลือก");
}

function renderRiskSummary(scope) {
  const cards = [
    ["Overdue", scope.risk.overdue, "Expected close date เลยกำหนด", "danger"],
    ["No Expected Date", scope.risk.noExpected, "ยังไม่มีวันที่คาดว่าจะปิด", "warn"],
    ["Stale > 30 days", scope.risk.stale30, "Stage ไม่ขยับเกิน 30 วัน", "danger"],
    ["Not Contacted", scope.risk.notContacted, "ต้องเริ่ม follow-up", "info"],
  ];
  els.riskSummary.innerHTML = cards
    .map(
      ([title, item, hint, tone]) => `
        <div class="risk-card">
          <strong>${escapeHtml(title)}</strong>
          <div class="value">${compactMoney(item.amount)}</div>
          <p class="note">${item.count.toLocaleString("th-TH")} deals · ${escapeHtml(hint)}</p>
          <span class="badge ${tone}">${escapeHtml(tone === "danger" ? "ต้องจัดการ" : "ติดตาม")}</span>
        </div>
      `,
    )
    .join("");
}

function renderRiskByStage(scope) {
  const riskyDeals = scope.openAll.filter((deal) => riskScore(deal) > 0);
  const stageMap = new Map();
  for (const deal of riskyDeals) {
    if (!stageMap.has(deal.stage)) stageMap.set(deal.stage, { stage: deal.stage, count: 0, amount: 0 });
    const item = stageMap.get(deal.stage);
    item.count += 1;
    item.amount += deal.amount;
  }
  const rows = Array.from(stageMap.values()).sort((a, b) => b.amount - a.amount).slice(0, 12);
  const maxAmount = Math.max(1, ...rows.map((row) => row.amount));
  els.riskByStage.innerHTML = rows.length
    ? rows
        .map(
          (row) => `
            <div class="rank-row">
              <div class="rank-name">${escapeHtml(row.stage)}</div>
              <div class="progress-track">
                <div class="progress-fill" style="width:${(row.amount / maxAmount) * 100}%"></div>
              </div>
              <div class="rank-value">${compactMoney(row.amount)}<br><span class="muted">${row.count} deals</span></div>
            </div>
          `,
        )
        .join("")
    : `<div class="empty">ไม่มี risk deal ใน filter นี้</div>`;
}

function renderRiskDealsTable(scope) {
  const rows = scope.openAll
    .filter((deal) => riskScore(deal) > 0)
    .filter(dealMatchesSearch)
    .sort((a, b) => riskScore(b) - riskScore(a) || b.amount - a.amount)
    .slice(0, 80);
  els.riskDealsTable.innerHTML = dealTable(rows, "ไม่มี pipeline risk ใน filter นี้");
}

function dealTable(rows, emptyText) {
  if (!rows.length) return `<div class="empty">${escapeHtml(emptyText)}</div>`;
  return `
    <table>
      <thead>
        <tr>
          <th>Sale</th>
          <th class="company-cell">Company / Deal</th>
          <th>Stage</th>
          <th>Pipeline</th>
          <th>Product</th>
          <th class="num">Amount</th>
          <th class="num">Forecast</th>
          <th>Expected</th>
          <th>Risk</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (deal) => `
              <tr>
                <td><strong>${escapeHtml(deal.saleName)}</strong><br><span class="muted">${escapeHtml(deal.responsible)}</span></td>
                <td class="company-cell"><strong>${escapeHtml(deal.company || "-")}</strong><br><span class="muted">${escapeHtml(deal.dealName || `ID ${deal.id}`)}</span></td>
                <td>${escapeHtml(deal.stage)}</td>
                <td>${escapeHtml(deal.pipeline)}</td>
                <td>${escapeHtml(deal.product)}</td>
                <td class="num">${money(deal.amount)}</td>
                <td class="num">${money(deal.forecastAmount)}</td>
                <td>${escapeHtml(deal.expectedDate || "-")}<br><span class="muted">${deal.stageAgeDays ?? "-"} stage days</span></td>
                <td>${riskBadges(deal)}</td>
              </tr>
            `,
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function riskBadges(deal) {
  const badges = [];
  if (deal.risk.overdue) badges.push(`<span class="badge danger">Overdue</span>`);
  if (deal.risk.noExpected) badges.push(`<span class="badge warn">No date</span>`);
  if (deal.risk.stale30) badges.push(`<span class="badge danger">Stale</span>`);
  if (deal.risk.notContacted) badges.push(`<span class="badge info">Not contacted</span>`);
  if (!badges.length) badges.push(`<span class="badge good">Clean</span>`);
  return `<div class="badge-list">${badges.join("")}</div>`;
}

function renderDataQuality() {
  els.assumptionList.innerHTML = dashboardData.metadata.assumptions
    .map((item) => `<div class="note-card">${escapeHtml(item)}</div>`)
    .join("");
  const q = dashboardData.quality;
  const cards = [
    ["Pre-WON counted as Won", q.preWonAsWon],
    ["Pre-LOST counted as Lost", q.preLostAsLost],
    ["Won total after normalization", q.won],
    ["Lost total after normalization", q.lost],
    ["Open pipeline after normalization", q.open],
    ["Overdue open", q.overdueOpen],
    ["Sales with zero target", q.zeroTargetSales],
  ];
  els.normalizationStats.innerHTML = cards
    .map(
      ([title, item]) => `
        <div class="risk-card">
          <strong>${escapeHtml(title)}</strong>
          <div class="value">${compactMoney(item.amount)}</div>
          <p class="note">${item.count.toLocaleString("th-TH")} deals</p>
        </div>
      `,
    )
    .join("");

  const rows = dashboardData.unmappedResponsibles;
  els.unmappedTable.innerHTML = rows.length
    ? `
      <table>
        <thead>
          <tr>
            <th>Responsible</th>
            <th class="num">Deals</th>
            <th class="num">Amount</th>
            <th>Suggested Action</th>
          </tr>
        </thead>
        <tbody>
          ${rows
            .map(
              (row) => `
                <tr>
                  <td>${escapeHtml(row.responsible)}</td>
                  <td class="num">${row.count.toLocaleString("th-TH")}</td>
                  <td class="num">${money(row.amount)}</td>
                  <td><span class="badge warn">เพิ่มใน target หรือ alias mapping</span></td>
                </tr>
              `,
            )
            .join("")}
        </tbody>
      </table>
    `
    : `<div class="empty">ทุก Responsible map กับ Sale Target แล้ว</div>`;
}

function renderTablesOnly() {
  const scope = calcScope();
  if (state.tab === "overview") renderTopDealsTable(scope);
  if (state.tab === "sales") renderSalesTable(scope);
  if (state.tab === "pipeline") {
    renderRiskSummary(scope);
    renderRiskByStage(scope);
    renderRiskDealsTable(scope);
  }
  if (state.tab === "data") renderDataQuality();
}

function render() {
  const scope = calcScope();
  renderMeta();
  renderKpis(scope);
  renderMonthlyChart(scope);
  renderTeamChart(scope);
  renderCategoryMix(scope);
  renderStatusMix(scope);
  renderTopDealsTable(scope);
  renderSalesTable(scope);
  renderRiskSummary(scope);
  renderRiskByStage(scope);
  renderRiskDealsTable(scope);
  renderDataQuality();
}

initFilters();
render();
