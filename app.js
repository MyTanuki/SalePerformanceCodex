const STORED_DATA_KEY = "sale-performance-dashboard-uploaded-deals";
let dashboardData = loadStoredDashboardData() || window.DASHBOARD_DATA;

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
  settingsModal: document.querySelector("#settingsModal"),
  closeSettings: document.querySelector("#closeSettings"),
  dealCsvInput: document.querySelector("#dealCsvInput"),
  applyDealCsv: document.querySelector("#applyDealCsv"),
  clearDealCsv: document.querySelector("#clearDealCsv"),
  uploadStatus: document.querySelector("#uploadStatus"),
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
  dealDetailTable: document.querySelector("#dealDetailTable"),
  salesTable: document.querySelector("#salesTable"),
  riskSummary: document.querySelector("#riskSummary"),
  riskByStage: document.querySelector("#riskByStage"),
  riskDealsTable: document.querySelector("#riskDealsTable"),
  assumptionList: document.querySelector("#assumptionList"),
  normalizationStats: document.querySelector("#normalizationStats"),
  unmappedTable: document.querySelector("#unmappedTable"),
};

const baseTargetSales = dashboardData.sales.filter((sale) => sale.hasTarget).map((sale) => clone(sale));

const RESPONSIBLE_ALIASES = {
  "chananthicha wongsan": "Chananthicha",
  chananthicha: "Chananthicha",
  "phongthorn meesaeng": "Phongthorn",
  phongthorn: "Phongthorn",
  "pimpilai puanglamyai": "Pimpilai",
  pimpilai: "Pimpilai",
  "ปณกา องคานนท": "Punika",
  punika: "Punika",
  "เสารแกว สงหลสาย": "Saokeaw",
  saokeaw: "Saokeaw",
  "parinya jinanukulwong": "Parinya",
  parinya: "Parinya",
  "netnapha pangdee": "Netnapha",
  netnapha: "Netnapha",
  "pornpat putthipornsavad": "Pornpat",
  pornpat: "Pornpat",
  "weerachai nilnampetch": "Weerachai",
  weerachai: "Weerachai",
  "soontaree treesup": "Soontaree",
  soontaree: "Soontaree",
  "kullapattra jongsiadison": "Kullapattra",
  kullapattra: "Kullapattra",
  "boonkiat jedsdaariyajit": "Boonkiat",
  boonkiat: "Boonkiat",
  "sutasinee lertpongsuk": "Sutasinee Lertpongsuk",
  "phannipha bothmart": "Phannipha Bothmart",
  "maranee kaewkomtai": "Maranee Kaewkomtai",
  "jongprasert kuvijitrijaru": "Jongprasert Kuvijitrijaru",
  "sunipa samngamya": "Sunipa Samngamya",
  sunipa: "Sunipa Samngamya",
  "chutidet jakrungroj": "Chutidet",
  chutidet: "Chutidet",
  "chaiyawats sukmanee": "Chaiyawat",
  chaiyawat: "Chaiyawat",
  "poowathep ratsamee": "Poowathep",
  poowathep: "Poowathep",
  "panuchanart alex auntaya": "Panuchanart",
  panuchanart: "Panuchanart",
  "thitirat phaenthanee": "Thitirat",
  thitirat: "Thitirat",
  "emorn yuenyong": "Emorn",
  emorn: "Emorn",
};

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function loadStoredDashboardData() {
  try {
    const raw = window.localStorage.getItem(STORED_DATA_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function storeDashboardData(data) {
  try {
    window.localStorage.setItem(STORED_DATA_KEY, JSON.stringify(data));
  } catch {
    setUploadStatus("Upload สำเร็จ แต่ browser ไม่สามารถบันทึกข้อมูลไว้หลัง refresh ได้", "error");
  }
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function cleanText(value) {
  return String(value ?? "").replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
}

function normalizeName(value) {
  return cleanText(value)
    .normalize("NFKD")
    .replace(/\([^)]*\)/g, "")
    .replace(/\p{M}/gu, "")
    .replace(/[^\p{L}\p{N}\s]/gu, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function makeKey(value) {
  return normalizeName(value).replace(/\s+/g, "-") || "unknown";
}

function parseMoneyValue(value) {
  const text = String(value ?? "").trim();
  if (!text) return 0;
  const compact = text.replace(/\s/g, "");
  if (!compact || /^-+$/.test(compact)) return 0;
  const cleaned = compact.replace(/,/g, "").replace(/[฿]/g, "").replace(/[^\d.-]/g, "");
  if (!cleaned || cleaned === "-" || cleaned === "." || cleaned === "-.") return 0;
  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
}

function roundMoney(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;
  const source = String(text || "").replace(/^\uFEFF/, "");

  for (let i = 0; i < source.length; i += 1) {
    const ch = source[i];
    const next = source[i + 1];
    if (inQuotes) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i += 1;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cell += ch;
      }
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell.replace(/\r$/, ""));
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += ch;
    }
  }
  if (cell.length || row.length) {
    row.push(cell.replace(/\r$/, ""));
    rows.push(row);
  }
  if (!rows.length) return [];
  const headers = rows[0].map((header) => header.trim());
  return rows
    .slice(1)
    .filter((values) => values.some((value) => cleanText(value)))
    .map((values) => {
      const record = {};
      headers.forEach((header, index) => {
        record[header] = values[index] ?? "";
      });
      return record;
    });
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
  return makeDateParts(`${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`);
}

function makeDateParts(isoDate) {
  const [year, month, day] = isoDate.split("-").map(Number);
  return {
    date: isoDate,
    year,
    month,
    day,
    monthKey: `${year}-${String(month).padStart(2, "0")}`,
    ts: Date.UTC(year, month - 1, day),
  };
}

function localIsoDate(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
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
  document.addEventListener("click", (event) => {
    const settingsButton = event.target.closest("#openSettings");
    if (settingsButton) openSettings();
  });
  els.closeSettings.addEventListener("click", closeSettings);
  els.settingsModal.addEventListener("click", (event) => {
    if (event.target === els.settingsModal) closeSettings();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !els.settingsModal.hidden) closeSettings();
  });
  els.dealCsvInput.addEventListener("change", () => {
    const file = els.dealCsvInput.files?.[0];
    setUploadStatus(file ? `เลือกไฟล์: ${file.name}` : "ยังไม่ได้เลือกไฟล์");
  });
  els.applyDealCsv.addEventListener("click", handleDealCsvUpload);
  els.clearDealCsv.addEventListener("click", () => {
    els.dealCsvInput.value = "";
    window.localStorage.removeItem(STORED_DATA_KEY);
    dashboardData = window.DASHBOARD_DATA;
    state.group = "all";
    state.sale = "all";
    renderGroupOptions();
    renderSaleOptions();
    render();
    setUploadStatus("Reset กลับไปใช้ข้อมูลตั้งต้นแล้ว", "success");
  });
}

function openSettings() {
  els.settingsModal.hidden = false;
  els.dealCsvInput.focus();
}

function closeSettings() {
  els.settingsModal.hidden = true;
}

function setUploadStatus(message, tone = "") {
  els.uploadStatus.textContent = message;
  els.uploadStatus.className = `upload-status ${tone}`.trim();
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
    <button type="button" class="settings-button" id="openSettings" aria-label="Open settings">
      <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M12 15.5A3.5 3.5 0 1 0 12 8a3.5 3.5 0 0 0 0 7.5Z"></path>
        <path d="M19.4 15a1.8 1.8 0 0 0 .36 1.98l.05.05a2.1 2.1 0 1 1-2.97 2.97l-.05-.05a1.8 1.8 0 0 0-1.98-.36 1.8 1.8 0 0 0-1.09 1.65V21.4a2.1 2.1 0 1 1-4.2 0v-.08A1.8 1.8 0 0 0 8.43 19.7a1.8 1.8 0 0 0-1.98.36l-.05.05a2.1 2.1 0 1 1-2.97-2.97l.05-.05a1.8 1.8 0 0 0 .36-1.98 1.8 1.8 0 0 0-1.65-1.09H2.1a2.1 2.1 0 1 1 0-4.2h.08A1.8 1.8 0 0 0 3.8 8.73a1.8 1.8 0 0 0-.36-1.98l-.05-.05A2.1 2.1 0 1 1 6.36 3.73l.05.05a1.8 1.8 0 0 0 1.98.36A1.8 1.8 0 0 0 9.48 2.5V2.1a2.1 2.1 0 1 1 4.2 0v.08a1.8 1.8 0 0 0 1.09 1.65 1.8 1.8 0 0 0 1.98-.36l.05-.05a2.1 2.1 0 1 1 2.97 2.97l-.05.05a1.8 1.8 0 0 0-.36 1.98 1.8 1.8 0 0 0 1.65 1.09h.39a2.1 2.1 0 1 1 0 4.2h-.08A1.8 1.8 0 0 0 19.4 15Z"></path>
      </svg>
    </button>
    <strong>As of ${escapeHtml(meta.asOfDate)}</strong><br>
    Deals: ${dashboardData.quality.dealRows.toLocaleString("th-TH")} rows<br>
    Targets: ${dashboardData.quality.targetRows.toLocaleString("th-TH")} sales<br>
    Built: ${escapeHtml(generated)}
  `;
}

async function handleDealCsvUpload() {
  const file = els.dealCsvInput.files?.[0];
  if (!file) {
    setUploadStatus("กรุณาเลือกไฟล์ .csv ก่อน", "error");
    return;
  }
  if (!file.name.toLowerCase().endsWith(".csv")) {
    setUploadStatus("รองรับเฉพาะไฟล์ .csv เท่านั้น", "error");
    return;
  }

  try {
    setUploadStatus("กำลังอ่านไฟล์...");
    const text = await file.text();
    const rows = parseCsv(text);
    validateDealRows(rows);
    dashboardData = buildDashboardDataFromDeals(rows, file.name);
    storeDashboardData(dashboardData);
    state.group = "all";
    state.sale = "all";
    state.search = "";
    els.searchInput.value = "";
    renderGroupOptions();
    renderSaleOptions();
    render();
    setUploadStatus(`Upload สำเร็จ: ${file.name} (${rows.length.toLocaleString("th-TH")} deals)`, "success");
  } catch (error) {
    setUploadStatus(error.message || "ไม่สามารถอ่านไฟล์นี้ได้", "error");
  }
}

function validateDealRows(rows) {
  if (!rows.length) throw new Error("ไฟล์นี้ไม่มีข้อมูล deal");
  const required = ["ID", "Pipeline", "Stage", "Responsible", "Income", "Created", "Deal Type", "Product Type", "Company", "Expected close date"];
  const missing = required.filter((column) => !(column in rows[0]));
  if (missing.length) {
    throw new Error(`รูปแบบ CSV ไม่ตรงกับไฟล์ DEAL ตั้งต้น: ขาดคอลัมน์ ${missing.join(", ")}`);
  }
}

function buildDashboardDataFromDeals(dealRows, fileName) {
  const asOfDate = localIsoDate();
  const today = makeDateParts(asOfDate);
  const dayMs = 24 * 60 * 60 * 1000;
  const targetSales = baseTargetSales.map((sale) => ({
    ...clone(sale),
    responsibleNames: [],
  }));
  const saleByKey = new Map(targetSales.map((sale) => [sale.key, sale]));
  const targetNameToKey = new Map(targetSales.map((sale) => [normalizeName(sale.name), sale.key]));
  const factsMap = new Map();
  const openDeals = [];
  const dealDetails = [];
  const unmappedResponsibles = new Map();
  const quality = {
    dealRows: dealRows.length,
    targetRows: targetSales.length,
    preWonAsWon: { count: 0, amount: 0 },
    preLostAsLost: { count: 0, amount: 0 },
    won: { count: 0, amount: 0 },
    lost: { count: 0, amount: 0 },
    open: { count: 0, amount: 0 },
    overdueOpen: { count: 0, amount: 0 },
    noExpectedOpen: { count: 0, amount: 0 },
    stale30Open: { count: 0, amount: 0 },
    notContactedOpen: { count: 0, amount: 0 },
  };

  const mapResponsible = (responsible) => {
    const normalized = normalizeName(responsible);
    if (!normalized) return null;
    const alias = RESPONSIBLE_ALIASES[normalized];
    if (alias) return targetNameToKey.get(normalizeName(alias)) || makeKey(alias);
    for (const sale of targetSales) {
      const targetNormalized = normalizeName(sale.name);
      if (normalized === targetNormalized || normalized.startsWith(`${targetNormalized} `)) return sale.key;
    }
    return `unmapped:${makeKey(responsible)}`;
  };

  for (const row of dealRows) {
    const responsible = cleanText(row.Responsible) || "Unknown";
    const saleKey = mapResponsible(responsible);
    const amount = parseMoneyValue(row.Income);
    const rawStage = cleanText(row.Stage);
    const status = normalizeDealStage(rawStage);
    const category = inferDealCategory(row);
    const expected = parseDateValue(row["Expected close date"]);
    const stageChanged = parseDateValue(row["Stage change date"]);
    const created = parseDateValue(row.Created);
    const closeDate = stageChanged || expected || created;

    if (!saleByKey.has(saleKey)) {
      saleByKey.set(saleKey, {
        key: saleKey,
        name: responsible,
        group: "Unmapped",
        hasTarget: false,
        hasPositiveTarget: false,
        responsibleNames: [],
        monthly: Object.fromEntries(dashboardData.months.map((month) => [month, { renew: 0, new: 0, total: 0 }])),
        targetAnnual: { renew: 0, new: 0, total: 0 },
      });
    }

    const sale = saleByKey.get(saleKey);
    if (!sale.responsibleNames.includes(responsible)) sale.responsibleNames.push(responsible);

    if (sale.group === "Unmapped") {
      if (!unmappedResponsibles.has(responsible)) unmappedResponsibles.set(responsible, { responsible, count: 0, amount: 0 });
      const item = unmappedResponsibles.get(responsible);
      item.count += 1;
      item.amount += amount;
    }

    if (rawStage === "Pre-WON") addQuality(quality.preWonAsWon, amount);
    if (rawStage === "Pre-LOST") addQuality(quality.preLostAsLost, amount);

    const detail = {
      id: cleanText(row.ID),
      saleKey,
      saleName: sale.name,
      group: sale.group,
      responsible,
      category,
      status,
      stage: rawStage,
      pipeline: cleanText(row.Pipeline),
      product: cleanText(row["Product Type"]) || "(blank)",
      company: cleanText(row.Company),
      contact: cleanText(row.Contact),
      dealName: cleanText(row["Deal Name"]),
      amount: roundMoney(amount),
      forecastAmount: roundMoney(status === "open" ? amount * (dashboardData.stageWeights[rawStage] ?? 0.1) : 0),
      expectedDate: expected?.date || "",
      expectedMonth: expected?.monthKey || "",
      stageChangeDate: stageChanged?.date || "",
      stageMonth: stageChanged?.monthKey || "",
      createdDate: created?.date || "",
      trackingMonth: status === "open" ? expected?.monthKey || "" : closeDate?.monthKey || "",
      stageAgeDays: stageChanged ? Math.floor((today.ts - stageChanged.ts) / dayMs) : null,
      risk: {
        overdue: Boolean(expected && expected.ts < today.ts),
        noExpected: !expected,
        stale30: Boolean(stageChanged && stageChanged.ts < today.ts - 30 * dayMs),
        notContacted: rawStage === "Not Contacted",
      },
    };
    dealDetails.push(detail);

    if (status === "won" || status === "lost") {
      addQuality(status === "won" ? quality.won : quality.lost, amount);
      const monthKey = closeDate?.monthKey || "";
      if (monthKey.startsWith("2026-")) {
        addFact(factsMap, {
          saleKey,
          monthKey,
          category,
          status,
          group: sale.group,
          saleName: sale.name,
          amount,
        });
      }
      continue;
    }

    addQuality(quality.open, amount);
    const noExpected = !expected;
    const overdue = Boolean(expected && expected.ts < today.ts);
    const stale30 = Boolean(stageChanged && stageChanged.ts < today.ts - 30 * dayMs);
    const notContacted = rawStage === "Not Contacted";
    if (noExpected) addQuality(quality.noExpectedOpen, amount);
    if (overdue) addQuality(quality.overdueOpen, amount);
    if (stale30) addQuality(quality.stale30Open, amount);
    if (notContacted) addQuality(quality.notContactedOpen, amount);

    const weight = dashboardData.stageWeights[rawStage] ?? 0.1;
    openDeals.push({
      id: cleanText(row.ID),
      saleKey,
      saleName: sale.name,
      group: sale.group,
      responsible,
      category,
      stage: rawStage,
      pipeline: cleanText(row.Pipeline),
      product: cleanText(row["Product Type"]) || "(blank)",
      company: cleanText(row.Company),
      dealName: cleanText(row["Deal Name"]),
      amount: roundMoney(amount),
      forecastAmount: roundMoney(amount * weight),
      weight,
      expectedDate: expected?.date || "",
      expectedMonth: expected?.monthKey || "",
      stageChangeDate: stageChanged?.date || "",
      stageAgeDays: stageChanged ? Math.floor((today.ts - stageChanged.ts) / dayMs) : null,
      risk: { overdue, noExpected, stale30, notContacted },
    });
  }

  for (const bucket of Object.values(quality)) {
    if (bucket && typeof bucket === "object" && "amount" in bucket) bucket.amount = roundMoney(bucket.amount);
  }

  const sales = Array.from(saleByKey.values())
    .map((sale) => ({
      ...sale,
      responsibleNames: sale.responsibleNames.sort((a, b) => a.localeCompare(b, "th")),
    }))
    .sort((a, b) => {
      if (a.group === "Unmapped" && b.group !== "Unmapped") return 1;
      if (b.group === "Unmapped" && a.group !== "Unmapped") return -1;
      return a.group.localeCompare(b.group, "th") || a.name.localeCompare(b.name, "th");
    });

  return {
    ...dashboardData,
    metadata: {
      ...dashboardData.metadata,
      generatedAt: new Date().toISOString(),
      asOfDate,
      sourceFiles: {
        ...dashboardData.metadata.sourceFiles,
        deals: fileName,
      },
    },
    sales,
    facts: Array.from(factsMap.values()).map((fact) => ({ ...fact, amount: roundMoney(fact.amount) })),
    openDeals,
    dealDetails,
    quality: {
      ...quality,
      zeroTargetSales: {
        count: targetSales.filter((sale) => !sale.hasPositiveTarget).length,
        amount: 0,
      },
    },
    unmappedResponsibles: Array.from(unmappedResponsibles.values())
      .map((item) => ({ ...item, amount: roundMoney(item.amount) }))
      .sort((a, b) => b.amount - a.amount),
  };
}

function addQuality(bucket, amount) {
  bucket.count += 1;
  bucket.amount += amount;
}

function addFact(factsMap, fact) {
  const key = [fact.saleKey, fact.monthKey, fact.category, fact.status].join("|");
  if (!factsMap.has(key)) factsMap.set(key, { ...fact, amount: 0, count: 0 });
  const item = factsMap.get(key);
  item.amount += fact.amount;
  item.count += 1;
}

function normalizeDealStage(stage) {
  const lower = cleanText(stage).toLowerCase();
  if (lower === "deal won" || lower === "pre-won") return "won";
  if (lower === "deal lost" || lower === "pre-lost") return "lost";
  return "open";
}

function inferDealCategory(row) {
  const pipeline = cleanText(row.Pipeline).toLowerCase();
  const dealType = cleanText(row["Deal Type"]).toLowerCase();
  if (pipeline.includes("renew") || dealType.startsWith("re-new")) return "renew";
  return "new";
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

function allDealDetails() {
  if (Array.isArray(dashboardData.dealDetails)) return dashboardData.dealDetails;
  return dashboardData.openDeals.map((deal) => ({
    ...deal,
    status: "open",
    trackingMonth: deal.expectedMonth,
  }));
}

function renderDealDetailTable(scope) {
  const q = normalizeSearch(state.search);
  const rows = allDealDetails()
    .filter((deal) => scope.saleSet.has(deal.saleKey))
    .filter((deal) => categoryMatches(deal.category))
    .filter((deal) => {
      if (state.periodMode === "year") return deal.trackingMonth?.startsWith("2026-");
      return deal.trackingMonth && scope.monthSet.has(deal.trackingMonth);
    })
    .filter((deal) => {
      if (!q) return true;
      const text = `${deal.saleName} ${deal.responsible} ${deal.company} ${deal.dealName} ${deal.pipeline} ${deal.stage} ${deal.id}`.toLowerCase();
      return text.includes(q);
    })
    .sort((a, b) => {
      const statusOrder = { open: 1, won: 2, lost: 3 };
      return (
        (statusOrder[a.status] || 9) - (statusOrder[b.status] || 9) ||
        (b.amount || 0) - (a.amount || 0)
      );
    })
    .slice(0, 300);

  els.dealDetailTable.innerHTML = dealDetailTable(rows, "ไม่มี deal detail ในเดือนที่เลือก");
}

function dealDetailTable(rows, emptyText) {
  if (!rows.length) return `<div class="empty">${escapeHtml(emptyText)}</div>`;
  return `
    <table>
      <thead>
        <tr>
          <th>ID</th>
          <th>Sale</th>
          <th class="company-cell">Company / Deal</th>
          <th>Status</th>
          <th>Stage</th>
          <th>Pipeline</th>
          <th>Product</th>
          <th class="num">Amount</th>
          <th class="num">Forecast</th>
          <th>Expected</th>
          <th>Stage Change</th>
          <th>Risk / Progress</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (deal) => `
              <tr>
                <td>${escapeHtml(deal.id || "-")}</td>
                <td><strong>${escapeHtml(deal.saleName)}</strong><br><span class="muted">${escapeHtml(deal.responsible)}</span></td>
                <td class="company-cell"><strong>${escapeHtml(deal.company || "-")}</strong><br><span class="muted">${escapeHtml(deal.dealName || "-")}</span></td>
                <td>${statusBadge(deal.status)}</td>
                <td>${escapeHtml(deal.stage)}</td>
                <td>${escapeHtml(deal.pipeline)}</td>
                <td>${escapeHtml(deal.product)}</td>
                <td class="num">${money(deal.amount)}</td>
                <td class="num">${money(deal.forecastAmount || 0)}</td>
                <td>${escapeHtml(deal.expectedDate || "-")}</td>
                <td>${escapeHtml(deal.stageChangeDate || "-")}<br><span class="muted">${deal.stageAgeDays ?? "-"} stage days</span></td>
                <td>${deal.status === "open" ? riskBadges(deal) : progressBadge(deal.status)}</td>
              </tr>
            `,
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function statusBadge(status) {
  if (status === "won") return `<span class="badge good">Won</span>`;
  if (status === "lost") return `<span class="badge danger">Lost</span>`;
  return `<span class="badge info">Open</span>`;
}

function progressBadge(status) {
  if (status === "won") return `<span class="badge good">สำเร็จแล้ว</span>`;
  if (status === "lost") return `<span class="badge danger">ปิด Lost แล้ว</span>`;
  return `<span class="badge info">ติดตามต่อ</span>`;
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
  if (state.tab === "deals") renderDealDetailTable(scope);
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
  renderDealDetailTable(scope);
  renderSalesTable(scope);
  renderRiskSummary(scope);
  renderRiskByStage(scope);
  renderRiskDealsTable(scope);
  renderDataQuality();
}

initFilters();
render();
