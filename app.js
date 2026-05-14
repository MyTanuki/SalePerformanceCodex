const STORED_DATA_KEY = "sale-performance-dashboard-workspace-v2";
const initialDashboardData = JSON.parse(JSON.stringify(window.DASHBOARD_DATA || createEmptyDashboardData()));
const storedDashboardData = loadStoredDashboardData();
let dashboardData = storedDashboardData || JSON.parse(JSON.stringify(initialDashboardData));

const state = {
  viewMode: "group",
  group: "all",
  sale: "all",
  periodType: "year",
  periodYear: "2026",
  periodQuarter: "2026-Q2",
  periodWeek: "2026-W20",
  startMonth: "2026-01",
  endMonth: "2026-12",
  leadStartWeek: "2026-W23",
  leadEndWeek: "2026-W26",
  leadSearch: "",
  category: "all",
  showUnmapped: false,
  search: "",
  tab: "overview",
  statusFilters: {
    won: true,
    commit: true,
    upside: true,
    open: true,
    lost: true,
  },
  transactionSort: {
    key: "amount",
    dir: "desc",
  },
  dealDetailSort: {
    key: "status",
    dir: "asc",
  },
};

const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const saleTypes = ["renew", "new"];
const defaultPipelineOrder = [
  "2027 Budget Setup (งบปี 70)",
  "2028 Budget Setup (งบปี 71)",
  "Auto Renew",
  "By Chance Project",
  "Connectivity",
  "Corporate",
  "Government Deal Competition",
  "Government Deal Focus",
  "Government Deal Steal",
  "Inside NEW",
  "Partner",
  "Special Project",
  "Subscription Renew",
  "Work Request Presale BIZ",
  "Work Request Presale SI",
];

function createEmptyDashboardData() {
  const months = Array.from({ length: 5 * 12 }, (_, index) => {
    const year = 2024 + Math.floor(index / 12);
    const month = (index % 12) + 1;
    return `${year}-${String(month).padStart(2, "0")}`;
  });
  return {
    metadata: {
      generatedAt: new Date().toISOString(),
      asOfDate: localIsoDate(),
      sourceFiles: {
        deals: "",
        target: "",
        mapping: "",
      },
      assumptions: [
        "Web starts empty. Upload Sale Target and Deal CSV to build the dashboard.",
        "Stage Mapping, Sale Target, and Pipeline Matching rules can be auto-loaded from ./data.",
      ],
    },
    months,
    stageWeights: {
      Commit: 0.8,
      Upside: 0.5,
      Open: 0.1,
    },
    mappings: {
      pipelineGroupMap: {},
      pipelineRuleMap: {},
      stageMap: {},
      dealTypeMap: {},
      salesGroupMap: {},
      counts: { pipelines: 0, pipelineRules: 0, stages: 0, dealTypes: 0, sales: 0 },
    },
    sales: [],
    facts: [],
    openDeals: [],
    dealDetails: [],
    quality: emptyQuality(),
    unmappedResponsibles: [],
  };
}

function emptyQuality() {
  return {
    dealRows: 0,
    targetRows: 0,
    preWonAsWon: { count: 0, amount: 0 },
    preLostAsLost: { count: 0, amount: 0 },
    won: { count: 0, amount: 0 },
    lost: { count: 0, amount: 0 },
    open: { count: 0, amount: 0 },
    overdueOpen: { count: 0, amount: 0 },
    noExpectedOpen: { count: 0, amount: 0 },
    stale30Open: { count: 0, amount: 0 },
    notContactedOpen: { count: 0, amount: 0 },
    pipelineGroupMismatch: { count: 0, amount: 0 },
    pipelineIncluded: { count: 0, amount: 0 },
    pipelineExcluded: { count: 0, amount: 0 },
    pipelineUnmapped: { count: 0, amount: 0 },
    zeroTargetSales: { count: 0, amount: 0 },
  };
}

const els = {
  metaBox: document.querySelector("#metaBox"),
  settingsModal: document.querySelector("#settingsModal"),
  closeSettings: document.querySelector("#closeSettings"),
  dealModal: document.querySelector("#dealModal"),
  closeDealModal: document.querySelector("#closeDealModal"),
  dealModalTitle: document.querySelector("#dealModalTitle"),
  dealModalSubtitle: document.querySelector("#dealModalSubtitle"),
  dealModalContent: document.querySelector("#dealModalContent"),
  targetCsvInput: document.querySelector("#targetCsvInput"),
  dealCsvInput: document.querySelector("#dealCsvInput"),
  mappingCsvInput: document.querySelector("#mappingCsvInput"),
  targetAutoFileName: document.querySelector("#targetAutoFileName"),
  mappingAutoFileName: document.querySelector("#mappingAutoFileName"),
  applyCsvFiles: document.querySelector("#applyCsvFiles"),
  clearCsvFiles: document.querySelector("#clearCsvFiles"),
  uploadStatus: document.querySelector("#uploadStatus"),
  emptyDataBanner: document.querySelector("#emptyDataBanner"),
  savePipelineMatching: document.querySelector("#savePipelineMatching"),
  clearPipelineMatching: document.querySelector("#clearPipelineMatching"),
  pipelineMatchingSummary: document.querySelector("#pipelineMatchingSummary"),
  pipelineMatchingMatrix: document.querySelector("#pipelineMatchingMatrix"),
  viewMode: document.querySelector("#viewMode"),
  groupFilterContainer: document.querySelector("#groupFilterContainer"),
  salesFilterContainer: document.querySelector("#salesFilterContainer"),
  periodType: document.querySelector("#periodType"),
  periodYear: document.querySelector("#periodYear"),
  periodQuarter: document.querySelector("#periodQuarter"),
  periodWeek: document.querySelector("#periodWeek"),
  startMonth: document.querySelector("#startMonth"),
  endMonth: document.querySelector("#endMonth"),
  leadStartWeek: document.querySelector("#leadStartWeek"),
  leadEndWeek: document.querySelector("#leadEndWeek"),
  leadSearchInput: document.querySelector("#leadSearchInput"),
  periodSubYear: document.querySelector("#periodSubYear"),
  periodSubQuarter: document.querySelector("#periodSubQuarter"),
  periodSubWeek: document.querySelector("#periodSubWeek"),
  periodSubRange: document.querySelector("#periodSubRange"),
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
  newLeadSummary: document.querySelector("#newLeadSummary"),
  newLeadWeeklyChart: document.querySelector("#newLeadWeeklyChart"),
  newLeadTable: document.querySelector("#newLeadTable"),
  riskSummary: document.querySelector("#riskSummary"),
  riskByStage: document.querySelector("#riskByStage"),
  riskDealsTable: document.querySelector("#riskDealsTable"),
  assumptionList: document.querySelector("#assumptionList"),
  normalizationStats: document.querySelector("#normalizationStats"),
  unmappedTable: document.querySelector("#unmappedTable"),
  salesStatusFilters: document.querySelector("#salesStatusFilters"),
  totalSalesChart: document.querySelector("#totalSalesChart"),
  renewSalesChart: document.querySelector("#renewSalesChart"),
  newSalesChart: document.querySelector("#newSalesChart"),
  cumulativeSalesChart: document.querySelector("#cumulativeSalesChart"),
  transactionAndInputs: Array.from(document.querySelectorAll("[data-transaction-and]")),
  transactionNotInputs: Array.from(document.querySelectorAll("[data-transaction-not]")),
  transactionTable: document.querySelector("#transactionTable"),
  transactionTotalAmount: document.querySelector("#transactionTotalAmount"),
};

function currentTargetSales() {
  return dashboardData.sales.filter((sale) => sale.hasTarget).map((sale) => clone(sale));
}

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

const DEFAULT_MAPPINGS = createDefaultMappings();

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
    const parsed = raw ? JSON.parse(raw) : null;
    if (!parsed) return null;
    if (!Array.isArray(parsed.months) || parsed.months.length < 60) return null;
    if (!parsed.mappings) return null;
    return parsed;
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

function hasDealData(data = dashboardData) {
  return (data?.quality?.dealRows || 0) > 0 || (data?.dealDetails || []).length > 0;
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
  const rows = parseCsvRows(text);
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

function parseCsvRows(text) {
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
  return rows;
}

function createDefaultMappings() {
  const stageMap = new Map([
    ["deal won", "Won"],
    ["pre-won", "Won"],
    ["pre-won", "Won"],
    ["deal lost", "Lost"],
    ["pre-lost", "Lost"],
    ["pre-lost", "Lost"],
    ["commit", "Commit"],
    ["upside", "Upside"],
  ]);
  const dealTypeMap = new Map([
    ["cross sell", "New"],
    ["new sell", "New"],
    ["up sell", "New"],
    ["re-new decrease", "Renew"],
    ["re-new macd", "Renew"],
    ["re-new same", "Renew"],
    ["re-new up sell", "Renew"],
  ]);
  const salesGroupMap = new Map();
  Object.entries(RESPONSIBLE_ALIASES).forEach(([source, alias]) => {
    salesGroupMap.set(source, { displayName: alias, group: "" });
  });
  return {
    pipelineGroupMap: new Map(),
    pipelineRuleMap: new Map(),
    stageMap,
    dealTypeMap,
    salesGroupMap,
    counts: { pipelines: 0, pipelineRules: 0, stages: stageMap.size, dealTypes: dealTypeMap.size, sales: salesGroupMap.size },
  };
}

function activeMappings() {
  return dashboardData.mappings ? hydrateMappings(dashboardData.mappings) : DEFAULT_MAPPINGS;
}

function hydrateMappings(mappingData) {
  if (!mappingData || mappingData.pipelineGroupMap instanceof Map) return mappingData || DEFAULT_MAPPINGS;
  return {
    pipelineGroupMap: new Map(Object.entries(mappingData.pipelineGroupMap || {})),
    pipelineRuleMap: new Map(Object.entries(mappingData.pipelineRuleMap || {}).map(([key, values]) => [key, new Set(values || [])])),
    stageMap: new Map(Object.entries(mappingData.stageMap || {})),
    dealTypeMap: new Map(Object.entries(mappingData.dealTypeMap || {})),
    salesGroupMap: new Map(Object.entries(mappingData.salesGroupMap || {})),
    counts: mappingData.counts || { pipelines: 0, pipelineRules: 0, stages: 0, dealTypes: 0, sales: 0 },
  };
}

function serializeMappings(mappings) {
  const mapToObject = (map) => Object.fromEntries(map instanceof Map ? map.entries() : Object.entries(map || {}));
  return {
    pipelineGroupMap: mapToObject(mappings.pipelineGroupMap),
    pipelineRuleMap: Object.fromEntries(
      Array.from(mappings.pipelineRuleMap instanceof Map ? mappings.pipelineRuleMap.entries() : []).map(([key, values]) => [
        key,
        Array.from(values || []),
      ]),
    ),
    stageMap: mapToObject(mappings.stageMap),
    dealTypeMap: mapToObject(mappings.dealTypeMap),
    salesGroupMap: mapToObject(mappings.salesGroupMap),
    counts: mappings.counts,
  };
}

function mappingCounts(mappings) {
  return {
    pipelines: mappings.pipelineGroupMap?.size || 0,
    pipelineRules: Array.from(mappings.pipelineRuleMap?.values?.() || []).reduce((total, values) => total + values.size, 0),
    stages: mappings.stageMap?.size || 0,
    dealTypes: mappings.dealTypeMap?.size || 0,
    sales: mappings.salesGroupMap?.size || 0,
  };
}

function mergeMappings(baseMappings, updateMappings) {
  const merged = {
    pipelineGroupMap: new Map(baseMappings.pipelineGroupMap),
    pipelineRuleMap: new Map(Array.from(baseMappings.pipelineRuleMap || []).map(([key, values]) => [key, new Set(values)])),
    stageMap: new Map(baseMappings.stageMap),
    dealTypeMap: new Map(baseMappings.dealTypeMap),
    salesGroupMap: new Map(baseMappings.salesGroupMap),
  };
  updateMappings.pipelineGroupMap.forEach((group, pipeline) => merged.pipelineGroupMap.set(pipeline, group));
  updateMappings.pipelineRuleMap?.forEach((pipelines, key) => {
    if (!merged.pipelineRuleMap.has(key)) merged.pipelineRuleMap.set(key, new Set());
    pipelines.forEach((pipeline) => merged.pipelineRuleMap.get(key).add(pipeline));
  });
  updateMappings.stageMap.forEach((stage, source) => merged.stageMap.set(source, stage));
  updateMappings.dealTypeMap.forEach((dealType, source) => merged.dealTypeMap.set(source, dealType));
  updateMappings.salesGroupMap.forEach((sale, source) => merged.salesGroupMap.set(source, sale));
  merged.counts = mappingCounts(merged);
  return merged;
}

function parseMappingCsv(text) {
  const rows = parseCsvRows(text).map((row) => row.map(cleanText));
  const pipelineGroupMap = new Map();
  const pipelineRuleMap = new Map();
  const stageMap = new Map();
  const dealTypeMap = new Map();
  const salesGroupMap = new Map();
  let section = "";
  let pipelineHeaders = [];

  for (const row of rows) {
    if (!row.some(Boolean)) continue;
    const first = row[0].toLowerCase();
    const second = row[1]?.toLowerCase() || "";
    if ((first.includes("sale group") && first.includes("pipeline")) || (first === "sale group" && second === "sale type")) {
      section = second === "sale type" ? "pipelineRules" : "pipeline";
      pipelineHeaders = row.slice(second === "sale type" ? 2 : 1);
      continue;
    }
    if (first === "stage (source)" || (first === "stage" && second === "stage")) {
      section = "stage";
      continue;
    }
    if (first === "deal type" && second === "deal type") {
      section = "dealType";
      continue;
    }
    if (first === "sales" && second === "sale group") {
      section = "sales";
      continue;
    }

    if (section === "pipeline") {
      const group = row[0];
      if (!group) continue;
      pipelineHeaders.forEach((pipeline, index) => {
        if (pipeline && isMappingSelected(row[index + 1])) pipelineGroupMap.set(normalizeName(pipeline), group);
      });
    } else if (section === "pipelineRules") {
      const group = row[0];
      const saleType = normalizeSaleType(row[1]);
      if (!group || !saleType) continue;
      const key = pipelineRuleKey(group, saleType);
      if (!pipelineRuleMap.has(key)) pipelineRuleMap.set(key, new Set());
      pipelineHeaders.forEach((pipeline, index) => {
        if (pipeline && isMappingSelected(row[index + 2])) pipelineRuleMap.get(key).add(normalizeName(pipeline));
      });
    } else if (section === "stage") {
      if (row[0] && row[1]) stageMap.set(normalizeName(row[0]), row[1]);
    } else if (section === "dealType") {
      if (row[0] && row[1]) dealTypeMap.set(normalizeName(row[0]), row[1]);
    } else if (section === "sales") {
      if (row[0] && row[1]) {
        const displayName = cleanText(row[0]).replace(/\s+/g, " ");
        salesGroupMap.set(normalizeName(row[0]), { displayName, group: row[1] });
      }
    }
  }

  return {
    pipelineGroupMap,
    pipelineRuleMap,
    stageMap,
    dealTypeMap,
    salesGroupMap,
    counts: {
      pipelines: pipelineGroupMap.size,
      pipelineRules: Array.from(pipelineRuleMap.values()).reduce((total, values) => total + values.size, 0),
      stages: stageMap.size,
      dealTypes: dealTypeMap.size,
      sales: salesGroupMap.size,
    },
  };
}

function validateMappings(mappings) {
  const counts = mappingCounts(mappings);
  const total = counts.pipelines + counts.pipelineRules + counts.stages + counts.dealTypes + counts.sales;
  if (!total) {
    throw new Error("ไฟล์ Mapping CSV ต้องมีอย่างน้อย 1 ตาราง: Pipeline Rule, Pipeline, Stage, Deal Type หรือ Sales");
  }
}

function isMappingSelected(value) {
  const text = cleanText(value).toLowerCase();
  return ["x", "button down", "buttondown", "selected", "yes", "true", "1"].includes(text);
}

function normalizeSaleType(value) {
  const text = cleanText(value).toLowerCase();
  if (text === "new") return "new";
  if (text === "renew" || text === "re-new" || text === "renewal") return "renew";
  return "";
}

function pipelineRuleKey(group, saleType) {
  return `${normalizeName(group)}|${normalizeSaleType(saleType) || cleanText(saleType).toLowerCase()}`;
}

function pipelineCountingResult(group, category, pipeline, mappings = activeMappings()) {
  const rules = mappings.pipelineRuleMap || new Map();
  if (!rules.size) return { status: "included", included: true, label: "Included" };
  const key = pipelineRuleKey(group, category);
  const allowedPipelines = rules.get(key);
  if (!allowedPipelines) return { status: "unmapped", included: false, label: "Unmapped rule" };
  if (allowedPipelines.has(normalizeName(pipeline))) return { status: "included", included: true, label: "Included" };
  return { status: "excluded", included: false, label: "Excluded pipeline" };
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

function monthLabel(month) {
  if (isIsoWeekKey(month)) return weekLabel(month);
  const index = Number(month.slice(5, 7)) - 1;
  return `${monthNames[index] || month} ${month.slice(0, 4)}`;
}

function displayDate(dateText) {
  const parts = parseDateValue(dateText);
  if (!parts) return cleanText(dateText) || "-";
  return parts.date;
}

function periodMonths() {
  const months = dashboardData.months || [];
  if (state.periodType === "week") return [state.periodWeek];
  if (state.periodType === "year") {
    return months.filter((month) => month.startsWith(`${state.periodYear}-`));
  }
  if (state.periodType === "quarter") {
    const [year, quarterToken] = state.periodQuarter.split("-Q");
    const quarter = Number(quarterToken);
    const startMonth = (quarter - 1) * 3 + 1;
    return months.filter((month) => {
      if (!month.startsWith(`${year}-`)) return false;
      const monthNumber = Number(month.slice(5, 7));
      return monthNumber >= startMonth && monthNumber < startMonth + 3;
    });
  }
  return months.filter((month) => month >= state.startMonth && month <= state.endMonth);
}

function isIsoWeekKey(value) {
  return /^\d{4}-W\d{2}$/.test(String(value || ""));
}

function weekLabel(weekKey) {
  const range = isoWeekRange(weekKey);
  return `${weekKey} (${range.start.date} - ${range.end.date})`;
}

function isoWeekRange(weekKey) {
  const [yearText, weekText] = weekKey.split("-W");
  const year = Number(yearText);
  const week = Number(weekText);
  const jan4 = new Date(Date.UTC(year, 0, 4));
  const jan4Day = jan4.getUTCDay() || 7;
  const startTs = jan4.getTime() - (jan4Day - 1) * 86400000 + (week - 1) * 7 * 86400000;
  const start = new Date(startTs);
  const end = new Date(startTs + 6 * 86400000);
  return {
    start: makeDateParts(dateToIso(start)),
    end: makeDateParts(dateToIso(end)),
  };
}

function dateToIso(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function isoWeekKeyFromDate(value) {
  const parts = typeof value === "string" ? parseDateValue(value) : value;
  if (!parts) return "";
  const date = new Date(parts.ts);
  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - day);
  const isoYear = date.getUTCFullYear();
  const yearStart = new Date(Date.UTC(isoYear, 0, 1));
  const week = Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
  return `${isoYear}-W${String(week).padStart(2, "0")}`;
}

function weeksInIsoYear(year) {
  return Number(isoWeekKeyFromDate(`${year}-12-28`).slice(6));
}

function dealTrackingDate(deal) {
  return deal.trackingDate || (deal.status === "open" ? deal.expectedDate : deal.stageChangeDate || deal.expectedDate || deal.createdDate);
}

function dealPeriodKey(deal, periodKey) {
  if (isIsoWeekKey(periodKey)) return isoWeekKeyFromDate(dealTrackingDate(deal));
  return deal.trackingMonth || (deal.status === "open" ? deal.expectedMonth : deal.stageMonth || deal.expectedMonth);
}

function dealPeriodLabel(deal) {
  const key = dealPeriodKey(deal, state.periodType === "week" ? state.periodWeek : deal.trackingMonth);
  return key ? monthLabel(key) : "-";
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

function selectedSaleRecord() {
  if (state.sale === "all") return null;
  return dashboardData.sales.find((sale) => sale.key === state.sale) || null;
}

function itemMatchesSaleRecord(item, sale) {
  if (!sale) return true;
  if (item.saleKey === sale.key) return true;
  const saleTerms = [sale.name, ...(sale.responsibleNames || [])].map(normalizeName).filter(Boolean);
  const itemTerms = [item.saleName, item.responsible].map(normalizeName).filter(Boolean);
  return itemTerms.some((itemTerm) =>
    saleTerms.some((saleTerm) => itemTerm === saleTerm || itemTerm.startsWith(`${saleTerm} `) || saleTerm.startsWith(`${itemTerm} `)),
  );
}

function itemMatchesSaleScope(item, scope) {
  return scope.saleSet.has(item.saleKey) || (scope.selectedSale && itemMatchesSaleRecord(item, scope.selectedSale));
}

function dealSearchText(deal) {
  return normalizeSearch(
    `${deal.id} ${deal.saleName} ${deal.responsible} ${deal.group} ${deal.company} ${deal.dealName} ${deal.pipeline} ${deal.stage} ${deal.rawStage || ""} ${deal.dealType || ""} ${deal.product || ""} ${deal.category || ""} ${performanceBucket(deal)} ${deal.expectedDate || ""} ${deal.stageChangeDate || ""} ${deal.createdDate || ""}`,
  );
}

function dealMatchesGlobalSearch(deal, search = state.search) {
  const q = normalizeSearch(search);
  if (!q) return true;
  return dealSearchText(deal).includes(q);
}

function dealMatchesPeriod(deal, scope, periodKey = null) {
  if (periodKey) return dealPeriodKey(deal, periodKey) === periodKey;
  return scope.monthSet.has(dealPeriodKey(deal, scope.months[0]));
}

function filteredDeals(scope, options = {}) {
  const {
    category = "state",
    period = true,
    periodKey = null,
    countingIncluded = true,
    globalSearch = true,
    statuses = null,
  } = options;
  return allDealDetails()
    .filter((deal) => itemMatchesSaleScope(deal, scope))
    .filter((deal) => categoryMatches(deal.category))
    .filter((deal) => category === "state" || category === "all" || deal.category === category)
    .filter((deal) => !countingIncluded || deal.countingIncluded !== false)
    .filter((deal) => !period || dealMatchesPeriod(deal, scope, periodKey))
    .filter((deal) => !globalSearch || dealMatchesGlobalSearch(deal))
    .filter((deal) => !statuses || statuses.includes(performanceBucket(deal)));
}

function targetForSale(sale, months, category = state.category) {
  return months.reduce((total, month) => {
    if (isIsoWeekKey(month)) return total + targetForSaleWeek(sale, month, category);
    const bucket = sale.monthly[month] || { renew: 0, new: 0, total: 0 };
    if (category === "renew") return total + bucket.renew;
    if (category === "new") return total + bucket.new;
    return total + bucket.total;
  }, 0);
}

function targetForSaleWeek(sale, weekKey, category = state.category) {
  const range = isoWeekRange(weekKey);
  let total = 0;
  for (let ts = range.start.ts; ts <= range.end.ts; ts += 86400000) {
    const date = new Date(ts);
    const month = `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
    const bucket = sale.monthly[month] || { renew: 0, new: 0, total: 0 };
    const daysInMonth = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0)).getUTCDate();
    const dayTarget =
      category === "renew" ? bucket.renew / daysInMonth : category === "new" ? bucket.new / daysInMonth : bucket.total / daysInMonth;
    total += dayTarget;
  }
  return total;
}

function calcScope() {
  const months = periodMonths();
  const monthSet = new Set(months);
  const sales = visibleSales();
  const saleSet = new Set(sales.map((sale) => sale.key));
  const selectedSale = selectedSaleRecord();
  const baseScope = { months, monthSet, saleSet, selectedSale };
  const periodDeals = filteredDeals(baseScope);
  const facts = periodDeals
    .filter((deal) => deal.status === "won" || deal.status === "lost")
    .map((deal) => ({
      saleKey: deal.saleKey,
      monthKey: dealPeriodKey(deal, months[0]),
      category: deal.category,
      status: deal.status,
      group: deal.group,
      saleName: deal.saleName,
      amount: deal.amount,
      count: 1,
    }));
  const actualFacts = facts.filter((fact) => fact.status === "won");
  const lostFacts = facts.filter((fact) => fact.status === "lost");
  const openPeriod = periodDeals.filter((deal) => deal.status === "open");
  const openAll = openPeriod;

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
    selectedSale,
    periodDeals,
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
  renderPeriodOptions();
  renderGroupOptions();
  renderSaleOptions();
  renderFilterVisibility();

  els.viewMode.addEventListener("change", () => {
    state.viewMode = els.viewMode.value;
    renderFilterVisibility();
    renderSaleOptions();
    render();
  });
  els.periodType.addEventListener("change", () => {
    state.periodType = els.periodType.value;
    renderFilterVisibility();
    render();
  });
  els.periodYear.addEventListener("change", () => {
    state.periodYear = els.periodYear.value;
    render();
  });
  els.periodQuarter.addEventListener("change", () => {
    state.periodQuarter = els.periodQuarter.value;
    render();
  });
  els.periodWeek.addEventListener("change", () => {
    state.periodWeek = els.periodWeek.value;
    render();
  });
  els.startMonth.addEventListener("change", () => {
    state.startMonth = els.startMonth.value;
    if (state.startMonth > state.endMonth) {
      state.endMonth = state.startMonth;
      els.endMonth.value = state.endMonth;
    }
    render();
  });
  els.endMonth.addEventListener("change", () => {
    state.endMonth = els.endMonth.value;
    if (state.endMonth < state.startMonth) {
      state.startMonth = state.endMonth;
      els.startMonth.value = state.startMonth;
    }
    render();
  });
  els.leadStartWeek.addEventListener("change", () => {
    state.leadStartWeek = els.leadStartWeek.value;
    if (state.leadStartWeek > state.leadEndWeek) {
      state.leadEndWeek = state.leadStartWeek;
      els.leadEndWeek.value = state.leadEndWeek;
    }
    renderTablesOnly();
  });
  els.leadEndWeek.addEventListener("change", () => {
    state.leadEndWeek = els.leadEndWeek.value;
    if (state.leadEndWeek < state.leadStartWeek) {
      state.leadStartWeek = state.leadEndWeek;
      els.leadStartWeek.value = state.leadStartWeek;
    }
    renderTablesOnly();
  });
  els.leadSearchInput.addEventListener("input", () => {
    state.leadSearch = els.leadSearchInput.value;
    if (state.tab === "leads") renderTablesOnly();
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
  els.salesStatusFilters.querySelectorAll("input[data-sales-status]").forEach((input) => {
    input.addEventListener("change", () => {
      state.statusFilters[input.dataset.salesStatus] = input.checked;
      renderTablesOnly();
    });
  });
  [...els.transactionAndInputs, ...els.transactionNotInputs].forEach((input) => {
    input.addEventListener("input", () => {
      if (state.tab === "sales") renderTablesOnly();
    });
  });
  els.transactionTable.addEventListener("click", (event) => {
    const sortButton = event.target.closest("[data-transaction-sort]");
    if (sortButton) {
      const key = sortButton.dataset.transactionSort;
      if (state.transactionSort.key === key) {
        state.transactionSort.dir = state.transactionSort.dir === "asc" ? "desc" : "asc";
      } else {
        state.transactionSort = { key, dir: key === "amount" ? "desc" : "asc" };
      }
      renderTablesOnly();
      return;
    }
    const dealRow = event.target.closest("[data-deal-key]");
    if (dealRow) openDealModal(dealRow.dataset.dealKey);
  });
  els.dealDetailTable.addEventListener("click", (event) => {
    const sortButton = event.target.closest("[data-deal-detail-sort]");
    if (sortButton) {
      const key = sortButton.dataset.dealDetailSort;
      if (state.dealDetailSort.key === key) {
        state.dealDetailSort.dir = state.dealDetailSort.dir === "asc" ? "desc" : "asc";
      } else {
        state.dealDetailSort = { key, dir: ["amount", "forecastAmount", "stageAgeDays"].includes(key) ? "desc" : "asc" };
      }
      renderTablesOnly();
      return;
    }
    const dealRow = event.target.closest("[data-deal-key]");
    if (dealRow) openDealModal(dealRow.dataset.dealKey);
  });
  [els.topDealsTable, els.riskDealsTable].forEach((table) => {
    table.addEventListener("click", (event) => {
      const dealRow = event.target.closest("[data-deal-key]");
      if (dealRow) openDealModal(dealRow.dataset.dealKey);
    });
  });
  els.newLeadTable.addEventListener("click", (event) => {
    const dealRow = event.target.closest("[data-deal-key]");
    if (dealRow) openDealModal(dealRow.dataset.dealKey);
  });
  els.savePipelineMatching.addEventListener("click", savePipelineMatching);
  els.clearPipelineMatching.addEventListener("click", clearPipelineMatching);
  document.addEventListener("click", (event) => {
    const settingsButton = event.target.closest("#openSettings");
    if (settingsButton) openSettings();
  });
  els.closeSettings.addEventListener("click", closeSettings);
  els.closeDealModal.addEventListener("click", closeDealModal);
  els.settingsModal.addEventListener("click", (event) => {
    if (event.target === els.settingsModal) closeSettings();
  });
  els.dealModal.addEventListener("click", (event) => {
    if (event.target === els.dealModal) closeDealModal();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    if (!els.dealModal.hidden) closeDealModal();
    if (!els.settingsModal.hidden) closeSettings();
  });
  [els.targetCsvInput, els.dealCsvInput, els.mappingCsvInput].forEach((input) => {
    input.addEventListener("change", () => {
      const targetFile = els.targetCsvInput.files?.[0]?.name;
      const dealFile = els.dealCsvInput.files?.[0]?.name;
      const mappingFile = els.mappingCsvInput.files?.[0]?.name;
      if (input === els.targetCsvInput && targetFile) setAutoLoadedFileName("target", "");
      if (input === els.mappingCsvInput && mappingFile) setAutoLoadedFileName("mapping", "");
      const selected = [
        targetFile && `Target: ${targetFile}`,
        dealFile && `Deal: ${dealFile}`,
        mappingFile && `Mapping: ${mappingFile}`,
      ].filter(Boolean);
      setUploadStatus(selected.length ? selected.join(" | ") : "ยังไม่ได้เลือกไฟล์ใหม่");
    });
  });
  els.applyCsvFiles.addEventListener("click", handleCsvRefresh);
  els.clearCsvFiles.addEventListener("click", () => {
    els.targetCsvInput.value = "";
    els.dealCsvInput.value = "";
    els.mappingCsvInput.value = "";
    setAutoLoadedFileName("target", "");
    setAutoLoadedFileName("mapping", "");
    window.localStorage.removeItem(STORED_DATA_KEY);
    dashboardData = JSON.parse(JSON.stringify(initialDashboardData));
    state.group = "all";
    state.sale = "all";
    renderPeriodOptions();
    renderGroupOptions();
    renderSaleOptions();
    renderFilterVisibility();
    render();
    setUploadStatus("Reset กลับไปใช้ข้อมูลตั้งต้นแล้ว", "success");
    autoLoadDefaultSetupFiles();
  });
}

function renderPeriodOptions() {
  const yearOptions = ["2024", "2025", "2026", "2027", "2028"];
  els.periodYear.innerHTML = yearOptions.map((year) => `<option value="${year}">${year}</option>`).join("");
  els.periodYear.value = state.periodYear;

  const quarterOptions = [];
  for (let year = 2026; year <= 2028; year += 1) {
    for (let quarter = 1; quarter <= 4; quarter += 1) {
      quarterOptions.push({ value: `${year}-Q${quarter}`, label: `${ordinalQuarter(quarter)} ${year}` });
    }
  }
  els.periodQuarter.innerHTML = quarterOptions.map((quarter) => `<option value="${quarter.value}">${quarter.label}</option>`).join("");
  els.periodQuarter.value = state.periodQuarter;

  const weekOptions = [];
  for (let year = 2026; year <= 2028; year += 1) {
    for (let week = 1; week <= weeksInIsoYear(year); week += 1) {
      const value = `${year}-W${String(week).padStart(2, "0")}`;
      weekOptions.push({ value, label: weekLabel(value) });
    }
  }
  els.periodWeek.innerHTML = weekOptions.map((week) => `<option value="${week.value}">${escapeHtml(week.label)}</option>`).join("");
  if (!weekOptions.some((week) => week.value === state.periodWeek)) state.periodWeek = weekOptions[0]?.value || "2026-W01";
  els.periodWeek.value = state.periodWeek;

  const leadWeeks = leadWeekOptions();
  const leadOptions = leadWeeks.map((week) => `<option value="${week.value}">${escapeHtml(week.label)}</option>`).join("");
  els.leadStartWeek.innerHTML = leadOptions;
  els.leadEndWeek.innerHTML = leadOptions;
  if (!leadWeeks.some((week) => week.value === state.leadStartWeek)) state.leadStartWeek = leadWeeks.find((week) => week.value === "2026-W23")?.value || leadWeeks[0]?.value || "2026-W01";
  if (!leadWeeks.some((week) => week.value === state.leadEndWeek)) state.leadEndWeek = leadWeeks.find((week) => week.value === "2026-W26")?.value || leadWeeks[leadWeeks.length - 1]?.value || state.leadStartWeek;
  if (state.leadStartWeek > state.leadEndWeek) state.leadEndWeek = state.leadStartWeek;
  els.leadStartWeek.value = state.leadStartWeek;
  els.leadEndWeek.value = state.leadEndWeek;

  const rangeMonths = (dashboardData.months || []).filter((month) => month >= "2026-01" && month <= "2028-12");
  const rangeOptions = rangeMonths.map((month) => `<option value="${month}">${monthLabel(month)}</option>`).join("");
  els.startMonth.innerHTML = rangeOptions;
  els.endMonth.innerHTML = rangeOptions;
  if (!rangeMonths.includes(state.startMonth)) state.startMonth = rangeMonths[0] || "2026-01";
  if (!rangeMonths.includes(state.endMonth)) state.endMonth = rangeMonths[rangeMonths.length - 1] || "2026-12";
  els.startMonth.value = state.startMonth;
  els.endMonth.value = state.endMonth;
}

function leadWeekOptions() {
  const weekSet = new Set();
  allDealDetails().forEach((deal) => {
    const week = isoWeekKeyFromDate(deal.createdDate);
    if (week) weekSet.add(week);
  });
  const years = new Set((dashboardData.months || []).map((month) => Number(month.slice(0, 4))).filter(Boolean));
  if (!years.size) {
    years.add(2026);
    years.add(2027);
    years.add(2028);
  }
  for (const year of years) {
    for (let week = 1; week <= weeksInIsoYear(year); week += 1) weekSet.add(`${year}-W${String(week).padStart(2, "0")}`);
  }
  return Array.from(weekSet)
    .sort()
    .map((week) => ({ value: week, label: weekLabel(week) }));
}

function ordinalQuarter(quarter) {
  return ["1st", "2nd", "3rd", "4th"][quarter - 1] || `${quarter}th`;
}

function renderFilterVisibility() {
  els.viewMode.value = state.viewMode;
  els.periodType.value = state.periodType;
  els.groupFilterContainer.hidden = false;
  els.salesFilterContainer.hidden = false;
  els.periodSubYear.hidden = state.periodType !== "year";
  els.periodSubQuarter.hidden = state.periodType !== "quarter";
  els.periodSubWeek.hidden = state.periodType !== "week";
  els.periodSubRange.hidden = state.periodType !== "range";
}

function openSettings() {
  renderPipelineMatchingSettings();
  els.settingsModal.hidden = false;
  els.targetCsvInput.focus();
}

function closeSettings() {
  els.settingsModal.hidden = true;
}

function closeDealModal() {
  els.dealModal.hidden = true;
}

function setUploadStatus(message, tone = "") {
  els.uploadStatus.textContent = message;
  els.uploadStatus.className = `upload-status ${tone}`.trim();
}

function setAutoLoadedFileName(kind, fileName) {
  const target = kind === "target" ? els.targetAutoFileName : els.mappingAutoFileName;
  if (!target) return;
  target.hidden = !fileName;
  target.textContent = fileName ? `${fileName} - โหลดขึ้นระบบเรียบร้อยแล้ว` : "";
}

function pipelineMatchingGroups() {
  return Array.from(
    new Set(
      (dashboardData.sales || [])
        .filter((sale) => sale.group && sale.group !== "Unmapped" && (sale.hasTarget || sale.hasPositiveTarget))
        .map((sale) => sale.group),
    ),
  ).sort((a, b) => a.localeCompare(b, "th"));
}

function pipelineMatchingPipelines() {
  const seen = new Set();
  const add = (pipeline) => {
    const clean = cleanText(pipeline);
    const key = normalizeName(clean);
    if (!clean || seen.has(key)) return null;
    seen.add(key);
    return clean;
  };
  const ordered = defaultPipelineOrder.map(add).filter(Boolean);
  const fromDeals = Array.from(new Set(allDealDetails().map((deal) => cleanText(deal.pipeline)).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b, "th"),
  );
  fromDeals.forEach((pipeline) => {
    const added = add(pipeline);
    if (added) ordered.push(added);
  });
  return ordered;
}

function renderPipelineMatchingSettings() {
  const groups = pipelineMatchingGroups();
  const pipelines = pipelineMatchingPipelines();
  const mappings = activeMappings();
  const ruleCount = mappingCounts(mappings).pipelineRules || 0;
  els.pipelineMatchingSummary.textContent = ruleCount
    ? `มี Pipeline Matching ที่เลือกไว้ ${ruleCount.toLocaleString("th-TH")} จุด ระบบจะนับยอดเฉพาะรายการ Included`
    : "ยังไม่กำหนด Pipeline Matching ระบบจะนับทุก Pipeline เหมือนเดิม";
  if (!groups.length || !pipelines.length) {
    els.pipelineMatchingMatrix.innerHTML = `<div class="empty">ยังไม่มีข้อมูล Sale Group หรือ Pipeline สำหรับสร้าง Matrix</div>`;
    return;
  }
  const header = `
    <thead>
      <tr>
        <th class="sticky-col group-col">Sale Group</th>
        <th class="sticky-col type-col">Sale Type</th>
        ${pipelines.map((pipeline) => `<th class="pipeline-heading"><span>${escapeHtml(pipeline)}</span></th>`).join("")}
      </tr>
    </thead>
  `;
  const body = groups
    .flatMap((group) =>
      saleTypes.map((saleType) => {
        const allowed = mappings.pipelineRuleMap.get(pipelineRuleKey(group, saleType)) || new Set();
        return `
          <tr>
            <td class="sticky-col group-col"><strong>${escapeHtml(group)}</strong></td>
            <td class="sticky-col type-col">${escapeHtml(saleType === "renew" ? "Renew" : "New")}</td>
            ${pipelines
              .map((pipeline) => {
                const checked = allowed.has(normalizeName(pipeline)) ? " checked" : "";
                return `
                  <td class="pipeline-match-cell">
                    <input type="checkbox" aria-label="${escapeHtml(`${group} ${saleType} ${pipeline}`)}" data-match-group="${escapeHtml(group)}" data-match-type="${saleType}" data-match-pipeline="${escapeHtml(pipeline)}"${checked} />
                  </td>
                `;
              })
              .join("")}
          </tr>
        `;
      }),
    )
    .join("");
  els.pipelineMatchingMatrix.innerHTML = `<table class="pipeline-matrix-table">${header}<tbody>${body}</tbody></table>`;
}

function collectPipelineMatchingRules() {
  const pipelineRuleMap = new Map();
  pipelineMatchingGroups().forEach((group) => {
    saleTypes.forEach((saleType) => pipelineRuleMap.set(pipelineRuleKey(group, saleType), new Set()));
  });
  els.pipelineMatchingMatrix.querySelectorAll("input[data-match-pipeline]:checked").forEach((input) => {
    const key = pipelineRuleKey(input.dataset.matchGroup, input.dataset.matchType);
    if (!pipelineRuleMap.has(key)) pipelineRuleMap.set(key, new Set());
    pipelineRuleMap.get(key).add(normalizeName(input.dataset.matchPipeline));
  });
  return pipelineRuleMap;
}

function applyPipelineMatchingRules(pipelineRuleMap, message) {
  const mappings = activeMappings();
  mappings.pipelineRuleMap = pipelineRuleMap;
  mappings.counts = mappingCounts(mappings);
  dashboardData = remapDashboardDataWithMappings(
    {
      ...dashboardData,
      mappings: serializeMappings(mappings),
      metadata: {
        ...dashboardData.metadata,
        generatedAt: new Date().toISOString(),
        sourceFiles: {
          ...dashboardData.metadata.sourceFiles,
          mapping: "Pipeline Matching Setting",
        },
      },
    },
    mappings,
  );
  storeDashboardData(dashboardData);
  renderPeriodOptions();
  renderGroupOptions();
  renderSaleOptions();
  renderFilterVisibility();
  render();
  renderPipelineMatchingSettings();
  setUploadStatus(message, "success");
}

function savePipelineMatching() {
  applyPipelineMatchingRules(collectPipelineMatchingRules(), "บันทึก Pipeline Matching และคำนวณ Dashboard ใหม่แล้ว");
}

function clearPipelineMatching() {
  applyPipelineMatchingRules(new Map(), "ล้าง Pipeline Matching แล้ว ระบบกลับไปนับทุก Pipeline");
}

async function autoLoadDefaultPipelineMatching() {
  const currentMappings = activeMappings();
  if ((mappingCounts(currentMappings).pipelineRules || 0) > 0) return;
  try {
    const response = await fetch("./data/Stage Mapping Button.csv", { cache: "no-store" });
    if (!response.ok) return;
    const defaultRules = parseMappingCsv(await response.text());
    if (!(mappingCounts(defaultRules).pipelineRules || 0)) return;
    const mappings = mergeMappings(currentMappings, defaultRules);
    dashboardData = remapDashboardDataWithMappings(
      {
        ...dashboardData,
        mappings: serializeMappings(mappings),
        metadata: {
          ...dashboardData.metadata,
          generatedAt: new Date().toISOString(),
          sourceFiles: {
            ...dashboardData.metadata.sourceFiles,
            mapping: "./data/Stage Mapping Button.csv",
          },
        },
      },
      mappings,
    );
    storeDashboardData(dashboardData);
    renderPeriodOptions();
    renderGroupOptions();
    renderSaleOptions();
    renderFilterVisibility();
    render();
  } catch {
    // The dashboard can still run without the optional default rule file.
  }
}

async function autoLoadDefaultSetupFiles() {
  const statusParts = [];
  let mappings = activeMappings();
  try {
    const stageResponse = await fetch("./data/Stage Mapping.csv", { cache: "no-store" });
    if (stageResponse.ok) {
      const stageMappings = parseMappingCsv(await stageResponse.text());
      if (mappingCounts(stageMappings).stages || mappingCounts(stageMappings).dealTypes || mappingCounts(stageMappings).pipelines) {
        mappings = mergeMappings(mappings, stageMappings);
        statusParts.push("Stage Mapping");
        setAutoLoadedFileName("mapping", "Stage Mapping.csv");
      }
    }

    const buttonResponse = await fetch("./data/Stage Mapping Button.csv", { cache: "no-store" });
    if (buttonResponse.ok) {
      const buttonMappings = parseMappingCsv(await buttonResponse.text());
      if (mappingCounts(buttonMappings).pipelineRules) {
        mappings = mergeMappings(mappings, buttonMappings);
        statusParts.push("Pipeline Matching");
      }
    }

    dashboardData = remapDashboardDataWithMappings(
      {
        ...dashboardData,
        mappings: serializeMappings(mappings),
        metadata: {
          ...dashboardData.metadata,
          generatedAt: new Date().toISOString(),
          sourceFiles: {
            ...dashboardData.metadata.sourceFiles,
            mapping: statusParts.join(" + "),
          },
        },
      },
      mappings,
    );

    const targetResponse = await fetch("./data/Sale Target.csv", { cache: "no-store" });
    if (targetResponse.ok) {
      const targetRows = parseCsv(await targetResponse.text());
      if (targetRows.length) {
        validateTargetRows(targetRows);
        const targetSales = buildTargetSalesFromRows(targetRows, mappings);
        dashboardData = mergeTargetSalesIntoDashboard(dashboardData, targetSales, "./data/Sale Target.csv");
        statusParts.push("Sale Target");
        setAutoLoadedFileName("target", "Sale Target.csv");
      }
    }

    storeDashboardData(dashboardData);
    renderPeriodOptions();
    renderGroupOptions();
    renderSaleOptions();
    renderFilterVisibility();
    render();
    if (statusParts.length) setUploadStatus(`Auto-load สำเร็จ: ${statusParts.join(" | ")}`, "success");
  } catch (error) {
    setUploadStatus(error.message || "Auto-load setup file ไม่สำเร็จ", "error");
  }
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

async function handleCsvRefresh() {
  const targetFile = els.targetCsvInput.files?.[0];
  const dealFile = els.dealCsvInput.files?.[0];
  const mappingFile = els.mappingCsvInput.files?.[0];
  if (!targetFile && !dealFile && !mappingFile) {
    setUploadStatus("กรุณาเลือก Sale Target.csv, Deal.csv หรือ Data Mapping.csv อย่างน้อย 1 ไฟล์", "error");
    return;
  }
  if ([targetFile, dealFile, mappingFile].some((file) => file && !file.name.toLowerCase().endsWith(".csv"))) {
    setUploadStatus("รองรับเฉพาะไฟล์ .csv เท่านั้น", "error");
    return;
  }

  try {
    setUploadStatus("กำลังอ่านและประมวลผลไฟล์...");
    let targetSales = currentTargetSales();
    let mappings = activeMappings();
    const statusParts = [];

    if (mappingFile) {
      const uploadedMappings = parseMappingCsv(await mappingFile.text());
      validateMappings(uploadedMappings);
      mappings = mergeMappings(mappings, uploadedMappings);
      dashboardData = remapDashboardDataWithMappings({
        ...dashboardData,
        mappings: serializeMappings(mappings),
        metadata: {
          ...dashboardData.metadata,
          generatedAt: new Date().toISOString(),
          sourceFiles: {
            ...dashboardData.metadata.sourceFiles,
            mapping: mappingFile.name,
          },
        },
      }, mappings);
      statusParts.push(
        `Mapping ${uploadedMappings.counts.pipelines}/${uploadedMappings.counts.stages}/${uploadedMappings.counts.dealTypes}/${uploadedMappings.counts.sales}`,
      );
    }

    if (targetFile) {
      const targetRows = parseCsv(await targetFile.text());
      validateTargetRows(targetRows);
      targetSales = buildTargetSalesFromRows(targetRows, mappings);
      dashboardData = mergeTargetSalesIntoDashboard(dashboardData, targetSales, targetFile.name);
      statusParts.push(`Target ${targetRows.length.toLocaleString("th-TH")} rows`);
    }

    if (dealFile) {
      const dealRows = parseCsv(await dealFile.text());
      validateDealRows(dealRows);
      dashboardData = buildDashboardDataFromDeals(dealRows, dealFile.name, targetSales, mappings);
      statusParts.push(`Deal ${dealRows.length.toLocaleString("th-TH")} rows`);
    }

    storeDashboardData(dashboardData);
    state.group = "all";
    state.sale = "all";
  state.search = "";
  state.leadSearch = "";
  els.searchInput.value = "";
  els.leadSearchInput.value = "";
    renderPeriodOptions();
    renderGroupOptions();
    renderSaleOptions();
    renderFilterVisibility();
    render();
    if (!els.settingsModal.hidden) renderPipelineMatchingSettings();
    setUploadStatus(`Refresh สำเร็จ: ${statusParts.join(" | ")}`, "success");
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

function validateTargetRows(rows) {
  if (!rows.length) throw new Error("ไฟล์ Sale Target ไม่มีข้อมูล");
  const required = ["sales", "sale group"];
  const missing = required.filter((column) => !(column in rows[0]));
  if (missing.length) {
    throw new Error(`รูปแบบ CSV ไม่ตรงกับไฟล์ Sale Target: ขาดคอลัมน์ ${missing.join(", ")}`);
  }
}

function buildTargetSalesFromRows(rows, mappings = activeMappings()) {
  const months = dashboardData.months || initialDashboardData.months || [];
  return rows
    .filter((row) => cleanText(row.sales))
    .map((row) => {
      const saleName = cleanText(row.sales).replace(/\([^)]*\)/g, "").replace(/\s+/g, " ").trim();
      const mappedSale = mappings.salesGroupMap.get(normalizeName(saleName));
      const monthly = {};
      months.forEach((month) => {
        monthly[month] = {
          renew: parseMoneyValue(row[`renew ${month}`]),
          new: parseMoneyValue(row[`new ${month}`]),
        };
        monthly[month].total = monthly[month].renew + monthly[month].new;
      });
      for (let year = 2024; year <= 2028; year += 1) {
        const yearMonths = months.filter((month) => month.startsWith(`${year}-`));
        const renewMonthly = sum(yearMonths, (month) => monthly[month].renew);
        const newMonthly = sum(yearMonths, (month) => monthly[month].new);
        const annualRenew = parseMoneyValue(row[`renew ${year}`]);
        const annualNew = parseMoneyValue(row[`new ${year}`]);
        if (!renewMonthly && annualRenew > 0 && yearMonths.length) {
          yearMonths.forEach((month) => {
            monthly[month].renew = annualRenew / yearMonths.length;
            monthly[month].total = monthly[month].renew + monthly[month].new;
          });
        }
        if (!newMonthly && annualNew > 0 && yearMonths.length) {
          yearMonths.forEach((month) => {
            monthly[month].new = annualNew / yearMonths.length;
            monthly[month].total = monthly[month].renew + monthly[month].new;
          });
        }
      }
      const targetAnnual = {
        renew: roundMoney(sum(months, (month) => monthly[month].renew)),
        new: roundMoney(sum(months, (month) => monthly[month].new)),
        total: roundMoney(sum(months, (month) => monthly[month].total)),
      };
      return {
        key: makeKey(saleName),
        name: mappedSale?.displayName || saleName,
        group: mappedSale?.group || cleanText(row["sale group"]) || "No group",
        hasTarget: true,
        hasPositiveTarget: targetAnnual.total > 0,
        responsibleNames: [],
        monthly,
        targetAnnual,
      };
    });
}

function mergeTargetSalesIntoDashboard(data, targetSales, fileName) {
  const targetMap = new Map(targetSales.map((sale) => [sale.key, sale]));
  const sales = [
    ...targetSales.map((sale) => {
      const previous = data.sales.find((item) => item.key === sale.key);
      return {
        ...sale,
        responsibleNames: previous?.responsibleNames || [],
      };
    }),
    ...data.sales.filter((sale) => !sale.hasTarget && !targetMap.has(sale.key)),
  ].sort((a, b) => a.group.localeCompare(b.group, "th") || a.name.localeCompare(b.name, "th"));
  const enrich = (item) => {
    const target = targetMap.get(item.saleKey);
    return target ? { ...item, saleName: target.name, group: target.group } : item;
  };
  return {
    ...data,
    metadata: {
      ...data.metadata,
      generatedAt: new Date().toISOString(),
      sourceFiles: {
        ...data.metadata.sourceFiles,
        target: fileName,
      },
    },
    sales,
    facts: data.facts.map(enrich),
    openDeals: data.openDeals.map(enrich),
    dealDetails: (data.dealDetails || []).map(enrich),
    quality: {
      ...data.quality,
      targetRows: targetSales.length,
      zeroTargetSales: {
        count: targetSales.filter((sale) => !sale.hasPositiveTarget).length,
        amount: 0,
      },
    },
  };
}

function remapDashboardDataWithMappings(data, mappings) {
  const months = data.months || dashboardData.months || [];
  const today = makeDateParts(localIsoDate());
  const dayMs = 24 * 60 * 60 * 1000;
  const salesByKey = new Map((data.sales || []).map((sale) => [sale.key, sale]));
  const factsMap = new Map();
  const openDeals = [];
  const quality = {
    dealRows: (data.dealDetails || []).length || data.quality?.dealRows || 0,
    targetRows: (data.sales || []).filter((sale) => sale.hasTarget).length,
    preWonAsWon: { count: 0, amount: 0 },
    preLostAsLost: { count: 0, amount: 0 },
    won: { count: 0, amount: 0 },
    lost: { count: 0, amount: 0 },
    open: { count: 0, amount: 0 },
    overdueOpen: { count: 0, amount: 0 },
    noExpectedOpen: { count: 0, amount: 0 },
    stale30Open: { count: 0, amount: 0 },
    notContactedOpen: { count: 0, amount: 0 },
    pipelineGroupMismatch: { count: 0, amount: 0 },
    pipelineIncluded: { count: 0, amount: 0 },
    pipelineExcluded: { count: 0, amount: 0 },
    pipelineUnmapped: { count: 0, amount: 0 },
  };

  const sourceDetails = Array.isArray(data.dealDetails)
    ? data.dealDetails
    : (data.openDeals || []).map((deal) => ({
        ...deal,
        status: "open",
        trackingMonth: deal.expectedMonth,
      }));
  const details = sourceDetails.map((deal) => {
    const rawStage = cleanText(deal.rawStage || deal.stage);
    const stage = mappedStage(rawStage, mappings);
    const status = normalizeDealStage(stage);
    const rawDealType = cleanText(deal.rawDealType || deal.dealType);
    const dealType = mappedDealType(rawDealType, mappings);
    const category = dealCategory(dealType, { Pipeline: deal.pipeline, "Deal Type": rawDealType });
    const sale = salesByKey.get(deal.saleKey);
    const pipelineGroup = mappings.pipelineGroupMap.get(normalizeName(deal.pipeline)) || "";
    const counting = pipelineCountingResult(sale?.group || deal.group, category, deal.pipeline, mappings);
    const expected = parseDateValue(deal.expectedDate);
    const stageChanged = parseDateValue(deal.stageChangeDate);
    const created = parseDateValue(deal.createdDate);
    const closeDate = stageChanged || expected || created;
    const stageAgeDays = stageChanged ? Math.floor((today.ts - stageChanged.ts) / dayMs) : deal.stageAgeDays ?? null;
    const trackingMonth = status === "open" ? expected?.monthKey || deal.expectedMonth || "" : closeDate?.monthKey || deal.trackingMonth || "";
    const trackingDate = status === "open" ? expected?.date || "" : closeDate?.date || "";
    const amount = Number(deal.amount) || 0;
    const forecastAmount = status === "open" ? roundMoney(amount * stageWeightFor(rawStage, stage)) : 0;
    const risk = {
      overdue: Boolean(expected && expected.ts < today.ts),
      noExpected: !expected,
      stale30: Boolean(stageChanged && stageChanged.ts < today.ts - 30 * dayMs),
      notContacted: stage === "Not Contacted" || rawStage === "Not Contacted",
    };
    const remapped = {
      ...deal,
      saleName: sale?.name || deal.saleName,
      group: sale?.group || deal.group,
      category,
      status,
      stage,
      matchedStage: stage,
      rawStage,
      dealType,
      rawDealType,
      pipelineGroup,
      pipelineGroupMatch: !pipelineGroup || pipelineGroup === (sale?.group || deal.group),
      countingStatus: counting.status,
      countingIncluded: counting.included,
      countingLabel: counting.label,
      forecastAmount,
      expectedMonth: expected?.monthKey || deal.expectedMonth || "",
      stageMonth: stageChanged?.monthKey || deal.stageMonth || "",
      trackingMonth,
      trackingDate,
      stageAgeDays,
      risk,
    };

    if (rawStage === "Pre-WON") addQuality(quality.preWonAsWon, amount);
    if (rawStage === "Pre-LOST") addQuality(quality.preLostAsLost, amount);
    if (pipelineGroup && pipelineGroup !== remapped.group) addQuality(quality.pipelineGroupMismatch, amount);
    addQuality(
      counting.status === "included"
        ? quality.pipelineIncluded
        : counting.status === "excluded"
          ? quality.pipelineExcluded
          : quality.pipelineUnmapped,
      amount,
    );

    if (!counting.included) return remapped;

    if (status === "won" || status === "lost") {
      addQuality(status === "won" ? quality.won : quality.lost, amount);
      if (months.includes(trackingMonth)) {
        addFact(factsMap, {
          saleKey: remapped.saleKey,
          monthKey: trackingMonth,
          category,
          status,
          group: remapped.group,
          saleName: remapped.saleName,
          amount,
        });
      }
    } else {
      addQuality(quality.open, amount);
      if (risk.noExpected) addQuality(quality.noExpectedOpen, amount);
      if (risk.overdue) addQuality(quality.overdueOpen, amount);
      if (risk.stale30) addQuality(quality.stale30Open, amount);
      if (risk.notContacted) addQuality(quality.notContactedOpen, amount);
      openDeals.push({
        ...remapped,
        weight: stageWeightFor(rawStage, stage),
      });
    }
    return remapped;
  });

  for (const bucket of Object.values(quality)) {
    if (bucket && typeof bucket === "object" && "amount" in bucket) bucket.amount = roundMoney(bucket.amount);
  }

  return {
    ...data,
    facts: Array.from(factsMap.values()).map((fact) => ({ ...fact, amount: roundMoney(fact.amount) })),
    openDeals,
    dealDetails: details,
    quality: {
      ...quality,
      zeroTargetSales: data.quality?.zeroTargetSales || {
        count: (data.sales || []).filter((sale) => !sale.hasPositiveTarget).length,
        amount: 0,
      },
    },
  };
}

function buildDashboardDataFromDeals(dealRows, fileName, targetSalesOverride = currentTargetSales(), mappings = activeMappings()) {
  const asOfDate = localIsoDate();
  const today = makeDateParts(asOfDate);
  const dayMs = 24 * 60 * 60 * 1000;
  const targetSales = targetSalesOverride.map((sale) => ({
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
    pipelineGroupMismatch: { count: 0, amount: 0 },
    pipelineIncluded: { count: 0, amount: 0 },
    pipelineExcluded: { count: 0, amount: 0 },
    pipelineUnmapped: { count: 0, amount: 0 },
  };

  const mapResponsible = (responsible) => {
    const normalized = normalizeName(responsible);
    if (!normalized) return null;
    const mappedSale = mappings.salesGroupMap.get(normalized);
    if (mappedSale) return targetNameToKey.get(normalizeName(mappedSale.displayName)) || makeKey(mappedSale.displayName);
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
    const displayStage = mappedStage(rawStage, mappings);
    const status = normalizeDealStage(displayStage);
    const dealType = mappedDealType(row["Deal Type"], mappings);
    const category = dealCategory(dealType, row);
    const expected = parseDateValue(row["Expected close date"]);
    const stageChanged = parseDateValue(row["Stage change date"]);
    const created = parseDateValue(row.Created);
    const closeDate = stageChanged || expected || created;
    const pipeline = cleanText(row.Pipeline);
    const pipelineGroup = mappings.pipelineGroupMap.get(normalizeName(pipeline)) || "";

    if (!saleByKey.has(saleKey)) {
      const mappedSale = mappings.salesGroupMap.get(normalizeName(responsible));
      saleByKey.set(saleKey, {
        key: saleKey,
        name: mappedSale?.displayName || responsible,
        group: mappedSale?.group || "Unmapped",
        hasTarget: false,
        hasPositiveTarget: false,
        responsibleNames: [],
        monthly: Object.fromEntries(dashboardData.months.map((month) => [month, { renew: 0, new: 0, total: 0 }])),
        targetAnnual: { renew: 0, new: 0, total: 0 },
      });
    }

    const sale = saleByKey.get(saleKey);
    if (!sale.responsibleNames.includes(responsible)) sale.responsibleNames.push(responsible);
    const counting = pipelineCountingResult(sale.group, category, pipeline, mappings);

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
      stage: displayStage,
      matchedStage: displayStage,
      rawStage,
      pipeline,
      pipelineGroup,
      pipelineGroupMatch: !pipelineGroup || pipelineGroup === sale.group,
      countingStatus: counting.status,
      countingIncluded: counting.included,
      countingLabel: counting.label,
      dealType,
      rawDealType: cleanText(row["Deal Type"]),
      product: cleanText(row["Product Type"]) || "(blank)",
      company: cleanText(row.Company),
      contact: cleanText(row.Contact),
      dealName: cleanText(row["Deal Name"]),
      amount: roundMoney(amount),
      forecastAmount: roundMoney(status === "open" ? amount * stageWeightFor(rawStage, displayStage) : 0),
      expectedDate: expected?.date || "",
      expectedMonth: expected?.monthKey || "",
      stageChangeDate: stageChanged?.date || "",
      stageMonth: stageChanged?.monthKey || "",
      createdDate: created?.date || "",
      trackingMonth: status === "open" ? expected?.monthKey || "" : closeDate?.monthKey || "",
      trackingDate: status === "open" ? expected?.date || "" : closeDate?.date || "",
      stageAgeDays: stageChanged ? Math.floor((today.ts - stageChanged.ts) / dayMs) : null,
      risk: {
        overdue: Boolean(expected && expected.ts < today.ts),
        noExpected: !expected,
        stale30: Boolean(stageChanged && stageChanged.ts < today.ts - 30 * dayMs),
        notContacted: displayStage === "Not Contacted" || rawStage === "Not Contacted",
      },
    };
    dealDetails.push(detail);
    if (pipelineGroup && pipelineGroup !== sale.group) addQuality(quality.pipelineGroupMismatch, amount);
    addQuality(
      counting.status === "included"
        ? quality.pipelineIncluded
        : counting.status === "excluded"
          ? quality.pipelineExcluded
          : quality.pipelineUnmapped,
      amount,
    );

    if (!counting.included) continue;

    if (status === "won" || status === "lost") {
      addQuality(status === "won" ? quality.won : quality.lost, amount);
      const monthKey = closeDate?.monthKey || "";
      if ((dashboardData.months || []).includes(monthKey)) {
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
    const notContacted = displayStage === "Not Contacted" || rawStage === "Not Contacted";
    if (noExpected) addQuality(quality.noExpectedOpen, amount);
    if (overdue) addQuality(quality.overdueOpen, amount);
    if (stale30) addQuality(quality.stale30Open, amount);
    if (notContacted) addQuality(quality.notContactedOpen, amount);

    const weight = stageWeightFor(rawStage, displayStage);
    openDeals.push({
      id: cleanText(row.ID),
      saleKey,
      saleName: sale.name,
      group: sale.group,
      responsible,
      category,
      stage: displayStage,
      matchedStage: displayStage,
      rawStage,
      pipeline,
      pipelineGroup,
      pipelineGroupMatch: !pipelineGroup || pipelineGroup === sale.group,
      countingStatus: counting.status,
      countingIncluded: counting.included,
      countingLabel: counting.label,
      dealType,
      rawDealType: cleanText(row["Deal Type"]),
      product: cleanText(row["Product Type"]) || "(blank)",
      company: cleanText(row.Company),
      dealName: cleanText(row["Deal Name"]),
      amount: roundMoney(amount),
      forecastAmount: roundMoney(amount * weight),
      weight,
      expectedDate: expected?.date || "",
      expectedMonth: expected?.monthKey || "",
      trackingDate: expected?.date || "",
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
    mappings: serializeMappings(mappings),
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
  if (lower === "won" || lower === "deal won" || lower === "pre-won") return "won";
  if (lower === "lost" || lower === "deal lost" || lower === "pre-lost") return "lost";
  return "open";
}

function inferDealCategory(row) {
  const pipeline = cleanText(row.Pipeline).toLowerCase();
  const dealType = cleanText(row["Deal Type"]).toLowerCase();
  if (pipeline.includes("renew") || dealType.startsWith("re-new")) return "renew";
  return "new";
}

function mappedStage(stage, mappings = activeMappings()) {
  const raw = cleanText(stage);
  return mappings.stageMap.get(normalizeName(raw)) || raw || "Open";
}

function mappedDealType(dealType, mappings = activeMappings()) {
  const raw = cleanText(dealType);
  return mappings.dealTypeMap.get(normalizeName(raw)) || raw || "New";
}

function dealCategory(dealType, row) {
  const mapped = cleanText(dealType).toLowerCase();
  if (mapped.includes("renew")) return "renew";
  if (mapped.includes("new")) return "new";
  return inferDealCategory(row);
}

function stageWeightFor(rawStage, displayStage) {
  const weights = dashboardData.stageWeights || {};
  return weights[displayStage] ?? weights[rawStage] ?? weights[cleanText(displayStage)] ?? 0.1;
}

function renderKpis(scope) {
  const rows = [
    kpiSummaryRow("ยอดขายรวม", "total", scope.sales, scope.months, scope.actualFacts),
    kpiSummaryRow("Renew", "renew", scope.sales, scope.months, scope.actualFacts),
    kpiSummaryRow("New", "new", scope.sales, scope.months, scope.actualFacts),
  ];

  els.kpiGrid.innerHTML = rows
    .map(
      (row) => `
        <article class="kpi-summary-card ${row.tone}">
          <div class="kpi-summary-head">
            <strong>${escapeHtml(row.label)}</strong>
            <span>${percent(row.achievement)}</span>
          </div>
          <div class="kpi-summary-meta">
            <span>ยอดขาย Won: <b>${compactMoney(row.actual)}</b></span>
            <span>เป้าหมาย: <b>${compactMoney(row.target)}</b></span>
          </div>
          <div class="progress-track kpi-progress-track">
            <div class="progress-fill" style="width:${Math.min(100, row.achievement * 100)}%"></div>
          </div>
          <p class="note">${escapeHtml(periodLabel())}</p>
        </article>
      `,
    )
    .join("");
}

function kpiSummaryRow(label, category, sales, months, actualFacts) {
  const target = sum(sales, (sale) => targetForSale(sale, months, category === "total" ? "all" : category));
  const actual = sum(
    actualFacts.filter((fact) => category === "total" || fact.category === category),
    (fact) => fact.amount,
  );
  const achievement = target ? actual / target : actual > 0 ? 1 : 0;
  return {
    label,
    target,
    actual,
    achievement,
    tone: achievement >= 1 ? "good" : achievement >= 0.7 ? "warn" : "danger",
  };
}

function periodLabel() {
  if (state.periodType === "year") return `Year ${state.periodYear}`;
  if (state.periodType === "quarter") {
    const [year, quarter] = state.periodQuarter.split("-Q");
    return `${ordinalQuarter(Number(quarter))} ${year}`;
  }
  if (state.periodType === "week") return weekLabel(state.periodWeek);
  return `${monthLabel(state.startMonth)} - ${monthLabel(state.endMonth)}`;
}

function metricForMonth(month, scope, category = state.category) {
  const target = sum(scope.sales, (sale) => targetForSale(sale, [month], category));
  const monthDeals = filteredDeals(scope, { category, periodKey: month });
  const actual = sum(
    monthDeals.filter((deal) => deal.status === "won"),
    (deal) => deal.amount,
  );
  const forecast = sum(
    monthDeals.filter((deal) => deal.status === "open"),
    (deal) => deal.forecastAmount,
  );
  return { month, target, actual, forecast };
}

function renderMonthlyChart(scope) {
  const rows = scope.months.map((month) => metricForMonth(month, scope));
  const maxValue = Math.max(1, ...rows.flatMap((row) => [row.target, row.actual, row.forecast]));
  els.monthlyChart.innerHTML = rows
    .map((row) => {
      return `
        <div class="month-row">
          <div class="month-name">${monthLabel(row.month)}</div>
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
      const groupScope = { ...scope, saleSet };
      const actual = sum(
        scope.actualFacts.filter((fact) => itemMatchesSaleScope(fact, groupScope)),
        (fact) => fact.amount,
      );
      const forecast = sum(
        scope.openPeriod.filter((deal) => itemMatchesSaleScope(deal, groupScope)),
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
      const filteredOut = state.category !== "all" && state.category !== category;
      const target = filteredOut ? 0 : sum(scope.sales, (sale) => targetForSale(sale, scope.months, category));
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

function riskScore(deal) {
  return (
    (deal.risk.overdue ? 4 : 0) +
    (deal.risk.noExpected ? 3 : 0) +
    (deal.risk.stale30 ? 2 : 0) +
    (deal.risk.notContacted ? 1 : 0)
  );
}

function dealMatchesSearch(deal) {
  return dealMatchesGlobalSearch(deal);
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

function performanceBucket(deal) {
  if (deal.status === "won") return "won";
  if (deal.status === "lost") return "lost";
  const stage = normalizeName(deal.stage);
  if (stage === "commit") return "commit";
  if (stage === "upside" || stage === "upsde") return "upside";
  return "open";
}

function filteredPerformanceDeals(scope, category = "all") {
  return filteredDeals(scope, { category });
}

function buildPerformanceRows(scope, category) {
  const deals = filteredPerformanceDeals(scope, category);
  const targetCategory = category === "all" ? state.category : category;
  const targetIsFilteredOut = state.category !== "all" && category !== "all" && category !== state.category;
  const rows = scope.sales.map((sale) => ({
    sale,
    target: targetIsFilteredOut ? 0 : targetForSale(sale, scope.months, targetCategory),
    won: 0,
    commit: 0,
    upside: 0,
    open: 0,
    lost: 0,
  }));
  const rowMap = new Map(rows.map((row) => [row.sale.key, row]));
  deals.forEach((deal) => {
    const row = rowMap.get(deal.saleKey) || rows.find((item) => itemMatchesSaleRecord(deal, item.sale));
    if (!row) return;
    row[performanceBucket(deal)] += deal.amount;
  });
  return rows
    .filter((row) => row.target > 0 || ["won", "commit", "upside", "open", "lost"].some((key) => row[key] > 0))
    .sort((a, b) => b.target - a.target || b.won - a.won || a.sale.name.localeCompare(b.sale.name, "th"));
}

function renderSalesPerformanceCharts(scope) {
  renderStackedPerformanceChart(els.totalSalesChart, buildPerformanceRows(scope, "all"), "total");
  renderStackedPerformanceChart(els.renewSalesChart, buildPerformanceRows(scope, "renew"), "renew");
  renderStackedPerformanceChart(els.newSalesChart, buildPerformanceRows(scope, "new"), "new");
  renderCumulativePerformanceChart(scope);
}

function renderStackedPerformanceChart(container, rows, variant) {
  if (!rows.length) {
    container.innerHTML = `<div class="empty">ไม่มีข้อมูลสำหรับกราฟนี้</div>`;
    return;
  }
  const visibleKeys = ["won", "commit", "upside", "open", "lost"].filter((key) => state.statusFilters[key]);
  const maxValue = Math.max(
    1,
    ...rows.map((row) => Math.max(row.target, sum(visibleKeys, (key) => row[key]))),
  );
  container.innerHTML = `
    <div class="chart-legend-inline">
      <span class="legend-target">Target</span>
      ${visibleKeys.map((key) => `<span class="legend-${key}">${escapeHtml(statusLabel(key))}</span>`).join("")}
    </div>
    <div class="performance-row-list ${variant}">
      ${rows
        .slice(0, 24)
        .map((row) => {
          const actualTotal = sum(visibleKeys, (key) => row[key]);
          const achievement = row.target ? row.won / row.target : row.won > 0 ? 1 : 0;
          return `
            <div class="performance-row">
              <div class="performance-label">
                <strong>${escapeHtml(row.sale.name)}</strong>
                <span>${percent(achievement)} Won / Target</span>
              </div>
              <div class="performance-bars">
                <div class="target-outline" style="width:${(row.target / maxValue) * 100}%"></div>
                <div class="stage-stack">
                  ${visibleKeys
                    .map(
                      (key) => `
                        <span
                          class="stage-segment stage-${key}"
                          title="${escapeHtml(statusLabel(key))}: ${money(row[key])}"
                          style="width:${(row[key] / maxValue) * 100}%"
                        ></span>
                      `,
                    )
                    .join("")}
                </div>
              </div>
              <div class="performance-values">
                <strong>${compactMoney(actualTotal)}</strong>
                <span>Target ${compactMoney(row.target)}</span>
              </div>
            </div>
          `;
        })
        .join("")}
    </div>
  `;
}

function statusLabel(key) {
  return {
    won: "Won",
    commit: "Commit",
    upside: "Upside",
    open: "Open",
    lost: "Lost",
  }[key] || key;
}

function renderCumulativePerformanceChart(scope) {
  const months = scope.months;
  if (!months.length) {
    els.cumulativeSalesChart.innerHTML = `<div class="empty">ไม่มีช่วงเวลาสำหรับแสดงกราฟ</div>`;
    return;
  }
  const visibleKeys = ["won", "commit", "upside", "open", "lost"].filter((key) => state.statusFilters[key]);
  const monthlyDeals = filteredPerformanceDeals(scope, "all");
  const monthly = months.map((month) => ({
    month,
    target: sum(scope.sales, (sale) => targetForSale(sale, [month], state.category)),
    won: 0,
    commit: 0,
    upside: 0,
    open: 0,
    lost: 0,
  }));
  const monthlyMap = new Map(monthly.map((item) => [item.month, item]));
  monthlyDeals.forEach((deal) => {
    const item = monthlyMap.get(dealPeriodKey(deal, scope.months[0]));
    if (!item) return;
    item[performanceBucket(deal)] += deal.amount;
  });
  let targetRunning = 0;
  const statusRunning = Object.fromEntries(visibleKeys.map((key) => [key, 0]));
  const cumulative = monthly.map((item) => {
    targetRunning += item.target;
    const result = { month: item.month, target: targetRunning };
    visibleKeys.forEach((key) => {
      statusRunning[key] += item[key];
      result[key] = statusRunning[key];
    });
    return result;
  });
  const maxValue = Math.max(
    1,
    ...cumulative.flatMap((item) => [item.target, sum(visibleKeys, (key) => item[key])]),
  );
  const width = Math.max(760, months.length * 72);
  const height = 320;
  const chartHeight = 240;
  const left = 46;
  const bottom = 274;
  const plotWidth = width - left - 22;
  const barWidth = Math.min(28, Math.max(16, plotWidth / Math.max(months.length * 2, 1)));
  const xAt = (index) => left + ((index + 0.5) / months.length) * plotWidth;
  const yAt = (value) => bottom - (value / maxValue) * chartHeight;
  const targetPoints = cumulative.map((item, index) => `${xAt(index)},${yAt(item.target)}`).join(" ");
  const bars = cumulative
    .map((item, index) => {
      let used = 0;
      const parts = visibleKeys
        .map((key) => {
          const current = item[key] || 0;
          const y = yAt(used + current);
          const heightValue = (current / maxValue) * chartHeight;
          used += current;
          return `<rect class="cum-${key}" x="${xAt(index) - barWidth / 2}" y="${y}" width="${barWidth}" height="${Math.max(0, heightValue)}" rx="2"></rect>`;
        })
        .join("");
      return `${parts}<text x="${xAt(index)}" y="${bottom + 20}" text-anchor="middle">${escapeHtml(item.month)}</text>`;
    })
    .join("");
  const guides = [0, 0.25, 0.5, 0.75, 1]
    .map((ratio) => {
      const value = maxValue * ratio;
      const y = yAt(value);
      return `<g><line x1="${left}" y1="${y}" x2="${width - 22}" y2="${y}" class="chart-guide"></line><text x="${left - 8}" y="${y + 4}" text-anchor="end">${escapeHtml(compactMoney(value))}</text></g>`;
    })
    .join("");
  els.cumulativeSalesChart.innerHTML = `
    <div class="chart-legend-inline">
      <span class="legend-target-line">Target สะสม</span>
      ${visibleKeys.map((key) => `<span class="legend-${key}">${escapeHtml(statusLabel(key))}</span>`).join("")}
    </div>
    <div class="svg-chart-wrap">
      <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Cumulative sales chart">
        ${guides}
        ${bars}
        <polyline class="target-line" points="${targetPoints}"></polyline>
      </svg>
    </div>
  `;
}

function renderTransactionTable(scope) {
  const andTerms = els.transactionAndInputs.map((input) => normalizeSearch(input.value)).filter(Boolean);
  const notTerms = els.transactionNotInputs.map((input) => normalizeSearch(input.value)).filter(Boolean);
  const globalTerm = normalizeSearch(state.search);
  const rows = filteredPerformanceDeals(scope, "all")
    .filter((deal) => categoryMatches(deal.category))
    .filter((deal) => state.statusFilters[performanceBucket(deal)])
    .filter((deal) => {
      const text = transactionSearchText(deal);
      const tokens = transactionSearchTokens(deal);
      if (globalTerm && !transactionTermMatches(globalTerm, text, tokens)) return false;
      if (andTerms.some((term) => !transactionTermMatches(term, text, tokens))) return false;
      if (notTerms.some((term) => transactionTermMatches(term, text, tokens))) return false;
      return true;
    });
  sortTransactionRows(rows);
  els.transactionTotalAmount.textContent = money(sum(rows, (row) => row.amount));
  els.transactionTable.innerHTML = transactionTable(rows.slice(0, 500));
}

function transactionSearchText(deal) {
  const expectedCloseDate = displayDate(deal.expectedDate);
  return normalizeSearch(
    `${deal.id} ${deal.saleName} ${deal.responsible} ${deal.company} ${deal.dealName} ${deal.pipeline} ${deal.stage} ${deal.dealType || ""} ${deal.product} ${deal.category} ${performanceBucket(deal)} ${deal.expectedDate || ""} ${expectedCloseDate}`,
  );
}

function transactionSearchTokens(deal) {
  const expectedCloseDate = displayDate(deal.expectedDate);
  return new Set(
    [
      deal.id,
      deal.saleName,
      deal.responsible,
      deal.company,
      deal.dealName,
      deal.pipeline,
      deal.stage,
      deal.dealType,
      deal.product,
      deal.category,
      deal.expectedDate,
      expectedCloseDate,
      performanceBucket(deal),
    ]
      .map(normalizeSearch)
      .flatMap((value) => value.split(/[^\p{L}\p{N}]+/u))
      .filter(Boolean),
  );
}

function transactionTermMatches(term, text, tokens) {
  if (term === "new" || term === "renew") return tokens.has(term);
  return text.includes(term);
}

function sortTransactionRows(rows) {
  const { key, dir } = state.transactionSort;
  const factor = dir === "asc" ? 1 : -1;
  rows.sort((a, b) => {
    if (key === "amount") return factor * ((a.amount || 0) - (b.amount || 0));
    if (key === "expectedDate") return factor * ((parseDateValue(a.expectedDate)?.ts || 0) - (parseDateValue(b.expectedDate)?.ts || 0));
    return factor * String(a[key] || "").localeCompare(String(b[key] || ""), "th");
  });
}

function transactionTable(rows) {
  if (!rows.length) return `<div class="empty">ไม่มี Transaction ที่ตรงกับเงื่อนไข</div>`;
  return `
    <table class="transaction-table">
      <thead>
        <tr>
          ${transactionHeader("id", "ID", false, "tx-id-col")}
          ${transactionHeader("saleName", "Sale", false, "tx-sale-col")}
          ${transactionHeader("company", "Company / Deal", false, "tx-company-col")}
          ${transactionHeader("category", "Type", false, "tx-type-col")}
          ${transactionHeader("stage", "Matched Stage", false, "tx-stage-col")}
          ${transactionHeader("expectedDate", "Expected", false, "tx-expected-col")}
          ${transactionHeader("amount", "Amount", true, "tx-amount-col")}
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (deal) => `
              <tr class="clickable-row" data-deal-key="${escapeHtml(dealDetailKey(deal))}">
                <td class="tx-id-col">${escapeHtml(deal.id || "-")}</td>
                <td class="tx-sale-col"><strong>${escapeHtml(deal.saleName)}</strong><br><span class="muted">${escapeHtml(deal.group)}</span></td>
                <td class="company-cell tx-company-col"><strong>${escapeHtml(deal.company || "-")}</strong><br><span class="muted">${escapeHtml(deal.dealName || "-")}</span></td>
                <td class="tx-type-col">${escapeHtml(deal.category)}</td>
                <td class="tx-stage-col"><strong>${escapeHtml(deal.stage)}</strong><br><span class="muted">DB: ${escapeHtml(deal.rawStage || deal.stage)}</span><br>${stageBucketBadge(performanceBucket(deal))}</td>
                <td class="tx-expected-col">${escapeHtml(displayDate(deal.expectedDate))}</td>
                <td class="num tx-amount-col">${money(deal.amount)}</td>
              </tr>
            `,
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function transactionHeader(key, label, numeric = false, extraClass = "") {
  const active = state.transactionSort.key === key;
  const icon = active ? (state.transactionSort.dir === "asc" ? "↑" : "↓") : "↕";
  const className = [numeric ? "num" : "", extraClass].filter(Boolean).join(" ");
  return `<th class="${className}"><button type="button" class="sort-button" data-transaction-sort="${key}">${escapeHtml(label)} ${icon}</button></th>`;
}

function renderDealDetailTable(scope) {
  const rows = filteredDeals(scope, { countingIncluded: false });
  sortDealDetailRows(rows);
  const visibleRows = rows.slice(0, 300);

  els.dealDetailTable.innerHTML = dealDetailTable(visibleRows, "ไม่มี deal detail ในเดือนที่เลือก");
}

function sortDealDetailRows(rows) {
  const { key, dir } = state.dealDetailSort;
  const factor = dir === "asc" ? 1 : -1;
  const statusOrder = { open: 1, won: 2, lost: 3 };
  const riskOrder = (deal) => riskScore(deal);
  const valueFor = (deal) => {
    if (key === "companyDeal") return `${deal.company || ""} ${deal.dealName || ""}`;
    if (key === "status") return statusOrder[deal.status] || 9;
    if (key === "risk") return deal.status === "open" ? riskOrder(deal) : deal.status === "won" ? -1 : -2;
    return deal[key];
  };
  rows.sort((a, b) => {
    const aValue = valueFor(a);
    const bValue = valueFor(b);
    let result;
    if (["amount", "forecastAmount", "stageAgeDays", "status", "risk"].includes(key)) {
      result = (Number(aValue) || 0) - (Number(bValue) || 0);
    } else {
      result = String(aValue || "").localeCompare(String(bValue || ""), "th", { numeric: true });
    }
    if (result === 0) result = (b.amount || 0) - (a.amount || 0);
    return factor * result;
  });
}

function dealDetailTable(rows, emptyText) {
  if (!rows.length) return `<div class="empty">${escapeHtml(emptyText)}</div>`;
  return `
    <table>
      <thead>
        <tr>
          ${dealDetailHeader("id", "ID")}
          ${dealDetailHeader("saleName", "Sale")}
          ${dealDetailHeader("companyDeal", "Company / Deal", false, "company-cell")}
          ${dealDetailHeader("status", "Status")}
          ${dealDetailHeader("stage", "Matched Stage")}
          ${dealDetailHeader("pipeline", "Pipeline")}
          ${dealDetailHeader("countingStatus", "Counting")}
          ${dealDetailHeader("product", "Product")}
          ${dealDetailHeader("amount", "Amount", true)}
          ${dealDetailHeader("forecastAmount", "Forecast", true)}
          ${dealDetailHeader("expectedDate", "Expected")}
          ${dealDetailHeader("stageChangeDate", "Stage Change")}
          ${dealDetailHeader("risk", "Risk / Progress")}
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (deal) => `
              <tr class="clickable-row" data-deal-key="${escapeHtml(dealDetailKey(deal))}">
                <td>${escapeHtml(deal.id || "-")}</td>
                <td><strong>${escapeHtml(deal.saleName)}</strong><br><span class="muted">${escapeHtml(deal.responsible)}</span></td>
                <td class="company-cell"><strong>${escapeHtml(deal.company || "-")}</strong><br><span class="muted">${escapeHtml(deal.dealName || "-")}</span></td>
                <td>${statusBadge(deal.status)}</td>
                <td><strong>${escapeHtml(deal.stage)}</strong><br><span class="muted">DB: ${escapeHtml(deal.rawStage || deal.stage)}</span></td>
                <td>${escapeHtml(deal.pipeline)}${deal.pipelineGroup && !deal.pipelineGroupMatch ? `<br><span class="badge warn">${escapeHtml(deal.pipelineGroup)}</span>` : ""}</td>
                <td>${countingBadge(deal)}</td>
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

function dealDetailHeader(key, label, numeric = false, extraClass = "") {
  const active = state.dealDetailSort.key === key;
  const icon = active ? (state.dealDetailSort.dir === "asc" ? "↑" : "↓") : "↕";
  const className = [numeric ? "num" : "", extraClass].filter(Boolean).join(" ");
  return `<th class="${className}"><button type="button" class="sort-button" data-deal-detail-sort="${key}">${escapeHtml(label)} ${icon}</button></th>`;
}

function leadCreatedWeek(deal) {
  return isoWeekKeyFromDate(deal.createdDate);
}

function filteredLeadDeals() {
  const sales = visibleSales();
  const saleSet = new Set(sales.map((sale) => sale.key));
  const selectedSale = selectedSaleRecord();
  const scope = { saleSet, selectedSale, months: [], monthSet: new Set() };
  const searchTerms = [state.search, state.leadSearch].map(normalizeSearch).filter(Boolean);
  return filteredDeals(scope, { period: false, globalSearch: false })
    .map((deal) => ({ ...deal, createdWeek: leadCreatedWeek(deal) }))
    .filter((deal) => deal.createdWeek && deal.createdWeek >= state.leadStartWeek && deal.createdWeek <= state.leadEndWeek)
    .filter((deal) => {
      if (!searchTerms.length) return true;
      const text = dealSearchText(deal);
      return searchTerms.every((term) => text.includes(term));
    })
    .sort((a, b) => a.createdWeek.localeCompare(b.createdWeek) || b.amount - a.amount);
}

function renderNewLead() {
  const deals = filteredLeadDeals();
  renderNewLeadSummary(deals);
  renderNewLeadWeeklyChart(deals);
  renderNewLeadTable(deals);
}

function renderNewLeadSummary(deals) {
  const weekCount = new Set(deals.map((deal) => deal.createdWeek)).size;
  const saleCount = new Set(deals.map((deal) => deal.saleKey)).size;
  const amount = sum(deals, (deal) => deal.amount);
  const avgPerWeek = weekCount ? deals.length / weekCount : 0;
  const cards = [
    ["Total Leads", deals.length.toLocaleString("th-TH"), `${weekLabel(state.leadStartWeek)} ถึง ${weekLabel(state.leadEndWeek)}`],
    ["Pipeline Amount", compactMoney(amount), "มูลค่า Lead ที่สร้างในช่วงนี้"],
    ["Active Sales", saleCount.toLocaleString("th-TH"), "จำนวน Sale ที่สร้าง Lead"],
    ["Avg Leads / Week", avgPerWeek.toFixed(1), "ใช้ดู consistency ในการหา lead"],
  ];
  els.newLeadSummary.innerHTML = cards
    .map(
      ([label, value, note]) => `
        <div class="lead-summary-card">
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(value)}</strong>
          <small>${escapeHtml(note)}</small>
        </div>
      `,
    )
    .join("");
}

function renderNewLeadWeeklyChart(deals) {
  const weeks = leadWeekOptions()
    .map((week) => week.value)
    .filter((week) => week >= state.leadStartWeek && week <= state.leadEndWeek);
  const rows = weeks.map((week) => ({ week, count: 0, amount: 0 }));
  const rowMap = new Map(rows.map((row) => [row.week, row]));
  deals.forEach((deal) => {
    const row = rowMap.get(deal.createdWeek);
    if (!row) return;
    row.count += 1;
    row.amount += deal.amount;
  });
  const maxCount = Math.max(1, ...rows.map((row) => row.count));
  const maxAmount = Math.max(1, ...rows.map((row) => row.amount));
  els.newLeadWeeklyChart.innerHTML = rows.length
    ? rows
        .map(
          (row) => `
            <div class="lead-week-row">
              <div class="lead-week-name">${escapeHtml(row.week)}</div>
              <div class="bar-group">
                <div class="bar-track" title="${row.count.toLocaleString("th-TH")} leads">
                  <div class="bar actual" style="width:${(row.count / maxCount) * 100}%"></div>
                </div>
                <div class="bar-track" title="${money(row.amount)}">
                  <div class="bar forecast" style="width:${(row.amount / maxAmount) * 100}%"></div>
                </div>
              </div>
              <div class="lead-week-value">${row.count.toLocaleString("th-TH")} leads<br><span class="muted">${compactMoney(row.amount)}</span></div>
            </div>
          `,
        )
        .join("")
    : `<div class="empty">ไม่มี Lead ในช่วง Week ที่เลือก</div>`;
}

function renderNewLeadTable(deals) {
  if (!deals.length) {
    els.newLeadTable.innerHTML = `<div class="empty">ไม่มี Lead ในช่วง Week ที่เลือก</div>`;
    return;
  }
  const grouped = new Map();
  deals.forEach((deal) => {
    if (!grouped.has(deal.createdWeek)) grouped.set(deal.createdWeek, []);
    grouped.get(deal.createdWeek).push(deal);
  });
  const rowsHtml = Array.from(grouped.entries())
    .map(([week, weekDeals]) => {
      const totalAmount = sum(weekDeals, (deal) => deal.amount);
      const groupHeader = `
        <tr class="week-group-row">
          <td colspan="8"><strong>${escapeHtml(weekLabel(week))}</strong><span>${weekDeals.length.toLocaleString("th-TH")} leads · ${compactMoney(totalAmount)}</span></td>
        </tr>
      `;
      const detailRows = weekDeals
        .sort((a, b) => a.saleName.localeCompare(b.saleName, "th") || b.amount - a.amount)
        .slice(0, 120)
        .map(
          (deal) => `
            <tr class="clickable-row" data-deal-key="${escapeHtml(dealDetailKey(deal))}">
              <td>${escapeHtml(deal.createdDate || "-")}</td>
              <td><strong>${escapeHtml(deal.saleName)}</strong><br><span class="muted">${escapeHtml(deal.group)}</span></td>
              <td class="company-cell"><strong>${escapeHtml(deal.company || "-")}</strong><br><span class="muted">${escapeHtml(deal.dealName || `ID ${deal.id}`)}</span></td>
              <td>${escapeHtml(deal.category)}</td>
              <td><strong>${escapeHtml(deal.stage)}</strong><br><span class="muted">DB: ${escapeHtml(deal.rawStage || deal.stage)}</span></td>
              <td>${escapeHtml(deal.pipeline || "-")}</td>
              <td class="num">${money(deal.amount)}</td>
              <td>${deal.status === "open" ? riskBadges(deal) : progressBadge(deal.status)}</td>
            </tr>
          `,
        )
        .join("");
      return groupHeader + detailRows;
    })
    .join("");
  els.newLeadTable.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Created</th>
          <th>Sale</th>
          <th class="company-cell">Company / Deal</th>
          <th>Type</th>
          <th>Matched Stage</th>
          <th>Pipeline</th>
          <th class="num">Amount</th>
          <th>Risk / Progress</th>
        </tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  `;
}

function dealDetailKey(deal) {
  return [deal.id, deal.saleKey, deal.trackingMonth || deal.expectedMonth, deal.amount, deal.company]
    .map((value) => cleanText(value))
    .join("||");
}

function openDealModal(key) {
  const deal = allDealDetails().find((item) => dealDetailKey(item) === key);
  if (!deal) return;
  els.dealModalTitle.textContent = deal.company || deal.dealName || `Deal ${deal.id || ""}`.trim();
  els.dealModalSubtitle.textContent = `${deal.dealName || "No deal name"}${deal.id ? ` | ID ${deal.id}` : ""}`;
  els.dealModalContent.innerHTML = dealModalHtml(deal);
  els.dealModal.hidden = false;
  els.closeDealModal.focus();
}

function dealModalHtml(deal) {
  const bucket = performanceBucket(deal);
  const risk = deal.status === "open" ? riskScore(deal) : 0;
  return `
    <div class="deal-modal-summary">
      ${dealInsightCard("Amount", money(deal.amount), deal.category === "renew" ? "Renew revenue" : "New revenue")}
      ${dealInsightCard("Forecast", money(deal.forecastAmount || 0), bucket === "won" ? "Actual counted" : statusLabel(bucket))}
      ${dealInsightCard("Timeline", deal.trackingMonth ? monthLabel(deal.trackingMonth) : "-", "เดือนที่ใช้ติดตาม")}
      ${dealInsightCard("Risk Score", String(risk), risk ? "ต้องติดตามต่อ" : "ไม่มี risk flag หลัก")}
    </div>

    <div class="deal-modal-grid">
      <section class="deal-section">
        <h3>Commercial View</h3>
        ${detailLine("Status", `${statusBadge(deal.status)} ${escapeHtml(statusLabel(bucket))}`, true)}
        ${detailLine("Deal Type", `${escapeHtml(deal.category)}${deal.dealType ? ` / ${escapeHtml(deal.dealType)}` : ""}`, true)}
        ${detailLine("Sale Owner", `${escapeHtml(deal.saleName)} <span class="muted">(${escapeHtml(deal.group)})</span>`, true)}
        ${detailLine("Responsible", deal.responsible || "-")}
        ${detailLine("Product", deal.product || "-")}
      </section>

      <section class="deal-section">
        <h3>Pipeline & Mapping</h3>
        ${detailLine("Pipeline", `${escapeHtml(deal.pipeline || "-")}${deal.pipelineGroup ? ` <span class="muted">expected ${escapeHtml(deal.pipelineGroup)}</span>` : ""}`, true)}
        ${detailLine("Pipeline Fit", deal.pipelineGroup && !deal.pipelineGroupMatch ? `<span class="badge warn">Sale Group mismatch</span>` : `<span class="badge good">Matched / No rule</span>`, true)}
        ${detailLine("Counting Rule", countingBadge(deal), true)}
        ${detailLine("Matched Stage", escapeHtml(deal.stage || "-"), true)}
        ${detailLine("Raw Stage", deal.rawStage || deal.stage || "-")}
        ${detailLine("Raw Deal Type", deal.rawDealType || deal.dealType || "-")}
      </section>

      <section class="deal-section">
        <h3>Timing & Progress</h3>
        ${detailLine("Expected Close", deal.expectedDate || "-")}
        ${detailLine("Stage Change", deal.stageChangeDate || "-")}
        ${detailLine("Stage Age", deal.stageAgeDays == null ? "-" : `${deal.stageAgeDays.toLocaleString("th-TH")} days`)}
        ${detailLine("Created", deal.createdDate || "-")}
        ${detailLine("Tracking Month", deal.trackingMonth ? monthLabel(deal.trackingMonth) : "-")}
      </section>

      <section class="deal-section">
        <h3>Risk & Next Analysis</h3>
        ${detailLine("Risk Flags", deal.status === "open" ? riskBadges(deal) : progressBadge(deal.status), true)}
        <div class="analysis-list">
          ${dealRecommendations(deal).map((item) => `<div>${escapeHtml(item)}</div>`).join("")}
        </div>
      </section>
    </div>

    <section class="deal-section full">
      <h3>Customer / Deal Context</h3>
      ${detailLine("Company", deal.company || "-")}
      ${detailLine("Deal Name", deal.dealName || "-")}
      ${detailLine("Contact", deal.contact || "-")}
    </section>
  `;
}

function dealInsightCard(label, value, note) {
  return `
    <div class="deal-insight">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <small>${escapeHtml(note)}</small>
    </div>
  `;
}

function detailLine(label, value, html = false) {
  return `
    <div class="detail-line">
      <span>${escapeHtml(label)}</span>
      <strong>${html ? value : escapeHtml(value || "-")}</strong>
    </div>
  `;
}

function dealRecommendations(deal) {
  const items = [];
  if (deal.status === "won") items.push("ใช้เป็น Actual ในเดือนปิดการขาย และตรวจสอบว่าถูกจัดเข้า New/Renew ถูกต้อง");
  if (deal.status === "lost") items.push("วิเคราะห์เหตุผลการ Lost เพิ่มเติม เพื่อแยกปัญหาด้านราคา, timing หรือ product fit");
  if (deal.status === "open") {
    if (deal.risk?.overdue) items.push("Expected close date เลยกำหนดแล้ว ควรอัปเดตเดือนปิดหรือ action ถัดไป");
    if (deal.risk?.noExpected) items.push("ยังไม่มี Expected close date ทำให้ forecast รายเดือนไม่แม่น ควรเติมวันที่คาดว่าจะปิด");
    if (deal.risk?.stale30) items.push("Stage ไม่ขยับเกิน 30 วัน ควร review next step กับ Sale Owner");
    if (deal.risk?.notContacted) items.push("ยังไม่ติดต่อ ควรจัดลำดับความสำคัญก่อนนับเป็น pipeline ที่มีน้ำหนัก");
    if (!items.length) items.push("ไม่มี risk flag หลัก ใช้ติดตามจำนวนเงิน, stage และ expected close ตามรอบ pipeline review");
  }
  if (deal.pipelineGroup && !deal.pipelineGroupMatch) items.push("Pipeline ไม่ตรงกับ Sale Group ตาม Data Mapping ควรตรวจ owner หรือ pipeline");
  if (deal.category === "renew") items.push("เป็น Renew ควรตรวจวันหมดสัญญา/รอบต่ออายุ เพื่อกันหลุดเป้า recurring revenue");
  if (deal.category === "new") items.push("เป็น New ควรดู source, product fit และ stage conversion เพื่อใช้ปรับแผนหา pipeline เพิ่ม");
  return items;
}

function statusBadge(status) {
  if (status === "won") return `<span class="badge good">Won</span>`;
  if (status === "lost") return `<span class="badge danger">Lost</span>`;
  return `<span class="badge info">Open</span>`;
}

function stageBucketBadge(bucket) {
  const className = {
    won: "good",
    lost: "danger",
    commit: "warn",
    upside: "info",
    open: "info",
  }[bucket] || "info";
  return `<span class="badge ${className}">${escapeHtml(statusLabel(bucket))}</span>`;
}

function countingBadge(deal) {
  if (deal.countingStatus === "excluded") return `<span class="badge danger">Excluded</span>`;
  if (deal.countingStatus === "unmapped") return `<span class="badge warn">Unmapped</span>`;
  return `<span class="badge good">Included</span>`;
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
          <th>Matched Stage</th>
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
              <tr class="clickable-row" data-deal-key="${escapeHtml(dealDetailKey(deal))}">
                <td><strong>${escapeHtml(deal.saleName)}</strong><br><span class="muted">${escapeHtml(deal.responsible)}</span></td>
                <td class="company-cell"><strong>${escapeHtml(deal.company || "-")}</strong><br><span class="muted">${escapeHtml(deal.dealName || `ID ${deal.id}`)}</span></td>
                <td><strong>${escapeHtml(deal.stage)}</strong><br><span class="muted">DB: ${escapeHtml(deal.rawStage || deal.stage)}</span></td>
                <td>${escapeHtml(deal.pipeline)}${deal.pipelineGroup && !deal.pipelineGroupMatch ? `<br><span class="badge warn">${escapeHtml(deal.pipelineGroup)}</span>` : ""}</td>
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
  const mappingCounts = dashboardData.mappings?.counts || { pipelines: 0, stages: 0, dealTypes: 0, sales: 0 };
  const cards = [
    [
      "Data Mapping loaded",
      { count: mappingCounts.pipelines + (mappingCounts.pipelineRules || 0) + mappingCounts.stages + mappingCounts.dealTypes + mappingCounts.sales, amount: 0 },
    ],
    ["Pipeline rules included", q.pipelineIncluded || { count: 0, amount: 0 }],
    ["Pipeline rules excluded", q.pipelineExcluded || { count: 0, amount: 0 }],
    ["Pipeline rules unmapped", q.pipelineUnmapped || { count: 0, amount: 0 }],
    ["Pre-WON counted as Won", q.preWonAsWon],
    ["Pre-LOST counted as Lost", q.preLostAsLost],
    ["Won total after normalization", q.won],
    ["Lost total after normalization", q.lost],
    ["Open pipeline after normalization", q.open],
    ["Pipeline / Sale Group mismatch", q.pipelineGroupMismatch || { count: 0, amount: 0 }],
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
  if (state.tab === "sales") {
    renderSalesPerformanceCharts(scope);
    renderTransactionTable(scope);
  }
  if (state.tab === "deals") renderDealDetailTable(scope);
  if (state.tab === "leads") renderNewLead();
  if (state.tab === "pipeline") {
    renderRiskSummary(scope);
    renderRiskByStage(scope);
    renderRiskDealsTable(scope);
  }
  if (state.tab === "data") renderDataQuality();
}

function render() {
  const scope = calcScope();
  els.emptyDataBanner.hidden = hasDealData();
  renderMeta();
  renderKpis(scope);
  renderMonthlyChart(scope);
  renderTeamChart(scope);
  renderCategoryMix(scope);
  renderStatusMix(scope);
  renderTopDealsTable(scope);
  renderDealDetailTable(scope);
  renderNewLead();
  renderSalesPerformanceCharts(scope);
  renderTransactionTable(scope);
  renderRiskSummary(scope);
  renderRiskByStage(scope);
  renderRiskDealsTable(scope);
  renderDataQuality();
}

initFilters();
render();
autoLoadDefaultSetupFiles();
