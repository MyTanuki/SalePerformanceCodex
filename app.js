const STORED_DATA_KEY = "sale-performance-dashboard-workspace-v2";
const THEME_KEY = "sale-performance-dashboard-theme";
const initialDashboardData = JSON.parse(JSON.stringify(window.DASHBOARD_DATA || createEmptyDashboardData()));
const storedDashboardData = loadStoredDashboardData();
let dashboardData = storedDashboardData || JSON.parse(JSON.stringify(initialDashboardData));

const state = {
  viewMode: "group",
  group: "all",
  sale: "all",
  periodType: "year",
  periodYear: currentYearKey(),
  periodQuarter: currentQuarterKey(),
  periodMonth: currentMonthKey(),
  startMonth: currentYearStartMonthKey(),
  endMonth: currentMonthKey(),
  leadStartWeek: previousIsoWeekKey(localIsoDate(), 3),
  leadEndWeek: isoWeekKeyFromDate(localIsoDate()),
  leadSearch: "",
  leadCategory: "all",
  category: "all",
  showUnmapped: false,
  revenueExpectedFallback: true,
  revenueSubtab: "performance",
  invoiceDateMode: "posting",
  revenueTargetVisible: false,
  revenueDistributionVisible: false,
  revenueDetailVisible: false,
  targetContributionVisible: false,
  revenueDataChecksVisible: false,
  search: "",
  tab: "overview",
  statusFilters: {
    won: true,
    commit: false,
    upside: false,
    open: false,
    lost: false,
  },
  transactionSort: {
    key: "amount",
    dir: "desc",
  },
  dealDetailSort: {
    key: "status",
    dir: "asc",
  },
  columnFilters: {
    transaction: {},
    allDeals: {},
    revenue: {},
    leads: {},
  },
  revenueTargetExpanded: {},
  revenueDistributionExpanded: {},
  invoiceAnalysisExpanded: {},
};

const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const saleTypes = ["renew", "new"];
const tableFilterSources = {
  transaction: [],
  allDeals: [],
  revenue: [],
  leads: [],
};
const tabDescriptions = {
  overview: "แสดงภาพรวมยอดขาย โครงสร้าง New/Renew และ Top Open Opportunities เฉพาะ Deals ที่กำหนดให้นับยอดขาย",
  sales: "แสดงเฉพาะ Deals ที่กำหนดให้นับยอดขาย",
  revenue: "สรุปกระแสรายได้รายเดือนจาก Deals Won / Pre-WON ที่กำหนดให้นับยอดขาย โดยใช้ช่วงสัญญาเพื่อสนับสนุนฝ่ายบัญชี การเงิน และผู้บริหาร",
  deals: "แสดงรายละเอียดทุก Deal ในช่วงเวลาที่เลือก รวมทั้ง Deals ที่ถูกนับและไม่นับยอดขาย เพื่อใช้ตรวจสอบความคืบหน้า",
  leads: "แสดง Leads จาก Created Date ตาม ISO Week โดยใช้เฉพาะ Filter ในหน้า New Deals",
  pipeline: "แสดงเฉพาะ Open Deals ที่กำหนดให้นับยอดขายและมีสัญญาณความเสี่ยงที่ควรติดตาม",
  data: "แสดงคุณภาพข้อมูล Mapping และข้อสมมติฐานที่ใช้คำนวณ Dashboard",
};
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
        invoice: "",
      },
      assumptions: [
        "Web starts empty. Upload Sale Target and Deal CSV to build the dashboard.",
        "Stage Mapping and Sale Target can be auto-loaded from ./data.",
        "Revenue Analysis uses Won / Pre-WON deals that pass counting rules. One-time/OTC deals are recognized in the contract start month; recurring deals are spread by Contract Start/End, Contract Period, or MRR fallback.",
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
    invoiceRows: [],
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
  themeToggleButton: null,
  settingsModal: document.querySelector("#settingsModal"),
  closeSettings: document.querySelector("#closeSettings"),
  fieldInfoModal: document.querySelector("#fieldInfoModal"),
  closeFieldInfo: document.querySelector("#closeFieldInfo"),
  fieldInfoContent: document.querySelector("#fieldInfoContent"),
  dealModal: document.querySelector("#dealModal"),
  closeDealModal: document.querySelector("#closeDealModal"),
  dealModalTitle: document.querySelector("#dealModalTitle"),
  dealModalSubtitle: document.querySelector("#dealModalSubtitle"),
  dealModalContent: document.querySelector("#dealModalContent"),
  targetCsvInput: document.querySelector("#targetCsvInput"),
  dealCsvInput: document.querySelector("#dealCsvInput"),
  mappingCsvInput: document.querySelector("#mappingCsvInput"),
  invoiceCsvInput: document.querySelector("#invoiceCsvInput"),
  targetAutoFileName: document.querySelector("#targetAutoFileName"),
  mappingAutoFileName: document.querySelector("#mappingAutoFileName"),
  applyCsvFiles: document.querySelector("#applyCsvFiles"),
  clearCsvFiles: document.querySelector("#clearCsvFiles"),
  uploadStatus: document.querySelector("#uploadStatus"),
  emptyDataBanner: document.querySelector("#emptyDataBanner"),
  dashboardFilterShell: document.querySelector(".dashboard-filter-shell"),
  periodFilterCard: document.querySelector(".period-filter-card"),
  leadToolbarCard: document.querySelector(".lead-toolbar-card"),
  savePipelineMatching: document.querySelector("#savePipelineMatching"),
  clearPipelineMatching: document.querySelector("#clearPipelineMatching"),
  pipelineMatchingSummary: document.querySelector("#pipelineMatchingSummary"),
  pipelineMatchingMatrix: document.querySelector("#pipelineMatchingMatrix"),
  viewMode: document.querySelector("#viewMode"),
  groupFilterContainer: document.querySelector("#groupFilterContainer"),
  salesFilterContainer: document.querySelector("#salesFilterContainer"),
  periodTypeButtons: Array.from(document.querySelectorAll("[data-period-type]")),
  periodYear: document.querySelector("#periodYear"),
  periodQuarter: document.querySelector("#periodQuarter"),
  periodMonth: document.querySelector("#periodMonth"),
  startMonth: document.querySelector("#startMonth"),
  endMonth: document.querySelector("#endMonth"),
  leadStartWeek: document.querySelector("#leadStartWeek"),
  leadEndWeek: document.querySelector("#leadEndWeek"),
  leadSearchInput: document.querySelector("#leadSearchInput"),
  periodSubYear: document.querySelector("#periodSubYear"),
  periodSubQuarter: document.querySelector("#periodSubQuarter"),
  periodSubMonth: document.querySelector("#periodSubMonth"),
  periodSubRange: document.querySelector("#periodSubRange"),
  groupSelect: document.querySelector("#groupSelect"),
  saleSelect: document.querySelector("#saleSelect"),
  searchInput: document.querySelector("#searchInput"),
  showUnmapped: document.querySelector("#showUnmapped"),
  tabDescription: document.querySelector("#tabDescription"),
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
  revenueSummary: document.querySelector("#revenueSummary"),
  revenueSubtabButtons: Array.from(document.querySelectorAll("[data-revenue-subtab]")),
  revenueSubtabPanels: Array.from(document.querySelectorAll("[data-revenue-panel]")),
  revenuePerformanceSummary: document.querySelector("#revenuePerformanceSummary"),
  revenuePerformanceChart: document.querySelector("#revenuePerformanceChart"),
  revenuePerformanceBySale: document.querySelector("#revenuePerformanceBySale"),
  revenuePerformanceDetail: document.querySelector("#revenuePerformanceDetail"),
  revenueMonthlyChart: document.querySelector("#revenueMonthlyChart"),
  revenueBySale: document.querySelector("#revenueBySale"),
  revenueTargetTitle: document.querySelector("#revenueTargetTitle"),
  revenueTargetDescription: document.querySelector("#revenueTargetDescription"),
  revenueTargetPanel: document.querySelector("#revenueTargetPanel"),
  revenueTargetToggle: document.querySelector("#revenueTargetToggle"),
  revenueTargetBody: document.querySelector("#revenueTargetBody"),
  revenueTargetSummary: document.querySelector("#revenueTargetSummary"),
  revenueTargetTable: document.querySelector("#revenueTargetTable"),
  exportRevenueTargetCsv: document.querySelector("#exportRevenueTargetCsv"),
  revenueDistributionTitle: document.querySelector("#revenueDistributionTitle"),
  revenueDistributionDescription: document.querySelector("#revenueDistributionDescription"),
  revenueDistributionPanel: document.querySelector("#revenueDistributionPanel"),
  revenueDistributionToggle: document.querySelector("#revenueDistributionToggle"),
  revenueDistributionBody: document.querySelector("#revenueDistributionBody"),
  revenueDistributionSummary: document.querySelector("#revenueDistributionSummary"),
  revenueDistributionTable: document.querySelector("#revenueDistributionTable"),
  exportRevenueDistributionCsv: document.querySelector("#exportRevenueDistributionCsv"),
  invoiceAnalysisTitle: document.querySelector("#invoiceAnalysisTitle"),
  invoiceAnalysisDescription: document.querySelector("#invoiceAnalysisDescription"),
  invoiceDateModeButtons: Array.from(document.querySelectorAll("[data-invoice-date-mode]")),
  invoiceAnalysisSummary: document.querySelector("#invoiceAnalysisSummary"),
  invoiceAnalysisTable: document.querySelector("#invoiceAnalysisTable"),
  exportInvoiceAnalysisCsv: document.querySelector("#exportInvoiceAnalysisCsv"),
  exportRevenueDetailCsv: document.querySelector("#exportRevenueDetailCsv"),
  revenueDetailMonthlyTotal: document.querySelector("#revenueDetailMonthlyTotal"),
  revenueDetailPanel: document.querySelector("#revenueDetailPanel"),
  revenueDetailToggle: document.querySelector("#revenueDetailToggle"),
  revenueDetailBody: document.querySelector("#revenueDetailBody"),
  targetContributionPanel: document.querySelector("#targetContributionPanel"),
  targetContributionToggle: document.querySelector("#targetContributionToggle"),
  targetContributionBody: document.querySelector("#targetContributionBody"),
  revenueExpectedFallback: document.querySelector("#revenueExpectedFallback"),
  revenueDetailTable: document.querySelector("#revenueDetailTable"),
  revenueDataChecksPanel: document.querySelector("#revenueDataChecksPanel"),
  revenueDataChecksToggle: document.querySelector("#revenueDataChecksToggle"),
  revenueDataChecksBody: document.querySelector("#revenueDataChecksBody"),
  revenueExceptionTable: document.querySelector("#revenueExceptionTable"),
  revenueAndInputs: Array.from(document.querySelectorAll("[data-revenue-and]")),
  revenueNotInputs: Array.from(document.querySelectorAll("[data-revenue-not]")),
  transactionAndInputs: Array.from(document.querySelectorAll("[data-transaction-and]")),
  transactionNotInputs: Array.from(document.querySelectorAll("[data-transaction-not]")),
  transactionTable: document.querySelector("#transactionTable"),
  transactionTotalAmount: document.querySelector("#transactionTotalAmount"),
  allDealsAndInputs: Array.from(document.querySelectorAll("[data-all-deals-and]")),
  allDealsNotInputs: Array.from(document.querySelectorAll("[data-all-deals-not]")),
  allDealsTotalAmount: document.querySelector("#allDealsTotalAmount"),
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

function parsePositiveInt(value) {
  const number = Math.round(parseMoneyValue(value));
  return Number.isFinite(number) && number > 0 ? number : 0;
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

function currentYearKey() {
  return localIsoDate().slice(0, 4);
}

function currentMonthKey() {
  return localIsoDate().slice(0, 7);
}

function currentYearStartMonthKey() {
  return `${currentYearKey()}-01`;
}

function currentQuarterKey() {
  const month = Number(currentMonthKey().slice(5, 7));
  const quarter = Math.max(1, Math.ceil(month / 3));
  return `${currentYearKey()}-Q${quarter}`;
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

function signedMoney(value) {
  const rounded = Math.round(value || 0);
  if (rounded > 0) return `+${money(rounded)}`;
  if (rounded < 0) return `-${money(Math.abs(rounded))}`;
  return money(0);
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
  if (!year || !month) return "";
  const date = new Date(Date.UTC(year, month, 0));
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function addMonthsToDate(dateText, offset) {
  const parts = parseDateValue(dateText);
  if (!parts) return "";
  const target = new Date(Date.UTC(parts.year, parts.month - 1 + offset, 1));
  const year = target.getUTCFullYear();
  const month = target.getUTCMonth() + 1;
  const lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();
  const day = Math.min(parts.day, lastDay);
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function displayDate(dateText) {
  const parts = parseDateValue(dateText);
  if (!parts) return cleanText(dateText) || "-";
  return parts.date;
}

function displayDateDmy(dateText) {
  const parts = parseDateValue(dateText);
  if (!parts) return cleanText(dateText) || "-";
  return `${String(parts.day).padStart(2, "0")}/${String(parts.month).padStart(2, "0")}/${parts.year}`;
}

function displayDateRangeDmy(value) {
  const text = cleanText(value);
  if (!text || text === "-") return text || "-";
  return text.replace(/\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/g, (dateText) => displayDateDmy(dateText));
}

function periodMonths() {
  const months = dashboardData.months || [];
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
  if (state.periodType === "month") {
    return months.filter((month) => month === state.periodMonth);
  }
  return months.filter((month) => month >= state.startMonth && month <= state.endMonth);
}

function isIsoWeekKey(value) {
  return /^\d{4}-W\d{2}$/.test(String(value || ""));
}

function weekLabel(weekKey) {
  const range = isoWeekRange(weekKey);
  return `${weekKey} (${range.start.date})`;
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

function previousIsoWeekKey(value = localIsoDate(), weeksBack = 1) {
  const parts = parseDateValue(value);
  if (!parts) return "";
  return isoWeekKeyFromDate(makeDateParts(dateToIso(new Date(parts.ts - weeksBack * 7 * 86400000))));
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
  const key = dealPeriodKey(deal, deal.trackingMonth);
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
    `${deal.id} ${deal.saleName} ${deal.responsible} ${deal.group} ${deal.company} ${deal.dealName} ${deal.pipeline} ${deal.stage} ${deal.rawStage || ""} ${deal.dealType || ""} ${deal.product || ""} ${deal.category || ""} ${deal.billingType || ""} ${deal.contractStartDate || ""} ${deal.contractEndDate || ""} ${deal.servicePeriod || ""} ${performanceBucket(deal)} ${deal.expectedDate || ""} ${deal.stageChangeDate || ""} ${deal.createdDate || ""}`,
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
    globalCategory = true,
    statuses = null,
  } = options;
  return allDealDetails()
    .filter((deal) => itemMatchesSaleScope(deal, scope))
    .filter((deal) => !globalCategory || categoryMatches(deal.category))
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
  initTheme();
  renderPeriodOptions();
  renderGroupOptions();
  renderSaleOptions();
  renderFilterVisibility();
  setupSelectControls();
  syncSelectControls();

  els.viewMode.addEventListener("change", () => {
    state.viewMode = els.viewMode.value;
    renderFilterVisibility();
    renderSaleOptions();
    syncSelectControls();
    render();
  });
  els.periodTypeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.periodType = button.dataset.periodType;
      renderFilterVisibility();
      render();
    });
  });
  els.periodYear.addEventListener("change", () => {
    state.periodYear = els.periodYear.value;
    syncSelectControls();
    render();
  });
  els.periodQuarter.addEventListener("change", () => {
    state.periodQuarter = els.periodQuarter.value;
    syncSelectControls();
    render();
  });
  els.periodMonth.addEventListener("change", () => {
    state.periodMonth = els.periodMonth.value;
    syncSelectControls();
    render();
  });
  els.startMonth.addEventListener("change", () => {
    state.startMonth = els.startMonth.value;
    if (state.startMonth > state.endMonth) {
      state.endMonth = state.startMonth;
      els.endMonth.value = state.endMonth;
    }
    syncSelectControls();
    render();
  });
  els.endMonth.addEventListener("change", () => {
    state.endMonth = els.endMonth.value;
    if (state.endMonth < state.startMonth) {
      state.startMonth = state.endMonth;
      els.startMonth.value = state.startMonth;
    }
    syncSelectControls();
    render();
  });
  els.leadStartWeek.addEventListener("change", () => {
    state.leadStartWeek = els.leadStartWeek.value;
    if (state.leadStartWeek > state.leadEndWeek) {
      state.leadEndWeek = state.leadStartWeek;
      els.leadEndWeek.value = state.leadEndWeek;
    }
    syncSelectControls();
    renderFilteredView();
  });
  els.leadEndWeek.addEventListener("change", () => {
    state.leadEndWeek = els.leadEndWeek.value;
    if (state.leadEndWeek < state.leadStartWeek) {
      state.leadStartWeek = state.leadEndWeek;
      els.leadStartWeek.value = state.leadStartWeek;
    }
    syncSelectControls();
    renderFilteredView();
  });
  els.leadSearchInput.addEventListener("input", () => {
    state.leadSearch = els.leadSearchInput.value;
    if (state.tab === "leads") renderFilteredView();
  });
  els.groupSelect.addEventListener("change", () => {
    state.group = els.groupSelect.value;
    renderSaleOptions();
    syncSelectControls();
    render();
  });
  els.saleSelect.addEventListener("change", () => {
    state.sale = els.saleSelect.value;
    syncSelectControls();
    render();
  });
  els.searchInput.addEventListener("input", () => {
    state.search = els.searchInput.value;
    renderFilteredView();
  });
  els.showUnmapped.addEventListener("change", () => {
    state.showUnmapped = els.showUnmapped.checked;
    renderGroupOptions();
    renderSaleOptions();
    syncSelectControls();
    render();
  });
  document.querySelectorAll(".segment button[data-category]").forEach((button) => {
    button.addEventListener("click", () => {
      if (state.tab === "leads") {
        state.leadCategory = button.dataset.category;
        renderCategoryButtons();
        renderFilteredView();
        return;
      }
      state.category = button.dataset.category;
      renderCategoryButtons();
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
      renderFilteredView();
    });
  });
  els.revenueSubtabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.revenueSubtab = button.dataset.revenueSubtab || "performance";
      renderRevenueSubtabs();
    });
  });
  els.invoiceDateModeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.invoiceDateMode = button.dataset.invoiceDateMode || "posting";
      renderFilteredView();
    });
  });
  els.salesStatusFilters.querySelectorAll("input[data-sales-status]").forEach((input) => {
    input.addEventListener("change", () => {
      state.statusFilters[input.dataset.salesStatus] = input.checked;
      renderFilteredView();
    });
  });
  [...els.transactionAndInputs, ...els.transactionNotInputs].forEach((input) => {
    input.addEventListener("input", () => {
      if (state.tab === "sales") renderFilteredView();
    });
  });
  [...els.allDealsAndInputs, ...els.allDealsNotInputs].forEach((input) => {
    input.addEventListener("input", () => {
      if (state.tab === "deals") renderFilteredView();
    });
  });
  [...els.revenueAndInputs, ...els.revenueNotInputs].forEach((input) => {
    input.addEventListener("input", () => {
      if (state.tab === "revenue") renderFilteredView();
    });
  });
  els.revenueExpectedFallback?.addEventListener("change", () => {
    state.revenueExpectedFallback = els.revenueExpectedFallback.checked;
    if (state.tab === "revenue") renderFilteredView();
  });
  els.revenueTargetToggle?.addEventListener("click", () => {
    state.revenueTargetVisible = !state.revenueTargetVisible;
    renderRevenueTargetVisibility();
  });
  els.revenueDistributionToggle?.addEventListener("click", () => {
    state.revenueDistributionVisible = !state.revenueDistributionVisible;
    renderRevenueDistributionVisibility();
  });
  els.revenueDetailToggle?.addEventListener("click", () => {
    state.revenueDetailVisible = !state.revenueDetailVisible;
    renderRevenueDetailVisibility();
  });
  els.targetContributionToggle?.addEventListener("click", () => {
    state.targetContributionVisible = !state.targetContributionVisible;
    renderTargetContributionVisibility();
  });
  els.revenueDataChecksToggle?.addEventListener("click", () => {
    state.revenueDataChecksVisible = !state.revenueDataChecksVisible;
    renderRevenueDataChecksVisibility();
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
      renderFilteredView();
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
      renderFilteredView();
      return;
    }
    const dealRow = event.target.closest("[data-deal-key]");
    if (dealRow) openDealModal(dealRow.dataset.dealKey);
  });
  [els.revenuePerformanceDetail, els.revenueDetailTable, els.revenueExceptionTable].forEach((table) => {
    table?.addEventListener("click", (event) => {
      const dealRow = event.target.closest("[data-deal-key]");
      if (dealRow) openDealModal(dealRow.dataset.dealKey);
    });
  });
  els.revenueTargetTable?.addEventListener("click", (event) => {
    const toggle = event.target.closest("[data-revenue-target-toggle]");
    if (!toggle) return;
    const key = toggle.dataset.revenueTargetToggle;
    state.revenueTargetExpanded[key] = !state.revenueTargetExpanded[key];
    renderFilteredView();
  });
  els.revenueDistributionTable?.addEventListener("click", (event) => {
    const toggle = event.target.closest("[data-revenue-distribution-toggle]");
    if (!toggle) return;
    const key = toggle.dataset.revenueDistributionToggle;
    state.revenueDistributionExpanded[key] = !state.revenueDistributionExpanded[key];
    renderFilteredView();
  });
  els.invoiceAnalysisTable?.addEventListener("click", (event) => {
    const toggle = event.target.closest("[data-invoice-analysis-toggle]");
    if (!toggle) return;
    const key = toggle.dataset.invoiceAnalysisToggle;
    state.invoiceAnalysisExpanded[key] = !state.invoiceAnalysisExpanded[key];
    renderFilteredView();
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
  document.addEventListener("click", (event) => {
    const filterButton = event.target.closest("[data-column-filter-table]");
    if (filterButton) {
      event.preventDefault();
      event.stopPropagation();
      openColumnFilter(filterButton.dataset.columnFilterTable, filterButton.dataset.columnFilterKey, filterButton);
      return;
    }
    if (!event.target.closest("#columnFilterPopover")) closeColumnFilter();
  });
  els.savePipelineMatching.addEventListener("click", savePipelineMatching);
  els.clearPipelineMatching.addEventListener("click", clearPipelineMatching);
  document.addEventListener("click", (event) => {
    const settingsButton = event.target.closest("#openSettings");
    if (settingsButton) openSettings();
    const fieldInfoButton = event.target.closest("#openFieldInfo");
    if (fieldInfoButton) openFieldInfo();
  });
  els.closeSettings.addEventListener("click", closeSettings);
  els.closeFieldInfo.addEventListener("click", closeFieldInfo);
  els.closeDealModal.addEventListener("click", closeDealModal);
  els.settingsModal.addEventListener("click", (event) => {
    if (event.target === els.settingsModal) closeSettings();
  });
  els.fieldInfoModal.addEventListener("click", (event) => {
    if (event.target === els.fieldInfoModal) closeFieldInfo();
  });
  els.dealModal.addEventListener("click", (event) => {
    if (event.target === els.dealModal) closeDealModal();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeWeekSelects();
    if (event.key !== "Escape") return;
    if (!els.dealModal.hidden) closeDealModal();
    if (!els.fieldInfoModal.hidden) closeFieldInfo();
    if (!els.settingsModal.hidden) closeSettings();
  });
  document.addEventListener("click", (event) => {
    if (!event.target.closest(".week-select-control")) closeWeekSelects();
  });
  [els.targetCsvInput, els.dealCsvInput, els.mappingCsvInput, els.invoiceCsvInput].forEach((input) => {
    input.addEventListener("change", () => {
      const targetFile = els.targetCsvInput.files?.[0]?.name;
      const dealFile = els.dealCsvInput.files?.[0]?.name;
      const mappingFile = els.mappingCsvInput.files?.[0]?.name;
      const invoiceFile = els.invoiceCsvInput.files?.[0]?.name;
      if (input === els.targetCsvInput && targetFile) setAutoLoadedFileName("target", "");
      if (input === els.mappingCsvInput && mappingFile) setAutoLoadedFileName("mapping", "");
      const selected = [
        targetFile && `Target: ${targetFile}`,
        dealFile && `Deal: ${dealFile}`,
        mappingFile && `Mapping: ${mappingFile}`,
        invoiceFile && `Invoice: ${invoiceFile}`,
      ].filter(Boolean);
      setUploadStatus(selected.length ? selected.join(" | ") : "ยังไม่ได้เลือกไฟล์ใหม่");
    });
  });
  els.applyCsvFiles.addEventListener("click", handleCsvRefresh);
  els.exportRevenueTargetCsv?.addEventListener("click", exportRevenueTargetCsv);
  els.exportRevenueDistributionCsv?.addEventListener("click", exportRevenueDistributionCsv);
  els.exportInvoiceAnalysisCsv?.addEventListener("click", exportInvoiceAnalysisCsv);
  els.exportRevenueDetailCsv?.addEventListener("click", exportRevenueDetailCsv);
  els.clearCsvFiles.addEventListener("click", () => {
    els.targetCsvInput.value = "";
    els.dealCsvInput.value = "";
    els.mappingCsvInput.value = "";
    els.invoiceCsvInput.value = "";
    setAutoLoadedFileName("target", "");
    setAutoLoadedFileName("mapping", "");
    window.localStorage.removeItem(STORED_DATA_KEY);
    dashboardData = JSON.parse(JSON.stringify(initialDashboardData));
    state.group = "all";
    state.sale = "all";
    resetColumnFilters();
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
  const currentYear = currentYearKey();
  const currentQuarter = currentQuarterKey();
  const currentMonth = currentMonthKey();
  const currentYearStart = currentYearStartMonthKey();
  const allMonths = Array.from(new Set(dashboardData.months || [])).sort();
  const yearOptions = Array.from(
    new Set([
      ...allMonths.map((month) => month.slice(0, 4)).filter(Boolean),
      "2024",
      "2025",
      "2026",
      "2027",
      "2028",
      currentYear,
    ]),
  ).sort();
  els.periodYear.innerHTML = yearOptions.map((year) => `<option value="${year}">${year}</option>`).join("");
  if (!yearOptions.includes(state.periodYear)) state.periodYear = yearOptions.includes(currentYear) ? currentYear : yearOptions[0] || currentYear;
  els.periodYear.value = state.periodYear;

  const quarterOptions = [];
  yearOptions.forEach((year) => {
    for (let quarter = 1; quarter <= 4; quarter += 1) {
      quarterOptions.push({ value: `${year}-Q${quarter}`, label: `${ordinalQuarter(quarter)} ${year}` });
    }
  });
  els.periodQuarter.innerHTML = quarterOptions.map((quarter) => `<option value="${quarter.value}">${quarter.label}</option>`).join("");
  if (!quarterOptions.some((quarter) => quarter.value === state.periodQuarter)) {
    state.periodQuarter = quarterOptions.some((quarter) => quarter.value === currentQuarter) ? currentQuarter : quarterOptions[0]?.value || currentQuarter;
  }
  els.periodQuarter.value = state.periodQuarter;

  const monthOptions = allMonths.length ? allMonths : [currentMonth];
  const monthOptionHtml = monthOptions.map((month) => `<option value="${month}">${monthLabel(month)}</option>`).join("");
  els.periodMonth.innerHTML = monthOptionHtml;
  if (!monthOptions.includes(state.periodMonth)) state.periodMonth = monthOptions.includes(currentMonth) ? currentMonth : monthOptions[0] || currentMonth;
  els.periodMonth.value = state.periodMonth;

  const leadWeeks = leadWeekOptions();
  const leadOptions = leadWeeks.map((week) => `<option value="${week.value}">${escapeHtml(week.label)}</option>`).join("");
  els.leadStartWeek.innerHTML = leadOptions;
  els.leadEndWeek.innerHTML = leadOptions;
  const previousWeek = previousIsoWeekKey(localIsoDate(), 3);
  const currentWeek = isoWeekKeyFromDate(localIsoDate());
  if (!leadWeeks.some((week) => week.value === state.leadStartWeek)) state.leadStartWeek = leadWeeks.find((week) => week.value === previousWeek)?.value || leadWeeks[0]?.value || "2026-W01";
  if (!leadWeeks.some((week) => week.value === state.leadEndWeek)) state.leadEndWeek = leadWeeks.find((week) => week.value === currentWeek)?.value || leadWeeks[leadWeeks.length - 1]?.value || state.leadStartWeek;
  if (state.leadStartWeek > state.leadEndWeek) state.leadEndWeek = state.leadStartWeek;
  els.leadStartWeek.value = state.leadStartWeek;
  els.leadEndWeek.value = state.leadEndWeek;
  syncSelectControls();

  els.startMonth.innerHTML = monthOptionHtml;
  els.endMonth.innerHTML = monthOptionHtml;
  if (!monthOptions.includes(state.startMonth)) {
    state.startMonth = monthOptions.includes(currentYearStart) ? currentYearStart : monthOptions.find((month) => month.startsWith(`${currentYear}-`)) || monthOptions[0] || currentYearStart;
  }
  if (!monthOptions.includes(state.endMonth)) {
    const currentOrEarlierMonth = monthOptions.filter((month) => month <= currentMonth).pop();
    state.endMonth = monthOptions.includes(currentMonth) ? currentMonth : currentOrEarlierMonth || monthOptions[monthOptions.length - 1] || currentMonth;
  }
  if (state.startMonth > state.endMonth) state.startMonth = state.endMonth;
  els.startMonth.value = state.startMonth;
  els.endMonth.value = state.endMonth;
  syncSelectControls();
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

function setupSelectControls() {
  document.querySelectorAll("select").forEach(setupWeekSelectControl);
}

function setupWeekSelectControl(select) {
  if (!select || select.dataset.weekSelectReady) return;
  select.dataset.weekSelectReady = "true";
  select.classList.add("native-week-select");
  const wrapper = document.createElement("div");
  wrapper.className = "week-select-control";
  const button = document.createElement("button");
  button.type = "button";
  button.className = "week-select-button";
  button.setAttribute("aria-haspopup", "listbox");
  const list = document.createElement("div");
  list.className = "week-select-list";
  list.hidden = true;
  list.setAttribute("role", "listbox");
  select.insertAdjacentElement("afterend", wrapper);
  wrapper.append(button, list);
  button.addEventListener("click", () => toggleWeekSelect(select));
  button.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown" || event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      toggleWeekSelect(select, true);
    }
  });
}

function syncSelectControls() {
  document.querySelectorAll("select").forEach(syncSelectControl);
}

function syncSelectControl(select) {
  if (!select) return;
  const wrapper = select.nextElementSibling?.classList?.contains("week-select-control") ? select.nextElementSibling : null;
  if (!wrapper) return;
  const button = wrapper.querySelector(".week-select-button");
  const list = wrapper.querySelector(".week-select-list");
  const selected = select.selectedOptions[0];
  button.textContent = selected?.textContent || "-";
  list.innerHTML = Array.from(select.options)
    .map(
      (option) => `
        <button type="button" class="${option.value === select.value ? "active" : ""}" data-week-value="${escapeHtml(option.value)}" role="option" aria-selected="${option.value === select.value}" title="${escapeHtml(option.textContent)}">
          ${escapeHtml(option.textContent)}
        </button>
      `,
    )
    .join("");
  list.querySelectorAll("[data-week-value]").forEach((item) => {
    item.addEventListener("click", () => {
      select.value = item.dataset.weekValue;
      select.dispatchEvent(new Event("change", { bubbles: true }));
      closeWeekSelects();
    });
  });
}

function toggleWeekSelect(select, open) {
  const wrapper = select?.nextElementSibling?.classList?.contains("week-select-control") ? select.nextElementSibling : null;
  if (!wrapper) return;
  const list = wrapper.querySelector(".week-select-list");
  const button = wrapper.querySelector(".week-select-button");
  const shouldOpen = open ?? list.hidden;
  closeWeekSelects();
  syncSelectControl(select);
  list.hidden = !shouldOpen;
  button.setAttribute("aria-expanded", String(shouldOpen));
  if (!shouldOpen) return;
  requestAnimationFrame(() => {
    const active = list.querySelector(".active");
    if (active) active.scrollIntoView({ block: "center" });
  });
}

function closeWeekSelects() {
  document.querySelectorAll(".week-select-list").forEach((list) => {
    list.hidden = true;
  });
  document.querySelectorAll(".week-select-button").forEach((button) => {
    button.setAttribute("aria-expanded", "false");
  });
}

function ordinalQuarter(quarter) {
  return ["1st", "2nd", "3rd", "4th"][quarter - 1] || `${quarter}th`;
}

function renderFilterVisibility() {
  els.viewMode.value = state.viewMode;
  els.periodTypeButtons.forEach((button) => {
    const active = button.dataset.periodType === state.periodType;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", String(active));
  });
  els.groupFilterContainer.hidden = false;
  els.salesFilterContainer.hidden = false;
  els.periodSubYear.hidden = state.periodType !== "year";
  els.periodSubQuarter.hidden = state.periodType !== "quarter";
  els.periodSubMonth.hidden = state.periodType !== "month";
  els.periodSubRange.hidden = state.periodType !== "range";
  syncSelectControls();
}

function openSettings() {
  renderPipelineMatchingSettings();
  els.settingsModal.hidden = false;
  els.targetCsvInput.focus();
}

function closeSettings() {
  els.settingsModal.hidden = true;
}

function openFieldInfo() {
  renderFieldInfoSummary();
  els.fieldInfoModal.hidden = false;
  els.closeFieldInfo.focus();
}

function closeFieldInfo() {
  els.fieldInfoModal.hidden = true;
}

function closeDealModal() {
  els.dealModal.hidden = true;
}

function fieldCheckRows() {
  const rows = [];
  if (Array.isArray(dashboardData.dealDetails)) rows.push(...dashboardData.dealDetails);
  if (Array.isArray(dashboardData.openDeals)) rows.push(...dashboardData.openDeals);
  return rows;
}

function invoiceCheckRows() {
  return Array.isArray(dashboardData.invoiceRows) ? dashboardData.invoiceRows : [];
}

function hasFieldValue(key) {
  return fieldCheckRows().some((row) => {
    const value = row?.[key];
    if (typeof value === "number") return Number.isFinite(value) && value !== 0;
    return cleanText(value) !== "";
  });
}

function hasPositiveFieldValue(key) {
  return fieldCheckRows().some((row) => Number(row?.[key] || 0) > 0);
}

function hasAnyDealWithStatus(status) {
  return fieldCheckRows().some((row) => cleanText(row?.status).toLowerCase() === status);
}

function hasTargetData() {
  return (dashboardData.sales || []).some((sale) => {
    if (sale.hasTarget) return true;
    const annual = sale.targetAnnual || {};
    if (Number(annual.total || 0) > 0 || Number(annual.new || 0) > 0 || Number(annual.renew || 0) > 0) return true;
    return Object.values(sale.monthly || {}).some((month) => Number(month?.total || 0) > 0);
  });
}

function hasMappingData(kind) {
  return Number(mappingCounts(activeMappings())[kind] || 0) > 0;
}

function hasAnyPipelineRule() {
  return Number(mappingCounts(activeMappings()).pipelineRules || 0) > 0;
}

function hasInvoiceFieldValue(key) {
  return invoiceCheckRows().some((row) => {
    const value = row?.[key];
    if (typeof value === "number") return Number.isFinite(value) && value !== 0;
    return cleanText(value) !== "";
  });
}

function fieldRequirementGroups() {
  const saleOwner = () => hasFieldValue("saleName") || hasFieldValue("responsible");
  const dealTitle = () => hasFieldValue("company") || hasFieldValue("dealName");
  const amount = () => hasPositiveFieldValue("amount") || hasPositiveFieldValue("forecastAmount");
  const category = () => hasFieldValue("category") || hasFieldValue("dealType") || hasFieldValue("rawDealType");

  return [
    {
      tab: "Overview",
      fields: [
        ["Sale Target", hasTargetData],
        ["Sale / Responsible", saleOwner],
        ["Sale Group", () => hasFieldValue("group")],
        ["Company / Deal", dealTitle],
        ["Income / Amount", amount],
        ["Stage", () => hasFieldValue("stage") || hasFieldValue("rawStage")],
        ["Pipeline", () => hasFieldValue("pipeline")],
        ["Deal Type (New/Renew)", category],
        ["Expected Close Date", () => hasFieldValue("expectedDate")],
      ],
    },
    {
      tab: "Sales Performance",
      fields: [
        ["Sale Target", hasTargetData],
        ["Sale / Responsible", saleOwner],
        ["Sale Group", () => hasFieldValue("group")],
        ["Won / Pre-WON Stage", () => hasAnyDealWithStatus("won")],
        ["Lost / Pre-LOST Stage", () => hasAnyDealWithStatus("lost")],
        ["Deal Type (New/Renew)", category],
        ["Expected Close Date", () => hasFieldValue("expectedDate")],
        ["Income / Amount", amount],
        ["Pipeline Matching", hasAnyPipelineRule],
      ],
    },
    {
      tab: "Revenue Analysis",
      fields: [
        ["Sale / Responsible", saleOwner],
        ["Company / Deal", dealTitle],
        ["Won / Pre-WON Stage", () => hasAnyDealWithStatus("won")],
        ["Billing Type", () => hasFieldValue("billingType")],
        ["Contract Start Date", () => hasFieldValue("contractStartDate")],
        ["Contract End Date", () => hasFieldValue("contractEndDate")],
        ["Contract Period", () => hasPositiveFieldValue("contractPeriodMonths")],
        ["MRR", () => hasPositiveFieldValue("mrr")],
        ["Income / Contract Amount", amount],
        ["Expected Close Date (Fallback)", () => hasFieldValue("expectedDate")],
      ],
    },
    {
      tab: "Invoice Analysis",
      fields: [
        ["Invoice CSV", () => invoiceCheckRows().length > 0],
        ["Posting Date", () => hasInvoiceFieldValue("postingDate")],
        ["Sale Name", () => hasInvoiceFieldValue("saleName")],
        ["Sale Group", () => hasInvoiceFieldValue("group")],
        ["Customer / Project", () => hasInvoiceFieldValue("company") || hasInvoiceFieldValue("dealName")],
        ["New / Renew", () => hasInvoiceFieldValue("category")],
        ["Invoice Amount", () => hasInvoiceFieldValue("amount")],
      ],
    },
    {
      tab: "All Deals",
      fields: [
        ["ID", () => hasFieldValue("id")],
        ["Sale / Responsible", saleOwner],
        ["Sale Group", () => hasFieldValue("group")],
        ["Company / Deal", dealTitle],
        ["Deal Type (New/Renew)", category],
        ["Stage", () => hasFieldValue("stage") || hasFieldValue("rawStage")],
        ["Pipeline", () => hasFieldValue("pipeline")],
        ["Created Date", () => hasFieldValue("createdDate")],
        ["Expected Close Date", () => hasFieldValue("expectedDate")],
        ["Income / Amount", amount],
      ],
    },
    {
      tab: "Pipeline Risk",
      fields: [
        ["Open Stage", () => hasAnyDealWithStatus("open")],
        ["Sale / Responsible", saleOwner],
        ["Sale Group", () => hasFieldValue("group")],
        ["Pipeline", () => hasFieldValue("pipeline")],
        ["Expected Close Date", () => hasFieldValue("expectedDate")],
        ["Stage Change Date", () => hasFieldValue("stageChangeDate")],
        ["Income / Amount", amount],
      ],
    },
    {
      tab: "Data Quality",
      fields: [
        ["Stage Mapping", () => hasMappingData("stages")],
        ["Deal Type Mapping", () => hasMappingData("dealTypes")],
        ["Sales Mapping", () => hasMappingData("sales")],
        ["Pipeline Matching", hasAnyPipelineRule],
        ["Pipeline", () => hasFieldValue("pipeline")],
        ["Stage", () => hasFieldValue("stage") || hasFieldValue("rawStage")],
        ["Responsible", () => hasFieldValue("responsible")],
        ["Deal Type", () => hasFieldValue("dealType") || hasFieldValue("rawDealType")],
        ["Expected Close Date", () => hasFieldValue("expectedDate")],
      ],
    },
    {
      tab: "New Deals",
      fields: [
        ["Created Date", () => hasFieldValue("createdDate")],
        ["Sale / Responsible", saleOwner],
        ["Sale Group", () => hasFieldValue("group")],
        ["Company / Deal", dealTitle],
        ["Deal Type (New/Renew)", category],
        ["Stage", () => hasFieldValue("stage") || hasFieldValue("rawStage")],
        ["Pipeline", () => hasFieldValue("pipeline")],
        ["Income / Amount", amount],
      ],
    },
  ];
}

function renderFieldInfoSummary() {
  els.fieldInfoContent.innerHTML = fieldRequirementGroups()
    .map(
      (group) => `
        <section class="field-info-section">
          <h3>${escapeHtml(group.tab)}</h3>
          <table class="field-info-table">
            <thead>
              <tr>
                <th>Field</th>
                <th class="status-col">Status</th>
              </tr>
            </thead>
            <tbody>
              ${group.fields
                .map(([label, check]) => {
                  const ok = Boolean(check());
                  return `
                    <tr>
                      <td>${escapeHtml(label)}</td>
                      <td class="status-col">
                        <span class="field-check ${ok ? "ok" : "missing"}" aria-label="${ok ? "Checked" : "Cross"}">
                          ${ok ? "✓" : "×"}
                        </span>
                      </td>
                    </tr>
                  `;
                })
                .join("")}
            </tbody>
          </table>
        </section>
      `,
    )
    .join("");
}

function hasMissingFieldRequirements() {
  return fieldRequirementGroups().some((group) => group.fields.some(([, check]) => !check()));
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

async function autoLoadDefaultSetupFiles() {
  const statusParts = [];
  let mappings = activeMappings();
  try {
    const stageResponse = await fetch("./data/Stage Mapping.csv", { cache: "no-store" });
    if (stageResponse.ok) {
      const stageMappings = parseMappingCsv(await stageResponse.text());
      const stageCounts = mappingCounts(stageMappings);
      if (stageCounts.stages || stageCounts.dealTypes || stageCounts.pipelines || stageCounts.pipelineRules || stageCounts.sales) {
        mappings = mergeMappings(mappings, stageMappings);
        statusParts.push("Stage Mapping");
        setAutoLoadedFileName("mapping", "Stage Mapping.csv");
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
  syncSelectControls();
}

function renderSaleOptions() {
  const sales = salesForOptions();
  if (state.sale !== "all" && !sales.some((sale) => sale.key === state.sale)) state.sale = "all";
  els.saleSelect.innerHTML = `<option value="all">ทุกคน</option>${sales
    .map((sale) => `<option value="${escapeHtml(sale.key)}">${escapeHtml(sale.name)}</option>`)
    .join("")}`;
  els.saleSelect.value = state.sale;
  syncSelectControls();
}

function renderMeta() {
  const meta = dashboardData.metadata;
  const generated = new Date(meta.generatedAt).toLocaleString("th-TH", {
    dateStyle: "medium",
    timeStyle: "short",
  });
  const isDark = document.documentElement.dataset.theme === "dark";
  const hasMissingFields = hasMissingFieldRequirements();
  els.metaBox.innerHTML = `
    <button type="button" class="settings-button" id="openSettings" aria-label="Open settings">
      <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M12 15.5A3.5 3.5 0 1 0 12 8a3.5 3.5 0 0 0 0 7.5Z"></path>
        <path d="M19.4 15a1.8 1.8 0 0 0 .36 1.98l.05.05a2.1 2.1 0 1 1-2.97 2.97l-.05-.05a1.8 1.8 0 0 0-1.98-.36 1.8 1.8 0 0 0-1.09 1.65V21.4a2.1 2.1 0 1 1-4.2 0v-.08A1.8 1.8 0 0 0 8.43 19.7a1.8 1.8 0 0 0-1.98.36l-.05.05a2.1 2.1 0 1 1-2.97-2.97l.05-.05a1.8 1.8 0 0 0 .36-1.98 1.8 1.8 0 0 0-1.65-1.09H2.1a2.1 2.1 0 1 1 0-4.2h.08A1.8 1.8 0 0 0 3.8 8.73a1.8 1.8 0 0 0-.36-1.98l-.05-.05A2.1 2.1 0 1 1 6.36 3.73l.05.05a1.8 1.8 0 0 0 1.98.36A1.8 1.8 0 0 0 9.48 2.5V2.1a2.1 2.1 0 1 1 4.2 0v.08a1.8 1.8 0 0 0 1.09 1.65 1.8 1.8 0 0 0 1.98-.36l.05-.05a2.1 2.1 0 1 1 2.97 2.97l-.05.05a1.8 1.8 0 0 0-.36 1.98 1.8 1.8 0 0 0 1.65 1.09h.39a2.1 2.1 0 1 1 0 4.2h-.08A1.8 1.8 0 0 0 19.4 15Z"></path>
      </svg>
    </button>
    <button type="button" class="info-button ${hasMissingFields ? "has-missing-fields" : ""}" id="openFieldInfo" aria-label="${hasMissingFields ? "Open field information, missing fields found" : "Open field information"}">
      <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="9"></circle>
        <path d="M12 11v5"></path>
        <path d="M12 8h.01"></path>
      </svg>
    </button>
    <button type="button" class="theme-toggle-button" id="themeToggle" aria-label="Toggle dark mode" aria-pressed="${isDark}">
      <svg class="theme-sun" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="4"></circle>
        <path d="M12 2v2"></path>
        <path d="M12 20v2"></path>
        <path d="m4.93 4.93 1.41 1.41"></path>
        <path d="m17.66 17.66 1.41 1.41"></path>
        <path d="M2 12h2"></path>
        <path d="M20 12h2"></path>
        <path d="m6.34 17.66-1.41 1.41"></path>
        <path d="m19.07 4.93-1.41 1.41"></path>
      </svg>
      <svg class="theme-moon" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M20.5 14.5A8.5 8.5 0 0 1 9.5 3.5 7 7 0 1 0 20.5 14.5Z"></path>
      </svg>
    </button>
    <strong>As of ${escapeHtml(meta.asOfDate)}</strong><br>
    Deals: ${dashboardData.quality.dealRows.toLocaleString("th-TH")} rows<br>
    Targets: ${dashboardData.quality.targetRows.toLocaleString("th-TH")} sales<br>
    Invoices: ${(dashboardData.invoiceRows || []).length.toLocaleString("th-TH")} rows<br>
    Built: ${escapeHtml(generated)}
  `;
  els.themeToggleButton = document.querySelector("#themeToggle");
}

function setTheme(theme) {
  const safeTheme = theme === "dark" ? "dark" : "light";
  document.documentElement.dataset.theme = safeTheme;
  window.localStorage.setItem(THEME_KEY, safeTheme);
  if (els.themeToggleButton) els.themeToggleButton.setAttribute("aria-pressed", String(safeTheme === "dark"));
}

function initTheme() {
  const storedTheme = window.localStorage.getItem(THEME_KEY);
  setTheme(storedTheme === "dark" ? "dark" : "light");
  els.metaBox.addEventListener("click", (event) => {
    const button = event.target.closest("#themeToggle");
    if (!button) return;
    setTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");
  });
}

async function handleCsvRefresh() {
  const targetFile = els.targetCsvInput.files?.[0];
  const dealFile = els.dealCsvInput.files?.[0];
  const mappingFile = els.mappingCsvInput.files?.[0];
  const invoiceFile = els.invoiceCsvInput.files?.[0];
  if (!targetFile && !dealFile && !mappingFile && !invoiceFile) {
    setUploadStatus("กรุณาเลือก Sale Target.csv, Deal.csv, Data Mapping.csv หรือ Invoice.csv อย่างน้อย 1 ไฟล์", "error");
    return;
  }
  if ([targetFile, dealFile, mappingFile, invoiceFile].some((file) => file && !file.name.toLowerCase().endsWith(".csv"))) {
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

    if (invoiceFile) {
      const invoiceRows = parseCsv(await invoiceFile.text());
      validateInvoiceRows(invoiceRows);
      dashboardData = mergeInvoiceRowsIntoDashboard(dashboardData, buildInvoiceRowsFromRows(invoiceRows, mappings), invoiceFile.name);
      statusParts.push(`Invoice ${invoiceRows.length.toLocaleString("th-TH")} rows`);
    }

    storeDashboardData(dashboardData);
    state.group = "all";
    state.sale = "all";
    state.search = "";
    state.leadSearch = "";
    els.searchInput.value = "";
    els.leadSearchInput.value = "";
    resetColumnFilters();
    renderPeriodOptions();
    renderGroupOptions();
    renderSaleOptions();
    renderFilterVisibility();
    render();
    if (!els.settingsModal.hidden) renderPipelineMatchingSettings();
    if (!els.fieldInfoModal.hidden) renderFieldInfoSummary();
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

function validateInvoiceRows(rows) {
  if (!rows.length) throw new Error("ไฟล์ Invoice ไม่มีข้อมูล");
  const required = ["Document No.", "Posting Date", "Customer code", "Description", "Amount", "Sale Name", "SALES TEAM (Dimension)", "New/Renew", "Type"];
  const missing = required.filter((column) => !(column in rows[0]));
  if (missing.length) {
    throw new Error(`รูปแบบ CSV ไม่ตรงกับไฟล์ Invoice: ขาดคอลัมน์ ${missing.join(", ")}`);
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

function cleanInvoiceSaleName(value) {
  return cleanText(value)
    .replace(/\[[^\]]*\]/g, "")
    .replace(/\([^)]*\)/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function saleRecordForInvoiceName(rawSaleName, mappings = activeMappings()) {
  const saleName = cleanInvoiceSaleName(rawSaleName) || cleanText(rawSaleName) || "Unknown";
  const normalized = normalizeName(saleName);
  const mappedSale = mappings.salesGroupMap.get(normalizeName(rawSaleName)) || mappings.salesGroupMap.get(normalized);
  const targetSales = dashboardData.sales || [];
  const mappedName = mappedSale?.displayName || saleName;
  const mappedNormalized = normalizeName(mappedName);
  const matched = targetSales.find((sale) => {
    const saleNormalized = normalizeName(sale.name);
    return (
      normalized === saleNormalized ||
      mappedNormalized === saleNormalized ||
      normalized.includes(` ${saleNormalized}`) ||
      mappedNormalized.includes(` ${saleNormalized}`) ||
      saleNormalized.includes(normalized)
    );
  });
  if (matched) return matched;
  return {
    key: makeKey(mappedName),
    name: mappedName,
    group: mappedSale?.group || "",
  };
}

function buildInvoiceRowsFromRows(rows, mappings = activeMappings()) {
  return rows
    .map((row, index) => {
      const postingDate = parseDateValue(row["Posting Date"]);
      const serviceStart = parseInvoiceServiceStart(row.Period);
      const sale = saleRecordForInvoiceName(row["Sale Name"], mappings);
      const amount = parseMoneyValue(row.Amount);
      const category = normalizeSaleType(row["New/Renew"]) || dealCategory(row["New/Renew"], row);
      const group = sale.group || cleanText(row["SALES TEAM (Dimension)"]) || "Unmapped";
      const customer = cleanText(row["Customer code"]) || cleanText(row["Sell-to Customer No."]) || "-";
      const description = cleanText(row.Description) || cleanText(row.No) || "-";
      return {
        id: cleanText(row["Document No."]) || `invoice-${index + 1}`,
        documentNo: cleanText(row["Document No."]),
        postingDate: postingDate ? displayDateDmy(postingDate.date) : cleanText(row["Posting Date"]),
        postingMonthKey: postingDate?.monthKey || "",
        serviceDate: serviceStart ? displayDateDmy(serviceStart.date) : "",
        serviceMonthKey: serviceStart?.monthKey || "",
        monthKey: postingDate?.monthKey || "",
        saleKey: sale.key,
        saleName: sale.name,
        responsible: cleanText(row["Sale Name"]),
        group,
        category,
        company: customer,
        dealName: description,
        contractRange: displayDateRangeDmy(row.Period) || (postingDate ? displayDateDmy(postingDate.date) : "-"),
        billingType: cleanText(row.Type) || "-",
        product: cleanText(row.Product),
        quantity: parseMoneyValue(row.Quantity),
        amount: roundMoney(amount),
        amountIncludingVat: roundMoney(parseMoneyValue(row["Amount Including VAT"])),
        invoiceSearch: normalizeSearch(
          `${row["Document No."] || ""} ${row["Posting Date"] || ""} ${customer} ${description} ${row.Period || ""} ${row["Sale Name"] || ""} ${group} ${row.Product || ""} ${row["New/Renew"] || ""} ${row.Type || ""} ${money(amount)}`,
        ),
      };
    })
    .filter((row) => row.monthKey && row.amount);
}

function parseInvoiceServiceStart(period) {
  const text = cleanText(period);
  if (!text) return null;
  const match = text.match(/(\d{1,4}[-/]\d{1,2}[-/]\d{1,4})/);
  return match ? parseDateValue(match[1]) : null;
}

function mergeInvoiceRowsIntoDashboard(data, invoiceRows, fileName) {
  return {
    ...data,
    metadata: {
      ...data.metadata,
      generatedAt: new Date().toISOString(),
      sourceFiles: {
        ...data.metadata.sourceFiles,
        invoice: fileName,
      },
    },
    invoiceRows,
  };
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
    invoiceRows: (data.invoiceRows || []).map(enrich),
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
    const contractStart = parseDateValue(row["Contract Start Date"]);
    const contractEnd = parseDateValue(row["Contract End Date"]);
    const startDate = parseDateValue(row["Start date"]);
    const closedDate = parseDateValue(row.Closed);

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
      billingType: cleanText(row["Billing Type"]),
      servicePeriod: cleanText(row["Service Period"]),
      startDate: startDate?.date || cleanText(row["Start date"]),
      closedDate: closedDate?.date || cleanText(row.Closed),
      contractStartDate: contractStart?.date || cleanText(row["Contract Start Date"]),
      contractEndDate: contractEnd?.date || cleanText(row["Contract End Date"]),
      contractPeriodMonths: parsePositiveInt(row["Contract Period (จำนวนเดือน)"]),
      mrr: roundMoney(parseMoneyValue(row.MRR)),
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
        <article class="kpi-summary-card ${row.tone} ${row.categoryKey}">
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
    categoryKey: category,
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
  if (state.periodType === "month") return monthLabel(state.periodMonth);
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
  const maxValue = Math.max(1, ...rows.flatMap((row) => [row.target, row.actual]));
  els.monthlyChart.innerHTML = rows
    .map((row) => {
      return `
        <div class="month-row">
          <div class="month-name">${monthLabel(row.month)}</div>
          <div class="sales-month-bar-space">
            <div class="sales-month-target-base" style="width:${row.target ? Math.min(100, Math.max(2, (row.target / maxValue) * 100)) : 0}%" title="Target ${money(row.target)}"></div>
            <div class="sales-month-actual-bar" style="width:${row.actual ? Math.min(100, Math.max(2, (row.actual / maxValue) * 100)) : 0}%" title="Actual ${money(row.actual)}"></div>
            <div class="sales-month-target-outline" style="width:${row.target ? Math.min(100, Math.max(2, (row.target / maxValue) * 100)) : 0}%"></div>
          </div>
          <div class="bar-label">
            Actual ${compactMoney(row.actual)} / Target ${compactMoney(row.target)}<br>
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
        <div class="mix-card ${category}">
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

function salesAmountVariant(category = state.category) {
  if (category === "renew") return "renew";
  if (category === "new") return "new";
  return "total";
}

function salesAmountStatusColor(key, variant = "total") {
  const colorSets = {
    renew: {
      won: "var(--renew)",
      commit: "#14b8a6",
      upside: "#99f6e4",
      open: "rgba(22, 101, 52, 0.28)",
      lost: "var(--lost)",
    },
    new: {
      won: "var(--new-sale)",
      commit: "#7e22ce",
      upside: "#c084fc",
      open: "rgba(107, 33, 168, 0.28)",
      lost: "var(--lost)",
    },
    total: {
      won: "var(--total)",
      commit: "#1d4ed8",
      upside: "#60a5fa",
      open: "rgba(30, 58, 138, 0.24)",
      lost: "var(--lost)",
    },
  };
  return colorSets[variant]?.[key] || colorSets.total[key] || "var(--line)";
}

function legendStatusItem(key, variant) {
  return `<span class="legend-${key}" style="--legend-color:${salesAmountStatusColor(key, variant)}">${escapeHtml(statusLabel(key))}</span>`;
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
    <div class="chart-legend-inline ${variant}">
      <span class="legend-target">Target</span>
      ${visibleKeys.map((key) => legendStatusItem(key, variant)).join("")}
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
  const variant = salesAmountVariant();
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
          return `<rect class="cum-${key}" x="${xAt(index) - barWidth / 2}" y="${y}" width="${barWidth}" height="${Math.max(0, heightValue)}" rx="2" style="fill:${salesAmountStatusColor(key, variant)}"></rect>`;
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
    <div class="chart-legend-inline ${variant}">
      <span class="legend-target-line">Target สะสม</span>
      ${visibleKeys.map((key) => legendStatusItem(key, variant)).join("")}
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
  const baseRows = filteredPerformanceDeals(scope, "all")
    .filter((deal) => categoryMatches(deal.category))
    .filter((deal) => state.statusFilters[performanceBucket(deal)])
    .filter((deal) => dealMatchesAndNotTerms(deal, andTerms, notTerms));
  tableFilterSources.transaction = baseRows;
  const rows = applyColumnFilters("transaction", baseRows);
  sortTransactionRows(rows);
  els.transactionTotalAmount.textContent = money(sum(rows, (row) => row.amount));
  els.transactionTable.innerHTML = transactionTable(rows.slice(0, 500));
}

function dealMatchesAndNotTerms(deal, andTerms, notTerms) {
  const text = transactionSearchText(deal);
  const tokens = transactionSearchTokens(deal);
  if (andTerms.some((term) => !transactionTermMatches(term, text, tokens))) return false;
  if (notTerms.some((term) => transactionTermMatches(term, text, tokens))) return false;
  return true;
}

function applyColumnFilters(tableKey, rows) {
  const filters = state.columnFilters[tableKey] || {};
  const activeEntries = Object.entries(filters).filter(([, values]) => Array.isArray(values));
  if (!activeEntries.length) return rows;
  return rows.filter((row) =>
    activeEntries.every(([key, values]) => values.includes(columnFilterValue(row, key))),
  );
}

function columnFilterValue(deal, key) {
  if (key === "company" || key === "companyDeal") return cleanText(`${deal.company || "-"} / ${deal.dealName || deal.id || "-"}`);
  if (key === "status") return statusLabel(deal.status || performanceBucket(deal));
  if (key === "risk") return riskFilterLabel(deal);
  if (key === "countingStatus") return deal.countingLabel || countingStatusLabel(deal.countingStatus);
  if (key === "amount") return money(deal.amount || 0);
  if (key === "monthlyRevenue") return money(deal.monthlyRevenue || 0);
  if (key === "contractAmount") return money(deal.contractAmount || deal.amount || 0);
  if (key === "forecastAmount") return money(deal.forecastAmount || 0);
  if (key === "expectedDate") return displayDate(deal.expectedDate);
  if (key === "contractRange") return cleanText(deal.contractRange || "-");
  if (key === "flags") return cleanText(deal.flagsLabel || "-");
  if (key === "stageChangeDate") return displayDate(deal.stageChangeDate);
  if (key === "createdDate") return displayDate(deal.createdDate);
  if (key === "saleName") return cleanText(deal.saleName || deal.responsible || "-");
  return cleanText(deal[key] ?? "-") || "-";
}

function countingStatusLabel(status) {
  if (status === "excluded") return "Excluded";
  if (status === "unmapped") return "Unmapped";
  return "Included";
}

function riskFilterLabel(deal) {
  if (deal.status !== "open") return statusLabel(deal.status);
  const labels = [];
  if (deal.risk?.overdue) labels.push("Overdue");
  if (deal.risk?.noExpected) labels.push("No Expected Date");
  if (deal.risk?.stale30) labels.push("Stale > 30 days");
  if (deal.risk?.notContacted) labels.push("Not Contacted");
  return labels.length ? labels.join(", ") : "No Risk";
}

function columnFilterButton(tableKey, key) {
  const active = Array.isArray(state.columnFilters[tableKey]?.[key]);
  return `<button type="button" class="column-filter-button${active ? " active" : ""}" data-column-filter-table="${escapeHtml(tableKey)}" data-column-filter-key="${escapeHtml(key)}" title="Filter column">▾</button>`;
}

function columnFilterOptions(tableKey, key) {
  const counts = new Map();
  (tableFilterSources[tableKey] || []).forEach((row) => {
    const value = columnFilterValue(row, key);
    counts.set(value, (counts.get(value) || 0) + 1);
  });
  return Array.from(counts.entries())
    .map(([value, count]) => ({ value, label: value || "-", count }))
    .sort((a, b) => a.label.localeCompare(b.label, "th", { numeric: true }));
}

function openColumnFilter(tableKey, key, anchor) {
  const options = columnFilterOptions(tableKey, key);
  const popover = ensureColumnFilterPopover();
  const saved = state.columnFilters[tableKey]?.[key];
  const selected = new Set(Array.isArray(saved) ? saved : options.map((option) => option.value));
  popover.innerHTML = `
    <div class="column-filter-title">เลือกข้อมูลที่ต้องการแสดง</div>
    <div class="column-filter-actions">
      <button type="button" data-column-filter-all>All</button>
      <button type="button" data-column-filter-none>None</button>
      <button type="button" data-column-filter-clear>Clear Filter</button>
    </div>
    <div class="column-filter-options">
      ${
        options.length
          ? options
              .map(
                (option) => `
                  <label class="column-filter-option">
                    <input type="checkbox" value="${escapeHtml(option.value)}" ${selected.has(option.value) ? "checked" : ""} />
                    <span>${escapeHtml(option.label)}</span>
                    <small>${option.count.toLocaleString("th-TH")}</small>
                  </label>
                `,
              )
              .join("")
          : `<div class="column-filter-empty">ไม่มีข้อมูลให้เลือก</div>`
      }
    </div>
    <div class="column-filter-footer">
      <button type="button" class="secondary-button" data-column-filter-close>Close</button>
      <button type="button" class="primary-button highlight-button" data-column-filter-apply>Apply</button>
    </div>
  `;
  popover.hidden = false;
  const rect = anchor.getBoundingClientRect();
  const width = 300;
  popover.style.left = `${Math.min(window.innerWidth - width - 12, Math.max(12, rect.left))}px`;
  popover.style.top = `${Math.min(window.innerHeight - 80, rect.bottom + 6)}px`;
  popover.dataset.tableKey = tableKey;
  popover.dataset.columnKey = key;

  popover.querySelector("[data-column-filter-all]")?.addEventListener("click", () => {
    popover.querySelectorAll(".column-filter-option input").forEach((input) => {
      input.checked = true;
    });
  });
  popover.querySelector("[data-column-filter-none]")?.addEventListener("click", () => {
    popover.querySelectorAll(".column-filter-option input").forEach((input) => {
      input.checked = false;
    });
  });
  popover.querySelector("[data-column-filter-clear]")?.addEventListener("click", () => {
    delete state.columnFilters[tableKey][key];
    closeColumnFilter();
    renderFilteredView();
  });
  popover.querySelector("[data-column-filter-close]")?.addEventListener("click", closeColumnFilter);
  popover.querySelector("[data-column-filter-apply]")?.addEventListener("click", () => {
    const checked = Array.from(popover.querySelectorAll(".column-filter-option input:checked")).map((input) => input.value);
    if (checked.length === options.length) {
      delete state.columnFilters[tableKey][key];
    } else {
      state.columnFilters[tableKey][key] = checked;
    }
    closeColumnFilter();
    renderFilteredView();
  });
}

function ensureColumnFilterPopover() {
  let popover = document.querySelector("#columnFilterPopover");
  if (!popover) {
    popover = document.createElement("div");
    popover.id = "columnFilterPopover";
    popover.className = "column-filter-popover";
    popover.hidden = true;
    document.body.appendChild(popover);
  }
  return popover;
}

function closeColumnFilter() {
  const popover = document.querySelector("#columnFilterPopover");
  if (popover) popover.hidden = true;
}

function resetColumnFilters() {
  state.columnFilters = { transaction: {}, allDeals: {}, revenue: {}, leads: {} };
  closeColumnFilter();
}

function transactionSearchText(deal) {
  const expectedCloseDate = displayDate(deal.expectedDate);
  const counting = countingSearchValues(deal).join(" ");
  const risk = riskFilterLabel(deal);
  return normalizeSearch(
    `${deal.id} ${deal.saleName} ${deal.responsible} ${deal.group} ${deal.company} ${deal.dealName} ${deal.pipeline} ${deal.stage} ${deal.rawStage || ""} ${deal.dealType || ""} ${deal.product} ${deal.category} ${deal.billingType || ""} ${deal.contractStartDate || ""} ${deal.contractEndDate || ""} ${deal.servicePeriod || ""} ${performanceBucket(deal)} ${deal.status || ""} ${counting} ${risk} ${deal.expectedDate || ""} ${expectedCloseDate} ${deal.stageChangeDate || ""}`,
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
      deal.rawStage,
      deal.dealType,
      deal.product,
      deal.category,
      deal.status,
      ...countingSearchValues(deal),
      riskFilterLabel(deal),
      deal.expectedDate,
      expectedCloseDate,
      deal.stageChangeDate,
      performanceBucket(deal),
    ]
      .map(normalizeSearch)
      .flatMap((value) => value.split(/[^\p{L}\p{N}]+/u))
      .filter(Boolean),
  );
}

function countingSearchValues(deal) {
  const status = deal.countingStatus || (deal.countingIncluded === false ? "excluded" : "included");
  const label = deal.countingLabel || countingStatusLabel(status);
  const values = [status, label, countingStatusLabel(status)];
  if (status === "excluded") values.push("not included", "not counted", "ไม่นับยอด");
  if (status === "included") values.push("counted", "นับยอด");
  if (status === "unmapped") values.push("no mapping", "unmapped rule");
  return values.filter(Boolean);
}

function transactionTermMatches(term, text, tokens) {
  if (term.startsWith("not ")) return transactionTermMatches(term.slice(4).trim(), text, tokens);
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
        ${
          rows.length
            ? rows
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
                .join("")
            : `<tr><td colspan="7"><div class="empty">ไม่มี Transaction ที่ตรงกับเงื่อนไข</div></td></tr>`
        }
      </tbody>
    </table>
  `;
}

function transactionHeader(key, label, numeric = false, extraClass = "") {
  const active = state.transactionSort.key === key;
  const icon = active ? (state.transactionSort.dir === "asc" ? "↑" : "↓") : "↕";
  const className = [numeric ? "num" : "", extraClass].filter(Boolean).join(" ");
  return `<th class="${className}"><div class="table-header-tools"><button type="button" class="sort-button" data-transaction-sort="${key}">${escapeHtml(label)} ${icon}</button>${columnFilterButton("transaction", key)}</div></th>`;
}

function renderDealDetailTable(scope) {
  const andTerms = els.allDealsAndInputs.map((input) => normalizeSearch(input.value)).filter(Boolean);
  const notTerms = els.allDealsNotInputs.map((input) => normalizeSearch(input.value)).filter(Boolean);
  const baseRows = filteredDeals(scope, { countingIncluded: false })
    .filter((deal) => dealMatchesAndNotTerms(deal, andTerms, notTerms));
  tableFilterSources.allDeals = baseRows;
  const rows = applyColumnFilters("allDeals", baseRows);
  sortDealDetailRows(rows);
  const visibleRows = rows.slice(0, 300);

  if (els.allDealsTotalAmount) els.allDealsTotalAmount.textContent = money(sum(rows, (row) => row.amount));
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
        ${
          rows.length
            ? rows
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
                .join("")
            : `<tr><td colspan="13"><div class="empty">${escapeHtml(emptyText)}</div></td></tr>`
        }
      </tbody>
    </table>
  `;
}

function dealDetailHeader(key, label, numeric = false, extraClass = "") {
  const active = state.dealDetailSort.key === key;
  const icon = active ? (state.dealDetailSort.dir === "asc" ? "↑" : "↓") : "↕";
  const className = [numeric ? "num" : "", extraClass].filter(Boolean).join(" ");
  return `<th class="${className}"><div class="table-header-tools"><button type="button" class="sort-button" data-deal-detail-sort="${key}">${escapeHtml(label)} ${icon}</button>${columnFilterButton("allDeals", key)}</div></th>`;
}

const revenueLogic = window.RevenueLogic.createRevenueLogic({
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
  getRevenueExpectedFallback: () => state.revenueExpectedFallback,
});

function revenueDateParts(...values) {
  return revenueLogic.revenueDateParts(...values);
}

function revenuePeriodMonths(deal) {
  return revenueLogic.revenuePeriodMonths(deal);
}

function isOneTimeBilling(value) {
  return revenueLogic.isOneTimeBilling(value);
}

function isRecurringBilling(value) {
  return revenueLogic.isRecurringBilling(value);
}

function allocateRevenueAmount(amount, monthsCount) {
  return revenueLogic.allocateRevenueAmount(amount, monthsCount);
}

function revenueScheduleForDeal(deal) {
  return revenueLogic.revenueScheduleForDeal(deal);
}

function targetYearMonths(year) {
  return revenueLogic.targetYearMonths(year);
}

function buildRecurringTargetFromRevenue(schedules, targetYear) {
  return revenueLogic.buildRecurringTargetFromRevenue(schedules, targetYear);
}

function buildRevenueDistributionFromRevenue(schedules, targetYear) {
  const months = targetYearMonths(targetYear);
  const monthSet = new Set(months);
  const rowsByKey = new Map();
  const monthlyTotals = Object.fromEntries(months.map((month) => [month, 0]));
  const dealKeys = new Set();

  schedules.forEach((schedule) => {
    schedule.entries
      .filter((entry) => monthSet.has(entry.monthKey))
      .forEach((entry) => {
        const deal = schedule.deal;
        const key = `${deal.saleKey || deal.saleName}|${deal.category || "-"}|${deal.group || "-"}`;
        if (!rowsByKey.has(key)) {
          rowsByKey.set(key, {
            key,
            saleName: deal.saleName || "-",
            group: deal.group || "-",
            category: deal.category || "-",
            monthly: Object.fromEntries(months.map((month) => [month, 0])),
            deals: new Set(),
            companies: new Set(),
            detailMap: new Map(),
          });
        }
        const row = rowsByKey.get(key);
        const detailKey = dealDetailKey(deal);
        if (!row.detailMap.has(detailKey)) {
          row.detailMap.set(detailKey, {
            key: detailKey,
            company: deal.company || "-",
            dealName: deal.dealName || deal.id || "-",
            contractRange: schedule.contractRange || "-",
            billingType: deal.billingType || "-",
            monthly: Object.fromEntries(months.map((month) => [month, 0])),
            postingDatesByMonth: {},
            total: 0,
            sourceMonths: new Set(),
          });
        }
        const detail = row.detailMap.get(detailKey);
        detail.postingDatesByMonth[entry.monthKey] = addMonthsToDate(schedule.contractStart, entry.monthIndex - 1) || `${entry.monthKey}-01`;
        row.monthly[entry.monthKey] = roundMoney(row.monthly[entry.monthKey] + entry.amount);
        detail.monthly[entry.monthKey] = roundMoney(detail.monthly[entry.monthKey] + entry.amount);
        detail.total = roundMoney(detail.total + entry.amount);
        detail.sourceMonths.add(entry.monthKey);
        row.deals.add(detailKey);
        row.companies.add(deal.company || deal.dealName || deal.id || "-");
        monthlyTotals[entry.monthKey] = roundMoney(monthlyTotals[entry.monthKey] + entry.amount);
        dealKeys.add(detailKey);
      });
  });

  const rows = Array.from(rowsByKey.values())
    .map((row) => ({
      ...row,
      total: roundMoney(sum(months, (month) => row.monthly[month] || 0)),
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

  return {
    sourceYear: targetYear,
    targetYear,
    months,
    rows,
    monthlyTotals,
    total: roundMoney(sum(months, (month) => monthlyTotals[month] || 0)),
    dealCount: dealKeys.size,
  };
}

function schedulesWithMonths(schedules, monthSet) {
  return revenueLogic.schedulesWithMonths(schedules, monthSet);
}

function invoiceMatchesGlobalSearch(invoice, search = state.search) {
  const q = normalizeSearch(search);
  if (!q) return true;
  return (invoice.invoiceSearch || normalizeSearch(Object.values(invoice).join(" "))).includes(q);
}

function revenueDistributionDeals(scope) {
  return filteredDeals(scope, { period: false, countingIncluded: true, globalSearch: true })
    .filter((deal) => deal.status === "won");
}

function invoiceRowsForAnalysis(scope) {
  const selectedSale = scope.selectedSale;
  return (dashboardData.invoiceRows || [])
    .filter((invoice) => categoryMatches(invoice.category))
    .filter((invoice) => state.group === "all" || invoice.group === state.group)
    .filter((invoice) => state.sale === "all" || itemMatchesSaleRecord(invoice, selectedSale))
    .filter((invoice) => invoiceMatchesGlobalSearch(invoice));
}

function invoiceAnalysisMonthKey(invoice) {
  const serviceMonth = invoice.serviceMonthKey || parseInvoiceServiceStart(invoice.contractRange)?.monthKey || "";
  if (state.invoiceDateMode === "service") return serviceMonth || invoice.postingMonthKey || invoice.monthKey || "";
  return invoice.postingMonthKey || invoice.monthKey || invoice.serviceMonthKey || "";
}

function buildInvoiceDistributionFromInvoices(invoices, targetYear) {
  const months = targetYearMonths(targetYear);
  const monthSet = new Set(months);
  const rowsByKey = new Map();
  const monthlyTotals = Object.fromEntries(months.map((month) => [month, 0]));
  const invoiceKeys = new Set();

  invoices
    .map((invoice) => ({ ...invoice, analysisMonthKey: invoiceAnalysisMonthKey(invoice) }))
    .filter((invoice) => monthSet.has(invoice.analysisMonthKey))
    .forEach((invoice) => {
      const key = `${invoice.saleKey || invoice.saleName}|${invoice.category || "-"}|${invoice.group || "-"}`;
      if (!rowsByKey.has(key)) {
        rowsByKey.set(key, {
          key,
          saleName: invoice.saleName || "-",
          group: invoice.group || "-",
          category: invoice.category || "-",
          monthly: Object.fromEntries(months.map((month) => [month, 0])),
          deals: new Set(),
          companies: new Set(),
          detailMap: new Map(),
        });
      }
      const row = rowsByKey.get(key);
      const detailKey = `${invoice.documentNo || invoice.id}|${invoice.company}|${invoice.dealName}|${invoice.contractRange}|${invoice.amount}`;
      if (!row.detailMap.has(detailKey)) {
        row.detailMap.set(detailKey, {
          key: detailKey,
          company: invoice.company || "-",
          dealName: invoice.dealName || invoice.documentNo || "-",
          postingDate: invoice.postingDate || "-",
          contractRange: invoice.contractRange || invoice.postingDate || "-",
          billingType: [invoice.category, invoice.billingType].filter(Boolean).join(" / ") || "-",
          monthly: Object.fromEntries(months.map((month) => [month, 0])),
          total: 0,
          sourceMonths: new Set(),
        });
      }
      const detail = row.detailMap.get(detailKey);
      row.monthly[invoice.analysisMonthKey] = roundMoney(row.monthly[invoice.analysisMonthKey] + invoice.amount);
      detail.monthly[invoice.analysisMonthKey] = roundMoney(detail.monthly[invoice.analysisMonthKey] + invoice.amount);
      detail.total = roundMoney(detail.total + invoice.amount);
      detail.sourceMonths.add(invoice.analysisMonthKey);
      row.deals.add(detailKey);
      row.companies.add(invoice.company || invoice.dealName || invoice.documentNo || "-");
      monthlyTotals[invoice.analysisMonthKey] = roundMoney(monthlyTotals[invoice.analysisMonthKey] + invoice.amount);
      invoiceKeys.add(detailKey);
    });

  const rows = Array.from(rowsByKey.values())
    .map((row) => ({
      ...row,
      total: roundMoney(sum(months, (month) => row.monthly[month] || 0)),
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

  return {
    sourceYear: targetYear,
    targetYear,
    months,
    rows,
    monthlyTotals,
    total: roundMoney(sum(months, (month) => monthlyTotals[month] || 0)),
    dealCount: invoiceKeys.size,
  };
}

function buildRevenueAnalysis(scope) {
  const baseDeals = revenueDistributionDeals(scope);
  const targetBaseDeals = filteredDeals(scope, { period: false, countingIncluded: true, globalSearch: false })
    .filter((deal) => deal.status === "won");
  const schedules = baseDeals.map(revenueScheduleForDeal);
  const targetSchedules = targetBaseDeals.map(revenueScheduleForDeal);
  const revenueYear = revenueTargetYearFromScope(scope);
  const recurringTarget = buildRecurringTargetFromRevenue(targetSchedules, revenueYear);
  const revenueDistribution = buildRevenueDistributionFromRevenue(schedulesWithMonths(schedules, scope.monthSet), revenueYear);
  const invoiceAnalysis = buildInvoiceDistributionFromInvoices(invoiceRowsForAnalysis(scope), revenueYear);
  const selectedEntries = schedules
    .flatMap((schedule) =>
      schedule.entries.map((entry) => ({
        ...entry,
        schedule,
      })),
    )
    .filter((entry) => scope.monthSet.has(entry.monthKey));
  const monthlyRows = scope.months.map((month) => {
    const entries = selectedEntries.filter((entry) => entry.monthKey === month);
    const recurringEntries = entries.filter((entry) => isRecurringBilling(entry.deal.billingType));
    const target = recurringTarget.monthlyTotals[month] || 0;
    const recurringAmount = roundMoney(sum(recurringEntries, (entry) => entry.amount));
    return {
      month,
      amount: roundMoney(sum(entries, (entry) => entry.amount)),
      recurringAmount,
      target,
      achievement: target ? recurringAmount / target : null,
      deals: new Set(entries.map((entry) => dealDetailKey(entry.deal))).size,
      recurringDeals: new Set(recurringEntries.map((entry) => dealDetailKey(entry.deal))).size,
    };
  });
  const saleRows = Array.from(
    selectedEntries.reduce((map, entry) => {
      const key = entry.deal.saleKey || entry.deal.saleName;
      if (!map.has(key)) {
        map.set(key, {
          saleName: entry.deal.saleName || "-",
          group: entry.deal.group || "-",
          amount: 0,
          deals: new Set(),
        });
      }
      const row = map.get(key);
      row.amount += entry.amount;
      row.deals.add(dealDetailKey(entry.deal));
      return map;
    }, new Map()).values(),
  )
    .map((row) => ({ ...row, amount: roundMoney(row.amount), dealCount: row.deals.size }))
    .sort((a, b) => b.amount - a.amount);
  const detailRows = selectedEntries
    .map((entry) => revenueEntryRow(entry))
    .sort(sortRevenueRowsByTime);
  const exceptionRows = schedules
    .filter((schedule) => schedule.flags.length)
    .map((schedule) => revenueExceptionRow(schedule))
    .sort((a, b) => b.contractAmount - a.contractAmount);
  return {
    baseDeals,
    targetBaseDeals,
    schedules,
    targetSchedules,
    selectedEntries,
    monthlyRows,
    saleRows,
    detailRows,
    exceptionRows,
    recurringTarget,
    revenueDistribution,
    invoiceAnalysis,
  };
}

function revenueTargetYearFromScope(scope) {
  if (state.periodType === "year") return Number(state.periodYear) || Number(currentYearKey());
  if (state.periodType === "quarter") return Number(String(state.periodQuarter).slice(0, 4)) || Number(currentYearKey());
  if (state.periodType === "month") return Number(String(state.periodMonth).slice(0, 4)) || Number(currentYearKey());
  const firstMonth = scope.months[0] || state.startMonth || currentMonthKey();
  return Number(String(firstMonth).slice(0, 4)) || Number(currentYearKey());
}

function revenueEntryRow(entry) {
  const { deal, schedule } = entry;
  const revenueDate = `${entry.monthKey}-01`;
  return {
    ...deal,
    monthKey: entry.monthKey,
    monthLabel: monthLabel(entry.monthKey),
    revenueDate,
    saleName: deal.saleName || "-",
    companyDeal: `${deal.company || "-"} / ${deal.dealName || deal.id || "-"}`,
    category: deal.category || "-",
    billingType: deal.billingType || "-",
    contractRange: schedule.contractRange,
    contractAmount: deal.amount || 0,
    monthlyRevenue: entry.amount,
    revenueRule: schedule.revenueRule,
    flagsLabel: schedule.flags.length ? schedule.flags.join(", ") : "Ready",
    dealKey: dealDetailKey(deal),
    monthPosition: `${entry.monthIndex}/${entry.monthsCount}`,
  };
}

function sortRevenueRowsByTime(a, b) {
  const dateDiff =
    (parseDateValue(a.revenueDate)?.ts || 0) - (parseDateValue(b.revenueDate)?.ts || 0) ||
    (parseDateValue(a.contractStartDate)?.ts || 0) - (parseDateValue(b.contractStartDate)?.ts || 0) ||
    (parseDateValue(a.stageChangeDate)?.ts || 0) - (parseDateValue(b.stageChangeDate)?.ts || 0) ||
    (parseDateValue(a.expectedDate)?.ts || 0) - (parseDateValue(b.expectedDate)?.ts || 0);
  if (dateDiff) return dateDiff;
  return (
    a.saleName.localeCompare(b.saleName, "th", { numeric: true }) ||
    String(a.id || "").localeCompare(String(b.id || ""), "th", { numeric: true }) ||
    b.monthlyRevenue - a.monthlyRevenue
  );
}

function revenueExceptionRow(schedule) {
  const { deal } = schedule;
  return {
    id: deal.id || "-",
    saleName: deal.saleName || "-",
    companyDeal: `${deal.company || "-"} / ${deal.dealName || deal.id || "-"}`,
    category: deal.category || "-",
    contractAmount: deal.amount || 0,
    contractRange: schedule.contractRange,
    billingType: deal.billingType || "-",
    flagsLabel: schedule.flags.join(", "),
    revenueRule: schedule.revenueRule,
    dealKey: dealDetailKey(deal),
  };
}

function renderRevenueAnalysis(scope) {
  const analysis = buildRevenueAnalysis(scope);
  const performance = buildRevenuePerformance(analysis, scope);
  renderRevenueSubtabs();
  renderRevenueSummary(analysis, scope);
  renderRevenuePerformance(performance, scope);
  renderRevenueMonthlyChart(analysis);
  renderRevenueBySale(analysis);
  renderRecurringTargetFromRevenue(analysis.recurringTarget);
  renderRevenueDistribution(analysis.revenueDistribution);
  renderInvoiceAnalysis(analysis.invoiceAnalysis);
  renderRevenueDetailTable(analysis.detailRows);
  renderRevenueExceptionTable(analysis.exceptionRows);
}

function renderRevenueSubtabs() {
  els.revenueSubtabButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.revenueSubtab === state.revenueSubtab);
  });
  els.revenueSubtabPanels.forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.revenuePanel === state.revenueSubtab);
  });
  renderRevenueTargetVisibility();
  renderRevenueDistributionVisibility();
  renderRevenueDetailVisibility();
  renderTargetContributionVisibility();
  renderRevenueDataChecksVisibility();
}

function renderCollapsiblePanel(panel, body, toggle, visible) {
  panel?.classList.toggle("is-collapsed", !visible);
  if (body) body.hidden = !visible;
  if (toggle) {
    toggle.setAttribute("aria-expanded", String(visible));
    const icon = toggle.querySelector(".panel-toggle-icon");
    if (icon) icon.textContent = visible ? "-" : "+";
  }
}

function renderRevenueDetailVisibility() {
  renderCollapsiblePanel(els.revenueDetailPanel, els.revenueDetailBody, els.revenueDetailToggle, state.revenueDetailVisible);
}

function renderRevenueTargetVisibility() {
  renderCollapsiblePanel(els.revenueTargetPanel, els.revenueTargetBody, els.revenueTargetToggle, state.revenueTargetVisible);
}

function renderRevenueDistributionVisibility() {
  renderCollapsiblePanel(
    els.revenueDistributionPanel,
    els.revenueDistributionBody,
    els.revenueDistributionToggle,
    state.revenueDistributionVisible,
  );
}

function renderTargetContributionVisibility() {
  renderCollapsiblePanel(
    els.targetContributionPanel,
    els.targetContributionBody,
    els.targetContributionToggle,
    state.targetContributionVisible,
  );
}

function renderRevenueDataChecksVisibility() {
  renderCollapsiblePanel(
    els.revenueDataChecksPanel,
    els.revenueDataChecksBody,
    els.revenueDataChecksToggle,
    state.revenueDataChecksVisible,
  );
}

function buildRevenuePerformance(analysis, scope) {
  const recurringEntries = analysis.selectedEntries.filter((entry) => isRecurringBilling(entry.deal.billingType));
  const targetRows = analysis.recurringTarget.rows || [];
  const ownerMonthTargets = new Map();
  const ownerRows = new Map();
  const ownerKeyFor = (saleName, group) => `${saleName || "-"}|${group || "-"}`;

  targetRows.forEach((row) => {
    const ownerKey = ownerKeyFor(row.saleName, row.group);
    if (!ownerRows.has(ownerKey)) {
      ownerRows.set(ownerKey, {
        saleName: row.saleName || "-",
        group: row.group || "-",
        actual: 0,
        target: 0,
        gap: 0,
        deals: new Set(),
      });
    }
    const owner = ownerRows.get(ownerKey);
    scope.months.forEach((month) => {
      const amount = row.monthly?.[month] || 0;
      owner.target = roundMoney(owner.target + amount);
      ownerMonthTargets.set(`${ownerKey}|${month}`, roundMoney((ownerMonthTargets.get(`${ownerKey}|${month}`) || 0) + amount));
    });
  });

  recurringEntries.forEach((entry) => {
    const deal = entry.deal;
    const ownerKey = ownerKeyFor(deal.saleName, deal.group);
    if (!ownerRows.has(ownerKey)) {
      ownerRows.set(ownerKey, {
        saleName: deal.saleName || "-",
        group: deal.group || "-",
        actual: 0,
        target: 0,
        gap: 0,
        deals: new Set(),
      });
    }
    const owner = ownerRows.get(ownerKey);
    owner.actual = roundMoney(owner.actual + entry.amount);
    owner.deals.add(dealDetailKey(deal));
  });

  const monthlyRows = scope.months.map((month) => {
    const actual = sum(
      recurringEntries.filter((entry) => entry.monthKey === month),
      (entry) => entry.amount,
    );
    const target = analysis.recurringTarget.monthlyTotals[month] || 0;
    const dealCount = new Set(recurringEntries.filter((entry) => entry.monthKey === month).map((entry) => dealDetailKey(entry.deal))).size;
    return {
      month,
      actual: roundMoney(actual),
      target,
      gap: roundMoney(actual - target),
      achievement: target ? actual / target : null,
      dealCount,
    };
  });

  const saleRows = Array.from(ownerRows.values())
    .map((row) => ({
      ...row,
      gap: roundMoney(row.actual - row.target),
      achievement: row.target ? row.actual / row.target : null,
      dealCount: row.deals.size,
    }))
    .filter((row) => row.actual || row.target)
    .sort((a, b) => b.actual - a.actual || b.target - a.target || a.saleName.localeCompare(b.saleName, "th", { numeric: true }));

  const detailRows = recurringEntries
    .map((entry) => {
      const row = revenueEntryRow(entry);
      const ownerKey = ownerKeyFor(row.saleName, row.group);
      const monthTarget = ownerMonthTargets.get(`${ownerKey}|${entry.monthKey}`) || 0;
      const monthlyRevenue = row.monthlyRevenue || 0;
      return {
        ...row,
        performanceTarget: monthTarget,
        performanceContribution: monthTarget ? monthlyRevenue / monthTarget : null,
        performanceGap: roundMoney(monthlyRevenue - monthTarget),
      };
    })
    .sort(sortRevenueRowsByTime);

  const totalActual = roundMoney(sum(monthlyRows, (row) => row.actual));
  const totalTarget = roundMoney(sum(monthlyRows, (row) => row.target));
  return {
    monthlyRows,
    saleRows,
    detailRows,
    totalActual,
    totalTarget,
    totalGap: roundMoney(totalActual - totalTarget),
    achievement: totalTarget ? totalActual / totalTarget : null,
    dealCount: new Set(recurringEntries.map((entry) => dealDetailKey(entry.deal))).size,
  };
}

function renderRevenuePerformance(performance, scope) {
  renderRevenuePerformanceSummary(performance, scope);
  renderRevenuePerformanceChart(performance);
  renderRevenuePerformanceBySale(performance);
  renderRevenuePerformanceDetail(performance);
}

function renderRevenuePerformanceSummary(performance, scope) {
  if (!els.revenuePerformanceSummary) return;
  const cards = [
    ["Actual Recurring Revenue", compactMoney(performance.totalActual), `${monthLabel(scope.months[0] || "")} - ${monthLabel(scope.months[scope.months.length - 1] || "")}`],
    ["Recurring Target", compactMoney(performance.totalTarget), "ฐานจาก Revenue ปีก่อนหน้า"],
    ["Achievement", performance.achievement == null ? "-" : percent(performance.achievement), `Gap ${signedMoney(performance.totalGap)}`],
    ["Active Deals", performance.dealCount.toLocaleString("th-TH"), "เฉพาะ Billing = Recurring"],
  ];
  els.revenuePerformanceSummary.innerHTML = cards
    .map(
      ([label, value, note]) => `
        <div class="lead-summary-card revenue-card">
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(value)}</strong>
          <small>${escapeHtml(note)}</small>
        </div>
      `,
    )
    .join("");
}

function renderRevenuePerformanceChart(performance) {
  if (!els.revenuePerformanceChart) return;
  const rows = performance.monthlyRows.filter((row) => row.actual || row.target || performance.monthlyRows.length <= 12);
  const maxAmount = Math.max(1, ...rows.flatMap((row) => [row.actual || 0, row.target || 0]));
  els.revenuePerformanceChart.innerHTML = rows.length
    ? `
      <div class="chart-legend-inline renew revenue-performance-legend">
        <span class="legend-target">Recurring Target</span>
        <span class="legend-won" style="--legend-color:var(--renew)">Won / Actual</span>
      </div>
      ${rows
        .map(
          (row) => `
            <div class="revenue-performance-row">
              <div class="lead-week-name">${escapeHtml(monthLabel(row.month))}</div>
              <div class="revperf-bar-space">
                <div class="revperf-target-base" style="width:${row.target ? Math.min(100, Math.max(2, ((row.target || 0) / maxAmount) * 100)) : 0}%" title="Target ${money(row.target || 0)}"></div>
                <div class="revperf-actual-bar" style="width:${row.actual ? Math.min(100, Math.max(2, ((row.actual || 0) / maxAmount) * 100)) : 0}%" title="Won ${money(row.actual || 0)}"></div>
                <div class="revperf-target-outline" style="width:${row.target ? Math.min(100, Math.max(2, ((row.target || 0) / maxAmount) * 100)) : 0}%"></div>
              </div>
              <div class="lead-week-value">
                <span class="muted">Target ${money(row.target || 0)}</span><br>
                <span class="muted">Won ${money(row.actual || 0)}</span>
                <span class="achievement-pill revenue-achievement-inline ${achievementClass(row.achievement)}">${row.achievement == null ? "-" : percent(row.achievement)}</span><br>
                <span class="muted">Gap ${signedMoney(row.gap || 0)} · ${row.dealCount.toLocaleString("th-TH")} deals</span>
              </div>
            </div>
          `,
        )
        .join("")}
    `
    : `<div class="empty">ไม่มี Recurring revenue ในช่วงเวลาที่เลือก</div>`;
}

function renderRevenuePerformanceBySale(performance) {
  if (!els.revenuePerformanceBySale) return;
  if (!performance.saleRows.length) {
    els.revenuePerformanceBySale.innerHTML = `<div class="empty">ไม่มีข้อมูลราย Sale สำหรับช่วงเวลาที่เลือก</div>`;
    return;
  }
  const maxAmount = Math.max(1, ...performance.saleRows.map((row) => row.actual));
  els.revenuePerformanceBySale.innerHTML = `
    ${performance.saleRows
      .slice(0, 15)
      .map(
        (row) => `
          <div class="rank-row revenue-performance-sale-row">
            <div class="rank-name">
              ${escapeHtml(row.saleName)}<br>
              <span class="muted">${escapeHtml(row.group)}</span>
            </div>
            <div class="progress-track">
              <div class="progress-fill revenue-fill" style="width:${Math.max(2, (row.actual / maxAmount) * 100)}%"></div>
            </div>
            <div class="rank-value">
              ${money(row.actual)}<br>
              <span class="muted">Target ${money(row.target)}</span><br>
              <span class="${row.gap < 0 ? "negative" : "positive"}">${signedMoney(row.gap)}</span>
            </div>
          </div>
        `,
      )
      .join("")}
  `;
}

function renderRevenuePerformanceDetail(performance) {
  if (!els.revenuePerformanceDetail) return;
  const rows = performance.detailRows.slice(0, 500);
  if (!rows.length) {
    els.revenuePerformanceDetail.innerHTML = `<div class="empty">ไม่มีรายละเอียด Recurring revenue ในช่วงเวลาที่เลือก</div>`;
    return;
  }
  els.revenuePerformanceDetail.innerHTML = `
    <table class="revenue-performance-detail-table">
      <thead>
        <tr>
          <th>Revenue Month</th>
          <th>Sale</th>
          <th>Company / Deal</th>
          <th class="num">Monthly Revenue</th>
          <th class="num">Monthly Target</th>
          <th class="num">Contribution</th>
          <th class="num">Gap</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (row) => `
              <tr class="clickable-row" data-deal-key="${escapeHtml(row.dealKey)}">
                <td><strong>${escapeHtml(row.monthLabel)}</strong><br><span class="muted">${escapeHtml(row.revenueDate)} | ${escapeHtml(row.monthPosition)}</span></td>
                <td><strong>${escapeHtml(row.saleName)}</strong><br><span class="muted">${escapeHtml(row.group)}</span></td>
                <td class="company-cell"><strong>${escapeHtml(row.company || "-")}</strong><br><span class="muted">${escapeHtml(row.dealName || "-")}</span></td>
                <td>${escapeHtml(row.category)}</td>
                <td class="num"><strong>${money(row.monthlyRevenue || 0)}</strong></td>
                <td class="num">${money(row.performanceTarget || 0)}</td>
                <td class="num">${row.performanceContribution == null ? "-" : percent(row.performanceContribution)}</td>
                <td class="num ${row.performanceGap < 0 ? "negative" : "positive"}">${signedMoney(row.performanceGap || 0)}</td>
              </tr>
            `,
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function renderRevenueSummary(analysis, scope) {
  const scheduledAmount = sum(analysis.selectedEntries, (entry) => entry.amount);
  const contractAmount = sum(analysis.baseDeals, (deal) => deal.amount);
  const activeMonths = analysis.monthlyRows.filter((row) => row.amount).length;
  const recurringActual = sum(analysis.monthlyRows, (row) => row.recurringAmount || 0);
  const recurringTarget = sum(analysis.monthlyRows, (row) => row.target || 0);
  const recurringAchievement = recurringTarget ? recurringActual / recurringTarget : null;
  const cards = [
    ["Revenue in Period", compactMoney(scheduledAmount), `${monthLabel(scope.months[0] || "")} - ${monthLabel(scope.months[scope.months.length - 1] || "")}`],
    ["Won / Pre-WON Contracts", analysis.baseDeals.length.toLocaleString("th-TH"), compactMoney(contractAmount)],
    ["Recurring Achievement", recurringAchievement == null ? "-" : percent(recurringAchievement), `${compactMoney(recurringActual)} / ${compactMoney(recurringTarget)}`],
    ["Data Checks", analysis.exceptionRows.length.toLocaleString("th-TH"), "ควรตรวจสอบก่อนส่งต่อบัญชี/การเงิน"],
  ];
  els.revenueSummary.innerHTML = cards
    .map(
      ([label, value, note]) => `
        <div class="lead-summary-card revenue-card">
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(value)}</strong>
          <small>${escapeHtml(note)}</small>
        </div>
      `,
    )
    .join("");
}

function renderRevenueMonthlyChart(analysis) {
  const maxAmount = Math.max(1, ...analysis.monthlyRows.flatMap((row) => [row.amount || 0, row.recurringAmount || 0, row.target || 0]));
  const rows = analysis.monthlyRows.filter((row) => row.amount || row.target || analysis.monthlyRows.length <= 12);
  els.revenueMonthlyChart.innerHTML = rows.length
    ? `
      <div class="chart-legend-inline renew revenue-performance-legend">
        <span class="legend-target">Recurring Target</span>
        <span class="legend-won" style="--legend-color:var(--renew)">Actual Recurring</span>
      </div>
      ${rows
        .map(
          (row) => `
            <div class="revenue-month-row">
              <div class="lead-week-name">${escapeHtml(monthLabel(row.month))}</div>
              <div class="revperf-bar-space">
                <div class="revperf-target-base" style="width:${row.target ? Math.min(100, Math.max(2, ((row.target || 0) / maxAmount) * 100)) : 0}%" title="Recurring Target ${money(row.target || 0)}"></div>
                <div class="revperf-actual-bar" style="width:${row.recurringAmount ? Math.min(100, Math.max(2, ((row.recurringAmount || 0) / maxAmount) * 100)) : 0}%" title="Actual Recurring ${money(row.recurringAmount || 0)}"></div>
                <div class="revperf-target-outline" style="width:${row.target ? Math.min(100, Math.max(2, ((row.target || 0) / maxAmount) * 100)) : 0}%"></div>
              </div>
              <div class="lead-week-value">
                <span class="muted">Target ${money(row.target || 0)}</span><br>
                <span class="muted">Actual ${money(row.recurringAmount || 0)}</span>
                <span class="achievement-pill revenue-achievement-inline ${achievementClass(row.achievement)}">${row.achievement == null ? "-" : percent(row.achievement)}</span><br>
                <span class="muted">Gap ${signedMoney((row.recurringAmount || 0) - (row.target || 0))} · Total Revenue ${money(row.amount || 0)} · ${row.deals.toLocaleString("th-TH")} deals</span>
              </div>
            </div>
          `,
        )
        .join("")}
    `
    : `<div class="empty">ไม่มี Revenue schedule ในช่วงเวลาที่เลือก</div>`;
}

function achievementClass(value) {
  if (value == null) return "neutral";
  if (value >= 1) return "good";
  if (value >= 0.8) return "warn";
  return "danger";
}

function renderRevenueBySale(analysis) {
  const maxAmount = Math.max(1, ...analysis.saleRows.map((row) => row.amount));
  els.revenueBySale.innerHTML = analysis.saleRows.length
    ? analysis.saleRows
        .slice(0, 15)
        .map(
          (row) => `
            <div class="rank-row">
              <div class="rank-name">${escapeHtml(row.saleName)}<br><span class="muted">${escapeHtml(row.group)}</span></div>
              <div class="progress-track"><div class="progress-fill revenue-fill" style="width:${Math.max(2, (row.amount / maxAmount) * 100)}%"></div></div>
              <div class="rank-value">${money(row.amount)}<br>${row.dealCount.toLocaleString("th-TH")} deals</div>
            </div>
          `,
        )
        .join("")
    : `<div class="empty">ไม่มี Revenue ตามเงื่อนไขที่เลือก</div>`;
}

function renderRecurringTargetFromRevenue(target) {
  if (!els.revenueTargetSummary || !els.revenueTargetTable) return;
  if (els.revenueTargetTitle) els.revenueTargetTitle.textContent = `${target.targetYear} Recurring Target from ${target.sourceYear} Revenue`;
  if (els.revenueTargetDescription) {
    els.revenueTargetDescription.textContent = `ประเมิน Target ปี ${target.targetYear} จากกระแสรายได้ปี ${target.sourceYear} เฉพาะ Billing = Recurring โดยเลื่อนไปเดือนเดียวกันของปีถัดไป`;
  }
  const avgMonthly = target.months.length ? target.total / target.months.length : 0;
  const activeMonths = target.months.filter((month) => target.monthlyTotals[month]).length;
  const cards = [
    [`${target.targetYear} Recurring Target`, compactMoney(target.total), `ฐานจาก Revenue ${target.sourceYear}`],
    ["Recurring Deals", target.dealCount.toLocaleString("th-TH"), "เฉพาะ Billing = Recurring"],
    ["Avg Target / Month", compactMoney(avgMonthly), `${activeMonths.toLocaleString("th-TH")} active months`],
    ["Target Rows", target.rows.length.toLocaleString("th-TH"), "Sale + Type"],
  ];
  els.revenueTargetSummary.innerHTML = cards
    .map(
      ([label, value, note]) => `
        <div class="lead-summary-card revenue-card">
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(value)}</strong>
          <small>${escapeHtml(note)}</small>
        </div>
      `,
    )
    .join("");
  els.revenueTargetTable.innerHTML = recurringTargetTable(target, {
    expandedMap: state.revenueTargetExpanded,
    toggleAttr: "data-revenue-target-toggle",
    emptyMessage: `ไม่มี recurring revenue ปี ${target.sourceYear} สำหรับประเมิน Target ${target.targetYear}`,
    formatContractRange: displayDateRangeDmy,
  });
}

function renderRevenueDistribution(distribution) {
  if (!els.revenueDistributionSummary || !els.revenueDistributionTable) return;
  if (els.revenueDistributionTitle) els.revenueDistributionTitle.textContent = `${distribution.targetYear} Revenue Distribution`;
  if (els.revenueDistributionDescription) {
    els.revenueDistributionDescription.textContent = `กระจาย Revenue ปี ${distribution.targetYear} เป็นรายเดือนตามสัญญาของ Deals ที่อยู่ในขอบเขตการวิเคราะห์`;
  }
  const avgMonthly = distribution.months.length ? distribution.total / distribution.months.length : 0;
  const activeMonths = distribution.months.filter((month) => distribution.monthlyTotals[month]).length;
  const cards = [
    [`${distribution.targetYear} Revenue`, compactMoney(distribution.total), "ยอด Revenue ที่กระจายตามเดือน"],
    ["Revenue Deals", distribution.dealCount.toLocaleString("th-TH"), "Deals ที่มี revenue ในปีที่เลือก"],
    ["Avg Revenue / Month", compactMoney(avgMonthly), `${activeMonths.toLocaleString("th-TH")} active months`],
    ["Distribution Rows", distribution.rows.length.toLocaleString("th-TH"), "Sale + Type"],
  ];
  els.revenueDistributionSummary.innerHTML = cards
    .map(
      ([label, value, note]) => `
        <div class="lead-summary-card revenue-card">
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(value)}</strong>
          <small>${escapeHtml(note)}</small>
        </div>
      `,
    )
    .join("");
  els.revenueDistributionTable.innerHTML = recurringTargetTable(distribution, {
    expandedMap: state.revenueDistributionExpanded,
    toggleAttr: "data-revenue-distribution-toggle",
    emptyMessage: `ไม่มี Revenue ปี ${distribution.targetYear} ในขอบเขตการวิเคราะห์`,
    formatContractRange: displayDateRangeDmy,
  });
}

function renderInvoiceAnalysis(invoiceAnalysis) {
  if (!els.invoiceAnalysisSummary || !els.invoiceAnalysisTable) return;
  const isServiceMode = state.invoiceDateMode === "service";
  els.invoiceDateModeButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.invoiceDateMode === state.invoiceDateMode);
  });
  if (els.invoiceAnalysisTitle) els.invoiceAnalysisTitle.textContent = `${invoiceAnalysis.targetYear} Invoice Analysis`;
  if (els.invoiceAnalysisDescription) {
    els.invoiceAnalysisDescription.textContent = isServiceMode
      ? `กระจายยอด Invoice ปี ${invoiceAnalysis.targetYear} เป็นรายเดือนจาก Service Period โดยใช้ Amount ก่อน VAT และ fallback เป็น Posting Date เมื่อไม่มี Period`
      : `กระจายยอด Invoice ปี ${invoiceAnalysis.targetYear} เป็นรายเดือนจาก Posting Date โดยใช้ Amount ก่อน VAT`;
  }
  const avgMonthly = invoiceAnalysis.months.length ? invoiceAnalysis.total / invoiceAnalysis.months.length : 0;
  const activeMonths = invoiceAnalysis.months.filter((month) => invoiceAnalysis.monthlyTotals[month]).length;
  const sourceFile = dashboardData.metadata?.sourceFiles?.invoice || "ยังไม่ได้ Upload Invoice CSV";
  const cards = [
    [`${invoiceAnalysis.targetYear} Invoice Amount`, compactMoney(invoiceAnalysis.total), isServiceMode ? "ฐานเดือน: Service Period" : "ฐานเดือน: Posting Date"],
    ["Invoice Rows", invoiceAnalysis.dealCount.toLocaleString("th-TH"), sourceFile],
    ["Avg Invoice / Month", compactMoney(avgMonthly), `${activeMonths.toLocaleString("th-TH")} active months`],
    ["Distribution Rows", invoiceAnalysis.rows.length.toLocaleString("th-TH"), "Sale + Type"],
  ];
  els.invoiceAnalysisSummary.innerHTML = cards
    .map(
      ([label, value, note]) => `
        <div class="lead-summary-card revenue-card">
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(value)}</strong>
          <small>${escapeHtml(note)}</small>
        </div>
      `,
    )
    .join("");
  els.invoiceAnalysisTable.innerHTML = recurringTargetTable(invoiceAnalysis, {
    expandedMap: state.invoiceAnalysisExpanded,
    toggleAttr: "data-invoice-analysis-toggle",
    emptyMessage: `ไม่มี Invoice ปี ${invoiceAnalysis.targetYear} ในขอบเขตการวิเคราะห์`,
    formatContractRange: displayDateRangeDmy,
  });
}

function recurringTargetTable(target, options = {}) {
  if (!target.rows.length) return `<div class="empty">${escapeHtml(options.emptyMessage || `ไม่มี recurring revenue ปี ${target.sourceYear} สำหรับประเมิน Target ${target.targetYear}`)}</div>`;
  const monthHeaders = target.months.map((month) => `<th class="num">${escapeHtml(monthLabel(month).replace(` ${target.targetYear}`, ""))}</th>`).join("");
  const columnCount = target.months.length + 6;
  const totalRow = `
    <tr class="total-row">
      <td colspan="4"><strong>Total ${target.targetYear}</strong></td>
      ${target.months.map((month) => `<td class="num"><strong>${money(target.monthlyTotals[month] || 0)}</strong></td>`).join("")}
      <td class="num"><strong>${money(target.total)}</strong></td>
      <td class="num">${target.dealCount.toLocaleString("th-TH")}</td>
    </tr>
  `;
  return `
    <table class="target-table">
      <thead>
        <tr>
          <th>Sale</th>
          <th>Customer/Project</th>
          <th>Sale Group</th>
          <th>Last Deal Type</th>
          ${monthHeaders}
          <th class="num">Total</th>
          <th class="num">Deals</th>
        </tr>
      </thead>
      <tbody>
        ${target.rows.map((row) => recurringTargetRow(row, target, columnCount, options)).join("")}
        ${totalRow}
      </tbody>
    </table>
  `;
}

function recurringTargetRow(row, target, columnCount, options = {}) {
  const rowKey = row.key || `${row.saleName}|${row.group}|${row.category}`;
  const expandedMap = options.expandedMap || state.revenueTargetExpanded;
  const toggleAttr = options.toggleAttr || "data-revenue-target-toggle";
  const expanded = Boolean(expandedMap[rowKey]);
  const detailCount = (row.details || []).length;
  return `
    <tr class="${expanded ? "target-parent-row is-expanded" : "target-parent-row"}">
      <td>
        <button type="button" class="target-sale-toggle" ${toggleAttr}="${escapeHtml(rowKey)}" aria-expanded="${expanded}">
          <span class="target-toggle-mark" aria-hidden="true">${expanded ? "-" : "+"}</span>
          <span class="target-toggle-name">${escapeHtml(row.saleName)}</span>
          <small>${detailCount.toLocaleString("th-TH")} deals</small>
        </button>
      </td>
      <td class="muted"></td>
      <td>${escapeHtml(row.group)}</td>
      <td>${escapeHtml(row.category)}</td>
      ${target.months.map((month) => `<td class="num">${row.monthly[month] ? money(row.monthly[month]) : "-"}</td>`).join("")}
      <td class="num"><strong>${money(row.total)}</strong></td>
      <td class="num">${row.dealCount.toLocaleString("th-TH")}</td>
    </tr>
    ${expanded ? recurringTargetDetailRow(row, target, columnCount, options) : ""}
  `;
}

function recurringTargetDetailRow(row, target, columnCount, options = {}) {
  const details = row.details || [];
  const formatContractRange = options.formatContractRange || ((value) => value);
  if (!details.length) {
    return `
      <tr class="target-detail-row">
        <td colspan="${columnCount}">
          <div class="target-detail-panel empty">ไม่มีรายละเอียด Deal สำหรับแถวนี้</div>
        </td>
      </tr>
    `;
  }
  return details
    .map(
      (detail) => `
        <tr class="target-detail-inline-row">
          <td><strong>${escapeHtml(row.saleName)}</strong></td>
          <td class="target-detail-company">
            <strong>${escapeHtml(detail.company)}</strong>
            <span>${escapeHtml(detail.dealName)}</span>
          </td>
          <td>${escapeHtml(formatContractRange(detail.contractRange || "-"))}</td>
          <td>${escapeHtml(detail.billingType || "-")}</td>
          ${target.months.map((month) => `<td class="num">${detail.monthly?.[month] ? money(detail.monthly[month]) : "-"}</td>`).join("")}
          <td class="num"><strong>${money(detail.total || 0)}</strong></td>
          <td class="num"></td>
        </tr>
      `,
    )
    .join("");
}

function csvCell(value) {
  const text = String(value ?? "");
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function csvDate(value) {
  return displayDateDmy(value);
}

function csvDateRange(value) {
  return displayDateRangeDmy(value);
}

function exportRevenueTargetCsv() {
  const target = buildRevenueAnalysis(calcScope()).recurringTarget;
  exportRevenueBreakdownCsv(target, `${target.targetYear}-recurring-target-from-${target.sourceYear}-revenue.csv`, {
    formatContractRange: csvDateRange,
  });
}

function exportRevenueDistributionCsv() {
  const distribution = buildRevenueAnalysis(calcScope()).revenueDistribution;
  exportRevenueBreakdownCsv(distribution, `${distribution.targetYear}-revenue-distribution.csv`, {
    includePostingDate: true,
    splitByMonthPostingDate: true,
    formatContractRange: csvDateRange,
  });
}

function exportInvoiceAnalysisCsv() {
  const invoiceAnalysis = buildRevenueAnalysis(calcScope()).invoiceAnalysis;
  exportRevenueBreakdownCsv(invoiceAnalysis, `${invoiceAnalysis.targetYear}-invoice-analysis-${state.invoiceDateMode}.csv`, {
    includePostingDate: true,
    formatContractRange: csvDateRange,
  });
}

function exportRevenueBreakdownCsv(target, fileName, options = {}) {
  const formatContractRange = options.formatContractRange || ((value) => value);
  const headers = [
    "Sale",
    "Customer/Project",
    ...(options.includePostingDate ? ["Posting Date"] : []),
    "Sale Group / Contract",
    "Last Deal Type / Billing",
    ...target.months.map((month) => monthLabel(month)),
    "Total",
  ];
  const exportRows = target.rows.flatMap((row) =>
    (row.details || []).flatMap((detail) => {
      const detailName = `${detail.company || "-"}${detail.dealName ? ` | ${detail.dealName}` : ""}`;
      const baseColumns = [row.saleName, detailName];
      if (options.splitByMonthPostingDate) {
        return target.months
          .filter((month) => detail.monthly?.[month])
          .map((month) => {
            const amount = detail.monthly?.[month] || 0;
            const postingDate = detail.postingDatesByMonth?.[month] || `${month}-01`;
            return {
              customerProject: detailName,
              postingDate,
              amount,
              row: [
                ...baseColumns,
                csvDate(postingDate),
                formatContractRange(detail.contractRange || "-"),
                detail.billingType || "-",
                ...target.months.map((targetMonth) => (targetMonth === month ? amount : 0)),
                amount,
              ],
            };
          });
      }
      const postingDate = detail.postingDate || parseDateValue(detail.contractRange)?.date || "";
      const amount = detail.total || 0;
      return [
        {
          customerProject: detailName,
          postingDate,
          amount,
          row: [
            ...baseColumns,
            ...(options.includePostingDate ? [csvDate(postingDate)] : []),
            formatContractRange(detail.contractRange || "-"),
            detail.billingType || "-",
            ...target.months.map((month) => detail.monthly?.[month] || 0),
            amount,
          ],
        },
      ];
    }),
  );
  const rows = exportRows
    .sort(exportCustomerPostingAmountSort)
    .map((item) => item.row);
  const csv = [headers, ...rows].map((row) => row.map(csvCell).join(",")).join("\r\n");
  downloadTextFile(fileName, `\uFEFF${csv}`, "text/csv;charset=utf-8");
}

function exportCustomerPostingAmountSort(a, b) {
  const customer = String(a.customerProject || "").localeCompare(String(b.customerProject || ""), "th", { numeric: true });
  if (customer) return customer;
  const date = (parseDateValue(a.postingDate)?.ts || 0) - (parseDateValue(b.postingDate)?.ts || 0);
  if (date) return date;
  return (Number(a.amount) || 0) - (Number(b.amount) || 0);
}

function exportRevenueDetailCsv() {
  const rows = filteredRevenueDetailRows(buildRevenueAnalysis(calcScope()).detailRows).sort((a, b) => {
    const customerA = `${a.company || ""} ${a.dealName || ""}`;
    const customerB = `${b.company || ""} ${b.dealName || ""}`;
    const customer = customerA.localeCompare(customerB, "th", { numeric: true });
    if (customer) return customer;
    const date = (parseDateValue(a.revenueDate)?.ts || 0) - (parseDateValue(b.revenueDate)?.ts || 0);
    if (date) return date;
    return (Number(a.monthlyRevenue) || 0) - (Number(b.monthlyRevenue) || 0);
  });
  const headers = [
    "Deal ID",
    "Revenue Month",
    "Revenue Date",
    "ยอดเดือน",
    "ระยะสัญญา",
    "Sale",
    "Sale Group",
    "Company",
    "Deal",
    "Type",
    "Billing",
    "Contract",
    "Expected Close Date",
    "Monthly Revenue",
    "Contract Amount",
    "Rule",
    "Checks",
  ];
  const body = rows.map((row) => {
    const [revenueMonthNo, contractMonths] = splitMonthPosition(row.monthPosition);
    return [
      row.id || "",
      row.monthLabel || "",
      csvDate(row.revenueDate),
      revenueMonthNo,
      contractMonths,
      row.saleName || "",
      row.group || "",
      row.company || "",
      row.dealName || "",
      row.category || "",
      row.billingType || "",
      csvDateRange(row.contractRange),
      csvDate(row.expectedDate),
      row.monthlyRevenue || 0,
      row.contractAmount || 0,
      row.revenueRule || "",
      row.flagsLabel || "",
    ];
  });
  const csv = [headers, ...body].map((row) => row.map(csvCell).join(",")).join("\r\n");
  downloadTextFile(`revenue-detail-filtered.csv`, `\uFEFF${csv}`, "text/csv;charset=utf-8");
}

function splitMonthPosition(value) {
  const match = String(value || "").match(/^(\d+)\s*\/\s*(\d+)$/);
  return match ? [match[1], match[2]] : ["", ""];
}

function downloadTextFile(fileName, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

function renderRevenueDetailTable(detailRows) {
  const rows = filteredRevenueDetailRows(detailRows);
  if (els.revenueDetailMonthlyTotal) els.revenueDetailMonthlyTotal.textContent = money(sum(rows, (row) => row.monthlyRevenue || 0));
  els.revenueDetailTable.innerHTML = revenueDetailTable(rows.slice(0, 500));
}

function filteredRevenueDetailRows(detailRows) {
  const andTerms = els.revenueAndInputs.map((input) => normalizeSearch(input.value)).filter(Boolean);
  const notTerms = els.revenueNotInputs.map((input) => normalizeSearch(input.value)).filter(Boolean);
  const baseRows = detailRows.filter((row) => revenueRowMatchesAndNotTerms(row, andTerms, notTerms));
  tableFilterSources.revenue = baseRows;
  const rows = applyColumnFilters("revenue", baseRows).slice();
  rows.sort(sortRevenueRowsByTime);
  return rows;
}

function revenueRowMatchesAndNotTerms(row, andTerms, notTerms) {
  const text = revenueSearchText(row);
  const tokens = revenueSearchTokens(row);
  if (andTerms.some((term) => !transactionTermMatches(term, text, tokens))) return false;
  if (notTerms.some((term) => transactionTermMatches(term, text, tokens))) return false;
  return true;
}

function revenueSearchText(row) {
  const counting = countingSearchValues(row).join(" ");
  const risk = riskFilterLabel(row);
  const expectedCloseDate = displayDate(row.expectedDate);
  return normalizeSearch(
    `${row.id || ""} ${row.saleName || ""} ${row.responsible || ""} ${row.group || ""} ${row.company || ""} ${row.dealName || ""} ${row.pipeline || ""} ${row.stage || ""} ${row.rawStage || ""} ${row.dealType || ""} ${row.product || ""} ${row.category || ""} ${row.billingType || ""} ${performanceBucket(row)} ${row.status || ""} ${counting} ${risk} ${row.expectedDate || ""} ${expectedCloseDate} ${row.stageChangeDate || ""} ${row.monthKey || ""} ${row.monthLabel || ""} ${row.revenueDate || ""} ${row.revenueRule || ""} ${row.flagsLabel || ""} ${row.monthPosition || ""} ${money(row.monthlyRevenue || 0)} ${money(row.contractAmount || 0)}`,
  );
}

function revenueSearchTokens(row) {
  const expectedCloseDate = displayDate(row.expectedDate);
  return new Set(
    [
      row.id,
      row.saleName,
      row.responsible,
      row.company,
      row.dealName,
      row.pipeline,
      row.stage,
      row.rawStage,
      row.dealType,
      row.product,
      row.category,
      row.status,
      ...countingSearchValues(row),
      riskFilterLabel(row),
      row.expectedDate,
      expectedCloseDate,
      row.stageChangeDate,
      performanceBucket(row),
      row.monthKey,
      row.monthLabel,
      row.revenueDate,
      row.revenueRule,
      row.flagsLabel,
      row.billingType,
      row.monthPosition,
      money(row.monthlyRevenue || 0),
      money(row.contractAmount || 0),
    ]
      .map(normalizeSearch)
      .flatMap((value) => value.split(/[^\p{L}\p{N}]+/u))
      .filter(Boolean),
  );
}

function revenueDetailTable(rows) {
  return `
    <table class="revenue-detail-table">
      <colgroup>
        <col class="rev-month-col" />
        <col class="rev-sale-col" />
        <col class="rev-company-col" />
        <col class="rev-type-col" />
        <col class="rev-billing-col" />
        <col class="rev-contract-col" />
        <col class="rev-monthly-col" />
        <col class="rev-contract-amount-col" />
        <col class="rev-rule-col" />
        <col class="rev-checks-col" />
      </colgroup>
      <thead>
        <tr>
          ${revenueHeader("monthLabel", "Revenue Month")}
          ${revenueHeader("saleName", "Sale")}
          ${revenueHeader("companyDeal", "Company / Deal", false, "company-cell")}
          ${revenueHeader("category", "Type")}
          ${revenueHeader("billingType", "Billing")}
          ${revenueHeader("contractRange", "Contract")}
          ${revenueHeader("monthlyRevenue", "Monthly Revenue", true)}
          ${revenueHeader("contractAmount", "Contract Amount", true)}
          ${revenueHeader("revenueRule", "Rule")}
          ${revenueHeader("flags", "Checks")}
        </tr>
      </thead>
      <tbody>
        ${
          rows.length
            ? rows
                .map(
                  (row) => `
                    <tr class="clickable-row" data-deal-key="${escapeHtml(row.dealKey)}">
                      <td><strong>${escapeHtml(row.monthLabel)}</strong><br><span class="muted">${escapeHtml(displayDateDmy(row.revenueDate))} | ${escapeHtml(row.monthPosition)}</span></td>
                      <td><strong>${escapeHtml(row.saleName)}</strong><br><span class="muted">${escapeHtml(row.group || "-")}</span></td>
                      <td class="company-cell"><strong>${escapeHtml(row.company || "-")}</strong><br><span class="muted">${escapeHtml(row.dealName || "-")}</span></td>
                      <td>${escapeHtml(row.category)}</td>
                      <td>${escapeHtml(row.billingType)}</td>
                      <td>${escapeHtml(displayDateRangeDmy(row.contractRange))}</td>
                      <td class="num"><strong>${money(row.monthlyRevenue)}</strong></td>
                      <td class="num">${money(row.contractAmount)}</td>
                      <td>${escapeHtml(row.revenueRule)}</td>
                      <td>${revenueCheckBadge(row.flagsLabel)}</td>
                    </tr>
                  `,
                )
                .join("")
            : `<tr><td colspan="10"><div class="empty">ไม่มี Revenue detail ในช่วงเวลาที่เลือก</div></td></tr>`
        }
      </tbody>
    </table>
  `;
}

function revenueHeader(key, label, numeric = false, extraClass = "") {
  const className = [numeric ? "num" : "", extraClass].filter(Boolean).join(" ");
  return `<th class="${className}"><div class="table-header-tools"><span>${escapeHtml(label)}</span>${columnFilterButton("revenue", key)}</div></th>`;
}

function revenueCheckBadge(label) {
  if (!label || label === "Ready") return `<span class="badge good">Ready</span>`;
  return `<span class="badge warn">${escapeHtml(label)}</span>`;
}

function renderRevenueExceptionTable(exceptionRows) {
  els.revenueExceptionTable.innerHTML = exceptionRows.length
    ? `
      <table>
        <thead>
          <tr>
            <th>ID</th>
            <th>Sale</th>
            <th class="company-cell">Company / Deal</th>
            <th>Type</th>
            <th>Billing</th>
            <th>Contract</th>
            <th class="num">Amount</th>
            <th>Checks</th>
            <th>Suggested Action</th>
          </tr>
        </thead>
        <tbody>
          ${exceptionRows
            .map(
              (row) => `
                <tr class="clickable-row" data-deal-key="${escapeHtml(row.dealKey)}">
                  <td>${escapeHtml(row.id)}</td>
                  <td>${escapeHtml(row.saleName)}</td>
                  <td class="company-cell">${escapeHtml(row.companyDeal)}</td>
                  <td>${escapeHtml(row.category)}</td>
                  <td>${escapeHtml(row.billingType)}</td>
                  <td>${escapeHtml(row.contractRange)}</td>
                  <td class="num">${money(row.contractAmount)}</td>
                  <td>${revenueCheckBadge(row.flagsLabel)}</td>
                  <td>เติม/ตรวจ Contract Start, End, Period หรือ MRR แล้ว Refresh Dashboard</td>
                </tr>
              `,
            )
            .join("")}
        </tbody>
      </table>
    `
    : `<div class="empty">ไม่พบรายการที่ต้องตรวจสอบเพิ่มเติม</div>`;
}

function leadCreatedWeek(deal) {
  return isoWeekKeyFromDate(deal.createdDate);
}

function leadCreatedTs(deal) {
  return parseDateValue(deal.createdDate)?.ts || 0;
}

function sortLeadByCreatedDate(a, b) {
  return (
    leadCreatedTs(a) - leadCreatedTs(b) ||
    String(a.id || "").localeCompare(String(b.id || ""), "th", { numeric: true }) ||
    a.saleName.localeCompare(b.saleName, "th")
  );
}

function filteredLeadDeals() {
  const scope = calcScope();
  const searchTerms = [state.leadSearch].map(normalizeSearch).filter(Boolean);
  return filteredDeals(scope, { category: state.leadCategory, period: false, globalSearch: false, globalCategory: false })
    .map((deal) => ({ ...deal, createdWeek: leadCreatedWeek(deal) }))
    .filter((deal) => deal.createdWeek && deal.createdWeek >= state.leadStartWeek && deal.createdWeek <= state.leadEndWeek)
    .filter((deal) => {
      if (!searchTerms.length) return true;
      const text = dealSearchText(deal);
      return searchTerms.every((term) => text.includes(term));
    })
    .sort(sortLeadByCreatedDate);
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
    ["Total Leads", deals.length.toLocaleString("th-TH"), `ตั้งแต่ ${weekLabel(state.leadStartWeek)}\nถึง ${weekLabel(state.leadEndWeek)}`],
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
  tableFilterSources.leads = deals;
  const tableDeals = applyColumnFilters("leads", deals);
  let rowsHtml = `<tr><td colspan="8"><div class="empty">ไม่มี Lead ในช่วง Week ที่เลือก</div></td></tr>`;
  if (tableDeals.length) {
    const grouped = new Map();
    tableDeals.forEach((deal) => {
      if (!grouped.has(deal.createdWeek)) grouped.set(deal.createdWeek, []);
      grouped.get(deal.createdWeek).push(deal);
    });
    rowsHtml = Array.from(grouped.entries())
      .map(([week, weekDeals]) => {
        const totalAmount = sum(weekDeals, (deal) => deal.amount);
        const groupHeader = `
          <tr class="week-group-row">
            <td colspan="8"><strong>${escapeHtml(weekLabel(week))}</strong><span>${weekDeals.length.toLocaleString("th-TH")} leads · ${compactMoney(totalAmount)}</span></td>
          </tr>
        `;
        const detailRows = weekDeals
          .sort(sortLeadByCreatedDate)
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
  }
  els.newLeadTable.innerHTML = `
    <table>
      <thead>
        <tr>
          ${leadHeader("createdDate", "Created")}
          ${leadHeader("saleName", "Sale")}
          ${leadHeader("companyDeal", "Company / Deal", false, "company-cell")}
          ${leadHeader("category", "Type")}
          ${leadHeader("stage", "Matched Stage")}
          ${leadHeader("pipeline", "Pipeline")}
          ${leadHeader("amount", "Amount", true)}
          ${leadHeader("risk", "Risk / Progress")}
        </tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  `;
}

function leadHeader(key, label, numeric = false, extraClass = "") {
  const className = [numeric ? "num" : "", extraClass].filter(Boolean).join(" ");
  return `<th class="${className}"><div class="table-header-tools"><span>${escapeHtml(label)}</span>${columnFilterButton("leads", key)}</div></th>`;
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
        ${detailLine("Billing Type", deal.billingType || "-")}
        ${detailLine("MRR", deal.mrr ? money(deal.mrr) : "-")}
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
        ${detailLine("Contract Start", deal.contractStartDate || "-")}
        ${detailLine("Contract End", deal.contractEndDate || "-")}
        ${detailLine("Contract Period", deal.contractPeriodMonths ? `${deal.contractPeriodMonths.toLocaleString("th-TH")} months` : "-")}
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
  if (deal.status === "won" && !parseDateValue(deal.contractStartDate)) items.push("Revenue Analysis จะใช้วัน fallback หากไม่มี Contract Start Date ควรเติมข้อมูลก่อนส่งต่อบัญชี/การเงิน");
  if (deal.status === "won" && !parseDateValue(deal.contractEndDate) && !revenuePeriodMonths(deal)) items.push("ยังไม่มี Contract End Date หรือ Contract Period ทำให้ระบบกระจายรายได้แบบเดือนเดียวชั่วคราว");
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

function renderOverviewView(scope) {
  renderKpis(scope);
  renderMonthlyChart(scope);
  renderTeamChart(scope);
  renderCategoryMix(scope);
  renderStatusMix(scope);
  renderTopDealsTable(scope);
}

function renderSalesPerformanceView(scope) {
  renderKpis(scope);
  renderSalesPerformanceCharts(scope);
  renderTransactionTable(scope);
}

function renderPipelineRiskView(scope) {
  renderRiskSummary(scope);
  renderRiskByStage(scope);
  renderRiskDealsTable(scope);
}

function renderActiveTabView(scope) {
  if (state.tab === "overview") renderOverviewView(scope);
  if (state.tab === "sales") renderSalesPerformanceView(scope);
  if (state.tab === "revenue") renderRevenueAnalysis(scope);
  if (state.tab === "deals") renderDealDetailTable(scope);
  if (state.tab === "leads") renderNewLead();
  if (state.tab === "pipeline") renderPipelineRiskView(scope);
  if (state.tab === "data") renderDataQuality();
}

function renderAllDashboardViews(scope) {
  renderOverviewView(scope);
  renderDealDetailTable(scope);
  renderNewLead();
  renderRevenueAnalysis(scope);
  renderSalesPerformanceView(scope);
  renderPipelineRiskView(scope);
  renderDataQuality();
}

function renderFilteredView(options = {}) {
  const { layout = true, meta = false, allTabs = false } = options;
  if (layout) renderTabDescription();
  if (meta) renderMeta();
  syncSelectControls();
  const scope = calcScope();
  if (allTabs) {
    renderAllDashboardViews(scope);
    return;
  }
  renderActiveTabView(scope);
}

function renderTabDescription() {
  const isLeadsTab = state.tab === "leads";
  const hideKpiTabs = new Set(["revenue", "deals", "pipeline", "data", "leads"]);
  document.body.classList.toggle("is-leads-tab", isLeadsTab);
  renderCategoryButtons();
  if (els.dashboardFilterShell) els.dashboardFilterShell.hidden = false;
  if (els.periodFilterCard) els.periodFilterCard.hidden = isLeadsTab;
  if (els.leadToolbarCard) els.leadToolbarCard.hidden = !isLeadsTab;
  if (els.emptyDataBanner) els.emptyDataBanner.hidden = isLeadsTab || hasDealData();
  if (els.kpiGrid) els.kpiGrid.hidden = hideKpiTabs.has(state.tab);
  if (!els.tabDescription) return;
  els.tabDescription.hidden = false;
  els.tabDescription.textContent = tabDescriptions[state.tab] || "";
}

function renderCategoryButtons() {
  const activeCategory = state.tab === "leads" ? state.leadCategory : state.category;
  document.querySelectorAll(".segment button[data-category]").forEach((button) => {
    button.classList.toggle("active", button.dataset.category === activeCategory);
  });
}

function render() {
  renderFilteredView({ meta: true, allTabs: true });
}

initFilters();
render();
autoLoadDefaultSetupFiles();
