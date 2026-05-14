const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const DEAL_PATH =
  process.argv[2] || "C:\\Temp\\DEAL_20260506_2e04b49b_69fabfc019c48.csv";
const TARGET_PATH =
  process.argv[3] ||
  "C:\\Users\\noppadol.s\\OneDrive - 1-TO-ALL Co., Ltd\\Project\\1toAll\\Sales Dashboard\\Sales Amount\\Sale Target 2026_r1.5.csv";
const MAPPING_PATH =
  process.argv[4] ||
  (fs.existsSync("C:\\CodexWS-Antigravity\\Data Mapping.csv")
    ? "C:\\CodexWS-Antigravity\\Data Mapping.csv"
    : null);
const OUT_PATH = path.join(ROOT, "data", "dashboard-data.js");

const MONTHS = Array.from({ length: 5 * 12 }, (_, index) => {
  const year = 2026 + Math.floor(index / 12);
  const month = (index % 12) + 1;
  return `${year}-${String(month).padStart(2, "0")}`;
});
const TODAY = dateParts("2026-05-11");
const DAY_MS = 24 * 60 * 60 * 1000;

const STAGE_WEIGHTS = {
  Commit: 0.8,
  "Negotiations Started": 0.6,
  "Quotation Sent": 0.55,
  Upside: 0.5,
  "Qualified Pipeline": 0.4,
  "Deal Proposed": 0.35,
  "Deal (By Chance Project)": 0.25,
  "Pre-Qualified Pipeline": 0.2,
  Contacted: 0.15,
  "Contacted-OK": 0.15,
  "New Request": 0.08,
  "Not Contacted": 0.05,
  Backlog: 0.1,
  Inprogress: 0.1,
  Delivered: 0.7,
  Completed: 0.7,
  Deal: 0.25,
};

const RESPONSIBLE_ALIASES = {
  "chananthicha wongsan": "Chananthicha",
  chananthicha: "Chananthicha",
  "phongthorn meeshaeng": "Phongthorn",
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

function readCsv(filePath) {
  const text = fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, "");
  return parseCsv(text);
}

function parseCsv(text) {
  const rows = parseCsvRows(text);
  if (!rows.length) return [];
  const headers = rows[0].map((header) => header.trim());
  return rows.slice(1).map((values) => {
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

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];

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

function parseMoney(value) {
  const text = String(value ?? "").trim();
  if (!text) return 0;
  const compact = text.replace(/\s/g, "");
  if (!compact || /^-+$/.test(compact)) return 0;
  const cleaned = compact.replace(/,/g, "").replace(/[฿]/g, "").replace(/[^\d.-]/g, "");
  if (!cleaned || cleaned === "-" || cleaned === "." || cleaned === "-.") return 0;
  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
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

function canonicalSaleName(value) {
  return cleanText(value).replace(/\([^)]*\)/g, "").replace(/\s+/g, " ").trim();
}

function makeKey(value) {
  return normalizeName(value).replace(/\s+/g, "-") || "unknown";
}

function parseMappingCsv(filePath) {
  const fallback = {
    pipelineGroupMap: new Map(),
    pipelineRuleMap: new Map(),
    stageMap: new Map(),
    dealTypeMap: new Map(),
    salesGroupMap: new Map(),
    counts: { pipelines: 0, pipelineRules: 0, stages: 0, dealTypes: 0, sales: 0 },
  };
  if (!filePath || !fs.existsSync(filePath)) return fallback;

  const rows = parseCsvRows(fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, "")).map((row) => row.map(cleanText));
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

function serializeMappings(mappings) {
  return {
    pipelineGroupMap: Object.fromEntries(mappings.pipelineGroupMap),
    pipelineRuleMap: Object.fromEntries(Array.from(mappings.pipelineRuleMap || []).map(([key, values]) => [key, Array.from(values || [])])),
    stageMap: Object.fromEntries(mappings.stageMap),
    dealTypeMap: Object.fromEntries(mappings.dealTypeMap),
    salesGroupMap: Object.fromEntries(mappings.salesGroupMap),
    counts: mappings.counts,
  };
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

function pipelineCountingResult(group, category, pipeline, mappings) {
  const rules = mappings.pipelineRuleMap || new Map();
  if (!rules.size) return { status: "included", included: true, label: "Included" };
  const key = pipelineRuleKey(group, category);
  const allowedPipelines = rules.get(key);
  if (!allowedPipelines) return { status: "unmapped", included: false, label: "Unmapped rule" };
  if (allowedPipelines.has(normalizeName(pipeline))) return { status: "included", included: true, label: "Included" };
  return { status: "excluded", included: false, label: "Excluded pipeline" };
}

function parseDate(value) {
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
  return dateParts(`${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`);
}

function dateParts(isoDate) {
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

function normalizeStage(stage) {
  const raw = cleanText(stage);
  const lower = raw.toLowerCase();
  if (lower === "won" || lower === "deal won" || lower === "pre-won") return "won";
  if (lower === "lost" || lower === "deal lost" || lower === "pre-lost") return "lost";
  return "open";
}

function inferCategory(row) {
  const pipeline = cleanText(row.Pipeline).toLowerCase();
  const dealType = cleanText(row["Deal Type"]).toLowerCase();
  if (pipeline.includes("renew") || dealType.startsWith("re-new")) return "renew";
  return "new";
}

function stageWeight(stage) {
  return STAGE_WEIGHTS[cleanText(stage)] ?? 0.1;
}

function mappedStage(stage, mappings) {
  const raw = cleanText(stage);
  return mappings.stageMap.get(normalizeName(raw)) || raw || "Open";
}

function mappedDealType(dealType, mappings) {
  const raw = cleanText(dealType);
  return mappings.dealTypeMap.get(normalizeName(raw)) || raw || "New";
}

function dealCategory(dealType, row) {
  const mapped = cleanText(dealType).toLowerCase();
  if (mapped.includes("renew")) return "renew";
  if (mapped.includes("new")) return "new";
  return inferCategory(row);
}

function stageWeightFor(rawStage, displayStage) {
  return STAGE_WEIGHTS[displayStage] ?? STAGE_WEIGHTS[rawStage] ?? 0.1;
}

function addAgg(map, keyParts, amount, extra = {}) {
  const key = keyParts.join("|");
  if (!map.has(key)) {
    map.set(key, { ...extra, amount: 0, count: 0 });
  }
  const item = map.get(key);
  item.amount += amount;
  item.count += 1;
}

function roundMoney(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

const mappings = parseMappingCsv(MAPPING_PATH);
const targetRowsRaw = readCsv(TARGET_PATH).filter((row) => cleanText(row.sales));
const dealRows = readCsv(DEAL_PATH);

const targetSales = targetRowsRaw.map((row) => {
  const saleName = canonicalSaleName(row.sales);
  const mappedSale = mappings.salesGroupMap.get(normalizeName(saleName));
  const key = makeKey(saleName);
  const monthly = {};
  MONTHS.forEach((month) => {
    monthly[month] = {
      renew: parseMoney(row[`renew ${month}`]),
      new: parseMoney(row[`new ${month}`]),
    };
    monthly[month].total = monthly[month].renew + monthly[month].new;
  });
  for (let year = 2026; year <= 2028; year += 1) {
    const yearMonths = MONTHS.filter((month) => month.startsWith(`${year}-`));
    const renewMonthly = yearMonths.reduce((sum, month) => sum + monthly[month].renew, 0);
    const newMonthly = yearMonths.reduce((sum, month) => sum + monthly[month].new, 0);
    const annualRenew = parseMoney(row[`renew ${year}`]);
    const annualNew = parseMoney(row[`new ${year}`]);
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
  const targetRenew = MONTHS.reduce((sum, month) => sum + monthly[month].renew, 0);
  const targetNew = MONTHS.reduce((sum, month) => sum + monthly[month].new, 0);
  const targetAnnual = {
    renew: roundMoney(targetRenew),
    new: roundMoney(targetNew),
    total: roundMoney(targetRenew + targetNew),
  };
  return {
    key,
    name: mappedSale?.displayName || saleName,
    group: mappedSale?.group || cleanText(row["sale group"]) || "No group",
    hasTarget: true,
    hasPositiveTarget: targetAnnual.total > 0,
    responsibleNames: [],
    monthly,
    targetAnnual,
  };
});

const saleByKey = new Map(targetSales.map((sale) => [sale.key, sale]));
const targetNameToKey = new Map(targetSales.map((sale) => [normalizeName(sale.name), sale.key]));

function mapResponsible(responsible) {
  const normalized = normalizeName(responsible);
  if (!normalized) return null;
  const mappedSale = mappings.salesGroupMap.get(normalized);
  if (mappedSale) return targetNameToKey.get(normalizeName(mappedSale.displayName)) || makeKey(mappedSale.displayName);
  const alias = RESPONSIBLE_ALIASES[normalized];
  if (alias) return targetNameToKey.get(normalizeName(alias)) || makeKey(alias);

  for (const sale of targetSales) {
    const targetNormalized = normalizeName(sale.name);
    if (normalized === targetNormalized || normalized.startsWith(`${targetNormalized} `)) {
      return sale.key;
    }
  }

  return `unmapped:${makeKey(responsible)}`;
}

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

for (const row of dealRows) {
  const responsible = cleanText(row.Responsible) || "Unknown";
  const saleKey = mapResponsible(responsible);
  const amount = parseMoney(row.Income);
  const rawStage = cleanText(row.Stage);
  const displayStage = mappedStage(rawStage, mappings);
  const status = normalizeStage(displayStage);
  const dealType = mappedDealType(row["Deal Type"], mappings);
  const category = dealCategory(dealType, row);
  const expected = parseDate(row["Expected close date"]);
  const stageChanged = parseDate(row["Stage change date"]);
  const created = parseDate(row.Created);
  const closeDate = stageChanged || expected || created;
  const pipeline = cleanText(row.Pipeline);
  const pipelineGroup = mappings.pipelineGroupMap.get(normalizeName(pipeline)) || "";

  if (!saleByKey.has(saleKey)) {
    const mappedSale = mappings.salesGroupMap.get(normalizeName(responsible));
    const displayName = mappedSale?.displayName || responsible;
    saleByKey.set(saleKey, {
      key: saleKey,
      name: displayName,
      group: mappedSale?.group || "Unmapped",
      hasTarget: false,
      hasPositiveTarget: false,
      responsibleNames: [],
      monthly: Object.fromEntries(MONTHS.map((month) => [month, { renew: 0, new: 0, total: 0 }])),
      targetAnnual: { renew: 0, new: 0, total: 0 },
    });
  }
  const sale = saleByKey.get(saleKey);
  if (!sale.responsibleNames.includes(responsible)) sale.responsibleNames.push(responsible);
  const counting = pipelineCountingResult(sale.group, category, pipeline, mappings);

  if (sale.group === "Unmapped") {
    const key = responsible;
    if (!unmappedResponsibles.has(key)) unmappedResponsibles.set(key, { responsible, count: 0, amount: 0 });
    const item = unmappedResponsibles.get(key);
    item.count += 1;
    item.amount += amount;
  }

  if (rawStage === "Pre-WON") {
    quality.preWonAsWon.count += 1;
    quality.preWonAsWon.amount += amount;
  }
  if (rawStage === "Pre-LOST") {
    quality.preLostAsLost.count += 1;
    quality.preLostAsLost.amount += amount;
  }

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
    stageAgeDays: stageChanged ? Math.floor((TODAY.ts - stageChanged.ts) / DAY_MS) : null,
    risk: {
      overdue: Boolean(expected && expected.ts < TODAY.ts),
      noExpected: !expected,
      stale30: Boolean(stageChanged && stageChanged.ts < TODAY.ts - 30 * DAY_MS),
      notContacted: displayStage === "Not Contacted" || rawStage === "Not Contacted",
    },
  };
  dealDetails.push(detail);
  if (pipelineGroup && pipelineGroup !== sale.group) {
    quality.pipelineGroupMismatch.count += 1;
    quality.pipelineGroupMismatch.amount += amount;
  }
  const countingBucket =
    counting.status === "included"
      ? quality.pipelineIncluded
      : counting.status === "excluded"
        ? quality.pipelineExcluded
        : quality.pipelineUnmapped;
  countingBucket.count += 1;
  countingBucket.amount += amount;

  if (!counting.included) continue;

  if (status === "won" || status === "lost") {
    const monthKey = closeDate?.monthKey || null;
    const qualityBucket = status === "won" ? quality.won : quality.lost;
    qualityBucket.count += 1;
    qualityBucket.amount += amount;
    if (MONTHS.includes(monthKey)) {
      addAgg(
        factsMap,
        [saleKey, monthKey, category, status],
        amount,
        {
          saleKey,
          monthKey,
          category,
          status,
          group: sale.group,
          saleName: sale.name,
        },
      );
    }
    continue;
  }

  quality.open.count += 1;
  quality.open.amount += amount;

  const noExpected = !expected;
  const overdue = Boolean(expected && expected.ts < TODAY.ts);
  const stale30 = Boolean(stageChanged && stageChanged.ts < TODAY.ts - 30 * DAY_MS);
  const notContacted = displayStage === "Not Contacted" || rawStage === "Not Contacted";
  const weight = stageWeightFor(rawStage, displayStage);
  const stageAgeDays = stageChanged ? Math.floor((TODAY.ts - stageChanged.ts) / DAY_MS) : null;

  if (noExpected) {
    quality.noExpectedOpen.count += 1;
    quality.noExpectedOpen.amount += amount;
  }
  if (overdue) {
    quality.overdueOpen.count += 1;
    quality.overdueOpen.amount += amount;
  }
  if (stale30) {
    quality.stale30Open.count += 1;
    quality.stale30Open.amount += amount;
  }
  if (notContacted) {
    quality.notContactedOpen.count += 1;
    quality.notContactedOpen.amount += amount;
  }

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
    stageAgeDays,
    risk: {
      overdue,
      noExpected,
      stale30,
      notContacted,
    },
  });
}

const facts = Array.from(factsMap.values()).map((fact) => ({
  ...fact,
  amount: roundMoney(fact.amount),
}));

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

for (const bucket of Object.values(quality)) {
  if (bucket && typeof bucket === "object" && "amount" in bucket) bucket.amount = roundMoney(bucket.amount);
}

const data = {
  metadata: {
    generatedAt: new Date().toISOString(),
    asOfDate: TODAY.date,
    sourceFiles: {
      deals: DEAL_PATH,
      target: TARGET_PATH,
      mapping: MAPPING_PATH,
    },
    assumptions: [
      "Pre-WON is normalized as Won and counted as achieved sales.",
      "Pre-LOST is normalized as Lost and excluded from open pipeline forecast.",
      "Actual month for Won/Lost uses Stage change date, falling back to Expected close date and Created date.",
      "Stage, Deal Type, Pipeline group, and Sales group are mapped from Data Mapping.csv when available.",
    ],
  },
  months: MONTHS,
  stageWeights: STAGE_WEIGHTS,
  mappings: serializeMappings(mappings),
  sales,
  facts,
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

const js = `window.DASHBOARD_DATA = ${JSON.stringify(data, null, 2)};\n`;
fs.mkdirSync(path.dirname(OUT_PATH), { recursive: true });
fs.writeFileSync(OUT_PATH, js, "utf8");

console.log(`Built ${OUT_PATH}`);
console.log(`Deals: ${dealRows.length.toLocaleString()} | Target rows: ${targetSales.length.toLocaleString()}`);
console.log(`Facts: ${facts.length.toLocaleString()} | Open deals: ${openDeals.length.toLocaleString()}`);
