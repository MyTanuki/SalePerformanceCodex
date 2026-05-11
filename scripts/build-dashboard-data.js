const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const DEAL_PATH =
  process.argv[2] || "C:\\Temp\\DEAL_20260506_2e04b49b_69fabfc019c48.csv";
const TARGET_PATH =
  process.argv[3] ||
  "C:\\Users\\noppadol.s\\OneDrive - 1-TO-ALL Co., Ltd\\Project\\1toAll\\Sales Dashboard\\Sales Amount\\Sale Target 2026_r1.5.csv";
const OUT_PATH = path.join(ROOT, "data", "dashboard-data.js");

const MONTHS = Array.from({ length: 12 }, (_, i) => `2026-${String(i + 1).padStart(2, "0")}`);
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
  if (lower === "deal won" || lower === "pre-won" || lower === "pre-won") return "won";
  if (lower === "deal lost" || lower === "pre-lost" || lower === "pre-lost") return "lost";
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

const targetRowsRaw = readCsv(TARGET_PATH).filter((row) => cleanText(row.sales));
const dealRows = readCsv(DEAL_PATH);

const targetSales = targetRowsRaw.map((row) => {
  const saleName = canonicalSaleName(row.sales);
  const key = makeKey(saleName);
  const monthly = {};
  MONTHS.forEach((month) => {
    monthly[month] = {
      renew: parseMoney(row[`renew ${month}`]),
      new: parseMoney(row[`new ${month}`]),
    };
    monthly[month].total = monthly[month].renew + monthly[month].new;
  });
  const targetRenew = MONTHS.reduce((sum, month) => sum + monthly[month].renew, 0);
  const targetNew = MONTHS.reduce((sum, month) => sum + monthly[month].new, 0);
  const targetAnnual = {
    renew: roundMoney(targetRenew),
    new: roundMoney(targetNew),
    total: roundMoney(targetRenew + targetNew),
  };
  return {
    key,
    name: saleName,
    group: cleanText(row["sale group"]) || "No group",
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
};

for (const row of dealRows) {
  const responsible = cleanText(row.Responsible) || "Unknown";
  const saleKey = mapResponsible(responsible);
  const amount = parseMoney(row.Income);
  const rawStage = cleanText(row.Stage);
  const status = normalizeStage(rawStage);
  const category = inferCategory(row);
  const expected = parseDate(row["Expected close date"]);
  const stageChanged = parseDate(row["Stage change date"]);
  const created = parseDate(row.Created);
  const closeDate = stageChanged || expected || created;

  if (!saleByKey.has(saleKey)) {
    const displayName = responsible;
    saleByKey.set(saleKey, {
      key: saleKey,
      name: displayName,
      group: "Unmapped",
      hasTarget: false,
      hasPositiveTarget: false,
      responsibleNames: [],
      monthly: Object.fromEntries(MONTHS.map((month) => [month, { renew: 0, new: 0, total: 0 }])),
      targetAnnual: { renew: 0, new: 0, total: 0 },
    });
  }
  const sale = saleByKey.get(saleKey);
  if (!sale.responsibleNames.includes(responsible)) sale.responsibleNames.push(responsible);

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
    stage: rawStage,
    pipeline: cleanText(row.Pipeline),
    product: cleanText(row["Product Type"]) || "(blank)",
    company: cleanText(row.Company),
    contact: cleanText(row.Contact),
    dealName: cleanText(row["Deal Name"]),
    amount: roundMoney(amount),
    forecastAmount: roundMoney(status === "open" ? amount * stageWeight(rawStage) : 0),
    expectedDate: expected?.date || "",
    expectedMonth: expected?.monthKey || "",
    stageChangeDate: stageChanged?.date || "",
    stageMonth: stageChanged?.monthKey || "",
    createdDate: created?.date || "",
    trackingMonth: status === "open" ? expected?.monthKey || "" : closeDate?.monthKey || "",
    stageAgeDays: stageChanged ? Math.floor((TODAY.ts - stageChanged.ts) / DAY_MS) : null,
    risk: {
      overdue: Boolean(expected && expected.ts < TODAY.ts),
      noExpected: !expected,
      stale30: Boolean(stageChanged && stageChanged.ts < TODAY.ts - 30 * DAY_MS),
      notContacted: rawStage === "Not Contacted",
    },
  };
  dealDetails.push(detail);

  if (status === "won" || status === "lost") {
    const monthKey = closeDate?.monthKey || null;
    const qualityBucket = status === "won" ? quality.won : quality.lost;
    qualityBucket.count += 1;
    qualityBucket.amount += amount;
    if (monthKey?.startsWith("2026-")) {
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
  const notContacted = rawStage === "Not Contacted";
  const weight = stageWeight(rawStage);
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
    },
    assumptions: [
      "Pre-WON is normalized as Won and counted as achieved sales.",
      "Pre-LOST is normalized as Lost and excluded from open pipeline forecast.",
      "Actual month for Won/Lost uses Stage change date, falling back to Expected close date and Created date.",
      "Renew category uses Pipeline containing Renew or Deal Type starting with Re-New; everything else is New.",
    ],
  },
  months: MONTHS,
  stageWeights: STAGE_WEIGHTS,
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
