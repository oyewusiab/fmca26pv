/***************************************************************
 * ACTION ITEM SYSTEM – COMPLETE ENGINE
 ***************************************************************/

const ACTION_ITEM_RULES = {
  PAID_NO_CN: "PAID_NO_CN",
  UNPAID_NO_CN_30D: "UNPAID_NO_CN_30D",
  RELEASED_UNPAID_15D: "RELEASED_UNPAID_15D",
  MISSING_DATA: "MISSING_DATA",
  DUPLICATE_VOUCHER: "DUPLICATE_VOUCHER"
};

const ACTION_ITEM_UNITS = {
  PAYABLE: "PAYABLE",
  CPO: "CPO",
  ADMIN: "ADMIN"
};

const ACTION_ITEM_2026_COLUMNS = {
  STATUS: "STATUS",
  PMT_MONTH: "PMT MONTH",
  PAYEE: "PAYEE",
  ACCOUNT_OR_MAIL: "ACCOUNT OR MAIL",
  GROSS_AMOUNT: "GROSS AMOUNT",
  CATEGORIES: "CATEGORIES",
  CONTROL_NUMBER: "CONTROL NUMBER",
  DATE: "DATE",
  ACCOUNT_TYPE: "ACCOUNT TYPE",
  SUB_ACCOUNT_TYPE: "SUB ACCOUNT TYPE",
  CREATED_AT: "CREATED AT",
  RELEASED_AT: "RELEASED AT"
};

/***************************************************************
 * SHEET VALIDATION (No Auto-Creation)
 ***************************************************************/
function ensureActionItemSheets_() {
  let logSheet = null;
  let settingsSheet = null;

  try { logSheet = getSheet(CONFIG.SHEETS.ACTION_ITEMS_LOG); } catch (_) {}
  try { settingsSheet = getSheet(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS); } catch (_) {}

  if (!logSheet) {
    throw new Error(
      'Missing required sheet: ' + CONFIG.SHEETS.ACTION_ITEMS_LOG +
      '. Auto-creation is disabled.'
    );
  }

  return { logSheet, settingsSheet };
}

function getDefaultActionItemSettings_() {
  return {
    unit: {
      PAYABLE: {},
      CPO: {},
      ADMIN: {}
    }
  };
}

function normalizeActionItemStatus_(status) {
  return String(status || "").trim().toUpperCase();
}

function parseActionItemDate_(value, locale) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof parseDateTimeFlexible_ === "function") {
    const parsed = parseDateTimeFlexible_(value, locale);
    if (parsed instanceof Date && !isNaN(parsed.getTime())) return parsed;
  }

  const s = String(value || "").trim();
  if (!s) return null;

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const a = parseInt(m[1], 10);
    const b = parseInt(m[2], 10);
    const yyyy = parseInt(m[3], 10);
    const hh = parseInt(m[4] || "0", 10);
    const mm = parseInt(m[5] || "0", 10);
    const ss = parseInt(m[6] || "0", 10);

    let day = a;
    let month = b;
    if (a <= 12 && b <= 12) {
      const isUS = String(locale || "").toLowerCase().includes("us");
      if (isUS) {
        month = a;
        day = b;
      }
    } else if (b > 12) {
      day = b;
      month = a;
    }

    const d = new Date(yyyy, month - 1, day, hh, mm, ss);
    if (!isNaN(d.getTime())) return d;
  }

  const fallback = new Date(s);
  return isNaN(fallback.getTime()) ? null : fallback;
}

function resolveCol_(header, keys, fallbackCol) {
  for (let i = 0; i < keys.length; i++) {
    const col = header[keys[i]];
    if (col) return col;
  }
  return fallbackCol || null;
}

function canonicalHeaderKey_(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function getCanonicalHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const key = canonicalHeaderKey_(h);
    if (key) map[key] = i + 1;
  });
  return map;
}

function resolveCanonicalCol_(canonicalMap, names, fallbackCol) {
  const existingKeys = Object.keys(canonicalMap || {});
  for (let i = 0; i < names.length; i++) {
    const key = canonicalHeaderKey_(names[i]);
    if (canonicalMap[key]) return canonicalMap[key];

    // Fuzzy match supports annotated headers like:
    // "RULE_KEY (PAID_NO_CN, ...)" and trailing symbols.
    for (let j = 0; j < existingKeys.length; j++) {
      const candidate = existingKeys[j];
      if (
        candidate === key ||
        candidate.indexOf(key + " ") === 0 ||
        candidate.indexOf(" " + key + " ") >= 0 ||
        candidate.indexOf(key) >= 0
      ) {
        return canonicalMap[candidate];
      }
    }
  }
  return fallbackCol || null;
}

function map2026ActionItemColumns_(canonicalMap) {
  const out = {
    status: resolveCanonicalCol_(canonicalMap, [ACTION_ITEM_2026_COLUMNS.STATUS, "STATUS"]),
    pmtMonth: resolveCanonicalCol_(canonicalMap, [ACTION_ITEM_2026_COLUMNS.PMT_MONTH, "PMT MONTH"]),
    payee: resolveCanonicalCol_(canonicalMap, [ACTION_ITEM_2026_COLUMNS.PAYEE, "PAYEE"]),
    voucherNo: resolveCanonicalCol_(canonicalMap, [
      ACTION_ITEM_2026_COLUMNS.ACCOUNT_OR_MAIL,
      "ACCOUNT OR MAIL (VOUCHER NUMBER)",
      "ACCOUNT OR MAIL"
    ]),
    grossAmount: resolveCanonicalCol_(canonicalMap, [ACTION_ITEM_2026_COLUMNS.GROSS_AMOUNT, "GROSS AMOUNT"]),
    category: resolveCanonicalCol_(canonicalMap, [ACTION_ITEM_2026_COLUMNS.CATEGORIES, "CATEGORIES"]),
    subAccountType: resolveCanonicalCol_(canonicalMap, [ACTION_ITEM_2026_COLUMNS.SUB_ACCOUNT_TYPE, "SUB ACCOUNT TYPE", "SUB-ACCOUNT TYPE"]),
    controlNumber: resolveCanonicalCol_(canonicalMap, [
      ACTION_ITEM_2026_COLUMNS.CONTROL_NUMBER,
      "CONTROL NO.",
      "CONTROL NO",
      "CONTROL NUMBER"
    ]),
    date: resolveCanonicalCol_(canonicalMap, [ACTION_ITEM_2026_COLUMNS.DATE, "DATE"]),
    accountType: resolveCanonicalCol_(canonicalMap, [ACTION_ITEM_2026_COLUMNS.ACCOUNT_TYPE, "ACCOUNT TYPE"]),
    createdAt: resolveCanonicalCol_(canonicalMap, [ACTION_ITEM_2026_COLUMNS.CREATED_AT, "CREATED AT"]),
    releasedAt: resolveCanonicalCol_(canonicalMap, [ACTION_ITEM_2026_COLUMNS.RELEASED_AT, "RELEASED AT"])
  };

  return out;
}

function getAccountTypeSubRequirementMap_() {
  const out = {};
  try {
    const sheet = getSheet(CONFIG.SHEETS.SYSTEM_CONFIG);
    if (!sheet) return out;
    const lastRow = sheet.getLastRow();
    if (lastRow < 7) return out;

    const values = sheet.getRange(7, 2, lastRow - 6, 2).getValues();
    values.forEach((row) => {
      const accountType = String(row[0] || "").trim();
      if (!accountType) return;
      const rawSubs = String(row[1] || "").trim();
      const subs = rawSubs
        ? rawSubs.split(",").map((x) => String(x || "").trim()).filter((x) => x)
        : [];
      out[accountType] = subs;
    });
  } catch (_) {}
  return out;
}

function getActionItemSheetNameByYear_(year) {
  const y = String(year || "2026");
  if (y === "2026") return CONFIG.SHEETS.VOUCHERS_2026;
  if (y === "2025") return CONFIG.SHEETS.VOUCHERS_2025;
  if (y === "2024") return CONFIG.SHEETS.VOUCHERS_2024;
  if (y === "2023") return CONFIG.SHEETS.VOUCHERS_2023;
  if (y === "<2023") return CONFIG.SHEETS.VOUCHERS_BEFORE_2023;
  return CONFIG.SHEETS.VOUCHERS_2026;
}

function getActionItemCurrentVoucherMeta_(year, rowIndex, locale) {
  const sheetName = getActionItemSheetNameByYear_(year);
  if (!sheetName) return null;

  let sheet = null;
  try {
    sheet = getSheet(sheetName);
  } catch (_) {
    return null;
  }

  const idx = parseInt(rowIndex, 10);
  if (isNaN(idx) || idx < 2 || idx > sheet.getLastRow()) return null;

  const canonicalHeader = getCanonicalHeaderMap_(sheet);
  const cols = map2026ActionItemColumns_(canonicalHeader);
  if (!cols.status || !cols.voucherNo || !cols.payee || !cols.grossAmount) return null;

  const readWidth = Math.min(
    sheet.getMaxColumns(),
    Math.max(
      sheet.getLastColumn(),
      cols.status || 1,
      cols.pmtMonth || 1,
      cols.payee || 1,
      cols.voucherNo || 1,
      cols.grossAmount || 1,
      cols.category || 1,
      cols.subAccountType || 1,
      cols.controlNumber || 1,
      cols.date || 1,
      cols.accountType || 1,
      cols.createdAt || 1,
      cols.releasedAt || 1
    )
  );

  const row = sheet.getRange(idx, 1, 1, readWidth).getValues()[0];
  const status = normalizeActionItemStatus_(row[(cols.status || 1) - 1]);
  const controlNumber = String(row[(cols.controlNumber || 1) - 1] || "").trim();
  const createdAt = parseActionItemDate_(row[(cols.createdAt || 1) - 1], locale)
    || parseActionItemDate_(row[(cols.date || 1) - 1], locale);
  let releasedAt = parseActionItemDate_(row[(cols.releasedAt || 1) - 1], locale);
  if (!releasedAt && controlNumber) releasedAt = createdAt;

  const amountRaw = row[(cols.grossAmount || 1) - 1] || 0;
  const amount = typeof parseAmount === "function" ? parseAmount(amountRaw) : Number(amountRaw || 0);
  const baseAccountType = String(row[(cols.accountType || 1) - 1] || "").trim();
  const subAccountType = String(row[(cols.subAccountType || 1) - 1] || "").trim();

  return {
    status,
    pmtMonth: String(row[(cols.pmtMonth || 1) - 1] || "").trim(),
    voucherNumber: row[(cols.voucherNo || 1) - 1],
    payee: row[(cols.payee || 1) - 1],
    amount,
    controlNumber,
    category: String(row[(cols.category || 1) - 1] || "").trim(),
    accountType: subAccountType ? `${baseAccountType} (${subAccountType})` : baseAccountType,
    subAccountType,
    createdAt,
    releasedAt
  };
}

/***************************************************************
 * ROLE HELPERS
 ***************************************************************/
function roleToUnit_(role) {
  if (role === CONFIG.ROLES.CPO) return ACTION_ITEM_UNITS.CPO;
  if ([CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.PAYABLE_STAFF].includes(role))
    return ACTION_ITEM_UNITS.PAYABLE;
  return null;
}

function canViewAllUnits_(role) {
  return [CONFIG.ROLES.ADMIN, CONFIG.ROLES.DDFA, CONFIG.ROLES.DFA]
    .includes(role);
}

/***************************************************************
 * SETTINGS LOADER
 ***************************************************************/
function loadActionItemSettings_() {
  const defaults = getDefaultActionItemSettings_();
  let sheet = null;

  try {
    sheet = getSheet(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS);
  } catch (_) {
    return defaults;
  }

  if (sheet.getLastRow() < 2) return defaults;
  const data = sheet.getDataRange().getValues();

  const settings = defaults;

  for (let i = 1; i < data.length; i++) {
    const unit = String(data[i][0] || "").toUpperCase();
    const rule = String(data[i][1] || "").trim();
    const enabled = String(data[i][2]).toUpperCase() !== "FALSE";

    if (settings.unit[unit])
      settings.unit[unit][rule] = enabled;
  }

  return settings;
}

function isRuleEnabled_(settings, unit, ruleKey) {
  return settings?.unit?.[unit]?.[ruleKey] !== false;
}

/***************************************************************
 * CLEAN ACTION TEXT (Refined Statements)
 ***************************************************************/
function buildActionItemText_(ruleKey, unit, meta) {

  if (ruleKey === ACTION_ITEM_RULES.PAID_NO_CN) {
    if (unit === ACTION_ITEM_UNITS.PAYABLE) {
      return {
        title: "Paid Voucher Not Released to CPO",
        message: "Voucher is marked PAID but has not been released to CPO. Please release it to CPO Unit.",
        severity: "warning"
      };
    }
    if (unit === ACTION_ITEM_UNITS.ADMIN) {
      return {
        title: "Paid Voucher Missing Control Number",
        message: "Voucher is marked PAID but has no control number. Review release routing and correct the record.",
        severity: "warning"
      };
    }
    return {
      title: "Paid Voucher Awaiting Release from Payable",
      message: "Voucher is marked PAID but has not reached CPO. Please request release from Payable Unit.",
      severity: "info"
    };
  }

  if (ruleKey === ACTION_ITEM_RULES.UNPAID_NO_CN_30D) {
    return {
      title: "Voucher Delayed in Payable Unit",
      message: "Voucher is over 30 days old, UNPAID, and has no control number. Request it from Payable Unit for payment processing.",
      severity: "warning"
    };
  }

  if (ruleKey === ACTION_ITEM_RULES.RELEASED_UNPAID_15D) {
    return {
      title: "Released Voucher Still Marked Unpaid",
      message: "Voucher has been released for over 15 days but status is still UNPAID. Update the voucher payment status.",
      severity: "danger"
    };
  }

  if (ruleKey === ACTION_ITEM_RULES.MISSING_DATA) {
    return {
      title: "Voucher Missing Required Fields",
      message: "Voucher record is missing required classification fields (Category, Account Type, or required Sub Account Type).",
      severity: "warning"
    };
  }

  if (ruleKey === ACTION_ITEM_RULES.DUPLICATE_VOUCHER) {
    return {
      title: "Duplicate Voucher Number Detected",
      message: "Voucher number appears more than once in 2026 VOUCHERS. Verify and resolve duplicates.",
      severity: "danger"
    };
  }

  return { title: "Action Required", message: "", severity: "info" };
}

/***************************************************************
 * MAIN SYNC ENGINE (AUTO RESOLVE + UPSERT)
 ***************************************************************/
function syncActionItems_() {
  const now = new Date();
  const settings = loadActionItemSettings_();
  const locale = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSpreadsheetLocale();
  const sheets = ensureActionItemSheets_();
  const logSheet = sheets.logSheet;

  const sources = [
    { year: "2026", sheetName: CONFIG.SHEETS.VOUCHERS_2026 }
  ];

  const currentItems = [];
  const activeIds = new Set();
  const accountTypeRequirements = getAccountTypeSubRequirementMap_();
  const duplicateTracker = {};

  sources.forEach(src => {
    if (!src.sheetName) return;

    let voucherSheet = null;
    try {
      voucherSheet = getSheet(src.sheetName);
    } catch (_) {
      return;
    }

    const lastRow = voucherSheet.getLastRow();
    if (lastRow < 2) return;

    const canonicalHeader = getCanonicalHeaderMap_(voucherSheet);
    const cols = map2026ActionItemColumns_(canonicalHeader);

    if (!cols.status || !cols.voucherNo || !cols.payee || !cols.grossAmount || !cols.controlNumber) {
      throw new Error(
        '2026 VOUCHERS header does not match required Action Item columns. ' +
        'Expected at least: STATUS, PAYEE, ACCOUNT OR MAIL, GROSS AMOUNT, CONTROL NUMBER.'
      );
    }

    const readWidth = Math.min(
      voucherSheet.getMaxColumns(),
      Math.max(
        voucherSheet.getLastColumn(),
        cols.status || 1,
        cols.pmtMonth || 1,
        cols.payee || 1,
        cols.voucherNo || 1,
        cols.grossAmount || 1,
        cols.category || 1,
        cols.subAccountType || 1,
        cols.controlNumber || 1,
        cols.date || 1,
        cols.accountType || 1,
        cols.createdAt || 1,
        cols.releasedAt || 1
      )
    );
    const data = voucherSheet.getRange(2, 1, lastRow - 1, readWidth).getValues();

    data.forEach((row, i) => {
      const rowIndex = i + 2;
      const status = normalizeActionItemStatus_(row[(cols.status || 1) - 1]);
      const pmtMonth = String(row[(cols.pmtMonth || 1) - 1] || "").trim();
      const cn = String(row[(cols.controlNumber || 1) - 1] || "").trim();
      const hasControlNumber = cn !== "";

      const voucherNo = row[(cols.voucherNo || 1) - 1];
      const voucherNoClean = String(voucherNo || "").trim();
      const voucherNoKey = voucherNoClean.toUpperCase();
      const payee = row[(cols.payee || 1) - 1];
      const category = String(row[(cols.category || 1) - 1] || "").trim();
      const accountType = String(row[(cols.accountType || 1) - 1] || "").trim();
      const subAccountType = String(row[(cols.subAccountType || 1) - 1] || "").trim();
      const accountTypeLabel = subAccountType ? `${accountType} (${subAccountType})` : accountType;
      const amountRaw = row[(cols.grossAmount || 1) - 1] || 0;
      const amount = typeof parseAmount === "function" ? parseAmount(amountRaw) : Number(amountRaw || 0);

      const createdAt = parseActionItemDate_(row[(cols.createdAt || 1) - 1], locale)
        || parseActionItemDate_(row[(cols.date || 1) - 1], locale);
      let releasedAt = parseActionItemDate_(row[(cols.releasedAt || 1) - 1], locale);
      if (!releasedAt && hasControlNumber) releasedAt = createdAt;

      if (!String(voucherNo || "").trim() && !String(payee || "").trim() && amount <= 0) {
        return;
      }

      if (voucherNoKey) {
        if (!duplicateTracker[voucherNoKey]) duplicateTracker[voucherNoKey] = [];
        duplicateTracker[voucherNoKey].push({
          rowIndex,
          voucherNo,
          payee,
          amount,
          cn,
          category,
          accountType: accountTypeLabel,
          pmtMonth,
          voucherStatus: status
        });
      }

      const requiredSubs = accountTypeRequirements[accountType] || [];
      const subRequired = requiredSubs.length > 0;
      const missingFields = [];
      if (!category) missingFields.push("CATEGORIES");
      if (!accountType) missingFields.push("ACCOUNT_TYPE");
      if (subRequired && !subAccountType) missingFields.push("SUB_ACCOUNT_TYPE");
      if (missingFields.length > 0) {
        [ACTION_ITEM_UNITS.PAYABLE, ACTION_ITEM_UNITS.ADMIN].forEach(unit => {
          if (!isRuleEnabled_(settings, unit, ACTION_ITEM_RULES.MISSING_DATA)) return;
          const id = `${ACTION_ITEM_RULES.MISSING_DATA}:${unit}:${src.year}:${rowIndex}`;
          currentItems.push({
            id,
            year: src.year,
            unit,
            rule: ACTION_ITEM_RULES.MISSING_DATA,
            rowIndex,
            voucherNo,
            payee,
            amount,
            cn,
            category,
            accountType: accountTypeLabel,
            pmtMonth,
            voucherStatus: status,
            missingFields: missingFields.join(", ")
          });
          activeIds.add(id);
        });
      }

      // RULE 1: Paid with no control number
      if (status === "PAID" && !hasControlNumber) {
        [ACTION_ITEM_UNITS.PAYABLE, ACTION_ITEM_UNITS.CPO].forEach(unit => {
          if (!isRuleEnabled_(settings, unit, ACTION_ITEM_RULES.PAID_NO_CN)) return;
          const id = `${ACTION_ITEM_RULES.PAID_NO_CN}:${unit}:${src.year}:${rowIndex}`;
          currentItems.push({
            id,
            year: src.year,
            unit,
            rule: ACTION_ITEM_RULES.PAID_NO_CN,
            rowIndex,
            voucherNo,
            payee,
            amount,
            cn,
            category,
            accountType: accountTypeLabel,
            pmtMonth,
            voucherStatus: status
          });
          activeIds.add(id);
        });
      }

      // RULE 2: Unpaid, no control number, older than 30 days
      if (status === "UNPAID" && !hasControlNumber && createdAt) {
        const ageDays = Math.floor((now - createdAt) / 86400000);
        if (
          ageDays > 30 &&
          isRuleEnabled_(settings, ACTION_ITEM_UNITS.CPO, ACTION_ITEM_RULES.UNPAID_NO_CN_30D)
        ) {
          const id = `${ACTION_ITEM_RULES.UNPAID_NO_CN_30D}:CPO:${src.year}:${rowIndex}`;
          currentItems.push({
            id,
            year: src.year,
            unit: "CPO",
            rule: ACTION_ITEM_RULES.UNPAID_NO_CN_30D,
            rowIndex,
            voucherNo,
            payee,
            amount,
            cn,
            ageDays,
            category,
            accountType: accountTypeLabel,
            pmtMonth,
            voucherStatus: status
          });
          activeIds.add(id);
        }
      }

      // RULE 3: Released, still unpaid, older than 15 days
      if (status === "UNPAID" && hasControlNumber && releasedAt) {
        const ageDays = Math.floor((now - releasedAt) / 86400000);
        if (
          ageDays > 15 &&
          isRuleEnabled_(settings, ACTION_ITEM_UNITS.CPO, ACTION_ITEM_RULES.RELEASED_UNPAID_15D)
        ) {
          const id = `${ACTION_ITEM_RULES.RELEASED_UNPAID_15D}:CPO:${src.year}:${rowIndex}`;
          currentItems.push({
            id,
            year: src.year,
            unit: "CPO",
            rule: ACTION_ITEM_RULES.RELEASED_UNPAID_15D,
            rowIndex,
            voucherNo,
            payee,
            amount,
            cn,
            ageDays,
            category,
            accountType: accountTypeLabel,
            pmtMonth,
            voucherStatus: status
          });
          activeIds.add(id);
        }
      }
    });
  });

  Object.keys(duplicateTracker).forEach(voucherNoKey => {
    const entries = duplicateTracker[voucherNoKey];
    if (!entries || entries.length <= 1) return;

    entries.forEach(entry => {
      [ACTION_ITEM_UNITS.PAYABLE, ACTION_ITEM_UNITS.ADMIN].forEach(unit => {
        if (!isRuleEnabled_(settings, unit, ACTION_ITEM_RULES.DUPLICATE_VOUCHER)) return;
        const id = `${ACTION_ITEM_RULES.DUPLICATE_VOUCHER}:${unit}:2026:${entry.rowIndex}`;
        currentItems.push({
          id,
          year: "2026",
          unit,
          rule: ACTION_ITEM_RULES.DUPLICATE_VOUCHER,
          rowIndex: entry.rowIndex,
          voucherNo: entry.voucherNo,
          payee: entry.payee,
          amount: entry.amount,
          cn: entry.cn,
          category: entry.category,
          accountType: entry.accountType,
          pmtMonth: entry.pmtMonth,
          voucherStatus: entry.voucherStatus
        });
        activeIds.add(id);
      });
    });
  });

  const logData = logSheet.getDataRange().getValues();
  const logHeader = getCanonicalHeaderMap_(logSheet);
  const idCol = resolveCanonicalCol_(logHeader, ["ITEM_ID", "ITEM ID"]);
  const ruleCol = resolveCanonicalCol_(logHeader, ["RULE_KEY", "RULE KEY"]);
  const yearCol = resolveCanonicalCol_(logHeader, ["YEAR"]);
  const rowIndexCol = resolveCanonicalCol_(logHeader, ["ROW_INDEX", "ROW INDEX"]);
  const unitCol = resolveCanonicalCol_(logHeader, ["UNIT"]);
  const statusCol = resolveCanonicalCol_(logHeader, ["STATUS"]);
  const firstSeenCol = resolveCanonicalCol_(logHeader, ["FIRST_SEEN_AT", "FIRST SEEN AT"]);
  const lastSeenCol = resolveCanonicalCol_(logHeader, ["LAST_SEEN_AT", "LAST SEEN AT"]);
  const resolvedCol = resolveCanonicalCol_(logHeader, ["RESOLVED_AT", "RESOLVED AT"]);
  const resolvedByCol = resolveCanonicalCol_(logHeader, ["RESOLVED_BY", "RESOLVED BY"]);

  if (!idCol || !ruleCol || !yearCol || !rowIndexCol || !unitCol || !statusCol || !firstSeenCol || !lastSeenCol || !resolvedCol) {
    throw new Error("ACTION_ITEMS_LOG sheet header is invalid.");
  }

  const existing = new Map();
  for (let r = 1; r < logData.length; r++) {
    const id = logData[r][idCol - 1];
    if (id) existing.set(id, r + 1);
  }

  const toAppend = [];
  const logLastCol = logSheet.getLastColumn();
  currentItems.forEach(it => {
    if (!existing.has(it.id)) {
      const rowArr = new Array(logLastCol).fill("");
      rowArr[idCol - 1] = it.id;
      rowArr[ruleCol - 1] = it.rule;
      rowArr[yearCol - 1] = it.year;
      rowArr[rowIndexCol - 1] = it.rowIndex;
      rowArr[unitCol - 1] = it.unit;
      rowArr[statusCol - 1] = "PENDING";
      rowArr[firstSeenCol - 1] = now;
      rowArr[lastSeenCol - 1] = now;
      rowArr[resolvedCol - 1] = "";
      if (resolvedByCol) rowArr[resolvedByCol - 1] = "";
      toAppend.push(rowArr);
      return;
    }

    const rowNum = existing.get(it.id);
    logSheet.getRange(rowNum, ruleCol).setValue(it.rule);
    logSheet.getRange(rowNum, yearCol).setValue(it.year);
    logSheet.getRange(rowNum, rowIndexCol).setValue(it.rowIndex);
    logSheet.getRange(rowNum, unitCol).setValue(it.unit);
    logSheet.getRange(rowNum, statusCol).setValue("PENDING");
    logSheet.getRange(rowNum, lastSeenCol).setValue(now);
    logSheet.getRange(rowNum, resolvedCol).setValue("");
    if (resolvedByCol) logSheet.getRange(rowNum, resolvedByCol).setValue("");
  });

  if (toAppend.length) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, toAppend.length, logLastCol).setValues(toAppend);
  }

  existing.forEach((rowNum, id) => {
    if (activeIds.has(id)) return;
    logSheet.getRange(rowNum, statusCol).setValue("RESOLVED");
    logSheet.getRange(rowNum, resolvedCol).setValue(now);
    logSheet.getRange(rowNum, lastSeenCol).setValue(now);
    if (resolvedByCol) logSheet.getRange(rowNum, resolvedByCol).setValue("SYSTEM");
  });
}

/***************************************************************
 * PUBLIC API – DASHBOARD + NOTIFICATION PAGE
 ***************************************************************/
function parseActionItemFilters_(paramsOrUnit, statusArg) {
  const fromObject = paramsOrUnit && typeof paramsOrUnit === "object" && !Array.isArray(paramsOrUnit);
  const params = fromObject ? paramsOrUnit : { unit: paramsOrUnit, status: statusArg };
  return {
    unit: String(params.unit || "").toUpperCase(),
    status: String(params.status || "PENDING").toUpperCase()
  };
}

function getActionItems(token, paramsOrUnit, statusArg) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: "Session expired" };

    const filters = parseActionItemFilters_(paramsOrUnit, statusArg);
    syncActionItems_();

    const sheet = getSheet(CONFIG.SHEETS.ACTION_ITEMS_LOG);
    const data = sheet.getDataRange().getValues();
    const header = getCanonicalHeaderMap_(sheet);
    const locale = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSpreadsheetLocale();
    const unitCol = resolveCanonicalCol_(header, ["UNIT"]);
    const statusCol = resolveCanonicalCol_(header, ["STATUS"]);
    const ruleCol = resolveCanonicalCol_(header, ["RULE_KEY", "RULE KEY"]);
    const yearCol = resolveCanonicalCol_(header, ["YEAR"]);
    const rowIndexCol = resolveCanonicalCol_(header, ["ROW_INDEX", "ROW INDEX"]);
    const firstSeenCol = resolveCanonicalCol_(header, ["FIRST_SEEN_AT", "FIRST SEEN AT"]);
    const lastSeenCol = resolveCanonicalCol_(header, ["LAST_SEEN_AT", "LAST SEEN AT"]);

    if (!unitCol || !statusCol || !ruleCol || !yearCol || !rowIndexCol || !firstSeenCol || !lastSeenCol) {
      return { success: false, error: "ACTION_ITEMS_LOG header is invalid." };
    }

    const userUnit = roleToUnit_(session.role);
    const canAll = canViewAllUnits_(session.role);

    const items = [];
    const snapshotCache = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUnit = String(row[unitCol - 1] || "").toUpperCase();
      const rowStatus = String(row[statusCol - 1] || "").toUpperCase();

      if (!canAll && rowUnit !== userUnit) continue;
      if (filters.status && filters.status !== "ALL" && rowStatus !== filters.status) continue;
      if (filters.unit && filters.unit !== "ALL" && rowUnit !== filters.unit) continue;

      const text = buildActionItemText_(row[ruleCol - 1], rowUnit, {});
      const rowYear = row[yearCol - 1];
      const rowIndex = row[rowIndexCol - 1];
      const cacheKey = `${rowYear}:${rowIndex}`;
      if (!snapshotCache[cacheKey]) {
        snapshotCache[cacheKey] = getActionItemCurrentVoucherMeta_(rowYear, rowIndex, locale) || null;
      }
      const snap = snapshotCache[cacheKey];

      items.push({
        unit: rowUnit,
        year: rowYear,
        itemStatus: rowStatus,
        status: snap ? snap.status : rowStatus,
        voucherNumber: snap ? snap.voucherNumber : "",
        payee: snap ? snap.payee : "",
        amount: snap ? snap.amount : 0,
        controlNumber: snap ? snap.controlNumber : "",
        pmtMonth: snap ? snap.pmtMonth : "",
        category: snap ? snap.category : "",
        accountType: snap ? snap.accountType : "",
        rowIndex: rowIndex,
        title: text.title,
        message: text.message,
        severity: text.severity,
        firstSeenAt: row[firstSeenCol - 1],
        lastSeenAt: row[lastSeenCol - 1]
      });
    }

    const severityRank = { danger: 1, warning: 2, info: 3 };
    items.sort((a, b) => {
      const rankA = severityRank[String(a.severity || "info")] || 9;
      const rankB = severityRank[String(b.severity || "info")] || 9;
      return rankA - rankB;
    });

    return { success: true, items, count: items.length };
  } catch (e) {
    return { success: false, error: "Failed to load action items: " + e.message };
  }
}

function getActionItemCount(token, paramsOrUnit, statusArg) {
  const res = getActionItems(token, paramsOrUnit, statusArg);
  if (!res.success) return res;
  return { success: true, count: res.count };
}

function getActionItemSettings(token) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: "Session expired" };
    if (!canViewAllUnits_(session.role)) return { success: false, error: "Unauthorized" };
    const settings = loadActionItemSettings_();
    return { success: true, settings };
  } catch (e) {
    return { success: false, error: "Failed to get settings: " + e.message };
  }
}

function saveActionItemSettings(token, settingsObj) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: "Session expired" };
    if (!canViewAllUnits_(session.role)) return { success: false, error: "Unauthorized" };
    const sheet = getSheet(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS);
    const rows = [];
    const unitObj = settingsObj?.unit || {};
    const units = [ACTION_ITEM_UNITS.PAYABLE, ACTION_ITEM_UNITS.CPO, ACTION_ITEM_UNITS.ADMIN];

    units.forEach(u => {
      const rules = unitObj[u] || {};
      Object.keys(ACTION_ITEM_RULES).forEach(ruleKey => {
        const rule = ACTION_ITEM_RULES[ruleKey];
        const enabled = rules[rule] !== false;
        rows.push([u, rule, enabled]);
      });
    });

    // Clear existing settings and write new ones
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, 3).setValues(rows);

    return { success: true, message: "Action Items settings saved." };
  } catch (e) {
    return { success: false, error: "Failed to save settings: " + e.message };
  }
}
