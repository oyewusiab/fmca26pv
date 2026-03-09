# Payable Vouchers 2026 - Backend Code Guide

This guide contains the necessary Google Apps Script code to fix existing errors and implement the new Action Items and Announcements features.

## 1. Fixing `SPREADSHEET_ID` Error

The error `ReferenceError: SPREADSHEET_ID is not defined` indicates that the `CONFIG` object is not available when `getSheet()` is called. To make your code more robust, you should have a single, well-defined `Config.gs` file and ensure it's loaded first.

Here is a corrected and more robust way to handle your configuration and sheet access.

### `Config.gs`

Create a new script file named `Config.gs` if you don't have one and place your `CONFIG` object there.

```javascript
const CONFIG = {
  SPREADSHEET_ID: '1Idm6P45SG1JA9GEiFWdDT9GT9fgYshx4q5NdqYAsFsA',
  SHEETS: {
    VOUCHERS_2026: '2026 VOUCHERS',
    VOUCHERS_2025: '2025 VOUCHERS',
    VOUCHERS_2024: '2024 VOUCHERS',
    VOUCHERS_2023: '2023 VOUCHERS',
    VOUCHERS_BEFORE_2023: '<2023 VOUCHERS',
    USERS: 'USERS',
    ACTION_ITEMS_LOG: 'ACTION_ITEMS_LOG',
    ACTION_ITEMS_SETTINGS: 'ACTION_ITEMS_SETTINGS',
    ANNOUNCEMENTS: 'ANNOUNCEMENTS',
    SYSTEM_CONFIG: 'SYSTEM_CONFIG'
  },
  SESSION_DURATION: 8 * 60 * 60 * 1000, // 8 hours
  VOUCHER_COLUMNS: {
    STATUS: 1, PMT_MONTH: 2, PAYEE: 3, ACCOUNT_OR_MAIL: 4, PARTICULAR: 5,
    CONTRACT_SUM: 6, GROSS_AMOUNT: 7, NET: 8, VAT: 9, WHT: 10, STAMP_DUTY: 11,
    CATEGORIES: 12, TOTAL_GROSS: 13, CONTROL_NUMBER: 14, OLD_VOUCHER_NUMBER: 15,
    DATE: 16, ACCOUNT_TYPE: 17, CREATED_AT: 18, RELEASED_AT: 19
  },
  USER_COLUMNS: {
    NAME: 1, EMAIL: 2, PASSWORD: 3, ROLE: 4, ACTIVE: 5, USERNAME: 6,
    PHONE: 7, DEPARTMENT: 8, UPDATED_AT: 9
  },
  ROLES: {
    PAYABLE_STAFF: 'Payable Unit Staff',
    PAYABLE_HEAD: 'Payable Unit Head',
    CPO: 'CPO',
    AUDIT: 'Audit Unit',
    DDFA: 'DDFA',
    DFA: 'DFA',
    ADMIN: 'ADMIN'
  }
};
```

### `SheetUtils.gs`

Create a new file `SheetUtils.gs` for utility functions. The `getSheet` function is improved to prevent errors.

```javascript
/**
 * Gets a sheet by name. Throws an error if not found.
 * @param {string} sheetName - The name of the sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found in spreadsheet.`);
  }
  return sheet;
}

/**
 * Gets the header columns of a sheet as a map of name to column index (1-based).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @returns {Object.&lt;string, number&gt;} A map of header names to column numbers.
 */
function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  const map = {};
  headers.forEach((header, index) => {
    if (header) {
      map[String(header).toUpperCase().trim()] = index + 1;
    }
  });
  return map;
}
```

By separating these, you ensure your `CONFIG` is always available. Make sure this file is at the top of your script file list in the Apps Script editor.

## 2. Action Items Backend (`action-items.gs`)

Create a new script file named `action-items.gs` and paste the entire code below into it. This script provides all the backend functionality for creating, managing, and retrieving action items based on the rules you specified.
Create a new script file named `action-items.gs` and paste the code below. This aligns with your specific logic for the Action Item System engine.

```javascript
/***************************************************************
 * ACTION ITEM SYSTEM – COMPLETE ENGINE
 ***************************************************************/

const ACTION_ITEM_RULES = {
  PAID_NO_CN: "PAID_NO_CN",
  UNPAID_NO_CN_30D: "UNPAID_NO_CN_30D",
  RELEASED_UNPAID_15D: "RELEASED_UNPAID_15D"
};

const ACTION_ITEM_UNITS = {
  PAYABLE: "PAYABLE",
  CPO: "CPO"
};

/**
 * Ensures the necessary sheets for action items exist.
 */
/***************************************************************
 * INITIAL SETUP (Auto-creates required sheets)
 ***************************************************************/
function ensureActionItemSheets_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  if (!ss.getSheetByName(CONFIG.SHEETS.ACTION_ITEMS_LOG)) {
    const sheet = ss.insertSheet(CONFIG.SHEETS.ACTION_ITEMS_LOG);
    sheet.getRange(1, 1, 1, 13).setValues([[
      "ITEM_ID", "RULE_KEY", "UNIT", "YEAR", "ROW_INDEX", "VOUCHER_NO",
      "PAYEE", "AMOUNT", "CONTROL_NO", "STATUS", "FIRST_SEEN_AT",
      "LAST_SEEN_AT", "RESOLVED_AT"

  // ===== LOG SHEET =====
  let log = ss.getSheetByName(CONFIG.SHEETS.ACTION_ITEMS_LOG);
  if (!log) {
    log = ss.insertSheet(CONFIG.SHEETS.ACTION_ITEMS_LOG);
    log.getRange(1, 1, 1, 13).setValues([[
      "ITEM_ID",
      "RULE_KEY",
      "UNIT",
      "YEAR",
      "ROW_INDEX",
      "VOUCHER_NO",
      "PAYEE",
      "AMOUNT",
      "CONTROL_NO",
      "STATUS", // PENDING | RESOLVED
      "FIRST_SEEN_AT",
      "LAST_SEEN_AT",
      "RESOLVED_AT"
    ]]);
  }
  if (!ss.getSheetByName(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS)) {
    const sheet = ss.insertSheet(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS);
    sheet.getRange(1, 1, 1, 3).setValues([["UNIT", "RULE_KEY", "ENABLED"]]);
    const rows = [

  // ===== SETTINGS SHEET =====
  let settings = ss.getSheetByName(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS);
  if (!settings) {
    settings = ss.insertSheet(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS);
    settings.getRange(1, 1, 1, 3)
      .setValues([["UNIT", "RULE_KEY", "ENABLED"]]);

    const defaults = [
      [ACTION_ITEM_UNITS.PAYABLE, ACTION_ITEM_RULES.PAID_NO_CN, true],
      [ACTION_ITEM_UNITS.CPO, ACTION_ITEM_RULES.PAID_NO_CN, true],
      [ACTION_ITEM_UNITS.CPO, ACTION_ITEM_RULES.UNPAID_NO_CN_30D, true],
      [ACTION_ITEM_UNITS.CPO, ACTION_ITEM_RULES.RELEASED_UNPAID_15D, true]
    ];
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);

    settings.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }
}

/**
 * Converts various date string formats to a Date object.
 */
function parseDateTimeFlexible_(value) {
  if (value instanceof Date) return value;
  const s = String(value || "").trim();
  if (!s) return null;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}

/**
 * Maps a user role to an action item unit.
 */
/***************************************************************
 * ROLE HELPERS
 ***************************************************************/
function roleToUnit_(role) {
  if (role === CONFIG.ROLES.CPO) return ACTION_ITEM_UNITS.CPO;
  if ([CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.PAYABLE_STAFF].includes(role)) return ACTION_ITEM_UNITS.PAYABLE;
  if ([CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.PAYABLE_STAFF].includes(role))
    return ACTION_ITEM_UNITS.PAYABLE;
  return null;
}

/**
 * Checks if a role can view all action items.
 */
function canViewAllUnits_(role) {
  return [CONFIG.ROLES.ADMIN, CONFIG.ROLES.DDFA, CONFIG.ROLES.DFA].includes(role);
  return [CONFIG.ROLES.ADMIN, CONFIG.ROLES.DDFA, CONFIG.ROLES.DFA]
    .includes(role);
}

/**
 * Loads and parses the action item settings from its sheet.
 */
/***************************************************************
 * SETTINGS LOADER
 ***************************************************************/
function loadActionItemSettings_() {
  ensureActionItemSheets_();
  const sheet = getSheet(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS);
  const data = sheet.getDataRange().getValues();

  const settings = { unit: { PAYABLE: {}, CPO: {} } };

  for (let i = 1; i < data.length; i++) {
    const unit = String(data[i] || "").trim().toUpperCase();
    const rule = String(data[i] || "").trim();
    const enabled = data[i] !== false;
    if (unit && rule && settings.unit[unit]) {
    const unit = String(data[i][0] || "").toUpperCase();
    const rule = String(data[i][1] || "").trim();
    const enabled = String(data[i][2]).toUpperCase() !== "FALSE";

    if (settings.unit[unit])
      settings.unit[unit][rule] = enabled;
    }
  }

  return settings;
}

/**
 * Checks if a specific rule is enabled for a unit.
 */
function isRuleEnabled_(settings, unit, ruleKey) {
  if (!unit) return false;
  const value = settings?.unit?.[unit]?.[ruleKey];
  return value !== false; // Enabled by default
  return settings?.unit?.[unit]?.[ruleKey] !== false;
}

/**
 * Generates the human-readable text for an action item.
 */
/***************************************************************
 * CLEAN ACTION TEXT (Refined Statements)
 ***************************************************************/
function buildActionItemText_(ruleKey, unit, meta) {

  switch (ruleKey) {
    case ACTION_ITEM_RULES.PAID_NO_CN:
      return unit === ACTION_ITEM_UNITS.PAYABLE
        ? { title: "Paid Voucher Not Released", message: "This voucher is marked PAID but has no Control Number. Consider assigning one and releasing it to CPO.", severity: "warning" }
        : { title: "Paid Voucher Not Received", message: "This voucher is marked PAID but has not been received from the Payable Unit. Consider requesting it.", severity: "info" };
      if (unit === ACTION_ITEM_UNITS.PAYABLE) {
        return {
          title: "Paid Voucher Not Released to CPO",
          message: "This voucher is marked as PAID but has no Control Number. Please assign a Control Number and release it to the CPO Unit.",
          severity: "warning"
        };
      }
      return {
        title: "Paid Voucher Awaiting Release from Payable",
        message: "This voucher is marked as PAID but has not been released by the Payable Unit. Consider requesting it for processing.",
        severity: "info"
      };

    case ACTION_ITEM_RULES.UNPAID_NO_CN_30D:
      return { title: "Delayed Voucher in Payable", message: `This voucher has been in the Payable Unit for over ${meta.ageDays || 30} days without payment or release. Consider requesting it for processing.`, severity: "warning" };
      return {
        title: "Voucher Delayed in Payable Unit",
        message: `This voucher has remained UNPAID for over ${meta.ageDays || 30} days without a Control Number. Consider requesting it from Payable for payment processing.`,
        severity: "warning"
      };

    case ACTION_ITEM_RULES.RELEASED_UNPAID_15D:
      return { title: "Released Voucher Still Unpaid", message: `This voucher was released over ${meta.ageDays || 15} days ago but is still marked as UNPAID. Please update its status.`, severity: "danger" };
      return {
        title: "Released Voucher Still Marked Unpaid",
        message: `This voucher was released over ${meta.ageDays || 15} days ago but remains marked as UNPAID. Please update the payment status if completed.`,
        severity: "danger"
      };

    default:
      return { title: "Action Item", message: "An action is required for this voucher.", severity: "info" };
      return { title: "Action Required", message: "", severity: "info" };
  }
}

/**
 * Main sync function to find and log all current action items.
 * Should be run on a time-based trigger (e.g., every 15 minutes).
 */
/***************************************************************
 * MAIN SYNC ENGINE (AUTO RESOLVE + AUTO CREATE)
 ***************************************************************/
function syncActionItems_() {

  ensureActionItemSheets_();
  const settings = loadActionItemSettings_();
  const now = new Date();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const header = getHeaderMap_(sheet);

  const colCreated = header["CREATED AT"];
  const colReleased = header["RELEASED AT"];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const now = new Date();

  const currentItems = [];
  const activeIds = new Set();

  for (let i = 0; i < data.length; i++) {
  data.forEach((row, i) => {

    const rowIndex = i + 2;
    const status = String(data[i][CONFIG.VOUCHER_COLUMNS.STATUS - 1] || "").trim();
    const cn = String(data[i][CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER - 1] || "").trim();
    const createdAt = parseDateTimeFlexible_(data[i][CONFIG.VOUCHER_COLUMNS.CREATED_AT - 1]);
    const releasedAt = parseDateTimeFlexible_(data[i][header['RELEASED AT'] - 1]);

    const baseInfo = {
      rowIndex,
      voucherNo: String(data[i][CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1] || ""),
      payee: String(data[i][CONFIG.VOUCHER_COLUMNS.PAYEE - 1] || ""),
      amount: parseFloat(data[i][CONFIG.VOUCHER_COLUMNS.GROSS_AMOUNT - 1]) || 0,
      cn
    };
    const status = String(row[CONFIG.VOUCHER_COLUMNS.STATUS - 1] || "").trim();
    const cn = String(row[CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER - 1] || "").trim();
    const voucherNo = row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1];
    const payee = row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1];
    const amount = row[CONFIG.VOUCHER_COLUMNS.GROSS_AMOUNT - 1] || 0;

    // Rule 1: Paid but no CN
    const createdAt = row[colCreated - 1];
    const releasedAt = row[colReleased - 1] || createdAt;

    // RULE 1
    if (status === "Paid" && !cn) {
      if (isRuleEnabled_(settings, ACTION_ITEM_UNITS.PAYABLE, ACTION_ITEM_RULES.PAID_NO_CN))
        currentItems.push({ ...baseInfo, id: `PAID_NO_CN:PAYABLE:2026:${rowIndex}`, ruleKey: ACTION_ITEM_RULES.PAID_NO_CN, unit: ACTION_ITEM_UNITS.PAYABLE });
      if (isRuleEnabled_(settings, ACTION_ITEM_UNITS.CPO, ACTION_ITEM_RULES.PAID_NO_CN))
        currentItems.push({ ...baseInfo, id: `PAID_NO_CN:CPO:2026:${rowIndex}`, ruleKey: ACTION_ITEM_RULES.PAID_NO_CN, unit: ACTION_ITEM_UNITS.CPO });
      [ACTION_ITEM_UNITS.PAYABLE, ACTION_ITEM_UNITS.CPO]
        .forEach(unit => {
          if (!isRuleEnabled_(settings, unit, ACTION_ITEM_RULES.PAID_NO_CN)) return;
          const id = `${ACTION_ITEM_RULES.PAID_NO_CN}:${unit}:2026:${rowIndex}`;
          currentItems.push({ id, unit, rule: ACTION_ITEM_RULES.PAID_NO_CN, rowIndex, voucherNo, payee, amount, cn });
          activeIds.add(id);
        });
    }

    // Rule 2: Unpaid, no CN, >30 days
    if (status === "Unpaid" && !cn && createdAt) {
      const ageDays = Math.floor((now - createdAt) / (1000 * 60 * 60 * 24));
      if (ageDays > 30 && isRuleEnabled_(settings, ACTION_ITEM_UNITS.CPO, ACTION_ITEM_RULES.UNPAID_NO_CN_30D)) {
        currentItems.push({ ...baseInfo, id: `UNPAID_NO_CN_30D:CPO:2026:${rowIndex}`, ruleKey: ACTION_ITEM_RULES.UNPAID_NO_CN_30D, unit: ACTION_ITEM_UNITS.CPO, ageDays });
    // RULE 2
    if (status === "Unpaid" && !cn && createdAt instanceof Date) {
      const ageDays = Math.floor((now - createdAt) / 86400000);
      if (ageDays > 30 &&
          isRuleEnabled_(settings, ACTION_ITEM_UNITS.CPO, ACTION_ITEM_RULES.UNPAID_NO_CN_30D)) {
        const id = `${ACTION_ITEM_RULES.UNPAID_NO_CN_30D}:CPO:2026:${rowIndex}`;
        currentItems.push({ id, unit: "CPO", rule: ACTION_ITEM_RULES.UNPAID_NO_CN_30D, rowIndex, voucherNo, payee, amount, cn, ageDays });
        activeIds.add(id);
      }
    }

    // Rule 3: Unpaid, has CN, >15 days since release
    const releaseDate = releasedAt || (cn ? createdAt : null); // Fallback to createdAt if releasedAt is missing but CN exists
    if (status === "Unpaid" && cn && releaseDate) {
      const ageDays = Math.floor((now - releaseDate) / (1000 * 60 * 60 * 24));
      if (ageDays > 15 && isRuleEnabled_(settings, ACTION_ITEM_UNITS.CPO, ACTION_ITEM_RULES.RELEASED_UNPAID_15D)) {
        currentItems.push({ ...baseInfo, id: `RELEASED_UNPAID_15D:CPO:2026:${rowIndex}`, ruleKey: ACTION_ITEM_RULES.RELEASED_UNPAID_15D, unit: ACTION_ITEM_UNITS.CPO, ageDays });
    // RULE 3
    if (status === "Unpaid" && cn && releasedAt instanceof Date) {
      const ageDays = Math.floor((now - releasedAt) / 86400000);
      if (ageDays > 15 &&
          isRuleEnabled_(settings, ACTION_ITEM_UNITS.CPO, ACTION_ITEM_RULES.RELEASED_UNPAID_15D)) {
        const id = `${ACTION_ITEM_RULES.RELEASED_UNPAID_15D}:CPO:2026:${rowIndex}`;
        currentItems.push({ id, unit: "CPO", rule: ACTION_ITEM_RULES.RELEASED_UNPAID_15D, rowIndex, voucherNo, payee, amount, cn, ageDays });
        activeIds.add(id);
      }
    }
  }

  updateActionItemsLog_(currentItems, now);
}
  });

/**
 * Updates the ACTION_ITEMS_LOG sheet based on current findings.
 */
function updateActionItemsLog_(currentItems, now) {
  const logSheet = getSheet(CONFIG.SHEETS.ACTION_ITEMS_LOG);
  const logData = logSheet.getDataRange().getValues();
  const h = getHeaderMap_(logSheet);
  const idCol = h["ITEM_ID"];
  const statusCol = h["STATUS"];
  const logHeader = getHeaderMap_(logSheet);

  const logMap = new Map(); // id -> { rowIndex, status }
  const idCol = logHeader["ITEM_ID"];
  const statusCol = logHeader["STATUS"];
  const resolvedCol = logHeader["RESOLVED_AT"];
  const lastSeenCol = logHeader["LAST_SEEN_AT"];

  const existing = new Map();
  for (let r = 1; r < logData.length; r++) {
    const id = String(logData[r][idCol - 1] || "").trim();
    if (id) logMap.set(id, { rowIndex: r + 1, status: String(logData[r][statusCol - 1] || "") });
    const id = logData[r][idCol - 1];
    if (id) existing.set(id, r + 1);
  }

  const toAppend = [];
  const currentIds = new Set();

  currentItems.forEach(item => {
    currentIds.add(item.id);
    const existing = logMap.get(item.id);
    if (!existing) {
  currentItems.forEach(it => {
    if (!existing.has(it.id)) {
      toAppend.push([
        item.id, item.ruleKey, item.unit, "2026", item.rowIndex, item.voucherNo,
        item.payee, item.amount, item.cn, "PENDING", now, now, ""
        it.id, it.rule, it.unit, "2026",
        it.rowIndex, it.voucherNo, it.payee,
        it.amount, it.cn || "",
        "PENDING", now, now, ""
      ]);
    } else if (existing.status === "RESOLVED") {
      logSheet.getRange(existing.rowIndex, statusCol).setValue("PENDING");
      logSheet.getRange(existing.rowIndex, h["LAST_SEEN_AT"]).setValue(now);
      logSheet.getRange(existing.rowIndex, h["RESOLVED_AT"]).setValue("");
    } else {
      logSheet.getRange(existing.rowIndex, h["LAST_SEEN_AT"]).setValue(now);
      const rowNum = existing.get(it.id);
      logSheet.getRange(rowNum, statusCol).setValue("PENDING");
      logSheet.getRange(rowNum, lastSeenCol).setValue(now);
      logSheet.getRange(rowNum, resolvedCol).setValue("");
    }
  });

  if (toAppend.length) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, toAppend.length, toAppend.length).setValues(toAppend);
  }
  if (toAppend.length)
    logSheet.getRange(logSheet.getLastRow() + 1, 1, toAppend.length, 13)
      .setValues(toAppend);

  logMap.forEach((value, id) => {
    if (value.status === "PENDING" && !currentIds.has(id)) {
      logSheet.getRange(value.rowIndex, statusCol).setValue("RESOLVED");
      logSheet.getRange(value.rowIndex, h["RESOLVED_AT"]).setValue(now);
  // Auto resolve
  existing.forEach((rowNum, id) => {
    if (!activeIds.has(id)) {
      logSheet.getRange(rowNum, statusCol).setValue("RESOLVED");
      logSheet.getRange(rowNum, resolvedCol).setValue(now);
      logSheet.getRange(rowNum, lastSeenCol).setValue(now);
    }
  });
}

/**
 * API endpoint to get action items for the current user.
 */
function getActionItems(token, params = {}) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: "Session expired" };
/***************************************************************
 * PUBLIC API – DASHBOARD + NOTIFICATION PAGE
 ***************************************************************/
function getActionItems(token, unit, status) {

    // Only authorized roles can see action items
    const allowedRoles = [
      CONFIG.ROLES.PAYABLE_STAFF, CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.CPO,
      CONFIG.ROLES.ADMIN, CONFIG.ROLES.DDFA, CONFIG.ROLES.DFA
    ];
    if (!hasPermission(session.role, allowedRoles)) {
      return { success: true, items: [], count: 0 };
    }
  const session = getSession(token);
  if (!session) return { success: false, error: "Session expired" };

    syncActionItems_(); // Run a fresh sync on demand
  syncActionItems_();

    const logSheet = getSheet(CONFIG.SHEETS.ACTION_ITEMS_LOG);
    const data = logSheet.getDataRange().getValues();
    const h = getHeaderMap_(logSheet);
  const sheet = getSheet(CONFIG.SHEETS.ACTION_ITEMS_LOG);
  const data = sheet.getDataRange().getValues();
  const header = getHeaderMap_(sheet);

    const canViewAll = canViewAllUnits_(session.role);
    const userUnit = roleToUnit_(session.role);
  const userUnit = roleToUnit_(session.role);
  const canAll = canViewAllUnits_(session.role);

    const unitFilter = canViewAll ? (params.unit || "ALL") : userUnit;
    const statusFilter = params.status || "PENDING";
  const items = [];

    const items = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUnit = String(row[h["UNIT"] - 1] || "").toUpperCase();
      const rowStatus = String(row[h["STATUS"] - 1] || "").toUpperCase();
  for (let i = 1; i < data.length; i++) {

      if (unitFilter !== "ALL" && rowUnit !== unitFilter) continue;
      if (statusFilter !== "ALL" && rowStatus !== statusFilter) continue;
    const row = data[i];
    const rowUnit = row[header["UNIT"] - 1];
    const rowStatus = row[header["STATUS"] - 1];

      const ruleKey = String(row[h["RULE_KEY"] - 1] || "");
      const text = buildActionItemText_(ruleKey, rowUnit, {});
    if (!canAll && rowUnit !== userUnit) continue;
    if (status && rowStatus !== status) continue;
    if (unit && unit !== "ALL" && rowUnit !== unit) continue;

      items.push({
        ruleKey,
        unit: rowUnit,
        itemStatus: rowStatus,
        year: "2026",
        rowIndex: Number(row[h["ROW_INDEX"] - 1] || 0),
        voucherNumber: String(row[h["VOUCHER_NO"] - 1] || ""),
        payee: String(row[h["PAYEE"] - 1] || ""),
        grossAmount: Number(row[h["AMOUNT"] - 1] || 0),
        controlNumber: String(row[h["CONTROL_NO"] - 1] || ""),
        title: text.title,
        message: text.message,
        severity: text.severity
      });
    }
    const text = buildActionItemText_(
      row[header["RULE_KEY"] - 1],
      rowUnit,
      {}
    );

    return { success: true, items, count: items.length };
  } catch (e) {
    return { success: false, error: "Failed to get action items: " + e.message };
    items.push({
      unit: rowUnit,
      status: rowStatus,
      voucherNumber: row[header["VOUCHER_NO"] - 1],
      payee: row[header["PAYEE"] - 1],
      amount: row[header["AMOUNT"] - 1],
      controlNumber: row[header["CONTROL_NO"] - 1],
      title: text.title,
      message: text.message,
      severity: text.severity
    });
  }

  return { success: true, items, count: items.length };
}

/**
 * API endpoint to get only the count of action items.
 */
function getActionItemCount(token, params = {}) {
  const result = getActionItems(token, params);
  if (!result.success) return result;
  return { success: true, count: result.count || 0 };
function getActionItemCount(token) {
  const res = getActionItems(token);
  if (!res.success) return res;
  return { success: true, count: res.count };
}

/**
 * API endpoint to get action item settings (for admins).
 */
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

/**
 * API endpoint to save action item settings (for admins).
 */
function saveActionItemSettings(token, settingsObj) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: "Session expired" };
    if (!canViewAllUnits_(session.role)) return { success: false, error: "Unauthorized" };

    ensureActionItemSheets_();
    const sheet = getSheet(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS);
    const rows = [];
    const unitObj = settingsObj?.unit || {};
    const units = [ACTION_ITEM_UNITS.PAYABLE, ACTION_ITEM_UNITS.CPO];

    Object.keys(ACTION_ITEM_UNITS).forEach(unitKey => {
      const u = ACTION_ITEM_UNITS[unitKey];
    units.forEach(u => {
      const rules = unitObj[u] || {};
      Object.keys(ACTION_ITEM_RULES).forEach(ruleKey => {
        const enabled = rules[ruleKey] !== false;
        rows.push([u, ruleKey, enabled]);
        const rule = ACTION_ITEM_RULES[ruleKey];
        const enabled = rules[rule] !== false;
        rows.push([u, rule, enabled]);
      });
    });

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    }
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }
    // Clear existing settings and write new ones
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, 3).setValues(rows);

    CacheService.getScriptCache().remove("action_items_sync_ts");
    return { success: true, message: "Action Items settings saved." };
  } catch (e) {
    return { success: false, error: "Failed to save settings: " + e.message };
  }
}

```

### Setup

1.  **Create a time-based trigger** in your Apps Script project to run the `syncActionItems_` function periodically (e.g., every 15 or 30 minutes). This keeps the action items log up-to-date for fast retrieval.
2.  Ensure you add `getActionItems`, `getActionItemCount`, `getActionItemSettings`, and `saveActionItemSettings` to your main `doPost` and `doGet` handlers so the client-side API can call them.

## 3. Announcements Backend

Here is the backend code for the announcements feature. Add these functions to a new `Announcements.gs` file.

```javascript
/**
 * Ensures the ANNOUNCEMENTS sheet exists.
 */
function ensureAnnouncementsSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  if (!ss.getSheetByName(CONFIG.SHEETS.ANNOUNCEMENTS)) {
    const sheet = ss.insertSheet(CONFIG.SHEETS.ANNOUNCEMENTS);
    sheet.getRange(1, 1, 1, 8).setValues([[
      "ID", "MESSAGE", "DISPLAY_LOCATIONS", "TARGET_USERS",
      "EXPIRES_AT", "ALLOW_DISMISS", "CREATED_AT", "CREATED_BY"
    ]]);
    sheet.setColumnWidth(2, 400); // Message column
  }
}

/**
 * API: Creates a new announcement (Admin only).
 */
function createAnnouncement(token, announcement) {
  try {
    const session = getSession(token);
    if (!session || session.role !== CONFIG.ROLES.ADMIN) {
      return { success: false, error: "Unauthorized" };
    }

    ensureAnnouncementsSheet_();
    const sheet = getSheet(CONFIG.SHEETS.ANNOUNCEMENTS);

    const newRow = [
      Utilities.getUuid(),
      announcement.message,
      JSON.stringify(announcement.locations || []),
      JSON.stringify(announcement.targets || []),
      new Date(announcement.expiresAt),
      announcement.allowDismiss,
      new Date(),
      session.email
    ];

    sheet.appendRow(newRow);
    return { success: true, message: "Announcement created successfully." };

  } catch (e) {
    return { success: false, error: "Failed to create announcement: " + e.message };
  }
}

/**
 * API: Gets active announcements for the current user.
 */
function getActiveAnnouncements(token) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: "Session expired" };

    ensureAnnouncementsSheet_();
    const sheet = getSheet(CONFIG.SHEETS.ANNOUNCEMENTS);
    if (sheet.getLastRow() < 2) return { success: true, announcements: [] };

    const data = sheet.getDataRange().getValues();
    const h = getHeaderMap_(sheet);
    const now = new Date();
    const userProperties = PropertiesService.getUserProperties();
    const dismissed = JSON.parse(userProperties.getProperty('dismissedAnnouncements') || '{}');

    const announcements = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = row[h.ID - 1];
      const expiresAt = parseDateTimeFlexible_(row[h.EXPIRES_AT - 1]);

      if (expiresAt < now || dismissed[id]) continue;

      const targets = JSON.parse(row[h.TARGET_USERS - 1] || '[]');
      const isTargeted = targets.includes("ALL") || targets.includes(session.email) || targets.includes(session.role);

      if (isTargeted) {
        announcements.push({
          id: id,
          message: row[h.MESSAGE - 1],
          locations: JSON.parse(row[h.DISPLAY_LOCATIONS - 1] || '[]'),
          allowDismiss: row[h.ALLOW_DISMISS]
        });
      }
    }

    return { success: true, announcements };

  } catch (e) {
    return { success: false, error: "Failed to get announcements: " + e.message };
  }
}

/**
 * API: Marks an announcement as dismissed for the current user.
 */
function dismissAnnouncement(token, announcementId) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: "Session expired" };

    const userProperties = PropertiesService.getUserProperties();
    const dismissed = JSON.parse(userProperties.getProperty('dismissedAnnouncements') || '{}');
    dismissed[announcementId] = true;
    userProperties.setProperty('dismissedAnnouncements', JSON.stringify(dismissed));

    return { success: true };
  } catch (e) {
    return { success: false, error: "Failed to dismiss announcement: " + e.message };
  }
}
```

### Setup

1.  Add `createAnnouncement`, `getActiveAnnouncements`, and `dismissAnnouncement` to your main `doPost`/`doGet` handlers.
2.  The script will automatically create the `ANNOUNCEMENTS` sheet on first run.