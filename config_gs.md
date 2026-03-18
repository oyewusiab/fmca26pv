const CONFIG = {
  SPREADSHEET_ID: '1Idm6P45SG1JA9GEiFWdDT9GT9fgYshx4q5NdqYAsFsA',
  SHEETS: {
    VOUCHERS_2026: '2026 VOUCHERS',
    VOUCHERS_2025: '2025 VOUCHERS',
    VOUCHERS_2024: '2024 VOUCHERS',
    VOUCHERS_2023: '2023 VOUCHERS',
    VOUCHERS_BEFORE_2023: '<2023 VOUCHERS',
    TAX_PAYMENTS: 'TAX_PAYMENTS',
    TAX_SCHEDULE: 'TAX_SCHEDULE',
    TAX_REPORTS: 'TAX_REPORTS',
    USERS: 'USERS',
    ANNOUNCEMENTS: 'ANNOUNCEMENTS',
    ACTION_ITEMS_LOG: 'ACTION_ITEMS_LOG',
    ACTION_ITEMS_SETTINGS: 'ACTION_ITEMS_SETTINGS',
    SYSTEM_CONFIG: 'SYSTEM_CONFIG'
  },
  TAX_TYPES: {
    VAT: 'VAT',
    WHT: 'Withholding Tax',
    STAMP_DUTY: 'Stamp Duty'
  }, 
  SESSION_DURATION: 8 * 60 * 60 * 1000, // 8 hours
   VOUCHER_COLUMNS: {
    STATUS: 1, PMT_MONTH: 2, PAYEE: 3, ACCOUNT_OR_MAIL: 4, PARTICULAR: 5,
    CONTRACT_SUM: 6, GROSS_AMOUNT: 7, NET: 8, VAT: 9, WHT: 10, STAMP_DUTY: 11,
    CATEGORIES: 12, TOTAL_GROSS: 13, CONTROL_NUMBER: 14, OLD_VOUCHER_NUMBER: 15,
    DATE: 16, ACCOUNT_TYPE: 17, CREATED_AT: 18, RELEASED_AT: 19, ATTACHMENT_URL: 20,
    OLD_VOUCHER_AVAILABLE: 21
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
    ADMIN: 'ADMIN',
    TAX: 'Tax Unit'
  }
};

/**
 * Gets a sheet by name. Throws an error if not found.
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
 * Diagnostic utility
 */
function diagnosticCheck() {
  const requiredFunctions = [
    'getSheet'
  ];
  
  Logger.log('=== DIAGNOSTIC CHECK ===');
  
  requiredFunctions.forEach(funcName => {
    try {
      if (typeof eval(funcName) === 'function') {
        Logger.log('✅ ' + funcName);
      } else {
        Logger.log('❌ ' + funcName + ' - NOT A FUNCTION');
      }
    } catch (e) {
      Logger.log('❌ ' + funcName + ' - NOT DEFINED: ' + e.message);
    }
  });
}
