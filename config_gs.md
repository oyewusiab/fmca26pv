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
    SYSTEM_CONFIG: 'SYSTEM_CONFIG',
    DEBT_PROFILE_REQUESTS: 'DEBT_PROFILE_REQUESTS'
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
    OLD_VOUCHER_AVAILABLE: 21, SUB_ACCOUNT_TYPE: 22
  },
  DEBT_REQUEST_COLUMNS: {
    REQUEST_ID: 1, TIMESTAMP: 2, REQUESTER_EMAIL: 3, REQUESTER_NAME: 4,
    FILTERS: 5, STATUS: 6, APPROVER_EMAIL: 7, APPROVAL_DATE: 8,
    COMMENTS: 9, REPORT_ID: 10, REPORT_TITLE: 11, EXECUTIVE_SUMMARY: 12,
    ANALYSIS: 13, RECOMMENDATIONS: 14
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

/**
 * Gets the system configuration including Account Types and Sub Account Types.
 * Reads from SYSTEM_CONFIG sheet rows 7+.
 * B column: Account Type
 * C column: Comma-separated Sub Account Types
 */
function getSystemConfig(token) {
  try {
    // Validate session if token is provided
    if (typeof getSession === 'function' && token) {
      const session = getSession(token);
      if (!session) {
        return { success: false, error: 'Session expired' };
      }
    }
    
    // Default config object
    const config = {
      accountTypes: {} // Format: { "RFA": ["TRF", "LRF"], "CAPITAL": [] }
    };
    
    const sheet = getSheet(CONFIG.SHEETS.SYSTEM_CONFIG);
    if (!sheet) {
      return { success: false, error: 'SYSTEM_CONFIG sheet not found' };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 7) {
      // No data yet, return empty
      return { success: true, config: config };
    }
    
    // Get rows 7 to lastRow, columns B (2) and C (3)
    const values = sheet.getRange(7, 2, lastRow - 6, 2).getValues();
    
    values.forEach(row => {
      const accountType = String(row[0] || '').trim();
      if (!accountType) return; // Skip empty rows
      
      const subAccountsRaw = String(row[1] || '').trim();
      let subAccountsArray = [];
      
      if (subAccountsRaw) {
        // Split by comma and trim each element
        subAccountsArray = subAccountsRaw.split(',').map(s => s.trim()).filter(s => s.length > 0);
      }
      
      config.accountTypes[accountType] = subAccountsArray;
    });
    
    return { success: true, config: config };
    
  } catch (e) {
    return { success: false, error: 'Failed to get system config: ' + e.message };
  }
}

