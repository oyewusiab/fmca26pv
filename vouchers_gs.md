// ==================== VOUCHER CRUD OPERATIONS ====================

/**
 * Gets vouchers from a specific year sheet
 * @param {string} token - Session token
 * @param {string} year - Year to fetch (2026, 2025, 2024, 2023, <2023)
 * @param {string} filtersJson - JSON string of filters (optional)
 * @returns {Object} List of vouchers
 */
function parseVoucherFilters_(filtersValue) {
  if (!filtersValue) return {};

  if (typeof filtersValue === 'string') {
    try {
      return JSON.parse(filtersValue);
    } catch (e) {
      return {};
    }
  }

  if (typeof filtersValue === 'object' && !Array.isArray(filtersValue)) {
    return filtersValue;
  }

  return {};
}

function normalizeGetVouchersArgs_(yearOrOptions, filtersArg, pageArg, pageSizeArg) {
  let year = '2026';
  let page = 1;
  let pageSize = 50;
  let filters = {};

  if (yearOrOptions && typeof yearOrOptions === 'object' && !Array.isArray(yearOrOptions)) {
    year = String(yearOrOptions.year || '2026');
    page = parseInt(yearOrOptions.page, 10) || 1;
    pageSize = parseInt(yearOrOptions.pageSize, 10) || 50;
    filters = parseVoucherFilters_(yearOrOptions.filters);
  } else {
    year = String(yearOrOptions || '2026');
    page = parseInt(pageArg, 10) || 1;
    pageSize = parseInt(pageSizeArg, 10) || 50;
    filters = parseVoucherFilters_(filtersArg);
  }

  page = Math.max(page, 1);
  pageSize = Math.min(Math.max(pageSize, 1), 200);

  return { year, filters, page, pageSize };
}

function getVouchers(token, yearOrOptions, filtersArg, pageArg, pageSizeArg) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };

    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.AUDIT,
      CONFIG.ROLES.DDFA,
      CONFIG.ROLES.DFA,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized' };
    }

    const query = normalizeGetVouchersArgs_(yearOrOptions, filtersArg, pageArg, pageSizeArg);
    let year = query.year;
    let filters = query.filters;
    let page = query.page;
    let pageSize = query.pageSize;

    let sheetName;
    let has2026Format = false;

    switch (year) {
      case '2026': sheetName = CONFIG.SHEETS.VOUCHERS_2026; has2026Format = true; break;
      case '2025': sheetName = CONFIG.SHEETS.VOUCHERS_2025; break;
      case '2024': sheetName = CONFIG.SHEETS.VOUCHERS_2024; break;
      case '2023': sheetName = CONFIG.SHEETS.VOUCHERS_2023; break;
      case '<2023': sheetName = CONFIG.SHEETS.VOUCHERS_BEFORE_2023; break;
      default: sheetName = CONFIG.SHEETS.VOUCHERS_2026; has2026Format = true; year = '2026';
    }

    const sheet = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    const lastCol = has2026Format ? 18 : 17;

    if (lastRow <= 1) {
      return {
        success: true,
        vouchers: [],
        totalCount: 0,
        page: 1,
        pageSize,
        totalPages: 0,
        year
      };
    }

    const term = String(filters.searchTerm || '').trim().toLowerCase();
    const cleanedTerm = term.replace(/[₦,\s]/g, '');
    const termNumeric = cleanedTerm ? parseFloat(cleanedTerm) : NaN;

    const minA = filters.amountMin !== undefined && filters.amountMin !== ''
      ? parseFloat(filters.amountMin)
      : null;

    const maxA = filters.amountMax !== undefined && filters.amountMax !== ''
      ? parseFloat(filters.amountMax)
      : null;

    const releaseFilter = String(filters.release || 'All').trim();

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const filtered = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      if (!row[2] && !row[13] && !row[14]) continue;

      const rowIndex = i + 2;
      const voucher = rowToVoucher(row, rowIndex, has2026Format);

      let include = true;

      if (filters.status && filters.status !== 'All') {
        if (String(voucher.status || '').trim() !== filters.status) include = false;
      }

      if (include && filters.category && filters.category !== 'All') {
        if (String(voucher.categories || '').trim() !== filters.category) include = false;
      }

      if (include && releaseFilter !== 'All') {
        const hasCN = voucher.controlNumber && String(voucher.controlNumber).trim() !== '';
        if (releaseFilter === 'Released' && !hasCN) include = false;
        if (releaseFilter === 'Not Released' && hasCN) include = false;
      }

      if (include && minA !== null && !isNaN(minA)) {
        if (Number(voucher.grossAmount || 0) < minA) include = false;
      }

      if (include && maxA !== null && !isNaN(maxA)) {
        if (Number(voucher.grossAmount || 0) > maxA) include = false;
      }

      if (include && term) {
        const searchFields = [
          voucher.payee,
          voucher.accountOrMail,
          voucher.particular,
          voucher.controlNumber,
          voucher.oldVoucherNumber,
          String(voucher.grossAmount || '')
        ].join(' ').toLowerCase();

        const numericMatch = !isNaN(termNumeric) &&
          Math.abs(Number(voucher.grossAmount || 0) - termNumeric) < 0.01;

        if (!searchFields.includes(term) && !numericMatch) include = false;
      }

      if (include) filtered.push(voucher);
    }

    const totalCount = filtered.length;
    const totalPages = totalCount > 0 ? Math.ceil(totalCount / pageSize) : 0;

    const start = (page - 1) * pageSize;
    const end = Math.min(start + pageSize, totalCount);
    const vouchers = filtered.slice(start, end);

    return {
      success: true,
      vouchers,
      totalCount,
      page,
      pageSize,
      totalPages,
      year
    };

  } catch (error) {
    return { success: false, error: 'Failed to get vouchers: ' + error.message };
  }
}

/**
 * Gets a single voucher by row index
 * @param {string} token - Session token
 * @param {number} rowIndex - Row index
 * @param {string} year - Year sheet
 * @returns {Object} Voucher data
 */
function getVoucherByRow(token, rowIndex, year) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Determine sheet
    let sheetName = CONFIG.SHEETS.VOUCHERS_2026;
    let has2026Format = true;
    
    if (year && year !== '2026') {
      has2026Format = false;
      switch(year) {
        case '2025': sheetName = CONFIG.SHEETS.VOUCHERS_2025; break;
        case '2024': sheetName = CONFIG.SHEETS.VOUCHERS_2024; break;
        case '2023': sheetName = CONFIG.SHEETS.VOUCHERS_2023; break;
        case '<2023': sheetName = CONFIG.SHEETS.VOUCHERS_BEFORE_2023; break;
      }
    }
    
    const sheet = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    
    if (rowIndex < 2 || rowIndex > lastRow) {
      return { success: false, error: 'Voucher not found' };
    }
    
    const numCols = has2026Format ? 18 : 17;
    const row = sheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];
    const voucher = rowToVoucher(row, rowIndex, has2026Format);
    
    return { success: true, voucher: voucher };
    
  } catch (error) {
    return { success: false, error: 'Failed to get voucher: ' + error.message };
  }
}

/**
 * Batch update status for multiple vouchers by control number
 * @param {string} token - Session token
 * @param {string} controlNumber - Control number to match
 * @param {string} status - New status
 * @param {string} pmtMonth - Payment month
 * @returns {Object} Result with count of updated vouchers
 */
function batchUpdateStatus(token, controlNumber, status, pmtMonth) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    logAudit(session, 'BATCH_UPDATE_STATUS',
         'Batch updated ' + updateCount + ' voucher(s) with control number ' + controlNumber + ' to status ' + status,
         CONFIG.SHEETS.VOUCHERS_2026, null);

    // Check permission
    if (!hasPermission(session.role, [
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized: Only CPO can batch update' };
    }
    
    if (!controlNumber) {
      return { success: false, error: 'Control number is required' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const data = sheet.getDataRange().getValues();
    
    let updateCount = 0;
    const targetCN = controlNumber.trim().toUpperCase();
    
    for (let i = 1; i < data.length; i++) {
      const rowCN = String(data[i][CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER - 1]).trim().toUpperCase();
      
      if (rowCN === targetCN) {
        const rowIndex = i + 1;
        sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS).setValue(status);
        
        if (pmtMonth) {
          sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.PMT_MONTH).setValue(pmtMonth);
        }
        
        updateCount++;
      }
    }
    
    if (updateCount === 0) {
      return { success: false, error: 'No vouchers found with control number: ' + controlNumber };
    }
    
        clearAllVoucherCaches();
    return { 
      success: true, 
      message: `Updated ${updateCount} voucher(s) with control number: ${controlNumber}`,
      count: updateCount
    };

    return { 
      success: true, 
      message: `Updated ${updateCount} voucher(s) with control number: ${controlNumber}`,
      count: updateCount
    };
    
  } catch (error) {
    return { success: false, error: 'Batch update failed: ' + error.message };
  }
  
}

/**
 * Release selected vouchers to a unit:
 * - Auto-generates a control number CN-<UNIT>-<N>
 * - Assigns same control number to all selected rowIndexes
 * - Can be called only by Payable Staff / Head / Admin
 */
function releaseSelectedVouchers(token, rowIndexes, targetUnit) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };

    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Only Payable Unit or Admin can release vouchers.' };
    }

    if (!Array.isArray(rowIndexes) || rowIndexes.length === 0) {
      return { success: false, error: 'No vouchers selected.' };
    }

    if (!targetUnit || !String(targetUnit).trim()) {
      return { success: false, error: 'Target unit is required.' };
    }

    const cn = generateControlNumber(targetUnit);

    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const header = getHeaderMap_(sheet);
    const colReleasedAt = header['RELEASED AT'] || null;

    const lastRow = sheet.getLastRow();
    const released = [];

    rowIndexes.forEach(r => {
      const rowIndex = parseInt(r, 10);
      if (rowIndex >= 2 && rowIndex <= lastRow) {
        setControlNoAndReleasedAt_(sheet, rowIndex, cn, colReleasedAt);

        const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
        released.push({
          voucherNumber: row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1],
          payee: row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1],
          amount: row[CONFIG.VOUCHER_COLUMNS.GROSS_AMOUNT - 1]
        });

        if (typeof logAudit === 'function') {
          logAudit(session, 'RELEASE_VOUCHER', 'Released voucher to ' + targetUnit, CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
        }
      }
    });

    if (released.length === 0) {
      return { success: false, error: 'No valid vouchers found for release.' };
    }

    clearAllVoucherCaches();

    return {
      success: true,
      message: 'Released ' + released.length + ' voucher(s) to ' + targetUnit + ' with Control Number: ' + cn,
      count: released.length,
      controlNumber: cn,
      targetUnit: targetUnit,
      vouchers: released
    };

  } catch (error) {
    return { success: false, error: 'Release failed: ' + error.message };
  }
}

/**
 * Gets the next control number for a target unit
 * Format: CN-<UNIT>-<BATCH_NUMBER>
 * Example: CN-DFA-003
 * 
 * @param {string} token - Session token
 * @param {string} targetUnit - Target unit code (CPO, DFA, DDFA, Audit, etc.)
 * @returns {Object} { success: true, controlNumber: "CN-DFA-003" } or error
 */
function getNextControlNumber(token, targetUnit) {
    try {
        // Validate session
        const session = getSession(token);
        if (!session) {
            return { success: false, error: 'Session expired' };
        }

        // Check permissions
        if (!hasPermission(session.role, [
            CONFIG.ROLES.PAYABLE_STAFF,
            CONFIG.ROLES.PAYABLE_HEAD,
            CONFIG.ROLES.CPO,
            CONFIG.ROLES.ADMIN
        ])) {
            return { success: false, error: 'Unauthorized' };
        }

        // Validate target unit
        if (!targetUnit || !String(targetUnit).trim()) {
            return { success: false, error: 'Target unit is required' };
        }

        // Generate the control number
        const controlNumber = generateControlNumber(targetUnit);
        
        return { 
            success: true, 
            controlNumber: controlNumber 
        };

    } catch (error) {
        return { 
            success: false, 
            error: 'Failed to generate control number: ' + error.message 
        };
    }
}

/**
 * Generates a new control number for 2026 in the format:
 * CN-<UNIT>-<N>
 * where <UNIT> is the target unit code (CPO, Audit, DDFA, DFA)
 * and <N> is the next batch number for that unit in 2026.
 * 
 * @param {string} targetUnit - Target unit code
 * @returns {string} Control number (e.g., "CN-DFA-003")
 */
function generateControlNumber(targetUnit) {
    const unitCode = String(targetUnit || '').toUpperCase().trim();
    
    if (!unitCode) {
        throw new Error('Target unit is required to generate control number.');
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    // If no data rows, start with 001
    if (lastRow <= 1) {
        return 'CN-' + unitCode + '-001';
    }
    
    const cnCol = CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER;
    const cnValues = sheet.getRange(2, cnCol, lastRow - 1, 1).getValues();
    
    const prefix = 'CN-' + unitCode + '-';
    let maxNumber = 0;
    
    for (let i = 0; i < cnValues.length; i++) {
        const val = String(cnValues[i][0] || '').trim().toUpperCase();
        
        if (val.startsWith(prefix)) {
            const numPart = val.substring(prefix.length);
            const n = parseInt(numPart, 10);
            
            if (!isNaN(n) && n > maxNumber) {
                maxNumber = n;
            }
        }
    }
    
    // Format with leading zeros (3 digits)
    const nextNumber = String(maxNumber + 1).padStart(3, '0');
    
    return prefix + nextNumber;
}

/**
 * Converts a spreadsheet row to a voucher object.
 * @param {any[]} row
 * @param {number} rowIndex
 * @param {boolean} has2026Format
 * @returns {object}
 */
function rowToVoucher(row, rowIndex, has2026Format = true) {
  const cols = CONFIG.VOUCHER_COLUMNS;
  const voucher = {
    rowIndex: rowIndex,
    status: row[cols.STATUS - 1],
    pmtMonth: row[cols.PMT_MONTH - 1],
    payee: row[cols.PAYEE - 1],
    accountOrMail: row[cols.ACCOUNT_OR_MAIL - 1],
    particular: row[cols.PARTICULAR - 1],
    contractSum: parseAmount(row[cols.CONTRACT_SUM - 1]),
    grossAmount: parseAmount(row[cols.GROSS_AMOUNT - 1]),
    net: parseAmount(row[cols.NET - 1]),
    vat: parseAmount(row[cols.VAT - 1]),
    wht: parseAmount(row[cols.WHT - 1]),
    stampDuty: parseAmount(row[cols.STAMP_DUTY - 1]),
    categories: row[cols.CATEGORIES - 1],
    totalGross: parseAmount(row[cols.TOTAL_GROSS - 1]),
    controlNumber: row[cols.CONTROL_NUMBER - 1],
    oldVoucherNumber: row[cols.OLD_VOUCHER_NUMBER - 1],
    oldVoucherAvailable: row[cols.OLD_VOUCHER_AVAILABLE - 1],
    date: row[cols.DATE - 1] ? Utilities.formatDate(new Date(row[cols.DATE - 1]), "GMT", "yyyy-MM-dd") : '',
    accountType: row[cols.ACCOUNT_TYPE - 1],
    createdAt: row[cols.CREATED_AT - 1] ? new Date(row[cols.CREATED_AT - 1]).toISOString() : '',
    releasedAt: row[cols.RELEASED_AT - 1] ? new Date(row[cols.RELEASED_AT - 1]).toISOString() : '',
    attachmentUrl: has2026Format ? row[cols.ATTACHMENT_URL - 1] : ''
  };
  return voucher;
}

/**
 * Converts voucher object to row array for 2026 sheet
 * @param {Object} voucher - Voucher object
 * @returns {Array} Row data
 */
function voucherToRow(voucher, targetColumnCount) {
  const cols = CONFIG.VOUCHER_COLUMNS;
  const maxConfiguredCol = Math.max.apply(null, Object.keys(cols).map(k => cols[k]));
  const width = Math.max(1, Math.min(parseInt(targetColumnCount || maxConfiguredCol, 10), maxConfiguredCol));
  const row = new Array(width).fill('');

  function setCol(colIndex, value) {
    if (!colIndex || colIndex < 1 || colIndex > width) return;
    row[colIndex - 1] = value;
  }

  setCol(cols.STATUS, voucher.status || 'Unpaid');
  setCol(cols.PMT_MONTH, voucher.pmtMonth || '');
  setCol(cols.PAYEE, voucher.payee || '');
  setCol(cols.ACCOUNT_OR_MAIL, voucher.accountOrMail || '');
  setCol(cols.PARTICULAR, voucher.particular || '');
  setCol(cols.CONTRACT_SUM, voucher.contractSum || 0);
  setCol(cols.GROSS_AMOUNT, voucher.grossAmount || 0);
  setCol(cols.NET, voucher.net || 0);
  setCol(cols.VAT, voucher.vat || 0);
  setCol(cols.WHT, voucher.wht || 0);
  setCol(cols.STAMP_DUTY, voucher.stampDuty || 0);
  setCol(cols.CATEGORIES, voucher.categories || '');
  setCol(cols.TOTAL_GROSS, voucher.totalGross || voucher.grossAmount || 0);
  setCol(cols.CONTROL_NUMBER, voucher.controlNumber || '');
  setCol(cols.OLD_VOUCHER_NUMBER, voucher.oldVoucherNumber || '');
  setCol(cols.OLD_VOUCHER_AVAILABLE, voucher.oldVoucherAvailable || '');
  setCol(cols.DATE, voucher.date ? new Date(voucher.date) : new Date());
  setCol(cols.ACCOUNT_TYPE, voucher.accountType || '');
  setCol(cols.CREATED_AT, voucher.createdAt ? new Date(voucher.createdAt) : new Date());
  setCol(cols.RELEASED_AT, voucher.releasedAt ? new Date(voucher.releasedAt) : '');
  setCol(cols.ATTACHMENT_URL, voucher.attachmentUrl || '');

  return row;
}

/**
 * Creates a new voucher in 2026 VOUCHERS sheet
 * @param {string} token - Session token
 * @param {Object} voucher - Voucher data
 * @returns {Object} Result with new row index
 */
function createVoucher(token, voucher) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Check permission
    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized: Cannot create vouchers' };
    }
    
    // Validate required fields
    if (!voucher.payee || !voucher.payee.trim()) {
      return { success: false, error: 'Payee name is required' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    
    // Check for duplicate OLD VOUCHER NUMBER in 2026
    if (voucher.oldVoucherNumber && voucher.oldVoucherNumber.trim()) {
      const data = sheet.getDataRange().getValues();
      const newOldVN = voucher.oldVoucherNumber.trim().toUpperCase();
      
      for (let i = 1; i < data.length; i++) {
        const existingOldVN = String(data[i][CONFIG.VOUCHER_COLUMNS.OLD_VOUCHER_NUMBER - 1]).trim().toUpperCase();
        if (existingOldVN === newOldVN) {
          return { 
            success: false, 
            error: 'Duplicate: A voucher with this Old Voucher Number already exists in 2026 (Row ' + (i + 1) + ')'
          };
        }
      }
    }
    
    // Set defaults
    voucher.status = voucher.status || 'Unpaid';
    voucher.createdAt = getNigerianTimestamp();
    
    // Calculate TOTAL GROSS if not provided
    if (!voucher.totalGross) {
      voucher.totalGross = parseAmount(voucher.grossAmount);
    }
    
    // Recalculate NET on server: NET = GROSS - (VAT + WHT + STAMP DUTY)
    const gross = parseAmount(voucher.grossAmount);
    const vat = parseAmount(voucher.vat);
    const wht = parseAmount(voucher.wht);
    const stamp = parseAmount(voucher.stampDuty);
    voucher.net = gross - (vat + wht + stamp);

    // Create row array
    const rowData = voucherToRow(voucher, sheet.getLastColumn());
    
    // Append to sheet
    sheet.appendRow(rowData);
    
    const newRowIndex = sheet.getLastRow();
    
    // Clear cache after data change - MUST BE BEFORE RETURN
    clearAllVoucherCaches();
    
    // Log audit
    logAudit(session, 'CREATE_VOUCHER', 'Created new voucher', CONFIG.SHEETS.VOUCHERS_2026, newRowIndex);
    
    return { 
      success: true, 
      message: 'Voucher created successfully',
      rowIndex: newRowIndex,
      voucher: rowToVoucher(rowData, newRowIndex, true)
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to create voucher: ' + error.message };
  }
}

/**
 * Updates voucher status (for CPO and Payable units)
 * @param {string} token - Session token
 * @param {number} rowIndex - Row index
 * @param {string} status - New status (Paid/Unpaid/Cancelled)
 * @param {string} pmtMonth - Payment month (optional, for CPO)
 * @returns {Object} Result
 */
function updateVoucherStatus(token, rowIndex, status, pmtMonth) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Check permission
    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized: Cannot update status' };
    }
    
    // Validate status
    const validStatuses = ['Paid', 'Unpaid', 'Cancelled'];
    if (!validStatuses.includes(status)) {
      return { success: false, error: 'Invalid status. Must be: ' + validStatuses.join(', ') };
    }
    
    // Validate row index
    if (!rowIndex || rowIndex < 2) {
      return { success: false, error: 'Invalid voucher reference' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    if (rowIndex > lastRow) {
      return { success: false, error: 'Voucher not found' };
    }
    
    // Update status
    sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS).setValue(status);
    
    // Update payment month if provided (CPO function)
    if (pmtMonth !== undefined && pmtMonth !== null && pmtMonth !== '') {
      if (!hasPermission(session.role, [CONFIG.ROLES.CPO, CONFIG.ROLES.ADMIN])) {
        return { success: false, error: 'Only CPO can set payment month' };
      }
      sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.PMT_MONTH).setValue(pmtMonth);
    }
    
    // Clear cache
    clearAllVoucherCaches();
    
    // Log audit
    logAudit(session, 'UPDATE_STATUS', 'Changed status to ' + status, CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
    
    return { 
      success: true, 
      message: 'Status updated to: ' + status
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to update status: ' + error.message };
  }
}

/**
 * Updates an existing voucher in 2026 VOUCHERS sheet
 * @param {string} token - Session token
 * @param {number} rowIndex - Row index to update
 * @param {Object} voucher - Updated voucher data
 * @returns {Object} Result
 */
function updateVoucher(token, rowIndex, voucher) {
  try {
    // FIRST: Get session
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Check basic permission
    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized: Cannot update vouchers' };
    }
    
    // Validate row index
    if (!rowIndex || rowIndex < 2) {
      return { success: false, error: 'Invalid voucher reference' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    if (rowIndex > lastRow) {
      return { success: false, error: 'Voucher not found' };
    }
    
    // THEN: Get current data
    const sheetCols = sheet.getLastColumn();
    const currentRow = sheet.getRange(rowIndex, 1, 1, sheetCols).getValues()[0];
    const currentVoucher = rowToVoucher(currentRow, rowIndex, true);
    
    // NOW we can check restrictions based on current voucher state
    
    // 1) Block edit of PAID vouchers by non-CPO/Admin
    if (String(currentVoucher.status || '').toLowerCase() === 'paid') {
      if (session.role !== CONFIG.ROLES.CPO && session.role !== CONFIG.ROLES.ADMIN) {
        return { success: false, error: 'Paid vouchers can only be edited by CPO or Admin.' };
      }
    }
    
    // 2) Block edit of RELEASED vouchers (has control number) by non-CPO/Admin
    if (currentVoucher.controlNumber && currentVoucher.controlNumber.toString().trim() !== '') {
      if (session.role !== CONFIG.ROLES.CPO && session.role !== CONFIG.ROLES.ADMIN) {
        return { success: false, error: 'Released vouchers can only be edited by CPO or Admin.' };
      }
    }
    
    // 3) Payable Staff can only edit if status is Unpaid (draft)
    if (session.role === CONFIG.ROLES.PAYABLE_STAFF) {
      if (currentVoucher.status !== 'Unpaid') {
        return { success: false, error: 'Cannot edit: Voucher is no longer a draft' };
      }
    }
    
    // Check for duplicate OLD VOUCHER NUMBER (excluding current row)
    if (voucher.oldVoucherNumber && voucher.oldVoucherNumber.trim()) {
      const data = sheet.getDataRange().getValues();
      const newOldVN = voucher.oldVoucherNumber.trim().toUpperCase();
      
      for (let i = 1; i < data.length; i++) {
        if (i + 1 === rowIndex) continue; // Skip current row
        
        const existingOldVN = String(data[i][CONFIG.VOUCHER_COLUMNS.OLD_VOUCHER_NUMBER - 1]).trim().toUpperCase();
        if (existingOldVN === newOldVN) {
          return { 
            success: false, 
            error: 'Duplicate: This Old Voucher Number already exists in Row ' + (i + 1)
          };
        }
      }
    }
    
    // Preserve CREATED_AT from original
    voucher.createdAt = currentVoucher.createdAt;
    
    // Preserve status if not explicitly changing
    if (!voucher.status) {
      voucher.status = currentVoucher.status;
    }
    
    // Preserve control number if not provided
    if (!voucher.controlNumber) {
      voucher.controlNumber = currentVoucher.controlNumber;
    }
    
    // Calculate TOTAL GROSS
    if (!voucher.totalGross) {
      voucher.totalGross = parseAmount(voucher.grossAmount);
    }
    // In updateVoucher() function, after getting currentVoucher:

    // Check if voucher is released (has control number)
    if (currentVoucher.controlNumber && currentVoucher.controlNumber.toString().trim() !== '') {
        // Released vouchers can only be edited by CPO or Admin
        if (session.role !== CONFIG.ROLES.CPO && session.role !== CONFIG.ROLES.ADMIN) {
            return { 
                success: false, 
                error: 'This voucher has been released. Only CPO or Admin can edit released vouchers.' 
            };
        }
    }

    // Check if voucher is paid
    if (String(currentVoucher.status || '').toLowerCase() === 'paid') {
        if (session.role !== CONFIG.ROLES.CPO && session.role !== CONFIG.ROLES.ADMIN) {
            return { 
                success: false, 
                error: 'Paid vouchers can only be edited by CPO or Admin.' 
            };
        }
    }
    // Recalculate NET on server: NET = GROSS - (VAT + WHT + STAMP DUTY)
    const gross = parseAmount(voucher.grossAmount);
    const vat = parseAmount(voucher.vat);
    const wht = parseAmount(voucher.wht);
    const stamp = parseAmount(voucher.stampDuty);
    voucher.net = gross - (vat + wht + stamp);

    // Create updated row
    const rowData = voucherToRow(voucher, sheetCols);
    
    // Update the row
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    
    // Clear cache after data change - MUST BE BEFORE RETURN
    clearAllVoucherCaches();
    
    // Log audit
    logAudit(session, 'UPDATE_VOUCHER', 'Updated voucher', CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
    
    return { 
      success: true, 
      message: 'Voucher updated successfully',
      voucher: rowToVoucher(rowData, rowIndex, true)
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to update voucher: ' + error.message };
  }
}

// ==================== DELETION WORKFLOW ====================

/**
 * Request voucher deletion
 * Stores original status for later restoration
 */
function requestVoucherDelete(token, rowIndex, reason, previousStatus) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Check permission
    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.AUDIT,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized to request deletion' };
    }
    
    // Validate reason
    if (!reason || !reason.trim()) {
      return { success: false, error: 'Reason for deletion is required' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    rowIndex = parseInt(rowIndex, 10);
    
    if (rowIndex < 2 || rowIndex > lastRow) {
      return { success: false, error: 'Voucher not found' };
    }
    
    // Get the full row
    const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
    const currentStatus = String(row[CONFIG.VOUCHER_COLUMNS.STATUS - 1] || '').trim();
    const controlNumber = String(row[CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER - 1] || '').trim();
    const payee = row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1];
    const voucherNum = row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1];
    const grossAmount = parseAmount(row[CONFIG.VOUCHER_COLUMNS.GROSS_AMOUNT - 1]);
    
    // Already pending?
    if (currentStatus === 'Pending Deletion') {
      return { success: false, error: 'Deletion already requested for this voucher.' };
    }
    
    // Store original status in a note on the STATUS cell for later restoration
    const statusCell = sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS);
    statusCell.setNote('ORIGINAL_STATUS:' + currentStatus + '|REQUESTED_BY:' + session.email + '|REASON:' + reason);
    
    // Set status to Pending Deletion
    statusCell.setValue('Pending Deletion');
    
    // Determine approver
    let requiredApprover = 'Payable Unit Head or Admin';
    if (controlNumber !== '') {
      requiredApprover = 'CPO or Admin';
    }
    
    // Log audit
    logAudit(session, 'REQUEST_DELETE', 
    'Requested deletion for ' + voucherNum + ' (' + payee + '). Previous status: ' + currentStatus + '. Reason: ' + reason, 
    CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
    
    // Get approver emails
    const usersSheet = getSheet(CONFIG.SHEETS.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    const approverEmails = [];
    
    for (let i = 1; i < usersData.length; i++) {
      const role = usersData[i][CONFIG.USER_COLUMNS.ROLE - 1];
      const email = usersData[i][CONFIG.USER_COLUMNS.EMAIL - 1];
      const active = usersData[i][CONFIG.USER_COLUMNS.ACTIVE - 1];
      
      if (!active || !email) continue;
      
      if (controlNumber !== '') {
        if (role === CONFIG.ROLES.CPO || role === CONFIG.ROLES.ADMIN) {
          approverEmails.push(email);
        }
      } else {
        if (role === CONFIG.ROLES.PAYABLE_HEAD || role === CONFIG.ROLES.ADMIN) {
          approverEmails.push(email);
        }
      }
    }
    
    // Format amount
    const formattedAmount = '₦' + Number(grossAmount).toLocaleString('en-NG', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
    
    // Create notifications
    if (approverEmails.length > 0) {
      createNotifications(
        approverEmails,
        '🗑️ Deletion Approval Required',
        session.name + ' requested deletion of voucher ' + voucherNum + ' (' + payee + ') - ' + formattedAmount + '. REASON: ' + reason,
        'vouchers.html?filter=pending',
        'warning'
      );
    }
    
    // Clear cache
    clearAllVoucherCaches();
    
    return { 
      success: true, 
      message: 'Deletion request submitted. Awaiting approval from ' + requiredApprover + '.'
    };
    
  } catch (error) {
    console.log('requestVoucherDelete error: ' + error.message);
    return { success: false, error: 'Delete request failed: ' + error.message };
  }
}

/**
 * Approve and execute voucher deletion
 */
function approveVoucherDelete(token, rowIndex) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    const role = session.role;
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    rowIndex = parseInt(rowIndex, 10);
    
    if (rowIndex < 2 || rowIndex > lastRow) {
      return { success: false, error: 'Voucher not found' };
    }
    
    // Get voucher data
    const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
    const currentStatus = String(row[CONFIG.VOUCHER_COLUMNS.STATUS - 1] || '').trim();
    const controlNumber = String(row[CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER - 1] || '').trim();
    const voucherNum = row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1];
    const payee = row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1];
    
    if (currentStatus !== 'Pending Deletion') {
      return { success: false, error: 'Voucher is not pending deletion' };
    }
    
    // Check permissions based on voucher state
    if (controlNumber !== '') {
      // Released voucher - only CPO or Admin
      if (role !== CONFIG.ROLES.CPO && role !== CONFIG.ROLES.ADMIN) {
        return { success: false, error: 'Only CPO or Admin can approve deletion of released vouchers.' };
      }
    } else {
      // Normal voucher
      if (role !== CONFIG.ROLES.PAYABLE_HEAD && role !== CONFIG.ROLES.ADMIN) {
        return { success: false, error: 'Only Payable Head or Admin can approve deletion.' };
      }
    }
    
    // Get the note with request info before deleting
    const statusCell = sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS);
    const note = statusCell.getNote() || '';
    
    // Delete the row
    sheet.deleteRow(rowIndex);
    
    // Clear cache
    clearAllVoucherCaches();
    
    // Log audit
    logAudit(session, 'APPROVE_DELETE', 
             'Approved and deleted voucher ' + voucherNum + ' (' + payee + ')', 
             CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
    
    return { success: true, message: 'Voucher deleted successfully.' };
    
  } catch (error) {
    console.log('approveVoucherDelete error: ' + error.message);
    return { success: false, error: 'Delete approval failed: ' + error.message };
  }
}

/**
 * Reject voucher deletion request - Restores to ORIGINAL status
 */
function rejectVoucherDelete(token, rowIndex, reason) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    const role = session.role;
    
    // Only Payable Head, CPO, and Admin can reject
    if (role !== CONFIG.ROLES.PAYABLE_HEAD && 
        role !== CONFIG.ROLES.CPO && 
        role !== CONFIG.ROLES.ADMIN) {
      return { success: false, error: 'Unauthorized to reject deletion requests' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    rowIndex = parseInt(rowIndex, 10);
    
    if (rowIndex < 2 || rowIndex > lastRow) {
      return { success: false, error: 'Voucher not found' };
    }
    
    // Get voucher data
    const statusCell = sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS);
    const currentStatus = String(statusCell.getValue() || '').trim();
    
    if (currentStatus !== 'Pending Deletion') {
      return { success: false, error: 'Voucher is not pending deletion' };
    }
    
    // Get original status from note
    const note = statusCell.getNote() || '';
    let originalStatus = 'Unpaid'; // Default fallback
    let requesterEmail = '';
    
    if (note) {
      // Parse the note: ORIGINAL_STATUS:Paid|REQUESTED_BY:email@example.com|REASON:some reason
      const parts = note.split('|');
      for (const part of parts) {
        if (part.startsWith('ORIGINAL_STATUS:')) {
          originalStatus = part.replace('ORIGINAL_STATUS:', '').trim();
        }
        if (part.startsWith('REQUESTED_BY:')) {
          requesterEmail = part.replace('REQUESTED_BY:', '').trim();
        }
      }
    }
    
    // Get voucher details for notification
    const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
    const voucherNum = row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1] || '';
    const payee = row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1] || '';
    
    // Restore to original status
    statusCell.setValue(originalStatus);
    statusCell.clearNote(); // Clear the stored info
    
    // Clear cache
    clearAllVoucherCaches();
    
    // Log audit
    logAudit(session, 'REJECT_DELETE', 
             'Rejected deletion of ' + voucherNum + '. Restored to: ' + originalStatus + '. Reason: ' + (reason || 'Not specified'), 
             CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
    
    // Notify the requester
    if (requesterEmail) {
      createNotifications(
        [requesterEmail],
        '❌ Deletion Request Rejected',
        'Your deletion request for voucher ' + voucherNum + ' (' + payee + ') was rejected by ' + session.name + '. Reason: ' + (reason || 'Not specified') + '. Status restored to: ' + originalStatus,
        'vouchers.html',
        'warning'
      );
    }
    
    return { 
      success: true, 
      message: 'Deletion request rejected. Voucher restored to "' + originalStatus + '" status.' 
    };
    
  } catch (error) {
    console.log('rejectVoucherDelete error: ' + error.message);
    return { success: false, error: 'Failed to reject deletion: ' + error.message };
  }
}

/**
 * Cancel own deletion request - Restores to ORIGINAL status
 */
function cancelDeleteRequest(token, rowIndex) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Anyone who can request can also cancel
    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.AUDIT,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    rowIndex = parseInt(rowIndex, 10);
    
    if (rowIndex < 2 || rowIndex > lastRow) {
      return { success: false, error: 'Voucher not found' };
    }
    
    // Get voucher data
    const statusCell = sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS);
    const currentStatus = String(statusCell.getValue() || '').trim();
    
    if (currentStatus !== 'Pending Deletion') {
      return { success: false, error: 'Voucher is not pending deletion' };
    }
    
    // Get original status from note
    const note = statusCell.getNote() || '';
    let originalStatus = 'Unpaid'; // Default fallback
    
    if (note) {
      const parts = note.split('|');
      for (const part of parts) {
        if (part.startsWith('ORIGINAL_STATUS:')) {
          originalStatus = part.replace('ORIGINAL_STATUS:', '').trim();
          break;
        }
      }
    }
    
    // Get voucher details
    const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
    const voucherNum = row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1] || '';
    const payee = row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1] || '';
    
    // Restore to original status
    statusCell.setValue(originalStatus);
    statusCell.clearNote();
    
    // Clear cache
    clearAllVoucherCaches();
    
    // Log audit
    logAudit(session, 'CANCEL_DELETE_REQUEST', 
             'Cancelled deletion request for ' + voucherNum + '. Restored to: ' + originalStatus, 
             CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
    
    return { 
      success: true, 
      message: 'Deletion request cancelled. Voucher restored to "' + originalStatus + '" status.' 
    };
    
  } catch (error) {
    console.log('cancelDeleteRequest error: ' + error.message);
    return { success: false, error: 'Failed to cancel request: ' + error.message };
  }
}

function testControlNumberGeneration() {
    Logger.log('=== Testing Control Number Generation ===');
    
    // Test 1: Login
    Logger.log('\n--- Test 1: Login ---');
    const loginResult = login('admin@fmc.gov.ng', 'Admin@2026'); // Use your credentials
    
    if (!loginResult.success) {
        Logger.log('❌ Login failed: ' + loginResult.error);
        return;
    }
    Logger.log('✅ Login successful');
    Logger.log('Token: ' + loginResult.token.substring(0, 20) + '...');
    
    const token = loginResult.token;
    
    // Test 2: Generate for DFA
    Logger.log('\n--- Test 2: Generate CN for DFA ---');
    const dfaResult = getNextControlNumber(token, 'DFA');
    Logger.log('Result: ' + JSON.stringify(dfaResult));
    
    if (dfaResult.success) {
        Logger.log('✅ DFA Control Number: ' + dfaResult.controlNumber);
    } else {
        Logger.log('❌ Failed: ' + dfaResult.error);
    }
    
    // Test 3: Generate for CPO
    Logger.log('\n--- Test 3: Generate CN for CPO ---');
    const cpoResult = getNextControlNumber(token, 'CPO');
    Logger.log('Result: ' + JSON.stringify(cpoResult));
    
    if (cpoResult.success) {
        Logger.log('✅ CPO Control Number: ' + cpoResult.controlNumber);
    } else {
        Logger.log('❌ Failed: ' + cpoResult.error);
    }
    
    // Test 4: Generate for Audit
    Logger.log('\n--- Test 4: Generate CN for Audit ---');
    const auditResult = getNextControlNumber(token, 'Audit');
    Logger.log('Result: ' + JSON.stringify(auditResult));
    
    if (auditResult.success) {
        Logger.log('✅ Audit Control Number: ' + auditResult.controlNumber);
    } else {
        Logger.log('❌ Failed: ' + auditResult.error);
    }
    
    // Cleanup
    logout(token);
    Logger.log('\n=== Test Complete ===');
}
