/**
 * PAYABLE VOUCHER 2026 - Backend System
 * Federal Medical Centre, Abeokuta
 * Finance & Accounts Department
 * 
 * This script serves as the API backend for the Payment Voucher System
 */

// ==================== CACHING SYSTEM ====================

/**
 * Gets data from cache or fetches fresh if expired
 * @param {string} cacheKey - Unique cache key
 * @param {number} expirationSeconds - Cache expiration in seconds (max 21600 = 6 hours)
 * @param {Function} fetchFunction - Function to call if cache miss
 * @returns {any} Cached or fresh data
 */
function getCachedData(cacheKey, expirationSeconds, fetchFunction) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Cache corrupted, fetch fresh
    }
  }
  
  const freshData = fetchFunction();
  
  // Only cache if data is not too large (100KB limit per item)
  try {
    const jsonData = JSON.stringify(freshData);
    if (jsonData.length < 100000) {
      cache.put(cacheKey, jsonData, expirationSeconds);
    }
  } catch (e) {
    // Data too large to cache, that's okay
  }
  
  return freshData;
}

/**
 * Clears specific cache
 * @param {string} cacheKey - Cache key to clear
 */
function clearCache(cacheKey) {
  const cache = CacheService.getScriptCache();
  cache.remove(cacheKey);
}

/**
 * Clears all voucher-related caches
 */
function clearAllVoucherCaches() {
  const cache = CacheService.getScriptCache();
  cache.removeAll([
    'vouchers_2026_summary',
    'vouchers_2026_count',
    'categories_list',
    'dashboard_stats'
  ]);
}

// ==================== MAIN ENTRY POINTS ====================

/**
 * Handles GET requests
 * @param {Object} e - Event object with parameters
 * @returns {TextOutput} JSON response
 */
function doGet(e) {
    try {
        const action = (e.parameter && e.parameter.action) || '';
        const params = e.parameter || {};

        let result;

        switch (action) {
            // ---- AUTH ----
            case 'login':
                result = login(params.email, params.password);
                break;
            case 'validateSession':
                result = validateSession(params.token);
                break;
            case 'logout':
                result = logout(params.token);
                break;

            // ---- VOUCHERS (GET operations) ----
            case 'getVouchers':
                result = getVouchers(
                    params.token,
                    params.year || '2026',
                    params.filters || null,
                    parseInt(params.page, 10) || 1,
                    parseInt(params.pageSize, 10) || 50
                );
                break;
            case 'getVoucherByRow':
                result = getVoucherByRow(params.token, parseInt(params.rowIndex, 10), params.year);
                break;
            case 'getNextControlNumber':
                result = getNextControlNumber(params.token, params.targetUnit);
                break;
            case 'lookupVoucher':
                result = lookupVoucher(params.token, params.voucherNumber);
                break;

            // ---- DASHBOARD / REPORTS ----
            case 'getDashboardStats':
                result = getDashboardStats(params.token);
                break;
            case 'getSummary':
                result = getSummary(params.token, params.year);
                break;
            case 'getAllYearsSummary':
                result = getAllYearsSummary(params.token);
                break;
            case 'getDebtProfile':
                result = getDebtProfile(params.token);
                break;
            case 'getDebtProfileRequestStatus':
                result = getDebtProfileRequestStatus(params.token);
                break;
            case 'getDebtProfileFullData':
                result = getDebtProfileFullData(params.token, params.requestId);
                break;
            case 'generateDebtProfilePDF':
                result = generateDebtProfilePDF(params.token, params.requestId);
                break;
            case 'generateDebtProfileExcel':
                result = generateDebtProfileExcel(params.token, params.requestId);
                break;
            case 'getQuickStats':
                result = getQuickStats(params.token);
                break;
            case 'getUsers':
                result = getUsers(params.token);
                break;
            case 'getRolePermissions':
                result = getRolePermissions(params.token);
                break;

            // ---- NOTIFICATIONS ----
            case 'getNotifications':
                result = getNotifications(params.token, params.onlyUnread === 'true');
                break;

            case 'getUsers':
                result = getUsers(params.token);
                break;
            case 'getRolePermissions':
                result = getRolePermissions(params.token);
                break;

            // ---- NOTIFICATIONS ----
            case 'getNotifications':
                result = getNotifications(params.token, params.onlyUnread === 'true');
                break;

            // ---- AUDIT ----
            case 'getAuditTrail':
                result = getAuditTrail(params.token, parseInt(params.limit) || 50, parseInt(params.offset) || 0);
                break;

            // ---- ACTION ITEMS ----
            case 'getActionItems':
                result = getActionItems(params.token, params);
                break;
            case 'getActionItemCount':
                result = getActionItemCount(params.token, params);
                break;
            case 'getActionItemSettings':
                result = getActionItemSettings(params.token);
                break;

            // ---- CATEGORIES ----
            case 'getCategories':
                result = getCategories(params.token);
                break;
            case 'getSystemConfig':
                result = getSystemConfig(params.token);
                break;

            // ---- DELETIONS ----
            case 'getPendingDeletions':
                result = getPendingDeletions(params.token);
                break;

            // ---- PROFILE ----
            case 'getMyProfile':
                result = getMyProfile(params.token);
                break;

            // ---- ANNOUNCEMENTS ----
            case 'getActiveAnnouncements':
                result = getActiveAnnouncements(params.token, params.location);
                break;

            // ---- TAX (GET operations) ----
            case 'getTaxSummary':
                result = getTaxSummary(params.token, params.year);
                break;
            case 'getTaxByCategory':
                result = getTaxByCategory(params.token, params.year);
                break;
            case 'getTaxByMonth':
                result = getTaxByMonth(params.token, params.year);
                break;
            case 'getTaxByPayee':
                result = getTaxByPayee(params.token, params.year, params.filters);
                break;
            case 'getTaxPayments':
                result = getTaxPayments(params.token, params.year);
                break;
            case 'getTaxSchedule':
                result = getTaxSchedule(params.token, params.year);
                break;
            case 'getTaxCompliance':
                result = getTaxCompliance(params.token, params.year);
                break;

            default:
                result = { success: false, error: 'Unknown GET action: ' + action };
        }

        return ContentService.createTextOutput(JSON.stringify(result))
            .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
            success: false,
            error: 'doGet error: ' + err.message
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * Handles POST requests
 * @param {Object} e - Event object with post data
 * @returns {TextOutput} JSON response
 */
/**
 * Handles POST requests
 * @param {Object} e - Event object with post data
 * @returns {TextOutput} JSON response
 */
/**
 * Handles POST requests
 * @param {Object} e - Event object with post data
 * @returns {TextOutput} JSON response
 */
function doPost(e) {
    try {
        let payload;

        if (e.postData) {
            try {
                payload = JSON.parse(e.postData.contents);
            } catch (_) {
                payload = {};
            }
        } else {
            payload = {};
        }

        const action = payload.action || '';
        const token = payload.token || '';

        let result;

        switch (action) {
            // ---- AUTH ----
            case 'login':
                result = login(payload.email, payload.password);
                break;
            case 'changePassword':
                result = changePassword(token, payload.oldPassword, payload.newPassword);
                break;

            // ---- VOUCHERS ----
            case 'createVoucher':
                result = createVoucher(token, payload.voucher);
                break;
            case 'updateVoucher':
                result = updateVoucher(token, payload.rowIndex, payload.voucher);
                break;
            case 'updateStatus':
                result = updateVoucherStatus(token, payload.rowIndex, payload.status, payload.pmtMonth);
                break;
            case 'batchUpdateStatus':
                result = batchUpdateStatus(token, payload.controlNumber, payload.status, payload.pmtMonth);
                break;
            case 'assignControlNumber':
                result = assignControlNumber(token, payload.rowIndexes, payload.controlNumber);
                break;
            
            // ✅ COMBINED RELEASE CASES - Fixed!
            case 'releaseSelectedVouchers':
            case 'releaseVouchers':
            case 'releaseToUnit':
                result = releaseSelectedVouchers(token, payload.rowIndexes, payload.targetUnit);
                break;
            
            case 'releaseVouchersWithNotification':
                result = releaseVouchersWithNotification(token, payload.rowIndexes, payload.controlNumber, payload.targetUnit);
                break;
            case 'createVoucherFromLookup':
                result = createVoucherFromLookup(token, payload.lookupResult, payload.additionalData);
                break;

            // ---- DELETE WORKFLOW ----
            case 'requestDelete':
                result = requestVoucherDelete(token, payload.rowIndex, payload.reason, payload.previousStatus);
                break;
            case 'cancelDeleteRequest':
                result = cancelDeleteRequest(token, payload.rowIndex);
                break;
            case 'approveDelete':
                result = approveVoucherDelete(token, payload.rowIndex);
                break;
            case 'rejectDelete':
                result = rejectVoucherDelete(token, payload.rowIndex, payload.reason);
                break;
            case 'deleteVoucher':
                result = deleteVoucher(token, payload.rowIndex);
                break;

            // ---- USERS ----
            case 'createUser':
                result = createUser(token, payload.user);
                break;
            case 'updateUser':
                result = updateUser(token, payload.rowIndex, payload.user);
                break;
            case 'deleteUser':
                result = deleteUser(token, payload.rowIndex);
                break;

            // ---- NOTIFICATIONS ----
            case 'markNotificationRead':
                result = markNotificationRead(token, payload.rowIndex);
                break;
            case 'markAllNotificationsRead':
                result = markAllNotificationsRead(token);
                break;

            // ---- ACTION ITEMS ----
            case 'saveActionItemSettings':
                result = saveActionItemSettings(token, payload.settings);
                break;

            // ---- PASSWORD RESET ----
            case 'requestPasswordReset':
                result = requestPasswordReset(payload.identifier);
                break;
            case 'resetPasswordWithOtp':
                result = resetPasswordWithOtp(payload.identifier, payload.otp, payload.newPassword);
                break;

            // ---- PROFILE ----
            case 'updateMyProfile':
                result = updateMyProfile(token, payload.profile);
                break;

            // ---- DEBT PROFILE WORKFLOW ----
            case 'requestDebtProfile':
                result = requestDebtProfile(token, payload);
                break;
            case 'approveDebtProfile':
                result = handleDebtProfileApproval(token, payload.requestId, 'approve', payload.comments);
                break;
            case 'rejectDebtProfile':
                result = handleDebtProfileApproval(token, payload.requestId, 'reject', payload.comments);
                break;

            // ---- ANNOUNCEMENTS ----
            case 'createAnnouncement':
                result = createAnnouncement(token, payload.announcement || payload);
                break;
            case 'dismissAnnouncement':
                result = dismissAnnouncement(token, payload.announcementId);
                break;

            // TAX OPERATIONS
            case 'getTaxSummary':
                result = getTaxSummary(token, payload.year);
                break;
            case 'getTaxByCategory':
                result = getTaxByCategory(token, payload.year);
                break;
            case 'getTaxByMonth':
                result = getTaxByMonth(token, payload.year);
                break;
            case 'getTaxByPayee':
                result = getTaxByPayee(token, payload.year, payload.filters);
                break;
            case 'recordTaxPayment':
                result = recordTaxPayment(token, payload.payment);
                break;
            case 'getTaxPayments':
                result = getTaxPayments(token, payload.year);
                break;
            case 'updateTaxPayment':
                result = updateTaxPayment(token, payload.paymentId, payload.updates);
                break;
            case 'deleteTaxPayment':
                result = deleteTaxPayment(token, payload.paymentId);
                break;
            case 'getTaxSchedule':
                result = getTaxSchedule(token, payload.year);
                break;
            case 'createTaxSchedule':
                result = createTaxSchedule(token, payload.schedule);
                break;
            case 'updateTaxSchedule':
                result = updateTaxSchedule(token, payload.scheduleId, payload.updates);
                break;
            case 'getTaxCompliance':
                result = getTaxCompliance(token, payload.year);
                break;

            default:
                result = { success: false, error: 'Unknown action' };
        }

        return ContentService.createTextOutput(JSON.stringify(result))
            .setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({
            success: false,
            error: error.message
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

function updateStatus(token, rowIndex, status, pmtMonth) {
  return updateVoucherStatus(token, rowIndex, status, pmtMonth);
}

function requestDelete(token, rowIndex, reason, previousStatus) {
  return requestVoucherDelete(token, rowIndex, reason, previousStatus);
}

function approveDelete(token, rowIndex) {
  return approveVoucherDelete(token, rowIndex);
}

function rejectDelete(token, rowIndex, reason) {
  return rejectVoucherDelete(token, rowIndex, reason);
}

// ==================== UTILITY FUNCTIONS ====================

/**
 * Creates a JSON output with proper CORS headers
 * @param {Object} data - Data to return
 * @returns {TextOutput} Formatted JSON output
 */
function createJsonOutput(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * Direct delete a voucher (ADMIN only) - bypasses approval workflow
 * @param {string} token - Session token
 * @param {number} rowIndex - Row index of the voucher to delete
 */
function deleteVoucher(token, rowIndex) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // STRICTLY for Admin role
    if (session.role !== CONFIG.ROLES.ADMIN) {
      return { 
        success: false, 
        error: 'Unauthorized: Only an Administrator can delete vouchers directly.' 
      };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    rowIndex = parseInt(rowIndex, 10);
    
    if (rowIndex < 2 || rowIndex > lastRow) {
      return { success: false, error: 'Voucher not found' };
    }
    
    // Get voucher details before deleting for logging
    const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
    const voucherNum = row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1];
    const payee = row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1];
    
    // Delete the row
    sheet.deleteRow(rowIndex);
    
    // Clear cache
    clearAllVoucherCaches();
    
    // Log this critical action
    logAudit(session, 'DELETE_DIRECT', 
             'Admin directly deleted voucher: ' + voucherNum + ' (' + payee + ')', 
             CONFIG.SHEETS.VOUCHERS_2026, rowIndex, 
             { payee: payee, voucherNumber: voucherNum });
    
    return { success: true, message: 'Voucher permanently deleted by Admin.' };
    
  } catch (error) {
    console.log('deleteVoucher error: ' + error.message);
    return { success: false, error: 'Direct delete failed: ' + error.message };
  }
}

/**
 * Gets a sheet by name
 * @param {string} sheetName - Name of the sheet
 * @returns {Sheet} Google Sheet object
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  return sheet;
}

/**
 * Simple hash function for passwords
 * Uses SHA-256 for basic security
 * @param {string} password - Plain text password
 * @returns {string} Hashed password
 */
function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  let hash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hex = (rawHash[i] & 0xFF).toString(16);
    if (hex.length === 1) hex = '0' + hex;
    hash += hex;
  }
  return hash;
}

/**
 * Generates a unique session token
 * @param {string} email - User email
 * @returns {string} Session token
 */
function generateToken(email) {
  const timestamp = new Date().getTime();
  const random = Math.random().toString(36).substring(2);
  const raw = email + timestamp + random;
  return Utilities.base64Encode(raw);
}

/**
 * Helper to map header names to column indices (1-based)
 * Used by syncActionItems_
 */
function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h) map[String(h).toUpperCase().trim()] = i + 1;
  });
  return map;
}

/**
 * Helper to check if a user role is in the allowed list
 */
function hasPermission(userRole, allowedRoles) {
  if (!userRole) return false;
  if (Array.isArray(allowedRoles)) {
    return allowedRoles.includes(userRole);
  }
  return userRole === allowedRoles;
}

/**
 * Helper to invalidate Action Item cache
 * Forces a re-sync on the next read
 */
function invalidateActionItemCache_() {
  CacheService.getScriptCache().remove("action_items_sync_ts");
}

/**
 * Helper for Audit Logging timestamp
 */
function getNigerianTimestamp() {
  return Utilities.formatDate(new Date(), "Africa/Lagos", "yyyy-MM-dd HH:mm:ss");
}


/**
 * Get audit trail records (paged) - ADMIN ONLY
 * Returns newest records first.
 *
 * Sheet: AUDIT_TRAIL
 * Columns (9): TIMESTAMP, USER_EMAIL, USER_NAME, ROLE, ACTION, DESCRIPTION, SHEET, ROW_INDEX, EXTRA
 *
 * @param {string} token
 * @param {string|number} limit
 * @param {string|number} offset
 */
function getAuditTrail(token, limit, offset) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };

    // ADMIN ONLY
    if (session.role !== CONFIG.ROLES.ADMIN) {
      return { success: false, error: 'Unauthorized' };
    }

    limit = parseInt(limit, 10);
    offset = parseInt(offset, 10);

    if (isNaN(limit) || limit <= 0) limit = 50;
    if (limit > 200) limit = 200; // safety cap
    if (isNaN(offset) || offset < 0) offset = 0;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('AUDIT_TRAIL');
    if (!sheet) {
      return { success: true, records: [], total: 0 };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, records: [], total: 0 };
    }

    const total = lastRow - 1; // exclude header

    // We want newest first with offset:
    // endRow = lastRow - offset
    // startRow = endRow - limit + 1
    const endRow = lastRow - offset;
    if (endRow < 2) {
      return { success: true, records: [], total: total };
    }

    const startRow = Math.max(2, endRow - limit + 1);
    const numRows = endRow - startRow + 1;

    const values = sheet.getRange(startRow, 1, numRows, 9).getValues();

    // Reverse so newest appears first
    values.reverse();

    const records = values.map(r => ({
      timestamp: r[0],
      email: r[1],
      name: r[2],
      role: r[3],
      action: r[4],
      description: r[5],
      sheet: r[6],
      rowIndex: r[7],
      extra: r[8]
    }));

    return { success: true, records: records, total: total };

  } catch (e) {
    return { success: false, error: 'Failed to get audit trail: ' + e.message };
  }
}

/**
 * Logs audit events to AUDIT_TRAIL sheet
 */
function logAudit(session, action, description, sheetName, rowIndex, extra) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('AUDIT_TRAIL');
    if (!sheet) {
      sheet = ss.insertSheet('AUDIT_TRAIL');
      sheet.getRange(1, 1, 1, 9).setValues([[
        'TIMESTAMP','USER_EMAIL','USER_NAME','ROLE','ACTION','DESCRIPTION','SHEET','ROW_INDEX','EXTRA'
      ]]);
    }

    const ts = getNigerianTimestamp();
    const extraStr = extra ? JSON.stringify(extra) : '';

    sheet.appendRow([
      ts,
      session.email || '',
      session.name || '',
      session.role || '',
      action || '',
      description || '',
      sheetName || '',
      rowIndex || '',
      extraStr
    ]);

  } catch (e) {
    console.log('logAudit error: ' + e.message);
  }
}

/**
 * Get audit trail records (paged) - ADMIN ONLY
 * Returns newest records first.
 */
function getAuditTrail(token, limit, offset) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };

    // ADMIN ONLY
    if (session.role !== CONFIG.ROLES.ADMIN) {
      return { success: false, error: 'Unauthorized' };
    }

    limit = parseInt(limit, 10);
    offset = parseInt(offset, 10);

    if (isNaN(limit) || limit <= 0) limit = 50;
    if (limit > 200) limit = 200;
    if (isNaN(offset) || offset < 0) offset = 0;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('AUDIT_TRAIL');
    if (!sheet) return { success: true, records: [], total: 0 };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, records: [], total: 0 };

    const total = lastRow - 1;

    // Newest first:
    const endRow = lastRow - offset;
    if (endRow < 2) return { success: true, records: [], total: total };

    const startRow = Math.max(2, endRow - limit + 1);
    const numRows = endRow - startRow + 1;

    const values = sheet.getRange(startRow, 1, numRows, 9).getValues();
    values.reverse();

    const records = values.map(r => ({
      timestamp: r[0],
      email: r[1],
      name: r[2],
      role: r[3],
      action: r[4],
      description: r[5],
      sheet: r[6],
      rowIndex: r[7],
      extra: r[8]
    }));

    return { success: true, records, total };

  } catch (e) {
    return { success: false, error: 'Failed to get audit trail: ' + e.message };
  }
}

/**
 * Stores session in Properties Service
 * @param {string} token - Session token
 * @param {Object} userData - User data to store
 */
function storeSession(token, userData) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const session = {
    ...userData,
    createdAt: new Date().getTime(),
    expiresAt: new Date().getTime() + CONFIG.SESSION_DURATION
  };
  scriptProperties.setProperty('session_' + token, JSON.stringify(session));
}

/**
 * Retrieves session from Properties Service
 * @param {string} token - Session token
 * @returns {Object|null} Session data or null if invalid/expired
 */
function getSession(token) {
  if (!token) return null;
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const sessionStr = scriptProperties.getProperty('session_' + token);
  
  if (!sessionStr) return null;
  
  const session = JSON.parse(sessionStr);
  
  // Check if expired
  if (new Date().getTime() > session.expiresAt) {
    scriptProperties.deleteProperty('session_' + token);
    return null;
  }
  
  return session;
}

/**
 * Clears a session
 * @param {string} token - Session token to clear
 */
function clearSession(token) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('session_' + token);
}

/**
 * Validates if user has required permission
 * @param {string} userRole - User's role
 * @param {string[]} allowedRoles - Array of allowed roles
 * @returns {boolean} True if authorized
 */
function hasPermission(userRole, allowedRoles) {
  // ADMIN has all permissions
  if (userRole === CONFIG.ROLES.ADMIN) return true;
  return allowedRoles.includes(userRole);
}

/**
 * Formats currency value
 * @param {number|string} value - Numeric value
 * @returns {number} Parsed number
 */
function parseAmount(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  // Remove currency symbols and commas
  const cleaned = String(value).replace(/[₦,\s]/g, '');
  return parseFloat(cleaned) || 0;
}

/**
 * Gets current timestamp in Nigerian timezone
 * @returns {string} Formatted timestamp
 */
function getNigerianTimestamp() {
  const now = new Date();
  // Nigeria is UTC+1
  const offset = 1 * 60 * 60 * 1000;
  const nigerianTime = new Date(now.getTime() + offset);
  return Utilities.formatDate(nigerianTime, 'GMT+1', 'yyyy-MM-dd HH:mm:ss');
}

function requestPasswordReset(identifier) {
  try {
    // Always respond generically to prevent account enumeration
    const generic = { success: true, message: 'If the account exists, an OTP has been sent.' };

    if (!identifier) return generic;

    const input = String(identifier).trim().toLowerCase();
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    let email = '';
    let active = false;

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][CONFIG.USER_COLUMNS.EMAIL - 1] || '').trim().toLowerCase();
      const rowUsername = String(data[i][CONFIG.USER_COLUMNS.USERNAME - 1] || '').trim().toLowerCase();
      const isActiveVal = data[i][CONFIG.USER_COLUMNS.ACTIVE - 1];
      const isActive = (isActiveVal === true || String(isActiveVal).toUpperCase() === 'TRUE');

      if (rowEmail === input || (rowUsername && rowUsername === input)) {
        email = rowEmail;
        active = isActive;
        break;
      }
    }

    if (!email || !active) return generic;

    const cache = CacheService.getScriptCache();

    // Rate limit: 1 per 60 seconds
    const rlKey = 'pwdreset_rl_' + email;
    if (cache.get(rlKey)) return generic;
    cache.put(rlKey, '1', 60);

    const otp = String(Math.floor(100000 + Math.random() * 900000)); // 6 digits
    const otpHash = hashPassword(otp);

    // Store OTP for 10 minutes
    const otpKey = 'pwdreset_otp_' + email;
    cache.put(otpKey, otpHash, 10 * 60);

    MailApp.sendEmail({
      to: email,
      subject: '[PAYABLE VOUCHER] Password Reset OTP',
      htmlBody: `
        <p>Your password reset OTP is:</p>
        <h2 style="letter-spacing:2px;">${otp}</h2>
        <p>This OTP expires in <strong>10 minutes</strong>.</p>
        <p>If you did not request this, please ignore this email.</p>
      `
    });

    return generic;
  } catch (e) {
    return { success: false, error: 'Failed to send OTP: ' + e.message };
  }
}

function resetPasswordWithOtp(identifier, otp, newPassword) {
  try {
    if (!identifier || !otp || !newPassword) {
      return { success: false, error: 'Identifier, OTP and new password are required' };
    }
    if (String(newPassword).length < 6) {
      return { success: false, error: 'New password must be at least 6 characters' };
    }

    // Resolve identifier -> email
    const input = String(identifier).trim().toLowerCase();
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    let email = '';
    let userRowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][CONFIG.USER_COLUMNS.EMAIL - 1] || '').trim().toLowerCase();
      const rowUsername = String(data[i][CONFIG.USER_COLUMNS.USERNAME - 1] || '').trim().toLowerCase();

      if (rowEmail === input || (rowUsername && rowUsername === input)) {
        email = rowEmail;
        userRowIndex = i + 1; // sheet row
        break;
      }
    }

    if (!email || userRowIndex < 2) {
      return { success: false, error: 'Invalid OTP or account' };
    }

    // Verify OTP
    const cache = CacheService.getScriptCache();
    const otpKey = 'pwdreset_otp_' + email;
    const storedOtpHash = cache.get(otpKey);

    if (!storedOtpHash) {
      return { success: false, error: 'OTP expired or invalid' };
    }

    const providedHash = hashPassword(String(otp).trim());
    if (providedHash !== storedOtpHash) {
      return { success: false, error: 'OTP expired or invalid' };
    }

    // Update password
    const hashedNew = hashPassword(newPassword);
    sheet.getRange(userRowIndex, CONFIG.USER_COLUMNS.PASSWORD).setValue(hashedNew);

    // Clear OTP (one-time use)
    cache.remove(otpKey);

    return { success: true, message: 'Password reset successful. You can now login.' };

  } catch (e) {
    return { success: false, error: 'Reset failed: ' + e.message };
  }
}


// =============================================================================
// 2. AUTH.gs — LOGIN, SESSION, VALIDATION
// =============================================================================
// This section includes enhanced authentication with password hashing and migration.

// --- Core Authentication Functions ---

/**
 * Handles user login with email/username and password.
 * Features password hashing and automatic upgrade for plain-text passwords.
 * @param {string} identifier - User email or username.
 * @param {string} password - User password.
 * @returns {Object} Login result with token and user data.
 */
function login(identifier, password) {
  try {
    if (!identifier || !password) {
      return { success: false, error: 'Email/Username and password are required' };
    }

    const userRecord = findUserByEmailOrUsername_(identifier);

    if (!userRecord) {
      return { success: false, error: 'User not found' };
    }

    const { rowIndex, data: row } = userRecord;
    const cols = CONFIG.USER_COLUMNS;

    const isActiveVal = row[cols.ACTIVE - 1];
    const isActive = (isActiveVal === true || String(isActiveVal).toUpperCase() === 'TRUE');
    if (!isActive) {
      return { success: false, error: 'Account is deactivated. Contact administrator.' };
    }

    const storedPassword = String(row[cols.PASSWORD - 1] || '');
    const hashedInput = hashPassword(password);

    // IMPORTANT: The `storedPassword === password` check is for migrating old plain-text passwords.
    // It should be removed after all users have logged in at least once.
    if (storedPassword !== hashedInput && storedPassword !== password && storedPassword !== '') {
      return { success: false, error: 'Invalid email/username or password.' };
    }

    // Upgrade plain/empty password to hashed version
    if (storedPassword !== hashedInput) {
      const sheet = getSheet(CONFIG.SHEETS.USERS);
      sheet.getRange(rowIndex, cols.PASSWORD).setValue(hashedInput);
    }

    const userEmail = String(row[cols.EMAIL - 1] || '').trim().toLowerCase();
    const userData = {
      name: row[cols.NAME - 1],
      email: userEmail,
      username: String(row[cols.USERNAME - 1] || ''),
      role: row[cols.ROLE - 1],
      rowIndex: rowIndex
    };

    const token = generateToken(userEmail);
    storeSession(token, userData);

    return {
      success: true,
      token,
      user: {
        name: userData.name,
        email: userData.email,
        username: userData.username,
        role: userData.role
      }
    };
  } catch (error) {
    return { success: false, error: 'Login failed: ' + error.message };
  }
}

/**
 * Alias for login() - used by test functions
 */
function handleLogin(email, password) {
  return login(email, password);
}

/**
 * Alias for logout() - used by test functions
 */
function handleLogout(token) {
  return logout(token);
}

/**
 * Validates a session token.
 * @param {string} token - Session token.
 * @returns {Object} Validation result.
 */
function validateSession(token) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired or invalid' };
    }
    return {
      success: true,
      user: {
        name: session.name,
        email: session.email,
        role: session.role
      }
    };
  } catch (error) {
    return { success: false, error: 'Session validation failed: ' + error.message };
  }
}

/**
 * Logs out a user by clearing their session.
 * @param {string} token - Session token.
 * @returns {Object} Logout result.
 */
function logout(token) {
  try {
    clearSession(token);
    return { success: true, message: 'Logged out successfully' };
  } catch (error) {
    return { success: false, error: 'Logout failed: ' + error.message };
  }
}

/**
 * Changes a user's password.
 * @param {string} token - Session token.
 * @param {string} oldPassword - Current password.
 * @param {string} newPassword - New password.
 * @returns {Object} Result object.
 */
function changePassword(token, oldPassword, newPassword) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired. Please login again.' };
    }

    if (!newPassword || newPassword.length < 6) {
      return { success: false, error: 'New password must be at least 6 characters' };
    }

    const userRecord = findUserByEmailOrUsername_(session.email);
    if (!userRecord) {
      return { success: false, error: 'User not found' };
    }

    const { rowIndex, data: row } = userRecord;
    const cols = CONFIG.USER_COLUMNS;

    const storedPassword = String(row[cols.PASSWORD - 1]);
    const hashedOld = hashPassword(oldPassword);

    if (storedPassword !== hashedOld && storedPassword !== oldPassword) {
      return { success: false, error: 'Current password is incorrect' };
    }

    const hashedNew = hashPassword(newPassword);
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    sheet.getRange(rowIndex, cols.PASSWORD).setValue(hashedNew);

    return { success: true, message: 'Password changed successfully' };

  } catch (error) {
    return { success: false, error: 'Password change failed: ' + error.message };
  }
}

// --- Authentication Helpers ---

/**
 * Retrieves a session from cache by token.
 */
function getSession(token) {
    if (!token) return null;
    try {
        const cache = CacheService.getScriptCache();
        const sessionStr = cache.get('session_' + token);
        if (!sessionStr) return null;
        return JSON.parse(sessionStr);
    } catch (e) {
        return null;
    }
}

/**
 * Stores session data in cache.
 */
function storeSession(token, userData) {
  const cache = CacheService.getScriptCache();
  const sessionData = JSON.stringify({
    token: token,
    email: userData.email,
    name: userData.name,
    role: userData.role,
    username: userData.username,
    rowIndex: userData.rowIndex,
    createdAt: new Date().toISOString()
  });
  const sessionDuration = (CONFIG && CONFIG.SESSION_DURATION) || 28800; // Default 8 hours
  cache.put('session_' + token, sessionData, sessionDuration);
}

/**
 * Removes a session from cache.
 */
function clearSession(token) {
    if (token) {
        const cache = CacheService.getScriptCache();
        cache.remove('session_' + token);
    }
}

/**
 * Generates a unique token.
 */
function generateToken(userEmail) {
  return Utilities.getUuid();
}

/**
 * Hashes a password using SHA-256.
 * NOTE: In a real-world, high-security application, a slower, salted hashing
 * algorithm like bcrypt or scrypt would be preferable, but SHA-256 is a
 * good, available option in Apps Script.
 */
function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return digest.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * Finds a user row by email or username.
 * @param {string} identifier - The email or username to search for.
 * @returns {Object|null} An object with rowIndex and data, or null if not found.
 */
function findUserByEmailOrUsername_(identifier) {
  const input = String(identifier || '').trim().toLowerCase();
  if (!input) return null;

  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  const cols = CONFIG.USER_COLUMNS;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const userEmail = String(row[cols.EMAIL - 1] || '').trim().toLowerCase();
    const username = String(row[cols.USERNAME - 1] || '').trim().toLowerCase();

    if (userEmail === input || (username && username === input)) {
      return {
        rowIndex: i + 1,
        data: row
      };
    }
  }
  return null;
}

function normalizeUsername_(u) {
  return String(u || '').trim().toLowerCase();
}

function isUsernameTaken_(sheet, usernameLower, excludeRowIndex) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowIndex = i + 1;
    if (excludeRowIndex && rowIndex === excludeRowIndex) continue;

    const u = normalizeUsername_(data[i][CONFIG.USER_COLUMNS.USERNAME - 1]);
    if (u && u === usernameLower) return true;
  }
  return false;
}

/**
 * Helper: checks if a user's role is in the allowed list.
 */
function hasPermission(role, allowedRoles) {
    if (!role) return false;
    if (role === CONFIG.ROLES.ADMIN) return true;
    if (Array.isArray(allowedRoles)) return allowedRoles.includes(role);
    return role === allowedRoles;
}


// ==================== USER MANAGEMENT FUNCTIONS ====================

/**
 * Gets all users (ADMIN only)
 * @param {string} token - Session token
 * @returns {Object} List of users
 */
function getUsers(token) {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };

    // Only admins can list all users
    if (!hasPermission(session.role, [CONFIG.ROLES.ADMIN, CONFIG.ROLES.DFA, CONFIG.ROLES.DDFA])) {
        return { success: false, error: 'Unauthorized' };
    }

    try {
        const sheet = getSheet(CONFIG.SHEETS.USERS);
        const data = sheet.getDataRange().getValues();
        const cols = CONFIG.USER_COLUMNS;

        const users = [];
        for (let i = 1; i < data.length; i++) {
            users.push({
                rowIndex: i + 1,
                name: String(data[i][cols.NAME - 1] || ''),
                email: String(data[i][cols.EMAIL - 1] || ''),
                role: String(data[i][cols.ROLE - 1] || ''),
                active: String(data[i][cols.ACTIVE - 1] || 'true'),
                username: String(data[i][cols.USERNAME - 1] || ''),
                phone: String(data[i][cols.PHONE - 1] || ''),
                department: String(data[i][cols.DEPARTMENT - 1] || '')
            });
        }

        return { success: true, users: users, count: users.length };
    } catch (err) {
        return { success: false, error: 'Failed to get users: ' + err.message };
    }
}

/**
 * Creates a new user (ADMIN only)
 * @param {string} token - Session token
 * @param {Object} user - User data
 * @returns {Object} Result
 */
function createUser(token, user) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };

    // Only ADMIN can create users
    if (!hasPermission(session.role, [CONFIG.ROLES.ADMIN])) {
      return { success: false, error: 'Unauthorized: Admin access required' };
    }

    // Validate required fields
    if (!user || !user.name || !user.email || !user.role) {
      return { success: false, error: 'Name, email, and role are required' };
    }

    // Username required (as you requested)
    if (!user.username || !String(user.username).trim()) {
      return { success: false, error: 'Username is required' };
    }

    // Validate role
    const validRoles = Object.values(CONFIG.ROLES);
    if (!validRoles.includes(user.role)) {
      return { success: false, error: 'Invalid role. Valid roles: ' + validRoles.join(', ') };
    }

    // Validate email format
    const normalizedEmail = String(user.email).trim().toLowerCase();
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(normalizedEmail)) {
      return { success: false, error: 'Invalid email format' };
    }

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    // Check email uniqueness
    for (let i = 1; i < data.length; i++) {
      const existingEmail = String(data[i][CONFIG.USER_COLUMNS.EMAIL - 1] || '').trim().toLowerCase();
      if (existingEmail === normalizedEmail) {
        return { success: false, error: 'A user with this email already exists' };
      }
    }

    // Check username uniqueness (requires USERNAME column to exist in CONFIG)
    const usernameLower = String(user.username).trim().toLowerCase();
    if (CONFIG.USER_COLUMNS.USERNAME) {
      for (let i = 1; i < data.length; i++) {
        const existingUsername = String(data[i][CONFIG.USER_COLUMNS.USERNAME - 1] || '').trim().toLowerCase();
        if (existingUsername && existingUsername === usernameLower) {
          return { success: false, error: 'Username already exists. Choose another.' };
        }
      }
    }

    // Default password
    const defaultPassword = 'Welcome123';
    const hashedPassword = hashPassword(defaultPassword);

    // Build row (append-only safe)
    const newRow = [
      String(user.name).trim(),         // NAME
      normalizedEmail,                  // EMAIL
      hashedPassword,                   // PASSWORD
      user.role,                        // ROLE
      true                              // ACTIVE
    ];

    // If you already added new columns, append them:
    if (CONFIG.USER_COLUMNS.USERNAME) newRow.push(String(user.username).trim());
    if (CONFIG.USER_COLUMNS.PHONE) newRow.push(String(user.phone || '').trim());
    if (CONFIG.USER_COLUMNS.DEPARTMENT) newRow.push(String(user.department || '').trim());
    if (CONFIG.USER_COLUMNS.UPDATED_AT) newRow.push(new Date());

    sheet.appendRow(newRow);

    return {
      success: true,
      message: 'User created successfully. Default password is: ' + defaultPassword,
      user: {
        name: String(user.name).trim(),
        email: normalizedEmail,
        username: String(user.username).trim(),
        role: user.role
      }
    };

  } catch (error) {
    return { success: false, error: 'Failed to create user: ' + error.message };
  }
}

/**
 * Updates an existing user (ADMIN only)
 * @param {string} token - Session token
 * @param {number} rowIndex - Row index of user to update
 * @param {Object} user - Updated user data
 * @returns {Object} Result
 */
function updateUser(token, rowIndex, user) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Only ADMIN can update users
    if (!hasPermission(session.role, [CONFIG.ROLES.ADMIN])) {
      return { success: false, error: 'Unauthorized: Admin access required' };
    }
    
    // Validate row index
    if (!rowIndex || rowIndex < 2) {
      return { success: false, error: 'Invalid user reference' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const lastRow = sheet.getLastRow();
    
    if (rowIndex > lastRow) {
      return { success: false, error: 'User not found' };
    }
    
    // Get current data
    const currentData = sheet.getRange(rowIndex, 1, 1, 5).getValues()[0];
    
    // Update fields (keep password unchanged unless explicitly provided)
    if (user.name) {
      sheet.getRange(rowIndex, CONFIG.USER_COLUMNS.NAME).setValue(user.name.trim());
    }
    
    if (user.email) {
      // Check if new email conflicts with existing users
      const data = sheet.getDataRange().getValues();
      const normalizedEmail = user.email.trim().toLowerCase();
      
      for (let i = 1; i < data.length; i++) {
        if (i + 1 !== rowIndex) { // Skip current user
          const existingEmail = String(data[i][CONFIG.USER_COLUMNS.EMAIL - 1]).trim().toLowerCase();
          if (existingEmail === normalizedEmail) {
            return { success: false, error: 'Another user with this email already exists' };
          }
        }
      }
      
      sheet.getRange(rowIndex, CONFIG.USER_COLUMNS.EMAIL).setValue(normalizedEmail);
    }
    
    if (user.role) {
      const validRoles = Object.values(CONFIG.ROLES);
      if (!validRoles.includes(user.role)) {
        return { success: false, error: 'Invalid role' };
      }
      sheet.getRange(rowIndex, CONFIG.USER_COLUMNS.ROLE).setValue(user.role);
    }
    
    if (user.active !== undefined) {
      sheet.getRange(rowIndex, CONFIG.USER_COLUMNS.ACTIVE).setValue(user.active);
    }
    
    // Reset password if requested
    if (user.resetPassword) {
      const defaultPassword = 'Welcome123';
      const hashedPassword = hashPassword(defaultPassword);
      sheet.getRange(rowIndex, CONFIG.USER_COLUMNS.PASSWORD).setValue(hashedPassword);
      return { 
        success: true, 
        message: 'User updated. Password reset to: ' + defaultPassword 
      };
    }

    if (user.username !== undefined) {
      const usernameLower = normalizeUsername_(user.username);
      if (!usernameLower) return { success: false, error: 'Username is required' };

      if (isUsernameTaken_(sheet, usernameLower, rowIndex)) {
        return { success: false, error: 'Username already exists. Choose another.' };
      }

      sheet.getRange(rowIndex, CONFIG.USER_COLUMNS.USERNAME).setValue(user.username.trim());
    }
    
    return { success: true, message: 'User updated successfully' };
    
  } catch (error) {
    return { success: false, error: 'Failed to update user: ' + error.message };
  }
}

/**
 * Deletes a user (ADMIN only)
 * @param {string} token - Session token
 * @param {number} rowIndex - Row index of user to delete
 * @returns {Object} Result
 */
function deleteUser(token, rowIndex) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Only ADMIN can delete users
    if (!hasPermission(session.role, [CONFIG.ROLES.ADMIN])) {
      return { success: false, error: 'Unauthorized: Admin access required' };
    }
    
    // Validate row index
    if (!rowIndex || rowIndex < 2) {
      return { success: false, error: 'Invalid user reference' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const lastRow = sheet.getLastRow();
    
    if (rowIndex > lastRow) {
      return { success: false, error: 'User not found' };
    }
    
    // Get user email to prevent self-deletion
    const userEmail = sheet.getRange(rowIndex, CONFIG.USER_COLUMNS.EMAIL).getValue();
    if (String(userEmail).toLowerCase() === session.email.toLowerCase()) {
      return { success: false, error: 'You cannot delete your own account' };
    }
    
    // Delete the row
    sheet.deleteRow(rowIndex);
    
    return { success: true, message: 'User deleted successfully' };
    
  } catch (error) {
    return { success: false, error: 'Failed to delete user: ' + error.message };
  }
}

/**
 * Gets role permissions for frontend rendering
 * @param {string} role - User role
 * @returns {Object} Permission set
 */
function getRolePermissions(token) {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };

    // Define permissions per role
    const permissions = {
        [CONFIG.ROLES.ADMIN]: {
            canCreateVoucher: true, canEditVoucher: true, canDeleteVoucher: true,
            canManageUsers: true, canViewReports: true, canRelease: true,
            canApproveDelete: true, canViewAudit: true, canManageSettings: true
        },
        [CONFIG.ROLES.DFA]: {
            canCreateVoucher: false, canEditVoucher: false, canDeleteVoucher: false,
            canManageUsers: false, canViewReports: true, canRelease: false,
            canApproveDelete: true, canViewAudit: true, canManageSettings: false
        },
        [CONFIG.ROLES.DDFA]: {
            canCreateVoucher: false, canEditVoucher: false, canDeleteVoucher: false,
            canManageUsers: false, canViewReports: true, canRelease: false,
            canApproveDelete: true, canViewAudit: true, canManageSettings: false
        },
        [CONFIG.ROLES.CPO]: {
            canCreateVoucher: false, canEditVoucher: true, canDeleteVoucher: false,
            canManageUsers: false, canViewReports: true, canRelease: false,
            canApproveDelete: false, canViewAudit: false, canManageSettings: false
        },
        [CONFIG.ROLES.AUDIT]: {
            canCreateVoucher: false, canEditVoucher: false, canDeleteVoucher: false,
            canManageUsers: false, canViewReports: true, canRelease: false,
            canApproveDelete: false, canViewAudit: true, canManageSettings: false
        },
        [CONFIG.ROLES.PAYABLE_HEAD]: {
            canCreateVoucher: true, canEditVoucher: true, canDeleteVoucher: true,
            canManageUsers: false, canViewReports: true, canRelease: true,
            canApproveDelete: false, canViewAudit: false, canManageSettings: false
        },
        [CONFIG.ROLES.PAYABLE_STAFF]: {
            canCreateVoucher: true, canEditVoucher: true, canDeleteVoucher: false,
            canManageUsers: false, canViewReports: false, canRelease: true,
            canApproveDelete: false, canViewAudit: false, canManageSettings: false
        },
                [CONFIG.ROLES.TAX]: { // NEW ROLE PERMISSIONS
            canCreateVoucher: false, canEditVoucher: false, canDeleteVoucher: false,
            canManageUsers: false, canViewReports: true, canRelease: false,
            canApproveDelete: false, canViewAudit: false, canManageSettings: false,
            canAccessTax: true // Only Tax can access Tax page
        }
    };

    const rolePerms = permissions[session.role] || {};
    return { success: true, permissions: rolePerms, role: session.role };
}

/**
 * Gets available categories from SYSTEM_CONFIG
 * @param {string} token - Session token
 * @returns {Object} Categories list
 */
function getCategories(token) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.SYSTEM_CONFIG);
    const data = sheet.getDataRange().getValues();
    
    // Find CATEGORIES row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'CATEGORIES') {
        const categoriesStr = data[i][1];
        const categories = categoriesStr.split(',').map(c => c.trim()).filter(c => c);
        return { success: true, categories: categories };
      }
    }
    
    // Default categories if not found
    return { 
      success: true, 
      categories: ['Drugs', 'Equipment', 'Services', 'Consultancy', 'Maintenance', 'Supplies', 'Others']
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to get categories: ' + error.message };
  }
}

/**
 * Gets list of valid roles
 * @returns {Object} Roles list
 */
function getAvailableRoles() {
  return {
    success: true,
    roles: Object.values(CONFIG.ROLES)
  };
}

// ==================== SETUP & TESTING FUNCTIONS ====================

/**
 * One-time setup: Creates admin user with password
 * Run this function once to set up your admin account
 */
function setupAdminUser() {
  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  
  // Check if admin exists
  let adminExists = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][CONFIG.USER_COLUMNS.ROLE - 1] === 'ADMIN') {
      adminExists = true;
      // Set password for existing admin
      const hashedPassword = hashPassword('Admin@2026');
      sheet.getRange(i + 1, CONFIG.USER_COLUMNS.PASSWORD).setValue(hashedPassword);
      sheet.getRange(i + 1, CONFIG.USER_COLUMNS.ACTIVE).setValue(true);
      Logger.log('Admin password updated. Email: ' + data[i][CONFIG.USER_COLUMNS.EMAIL - 1]);
      break;
    }
  }
  
  if (!adminExists) {
    // Create admin user
    const hashedPassword = hashPassword('Admin@2026');
    const newRow = ['Administrator', 'admin@fmc.gov.ng', hashedPassword, 'ADMIN', true];
    sheet.appendRow(newRow);
    Logger.log('Admin user created. Email: admin@fmc.gov.ng');
  }
  
  Logger.log('Setup complete! Default password is: Admin@2026');
}

/**
 * Test function to verify login works
 */
function testLogin() {
  const result = handleLogin('admin@fmc.gov.ng', 'Admin@2026');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * Test function to check session validation
 */
function testSession() {
  // First login
  const loginResult = handleLogin('admin@fmc.gov.ng', 'Admin@2026');
  Logger.log('Login result: ' + JSON.stringify(loginResult));
  
  if (loginResult.success) {
    // Test session validation
    const sessionResult = validateSession(loginResult.token);
    Logger.log('Session result: ' + JSON.stringify(sessionResult));
    
    // Test getting users
    const usersResult = getUsers(loginResult.token);
    Logger.log('Users result: ' + JSON.stringify(usersResult));
  }
}

/**
 * Clears all sessions (use if you have issues)
 */
function clearAllSessions() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const allProperties = scriptProperties.getProperties();
  
  for (let key in allProperties) {
    if (key.startsWith('session_')) {
      scriptProperties.deleteProperty(key);
    }
  }
  
  Logger.log('All sessions cleared');
}

function testMailAuthorization_() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), 'Test OTP Mail Permission', 'Mail permission granted.');
}

// ==================== PHASE 3 TEST FUNCTIONS ====================

/**
 * Test creating a voucher
 */
function testCreateVoucher() {
  // First login to get token
  const loginResult = handleLogin('admin@fmc.gov.ng', 'Admin@2026');
  
  if (!loginResult.success) {
    Logger.log('Login failed: ' + loginResult.error);
    return;
  }
  
  const token = loginResult.token;
  
  // Create test voucher
  const testVoucher = {
    payee: 'Test Contractor Ltd',
    accountOrMail: 'FMCA/RFA/TRF/236/2026',
    particular: 'Supply of Medical Equipment',
    contractSum: 23000000,
    grossAmount: 20000000,
    net: 1000000,
    vat: 1000000,
    wht: 500000,
    stampDuty: 50000,
    categories: 'Equipment',
    totalGross: 2000000,
    controlNumber: 'CN-2026-002',
    oldVoucherNumber: 'nil',
    date: '2026-01-16',
    accountType: 'TSA'
  };
  
  const result = createVoucher(token, testVoucher);
  Logger.log('Create voucher result: ' + JSON.stringify(result, null, 2));
}

/**
 * Test getting vouchers
 */
function testGetVouchers() {
  const loginResult = handleLogin('oyewusi.adebayo1@gmail.com', 'admin');
  
  if (!loginResult.success) {
    Logger.log('Login failed');
    return;
  }
  
  const result = getVouchers(loginResult.token, '2026', null);
  Logger.log('Get vouchers result: ' + JSON.stringify(result, null, 2));
}

/**
 * Test lookup feature
 * NOTE: This will only work if you have data in the 2025/2024/2023/<2023 sheets
 */
function testLookup() {
  const loginResult = handleLogin('admin@fmc.gov.ng', 'Admin@2026');
  
  if (!loginResult.success) {
    Logger.log('Login failed');
    return;
  }
  
  // Replace with an actual old voucher number from your previous year sheets
  const result = lookupVoucher(loginResult.token, 'SAMPLE-OVN-001', null);
  Logger.log('Lookup result: ' + JSON.stringify(result, null, 2));
}

/**
 * Test status update
 */
function testStatusUpdate() {
  const loginResult = handleLogin('admin@fmc.gov.ng', 'Admin@2026');
  
  if (!loginResult.success) {
    Logger.log('Login failed');
    return;
  }
  
  // Update row 2 (first data row) - adjust as needed
  const result = updateVoucherStatus(loginResult.token, 2, 'Paid', 'January');
  Logger.log('Status update result: ' + JSON.stringify(result, null, 2));
}

// ==================== SUMMARY & REPORTING FUNCTIONS ====================

/**
 * Gets comprehensive summary for a specific year
 * Calculations follow FMC definitions:
 * - Total Vouchers Raised: COUNT if ACCOUNT OR MAIL is not empty
 * - Paid Vouchers (amount): SUM(GROSS AMOUNT) where STATUS = "Paid"
 * - Unpaid Vouchers (amount): SUM(GROSS AMOUNT) where STATUS = "Unpaid"
 * - Cancelled Vouchers (amount): SUM(GROSS AMOUNT) where STATUS = "Cancelled"
 * - Total Contract Sum: SUM(CONTRACT SUM)
 * - Total Debt: Total Contract Sum - (Paid + Cancelled)
 * - Average Payment Rate: Paid voucher COUNT / Total voucher COUNT * 100
 * - Revalidated Vouchers: COUNT if OLD VOUCHER NUMBER is not empty OR
 *   OLD VOUCHER NO AVAILABLE? = "YES"
 */
function resolveVoucherSummaryColumns_(sheet) {
  const cfg = CONFIG.VOUCHER_COLUMNS || {};
  const header = getHeaderMap_(sheet);

  function colFromHeaderOrConfig_(headerNames, fallbackCol1Based) {
    for (let i = 0; i < headerNames.length; i++) {
      const idx = header[String(headerNames[i] || '').toUpperCase().trim()];
      if (idx) return idx - 1; // zero-based
    }
    return Math.max((fallbackCol1Based || 1) - 1, 0);
  }

  const out = {
    STATUS_COL: colFromHeaderOrConfig_(['STATUS'], cfg.STATUS),
    PMT_MONTH_COL: colFromHeaderOrConfig_(['PMT MONTH', 'PAYMENT MONTH'], cfg.PMT_MONTH),
    PAYEE_COL: colFromHeaderOrConfig_(['PAYEE'], cfg.PAYEE),
    ACCT_COL: colFromHeaderOrConfig_(['ACCOUNT OR MAIL', 'ACCOUNT OR EMAIL (VOUCHER NUMBER)', 'VOUCHER NUMBER', 'VOUCHER NO.', 'VOUCHER NO'], cfg.ACCOUNT_OR_MAIL),
    CONTRACT_COL: colFromHeaderOrConfig_(['CONTRACT SUM'], cfg.CONTRACT_SUM),
    GROSS_COL: colFromHeaderOrConfig_(['GROSS AMOUNT'], cfg.GROSS_AMOUNT),
    VAT_COL: colFromHeaderOrConfig_(['VAT'], cfg.VAT),
    WHT_COL: colFromHeaderOrConfig_(['WHT', 'WITHHOLDING TAX'], cfg.WHT),
    STAMP_COL: colFromHeaderOrConfig_(['STAMP DUTY'], cfg.STAMP_DUTY),
    CATEGORY_COL: colFromHeaderOrConfig_(['CATEGORIES', 'CATEGORY'], cfg.CATEGORIES),
    OLD_VN_COL: colFromHeaderOrConfig_(['OLD VOUCHER NUMBER', 'OLD VOUCHER NO'], cfg.OLD_VOUCHER_NUMBER),
    OLD_VN_AVAILABLE_COL: colFromHeaderOrConfig_(['OLD VOUCHER NO AVAILABLE?', 'OLD VOUCHER AVAILABLE'], cfg.OLD_VOUCHER_AVAILABLE),
    ACCOUNT_TYPE_COL: colFromHeaderOrConfig_(['ACCOUNT TYPE'], cfg.ACCOUNT_TYPE),
    SUB_ACCT_COL: colFromHeaderOrConfig_(['SUB ACCOUNT', 'SUB ACCOUNT TYPE'], cfg.SUB_ACCOUNT_TYPE)
  };

  out.LAST_COL = Math.max(
    out.STATUS_COL, out.PMT_MONTH_COL, out.PAYEE_COL, out.ACCT_COL,
    out.CONTRACT_COL, out.GROSS_COL, out.VAT_COL, out.WHT_COL,
    out.STAMP_COL, out.CATEGORY_COL, out.OLD_VN_COL, out.OLD_VN_AVAILABLE_COL,
    out.ACCOUNT_TYPE_COL, out.SUB_ACCT_COL
  ) + 1;

  return out;
}

function getSummary(token, year) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Roles allowed to view summary
    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.AUDIT,
      CONFIG.ROLES.DDFA,
      CONFIG.ROLES.DFA,
      CONFIG.ROLES.ADMIN,
      CONFIG.ROLES.TAX  // ADD TAX ROLE HERE
    ])) {
      return { success: false, error: 'Unauthorized' };
    }
    
    year = year || '2026';
    
    // Pick sheet name based on year
    let sheetName = CONFIG.SHEETS.VOUCHERS_2026;
    if (year !== '2026') {
      switch (year) {
        case '2025':  sheetName = CONFIG.SHEETS.VOUCHERS_2025;        break;
        case '2024':  sheetName = CONFIG.SHEETS.VOUCHERS_2024;        break;
        case '2023':  sheetName = CONFIG.SHEETS.VOUCHERS_2023;        break;
        case '<2023': sheetName = CONFIG.SHEETS.VOUCHERS_BEFORE_2023; break;
      }
    }
    
    const sheet = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      // No data rows
      return {
        success: true,
        year: year,
        summary: {
          totalVouchersRaised: 0,
          paidVouchers: 0,
          unpaidVouchers: 0,
          cancelledVouchers: 0,
          revalidatedVouchers: 0,
          revalidatedWithoutOldNumber: 0,
          totalProcessedContractSum: 0,
          totalPaidAmount: 0,
          totalUnpaidAmount: 0,
          totalCancelledAmount: 0,
          totalDebt: 0,
          averagePaymentPercent: 0
        },
        taxSummary: {
          totalVAT: 0,
          totalWHT: 0,
          totalStampDuty: 0,
          totalTaxLiability: 0,
          paidVAT: 0,
          paidWHT: 0,
          paidStampDuty: 0,
          outstandingVAT: 0,
          outstandingWHT: 0,
          outstandingStampDuty: 0
        },
        categoryBreakdown: [],
        monthlyBreakdown: [],
        accountTypeBreakdown: []
      };
    }
    
    // Column indices resolved by header first (supports actual Q/T layout), fallback to CONFIG.
    const c = resolveVoucherSummaryColumns_(sheet);
    const data = sheet.getRange(2, 1, lastRow - 1, c.LAST_COL).getValues();
    const STATUS_COL   = c.STATUS_COL;
    const PMT_MONTH_COL= c.PMT_MONTH_COL;
    const PAYEE_COL    = c.PAYEE_COL;
    const ACCT_COL     = c.ACCT_COL;
    const CONTRACT_COL = c.CONTRACT_COL;
    const GROSS_COL    = c.GROSS_COL;
    const VAT_COL      = c.VAT_COL;
    const WHT_COL      = c.WHT_COL;
    const STAMP_COL    = c.STAMP_COL;
    const CATEGORY_COL = c.CATEGORY_COL;
    const OLD_VN_COL   = c.OLD_VN_COL;
    const OLD_VN_AVAILABLE_COL = c.OLD_VN_AVAILABLE_COL;
    const ACCOUNT_TYPE_COL = c.ACCOUNT_TYPE_COL;
    const SUB_ACCT_COL = c.SUB_ACCT_COL;
    
    // ---------- INITIALISE COUNTERS ----------
    
    // Counts
    let totalVouchersRaised   = 0;  // count where ACCOUNT OR MAIL not empty
    let paidVoucherCount      = 0;  // count where STATUS = Paid
    let unpaidVoucherCount    = 0;  // count where STATUS = Unpaid
    let cancelledVoucherCount = 0;  // count where STATUS = Cancelled
    let revalidatedVouchers   = 0;  // count where old number exists OR availability = YES
    let revalidatedWithoutOldNumber = 0; // availability = YES and old number is empty
    
    // Amounts
    let totalContractSum      = 0;  // SUM(CONTRACT SUM)
    let totalPaidAmount       = 0;  // SUM(GROSS) where Paid
    let totalUnpaidAmount     = 0;  // SUM(GROSS) where Unpaid
    let totalCancelledAmount  = 0;  // SUM(GROSS) where Cancelled
    
    // Tax amounts
    let totalVAT = 0, totalWHT = 0, totalStampDuty = 0;
    let paidVAT = 0, paidWHT = 0, paidStampDuty = 0;
    
    // Category and monthly breakdown
    const categoryStats = {};   // { [category]: { vouchersRaised, amountPaid, balance, totalAmount } }
    const monthlyStats  = {};   // { [month]: { count, paidAmount, unpaidAmount } }
    const accountTypeStats = {}; // { [accountType]: { count, totalAmount, paidAmount, unpaidAmount, cancelledAmount, totalTax, paidTax } }
    
    const monthsList = [
      'January','February','March','April','May','June',
      'July','August','September','October','November','December'
    ];
    
    monthsList.forEach(m => {
      monthlyStats[m] = { count: 0, paidAmount: 0, unpaidAmount: 0 };
    });
    
    // ---------- PROCESS EACH ROW ----------
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // Basic values
      const rawStatus = String(row[STATUS_COL] || '').trim();
      const status    = rawStatus.toLowerCase();          // normalize to lower case
      const pmtMonth  = String(row[PMT_MONTH_COL] || '').trim();
      const payee     = String(row[PAYEE_COL]     || '').trim();
      const account   = String(row[ACCT_COL]      || '').trim();
      const category  = String(row[CATEGORY_COL]  || '').trim() || 'Uncategorized';
      const baseAccountType = String(row[ACCOUNT_TYPE_COL] || '').trim() || 'Unspecified';
      const subAccountType = String(row[SUB_ACCT_COL] || '').trim();
      const accountKey = subAccountType ? `${baseAccountType}::${subAccountType}` : `${baseAccountType}::`;
      
      const contractSum = parseAmount(row[CONTRACT_COL]);
      const grossAmount = parseAmount(row[GROSS_COL]);
      const oldVoucher  = String(row[OLD_VN_COL] || '').trim();
      const oldVoucherAvailable = String(row[OLD_VN_AVAILABLE_COL] || '').trim().toLowerCase();
      const hasOldVoucherNumber = !!oldVoucher;
      const markedOldVoucherAvailable = oldVoucherAvailable === 'yes';
      
      // Tax amounts
      const vat = parseAmount(row[VAT_COL]);
      const wht = parseAmount(row[WHT_COL]);
      const stampDuty = parseAmount(row[STAMP_COL]);
      
      // Skip completely empty rows
      if (!payee && !account && !contractSum && !grossAmount) continue;
      
      // Total Vouchers Raised = count if ACCOUNT OR MAIL not empty
      if (account) {
        totalVouchersRaised++;
      }
      
      // Total Contract Sum
      totalContractSum += contractSum;
      
      // Revalidated Vouchers
      if (hasOldVoucherNumber || markedOldVoucherAvailable) {
        revalidatedVouchers++;
        if (!hasOldVoucherNumber && markedOldVoucherAvailable) {
          revalidatedWithoutOldNumber++;
        }
      }
      
      // Tax accumulation (exclude cancelled)
      if (status !== 'cancelled') {
        totalVAT += vat;
        totalWHT += wht;
        totalStampDuty += stampDuty;
        
        if (status === 'paid') {
          paidVAT += vat;
          paidWHT += wht;
          paidStampDuty += stampDuty;
        }
      }
      
      // Paid / Unpaid / Cancelled - amounts & counts
      if (status === 'paid') {
        paidVoucherCount++;
        totalPaidAmount += grossAmount;
      } else if (status === 'unpaid') {
        unpaidVoucherCount++;
        totalUnpaidAmount += grossAmount;
      } else if (status === 'cancelled') {
        cancelledVoucherCount++;
        totalCancelledAmount += grossAmount;
      }
      // other statuses ignored for these sums

      // ----- Account type breakdown -----
      if (!accountTypeStats[accountKey]) {
        accountTypeStats[accountKey] = {
          count: 0,
          totalAmount: 0,
          paidAmount: 0,
          unpaidAmount: 0,
          cancelledAmount: 0,
          totalTax: 0,
          paidTax: 0
        };
      }

      const at = accountTypeStats[accountKey];
      at.count++;
      at.totalAmount += grossAmount;
      at.totalTax += (vat + wht + stampDuty);

      if (status === 'paid') {
        at.paidAmount += grossAmount;
        at.paidTax += (vat + wht + stampDuty);
      } else if (status === 'unpaid') {
        at.unpaidAmount += grossAmount;
      } else if (status === 'cancelled') {
        at.cancelledAmount += grossAmount;
      }
      
      // ----- Category breakdown -----
      if (!categoryStats[category]) {
        categoryStats[category] = {
          vouchersRaised: 0,
          amountPaid: 0,
          balance: 0,
          totalAmount: 0
        };
      }
      const cat = categoryStats[category];
      
      // Vouchers raised in this category (count of records with account)
      if (account) {
        cat.vouchersRaised++;
      }
      
      // Total amount in this category (all gross)
      cat.totalAmount += grossAmount;
      
      if (status === 'paid') {
        cat.amountPaid += grossAmount;
      } else if (status === 'unpaid') {
        cat.balance += grossAmount;
      }
      // Cancelled amounts reduce debt; not counted in balance
      
      // ----- Monthly breakdown -----
      if (pmtMonth && monthsList.includes(pmtMonth)) {
        monthlyStats[pmtMonth].count++;
        if (status === 'paid') {
          monthlyStats[pmtMonth].paidAmount += grossAmount;
        } else if (status === 'unpaid') {
          monthlyStats[pmtMonth].unpaidAmount += grossAmount;
        }
      }
    }
    
    // ---------- FINAL CALCULATIONS ----------
    
    // Total Debt = Total Contract Sum - (Paid + Cancelled)
    const totalDebt = totalContractSum - (totalPaidAmount + totalCancelledAmount);
    
    // Average Payment Rate = Paid vouchers (count) / Total raised (count)
    const averagePaymentPercent = totalVouchersRaised > 0
      ? ((paidVoucherCount / totalVouchersRaised) * 100).toFixed(2)
      : '0.00';
    
    // Category breakdown array
    const categoryBreakdown = Object.keys(categoryStats).map(catName => {
      const c = categoryStats[catName];
      
      const percentagePaid = c.totalAmount > 0
        ? ((c.amountPaid / c.totalAmount) * 100).toFixed(2)
        : '0.00';
      
      const percentOfTotalPayment = totalPaidAmount > 0
        ? ((c.amountPaid / totalPaidAmount) * 100).toFixed(2)
        : '0.00';
      
      return {
        category: catName,
        vouchersRaised: c.vouchersRaised,
        amountPaid: c.amountPaid,
        balance: c.balance,
        percentagePaid: parseFloat(percentagePaid),
        percentOfTotalPayment: parseFloat(percentOfTotalPayment)
      };
    });
    
    // Monthly breakdown array
    const monthlyBreakdown = monthsList.map(month => ({
      month: month,
      count: monthlyStats[month].count,
      paidAmount: monthlyStats[month].paidAmount,
      unpaidAmount: monthlyStats[month].unpaidAmount
    }));

    const accountTypeBreakdown = Object.keys(accountTypeStats)
      .map(accountKey => {
        const t = accountTypeStats[accountKey];
        const [baseType, subType] = accountKey.split('::');
        const paymentRate = t.totalAmount > 0 ? ((t.paidAmount / t.totalAmount) * 100) : 0;
        const taxComplianceRate = t.totalTax > 0 ? ((t.paidTax / t.totalTax) * 100) : 0;
        return {
          accountType: baseType,
          subAccountType: subType,
          displayName: subType ? `${baseType} (${subType})` : baseType,
          count: t.count,
          totalAmount: t.totalAmount,
          paidAmount: t.paidAmount,
          unpaidAmount: t.unpaidAmount,
          cancelledAmount: t.cancelledAmount,
          paymentRate: parseFloat(paymentRate.toFixed(2)),
          totalTax: t.totalTax,
          paidTax: t.paidTax,
          outstandingTax: t.totalTax - t.paidTax,
          taxComplianceRate: parseFloat(taxComplianceRate.toFixed(2))
        };
      })
      .sort((a, b) => b.totalAmount - a.totalAmount);
    
    return {
      success: true,
      year: year,
      summary: {
        // COUNTS
        totalVouchersRaised: totalVouchersRaised,
        paidVouchers:        paidVoucherCount,
        unpaidVouchers:      unpaidVoucherCount,
        cancelledVouchers:   cancelledVoucherCount,
        revalidatedVouchers: revalidatedVouchers,
        revalidatedWithoutOldNumber: revalidatedWithoutOldNumber,
        
        // AMOUNTS
        totalProcessedContractSum: totalContractSum,
        totalPaidAmount:           totalPaidAmount,
        totalUnpaidAmount:         totalUnpaidAmount,
        totalCancelledAmount:      totalCancelledAmount,
        totalDebt:                 totalDebt,
        
        // PERCENTAGE
        averagePaymentPercent:     parseFloat(averagePaymentPercent)
      },
      taxSummary: {
        totalVAT: totalVAT,
        totalWHT: totalWHT,
        totalStampDuty: totalStampDuty,
        totalTaxLiability: totalVAT + totalWHT + totalStampDuty,
        paidVAT: paidVAT,
        paidWHT: paidWHT,
        paidStampDuty: paidStampDuty,
        outstandingVAT: totalVAT - paidVAT,
        outstandingWHT: totalWHT - paidWHT,
        outstandingStampDuty: totalStampDuty - paidStampDuty
      },
      categoryBreakdown: categoryBreakdown,
      monthlyBreakdown: monthlyBreakdown,
      accountTypeBreakdown: accountTypeBreakdown
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to get summary: ' + error.message };
  }
}

/**
 * Gets debt profile by category
 * @param {string} token - Session token
 * @returns {Object} Debt profile data
 */
function getDebtProfile(token) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const data = sheet.getDataRange().getValues();
    
    const debtByCategory = {};
    const debtByPayee = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1]) continue;
      
      const status = String(row[CONFIG.VOUCHER_COLUMNS.STATUS - 1]).trim();
      
      // Only count unpaid vouchers as debt
      if (status !== 'Unpaid' && status !== 'Pending Deletion') continue;
      
      const payee = String(row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1]).trim();
      const category = String(row[CONFIG.VOUCHER_COLUMNS.CATEGORIES - 1]).trim() || 'Uncategorized';
      const grossAmount = parseAmount(row[CONFIG.VOUCHER_COLUMNS.GROSS_AMOUNT - 1]);
      
      // By category
      if (!debtByCategory[category]) {
        debtByCategory[category] = { amount: 0, count: 0 };
      }
      debtByCategory[category].amount += grossAmount;
      debtByCategory[category].count++;
      
      // By payee
      if (!debtByPayee[payee]) {
        debtByPayee[payee] = { amount: 0, count: 0 };
      }
      debtByPayee[payee].amount += grossAmount;
      debtByPayee[payee].count++;
    }
    
    // Convert to arrays and sort
    const categoryDebt = Object.keys(debtByCategory).map(cat => ({
      category: cat,
      amount: debtByCategory[cat].amount,
      count: debtByCategory[cat].count
    })).sort((a, b) => b.amount - a.amount);
    
    const payeeDebt = Object.keys(debtByPayee).map(payee => ({
      payee: payee,
      amount: debtByPayee[payee].amount,
      count: debtByPayee[payee].count
    })).sort((a, b) => b.amount - a.amount);
    
    // Top 10 debtors
    const topDebtors = payeeDebt.slice(0, 10);
    
    const totalDebt = categoryDebt.reduce((sum, c) => sum + c.amount, 0);
    
    return {
      success: true,
      totalDebt: totalDebt,
      debtByCategory: categoryDebt,
      topDebtors: topDebtors,
      allPayeeDebts: payeeDebt
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to get debt profile: ' + error.message };
  }
}

/**
 * Gets payment statistics for dashboard
 * OPTIMIZED - uses cache and limits recent vouchers
 * @param {string} token - Session token
 * @returns {Object} Dashboard stats
 */
function getDashboardStats(token) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Get 2026 summary (uses its own cache)
    const summary2026 = getSummary(token, '2026');
    
    if (!summary2026.success) {
      return summary2026;
    }
    
    // Get pending deletions count (quick query)
    let pendingCount = 0;
    try {
      const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const statusColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < statusColumn.length; i++) {
          if (statusColumn[i][0] === 'Pending Deletion') {
            pendingCount++;
          }
        }
      }
    } catch (e) {
      pendingCount = 0;
    }
    
    // Get only last 10 vouchers (not all vouchers)
    let recentVouchers = [];
    try {
      const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const startRow = Math.max(2, lastRow - 9); // Last 10 rows
        const numRows = lastRow - startRow + 1;
        const data = sheet.getRange(startRow, 1, numRows, 18).getValues();
        
        for (let i = data.length - 1; i >= 0; i--) {
          if (data[i][2]) { // Has payee
            recentVouchers.push(rowToVoucher(data[i], startRow + i, true));
          }
          if (recentVouchers.length >= 10) break;
        }
      }
    } catch (e) {
      recentVouchers = [];
    }
    
    return {
      success: true,
      stats: summary2026.summary,
      taxSummary: summary2026.taxSummary,  // ADD THIS LINE
      categoryBreakdown: summary2026.categoryBreakdown,
      monthlyBreakdown: summary2026.monthlyBreakdown,
      pendingDeletions: pendingCount,
      recentVouchers: recentVouchers,
      userRole: session.role,
      userName: session.name
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to get dashboard stats: ' + error.message };
  }
}

// ==================== TEST FUNCTIONS FOR PHASE 4 ====================

/**
 * Test summary function
 */
function testSummary() {
  const loginResult = handleLogin('admin@fmc.gov.ng', 'Admin@2026');
  
  if (!loginResult.success) {
    Logger.log('Login failed');
    return;
  }
  
  const result = getSummary(loginResult.token, '2026');
  Logger.log('Summary result: ' + JSON.stringify(result, null, 2));
}

/**
 * Test all years summary
 */
function testAllYearsSummary() {
  const loginResult = handleLogin('admin@fmc.gov.ng', 'Admin@2026');
  
  if (!loginResult.success) {
    Logger.log('Login failed');
    return;
  }
  
  const result = getAllYearsSummary(loginResult.token);
  Logger.log('All years summary: ' + JSON.stringify(result, null, 2));
}

/**
 * Test dashboard stats
 */
function testDashboardStats() {
  const loginResult = handleLogin('admin@fmc.gov.ng', 'Admin@2026');
  
  if (!loginResult.success) {
    Logger.log('Login failed');
    return;
  }
  
  const result = getDashboardStats(loginResult.token);
  Logger.log('Dashboard stats: ' + JSON.stringify(result, null, 2));
}

/**
 * Gets vouchers pending deletion (for Unit Head approval)
 * @param {string} token - Session token
 * @returns {Object} List of pending deletions
 */
function getPendingDeletions(token) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Only Unit Head and Admin can view pending deletions
    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: true, vouchers: [], count: 0 };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const data = sheet.getDataRange().getValues();
    
    const pendingVouchers = [];
    
    for (let i = 1; i < data.length; i++) {
      const status = data[i][CONFIG.VOUCHER_COLUMNS.STATUS - 1];
      
      if (status === 'Pending Deletion') {
        pendingVouchers.push(rowToVoucher(data[i], i + 1, true));
      }
    }
    
    return { 
      success: true, 
      vouchers: pendingVouchers,
      count: pendingVouchers.length
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to get pending deletions: ' + error.message };
  }
}

/**
 * Diagnostic function to check if all required functions exist
 */
function diagnosticCheck() {
  const functions = [
    'getSession',
    'hasPermission', 
    'getSheet',
    'parseAmount',
    'rowToVoucher',
    'voucherToRow',
    'getVouchers',
    'getSummary',
    'getPendingDeletions',
    'getDashboardStats',
    'handleLogin',
    'validateSession',
    'getNigerianTimestamp'
  ];
  
  const results = {};
  
  functions.forEach(funcName => {
    try {
      results[funcName] = typeof eval(funcName) === 'function' ? '✅ EXISTS' : '❌ NOT A FUNCTION';
    } catch (e) {
      results[funcName] = '❌ NOT DEFINED';
    }
  });
  
  Logger.log('=== DIAGNOSTIC RESULTS ===');
  for (const func in results) {
    Logger.log(func + ': ' + results[func]);
  }
  
  return results;
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[String(h).trim().toUpperCase()] = i + 1);
  return map;
}

/**
 * SMART LOOK-UP FEATURE
 * Parses voucher number to determine which year to search
 * Example: ABC/123/2025 → searches 2025 VOUCHERS first
 * Also searches OLD VOUCHER NUMBER column in all sheets
 * @param {string} token - Session token
 * @param {string} voucherNumber - The voucher number to search for
 * @returns {Object} Found voucher data or not found message
 */
function lookupVoucher(token, voucherNumber) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized: Cannot perform lookup' };
    }
    
    if (!voucherNumber || !voucherNumber.trim()) {
      return { success: false, error: 'Voucher Number is required for lookup' };
    }
    
    const searchValue = voucherNumber.trim().toUpperCase();
    
    // Determine which year to search based on voucher number suffix
    const yearFromVoucher = extractYearFromVoucherNumber(searchValue);
    
    // Build search order based on extracted year
    let sheetsToSearch = [];
    
    if (yearFromVoucher) {
      if (yearFromVoucher === '2025') {
        sheetsToSearch = [
          { name: CONFIG.SHEETS.VOUCHERS_2025, year: '2025', primary: true }
        ];
      } else if (yearFromVoucher === '2024') {
        sheetsToSearch = [
          { name: CONFIG.SHEETS.VOUCHERS_2024, year: '2024', primary: true }
        ];
      } else if (yearFromVoucher === '2023') {
        sheetsToSearch = [
          { name: CONFIG.SHEETS.VOUCHERS_2023, year: '2023', primary: true }
        ];
      } else if (['2022', '2021', '2020', '2019', '2018', '2017', '2016', '2015'].includes(yearFromVoucher)) {
        sheetsToSearch = [
          { name: CONFIG.SHEETS.VOUCHERS_BEFORE_2023, year: '<2023', primary: true }
        ];
      }
    } else {
      // No year detected, search all (most recent first)
      sheetsToSearch = [
        { name: CONFIG.SHEETS.VOUCHERS_2025, year: '2025', primary: true },
        { name: CONFIG.SHEETS.VOUCHERS_2024, year: '2024', primary: true },
        { name: CONFIG.SHEETS.VOUCHERS_2023, year: '2023', primary: true },
        { name: CONFIG.SHEETS.VOUCHERS_BEFORE_2023, year: '<2023', primary: true }
      ];
    }
    
    // STEP 1: Search in ACCOUNT_OR_MAIL column of the determined year sheet(s)
    for (const sheetInfo of sheetsToSearch) {
      try {
        const sheet = getSheet(sheetInfo.name);
        const lastRow = sheet.getLastRow();
        
        if (lastRow <= 1) continue;
        
        // Get Account/Mail column (Column D)
        const searchColumn = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
        
        for (let i = 0; i < searchColumn.length; i++) {
          const cellValue = String(searchColumn[i][0]).trim().toUpperCase();
          
          if (cellValue === searchValue) {
            // Found in Account/Mail column!
            const fullRow = sheet.getRange(i + 2, 1, 1, 17).getValues()[0];
            const voucher = rowToVoucher(fullRow, i + 2, false);
            
            // Check if voucher is PAID - cannot be revalidated
            if (voucher.status === 'Paid') {
              return {
                success: true,
                found: true,
                canRevalidate: false,
                voucher: voucher,
                sourceYear: sheetInfo.year,
                message: `Voucher found in ${sheetInfo.year} but CANNOT be revalidated because it is PAID.`,
                reason: 'PAID vouchers cannot be revalidated.'
              };
            }
            
            // Check if voucher is CANCELLED - needs authorization
            if (voucher.status === 'Cancelled') {
              return {
                success: true,
                found: true,
                canRevalidate: true,
                requiresAuthorization: true,
                voucher: voucher,
                sourceYear: sheetInfo.year,
                message: `Voucher found in ${sheetInfo.year}. Authorization required to revalidate CANCELLED voucher.`,
                warning: 'This voucher was CANCELLED. Unit Head approval is required to revalidate.'
              };
            }
            
            // Regular case - can revalidate
            return {
              success: true,
              found: true,
              canRevalidate: true,
              requiresAuthorization: false,
              voucher: voucher,
              sourceYear: sheetInfo.year,
              message: `Voucher found in ${sheetInfo.year} VOUCHERS`
            };
          }
        }
      } catch (sheetError) {
        console.log(`Could not search ${sheetInfo.name}: ${sheetError.message}`);
      }
    }
    
    // STEP 2: Search in OLD_VOUCHER_NUMBER column of ALL sheets (including 2026)
    const allSheets = [
      { name: CONFIG.SHEETS.VOUCHERS_2026, year: '2026' },
      { name: CONFIG.SHEETS.VOUCHERS_2025, year: '2025' },
      { name: CONFIG.SHEETS.VOUCHERS_2024, year: '2024' },
      { name: CONFIG.SHEETS.VOUCHERS_2023, year: '2023' },
      { name: CONFIG.SHEETS.VOUCHERS_BEFORE_2023, year: '<2023' }
    ];
    
    for (const sheetInfo of allSheets) {
      try {
        const sheet = getSheet(sheetInfo.name);
        const lastRow = sheet.getLastRow();
        
        if (lastRow <= 1) continue;
        
        // Get Old Voucher Number column (Column O = 15)
        const oldVNColumn = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
        
        for (let i = 0; i < oldVNColumn.length; i++) {
          const cellValue = String(oldVNColumn[i][0]).trim().toUpperCase();
          
          if (cellValue === searchValue) {
            // Found as Old Voucher Number - this means it was already revalidated!
            const has2026Format = sheetInfo.year === '2026';
            const numCols = has2026Format ? 18 : 17;
            const fullRow = sheet.getRange(i + 2, 1, 1, numCols).getValues()[0];
            const voucher = rowToVoucher(fullRow, i + 2, has2026Format);
            
            return {
              success: true,
              found: true,
              canRevalidate: false,
              alreadyRevalidated: true,
              voucher: voucher,
              sourceYear: sheetInfo.year,
              message: `This voucher was already revalidated in ${sheetInfo.year}.`,
              warning: 'Cannot revalidate again. The voucher already exists in the system.',
              existingVoucherNumber: voucher.accountOrMail
            };
          }
        }
      } catch (sheetError) {
        console.log(`Could not search old VN in ${sheetInfo.name}: ${sheetError.message}`);
      }
    }
    
    // Not found in any sheet
    return {
      success: true,
      found: false,
      message: 'Voucher not found in any records',
      searchedYears: sheetsToSearch.map(s => s.year),
      hint: yearFromVoucher 
        ? `Searched in ${yearFromVoucher} based on voucher number format.`
        : 'Searched in all available years.'
    };
    
  } catch (error) {
    return { success: false, error: 'Lookup failed: ' + error.message };
  }
}

/**
 * Extracts year from voucher number
 * Examples: "ABC/123/2025" → "2025", "XYZ-456-2024" → "2024"
 * @param {string} voucherNumber - The voucher number
 * @returns {string|null} Extracted year or null
 */
function extractYearFromVoucherNumber(voucherNumber) {
  if (!voucherNumber) return null;
  
  // Look for 4-digit year pattern at the end
  const patterns = [
    /\/(\d{4})$/,      // Ends with /2025
    /-(\d{4})$/,       // Ends with -2025
    /\.(\d{4})$/,      // Ends with .2025
    /\s(\d{4})$/,      // Ends with space 2025
    /(\d{4})$/         // Just ends with 4 digits
  ];
  
  for (const pattern of patterns) {
    const match = voucherNumber.match(pattern);
    if (match) {
      const year = match[1];
      // Validate it's a reasonable year (2015-2025)
      const yearNum = parseInt(year);
      if (yearNum >= 2015 && yearNum <= 2025) {
        return year;
      }
    }
  }
  
  return null;
}

/**
 * Assigns control number to multiple vouchers
 * Only Payable Unit can do this
 * @param {string} token - Session token
 * @param {Array} rowIndexes - Array of row indexes to update
 * @param {string} controlNumber - Control number to assign
 * @returns {Object} Result
 */
function assignControlNumber(token, rowIndexes, controlNumber) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };

    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized: Only Payable Unit can assign control numbers' };
    }

    if (!controlNumber || !controlNumber.trim()) {
      return { success: false, error: 'Control Number is required' };
    }

    if (!rowIndexes || !Array.isArray(rowIndexes) || rowIndexes.length === 0) {
      return { success: false, error: 'No vouchers selected' };
    }

    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const header = getHeaderMap_(sheet);
    const colReleasedAt = header["RELEASED AT"] || null;

    const cn = controlNumber.trim();
    let updatedCount = 0;

    for (const rowIndex of rowIndexes) {
      if (rowIndex >= 2 && rowIndex <= sheet.getLastRow()) {
        // ✅ set CN + RELEASED AT (only once)
        setControlNoAndReleasedAt_(sheet, rowIndex, cn, colReleasedAt);
        updatedCount++;
      }
    }

    return {
      success: true,
      message: `Control number "${cn}" assigned to ${updatedCount} voucher(s)`,
      count: updatedCount
    };

  } catch (error) {
    return { success: false, error: 'Failed to assign control number: ' + error.message };
  }
}

// ==================== QUICK STATS (Lightweight) ====================

/**
 * Gets just the counts - very fast for dashboard header
 * Use this instead of full getDashboardStats for initial page load
 */
function getQuickStats(token) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Check cache first (cache for 2 minutes)
    const cache = CacheService.getScriptCache();
    const cached = cache.get('quick_stats_2026');
    
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {}
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: true,
        totalVouchers: 0,
        paidCount: 0,
        unpaidCount: 0,
        cancelledCount: 0
      };
    }
    
    // Only get status column (Column A)
    const statusColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    let paid = 0, unpaid = 0, cancelled = 0, pending = 0;
    
    for (let i = 0; i < statusColumn.length; i++) {
      const status = String(statusColumn[i][0]).trim();
      switch(status) {
        case 'Paid': paid++; break;
        case 'Cancelled': cancelled++; break;
        case 'Pending Deletion': pending++; break;
        default: unpaid++; break;
      }
    }
    
    const result = {
      success: true,
      totalVouchers: statusColumn.length,
      paidCount: paid,
      unpaidCount: unpaid,
      cancelledCount: cancelled,
      pendingCount: pending
    };
    
    // Cache for 2 minutes
    try {
      cache.put('quick_stats_2026', JSON.stringify(result), 120);
    } catch (e) {}
    
    return result;
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Preload data into cache (run this periodically via trigger)
 */
function preloadCache() {
  try {
    // Create a dummy session for internal use
    const adminEmail = 'admin@fmc.gov.ng';
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === adminEmail && data[i][3] === 'ADMIN') {
        const token = generateToken(adminEmail);
        const session = {
          name: data[i][0],
          email: adminEmail,
          role: 'ADMIN'
        };
        storeSession(token, session);
        
        // Preload summary
        getSummary(token, '2026');
        
        // Clear the temporary session
        clearSession(token);
        
        Logger.log('Cache preloaded successfully');
        break;
      }
    }
  } catch (error) {
    Logger.log('Cache preload error: ' + error.message);
  }
}

/**
 * Release vouchers with control number assignment and notifications
 * @param {string} token - Session token
 * @param {Array} rowIndexes - Array of row indexes to release
 * @param {string} controlNumber - Control number to assign
 * @param {string} targetUnit - Target unit (CPO, Audit, etc.)
 * @returns {Object} Result
 */
function releaseVouchersWithNotification(token, rowIndexes, controlNumber, targetUnit) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Only Payable Unit can release
    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized: Only Payable Unit can release vouchers' };
    }
    
    if (!controlNumber || !controlNumber.trim()) {
      return { success: false, error: 'Control Number is required' };
    }
    
    if (!rowIndexes || !Array.isArray(rowIndexes) || rowIndexes.length === 0) {
      return { success: false, error: 'No vouchers selected' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const cn = controlNumber.trim();
    let updatedCount = 0;
    const releasedVouchers = [];
    
    for (const rowIndex of rowIndexes) {
      if (rowIndex >= 2 && rowIndex <= sheet.getLastRow()) {
        // Update control number
        const header = getHeaderMap_(sheet);
          const colReleasedAt = header["RELEASED AT"] || null;
          // ...
          setControlNoAndReleasedAt_(sheet, rowIndex, cn, colReleasedAt);
        
        // Get voucher details for notification
        const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
        releasedVouchers.push({
          voucherNumber: row[3], // Account/Mail
          payee: row[2],
          amount: row[6]
        });
        
        updatedCount++;
      }
    }
    
    // Clear cache
    clearAllVoucherCaches();
    
    // Send notification emails
    sendReleaseNotification(cn, targetUnit, releasedVouchers, session.name);
    
    return {
      success: true,
      message: `${updatedCount} voucher(s) released with Control Number: ${cn}. Notifications sent.`,
      count: updatedCount
    };
    
  } catch (error) {
    return { success: false, error: 'Release failed: ' + error.message };
  }
}

/**
 * Send release notification to relevant parties
 */
function sendReleaseNotification(controlNumber, targetUnit, vouchers, releasedBy) {
  try {
    const usersSheet = getSheet(CONFIG.SHEETS.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    
    // Find recipients: Admin, Payable Head, DDFA, DFA
    const recipientRoles = ['ADMIN', 'Payable Unit Head', 'DDFA', 'DFA'];
    const recipients = [];
    
    for (let i = 1; i < usersData.length; i++) {
      const role = usersData[i][3];
      const email = usersData[i][1];
      const active = usersData[i][4];
      
      if (recipientRoles.includes(role) && active && email) {
        recipients.push(email);
      }
    }
    
    if (recipients.length === 0) return;
    
    // Calculate total amount
    let totalAmount = 0;
    vouchers.forEach(v => {
      totalAmount += parseFloat(v.amount) || 0;
    });
    
    // Create voucher list HTML
    let voucherList = '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%;">';
    voucherList += '<tr style="background: #f4f4f4;"><th>Voucher No.</th><th>Payee</th><th>Amount</th></tr>';
    vouchers.forEach(v => {
      voucherList += `<tr><td>${v.voucherNumber}</td><td>${v.payee}</td><td>₦${Number(v.amount).toLocaleString()}</td></tr>`;
    });
    voucherList += `<tr style="background: #e8f5e9; font-weight: bold;"><td colspan="2">Total</td><td>₦${Number(totalAmount).toLocaleString()}</td></tr>`;
    voucherList += '</table>';
    
    // Email content
    const subject = `[PAYABLE VOUCHER] Vouchers Released - Control No: ${controlNumber}`;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #1a5f2a; color: white; padding: 20px; text-align: center;">
          <h2 style="margin: 0;">Voucher Release Notification</h2>
        </div>
        
        <div style="padding: 20px; background: #fff;">
          <p>Dear Team,</p>
          
          <p>The following vouchers have been released for processing:</p>
          
          <div style="background: #f9f9f9; padding: 15px; border-radius: 5px; margin: 15px 0;">
            <p><strong>Control Number:</strong> ${controlNumber}</p>
            <p><strong>Released To:</strong> ${targetUnit}</p>
            <p><strong>Released By:</strong> ${releasedBy}</p>
            <p><strong>Number of Vouchers:</strong> ${vouchers.length}</p>
            <p><strong>Total Amount:</strong> ₦${Number(totalAmount).toLocaleString()}</p>
          </div>
          
          <h4>Voucher Details:</h4>
          ${voucherList}
          
          <p style="margin-top: 20px;">Please take necessary action.</p>
          
          <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
          
          <p style="font-size: 12px; color: #666;">
            This is an automated notification from the Payable Voucher System.<br>
            Federal Medical Centre, Abeokuta - Finance & Accounts Department<br>
            <em>Powered by ABLEBIZ @ hello@ablebiz.com.ng</em>
          </p>
        </div>
      </div>
    `;
    
    // Send email
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      htmlBody: htmlBody
    });
    
  } catch (error) {
    console.log('Notification error: ' + error.message);
    // Don't fail the release if notification fails
  }
}

/**
 * Create notifications for given users
 */
function createNotifications(recipientEmails, title, message, link) {
  try {
    if (!recipientEmails || recipientEmails.length === 0) return;
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('NOTIFICATIONS');
    if (!sheet) {
      sheet = ss.insertSheet('NOTIFICATIONS');
      sheet.getRange(1,1,1,6).setValues([['TIMESTAMP','USER_EMAIL','TITLE','MESSAGE','LINK','READ']]);
    }
    
    const ts = getNigerianTimestamp();
    const rows = recipientEmails.map(email => [
      ts,
      String(email || '').toLowerCase(),
      title,
      message,
      link || '',
      false
    ]);
    
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    
  } catch (e) {
    console.log('createNotifications error: ' + e.message);
  }
}

/**
 * Get notifications for current user
 */
function getNotifications(token, onlyUnread) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };
    
    const email = String(session.email || '').toLowerCase();
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('NOTIFICATIONS');
    if (!sheet) return { success: true, notifications: [], unreadCount: 0 };
    
    const data = sheet.getDataRange().getValues();
    const notifications = [];
    let unreadCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = String(row[1] || '').toLowerCase();
      if (rowEmail !== email) continue;
      
      const read = row[5] === true || String(row[5]).toUpperCase() === 'TRUE';
      if (!read) unreadCount++;
      if (onlyUnread && read) continue;
      
      notifications.push({
        rowIndex: i + 1,
        timestamp: row[0],
        title: row[2],
        message: row[3],
        link: row[4],
        read: read
      });
    }
    
    return { success: true, notifications: notifications, unreadCount: unreadCount };
    
  } catch (e) {
    return { success: false, error: 'Failed to get notifications: ' + e.message };
  }
}

/**
 * Mark notification as read
 */
function markNotificationRead(token, rowIndex) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('NOTIFICATIONS');
    if (!sheet) return { success: false, error: 'No notifications sheet' };
    
    const lastRow = sheet.getLastRow();
    rowIndex = parseInt(rowIndex, 10);
    if (rowIndex < 2 || rowIndex > lastRow) {
      return { success: false, error: 'Invalid notification reference' };
    }
    
    sheet.getRange(rowIndex, 6).setValue(true); // READ column
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: 'Failed to mark notification read: ' + e.message };
  }
}

/**
 * Gets summary across all years for debt tracking
 * @param {string} token - Session token
 * @returns {Object} All years summary with grand totals
 */
function getAllYearsSummary(token) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Roles allowed to view summary
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
    
    const years = [
      { label: '<2023', sheet: CONFIG.SHEETS.VOUCHERS_BEFORE_2023 },
      { label: '2023', sheet: CONFIG.SHEETS.VOUCHERS_2023 },
      { label: '2024', sheet: CONFIG.SHEETS.VOUCHERS_2024 },
      { label: '2025', sheet: CONFIG.SHEETS.VOUCHERS_2025 },
      { label: '2026', sheet: CONFIG.SHEETS.VOUCHERS_2026 }
    ];
    
    const yearsSummary = [];
    let grandTotalVouchers = 0;
    let grandTotalAmount = 0;
    let grandTotalPaid = 0;
    let grandTotalRevalidated = 0;
    let runningBalance = 0;
    
    for (const yearInfo of years) {
      try {
        const sheet = getSheet(yearInfo.sheet);
        const lastRow = sheet.getLastRow();
        
        if (lastRow <= 1) {
          yearsSummary.push({
            label: yearInfo.label,
            balanceBroughtForward: runningBalance,
            totalVouchers: 0,
            totalAmount: 0,
            paidAmount: 0,
            revalidatedVouchers: 0,
            cancelledVouchers: 0,
            currentBalance: runningBalance
          });
          continue;
        }
        
        const has2026Format = yearInfo.label === '2026';
        let statusCol = 0;
        let accountCol = 3;
        let grossCol = 6;
        let oldVNCol = 14;
        let oldAvailCol = 20;
        let numCols = 17;

        if (has2026Format) {
          const c = resolveVoucherSummaryColumns_(sheet);
          statusCol = c.STATUS_COL;
          accountCol = c.ACCT_COL;
          grossCol = c.GROSS_COL;
          oldVNCol = c.OLD_VN_COL;
          oldAvailCol = c.OLD_VN_AVAILABLE_COL;
          numCols = c.LAST_COL;
        }

        const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
        
        let totalVouchers = 0;
        let totalAmount = 0;
        let paidAmount = 0;
        let unpaidAmount = 0;
        let revalidatedCount = 0;
        let cancelledCount = 0;
        
        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          const status = String(row[statusCol] || '').trim().toLowerCase();
          const account = String(row[accountCol] || '').trim();
          const grossAmount = parseAmount(row[grossCol]);
          const oldVN = String(row[oldVNCol] || '').trim();
          const oldVoucherAvailable = has2026Format ? String(row[oldAvailCol] || '').trim().toLowerCase() : '';
          
          if (!account && !grossAmount) continue;
          
          if (account) totalVouchers++;
          totalAmount += grossAmount;
          
          if (status === 'paid') {
            paidAmount += grossAmount;
          } else if (status === 'cancelled') {
            cancelledCount++;
          } else {
            unpaidAmount += grossAmount;
          }
          
          if (oldVN || oldVoucherAvailable === 'yes') revalidatedCount++;
        }
        
        const balanceBF = runningBalance;
        const currentBalance = balanceBF + unpaidAmount;
        runningBalance = currentBalance;
        
        yearsSummary.push({
          label: yearInfo.label,
          balanceBroughtForward: balanceBF,
          totalVouchers: totalVouchers,
          totalAmount: totalAmount,
          paidAmount: paidAmount,
          revalidatedVouchers: revalidatedCount,
          cancelledVouchers: cancelledCount,
          currentBalance: currentBalance
        });
        
        grandTotalVouchers += totalVouchers;
        grandTotalAmount += totalAmount;
        grandTotalPaid += paidAmount;
        grandTotalRevalidated += revalidatedCount;
        
      } catch (e) {
        yearsSummary.push({
          label: yearInfo.label,
          error: 'Could not load data: ' + e.message
        });
      }
    }
    
    return {
      success: true,
      yearsSummary: yearsSummary,
      grandTotals: {
        totalVouchers: grandTotalVouchers,
        totalAmount: grandTotalAmount,
        totalPaid: grandTotalPaid,
        totalRevalidated: grandTotalRevalidated,
        currentOutstandingBalance: runningBalance
      }
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to get all years summary: ' + error.message };
  }
}

/**
 * Release vouchers with full workflow support
 * Supports both Payable Unit and CPO release flows
 */
function releaseVouchers(data) {
  try {
    const session = getSession(data.token);
    if (!session) return { success: false, error: 'Session expired' };

    const role = session.role;
    const rowIndexes = data.rowIndexes;
    const controlNumber = (data.controlNumber || '').trim();
    const targetUnit = (data.targetUnit || '').trim();
    const purpose = (data.purpose || '').trim();
    const isCPORelease = !!data.isCPORelease;

    // Validate permissions
    const allowedRoles = [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.ADMIN
    ];

    if (!hasPermission(role, allowedRoles)) {
      return { success: false, error: 'Unauthorized to release vouchers' };
    }

    if (!rowIndexes || !Array.isArray(rowIndexes) || rowIndexes.length === 0) {
      return { success: false, error: 'No vouchers selected' };
    }

    if (!targetUnit) {
      return { success: false, error: 'Target unit is required' };
    }

    // For Payable Unit, control number is required
    if (!isCPORelease && !controlNumber) {
      return { success: false, error: 'Control number is required' };
    }

    // For CPO, purpose is required
    if (isCPORelease && !purpose) {
      return { success: false, error: 'Purpose is required for CPO release' };
    }

    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();

    // ✅ compute RELEASED AT column once (not inside loop)
    const header = getHeaderMap_(sheet);
    const colReleasedAt = header["RELEASED AT"] || null;

    let releasedVouchers = [];
    let totalAmount = 0;
    let skippedCount = 0;

    for (const r of rowIndexes) {
      const rowIndex = parseInt(r, 10);
      if (rowIndex < 2 || rowIndex > lastRow) continue;

      const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
      const currentCN = String(row[CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER - 1] || '').trim();

      // Check for double release (for non-CPO)
      if (!isCPORelease && currentCN !== '') {
        // Already has CN - only Head/Admin can proceed
        if (role !== CONFIG.ROLES.PAYABLE_HEAD && role !== CONFIG.ROLES.ADMIN) {
          skippedCount++;
          continue;
        }
      }

      // ✅ Update control number only when appropriate
      // - Payable release: set CN (required)
      // - CPO release: typically you don't assign CN here; you may be sending onward.
      if (!isCPORelease || !currentCN) {
        if (controlNumber) {
          setControlNoAndReleasedAt_(sheet, rowIndex, controlNumber, colReleasedAt);
        }
      }

      const amount = parseAmount(row[CONFIG.VOUCHER_COLUMNS.GROSS_AMOUNT - 1]);

      const voucherInfo = {
        rowIndex: rowIndex,
        voucherNumber: row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1],
        payee: row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1],
        amount: amount
      };

      releasedVouchers.push(voucherInfo);
      totalAmount += amount;

      // Log audit
      logAudit(
        session,
        'RELEASE_VOUCHER',
        `Released to ${targetUnit}${purpose ? ' - ' + purpose : ''}`,
        CONFIG.SHEETS.VOUCHERS_2026,
        rowIndex,
        { controlNumber: controlNumber || currentCN, targetUnit: targetUnit }
      );
    }

    if (releasedVouchers.length === 0) {
      return { success: false, error: 'No vouchers were released. They may already be assigned.' };
    }

    // Clear cache (recommended after updates)
    if (typeof clearAllVoucherCaches === 'function') {
      clearAllVoucherCaches();
    }

    // Send notification email
    sendReleaseNotification(
      controlNumber || 'N/A (CPO Release)',
      targetUnit,
      releasedVouchers,
      session.name,
      purpose
    );

    // Create in-app notifications
    const notifTitle = isCPORelease
      ? `CPO Released Vouchers to ${targetUnit}`
      : `Vouchers Released: ${controlNumber}`;

    const notifMessage =
      `${releasedVouchers.length} voucher(s) totaling ${formatCurrencyForNotif(totalAmount)} released to ${targetUnit}`;

    const recipientRoles = ['ADMIN', 'Payable Unit Head', 'DDFA', 'DFA'];
    const usersSheet = getSheet(CONFIG.SHEETS.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    const recipientEmails = [];

    for (let i = 1; i < usersData.length; i++) {
      if (recipientRoles.includes(usersData[i][3]) && usersData[i][4]) {
        recipientEmails.push(usersData[i][1]);
      }
    }

    createNotifications(recipientEmails, notifTitle, notifMessage, 'vouchers.html');

    let message = `${releasedVouchers.length} voucher(s) released to ${targetUnit}`;
    if (skippedCount > 0) {
      message += `. ${skippedCount} voucher(s) skipped (already released).`;
    }

    return {
      success: true,
      message: message,
      count: releasedVouchers.length,
      controlNumber: controlNumber,
      targetUnit: targetUnit,
      totalAmount: totalAmount
    };

  } catch (error) {
    return { success: false, error: 'Release failed: ' + error.message };
  }
}

/**
 * Helper function to format currency for notifications
 */
function formatCurrencyForNotif(amount) {
    return '₦' + Number(amount).toLocaleString('en-NG', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Test if updateVoucherStatus function exists
 * Run this from the Apps Script editor
 */
function testUpdateVoucherStatus() {
  Logger.log('Testing updateVoucherStatus...');
  
  // Check if function exists
  if (typeof updateVoucherStatus === 'function') {
    Logger.log('✅ updateVoucherStatus function EXISTS');
  } else {
    Logger.log('❌ updateVoucherStatus function NOT FOUND');
  }
  
  // List all functions
  Logger.log('--- All global functions ---');
  for (var name in this) {
    if (typeof this[name] === 'function') {
      Logger.log('Function: ' + name);
    }
  }
}

/**
 * Approve and execute voucher deletion
 * - Payable Head & Admin can approve for normal vouchers
 * - For PAID or RELEASED vouchers, only CPO or Admin can approve
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
    
    if (rowIndex < 2 || rowIndex > lastRow) {
      return { success: false, error: 'Voucher not found' };
    }
    
    // Get the full row data
    const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
    const currentStatus = String(row[CONFIG.VOUCHER_COLUMNS.STATUS - 1] || '').trim();
    const controlNumber = String(row[CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER - 1] || '').trim();
    const payee = row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1];
    const voucherNum = row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1];
    
    // Check if status is Pending Deletion
    if (currentStatus !== 'Pending Deletion') {
      return { success: false, error: 'Voucher is not marked as Pending Deletion.' };
    }
    
    // For RELEASED vouchers (has control number), only CPO or ADMIN can approve
    if (controlNumber !== '') {
      if (role !== CONFIG.ROLES.CPO && role !== CONFIG.ROLES.ADMIN) {
        return { success: false, error: 'Only CPO or Admin can approve deletion of released vouchers.' };
      }
    } else {
      // For non-released vouchers, Payable Head and Admin can approve
      if (role !== CONFIG.ROLES.PAYABLE_HEAD && role !== CONFIG.ROLES.ADMIN) {
        return { success: false, error: 'Only Payable Head or Admin can approve deletion.' };
      }
    }
    
    // Delete the row
    sheet.deleteRow(rowIndex);
    
    // Clear cache
    clearAllVoucherCaches();
    
    // Log audit
    logAudit(session, 'APPROVE_DELETE', 'Approved and deleted voucher: ' + voucherNum, CONFIG.SHEETS.VOUCHERS_2026, rowIndex, {
      payee: payee,
      voucherNumber: voucherNum
    });
    
    return { success: true, message: 'Voucher deleted successfully.' };
    
  } catch (error) {
    return { success: false, error: 'Delete approval failed: ' + error.message };
  }
}

// ==================== NOTIFICATION SYSTEM ====================

/**
 * Creates notifications for specified users
 * @param {Array} recipientEmails - Array of email addresses
 * @param {string} title - Notification title
 * @param {string} message - Notification message
 * @param {string} link - Link to relevant page
 * @param {string} type - Notification type (info, warning, success, danger)
 */
function createNotifications(recipientEmails, title, message, link, type) {
  try {
    if (!recipientEmails || recipientEmails.length === 0) return;
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('NOTIFICATIONS');
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('NOTIFICATIONS');
      sheet.getRange(1, 1, 1, 7).setValues([[
        'TIMESTAMP', 'USER_EMAIL', 'TITLE', 'MESSAGE', 'LINK', 'TYPE', 'READ'
      ]]);
      // Format header
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f0f0f0');
    }
    
    const ts = getNigerianTimestamp();
    const rows = recipientEmails.map(email => [
      ts,
      String(email || '').toLowerCase().trim(),
      title || '',
      message || '',
      link || '',
      type || 'info',
      false
    ]);
    
    // Append all notifications
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    }
    
  } catch (e) {
    console.log('createNotifications error: ' + e.message);
  }
}

/**
 * Get notifications for current user
 * @param {string} token - Session token
 * @param {boolean} onlyUnread - If true, return only unread notifications
 */
function getNotifications(token, onlyUnread) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };
    
    const email = String(session.email || '').toLowerCase().trim();
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('NOTIFICATIONS');
    
    if (!sheet) {
      return { success: true, notifications: [], unreadCount: 0 };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, notifications: [], unreadCount: 0 };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const notifications = [];
    let unreadCount = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowEmail = String(row[1] || '').toLowerCase().trim();
      
      if (rowEmail !== email) continue;
      
      const isRead = row[6] === true || String(row[6]).toUpperCase() === 'TRUE';
      
      if (!isRead) unreadCount++;
      if (onlyUnread && isRead) continue;
      
      notifications.push({
        rowIndex: i + 2, // 1-based, skip header
        timestamp: row[0],
        title: row[2],
        message: row[3],
        link: row[4],
        type: row[5] || 'info',
        read: isRead
      });
    }
    
    // Sort by timestamp descending (newest first)
    notifications.sort((a, b) => {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });
    
    return { 
      success: true, 
      notifications: notifications, 
      unreadCount: unreadCount 
    };
    
  } catch (e) {
    return { success: false, error: 'Failed to get notifications: ' + e.message };
  }
}

/**
 * Mark notification as read
 */
function markNotificationRead(token, rowIndex) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('NOTIFICATIONS');
    
    if (!sheet) return { success: false, error: 'Notifications not found' };
    
    const lastRow = sheet.getLastRow();
    rowIndex = parseInt(rowIndex, 10);
    
    if (rowIndex < 2 || rowIndex > lastRow) {
      return { success: false, error: 'Invalid notification' };
    }
    
    // Verify this notification belongs to the user
    const email = sheet.getRange(rowIndex, 2).getValue();
    if (String(email).toLowerCase() !== session.email.toLowerCase()) {
      return { success: false, error: 'Unauthorized' };
    }
    
    sheet.getRange(rowIndex, 7).setValue(true); // READ column
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: 'Failed to mark notification: ' + e.message };
  }
}

/**
 * Mark all notifications as read for current user
 */
function markAllNotificationsRead(token) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };
    
    const email = String(session.email || '').toLowerCase().trim();
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('NOTIFICATIONS');
    
    if (!sheet) return { success: true, message: 'No notifications' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, message: 'No notifications' };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    let count = 0;
    
    for (let i = 0; i < data.length; i++) {
      const rowEmail = String(data[i][1] || '').toLowerCase().trim();
      const isRead = data[i][6] === true || String(data[i][6]).toUpperCase() === 'TRUE';
      
      if (rowEmail === email && !isRead) {
        sheet.getRange(i + 2, 7).setValue(true);
        count++;
      }
    }
    
    return { success: true, message: `Marked ${count} notification(s) as read` };
    
  } catch (e) {
    return { success: false, error: 'Failed: ' + e.message };
  }
}

/**
 * Get pending deletions for approval
 */
function getPendingDeletions(token) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    // Only Unit Head, CPO, and Admin can view pending deletions
    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: true, vouchers: [], count: 0 };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: true, vouchers: [], count: 0 };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
    const pendingVouchers = [];
    
    for (let i = 0; i < data.length; i++) {
      const status = String(data[i][CONFIG.VOUCHER_COLUMNS.STATUS - 1] || "")
        .trim().toUpperCase();

      let cn = String(data[i][CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER - 1] || "").trim();
      cn = (cn === "-" || cn.toUpperCase() === "N/A") ? "" : cn; // optional cleanup
      
      if (status === 'Pending Deletion') {
        const voucher = rowToVoucher(data[i], i + 2, true);
        pendingVouchers.push(voucher);
      }
    }
    
    return { 
      success: true, 
      vouchers: pendingVouchers,
      count: pendingVouchers.length
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to get pending deletions: ' + error.message };
  }
}

/**
 * Reject voucher deletion request
 */
function rejectVoucherDelete(token, rowIndex, reason) {
  try {
    const session = getSession(token);
    if (!session) {
      return { success: false, error: 'Session expired' };
    }
    
    const role = session.role;
    
    // Only Payable Head, CPO, and Admin can reject
    if (!hasPermission(role, [
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized to reject deletion requests' };
    }
    
    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
    const lastRow = sheet.getLastRow();
    
    if (rowIndex < 2 || rowIndex > lastRow) {
      return { success: false, error: 'Voucher not found' };
    }
    
    const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
    const currentStatus = String(row[CONFIG.VOUCHER_COLUMNS.STATUS - 1] || '').trim();
    
    if (currentStatus !== 'Pending Deletion') {
      return { success: false, error: 'Voucher is not pending deletion' };
    }
    
    // Revert status to Unpaid
    sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS).setValue('Unpaid');
    
    // Log audit
    logAudit(session, 'REJECT_DELETE', 'Rejected deletion: ' + (reason || 'No reason provided'), 
             CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
    
    // Notify the requester (if we tracked who requested - for now notify all staff)
    const voucherNum = row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1];
    const payee = row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1];
    
    // Get Payable Staff emails
    const usersSheet = getSheet(CONFIG.SHEETS.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    const staffEmails = [];
    
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][3] === CONFIG.ROLES.PAYABLE_STAFF && usersData[i][4]) {
        staffEmails.push(usersData[i][1]);
      }
    }
    
    createNotifications(
      staffEmails,
      'Deletion Request Rejected',
      `Deletion of voucher ${voucherNum} (${payee}) was rejected. Reason: ${reason || 'Not specified'}`,
      'vouchers.html',
      'warning'
    );
    
    return { success: true, message: 'Deletion request rejected. Status reverted to Unpaid.' };
    
  } catch (error) {
    return { success: false, error: 'Failed to reject: ' + error.message };
  }
}

/**
 * Reject voucher deletion request
 * @param {string} token - Session token
 * @param {number} rowIndex - Row index of voucher
 * @param {string} reason - Reason for rejection
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
    const row = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
    const currentStatus = String(row[CONFIG.VOUCHER_COLUMNS.STATUS - 1] || '').trim();
    const voucherNum = row[CONFIG.VOUCHER_COLUMNS.ACCOUNT_OR_MAIL - 1] || '';
    const payee = row[CONFIG.VOUCHER_COLUMNS.PAYEE - 1] || '';
    
    if (currentStatus !== 'Pending Deletion') {
      return { success: false, error: 'Voucher is not pending deletion' };
    }
    
    // Revert status to Unpaid
    sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS).setValue('Unpaid');
    
    // Clear cache
    clearAllVoucherCaches();
    
    // Log audit
    logAudit(session, 'REJECT_DELETE', 
             'Rejected deletion of ' + voucherNum + '. Reason: ' + (reason || 'Not specified'), 
             CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
    
    // Send notification to Payable Staff about rejection
    try {
      const usersSheet = getSheet(CONFIG.SHEETS.USERS);
      const usersData = usersSheet.getDataRange().getValues();
      const staffEmails = [];
      
      for (let i = 1; i < usersData.length; i++) {
        const userRole = usersData[i][CONFIG.USER_COLUMNS.ROLE - 1];
        const email = usersData[i][CONFIG.USER_COLUMNS.EMAIL - 1];
        const active = usersData[i][CONFIG.USER_COLUMNS.ACTIVE - 1];
        
        if (userRole === CONFIG.ROLES.PAYABLE_STAFF && active && email) {
          staffEmails.push(email);
        }
      }
      
      if (staffEmails.length > 0) {
        createNotifications(
          staffEmails,
          '❌ Deletion Request Rejected',
          'Deletion of voucher ' + voucherNum + ' (' + payee + ') was rejected by ' + session.name + '. Reason: ' + (reason || 'Not specified'),
          'vouchers.html',
          'warning'
        );
      }
    } catch (notifError) {
      console.log('Notification error: ' + notifError.message);
      // Don't fail the main operation if notification fails
    }
    
    return { 
      success: true, 
      message: 'Deletion request rejected. Voucher status reverted to Unpaid.' 
    };
    
  } catch (error) {
    console.log('rejectVoucherDelete error: ' + error.message);
    return { success: false, error: 'Failed to reject deletion: ' + error.message };
  }
}




function getMyProfile(token) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };

  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const rowIndex = session.rowIndex;

  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  return {
    success: true,
    profile: {
      // Admin-only fields (read-only for user)
      name: row[CONFIG.USER_COLUMNS.NAME - 1],
      email: row[CONFIG.USER_COLUMNS.EMAIL - 1],
      username: row[CONFIG.USER_COLUMNS.USERNAME - 1] || '',
      role: row[CONFIG.USER_COLUMNS.ROLE - 1],

      // User-editable fields
      phone: row[CONFIG.USER_COLUMNS.PHONE - 1] || '',
      department: row[CONFIG.USER_COLUMNS.DEPARTMENT - 1] || ''
    }
  };
}

function updateMyProfile(token, profile) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };

  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const rowIndex = session.rowIndex;

  // Only allow user-editable fields
  if (profile && profile.phone !== undefined) {
    sheet.getRange(rowIndex, CONFIG.USER_COLUMNS.PHONE).setValue(String(profile.phone || '').trim());
  }
  if (profile && profile.department !== undefined) {
    sheet.getRange(rowIndex, CONFIG.USER_COLUMNS.DEPARTMENT).setValue(String(profile.department || '').trim());
  }
  if (CONFIG.USER_COLUMNS.UPDATED_AT) {
    sheet.getRange(rowIndex, CONFIG.USER_COLUMNS.UPDATED_AT).setValue(new Date());
  }

  return { success: true, message: 'Profile updated successfully' };
}

function listUsersMissingUsername() {
  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  const missing = [];
  for (let i = 1; i < data.length; i++) {
    const email = data[i][CONFIG.USER_COLUMNS.EMAIL - 1];
    const username = data[i][CONFIG.USER_COLUMNS.USERNAME - 1];
    if (!username || !String(username).trim()) {
      missing.push({ rowIndex: i + 1, email: email });
    }
  }
  Logger.log(JSON.stringify(missing, null, 2));
  return missing;
}
function legacySyncActionItems_() {
  const year = "2026";
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(year);
  const logSheet = ss.getSheetByName("ACTION_ITEMS_LOG");

  if (!sheet || !logSheet) throw new Error("Missing required sheet(s).");

  const header = getHeaderMap_(sheet);
  const colVN = header["VOUCHER NO."] || header["VOUCHER NO"] || header["ACCOUNT OR EMAIL (VOUCHER NUMBER)"] || header["VOUCHER NUMBER"];
  const colPayee = header["PAYEE"];
  const colAmt = header["GROSS AMOUNT"];
  const colStatus = header["STATUS"];
  const colCN = header["CONTROL NO."] || header["CONTROL NUMBER"] || header["CONTROL NO"];
  const colCreated = header["CREATED AT"];
  const colReleasedAt = header["RELEASED AT"];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const settings = legacyLoadActionItemSettings_(); // {unit:{PAYABLE:{RULE:true},CPO:{RULE:true}}}

  const now = new Date();
  const currentKeys = new Set();
  const newItems = [];

  function enabled_(unit, ruleKey) {
    return settings.unit?.[unit]?.[ruleKey] !== false; // default true
  }

  values.forEach((row, idx) => {
    const rowIndex = idx + 2;

    const status = String(row[colStatus - 1] || "").trim();
    const cn = String(row[colCN - 1] || "").trim();
    const createdAt = row[colCreated - 1];
    const releasedAt = row[colReleasedAt - 1];

    const voucherNo = colVN ? String(row[colVN - 1] || "").trim() : "";
    const payee = colPayee ? String(row[colPayee - 1] || "").trim() : "";
    const amount = colAmt ? Number(row[colAmt - 1] || 0) : 0;

    // RULE 1: PAID_NO_CN
    if (status === "Paid" && !cn) {
      if (enabled_("PAYABLE", "PAID_NO_CN")) {
        newItems.push(makeItem_("PAID_NO_CN", "PAYABLE", year, rowIndex, voucherNo, payee, amount, cn));
      }
      if (enabled_("CPO", "PAID_NO_CN")) {
        newItems.push(makeItem_("PAID_NO_CN", "CPO", year, rowIndex, voucherNo, payee, amount, cn));
      }
    }

    // RULE 2: UNPAID_NO_CN_30D (CPO only)
    if (status === "Unpaid" && !cn && createdAt instanceof Date) {
      const ageDays = Math.floor((now - createdAt) / (1000 * 60 * 60 * 24));
      if (ageDays > 30 && enabled_("CPO", "UNPAID_NO_CN_30D")) {
        newItems.push(makeItem_("UNPAID_NO_CN_30D", "CPO", year, rowIndex, voucherNo, payee, amount, cn, { ageDays }));
      }
    }

    // RULE 3: RELEASED_UNPAID_15D (CPO only)
    if (status === "Unpaid" && cn && releasedAt instanceof Date) {
      const ageDays = Math.floor((now - releasedAt) / (1000 * 60 * 60 * 24));
      if (ageDays > 15 && enabled_("CPO", "RELEASED_UNPAID_15D")) {
        newItems.push(makeItem_("RELEASED_UNPAID_15D", "CPO", year, rowIndex, voucherNo, payee, amount, cn, { ageDays }));
      }
    }
  });

  // Load log into map
  const logLastRow = logSheet.getLastRow();
  const logHeader = getHeaderMap_(logSheet);
  const logData = logLastRow >= 2
    ? logSheet.getRange(2, 1, logLastRow - 1, logSheet.getLastColumn()).getValues()
    : [];

  const mapById = new Map();
  logData.forEach((r, i) => {
    const id = String(r[logHeader["ITEM_ID"] - 1] || "").trim();
    if (id) mapById.set(id, { row: i + 2, status: r[logHeader["STATUS"] - 1] });
  });

  // Mark current as pending (upsert)
  const toAppend = [];
  newItems.forEach(item => {
    currentKeys.add(item.ITEM_ID);

    const existing = mapById.get(item.ITEM_ID);
    if (!existing) {
      toAppend.push(item);
    } else {
      // update LAST_SEEN_AT, ensure STATUS=PENDING, clear RESOLVED_AT if previously resolved
      const rowNum = existing.row;
      logSheet.getRange(rowNum, logHeader["STATUS"]).setValue("PENDING");
      logSheet.getRange(rowNum, logHeader["LAST_SEEN_AT"]).setValue(now);
      logSheet.getRange(rowNum, logHeader["RESOLVED_AT"]).setValue("");
    }
  });

  // Append new items
  if (toAppend.length) {
    const rows = toAppend.map(it => ([
      it.ITEM_ID,
      it.RULE_KEY,
      it.UNIT,
      it.YEAR,
      it.ROW_INDEX,
      it.VOUCHER_NO,
      it.PAYEE,
      it.AMOUNT,
      it.CONTROL_NO,
      "PENDING",
      now,
      now,
      ""
    ]));
    logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Resolve items no longer present (previously pending)
  logData.forEach((r, i) => {
    const id = String(r[logHeader["ITEM_ID"] - 1] || "").trim();
    const status = String(r[logHeader["STATUS"] - 1] || "").trim();
    if (id && status === "PENDING" && !currentKeys.has(id)) {
      const rowNum = i + 2;
      logSheet.getRange(rowNum, logHeader["STATUS"]).setValue("RESOLVED");
      logSheet.getRange(rowNum, logHeader["RESOLVED_AT"]).setValue(now);
      logSheet.getRange(rowNum, logHeader["LAST_SEEN_AT"]).setValue(now);
    }
  });
}


function legacyLoadActionItemSettings_() {
  const ss = SpreadsheetApp.getActive();
  const s = ss.getSheetByName("ACTION_ITEMS_SETTINGS");
  if (!s || s.getLastRow() < 2) return { unit: {} };

  const header = getHeaderMap_(s);
  const data = s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues();

  const out = { unit: {} };
  data.forEach(r => {
    const unit = String(r[header["UNIT"] - 1] || "").trim().toUpperCase();
    const rule = String(r[header["RULE_KEY"] - 1] || "").trim();
    const enabledVal = r[header["ENABLED"] - 1];
    const enabled = String(enabledVal).toUpperCase() !== "FALSE";

    if (!unit || !rule) return;
    if (!out.unit[unit]) out.unit[unit] = {};
    out.unit[unit][rule] = enabled;
  });
  return out;
}

function setControlNoAndReleasedAt_(sheet, rowIndex, controlNumber, colReleasedAt) {
  // Always set CN
  sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER).setValue(controlNumber);

  // Set RELEASED AT only if the column exists and is currently empty
  if (colReleasedAt) {
    const cell = sheet.getRange(rowIndex, colReleasedAt);
    const existing = cell.getValue();
    if (!existing) cell.setValue(new Date());
  }
}

function backfillReleasedAtFromCreatedAt_2026() {
  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const header = getHeaderMap_(sheet);

  const colCN = header["CONTROL NO."] || header["CONTROL NUMBER"] || header["CONTROL NO"];
  const colCreated = header["CREATED AT"];
  const colReleased = header["RELEASED AT"];

  if (!colCN || !colCreated || !colReleased) {
    throw new Error("Missing required columns: CONTROL NO / CREATED AT / RELEASED AT");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, updated: 0 };

  const locale = SpreadsheetApp.getActive().getSpreadsheetLocale(); // used for parsing strings safely
  let updated = 0;

  // Read only needed columns
  const cnVals = sheet.getRange(2, colCN, lastRow - 1, 1).getValues();
  const createdVals = sheet.getRange(2, colCreated, lastRow - 1, 1).getValues();
  const releasedVals = sheet.getRange(2, colReleased, lastRow - 1, 1).getValues();

  for (let i = 0; i < cnVals.length; i++) {
    const cn = String(cnVals[i][0] || "").trim();
    const releasedAt = releasedVals[i][0];

    if (!cn) continue;                 // no CN
    if (releasedAt) continue;          // already has RELEASED AT

    const createdAtRaw = createdVals[i][0];
    const createdAt = parseDateTimeFlexible_(createdAtRaw, locale);
    if (!createdAt) continue;

    sheet.getRange(i + 2, colReleased).setValue(createdAt);
    updated++;
  }

  return { success: true, updated };
}

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

    const payload = announcement && announcement.announcement ? announcement.announcement : announcement;
    const message = String(payload?.message || "").trim();
    const locations = Array.isArray(payload?.locations) ? payload.locations : [];
    const targets = Array.isArray(payload?.targets) && payload.targets.length ? payload.targets : ["ALL"];
    const expiresAt = parseDateTimeFlexible_(payload?.expiresAt);
    const allowDismiss = payload?.allowDismiss !== false;

    if (!message) return { success: false, error: "Message is required." };
    if (!locations.length) return { success: false, error: "Select at least one display location." };
    if (!expiresAt) return { success: false, error: "Invalid expiry date/time." };
    if (expiresAt <= new Date()) return { success: false, error: "Expiry date/time must be in the future." };

    ensureAnnouncementsSheet_();
    const sheet = getSheet(CONFIG.SHEETS.ANNOUNCEMENTS);
    const announcementId = Utilities.getUuid();

    const newRow = [
      announcementId,
      message,
      JSON.stringify(locations),
      JSON.stringify(targets.map(t => String(t).trim()).filter(Boolean)),
      expiresAt,
      allowDismiss,
      new Date(),
      session.email
    ];

    sheet.appendRow(newRow);
    return {
      success: true,
      message: "Announcement created successfully.",
      id: announcementId
    };

  } catch (e) {
    return { success: false, error: "Failed to create announcement: " + e.message };
  }
}

/**
 * API: Gets active announcements for the current user.
 */
function getActiveAnnouncements(token, location) {
  try {
    const session = token ? getSession(token) : null;
    const requestedLocation = String(location || "").toUpperCase();
    const allowPublicLogin = !session && requestedLocation === "LOGIN";

    if (!session && !allowPublicLogin) {
      return { success: false, error: "Session expired" };
    }

    ensureAnnouncementsSheet_();
    const sheet = getSheet(CONFIG.SHEETS.ANNOUNCEMENTS);
    if (sheet.getLastRow() < 2) return { success: true, announcements: [] };

    const data = sheet.getDataRange().getValues();
    const h = getHeaderMap_(sheet);
    const now = new Date();
    const userProperties = session ? PropertiesService.getUserProperties() : null;
    const dismissed = userProperties
      ? JSON.parse(userProperties.getProperty('dismissedAnnouncements') || '{}')
      : {};

    const announcements = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = row[h.ID - 1];
      const expiresAt = parseDateTimeFlexible_(row[h.EXPIRES_AT - 1]);
      if (!expiresAt || expiresAt < now) continue;
      if (session && dismissed[id]) continue;

      let locations = [];
      let targets = [];
      try { locations = JSON.parse(row[h.DISPLAY_LOCATIONS - 1] || "[]"); } catch (_) {}
      try { targets = JSON.parse(row[h.TARGET_USERS - 1] || "[]"); } catch (_) {}

      const locationList = locations.map(x => String(x).toUpperCase());
      const targetList = targets.map(x => String(x));
      const targetListUpper = targetList.map(x => x.toUpperCase());

      if (requestedLocation && !locationList.includes(requestedLocation)) continue;

      let isTargeted = false;
      if (session && session.role === CONFIG.ROLES.ADMIN) {
        isTargeted = true;
      } else if (session) {
        const emailLower = String(session.email || "").toLowerCase();
        const roleUpper = String(session.role || "").toUpperCase();
        isTargeted =
          targetListUpper.includes("ALL") ||
          targetList.some(t => String(t).toLowerCase() === emailLower) ||
          targetListUpper.includes(roleUpper);
      } else if (allowPublicLogin) {
        isTargeted = targetListUpper.includes("ALL") && locationList.includes("LOGIN");
      }

      if (isTargeted) {
        announcements.push({
          id,
          message: row[h.MESSAGE - 1],
          locations,
          targets: targetList,
          allowDismiss: String(row[h.ALLOW_DISMISS - 1]).toUpperCase() !== "FALSE",
          expiresAt: expiresAt,
          createdAt: row[h.CREATED_AT - 1],
          createdBy: row[h.CREATED_BY - 1]
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
    if (!announcementId) return { success: false, error: "Announcement id is required" };

    const userProperties = PropertiesService.getUserProperties();
    const dismissed = JSON.parse(userProperties.getProperty('dismissedAnnouncements') || '{}');
    dismissed[announcementId] = true;
    userProperties.setProperty('dismissedAnnouncements', JSON.stringify(dismissed));

    return { success: true };
  } catch (e) {
    return { success: false, error: "Failed to dismiss announcement: " + e.message };
  }
}

// ==================== TAX FUNCTIONS ====================

/**
 * Get comprehensive tax summary for a year
 */
// TAX FUNCTIONS REMOVED - MOVED TO tax_gs.md

// ==================== ANNOUNCEMENT FUNCTIONS ====================

function legacySendAnnouncement_(data) {
  Logger.log('sendAnnouncement called: ' + JSON.stringify(data));
  
  var user = getCurrentUser();
  if (!user || user.role !== 'Admin') {
    return { success: false, error: 'Only administrators can send announcements' };
  }
  
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Announcements');
    
    if (!sheet) {
      return { success: false, error: 'Legacy Announcements sheet not found' };
    }
    
    sheet.appendRow([
      data.id,
      data.title,
      data.message,
      JSON.stringify(data.locations),
      JSON.stringify(data.recipients),
      data.expiry,
      data.allowDismiss,
      data.createdBy,
      data.createdAt
    ]);
    
    Logger.log('Announcement saved successfully');
    return { success: true };
    
  } catch (e) {
    Logger.log('Error saving announcement: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function legacyGetAnnouncements_() {
  Logger.log('getAnnouncements called');
  
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Announcements');
    
    if (!sheet) {
      Logger.log('No Announcements sheet found');
      return { success: true, announcements: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var announcements = [];
    
    for (var i = 1; i < data.length; i++) {
      announcements.push({
        id: data[i][0],
        title: data[i][1],
        message: data[i][2],
        locations: JSON.parse(data[i][3] || '[]'),
        recipients: JSON.parse(data[i][4] || '[]'),
        expiry: data[i][5],
        allowDismiss: data[i][6],
        createdBy: data[i][7],
        createdAt: data[i][8]
      });
    }
    
    Logger.log('Found ' + announcements.length + ' announcements');
    return { success: true, announcements: announcements.reverse() };
    
  } catch (e) {
    Logger.log('Error getting announcements: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function legacyGetActiveAnnouncements_(params) {
  Logger.log('legacyGetActiveAnnouncements_ called: ' + JSON.stringify(params));
  
  var page = params.page;
  var username = params.username;
  var now = new Date();
  
  var result = legacyGetAnnouncements_();
  if (!result.success) return result;
  
  var active = [];
  
  for (var i = 0; i < result.announcements.length; i++) {
    var ann = result.announcements[i];
    
    // Check if expired
    if (new Date(ann.expiry) < now) {
      Logger.log('Announcement "' + ann.title + '" expired');
      continue;
    }
    
    // Check if for this page
    if (ann.locations.indexOf(page) === -1) {
      Logger.log('Announcement "' + ann.title + '" not for page: ' + page);
      continue;
    }
    
    // Check if for this user
    if (ann.recipients.indexOf('ALL') === -1 && ann.recipients.indexOf(username) === -1) {
      Logger.log('Announcement "' + ann.title + '" not for user: ' + username);
      continue;
    }
    
    active.push(ann);
  }
  
  Logger.log('Found ' + active.length + ' active announcements for ' + username + ' on ' + page);
  return { success: true, announcements: active };
}

function legacyDeleteAnnouncement_(params) {
  Logger.log('legacyDeleteAnnouncement_ called: ' + JSON.stringify(params));
  
  var user = getCurrentUser();
  if (!user || user.role !== 'Admin') {
    return { success: false, error: 'Only administrators can delete announcements' };
  }
  
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Announcements');
    
    if (!sheet) {
      return { success: false, error: 'Announcements sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === params.id) {
        sheet.deleteRow(i + 1);
        Logger.log('Announcement deleted');
        return { success: true };
      }
    }
    
    return { success: false, error: 'Announcement not found' };
    
  } catch (e) {
    Logger.log('Error deleting announcement: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Accepts Date object OR strings like "1/5/2026 20:20:22"
 * Uses spreadsheet locale to decide dd/mm vs mm/dd when ambiguous.
 */
function parseDateTimeFlexible_(value, locale) {
  if (value instanceof Date) return value;

  const s = String(value || "").trim();
  if (!s) return null;

  // Expect: d/m/yyyy hh:mm:ss OR m/d/yyyy hh:mm:ss
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) {
    const d2 = new Date(s);
    return isNaN(d2.getTime()) ? null : d2;
  }

  const a = parseInt(m[1], 10);
  const b = parseInt(m[2], 10);
  const yyyy = parseInt(m[3], 10);
  const hh = parseInt(m[4], 10);
  const mm = parseInt(m[5], 10);
  const ss = parseInt(m[6] || "0", 10);

  // Decide order
  // If one part > 12, it must be the day.
  // If ambiguous (both <= 12), use locale: en_US => mm/dd else dd/mm
  let day, month;
  if (a > 12) { day = a; month = b; }
  else if (b > 12) { day = b; month = a; }
  else {
    const isUS = String(locale || "").toLowerCase().includes("us");
    if (isUS) { month = a; day = b; } else { day = a; month = b; }
  }

  const d = new Date(yyyy, month - 1, day, hh, mm, ss);
  return isNaN(d.getTime()) ? null : d;
}

function toDate_(value) {
  if (value instanceof Date) return value;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function diagnosticCheck() { 
  const requiredFunctions = [ 
    'doGet',
    'doPost',
    'hasPermission',
    'updateVoucher',
    'updateVoucherStatus',
    'requestVoucherDelete',
    'getVouchers',
    'getSummary',
    'handleLogin',
    'validateSession',
    'getAllYearsSummary',
    'getCachedData',
    'assignControlNumber',
    'approveVoucherDelete',
    'batchUpdateStatus',
    'createNotifications',
    'clearCache',
    'clearAllVoucherCaches',
    'createJsonOutput',
    'clearSession',
    'changePassword',
    'createUser',
    'clearAllSessions',
    'createVoucher',
    'deleteUser',
    'deleteVoucher',
    'extractYearFromVoucherNumber',
    'formatCurrencyForNotif',
    'getNextControlNumber',
    'getNotifications',
    'getPendingDeletions',
    'getSheet',
    'generateToken',
    'getAuditTrail',
    'getSession',
    'getNigerianTimestamp',
    'generateControlNumber',
    'getUsers',
    'getRolePermissions',
    'getCategories',
    'getAvailableRoles',
    'getVoucherByRow',
    'hashPassword',
    'handleLogout',
    'lookupVoucher',
    'logAudit',
    'markNotificationRead',
    'markAllNotificationsRead',
    'parseAmount',
    'preloadCache',
    'releaseVouchersWithNotification',
    'releaseVouchers',
    'rejectVoucherDelete',
    'rowToVoucher',
    'releaseSelectedVouchers',
    'sendReleaseNotification',
    'storeSession',
    'setupAdminUser',
    'updateUser',
    'voucherToRow'
  ];

  const results = requiredFunctions.map(fn => {
    const exists = typeof this[fn] === 'function';
    return {
      functionName: fn,
      status: exists ? 'OK' : 'MISSING'
    };
  });

  const missing = results.filter(r => r.status === 'MISSING');

  return {
    success: missing.length === 0,
    totalChecked: requiredFunctions.length,
    missingCount: missing.length,
    missingFunctions: missing.map(m => m.functionName),
    details: results
  };

  function testNextControlNumber () {
  // First, create a session
  const loginResult = login('admin@fmc.gov.ng', 'Admin@2026'); // Use your actual credentials
  
  if (!loginResult.success) {
    Logger.log('Login failed: ' + loginResult.error);
    return;
  }
  
  Logger.log('Login successful, token: ' + loginResult.token);
  
  // Test getNextControlNumber
  const result = getNextControlNumber(loginResult.token, 'CPO');
  Logger.log('getNextControlNumber result: ' + JSON.stringify(result));
  
  // Should output something like:
  // { success: true, controlNumber: "CN-CPO-001" }
}
}

/**
 * DEBT PROFILE REQUEST & APPROVAL SYSTEM
 */

/**
 * Request a new Debt Profile generation
 */
function requestDebtProfile(token, filterData) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };
    
    // Permission check: Any user with reporting access can request, but only senior roles auto-approve
    if (!hasPermission(session.role, [CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.DDFA, CONFIG.ROLES.DFA, CONFIG.ROLES.ADMIN])) {
      return { success: false, error: 'Unauthorized: You do not have permission to request debt profile generation.' };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEETS.DEBT_PROFILE_REQUESTS);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.DEBT_PROFILE_REQUESTS);
      sheet.getRange(1, 1, 1, 14).setValues([[
        'REQUEST_ID', 'TIMESTAMP', 'REQUESTER_EMAIL', 'REQUESTER_NAME', 
        'FILTERS', 'STATUS', 'APPROVER_EMAIL', 'APPROVAL_DATE', 'COMMENTS', 'REPORT_ID',
        'REPORT_TITLE', 'EXECUTIVE_SUMMARY', 'ANALYSIS', 'RECOMMENDATIONS'
      ]]);
      sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#e0e0e0');
    }
    
    // Senior roles auto-approve
    const isSeniorRole = [CONFIG.ROLES.DDFA, CONFIG.ROLES.DFA, CONFIG.ROLES.ADMIN].includes(session.role);
    const status = isSeniorRole ? 'APPROVED' : 'PENDING';
    const requestId = 'REQ-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const ts = getNigerianTimestamp();
    
    // Extract metadata from filterData if provided as object
    let filters = '';
    let reportTitle = 'Debt Profile Report 2026';
    let execSummary = '';
    let analysis = '';
    let recommendations = '';
    
    if (typeof filterData === 'object') {
      // Handle the case where the API wraps everything in a 'filterData' key
      const actualData = filterData.filterData || filterData;
      
      filters = JSON.stringify(actualData.filters || {});
      reportTitle = actualData.title || reportTitle;
      execSummary = actualData.summary || '';
      analysis = actualData.analysis || '';
      recommendations = actualData.recommendations || '';
    } else {
      filters = typeof filterData === 'string' ? filterData : JSON.stringify(filterData || {});
    }
    
    sheet.appendRow([
      requestId, ts, session.email, session.name, 
      filters, status, 
      isSeniorRole ? session.email : '', 
      isSeniorRole ? ts : '', 
      isSeniorRole ? 'System Auto-Approved' : '', 
      '',
      reportTitle, execSummary, analysis, recommendations
    ]);
    
    // Ensure data is immediately available for subsequent calls in the same execution
    SpreadsheetApp.flush();
    
    if (!isSeniorRole) {
      const approverRoles = [CONFIG.ROLES.DDFA, CONFIG.ROLES.DFA, CONFIG.ROLES.ADMIN];
      const usersSheet = getSheet(CONFIG.SHEETS.USERS);
      const usersData = usersSheet.getDataRange().getValues();
      const approverEmails = [];
      
      for (let i = 1; i < usersData.length; i++) {
        if (approverRoles.includes(usersData[i][3]) && usersData[i][4]) {
          approverEmails.push(usersData[i][1]);
        }
      }
      
      createNotifications(
        approverEmails, 
        'Debt Profile Generation Request', 
        `${session.name} has requested a comprehensive Debt Profile Report. Approval required.`,
        'reports.html?tab=debt-profile',
        'warning'
      );
    }
    
    return { 
      success: true, 
      requestId: requestId, 
      message: isSeniorRole ? 'Report generated successfully.' : 'Request submitted for approval.',
      status: status
    };
  } catch (e) {
    return { success: false, error: 'Request failed: ' + e.message };
  }
}

function getDebtProfileRequestStatus(token) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DEBT_PROFILE_REQUESTS);
    if (!sheet) return { success: true, status: 'NONE' };
    
    const data = sheet.getDataRange().getValues();
    // Search backwards for the latest request by THIS user
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (row[2] === session.email) {
        return {
          success: true,
          requestId: row[0],
          timestamp: row[1],
          requester: row[3],
          filters: JSON.parse(row[4] || '{}'),
          status: row[5],
          approver: row[6],
          approvalDate: row[7],
          comments: row[8],
          narrative: {
            title: row[10],
            summary: row[11],
            analysis: row[12],
            recommendations: row[13]
          }
        };
      }
    }
    
    return { success: true, status: 'NONE' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function handleDebtProfileApproval(token, requestId, action, comments) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };
    
    if (!hasPermission(session.role, [CONFIG.ROLES.DDFA, CONFIG.ROLES.DFA, CONFIG.ROLES.ADMIN])) {
      return { success: false, error: 'Unauthorized.' };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DEBT_PROFILE_REQUESTS);
    if (!sheet) return { success: false, error: 'No requests.' };
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let requesterEmail = '';
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === requestId) {
        rowIndex = i + 1;
        requesterEmail = data[i][2];
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Request not found.' };
    
    const status = action === 'approve' ? 'APPROVED' : 'REJECTED';
    const ts = getNigerianTimestamp();
    
    sheet.getRange(rowIndex, 6).setValue(status);
    sheet.getRange(rowIndex, 7).setValue(session.email);
    sheet.getRange(rowIndex, 8).setValue(ts);
    sheet.getRange(rowIndex, 9).setValue(comments || '');
    
    createNotifications(
      [requesterEmail],
      'Debt Profile Request',
      `Your request for ${requestId} has been ${action}d by ${session.name}.`,
      'reports.html?tab=debt-profile',
      status === 'APPROVED' ? 'success' : 'danger'
    );
    
    return { success: true, message: `Request ${action}d successfully.` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getDebtProfileFullData(token, requestId) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };
    
    SpreadsheetApp.flush();
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const requestsSheet = ss.getSheetByName(CONFIG.SHEETS.DEBT_PROFILE_REQUESTS);
    if (!requestsSheet) return { success: false, error: 'Requests sheet missing.' };
    
    const requestsData = requestsSheet.getDataRange().getValues();
    let request = null;
    const reqCols = CONFIG.DEBT_REQUEST_COLUMNS;
    
    for (let i = 1; i < requestsData.length; i++) {
      if (requestsData[i][0] === requestId) {
        request = {
          id: requestsData[i][0],
          filters: JSON.parse(requestsData[i][reqCols.FILTERS - 1] || '{}'),
          status: requestsData[i][reqCols.STATUS - 1],
          title: requestsData[i][reqCols.REPORT_TITLE - 1] || 'Debt Profile Report',
          summary: requestsData[i][reqCols.EXECUTIVE_SUMMARY - 1] || '',
          analysis: requestsData[i][reqCols.ANALYSIS - 1] || '',
          recommendations: requestsData[i][reqCols.RECOMMENDATIONS - 1] || ''
        };
        break;
      }
    }
    
    if (!request) return { success: false, error: 'Request not found.' };
    if (request.status !== 'APPROVED' && session.role !== CONFIG.ROLES.ADMIN) {
      return { success: false, error: 'Request not approved yet.' };
    }
    
    const filters = request.filters;
    const startDate = filters.startDate ? new Date(filters.startDate) : null;
    const endDate = filters.endDate ? new Date(filters.endDate) : null;
    
    const vouchersSheet = ss.getSheetByName(CONFIG.SHEETS.VOUCHERS_2026);
    const data = vouchersSheet.getDataRange().getValues();
    const cols = CONFIG.VOUCHER_COLUMNS;
    const header = getHeaderMap_(vouchersSheet);
    const OLD_VN_COL = (header['OLD VOUCHER NUMBER'] || cols.OLD_VOUCHER_NUMBER) - 1;
    const OLD_VN_AVAILABLE_COL = (header['OLD VOUCHER NO AVAILABLE?'] || header['OLD VOUCHER AVAILABLE'] || cols.OLD_VOUCHER_AVAILABLE) - 1;
    const ACCOUNT_TYPE_COL = (header['ACCOUNT TYPE'] || cols.ACCOUNT_TYPE) - 1;
    const SUB_ACCOUNT_COL = (header['SUB ACCOUNT'] || header['SUB ACCOUNT TYPE'] || cols.SUB_ACCOUNT_TYPE) - 1;
    
    // ----- FETCH 2025 DATA FOR BALANCE B/F -----
    let balanceBF = 0;
    try {
      const sheet2025 = ss.getSheetByName(CONFIG.SHEETS.VOUCHERS_2025);
      if (sheet2025) {
        const data2025 = sheet2025.getDataRange().getValues();
        const cols2025 = CONFIG.VOUCHER_COLUMNS;
        for (let i = 1; i < data2025.length; i++) {
          const status = String(data2025[i][cols2025.STATUS - 1]).trim().toLowerCase();
          if (status === 'unpaid') {
            balanceBF += parseAmount(data2025[i][cols2025.GROSS_AMOUNT - 1]);
          }
        }
      }
    } catch (e) {
      Logger.log("Warning: Could not fetch 2025 data for B/F: " + e.message);
    }

    const result = {
      summary: { 
        totalDebt: 0, 
        count: 0, 
        overdueCount: 0, 
        overdueAmount: 0,
        balanceBF: balanceBF,
        revalidatedAmount: 0,
        currentUnpaid2026: 0,
        totalContractSum: 0,
        totalPayments: 0,
        paymentEfficiency: 0,
        debtGrowthRate: 0
      },
      currentMonth: {
        name: '',
        newObligations: 0,
        payments: 0,
        unpaid: 0
      },
      taxSummary: { totalVAT: 0, totalWHT: 0, totalStampDuty: 0, totalTax: 0 },
      byCategory: {},
      byDepartment: {},
      byAccountType: {},
      byAge: { '0-30 Days': 0, '31-90 Days': 0, '91+ Days': 0 },
      details: []
    };
    
    const now = new Date();
    const currentMonthName = Utilities.formatDate(now, "GMT", "MMMM");
    result.currentMonth.name = currentMonthName;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = String(row[cols.STATUS - 1]).trim().toLowerCase();
      const voucherMonth = String(row[cols.PMT_MONTH - 1]).trim();
      
      const grossAmount = parseAmount(row[cols.GROSS_AMOUNT - 1]);
      const contractSum = parseAmount(row[cols.CONTRACT_SUM - 1]);
      const hasOldVoucherNumber = !!String(row[OLD_VN_COL] || '').trim();
      const oldVoucherAvailable = String(row[OLD_VN_AVAILABLE_COL] || '').trim().toLowerCase();
      const isRevalidated = hasOldVoucherNumber || oldVoucherAvailable === 'yes';
      
      // Multi-metric tracking (Performance Summary)
      result.summary.totalContractSum += contractSum;
      if (status === 'paid') {
        result.summary.totalPayments += grossAmount;
        if (voucherMonth === currentMonthName) {
          result.currentMonth.payments += grossAmount;
        }
      }
      
      if (status !== 'unpaid') continue;
      
      // Filtering for report period if requested
      const dateVal = row[cols.DATE - 1];
      const voucherDate = dateVal ? new Date(dateVal) : null;
      if (startDate && voucherDate && voucherDate < startDate) continue;
      if (endDate && voucherDate && voucherDate > endDate) continue;
      
      const amount = grossAmount;
      const vat = parseAmount(row[cols.VAT - 1]);
      const wht = parseAmount(row[cols.WHT - 1]);
      const stampDuty = parseAmount(row[cols.STAMP_DUTY - 1]);
      
      const category = String(row[cols.CATEGORIES - 1]).trim() || 'Uncategorized';
      const baseDept = String(row[ACCOUNT_TYPE_COL]).trim() || 'General';
      const subDept = String(row[SUB_ACCOUNT_COL]).trim();
      const departmentKey = subDept ? `${baseDept}::${subDept}` : `${baseDept}::`;
      const createdAt = row[cols.CREATED_AT - 1] ? new Date(row[cols.CREATED_AT - 1]) : voucherDate;
      
      result.summary.totalDebt += amount;
      result.summary.count++;
      result.summary.currentUnpaid2026 += amount;
      if (isRevalidated) result.summary.revalidatedAmount += amount;

      if (voucherMonth === currentMonthName) {
        result.currentMonth.newObligations += amount;
        result.currentMonth.unpaid += amount;
      }
      
      result.taxSummary.totalVAT += vat;
      result.taxSummary.totalWHT += wht;
      result.taxSummary.totalStampDuty += stampDuty;
      result.taxSummary.totalTax += (vat + wht + stampDuty);
      
      if (!result.byCategory[category]) result.byCategory[category] = 0;
      result.byCategory[category] += amount;
      
      if (!result.byDepartment[departmentKey]) result.byDepartment[departmentKey] = 0;
      result.byDepartment[departmentKey] += amount;

      if (!result.byAccountType[baseDept]) {
        result.byAccountType[baseDept] = { total: 0, subTypes: {} };
      }
      result.byAccountType[baseDept].total += amount;
      const subTypeKey = subDept || 'Unspecified';
      if (!result.byAccountType[baseDept].subTypes[subTypeKey]) {
        result.byAccountType[baseDept].subTypes[subTypeKey] = 0;
      }
      result.byAccountType[baseDept].subTypes[subTypeKey] += amount;
      
      if (createdAt) {
        const ageDays = Math.floor((now - createdAt) / (1000 * 60 * 60 * 24));
        if (ageDays <= 30) result.byAge['0-30 Days'] += amount;
        else if (ageDays <= 90) result.byAge['31-90 Days'] += amount;
        else {
          result.byAge['91+ Days'] += amount;
          result.summary.overdueCount++;
          result.summary.overdueAmount += amount;
        }
      } else {
        result.byAge['0-30 Days'] += amount;
      }
      
      // Captured detail records (Full schedule for official report)
      result.details.push({
        payee: row[cols.PAYEE - 1],
        particular: row[cols.PARTICULAR - 1],
        amount: amount,
        date: voucherDate ? Utilities.formatDate(voucherDate, "GMT", "yyyy-MM-dd") : 'N/A',
        category: category,
        department: subDept ? `${baseDept} - ${subDept}` : baseDept
      });
    }
    
    // Final Performance Calculations
    const vouchersRaised2026 = result.summary.totalPayments + result.summary.currentUnpaid2026;
    result.summary.paymentEfficiency = vouchersRaised2026 > 0 
      ? (result.summary.totalPayments / vouchersRaised2026) * 100 
      : 0;
    
    result.summary.debtGrowthRate = result.summary.balanceBF > 0 
      ? ((result.summary.totalDebt / result.summary.balanceBF) - 1) * 100 
      : 0;

    // Convert and sort breakdowns for easier consumption in the template
    const categoryBreakdown = Object.entries(result.byCategory)
      .map(([name, amount]) => ({ name, amount }))
      .sort((a, b) => b.amount - a.amount);
      
    const departmentBreakdown = Object.entries(result.byDepartment)
      .map(([nameKey, amount]) => {
         const [baseType, subType] = nameKey.split('::');
         return { 
           name: baseType, 
           subType: subType,
           displayName: subType ? `${baseType} - ${subType}` : baseType,
           amount 
         };
      })
      .sort((a, b) => b.amount - a.amount);

    const byAccountType = {};
    Object.keys(result.byAccountType)
      .sort((a, b) => result.byAccountType[b].total - result.byAccountType[a].total)
      .forEach((parentType) => {
        const node = result.byAccountType[parentType];
        const sortedSub = {};
        Object.keys(node.subTypes)
          .sort((a, b) => node.subTypes[b] - node.subTypes[a])
          .forEach((subType) => {
            sortedSub[subType] = node.subTypes[subType];
          });
        byAccountType[parentType] = {
          total: node.total,
          subTypes: sortedSub
        };
      });

    return { 
      success: true, 
      data: result, 
      categoryBreakdown: categoryBreakdown, // SHOW ALL CATEGORIES
      departmentBreakdown: departmentBreakdown, // SHOW ALL DEPARTMENTS
      byAccountType: byAccountType,
      balanceBF: result.summary.balanceBF,
      total2026Unpaid: result.summary.currentUnpaid2026,
      revalidated: result.summary.revalidatedAmount,
      totalDebtProfile: result.summary.balanceBF + result.summary.currentUnpaid2026,
      performance: {
        contractSum: result.summary.totalContractSum,
        vouchersRaised: vouchersRaised2026,
        paymentsMade: result.summary.totalPayments,
        unpaidBalance: result.summary.currentUnpaid2026,
        efficiency: result.summary.paymentEfficiency.toFixed(1) + '%',
        growth: result.summary.debtGrowthRate.toFixed(1) + '%'
      },
      currentMonth: result.currentMonth,
      filters: filters, 
      narrative: {
        title: request.title,
        summary: request.summary,
        analysis: request.analysis,
        recommendations: request.recommendations
      },
      generatedAt: getNigerianTimestamp() 
    };
  } catch (error) {
    return { success: false, error: 'Data aggregation failed: ' + error.message };
  }
}

/**
 * Generates a professional PDF for the Debt Profile
 */
function generateDebtProfilePDF(token, requestId) {
  try {
    const res = getDebtProfileFullData(token, requestId);
    if (!res.success) return res;
    
    const data = res.data;
    const narrative = res.narrative;
    
    // Check if template exists to avoid script crash (which causes CORS/500 errors)
    let htmlTemplate;
    try {
      htmlTemplate = HtmlService.createTemplateFromFile('DebtReportTemplate');
    } catch (templateError) {
      return { 
        success: false, 
        error: "PDF Template 'DebtReportTemplate' not found in Google Apps Script project. Please create an HTML file with this name in the GAS editor and paste the template code." 
      };
    }

    // Robust assignment with defaults to prevent "is not defined" errors in template
    htmlTemplate.data = res.data || { summary: {}, details: [], taxSummary: {} };
    htmlTemplate.narrative = res.narrative || { title: 'Debt Profile', summary: '', analysis: '', recommendations: '' };
    htmlTemplate.categoryBreakdown = res.categoryBreakdown || [];
    htmlTemplate.departmentBreakdown = res.departmentBreakdown || [];
    htmlTemplate.balanceBF = res.balanceBF || 0;
    htmlTemplate.revalidated = res.revalidated || 0;
    htmlTemplate.total2026Unpaid = res.total2026Unpaid || 0;
    htmlTemplate.totalDebtProfile = res.totalDebtProfile || 0;
    htmlTemplate.performance = res.performance || { vouchersRaised: 0, paymentsMade: 0, efficiency: '0%', growth: '0%', contractSum: 0 };
    htmlTemplate.currentMonth = res.currentMonth || { name: 'Month', newObligations: 0, payments: 0, unpaid: 0 };
    htmlTemplate.generatedAt = res.generatedAt || new Date().toLocaleString();
    
    const html = htmlTemplate.evaluate().getContent();
    const blob = Utilities.newBlob(html, 'text/html', 'report.html');
    const pdf = blob.getAs('application/pdf').setName(narrative.title.replace(/\s+/g, '_') + '.pdf');
    
    // In GAS, we usually return a base64 string or a Drive file URL. 
    // For direct web apps, returning base64 is common for small files.
    return { 
      success: true, 
      pdfBase64: Utilities.base64Encode(pdf.getBytes()),
      fileName: pdf.getName()
    };
  } catch (e) {
    return { success: false, error: 'PDF Generation failed: ' + e.message };
  }
}

function generateDebtProfileExcel(token, requestId) {
  try {
    const res = getDebtProfileFullData(token, requestId);
    if (!res.success) return res;
    
    const data = res.data;
    const narrative = res.narrative;
    const catBreakdown = res.categoryBreakdown || [];
    const perf = res.performance || {};
    const cm = res.currentMonth || {};
    
    const ss = SpreadsheetApp.create('Debt_Profile_Analytical_Report_' + requestId);
    
    // --- SHEET 1: POSITION & SUMMARY ---
    const posSheet = ss.getSheets()[0];
    posSheet.setName('1-2. Summary & Position');
    
    posSheet.getRange('A1:C1').merge().setValue('FEDERAL MEDICAL CENTRE, ABEOKUTA').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
    posSheet.getRange('A2:C2').merge().setValue('DEBT PROFILE ANALYTICAL REPORT').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    posSheet.getRange('A3:C3').merge().setValue('As of ' + res.generatedAt).setFontStyle('italic').setHorizontalAlignment('center');
    
    posSheet.getRange('A5').setValue('1. OVERVIEW & OBJECTIVES').setFontWeight('bold').setBackground('#f3f3f3');
    posSheet.getRange('A6:C6').merge().setValue(narrative.summary).setWrap(true);
    posSheet.setRowHeight(6, 80);
    
    posSheet.getRange('A8').setValue('2. CURRENT FINANCIAL POSITION').setFontWeight('bold').setBackground('#f3f3f3');
    posSheet.getRange('A9:C9').setValues([['Liability Type', 'Amount (₦)', 'Description']]).setBackground('#003366').setFontColor('white');
    posSheet.getRange('A10:C13').setValues([
      ['Balance B/F (2025)', res.balanceBF, 'Unpaid vouchers carried into 2026'],
      ['Revalidated Vouchers', res.revalidated, 'Verified historical obligations'],
      ['Current Unpaid (2026)', res.total2026Unpaid, 'Newly incurred debt'],
      ['TOTAL DEBT PROFILE', res.totalDebtProfile, 'Total institutional liability']
    ]);
    posSheet.getRange('B10:B13').setNumberFormat('₦#,##0.00');
    posSheet.getRange('A13:C13').setFontWeight('bold').setBackground('#fff2cc');
    posSheet.setColumnWidth(1, 200);
    posSheet.setColumnWidth(2, 200);
    posSheet.setColumnWidth(3, 300);

    // --- SHEET 2: PERFORMANCE & COMPARISON ---
    const perfSheet = ss.insertSheet('3-4. Performance & Comparison');
    perfSheet.getRange('A1').setValue('3. FINANCIAL PERFORMANCE (2026)').setFontWeight('bold').setBackground('#f3f3f3');
    perfSheet.getRange('A2:B2').setValues([['Metric', 'Value']]).setBackground('#003366').setFontColor('white');
    perfSheet.getRange('A3:B8').setValues([
      ['Total Vouchers Raised', perf.vouchersRaised],
      ['Total Payments Made', perf.paymentsMade],
      ['Unpaid Balance', perf.unpaidBalance],
      ['Payment Efficiency', perf.efficiency],
      ['Debt Growth Rate', perf.growth],
      ['Total Contract Sum', perf.contractSum]
    ]);
    perfSheet.getRange('B3:B5').setNumberFormat('₦#,##0.00');
    perfSheet.getRange('B8').setNumberFormat('₦#,##0.00');
    
    perfSheet.getRange('A10').setValue('4. PERIOD COMPARISON').setFontWeight('bold').setBackground('#f3f3f3');
    perfSheet.getRange('A11:C11').setValues([['Metric', 'Current Month ('+cm.name+')', 'Cumulative (YTD)']]).setBackground('#003366').setFontColor('white');
    perfSheet.getRange('A12:C14').setValues([
      ['New Obligations', cm.newObligations, perf.vouchersRaised],
      ['Actual Payments', cm.payments, perf.paymentsMade],
      ['Unpaid Balance', cm.unpaid, res.totalDebtProfile]
    ]);
    perfSheet.getRange('B12:C14').setNumberFormat('₦#,##0.00');
    perfSheet.autoResizeColumns(1, 3);

    // --- SHEET 3: EXPENDITURE ANALYSIS ---
    const expSheet = ss.insertSheet('5. Expenditure Analysis');
    expSheet.getRange('A1').setValue('SECTORAL DEBT DISTRIBUTION (ALL CATEGORIES)').setFontWeight('bold');
    expSheet.getRange('A2:B2').setValues([['Category', 'Amount (₦)']]).setBackground('#003366').setFontColor('white');
    if (catBreakdown.length > 0) {
      const rows = catBreakdown.map(c => [c.name, c.amount]);
      expSheet.getRange(3, 1, rows.length, 2).setValues(rows);
      expSheet.getRange(3, 2, rows.length, 1).setNumberFormat('₦#,##0.00');
    }
    
    const taxRow = catBreakdown.length + 5;
    expSheet.getRange(taxRow, 1).setValue('STATUTORY TAX FORECAST').setFontWeight('bold');
    expSheet.getRange(taxRow + 1, 1, 1, 2).setValues([['Tax Type', 'Amount (₦)']]).setBackground('#003366').setFontColor('white');
    expSheet.getRange(taxRow + 2, 1, 4, 2).setValues([
      ['VAT', data.taxSummary.totalVAT],
      ['WHT', data.taxSummary.totalWHT],
      ['Stamp Duty', data.taxSummary.totalStampDuty],
      ['Total Liability', data.taxSummary.totalTax]
    ]);
    expSheet.getRange(taxRow + 2, 2, 4, 1).setNumberFormat('₦#,##0.00');
    expSheet.getRange(taxRow + 5, 1, 1, 2).setFontWeight('bold').setBackground('#fff2cc');
    
    expSheet.getRange('D1').setValue('5. ANALYTICAL COMMENTARY').setFontWeight('bold');
    expSheet.getRange('D2:G2').merge().setValue(narrative.analysis).setWrap(true);
    expSheet.setColumnWidth(4, 400);
    expSheet.autoResizeColumns(1, 3);

    // --- SHEET 4: RECOMMENDATIONS ---
    const recSheet = ss.insertSheet('6. Recommendations');
    recSheet.getRange('A1').setValue('6. STRATEGIC OBSERVATIONS & RECOMMENDATIONS').setFontWeight('bold').setBackground('#f3f3f3');
    recSheet.getRange('A2:E6').merge().setValue(narrative.recommendations).setWrap(true).setVerticalAlignment('top');
    recSheet.setRowHeight(2, 200);
    recSheet.setColumnWidth(1, 600);

    // --- SHEET 5: DETAILED SCHEDULE ---
    const detSheet = ss.insertSheet('7. Detailed Schedule');
    detSheet.appendRow(['Payee', 'Particular', 'Amount', 'Date', 'Category']);
    detSheet.getRange(1, 1, 1, 5).setBackground('#003366').setFontColor('white').setFontWeight('bold');
    if (data.details && data.details.length > 0) {
      const detailRows = data.details.map(d => [d.payee, d.particular, d.amount, d.date, d.category]);
      detSheet.getRange(2, 1, detailRows.length, 5).setValues(detailRows);
      detSheet.getRange(2, 3, detailRows.length, 1).setNumberFormat('₦#,##0.00');
    }
    detSheet.setFrozenRows(1);
    detSheet.autoResizeColumns(1, 5);
    
    SpreadsheetApp.flush();
    
    try {
      DriveApp.getFileById(ss.getId()).setAnonymousAccess(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      return { success: true, downloadUrl: ss.getUrl(), message: 'Excel generated. Permissions error: ' + e.message };
    }
    
    return { success: true, downloadUrl: ss.getUrl(), message: 'Comprehensive Analytical Excel report generated successfully.' };
  } catch (e) {
    return { success: false, error: 'Excel Generation failed: ' + e.message };
  }
}
