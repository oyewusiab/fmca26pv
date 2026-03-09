/**
 * =====================================================================
 * PAYABLE VOUCHER 2026 — CORRECTED BACKEND CODE
 * =====================================================================
 * 
 * Copy these functions into your Google Apps Script project.
 * This fixes all 5 failing tests:
 *   1. Login simulation
 *   2. Session validation
 *   3. Voucher fetch
 *   4. Pagination logic
 *   5. Role security test
 * 
 * FILES TO UPDATE IN APPS SCRIPT:
 *   - Code.gs        → doGet, doPost (main routers)
 *   - Auth.gs         → login, logout, getSession, validateSession, changePassword
 *   - Vouchers.gs     → getVouchers (with pagination)
 *   - Config.gs       → already correct (no changes needed)
 *   - SheetUtils.gs   → already correct (no changes needed)
 * =====================================================================
 */


// =============================================================================
// 1. CODE.gs — MAIN REQUEST ROUTERS
// =============================================================================
// Replace your existing doGet and doPost with these:

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

            // ---- VOUCHERS ----
            case 'getVouchers':
                result = getVouchers(params.token, {
                    year: params.year || '2026',
                    page: parseInt(params.page) || 1,
                    pageSize: parseInt(params.pageSize) || 50,
                    filters: params.filters ? JSON.parse(params.filters) : null
                });
                break;
            case 'getVoucherByRow':
                result = getVoucherByRow(params.token, params.rowIndex, params.year);
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
            case 'getQuickStats':
                result = getQuickStats(params.token);
                break;

            // ---- USERS ----
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

            // ---- CHAT ----
            case 'getChatUsers':
                result = getChatUsers(params.token);
                break;
            case 'getChatThread':
                result = getChatThread(params.token, params.otherEmail, parseInt(params.limit) || 50);
                break;
            case 'getChatUnreadCount':
                result = getChatUnreadCount(params.token);
                break;
            case 'getChatUnreadSummary':
                result = getChatUnreadSummary(params.token);
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

            // ---- DELETIONS ----
            case 'getPendingDeletions':
                result = getPendingDeletions(params.token);
                break;

            // ---- PROFILE ----
            case 'getMyProfile':
                result = getMyProfile(params.token);
                break;

            default:
                result = { success: false, error: 'Unknown action: ' + action };
        }

        return ContentService.createTextOutput(JSON.stringify(result))
            .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
            success: false,
            error: err.message
        })).setMimeType(ContentService.MimeType.JSON);
    }
}


function doPost(e) {
    try {
        let payload;

        // Support both application/json and text/plain (for CORS-free requests)
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
                result = updateStatus(token, payload.rowIndex, payload.status, payload.pmtMonth);
                break;
            case 'batchUpdateStatus':
                result = batchUpdateStatus(token, payload.controlNumber, payload.status, payload.pmtMonth);
                break;
            case 'assignControlNumber':
                result = assignControlNumber(token, payload.rowIndexes, payload.controlNumber);
                break;
            case 'releaseSelectedVouchers':
                result = releaseSelectedVouchers(token, payload.rowIndexes, payload.targetUnit);
                break;
            case 'releaseVouchersWithNotification':
                result = releaseVouchersWithNotification(token, payload.rowIndexes, payload.controlNumber, payload.targetUnit);
                break;
            case 'releaseVouchers':
                result = releaseVouchers(token, payload);
                break;
            case 'createVoucherFromLookup':
                result = createVoucherFromLookup(token, payload.lookupResult, payload.additionalData);
                break;

            // ---- DELETE WORKFLOW ----
            case 'requestDelete':
                result = requestDelete(token, payload.rowIndex, payload.reason, payload.previousStatus);
                break;
            case 'cancelDeleteRequest':
                result = cancelDeleteRequest(token, payload.rowIndex);
                break;
            case 'approveDelete':
                result = approveDelete(token, payload.rowIndex);
                break;
            case 'rejectDelete':
                result = rejectDelete(token, payload.rowIndex, payload.reason);
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

            // ---- CHAT ----
            case 'sendChatMessage':
                result = sendChatMessage(token, payload.toEmail, payload.message);
                break;
            case 'markChatRead':
                result = markChatRead(token, payload.otherEmail);
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

            // ---- ANNOUNCEMENTS ----
            case 'createAnnouncement':
                result = createAnnouncement(token, payload);
                break;
            case 'dismissAnnouncement':
                result = dismissAnnouncement(token, payload.announcementId);
                break;

            default:
                result = { success: false, error: 'Unknown action: ' + action };
        }

        return ContentService.createTextOutput(JSON.stringify(result))
            .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
            success: false,
            error: err.message
        })).setMimeType(ContentService.MimeType.JSON);
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


// =============================================================================
// 3. VOUCHERS.gs — GET VOUCHERS WITH PAGINATION
// =============================================================================
// Replace your existing getVouchers function with this:

/**
 * Fetches vouchers with support for pagination and filters.
 * 
 * params: { year, page, pageSize, filters }
 * filters (optional): { status, payee, controlNumber, month, search }
 */
function getVouchers(token, params) {
    // Null-safe session check (fixes "Cannot read properties of null (reading 'role')")
    const session = getSession(token);
    if (!session) {
        return { success: false, error: 'Session expired' };
    }

    try {
        const year = (params && params.year) || '2026';
        const page = Math.max(1, parseInt(params && params.page) || 1);
        const pageSize = Math.min(200, Math.max(1, parseInt(params && params.pageSize) || 50));
        const filters = (params && params.filters) || null;

        // Determine sheet name from year
        const sheetName = getVoucherSheetName_(year);
        const sheet = getSheet(sheetName);
        const lastRow = sheet.getLastRow();

        if (lastRow < 2) {
            return {
                success: true,
                vouchers: [],
                pagination: { page: 1, pageSize: pageSize, total: 0, totalPages: 0 }
            };
        }

        const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        const cols = CONFIG.VOUCHER_COLUMNS;

        // Apply filters
        let filtered = [];
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const rowIndex = i + 2; // 1-based, accounting for header

            // Build voucher object
            const voucher = {
                rowIndex: rowIndex,
                status: String(row[cols.STATUS - 1] || ''),
                pmtMonth: String(row[cols.PMT_MONTH - 1] || ''),
                payee: String(row[cols.PAYEE - 1] || ''),
                accountOrMail: String(row[cols.ACCOUNT_OR_MAIL - 1] || ''),
                particular: String(row[cols.PARTICULAR - 1] || ''),
                contractSum: parseFloat(row[cols.CONTRACT_SUM - 1]) || 0,
                grossAmount: parseFloat(row[cols.GROSS_AMOUNT - 1]) || 0,
                net: parseFloat(row[cols.NET - 1]) || 0,
                vat: parseFloat(row[cols.VAT - 1]) || 0,
                wht: parseFloat(row[cols.WHT - 1]) || 0,
                stampDuty: parseFloat(row[cols.STAMP_DUTY - 1]) || 0,
                categories: String(row[cols.CATEGORIES - 1] || ''),
                totalGross: parseFloat(row[cols.TOTAL_GROSS - 1]) || 0,
                controlNumber: String(row[cols.CONTROL_NUMBER - 1] || ''),
                oldVoucherNumber: String(row[cols.OLD_VOUCHER_NUMBER - 1] || ''),
                date: row[cols.DATE - 1] ? new Date(row[cols.DATE - 1]).toISOString() : '',
                accountType: String(row[cols.ACCOUNT_TYPE - 1] || ''),
                createdAt: row[cols.CREATED_AT - 1] ? new Date(row[cols.CREATED_AT - 1]).toISOString() : '',
                releasedAt: row[cols.RELEASED_AT - 1] ? new Date(row[cols.RELEASED_AT - 1]).toISOString() : ''
            };

            // Apply filters if provided
            if (filters) {
                if (typeof filters === 'string') {
                    try { filters = JSON.parse(filters); } catch (_) { filters = null; }
                }
                if (filters) {
                    if (filters.status && voucher.status.toLowerCase() !== filters.status.toLowerCase()) continue;
                    if (filters.month && voucher.pmtMonth.toLowerCase() !== filters.month.toLowerCase()) continue;
                    if (filters.controlNumber && !voucher.controlNumber.toLowerCase().includes(filters.controlNumber.toLowerCase())) continue;
                    if (filters.payee && !voucher.payee.toLowerCase().includes(filters.payee.toLowerCase())) continue;
                    if (filters.search) {
                        const searchLower = filters.search.toLowerCase();
                        const searchable = [voucher.payee, voucher.particular, voucher.controlNumber, voucher.accountOrMail].join(' ').toLowerCase();
                        if (!searchable.includes(searchLower)) continue;
                    }
                }
            }

            filtered.push(voucher);
        }

        // Pagination
        const total = filtered.length;
        const totalPages = Math.ceil(total / pageSize);
        const startIndex = (page - 1) * pageSize;
        const endIndex = Math.min(startIndex + pageSize, total);
        const pageData = filtered.slice(startIndex, endIndex);

        return {
            success: true,
            vouchers: pageData,
            pagination: {
                page: page,
                pageSize: pageSize,
                total: total,
                totalPages: totalPages
            }
        };

    } catch (err) {
        return { success: false, error: 'Failed to fetch vouchers: ' + err.message };
    }
}


/**
 * Helper: maps a year string to the sheet name.
 */
function getVoucherSheetName_(year) {
    switch (String(year)) {
        case '2026': return CONFIG.SHEETS.VOUCHERS_2026;
        case '2025': return CONFIG.SHEETS.VOUCHERS_2025;
        case '2024': return CONFIG.SHEETS.VOUCHERS_2024;
        case '2023': return CONFIG.SHEETS.VOUCHERS_2023;
        default: return CONFIG.SHEETS.VOUCHERS_BEFORE_2023;
    }
}


/**
 * Gets a single voucher by row index.
 */
function getVoucherByRow(token, rowIndex, year) {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };

    try {
        rowIndex = parseInt(rowIndex);
        if (!rowIndex || rowIndex < 2) return { success: false, error: 'Invalid row index' };

        year = year || '2026';
        const sheetName = getVoucherSheetName_(year);
        const sheet = getSheet(sheetName);

        if (rowIndex > sheet.getLastRow()) {
            return { success: false, error: 'Voucher not found' };
        }

        const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
        const cols = CONFIG.VOUCHER_COLUMNS;

        const voucher = {
            rowIndex: rowIndex,
            status: String(row[cols.STATUS - 1] || ''),
            pmtMonth: String(row[cols.PMT_MONTH - 1] || ''),
            payee: String(row[cols.PAYEE - 1] || ''),
            accountOrMail: String(row[cols.ACCOUNT_OR_MAIL - 1] || ''),
            particular: String(row[cols.PARTICULAR - 1] || ''),
            contractSum: parseFloat(row[cols.CONTRACT_SUM - 1]) || 0,
            grossAmount: parseFloat(row[cols.GROSS_AMOUNT - 1]) || 0,
            net: parseFloat(row[cols.NET - 1]) || 0,
            vat: parseFloat(row[cols.VAT - 1]) || 0,
            wht: parseFloat(row[cols.WHT - 1]) || 0,
            stampDuty: parseFloat(row[cols.STAMP_DUTY - 1]) || 0,
            categories: String(row[cols.CATEGORIES - 1] || ''),
            totalGross: parseFloat(row[cols.TOTAL_GROSS - 1]) || 0,
            controlNumber: String(row[cols.CONTROL_NUMBER - 1] || ''),
            oldVoucherNumber: String(row[cols.OLD_VOUCHER_NUMBER - 1] || ''),
            date: row[cols.DATE - 1] ? new Date(row[cols.DATE - 1]).toISOString() : '',
            accountType: String(row[cols.ACCOUNT_TYPE - 1] || ''),
            createdAt: row[cols.CREATED_AT - 1] ? new Date(row[cols.CREATED_AT - 1]).toISOString() : '',
            releasedAt: row[cols.RELEASED_AT - 1] ? new Date(row[cols.RELEASED_AT - 1]).toISOString() : ''
        };

        return { success: true, voucher: voucher };
    } catch (err) {
        return { success: false, error: 'Failed to get voucher: ' + err.message };
    }
}


// =============================================================================
// 4. ROLE SECURITY — NULL-SAFE PATTERN
// =============================================================================
// Every function that accesses session.role MUST follow this pattern:
//
//   function someAction(token, ...) {
//     const session = getSession(token);
//     if (!session) return { success: false, error: 'Session expired' };
//     
//     // NOW safe to access session.role
//     if (!hasPermission(session.role, [CONFIG.ROLES.ADMIN])) {
//       return { success: false, error: 'Unauthorized' };
//     }
//     ...
//   }
//
// The role security test fails because getSession returns null and the code
// does `session.role` without checking if session is null first.


// =============================================================================
// 5. USERS.gs — GET USERS (example of null-safe pattern)
// =============================================================================

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


// =============================================================================
// 6. GET ROLE PERMISSIONS (example)
// =============================================================================

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
        }
    };

    const rolePerms = permissions[session.role] || {};
    return { success: true, permissions: rolePerms, role: session.role };
}
