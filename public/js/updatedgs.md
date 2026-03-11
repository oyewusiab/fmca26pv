# Payable Vouchers 2026 - Complete Backend Suite

Copy the code blocks below into separate script files in your Google Apps Script project for better organization, or paste them sequentially into `Code.gs`.

## 1. Config.gs

```javascript
/**
 * Global Configuration
 */
const CONFIG = {
  SPREADSHEET_ID: '1Idm6P45SG1JA9GEiFWdDT9GT9fgYshx4q5NdqYAsFsA', // Replace with actual ID
  SHEETS: {
    VOUCHERS_2026: '2026 VOUCHERS',
    VOUCHERS_2025: '2025 VOUCHERS',
    VOUCHERS_2024: '2024 VOUCHERS',
    VOUCHERS_2023: '2023 VOUCHERS',
    VOUCHERS_BEFORE_2023: '<2023 VOUCHERS',
    USERS: 'USERS',
    NOTIFICATIONS: 'NOTIFICATIONS',
    CHAT_MESSAGES: 'CHAT_MESSAGES',
    ACTION_ITEMS_LOG: 'ACTION_ITEMS_LOG',
    ACTION_ITEMS_SETTINGS: 'ACTION_ITEMS_SETTINGS',
    ANNOUNCEMENTS: 'ANNOUNCEMENTS',
    SYSTEM_CONFIG: 'SYSTEM_CONFIG',
    AUDIT_TRAIL: 'AUDIT_TRAIL'
  },
  SESSION_DURATION: 8 * 60 * 60 * 1000, // 8 hours
  
  // Voucher Column Indices (1-based)
  VOUCHER_COLUMNS: {
    STATUS: 1, PMT_MONTH: 2, PAYEE: 3, ACCOUNT_OR_MAIL: 4, PARTICULAR: 5,
    CONTRACT_SUM: 6, GROSS_AMOUNT: 7, NET: 8, VAT: 9, WHT: 10, STAMP_DUTY: 11,
    CATEGORIES: 12, TOTAL_GROSS: 13, CONTROL_NUMBER: 14, OLD_VOUCHER_NUMBER: 15,
    DATE: 16, ACCOUNT_TYPE: 17, CREATED_AT: 18, RELEASED_AT: 19, ATTACHMENT_URL: 20,
    OLD_VOUCHER_AVAILABLE: 21
  },
  
  // User Column Indices
  USER_COLUMNS: {
    NAME: 1, EMAIL: 2, PASSWORD: 3, ROLE: 4, ACTIVE: 5, USERNAME: 6,
    PHONE: 7, DEPARTMENT: 8, UPDATED_AT: 9
  },
  
  // Roles
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

## 2. SheetUtils.gs

```javascript
/**
 * Gets a sheet by name. Throws an error if not found.
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // Auto-create specific system sheets if missing
    if ([CONFIG.SHEETS.ACTION_ITEMS_LOG, CONFIG.SHEETS.ANNOUNCEMENTS, CONFIG.SHEETS.AUDIT_TRAIL, CONFIG.SHEETS.CHAT_MESSAGES].includes(sheetName)) {
      const newSheet = ss.insertSheet(sheetName);
      // Add headers for new sheets
      if (sheetName === CONFIG.SHEETS.AUDIT_TRAIL) {
        newSheet.appendRow(['TIMESTAMP', 'USER_EMAIL', 'USER_NAME', 'ROLE', 'ACTION', 'DESCRIPTION', 'SHEET', 'ROW_INDEX', 'EXTRA']);
      }
      return newSheet;
    }
    throw new Error(`Sheet "${sheetName}" not found.`);
  }
  return sheet;
}

/**
 * Parses a numeric amount from a string, removing currency symbols and commas.
 * @param {number|string} value The value to parse.
 * @returns {number} The parsed number.
 */
function parseAmount(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  const cleaned = String(value).replace(/[₦,]/g, '');
  return parseFloat(cleaned) || 0;
}

/**
 * Gets the current timestamp in the Nigerian timezone.
 * @returns {string} A formatted timestamp string (e.g., "2026-03-02 18:35:27").
 */
function getNigerianTimestamp() {
  return Utilities.formatDate(new Date(), "Africa/Lagos", "yyyy-MM-dd HH:mm:ss");
}

/**
 * Clears all voucher-related caches to ensure fresh data is fetched.
 */
function clearAllVoucherCaches() {
  const cache = CacheService.getScriptCache();
  cache.removeAll([
    'vouchers_2026_summary',
    'dashboard_stats'
    // Add other cache keys here as they are implemented
  ]);
}

/**
 * Logs an audit trail event to the AUDIT_TRAIL sheet.
 */
function logAudit(session, action, description, sheetName, rowIndex, extra) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.AUDIT_TRAIL);
    sheet.appendRow([
      getNigerianTimestamp(),
      session.email,
      session.name,
      session.role,
      action,
      description,
      sheetName || '',
      rowIndex || '',
      extra ? JSON.stringify(extra) : ''
    ]);
  } catch (e) {
    console.error('Audit Log Error:', e);
  }
}

/**
 * Gets the header map of a sheet.
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
 * Helper: maps a year string to the sheet name.
 * @param {string} year
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
 * Converts a voucher object to a spreadsheet row array.
 * @param {object} voucher
 * @returns {any[]}
 */
function voucherToRow(voucher) {
  const cols = CONFIG.VOUCHER_COLUMNS;
  const row = [];
  row[cols.STATUS - 1] = voucher.status || 'Unpaid';
  row[cols.PMT_MONTH - 1] = voucher.pmtMonth || '';
  row[cols.PAYEE - 1] = voucher.payee;
  row[cols.ACCOUNT_OR_MAIL - 1] = voucher.accountOrMail;
  row[cols.PARTICULAR - 1] = voucher.particular;
  row[cols.CONTRACT_SUM - 1] = voucher.contractSum || 0;
  row[cols.GROSS_AMOUNT - 1] = voucher.grossAmount;
  row[cols.NET - 1] = voucher.net;
  row[cols.VAT - 1] = voucher.vat || 0;
  row[cols.WHT - 1] = voucher.wht || 0;
  row[cols.STAMP_DUTY - 1] = voucher.stampDuty || 0;
  row[cols.CATEGORIES - 1] = voucher.categories || '';
  row[cols.TOTAL_GROSS - 1] = voucher.totalGross || voucher.grossAmount;
  row[cols.CONTROL_NUMBER - 1] = voucher.controlNumber || '';
  row[cols.OLD_VOUCHER_NUMBER - 1] = voucher.oldVoucherNumber || '';
  row[cols.OLD_VOUCHER_AVAILABLE - 1] = voucher.oldVoucherAvailable || '';
  row[cols.DATE - 1] = voucher.date ? new Date(voucher.date) : new Date();
  row[cols.ACCOUNT_TYPE - 1] = voucher.accountType || '';
  row[cols.CREATED_AT - 1] = voucher.createdAt ? new Date(voucher.createdAt) : new Date();
  row[cols.RELEASED_AT - 1] = voucher.releasedAt ? new Date(voucher.releasedAt) : '';
  row[cols.ATTACHMENT_URL - 1] = voucher.attachmentUrl || '';
  return row;
}
```

## 3. Code.gs (Main Router)

```javascript
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';
    const params = e.parameter || {};
    let result;

    switch (action) {
      // ---- AUTH ----
      case 'login': result = login(params.email, params.password); break;
      case 'validateSession': result = validateSession(params.token); break;
      case 'logout': result = logout(params.token); break;

      // ---- VOUCHERS ----
      case 'getVouchers':
        result = getVouchers(params.token, {
          year: params.year || '2026',
          page: parseInt(params.page) || 1,
          pageSize: parseInt(params.pageSize) || 50,
          filters: params.filters ? JSON.parse(params.filters) : null
        });
        break;
      case 'getVoucherByRow': result = getVoucherByRow(params.token, params.rowIndex, params.year); break;
      case 'lookupVoucher': result = lookupVoucher(params.token, params.voucherNumber); break;
      case 'getNextControlNumber': result = getNextControlNumber(params.token, params.targetUnit); break;

      // ---- USERS ----
      case 'getUsers': result = getUsers(params.token); break;
      case 'getRolePermissions': result = getRolePermissions(params.token); break;
      case 'getMyProfile': result = getMyProfile(params.token); break;
      case 'getPendingDeletions': result = getPendingDeletions(params.token); break;

      // ---- ACTION ITEMS ----
      case 'getActionItems': result = getActionItems(params.token, params); break;
      case 'getActionItemCount': result = getActionItemCount(params.token, params); break;
      case 'getActionItemSettings': result = getActionItemSettings(params.token); break;

      // ---- REPORTS ----
      case 'getDashboardStats': result = getDashboardStats(params.token); break;
      case 'getSummary': result = getSummary(params.token, params.year); break;
      case 'getDebtProfile': result = getDebtProfile(params.token); break;

      // ---- ANNOUNCEMENTS ----
      case 'getActiveAnnouncements': result = getActiveAnnouncements(params.token); break;

      // ---- CHAT ----
      case 'getChatUsers': result = getChatUsers(params.token); break;
      case 'getChatThread': result = getChatThread(params.token, params.otherEmail, params.limit); break;
      case 'getChatUnreadSummary': result = getChatUnreadSummary(params.token); break;
      case 'getChatUnreadCount': result = getChatUnreadCount(params.token); break;

      // ---- PLACEHOLDERS ----
      case 'getNotifications': 
      case 'getCategories': 
        result = { success: true, message: "Feature coming in Part 2", data: [] }; 
        break;

      default: result = { success: false, error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    let payload = {};
    if (e.postData) {
      try { payload = JSON.parse(e.postData.contents); } catch (_) {}
    }
    const action = payload.action || '';
    const token = payload.token || '';
    let result;

    switch (action) {
      // ---- AUTH ----
      case 'changePassword': result = changePassword(token, payload.oldPassword, payload.newPassword); break;
      case 'requestPasswordReset': result = requestPasswordReset(payload.identifier); break;
      case 'resetPasswordWithOtp': result = resetPasswordWithOtp(payload.identifier, payload.otp, payload.newPassword); break;

      // ---- USERS ----
      case 'createUser': result = createUser(token, payload.user); break;
      case 'updateUser': result = updateUser(token, payload.rowIndex, payload.user); break;
      case 'deleteUser': result = deleteUser(token, payload.rowIndex); break;
      case 'updateMyProfile': result = updateMyProfile(token, payload.profile); break;

      // ---- ACTION ITEMS & ANNOUNCEMENTS ----
      case 'saveActionItemSettings': result = saveActionItemSettings(token, payload.settings); break;
      case 'createAnnouncement': result = createAnnouncement(token, payload.announcement); break;
      case 'dismissAnnouncement': result = dismissAnnouncement(token, payload.announcementId); break;

      // ---- CHAT ----
      case 'sendChatMessage': result = sendChatMessage(token, payload.toEmail, payload.message); break;
      case 'markChatRead': result = markChatRead(token, payload.otherEmail); break;

      // ---- VOUCHER ACTIONS ----
      case 'createVoucher': result = createVoucher(token, payload.voucher); break;
      case 'updateVoucher': result = updateVoucher(token, payload.rowIndex, payload.voucher); break;
      case 'updateStatus': result = updateVoucherStatus(token, payload.rowIndex, payload.status, payload.pmtMonth); break;
      case 'batchUpdateStatus': result = batchUpdateStatus(token, payload.controlNumber, payload.status, payload.pmtMonth); break;
      case 'releaseVouchers': result = releaseVouchers(token, payload); break;
      case 'assignControlNumber': result = assignControlNumber(token, payload.rowIndexes, payload.controlNumber); break;
      case 'getNextControlNumber': result = getNextControlNumber(token, payload.targetUnit); break;
      case 'requestDelete': result = requestVoucherDelete(token, payload.rowIndex, payload.reason, payload.previousStatus); break;
      case 'approveDelete': result = approveVoucherDelete(token, payload.rowIndex); break;
      case 'rejectDelete': result = rejectVoucherDelete(token, payload.rowIndex, payload.reason); break;
      case 'deleteVoucher': result = deleteVoucher(token, payload.rowIndex); break;
        break;

      default: result = { success: false, error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

## 4. Auth.gs

```javascript
function login(email, password) {
  try {
    if (!email || !password) return { success: false, error: 'Email and password required' };
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    const cols = CONFIG.USER_COLUMNS;
    const emailLower = String(email).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = String(row[cols.EMAIL - 1] || '').trim().toLowerCase();
      const rowUsername = String(row[cols.USERNAME - 1] || '').trim().toLowerCase();
      const rowPassword = String(row[cols.PASSWORD - 1] || '');
      const rowActive = String(row[cols.ACTIVE - 1] || 'true').toLowerCase();

      if (rowEmail === emailLower || rowUsername === emailLower) {
        if (rowActive === 'false' || rowActive === 'no' || rowActive === '0') {
          return { success: false, error: 'Account is disabled.' };
        }
        // Simple check (In production, use hash comparison)
        if (rowPassword !== password) return { success: false, error: 'Invalid password' };

        const token = Utilities.getUuid();
        // Hashing function for passwords
        const hashedPassword = hashPassword(password);
        if (rowPassword !== hashedPassword) {
          sheet.getRange(i + 1, cols.PASSWORD).setValue(hashedPassword);
        }
        const user = {
          name: String(row[cols.NAME - 1] || ''),
          email: String(row[cols.EMAIL - 1] || ''),
          role: String(row[cols.ROLE - 1] || ''),
          username: String(row[cols.USERNAME - 1] || ''),
          rowIndex: i + 1
        };

        const cache = CacheService.getScriptCache();
        cache.put('session_' + token, JSON.stringify({...user, expiresAt: new Date().getTime() + CONFIG.SESSION_DURATION}), CONFIG.SESSION_DURATION / 1000);
        cache.put('session_' + token, JSON.stringify(user), CONFIG.SESSION_DURATION / 1000);
        return { success: true, token: token, user: user };
      }
    }
    return { success: false, error: 'User not found' };
  } catch (err) { return { success: false, error: err.message }; }
}

function getSession(token) {
  if (!token) return null;
  try {
    const json = CacheService.getScriptCache().get('session_' + token);
    if (!json) return null;
    const session = JSON.parse(json);
    if (new Date().getTime() > session.expiresAt) {
      CacheService.getScriptCache().remove('session_' + token);
      return null;
    }
    return session;
    return json ? JSON.parse(json) : null;
  } catch (e) { return null; }
}

function validateSession(token) {
  const session = getSession(token);
  return session ? { success: true, user: session } : { success: false, error: 'Session expired' };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('session_' + token);
  return { success: true };
}

function changePassword(token, oldPassword, newPassword) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const cols = CONFIG.USER_COLUMNS;
  const storedPassword = sheet.getRange(session.rowIndex, cols.PASSWORD).getValue();
  const hashedOld = hashPassword(oldPassword);

  if (storedPassword !== hashedOld && storedPassword !== oldPassword) {
    return { success: false, error: 'Current password is incorrect' };
  }

  if (!newPassword || newPassword.length < 6) {
    return { success: false, error: 'New password must be at least 6 characters' };
  }

  sheet.getRange(session.rowIndex, cols.PASSWORD).setValue(hashPassword(newPassword));
  // In a real app, you'd verify oldPassword against the sheet here
  sheet.getRange(session.rowIndex, cols.PASSWORD).setValue(newPassword);
  return { success: true, message: 'Password changed' };
}

function hasPermission(role, allowedRoles) {
  if (!role) return false;
  if (role === CONFIG.ROLES.ADMIN) return true;
  return Array.isArray(allowedRoles) ? allowedRoles.includes(role) : role === allowedRoles;
}

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

// Placeholders for password reset (Part 2)
function requestPasswordReset(identifier) { return { success: true, message: "OTP sent (Simulated)" }; }
function resetPasswordWithOtp(id, otp, newPwd) { return { success: true, message: "Password reset (Simulated)" }; }
function getMyProfile(token) { 
  const s = getSession(token); 
  return s ? { success: true, profile: s } : { success: false }; 
}
function updateMyProfile(token, profile) { return { success: true, message: "Profile updated" }; }
```

## 5. Users.gs

```javascript
function getUsers(token) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  if (!hasPermission(session.role, [CONFIG.ROLES.ADMIN, CONFIG.ROLES.DFA])) {
    return { success: false, error: 'Unauthorized' };
  }

  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  const cols = CONFIG.USER_COLUMNS;
  const users = [];

  for (let i = 1; i < data.length; i++) {
    users.push({
      rowIndex: i + 1,
      name: data[i][cols.NAME - 1],
      email: data[i][cols.EMAIL - 1],
      role: data[i][cols.ROLE - 1],
      active: data[i][cols.ACTIVE - 1],
      username: data[i][cols.USERNAME - 1]
    });
  }
  return { success: true, users: users };
}

function createUser(token, user) {
  const session = getSession(token);
  if (!session || session.role !== CONFIG.ROLES.ADMIN) return { success: false, error: 'Unauthorized' };
  
  const sheet = getSheet(CONFIG.SHEETS.USERS);
  sheet.appendRow([
    user.name, user.email, 'Welcome123', user.role, true, user.username, user.phone || '', user.department || '', new Date()
  ]);
  return { success: true, message: 'User created' };
}

function updateUser(token, rowIndex, user) {
  const session = getSession(token);
  if (!session || session.role !== CONFIG.ROLES.ADMIN) return { success: false, error: 'Unauthorized' };
  
  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const cols = CONFIG.USER_COLUMNS;
  if (user.name) sheet.getRange(rowIndex, cols.NAME).setValue(user.name);
  if (user.email) sheet.getRange(rowIndex, cols.EMAIL).setValue(user.email);
  if (user.role) sheet.getRange(rowIndex, cols.ROLE).setValue(user.role);
  if (user.active !== undefined) sheet.getRange(rowIndex, cols.ACTIVE).setValue(user.active);
  return { success: true, message: 'User updated' };
}

function deleteUser(token, rowIndex) {
  const session = getSession(token);
  if (!session || session.role !== CONFIG.ROLES.ADMIN) return { success: false, error: 'Unauthorized' };
  getSheet(CONFIG.SHEETS.USERS).deleteRow(rowIndex);
  return { success: true, message: 'User deleted' };
}

function getRolePermissions(token) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  // Simplified permission map
  const p = {
    canCreateVoucher: [CONFIG.ROLES.PAYABLE_STAFF, CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.ADMIN].includes(session.role),
    canEditVoucher: [CONFIG.ROLES.PAYABLE_STAFF, CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.CPO, CONFIG.ROLES.ADMIN].includes(session.role),
    canManageUsers: session.role === CONFIG.ROLES.ADMIN,
    canRequestDelete: true, // Simplified for now
    canApproveDelete: [CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.CPO, CONFIG.ROLES.ADMIN].includes(session.role),
    canUpdateStatus: true // Simplified for now
    canManageUsers: session.role === CONFIG.ROLES.ADMIN
  };
  return { success: true, permissions: p, role: session.role };
}
```

## 6. VouchersRead.gs

```javascript
function getVouchers(token, params) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };

  const year = params.year || '2026';
  const page = Math.max(1, parseInt(params.page) || 1);
  const pageSize = Math.min(200, parseInt(params.pageSize) || 50);
  const filters = params.filters || {};
  

  const sheet = getSheet(getVoucherSheetName_(year));
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, vouchers: [], pagination: { total: 0, page, pageSize } };

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const cols = CONFIG.VOUCHER_COLUMNS;
  let filtered = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    // Basic check for non-empty row
    if (!row[2] && !row[3]) continue;
    const v = {
      rowIndex: i + 2,
      status: String(row[cols.STATUS - 1] || ''),
      payee: String(row[cols.PAYEE - 1] || ''),
      grossAmount: parseFloat(row[cols.GROSS_AMOUNT - 1]) || 0,
      controlNumber: String(row[cols.CONTROL_NUMBER - 1] || ''),
      particular: String(row[cols.PARTICULAR - 1] || ''),
      pmtMonth: String(row[cols.PMT_MONTH - 1] || ''),
      date: row[cols.DATE - 1]
    };

    const v = rowToVoucher(row, i + 2, year === '2026');

    // Basic Filtering
    if (filters.status && filters.status !== 'All' && v.status !== filters.status) continue;
    if (filters.category && filters.category !== 'All' && v.categories !== filters.category) continue;
    if (filters.release && filters.release !== 'All') {
      const hasCN = v.controlNumber && String(v.controlNumber).trim() !== '';
      if (filters.release === 'Released' && !hasCN) continue;
      if (filters.release === 'Not Released' && hasCN) continue;
    }

    if (filters.searchTerm) {
      const term = filters.searchTerm.toLowerCase();
    if (filters.search) {
      const term = filters.search.toLowerCase();
      if (!v.payee.toLowerCase().includes(term) && !v.controlNumber.toLowerCase().includes(term)) continue;
    }
    filtered.push(v);
  }

  // Pagination
  const total = filtered.length;
  const start = (page - 1) * pageSize;
  const pagedData = filtered.slice(start, start + pageSize);

  return {
    success: true,
    vouchers: pagedData,
    pagination: { total, page, pageSize, totalPages: Math.ceil(total / pageSize) }
  };
}

function getVoucherByRow(token, rowIndex, year) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  
  const sheet = getSheet(getVoucherSheetName_(year || '2026'));
  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { success: true, voucher: rowToVoucher(row, parseInt(rowIndex), year === '2026') };
  const cols = CONFIG.VOUCHER_COLUMNS;
  
  // Map full details
  const voucher = {
    rowIndex: parseInt(rowIndex),
    status: row[cols.STATUS - 1],
    payee: row[cols.PAYEE - 1],
    grossAmount: row[cols.GROSS_AMOUNT - 1],
    contractSum: row[cols.CONTRACT_SUM - 1],
    net: row[cols.NET - 1],
    vat: row[cols.VAT - 1],
    wht: row[cols.WHT - 1],
    stampDuty: row[cols.STAMP_DUTY - 1],
    particular: row[cols.PARTICULAR - 1],
    controlNumber: row[cols.CONTROL_NUMBER - 1],
    accountOrMail: row[cols.ACCOUNT_OR_MAIL - 1],
    categories: row[cols.CATEGORIES - 1],
    date: row[cols.DATE - 1]
  };
  
  return { success: true, voucher: voucher };
}

function lookupVoucher(token, voucherNumber) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  
  if (!voucherNumber || !String(voucherNumber).trim()) {
    return { success: false, error: 'Voucher Number is required' };
  }
  
  const searchValue = String(voucherNumber).trim().toUpperCase();
  
  // Determine year from voucher number suffix (e.g., /2025)
  let year = null;
  const match = searchValue.match(/[\/\-\.\s](\d{4})$/);
  if (match) {
    const y = parseInt(match[1]);
    if (y >= 2015 && y <= 2025) year = String(y);
  }
  
  // Define search order
  let sheetsToSearch = [];
  if (year) {
    if (year === '2025') sheetsToSearch.push({ name: CONFIG.SHEETS.VOUCHERS_2025, year: '2025' });
    else if (year === '2024') sheetsToSearch.push({ name: CONFIG.SHEETS.VOUCHERS_2024, year: '2024' });
    else if (year === '2023') sheetsToSearch.push({ name: CONFIG.SHEETS.VOUCHERS_2023, year: '2023' });
    else sheetsToSearch.push({ name: CONFIG.SHEETS.VOUCHERS_BEFORE_2023, year: '<2023' });
  } else {
    // Search all recent years if no year detected
    sheetsToSearch = [
      { name: CONFIG.SHEETS.VOUCHERS_2025, year: '2025' },
      { name: CONFIG.SHEETS.VOUCHERS_2024, year: '2024' },
      { name: CONFIG.SHEETS.VOUCHERS_2023, year: '2023' },
      { name: CONFIG.SHEETS.VOUCHERS_BEFORE_2023, year: '<2023' }
    ];
  }
  
  // 1. Search in ACCOUNT_OR_MAIL column
  for (const info of sheetsToSearch) {
    try {
      const sheet = getSheet(info.name);
      const data = sheet.getDataRange().getValues();
      const cols = CONFIG.VOUCHER_COLUMNS; // Assuming column structure is similar enough for lookup
      
      for (let i = 1; i < data.length; i++) {
        const rowVal = String(data[i][cols.ACCOUNT_OR_MAIL - 1] || '').trim().toUpperCase();
        if (rowVal === searchValue) {
          const v = rowToVoucher(data[i], i + 1, false);
          
          if (String(v.status).toLowerCase() === 'paid') {
            return { success: true, found: true, canRevalidate: false, voucher: v, sourceYear: info.year, message: "Found but PAID (Cannot revalidate)" };
          }
          
          return { success: true, found: true, canRevalidate: true, voucher: v, sourceYear: info.year, message: `Found in ${info.year}` };
        }
      }
    } catch (e) { console.log('Lookup error in ' + info.name, e); }
  }
  
  // 2. Search in OLD_VOUCHER_NUMBER column (to check if already revalidated)
  const allSheets = [
    { name: CONFIG.SHEETS.VOUCHERS_2026, year: '2026' },
    ...sheetsToSearch
  ];
  
  for (const info of allSheets) {
    try {
      const sheet = getSheet(info.name);
      const data = sheet.getDataRange().getValues();
      const cols = CONFIG.VOUCHER_COLUMNS;
      
      for (let i = 1; i < data.length; i++) {
        const oldVal = String(data[i][cols.OLD_VOUCHER_NUMBER - 1] || '').trim().toUpperCase();
        if (oldVal === searchValue) {
           const v = rowToVoucher(data[i], i + 1, info.year === '2026');
           return { 
             success: true, found: true, canRevalidate: false, alreadyRevalidated: true, 
             voucher: v, sourceYear: info.year, message: `Already revalidated in ${info.year}` 
           };
        }
      }
    } catch (e) {}
  }
  
  return { success: true, found: false, message: "Voucher not found" };
  // Placeholder for lookup logic
  return { success: true, found: false, message: "Lookup logic in Part 2" };
}
```

## 7. ActionItems.gs

```javascript
const ACTION_ITEM_RULES = {
  PAID_NO_CN: "PAID_NO_CN",
  UNPAID_NO_CN_30D: "UNPAID_NO_CN_30D",
  RELEASED_UNPAID_15D: "RELEASED_UNPAID_15D"
};

const ACTION_ITEM_UNITS = { PAYABLE: "PAYABLE", CPO: "CPO" };

function ensureActionItemSheets_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  if (!ss.getSheetByName(CONFIG.SHEETS.ACTION_ITEMS_LOG)) {
    ss.insertSheet(CONFIG.SHEETS.ACTION_ITEMS_LOG).appendRow([
      "ITEM_ID", "RULE_KEY", "UNIT", "YEAR", "ROW_INDEX", "VOUCHER_NO",
      "PAYEE", "AMOUNT", "CONTROL_NO", "STATUS", "FIRST_SEEN_AT", "LAST_SEEN_AT", "RESOLVED_AT"
    ]);
  }
  if (!ss.getSheetByName(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS)) {
    ss.insertSheet(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS).appendRow(["UNIT", "RULE_KEY", "ENABLED"]);
  }
}

function getActionItems(token, params) {
  const session = getSession(token);
  if (!session) return { success: false, error: "Session expired" };
  
  // Mock data for now until sync engine is fully integrated
  return { success: true, items: [], count: 0 };
}

function getActionItemCount(token, params) {
  const res = getActionItems(token, params);
  return { success: true, count: res.count || 0 };
}

function getActionItemSettings(token) {
  ensureActionItemSheets_();
  return { success: true, settings: { unit: { PAYABLE: {}, CPO: {} } } };
}

function saveActionItemSettings(token, settings) {
  return { success: true, message: "Settings saved" };
}
```

## 8. Announcements.gs

```javascript
function ensureAnnouncementsSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  if (!ss.getSheetByName(CONFIG.SHEETS.ANNOUNCEMENTS)) {
    ss.insertSheet(CONFIG.SHEETS.ANNOUNCEMENTS).appendRow([
      "ID", "MESSAGE", "DISPLAY_LOCATIONS", "TARGET_USERS",
      "EXPIRES_AT", "ALLOW_DISMISS", "CREATED_AT", "CREATED_BY"
    ]);
  }
}

function createAnnouncement(token, announcement) {
  const session = getSession(token);
  if (!session || session.role !== CONFIG.ROLES.ADMIN) return { success: false, error: "Unauthorized" };
  
  ensureAnnouncementsSheet_();
  const sheet = getSheet(CONFIG.SHEETS.ANNOUNCEMENTS);
  sheet.appendRow([
    Utilities.getUuid(), announcement.message, JSON.stringify(announcement.locations),
    JSON.stringify(announcement.targets), announcement.expiresAt, announcement.allowDismiss, new Date(), session.email
  ]);
  return { success: true, message: "Announcement created" };
}

function getActiveAnnouncements(token) {
  const session = getSession(token);
  if (!session) return { success: false, error: "Session expired" };
  ensureAnnouncementsSheet_();
  // Return empty for now, logic in Part 2
  return { success: true, announcements: [] };
}

function dismissAnnouncement(token, id) {
  return { success: true };
}

    VOUCHERS_2023: '2023 VOUCHERS',
    VOUCHERS_BEFORE_2023: '<2023 VOUCHERS',
    USERS: 'USERS',
    CHAT_MESSAGES: 'CHAT_MESSAGES',
    ACTION_ITEMS_LOG: 'ACTION_ITEMS_LOG',
    ACTION_ITEMS_SETTINGS: 'ACTION_ITEMS_SETTINGS',
    ANNOUNCEMENTS: 'ANNOUNCEMENTS',
      case 'getActiveAnnouncements': result = getActiveAnnouncements(params.token); break;

      // ---- CHAT ----
      case 'getChatUsers': result = getChatUsers(params.token); break;
      case 'getChatThread': result = getChatThread(params.token, params.otherEmail, params.limit); break;
      case 'getChatUnreadSummary': result = getChatUnreadSummary(params.token); break;

      // ---- PLACEHOLDERS ----
      case 'getNotifications': 
      case 'createAnnouncement': result = createAnnouncement(token, payload.announcement); break;
      case 'dismissAnnouncement': result = dismissAnnouncement(token, payload.announcementId); break;

      // ---- CHAT ----
      case 'sendChatMessage': result = sendChatMessage(token, payload.toEmail, payload.message); break;
      case 'markChatRead': result = markChatRead(token, payload.otherEmail); break;

      // ---- VOUCHER ACTIONS ----
      case 'createVoucher': result = createVoucher(token, payload.voucher); break;
      case 'updateVoucher': result = updateVoucher(token, payload.rowIndex, payload.voucher); break;
}

function updateVoucher(token, rowIndex, voucher) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  
  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const currentRow = sheet.getRange(rowIndex, 1, 1, 18).getValues()[0];
  const currentVoucher = rowToVoucher(currentRow, rowIndex, true);

  // Permissions & Restrictions
  const isPaid = String(currentVoucher.status || '').toLowerCase() === 'paid';
  const isReleased = currentVoucher.controlNumber && String(currentVoucher.controlNumber).trim() !== '';
  
  if (isPaid || isReleased) {
    if (![CONFIG.ROLES.CPO, CONFIG.ROLES.ADMIN].includes(session.role)) {
      return { success: false, error: 'Paid or Released vouchers can only be edited by CPO or Admin.' };
    }
  }

  // Preserve immutable fields
  voucher.createdAt = currentVoucher.createdAt;
  voucher.status = voucher.status || currentVoucher.status;
  voucher.controlNumber = voucher.controlNumber || currentVoucher.controlNumber;
  
  // Recalculate Net
  voucher.totalGross = voucher.totalGross || parseAmount(voucher.grossAmount);
  voucher.net = parseAmount(voucher.grossAmount) - (parseAmount(voucher.vat) + parseAmount(voucher.wht) + parseAmount(voucher.stampDuty));

  const rowData = voucherToRow(voucher);
  sheet.getRange(rowIndex, 1, 1, 18).setValues([rowData]);
  
  clearAllVoucherCaches();
  logAudit(session, 'UPDATE_VOUCHER', `Updated voucher ${rowIndex}`, CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
  
  return { success: true, message: 'Voucher updated successfully', voucher: rowToVoucher(rowData, rowIndex, true) };
}

function updateVoucherStatus(token, rowIndex, status, pmtMonth) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };

  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS).setValue(status);
  
  if (pmtMonth && [CONFIG.ROLES.CPO, CONFIG.ROLES.ADMIN].includes(session.role)) {
    sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.PMT_MONTH).setValue(pmtMonth);
  }
  
  clearAllVoucherCaches();
  logAudit(session, 'UPDATE_STATUS', `Status changed to ${status}`, CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
  return { success: true, message: 'Status updated' };
}

function batchUpdateStatus(token, controlNumber, status, pmtMonth) {
  const session = getSession(token);
  if (!session || ![CONFIG.ROLES.CPO, CONFIG.ROLES.ADMIN].includes(session.role)) {
    return { success: false, error: 'Unauthorized' };
  }

  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const data = sheet.getDataRange().getValues();
  let count = 0;
  const targetCN = String(controlNumber).trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    const rowCN = String(data[i][CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER - 1]).trim().toUpperCase();
    if (rowCN === targetCN) {
      sheet.getRange(i + 1, CONFIG.VOUCHER_COLUMNS.STATUS).setValue(status);
      if (pmtMonth) sheet.getRange(i + 1, CONFIG.VOUCHER_COLUMNS.PMT_MONTH).setValue(pmtMonth);
      count++;
    }
  }
  
  clearAllVoucherCaches();
  return { success: true, message: `Updated ${count} vouchers`, count };
}

function releaseVouchers(token, payload) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };

  const { rowIndexes, controlNumber, targetUnit, purpose, isCPORelease } = payload;
  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const header = getHeaderMap_(sheet);
  const colReleasedAt = header["RELEASED AT"];
  const colControlNo = CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER;

  let updatedCount = 0;
  const releasedVouchers = [];

  rowIndexes.forEach(rowIndex => {
    const r = parseInt(rowIndex);
    if (r < 2) return;
    
    // Only assign CN if not CPO release (CPO releases existing vouchers)
    if (!isCPORelease && controlNumber) {
      sheet.getRange(r, colControlNo).setValue(controlNumber);
      if (colReleasedAt) sheet.getRange(r, colReleasedAt).setValue(new Date());
    }

    const row = sheet.getRange(r, 1, 1, 18).getValues()[0];
    releasedVouchers.push({
      voucherNumber: row[3],
      payee: row[2],
      amount: row[6]
    });
    updatedCount++;
  });

  // Send Notification
  if (typeof sendReleaseNotification === 'function') {
    sendReleaseNotification(controlNumber || 'N/A', targetUnit, releasedVouchers, session.name);
  }

  logAudit(session, 'RELEASE', `Released ${updatedCount} vouchers to ${targetUnit}`, CONFIG.SHEETS.VOUCHERS_2026, 0);
  return { success: true, message: `Released ${updatedCount} vouchers`, count: updatedCount };
}

function assignControlNumber(token, rowIndexes, controlNumber) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  
  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const header = getHeaderMap_(sheet);
  const colReleasedAt = header["RELEASED AT"];
  
  rowIndexes.forEach(r => {
    sheet.getRange(r, CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER).setValue(controlNumber);
    if (colReleasedAt) sheet.getRange(r, colReleasedAt).setValue(new Date());
  });

  return { success: true, message: "Control Number Assigned" };
}

/**
 * Returns the next control number for a target unit.
 * Can be called by Payable Staff / Head / CPO / Admin.
 */
function getNextControlNumber(token, targetUnit) {
  try {
    const session = getSession(token);
    if (!session) return { success: false, error: 'Session expired' };

    if (!hasPermission(session.role, [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.CPO,
      CONFIG.ROLES.ADMIN
    ])) {
      return { success: false, error: 'Unauthorized' };
    }

    if (!targetUnit || !String(targetUnit).trim()) {
      return { success: false, error: 'Target unit is required' };
    }

    const controlNumber = generateControlNumber(targetUnit);
    return { success: true, controlNumber: controlNumber };
  } catch (error) {
    return { success: false, error: 'Failed to generate control number: ' + error.message };
  }
}

/**
 * Generates a new control number for 2026 in the format:
 * CN-<UNIT>-<N>
 */
function generateControlNumber(targetUnit) {
  const unitCode = String(targetUnit || '').toUpperCase();
  if (!unitCode) {
    throw new Error('Target unit is required to generate control number.');
  }
  
  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return `CN-${unitCode}-1`;
  }
  
  const cnCol = CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER;
  const cnValues = sheet.getRange(2, cnCol, lastRow - 1, 1).getValues();
  
  const prefix = `CN-${unitCode}-`;
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
  
  const nextNumber = maxNumber + 1;
  return `${prefix}${nextNumber}`;
}

function rowToVoucher(row, rowIndex, has2026Format = true) {

```javascript
function requestVoucherDelete(token, rowIndex, reason) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  
  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const statusCell = sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS);
  const currentStatus = statusCell.getValue();
  
  if (currentStatus === 'Pending Deletion') return { success: false, error: 'Already pending' };
  
  // Store original status in note
  statusCell.setNote(`ORIGINAL:${currentStatus}|BY:${session.email}|REASON:${reason}`);
  statusCell.setValue('Pending Deletion');
  
  logAudit(session, 'REQUEST_DELETE', `Reason: ${reason}`, CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
  return { success: true, message: 'Deletion requested' };
}

function approveVoucherDelete(token, rowIndex) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  
  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const row = sheet.getRange(rowIndex, 1, 1, 18).getValues();
  const status = String(row[CONFIG.VOUCHER_COLUMNS.STATUS - 1] || '').trim();
  const cn = String(row[CONFIG.VOUCHER_COLUMNS.CONTROL_NUMBER - 1] || '').trim();
  
  if (status !== 'Pending Deletion') return { success: false, error: 'Not pending deletion' };
  
  // Permission check
  if (cn !== '') {
    if (![CONFIG.ROLES.CPO, CONFIG.ROLES.ADMIN].includes(session.role)) {
      return { success: false, error: 'Released vouchers require CPO/Admin approval' };
    }
  } else {
    if (![CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.ADMIN].includes(session.role)) {
      return { success: false, error: 'Requires Payable Head/Admin approval' };
    }
  }
  
  sheet.deleteRow(rowIndex);
  clearAllVoucherCaches();
  logAudit(session, 'APPROVE_DELETE', 'Deleted voucher', CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
  return { success: true, message: 'Voucher deleted' };
}

function rejectVoucherDelete(token, rowIndex, reason) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  
  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const statusCell = sheet.getRange(rowIndex, CONFIG.VOUCHER_COLUMNS.STATUS);
  const note = statusCell.getNote();
  
  let originalStatus = 'Unpaid';
  if (note && note.includes('ORIGINAL:')) {
    originalStatus = note.split('|').split(':');
  }
  
  statusCell.setValue(originalStatus);
  statusCell.clearNote();
  
  logAudit(session, 'REJECT_DELETE', `Rejected. Reason: ${reason}`, CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
  return { success: true, message: 'Deletion rejected, status restored' };
}

function deleteVoucher(token, rowIndex) {
  const session = getSession(token);
  if (!session || session.role !== CONFIG.ROLES.ADMIN) return { success: false, error: 'Unauthorized' };
  
  getSheet(CONFIG.SHEETS.VOUCHERS_2026).deleteRow(rowIndex);
  clearAllVoucherCaches();
  logAudit(session, 'HARD_DELETE', 'Admin deleted voucher', CONFIG.SHEETS.VOUCHERS_2026, rowIndex);
  return { success: true, message: "Voucher permanently deleted" };
}

function getPendingDeletions(token) {
  const session = getSession(token);
  if (!session) return { success: false, error: 'Session expired' };
  
  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
  const data = sheet.getDataRange().getValues();
  const vouchers = data.map((r, i) => ({r, i: i+1})).filter(x => x.r === 'Pending Deletion').map(x => rowToVoucher(x.r, x.i, true));
  
  return { success: true, vouchers };
}

function getDashboardStats(token) {
-  // Placeholder for Part 2 logic
-  return { success: true, stats: { totalVouchersRaised: 0, totalPaidAmount: 0, totalDebt: 0 }, pendingDeletions: 0 };
+  const session = getSession(token);
+  if (!session) return { success: false, error: 'Session expired' };
+  
+  const summary = getSummary(token, '2026');
+  if (!summary.success) return summary;
+  
+  // Pending deletions count
+  let pendingCount = 0;
+  try {
+    const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
+    const data = sheet.getRange(2, CONFIG.VOUCHER_COLUMNS.STATUS, sheet.getLastRow()-1, 1).getValues();
+    pendingCount = data.filter(r => r === 'Pending Deletion').length;
+  } catch (e) {}
+  
+  return {
+    success: true,
+    stats: summary.summary,
+    pendingDeletions: pendingCount,
+    recentVouchers: [] // Can implement if needed
+  };
}

function getSummary(token, year) {
-  // Placeholder for Part 2 logic
-  return { success: true, summary: { totalVouchersRaised: 0, totalPaidAmount: 0, totalDebt: 0 } };
-}
-
-function getDebtProfile(token) {
-  // Placeholder for Part 2 logic
-  return { success: true, totalDebt: 0, debtByCategory: [] };
-}
-```
-
-## 10. ActionItems.gs
-```javascript
-const ACTION_ITEM_RULES = {
-  PAID_NO_CN: "PAID_NO_CN",
-  UNPAID_NO_CN_30D: "UNPAID_NO_CN_30D",
-  RELEASED_UNPAID_15D: "RELEASED_UNPAID_15D"
-};
-
-const ACTION_ITEM_UNITS = { PAYABLE: "PAYABLE", CPO: "CPO" };
-
-function ensureActionItemSheets_() {
-  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
-  if (!ss.getSheetByName(CONFIG.SHEETS.ACTION_ITEMS_LOG)) {
-    ss.insertSheet(CONFIG.SHEETS.ACTION_ITEMS_LOG).appendRow([
-      "ITEM_ID", "RULE_KEY", "UNIT", "YEAR", "ROW_INDEX", "VOUCHER_NO",
-      "PAYEE", "AMOUNT", "CONTROL_NO", "STATUS", "FIRST_SEEN_AT", "LAST_SEEN_AT", "RESOLVED_AT"
-    ]);
-  }
-  if (!ss.getSheetByName(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS)) {
-    ss.insertSheet(CONFIG.SHEETS.ACTION_ITEMS_SETTINGS).appendRow(["UNIT", "RULE_KEY", "ENABLED"]);
-  }
-}
-
-function getActionItems(token, params) {
-  const session = getSession(token);
-  if (!session) return { success: false, error: "Session expired" };
-  
-  // Mock data for now until sync engine is fully integrated
-  return { success: true, items: [], count: 0 };
-}
-
-function getActionItemCount(token, params) {
-  const res = getActionItems(token, params);
-  return { success: true, count: res.count || 0 };
-}
-
-function getActionItemSettings(token) {
-  ensureActionItemSheets_();
-  return { success: true, settings: { unit: { PAYABLE: {}, CPO: {} } } };
-}
-
-function saveActionItemSettings(token, settings) {
-  return { success: true, message: "Settings saved" };
-}
-```
-
-## 8. Announcements.gs
-
-```javascript
-function ensureAnnouncementsSheet_() {
-  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
-  if (!ss.getSheetByName(CONFIG.SHEETS.ANNOUNCEMENTS)) {
-    ss.insertSheet(CONFIG.SHEETS.ANNOUNCEMENTS).appendRow([
-      "ID", "MESSAGE", "DISPLAY_LOCATIONS", "TARGET_USERS",
-      "EXPIRES_AT", "ALLOW_DISMISS", "CREATED_AT", "CREATED_BY"
-    ]);
-  }
-}
-
-function createAnnouncement(token, announcement) {
-  const session = getSession(token);
-  if (!session || session.role !== CONFIG.ROLES.ADMIN) return { success: false, error: "Unauthorized" };
-  
-  ensureAnnouncementsSheet_();
-  const sheet = getSheet(CONFIG.SHEETS.ANNOUNCEMENTS);
-  sheet.appendRow([
-    Utilities.getUuid(), announcement.message, JSON.stringify(announcement.locations),
-    JSON.stringify(announcement.targets), announcement.expiresAt, announcement.allowDismiss, new Date(), session.email
-  ]);
-  return { success: true, message: "Announcement created" };
-}
-
-function getActiveAnnouncements(token) {
-  const session = getSession(token);
-  if (!session) return { success: false, error: "Session expired" };
-  ensureAnnouncementsSheet_();
-  // Return empty for now, logic in Part 2
-  return { success: true, announcements: [] };
-}
-
-function dismissAnnouncement(token, id) {
-  return { success: true };
-}
-```
+  const session = getSession(token);
+  if (!session) return { success: false, error: 'Session expired' };
+  
+  const sheetName = year === '2026' ? CONFIG.SHEETS.VOUCHERS_2026 : CONFIG.SHEETS.VOUCHERS_2025; // Simplified
+  const sheet = getSheet(sheetName);
+  const data = sheet.getDataRange().getValues();
+  
+  let totalRaised = 0, paidAmt = 0, unpaidAmt = 0, debt = 0, cancelledAmt = 0;
+  
+  for (let i = 1; i < data.length; i++) {
+    const row = data[i];
+    if (!row[3]) continue; // No account/mail
+    
+    const amt = parseAmount(row[6]); // Gross
+    const status = String(row[0]).toLowerCase();
+    
+    totalRaised++;
+    if (status === 'paid') paidAmt += amt;
+    else if (status === 'cancelled') cancelledAmt += amt;
+    else { unpaidAmt += amt; debt += amt; }
+  }
+  
+  return {
+    success: true,
+    summary: {
+      totalVouchersRaised: totalRaised,
+      totalPaidAmount: paidAmt,
+      totalUnpaidAmount: unpaidAmt,
+      totalDebt: debt,
+      totalCancelledAmount: cancelledAmt
+    }
+  };
+}
+
+function getDebtProfile(token) {
+  const session = getSession(token);
+  if (!session) return { success: false, error: 'Session expired' };
+  
+  const sheet = getSheet(CONFIG.SHEETS.VOUCHERS_2026);
+  const data = sheet.getDataRange().getValues();
+  const debtByCat = {};
+  
+  for (let i = 1; i < data.length; i++) {
+    const row = data[i];
+    const status = String(row[0]).trim();
+    if (status !== 'Unpaid' && status !== 'Pending Deletion') continue;
+    
+    const cat = String(row[11]).trim() || 'Uncategorized';
+    const amt = parseAmount(row[6]);
+    
+    if (!debtByCat[cat]) debtByCat[cat] = 0;
+    debtByCat[cat] += amt;
+  }
+  
+  const categories = Object.keys(debtByCat).map(c => ({ category: c, amount: debtByCat[c] }));
+  

```
