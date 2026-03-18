function runFullSystemDiagnostics() {

  const startTotal = Date.now();

  const report = {
    timestamp: new Date().toISOString(),
    status: "OK",
    performance: {},
    errors: [],
    warnings: [],
    tests: {}
  };

  const mark = (name, start) =>
    report.performance[name] = (Date.now() - start) + " ms";

  const fail = msg => {
    report.status = "FAILED";
    report.errors.push(msg);
  };

  const warn = msg => report.warnings.push(msg);

  const pass = (name, details = "") =>
    report.tests[name] = { ok: true, details };

  const run = (name, fn) => {
    const t0 = Date.now();
    try {
      const result = fn();
      mark(name, t0);
      report.tests[name] = { ok: true, result };
      return result;
    } catch (e) {
      mark(name, t0);
      fail(`${name}: ${e.message}`);
      report.tests[name] = { ok: false };
      return null;
    }
  };

  try {

    // =============================
    // CONFIG DEEP CHECK
    // =============================
    run("CONFIG validation", () => {
      if (!CONFIG.SPREADSHEET_ID) throw new Error("Missing SPREADSHEET_ID");
      if (!CONFIG.SHEETS.USERS) throw new Error("Missing USERS sheet config");
      return "CONFIG OK";
    });

    // =============================
    // SPREADSHEET ACCESS
    // =============================
    run("Spreadsheet access", () =>
      SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getName()
    );

    // =============================
    // GET USERS (REAL DATA TEST)
    // =============================
    const users = run("Fetch USERS sheet", () => {
      const sheet = getSheet(CONFIG.SHEETS.USERS);
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) throw new Error("No users found");
      return data.length - 1;
    });

    // =============================
    // LOGIN SIMULATION
    // =============================
    let sessionToken = null;

    run("Login simulation", () => {

      const sheet = getSheet(CONFIG.SHEETS.USERS);
      const row = sheet.getDataRange().getValues()[1];

      const identifier =
        row[CONFIG.USER_COLUMNS.EMAIL - 1] ||
        row[CONFIG.USER_COLUMNS.USERNAME - 1];

      const password = row[CONFIG.USER_COLUMNS.PASSWORD - 1];

      const res = handleLogin(identifier, password);

      if (!res.success) throw new Error("Login failed");

      sessionToken = res.data.token;

      return "Login successful";
    });

    // =============================
    // SESSION VALIDATION
    // =============================
    run("Session validation", () => {
      const session = getSession(sessionToken);
      if (!session) throw new Error("Session not retrievable");
      return session.email;
    });

    // =============================
    // VOUCHER FETCH
    // =============================
    run("Voucher fetch", () => {
      const res = getVouchers(sessionToken, { page: 1, limit: 5 });
      if (!res.success) throw new Error("Voucher fetch failed");
      return res.data.length + " vouchers";
    });

    // =============================
    // PAGINATION TEST
    // =============================
    run("Pagination logic", () => {

      const page1 = getVouchers(sessionToken, { page: 1, limit: 2 });
      const page2 = getVouchers(sessionToken, { page: 2, limit: 2 });

      if (JSON.stringify(page1.data) === JSON.stringify(page2.data)) {
        throw new Error("Pagination not working");
      }

      return "Pagination OK";
    });

    // =============================
    // ROLE-RESTRICTED ENDPOINT
    // =============================
    run("Role security test", () => {

      const session = getSession(sessionToken);

      if (!hasPermission(session.role, [CONFIG.ROLES.ADMIN])) {
        return "User is not admin (correct behaviour)";
      }

      return "Admin access confirmed";
    });

    // =============================
    // CACHE SPEED TEST
    // =============================
    run("Cache performance", () => {

      const cache = CacheService.getScriptCache();

      const t0 = Date.now();
      cache.put("speed", "x", 5);
      cache.get("speed");

      return (Date.now() - t0) + " ms";
    });

  } catch (e) {
    fail("Unexpected failure: " + e.message);
  }

  report.performance.total = (Date.now() - startTotal) + " ms";

  Logger.log(JSON.stringify(report, null, 2));

  return report;
}

function checkSpreadsheetSize() {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    
    let totalCells = 0;
    const report = [];
    
    sheets.forEach(sheet => {
        const name = sheet.getName();
        const maxRows = sheet.getMaxRows();
        const maxCols = sheet.getMaxColumns();
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        const cells = maxRows * maxCols;
        const usedCells = lastRow * lastCol;
        
        totalCells += cells;
        
        report.push({
            name: name,
            maxRows: maxRows,
            maxCols: maxCols,
            lastRowWithData: lastRow,
            lastColWithData: lastCol,
            totalCells: cells,
            usedCells: usedCells,
            wastedCells: cells - usedCells
        });
    });
    
    // Sort by total cells descending
    report.sort((a, b) => b.totalCells - a.totalCells);
    
    Logger.log('=== SPREADSHEET SIZE REPORT ===');
    Logger.log('Total cells in spreadsheet: ' + totalCells.toLocaleString());
    Logger.log('Cell limit: 10,000,000');
    Logger.log('Remaining: ' + (10000000 - totalCells).toLocaleString());
    Logger.log('');
    Logger.log('=== SHEETS BREAKDOWN ===');
    
    report.forEach(s => {
        Logger.log(`\n${s.name}:`);
        Logger.log(`  Max Rows: ${s.maxRows}, Max Cols: ${s.maxCols}`);
        Logger.log(`  Last Row with Data: ${s.lastRowWithData}, Last Col with Data: ${s.lastColWithData}`);
        Logger.log(`  Total Cells: ${s.totalCells.toLocaleString()}`);
        Logger.log(`  Used Cells: ${s.usedCells.toLocaleString()}`);
        Logger.log(`  Wasted Cells: ${s.wastedCells.toLocaleString()}`);
    });
    
    return report;
}

function listAllSheets() {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    
    Logger.log('=== ALL SHEETS IN SPREADSHEET ===');
    Logger.log('Total sheets: ' + sheets.length);
    Logger.log('');
    
    // Known/expected sheets
    const expectedSheets = [
        CONFIG.SHEETS.VOUCHERS_2026,
        CONFIG.SHEETS.VOUCHERS_2025,
        CONFIG.SHEETS.VOUCHERS_2024,
        CONFIG.SHEETS.VOUCHERS_2023,
        CONFIG.SHEETS.VOUCHERS_BEFORE_2023,
        CONFIG.SHEETS.USERS,
        CONFIG.SHEETS.SYSTEM_CONFIG,
        'AUDIT_TRAIL',
        'NOTIFICATIONS',
        'ACTION_ITEMS_LOG',
        'ACTION_ITEMS_SETTINGS',
        'Announcements',
        CONFIG.SHEETS.ANNOUNCEMENTS
    ].filter(Boolean); // Remove undefined
    
    const expectedSet = new Set(expectedSheets.map(s => s ? s.toUpperCase() : ''));
    
    sheets.forEach((sheet, index) => {
        const name = sheet.getName();
        const isExpected = expectedSet.has(name.toUpperCase());
        const rows = sheet.getLastRow();
        const cols = sheet.getLastColumn();
        const maxRows = sheet.getMaxRows();
        const maxCols = sheet.getMaxColumns();
        
        const status = isExpected ? '✅ EXPECTED' : '❓ UNEXPECTED';
        
        Logger.log(`${index + 1}. "${name}" - ${status}`);
        Logger.log(`   Data: ${rows} rows x ${cols} cols`);
        Logger.log(`   Max: ${maxRows} rows x ${maxCols} cols`);
        Logger.log(`   Cells: ${maxRows * maxCols}`);
        Logger.log('');
    });
    
    // List unexpected sheets
    const unexpectedSheets = sheets.filter(s => !expectedSet.has(s.getName().toUpperCase()));
    
    if (unexpectedSheets.length > 0) {
        Logger.log('=== UNEXPECTED SHEETS ===');
        unexpectedSheets.forEach(s => {
            Logger.log('- ' + s.getName());
        });
    }
    
    return {
        total: sheets.length,
        expected: sheets.length - unexpectedSheets.length,
        unexpected: unexpectedSheets.length,
        unexpectedNames: unexpectedSheets.map(s => s.getName())
    };
}

function deleteGenericSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let deletedCount = 0;

  console.log("Starting cleanup of " + sheets.length + " sheets...");

  // Loop backwards to keep indices stable
  for (let i = sheets.length - 1; i >= 0; i--) {
    const sheet = sheets[i];
    const name = sheet.getName();
    
    // Target names starting with "SHEET" (case-insensitive)
    if (/^SHEET/i.test(name)) {
      try {
        ss.deleteSheet(sheet);
        deletedCount++;
        
        // Every 5 sheets, log progress and flush
        if (deletedCount % 5 === 0) {
          SpreadsheetApp.flush();
          console.log(`Deleted ${deletedCount} sheets so far...`);
        }
      } catch (e) {
        // This usually triggers if it's the last sheet in the file
        console.warn(`Skipped ${name}: ${e.message}`);
      }
    }
  }

  console.log(`FINISH: Successfully deleted ${deletedCount} sheets.`);
}
