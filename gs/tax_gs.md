/**
 * PAYABLE VOUCHER 2026 - Tax Backend Module
 * Consolidated tax-related backend functions
 */
const TAX_MONTHS_ = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
];

function normalizeTaxType_(taxType) {
    const value = String(taxType || '').trim().toUpperCase();
    if (!value) return '';
    if (value === 'VAT') return 'VAT';
    if (value === 'WHT' || value === 'WITHHOLDING TAX' || value === 'WITHHOLDING') return 'WHT';
    if (value === 'STAMP DUTY' || value === 'STAMP') return 'Stamp Duty';
    if (value === 'ALL' || value === 'ALL TYPES' || value === 'ALL TAX TYPES' || value === 'ALL TAX TYPE') {
        return 'All tax type';
    }
    return String(taxType || '').trim();
}

function normalizeMonthName_(monthValue) {
    const value = String(monthValue || '').trim().toLowerCase();
    if (!value) return '';
    for (let i = 0; i < TAX_MONTHS_.length; i++) {
        if (TAX_MONTHS_[i].toLowerCase() === value) return TAX_MONTHS_[i];
    }
    return '';
}

function getTaxPaymentBreakdown_(year) {
    const out = {
        paidVAT: 0,
        paidWHT: 0,
        paidStampDuty: 0,
        paidAllTypes: 0,
        totalPaid: 0,
        byMonth: {}
    };

    let paymentSheet;
    try {
        paymentSheet = getSheet(CONFIG.SHEETS.TAX_PAYMENTS);
    } catch (_) {
        return out;
    }

    const lastRow = paymentSheet.getLastRow();
    if (lastRow <= 1) return out;

    const data = paymentSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const targetYear = String(year || '2026').trim();

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const paymentYear = String(row[4] || '2026').trim() || '2026';
        if (paymentYear !== targetYear) continue;

        const amount = parseAmount(row[5]);
        if (!(amount > 0)) continue;

        const taxType = normalizeTaxType_(row[2]);
        const periodMonth = normalizeMonthName_(row[3]);

        if (taxType === 'VAT') out.paidVAT += amount;
        else if (taxType === 'WHT') out.paidWHT += amount;
        else if (taxType === 'Stamp Duty') out.paidStampDuty += amount;
        else if (taxType === 'All tax type') out.paidAllTypes += amount;
        else out.paidAllTypes += amount;

        out.totalPaid += amount;

        if (periodMonth) {
            if (!out.byMonth[periodMonth]) {
                out.byMonth[periodMonth] = {
                    paidVAT: 0,
                    paidWHT: 0,
                    paidStampDuty: 0,
                    paidAllTypes: 0,
                    totalPaid: 0
                };
            }
            const bucket = out.byMonth[periodMonth];
            if (taxType === 'VAT') bucket.paidVAT += amount;
            else if (taxType === 'WHT') bucket.paidWHT += amount;
            else if (taxType === 'Stamp Duty') bucket.paidStampDuty += amount;
            else bucket.paidAllTypes += amount;
            bucket.totalPaid += amount;
        }
    }

    return out;
}

/**
 * Get comprehensive tax summary for a year
 * @param {string} token - Session token
 * @param {string} year - Target year (default 2026)
 * @returns {Object} Tax summary data
 */
function getTaxSummary(token, year) {
    try {
        const session = getSession(token);
        if (!session) return { success: false, error: 'Session expired' };

        if (!hasPermission(session.role, [CONFIG.ROLES.TAX, CONFIG.ROLES.ADMIN])) {
            return { success: false, error: 'Unauthorized' };
        }

        year = year || '2026';
        const sheetName = getSheetNameByYear(year);
        const sheet = getSheet(sheetName);
        const lastRow = sheet.getLastRow();

        if (lastRow <= 1) {
            return {
                success: true,
                year: year,
                summary: getDefaultTaxSummary()
            };
        }

        const header = getHeaderMap_(sheet);
        const cols = CONFIG.VOUCHER_COLUMNS;
        const statusCol = (header['STATUS'] || cols.STATUS) - 1;
        const vatCol = (header['VAT'] || cols.VAT) - 1;
        const whtCol = (header['WHT'] || header['WITHHOLDING TAX'] || cols.WHT) - 1;
        const stampCol = (header['STAMP DUTY'] || cols.STAMP_DUTY) - 1;
        const accountTypeCol = (header['ACCOUNT TYPE'] || cols.ACCOUNT_TYPE) - 1;
        const subAccountCol = (header['SUB ACCOUNT'] || header['SUB ACCOUNT TYPE'] || cols.SUB_ACCOUNT_TYPE) - 1;
        const numCols = Math.max(statusCol, vatCol, whtCol, stampCol, accountTypeCol, subAccountCol) + 1;
        const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

        let totalVAT = 0, totalWHT = 0, totalStampDuty = 0;
        const accountTypeStats = {};

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const status = String(row[statusCol] || '').trim().toLowerCase();
            const vat = parseAmount(row[vatCol]);
            const wht = parseAmount(row[whtCol]);
            const stampDuty = parseAmount(row[stampCol]);
            const baseAccountType = String(row[accountTypeCol] || '').trim() || 'Unspecified';
            const subAccountType = String(row[subAccountCol] || '').trim();
            const accountKey = subAccountType ? `${baseAccountType}::${subAccountType}` : `${baseAccountType}::`;

            // Skip cancelled vouchers
            if (status === 'cancelled') continue;

            totalVAT += vat;
            totalWHT += wht;
            totalStampDuty += stampDuty;

            if (!accountTypeStats[accountKey]) {
                accountTypeStats[accountKey] = {
                    count: 0,
                    totalVAT: 0,
                    totalWHT: 0,
                    totalStampDuty: 0
                };
            }

            const at = accountTypeStats[accountKey];
            at.count++;
            at.totalVAT += vat;
            at.totalWHT += wht;
            at.totalStampDuty += stampDuty;
        }

        // Paid tax is sourced strictly from TAX_PAYMENTS records.
        const paymentBreakdown = getTaxPaymentBreakdown_(year);
        const paidVAT = paymentBreakdown.paidVAT;
        const paidWHT = paymentBreakdown.paidWHT;
        const paidStampDuty = paymentBreakdown.paidStampDuty;
        const paidAllTypes = paymentBreakdown.paidAllTypes;

        const totalTaxLiability = totalVAT + totalWHT + totalStampDuty;
        const totalPaid = paymentBreakdown.totalPaid;
        const totalOutstanding = Math.max(totalTaxLiability - totalPaid, 0);
        const accountTypeBreakdown = Object.keys(accountTypeStats)
            .map(accountKey => {
                const x = accountTypeStats[accountKey];
                const parts = accountKey.split('::');
                const baseType = String(parts[0] || '').trim() || 'Unspecified';
                const subType = String(parts[1] || '').trim();
                const formattedType = subType ? `${baseType} (${subType})` : baseType;
                const totalTax = x.totalVAT + x.totalWHT + x.totalStampDuty;
                return {
                    accountType: formattedType,
                    baseAccountType: baseType,
                    subAccountType: subType,
                    count: x.count,
                    totalVAT: x.totalVAT,
                    totalWHT: x.totalWHT,
                    totalStampDuty: x.totalStampDuty,
                    totalTax: totalTax,
                    paidTax: 0,
                    outstanding: totalTax,
                    complianceRate: 0
                };
            })
            .sort((a, b) => b.totalTax - a.totalTax);

        return {
            success: true,
            year: year,
            summary: {
                totalVAT,
                totalWHT,
                totalStampDuty,
                totalTaxLiability,
                paidVAT,
                paidWHT,
                paidStampDuty,
                totalPaid,
                unallocatedPaid: paidAllTypes,
                outstandingVAT: Math.max(totalVAT - paidVAT, 0),
                outstandingWHT: Math.max(totalWHT - paidWHT, 0),
                outstandingStampDuty: Math.max(totalStampDuty - paidStampDuty, 0),
                totalOutstanding,
                complianceRate: totalTaxLiability > 0 ? Math.min((totalPaid / totalTaxLiability) * 100, 100) : 0,
                paidFromRemittanceOnly: true,
                accountTypeBreakdown
            }
        };

    } catch (error) {
        return { success: false, error: 'Failed to get tax summary: ' + error.message };
    }
}

/**
 * Get tax breakdown by category
 */
function getTaxByCategory(token, year) {
    try {
        const session = getSession(token);
        if (!session) return { success: false, error: 'Session expired' };

        if (!hasPermission(session.role, [CONFIG.ROLES.TAX, CONFIG.ROLES.ADMIN])) {
            return { success: false, error: 'Unauthorized' };
        }

        year = year || '2026';
        const sheetName = getSheetNameByYear(year);
        const sheet = getSheet(sheetName);
        const lastRow = sheet.getLastRow();

        if (lastRow <= 1) {
            return { success: true, year: year, categories: [] };
        }

        const data = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
        const cols = CONFIG.VOUCHER_COLUMNS;
        const categoryStats = {};
        const paymentBreakdown = getTaxPaymentBreakdown_(year);

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const status = String(row[cols.STATUS - 1] || '').trim().toLowerCase();
            if (status === 'cancelled') continue;

            const category = String(row[cols.CATEGORIES - 1] || '').trim() || 'Uncategorized';
            const vat = parseAmount(row[cols.VAT - 1]);
            const wht = parseAmount(row[cols.WHT - 1]);
            const stampDuty = parseAmount(row[cols.STAMP_DUTY - 1]);

            if (!categoryStats[category]) {
                categoryStats[category] = {
                    totalVAT: 0, totalWHT: 0, totalStampDuty: 0,
                    count: 0
                };
            }

            const cat = categoryStats[category];
            cat.totalVAT += vat;
            cat.totalWHT += wht;
            cat.totalStampDuty += stampDuty;
            cat.count++;
        }

        const categories = Object.keys(categoryStats).map(catName => {
            const c = categoryStats[catName];
            const totalTax = c.totalVAT + c.totalWHT + c.totalStampDuty;

            return {
                category: catName,
                totalVAT: c.totalVAT,
                totalWHT: c.totalWHT,
                totalStampDuty: c.totalStampDuty,
                totalTax: totalTax,
                paidTax: 0,
                outstanding: totalTax,
                count: c.count,
                complianceRate: 0,
                paymentAllocation: 'UNMAPPED'
            };
        }).sort((a, b) => b.totalTax - a.totalTax);

        return {
            success: true,
            year: year,
            categories: categories,
            totalRemitted: paymentBreakdown.totalPaid,
            paymentAllocationMode: 'UNMAPPED'
        };

    } catch (error) {
        return { success: false, error: 'Failed to get tax by category: ' + error.message };
    }
}

/**
 * Get tax breakdown by month
 */
function getTaxByMonth(token, year) {
    try {
        const session = getSession(token);
        if (!session) return { success: false, error: 'Session expired' };

        if (!hasPermission(session.role, [CONFIG.ROLES.TAX, CONFIG.ROLES.ADMIN])) {
            return { success: false, error: 'Unauthorized' };
        }

        year = year || '2026';
        const sheetName = getSheetNameByYear(year);
        const sheet = getSheet(sheetName);
        const lastRow = sheet.getLastRow();

        const monthlyStats = {};
        TAX_MONTHS_.forEach(m => {
            monthlyStats[m] = {
                totalVAT: 0, totalWHT: 0, totalStampDuty: 0,
                paidVAT: 0, paidWHT: 0, paidStampDuty: 0, paidAllTypes: 0,
                count: 0
            };
        });

        if (lastRow <= 1) {
            return {
                success: true,
                year: year,
                months: TAX_MONTHS_.map(month => ({
                    month,
                    totalVAT: 0, totalWHT: 0, totalStampDuty: 0,
                    totalTax: 0, paidTax: 0, outstanding: 0, count: 0
                }))
            };
        }

        const data = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
        const cols = CONFIG.VOUCHER_COLUMNS;

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const status = String(row[cols.STATUS - 1] || '').trim().toLowerCase();
            if (status === 'cancelled') continue;

            const pmtMonth = String(row[cols.PMT_MONTH - 1] || '').trim();
            if (!TAX_MONTHS_.includes(pmtMonth)) continue;

            const vat = parseAmount(row[cols.VAT - 1]);
            const wht = parseAmount(row[cols.WHT - 1]);
            const stampDuty = parseAmount(row[cols.STAMP_DUTY - 1]);

            const m = monthlyStats[pmtMonth];
            m.totalVAT += vat;
            m.totalWHT += wht;
            m.totalStampDuty += stampDuty;
            m.count++;
        }

        // Paid values come from remittance records by period month.
        const paymentBreakdown = getTaxPaymentBreakdown_(year);
        Object.keys(paymentBreakdown.byMonth).forEach(month => {
            if (!monthlyStats[month]) return;
            const paid = paymentBreakdown.byMonth[month];
            monthlyStats[month].paidVAT += paid.paidVAT;
            monthlyStats[month].paidWHT += paid.paidWHT;
            monthlyStats[month].paidStampDuty += paid.paidStampDuty;
            monthlyStats[month].paidAllTypes += paid.paidAllTypes;
        });

        const months = TAX_MONTHS_.map(month => {
            const m = monthlyStats[month];
            const totalTax = m.totalVAT + m.totalWHT + m.totalStampDuty;
            const paidTax = m.paidVAT + m.paidWHT + m.paidStampDuty + m.paidAllTypes;

            return {
                month,
                totalVAT: m.totalVAT,
                totalWHT: m.totalWHT,
                totalStampDuty: m.totalStampDuty,
                totalTax: totalTax,
                paidTax: paidTax,
                paidAllTypes: m.paidAllTypes,
                outstanding: Math.max(totalTax - paidTax, 0),
                count: m.count
            };
        });

        return { success: true, year: year, months: months, paidFromRemittanceOnly: true };

    } catch (error) {
        return { success: false, error: 'Failed to get tax by month: ' + error.message };
    }
}

/**
 * Get tax breakdown by payee with filters
 */
function getTaxByPayee(token, year, filters) {
    try {
        const session = getSession(token);
        if (!session) return { success: false, error: 'Session expired' };

        if (!hasPermission(session.role, [CONFIG.ROLES.TAX, CONFIG.ROLES.ADMIN])) {
            return { success: false, error: 'Unauthorized' };
        }

        year = year || '2026';
        filters = filters || {};
        const sheetName = getSheetNameByYear(year);
        const sheet = getSheet(sheetName);
        const lastRow = sheet.getLastRow();

        if (lastRow <= 1) {
            return { success: true, year: year, payees: [] };
        }

        const data = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
        const cols = CONFIG.VOUCHER_COLUMNS;
        const payeeStats = {};
        const paymentBreakdown = getTaxPaymentBreakdown_(year);

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const status = String(row[cols.STATUS - 1] || '').trim().toLowerCase();
            if (status === 'cancelled') continue;

            const payee = String(row[cols.PAYEE - 1] || '').trim();
            if (!payee) continue;

            // Apply filters
            if (filters.category && filters.category !== 'All') {
                const category = String(row[cols.CATEGORIES - 1] || '').trim();
                if (category !== filters.category) continue;
            }

            if (filters.month && filters.month !== 'All') {
                const month = String(row[cols.PMT_MONTH - 1] || '').trim();
                if (month !== filters.month) continue;
            }

            const vat = parseAmount(row[cols.VAT - 1]);
            const wht = parseAmount(row[cols.WHT - 1]);
            const stampDuty = parseAmount(row[cols.STAMP_DUTY - 1]);

            if (!payeeStats[payee]) {
                payeeStats[payee] = {
                    totalVAT: 0, totalWHT: 0, totalStampDuty: 0,
                    count: 0
                };
            }

            const p = payeeStats[payee];
            p.totalVAT += vat;
            p.totalWHT += wht;
            p.totalStampDuty += stampDuty;
            p.count++;
        }

        const payeesList = Object.keys(payeeStats).map(payeeName => {
            const p = payeeStats[payeeName];
            const totalTax = p.totalVAT + p.totalWHT + p.totalStampDuty;

            return {
                payee: payeeName,
                totalVAT: p.totalVAT,
                totalWHT: p.totalWHT,
                totalStampDuty: p.totalStampDuty,
                totalTax: totalTax,
                paidTax: 0,
                outstanding: totalTax,
                count: p.count,
                paymentAllocation: 'UNMAPPED'
            };
        }).sort((a, b) => b.totalTax - a.totalTax);

        return {
            success: true,
            year: year,
            payees: payeesList,
            filters: filters,
            totalRemitted: paymentBreakdown.totalPaid,
            paymentAllocationMode: 'UNMAPPED'
        };

    } catch (error) {
        return { success: false, error: 'Failed to get tax by payee: ' + error.message };
    }
}

/**
 * Record a tax payment
 */
function recordTaxPayment(token, payment) {
    try {
        const session = getSession(token);
        if (!session) return { success: false, error: 'Session expired' };

        if (!hasPermission(session.role, [CONFIG.ROLES.TAX, CONFIG.ROLES.ADMIN])) {
            return { success: false, error: 'Unauthorized' };
        }

        if (!payment || typeof payment !== 'object') {
            return { success: false, error: 'Payment details are required' };
        }

        const normalizedTaxType = normalizeTaxType_(payment.taxType);
        if (!['VAT', 'WHT', 'Stamp Duty', 'All tax type'].includes(normalizedTaxType)) {
            return { success: false, error: 'Invalid tax type' };
        }

        const normalizedMonth = normalizeMonthName_(payment.period);
        if (!normalizedMonth) {
            return { success: false, error: 'Valid remittance period month is required' };
        }

        const amount = parseAmount(payment.amount);
        if (!(amount > 0)) {
            return { success: false, error: 'Amount must be greater than zero' };
        }

        const sheet = getSheet(CONFIG.SHEETS.TAX_PAYMENTS);
        const id = 'TXP-' + Date.now();
        const now = new Date();
        const year = String(payment.year || '2026').trim() || '2026';

        sheet.appendRow([
            id,
            payment.date || now,
            normalizedTaxType,
            normalizedMonth,
            year,
            amount,
            payment.paymentMethod || '',
            payment.referenceNumber || '',
            payment.bank || '',
            payment.notes || '',
            session.name,
            now
        ]);

        // Log activity
        logAudit(session, 'TAX_PAYMENT_RECORDED', 
            `Recorded ${normalizedTaxType} payment: NGN ${amount.toLocaleString()}`,
            CONFIG.SHEETS.TAX_PAYMENTS,
            sheet.getLastRow(),
            { paymentId: id, taxType: normalizedTaxType, amount: amount, period: normalizedMonth, year: year }
        );

        return {
            success: true,
            message: 'Tax payment recorded successfully',
            payment: {
                id: id,
                ...payment,
                taxType: normalizedTaxType,
                period: normalizedMonth,
                year: year,
                amount: amount
            }
        };

    } catch (error) {
        return { success: false, error: 'Failed to record tax payment: ' + error.message };
    }
}

/**
 * Get tax payments
 */
function getTaxPayments(token, year) {
    try {
        const session = getSession(token);
        if (!session) return { success: false, error: 'Session expired' };

        if (!hasPermission(session.role, [CONFIG.ROLES.TAX, CONFIG.ROLES.ADMIN])) {
            return { success: false, error: 'Unauthorized' };
        }

        let sheet;
        try {
            sheet = getSheet(CONFIG.SHEETS.TAX_PAYMENTS);
        } catch (_) {
            return { success: true, payments: [] };
        }
        const lastRow = sheet.getLastRow();

        if (lastRow <= 1) return { success: true, payments: [] };

        const targetYear = String(year || '').trim();
        const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
        const payments = data
            .filter(row => {
                if (!targetYear) return true;
                const rowYear = String(row[4] || '2026').trim() || '2026';
                return rowYear === targetYear;
            })
            .map(row => ({
                id: row[0],
                date: row[1],
                taxType: normalizeTaxType_(row[2]),
                period: normalizeMonthName_(row[3]) || row[3],
                year: String(row[4] || '2026').trim() || '2026',
                amount: parseAmount(row[5]),
                paymentMethod: row[6],
                referenceNumber: row[7],
                bank: row[8],
                notes: row[9],
                createdBy: row[10],
                createdAt: row[11]
            }))
            .sort((a, b) => new Date(b.date) - new Date(a.date));

        return { success: true, payments: payments };

    } catch (error) {
        return { success: false, error: 'Failed to get tax payments: ' + error.message };
    }
}

/**
 * Get tax compliance status
 */
function getTaxCompliance(token, year) {
    try {
        const summaryRes = getTaxSummary(token, year);
        if (!summaryRes.success) return summaryRes;

        const summary = summaryRes.summary;
        
        return {
            success: true,
            year: year || '2026',
            compliance: {
                VAT: calculateTypeCompliance(summary.totalVAT, summary.paidVAT),
                WHT: calculateTypeCompliance(summary.totalWHT, summary.paidWHT),
                StampDuty: calculateTypeCompliance(summary.totalStampDuty, summary.paidStampDuty)
            },
            overallCompliance: summary.complianceRate
        };
    } catch (error) {
        return { success: false, error: 'Failed to get tax compliance: ' + error.message };
    }
}

/**
 * Helper to calculate compliance for a specific tax type
 */
function calculateTypeCompliance(liability, paid) {
    return {
        liability,
        paid,
        outstanding: liability - paid,
        complianceRate: liability > 0 ? (paid / liability * 100) : 0
    };
}

/**
 * Get tax schedules for a year
 */
function getTaxSchedule(token, year) {
    try {
        const session = getSession(token);
        if (!session) return { success: false, error: 'Session expired' };

        if (!hasPermission(session.role, [CONFIG.ROLES.TAX, CONFIG.ROLES.ADMIN])) {
            return { success: false, error: 'Unauthorized' };
        }

        const sheet = getSheet(CONFIG.SHEETS.TAX_SCHEDULE);
        const lastRow = sheet.getLastRow();

        if (lastRow <= 1) return { success: true, schedules: [] };

        const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
        const schedules = data.filter(row => !year || String(row[2]).trim() === year)
            .map(row => ({
                id: row[0],
                month: row[1],
                year: row[2],
                dueDate: row[3],
                status: row[4],
                amount: parseAmount(row[5]),
                taxType: row[6],
                notes: row[7],
                createdBy: row[8],
                createdAt: row[9],
                lastUpdated: row[10]
            }));

        return { success: true, schedules: schedules };

    } catch (error) {
        return { success: false, error: 'Failed to get tax schedule: ' + error.message };
    }
}

/**
 * Create a new tax schedule
 */
function createTaxSchedule(token, schedule) {
    try {
        const session = getSession(token);
        if (!session) return { success: false, error: 'Session expired' };

        if (!hasPermission(session.role, [CONFIG.ROLES.TAX, CONFIG.ROLES.ADMIN])) {
            return { success: false, error: 'Unauthorized' };
        }

        const sheet = getSheet(CONFIG.SHEETS.TAX_SCHEDULE);
        const id = 'TXS-' + Date.now();
        const now = new Date();

        sheet.appendRow([
            id,
            schedule.month || '',
            schedule.year || '2026',
            schedule.dueDate || '',
            schedule.status || 'Pending',
            parseAmount(schedule.amount),
            schedule.taxType || '',
            schedule.notes || '',
            session.name,
            now,
            now
        ]);

        return { success: true, message: 'Schedule created successfully', id: id };

    } catch (error) {
        return { success: false, error: 'Failed to create tax schedule: ' + error.message };
    }
}

/**
 * Update an existing tax schedule
 */
function updateTaxSchedule(token, scheduleId, updates) {
    try {
        const session = getSession(token);
        if (!session) return { success: false, error: 'Session expired' };

        if (!hasPermission(session.role, [CONFIG.ROLES.TAX, CONFIG.ROLES.ADMIN])) {
            return { success: false, error: 'Unauthorized' };
        }

        const sheet = getSheet(CONFIG.SHEETS.TAX_SCHEDULE);
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return { success: false, error: 'No schedules found' };

        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => r[0]);
        const rowIndex = ids.indexOf(scheduleId);

        if (rowIndex === -1) return { success: false, error: 'Schedule not found' };

        const row = rowIndex + 2;
        const now = new Date();

        if (updates.status) sheet.getRange(row, 5).setValue(updates.status);
        if (updates.amount !== undefined) sheet.getRange(row, 6).setValue(parseAmount(updates.amount));
        if (updates.notes !== undefined) sheet.getRange(row, 8).setValue(updates.notes);
        sheet.getRange(row, 11).setValue(now);

        return { success: true, message: 'Schedule updated successfully' };

    } catch (error) {
        return { success: false, error: 'Failed to update tax schedule: ' + error.message };
    }
}

/**
 * Helper to get default tax summary object
 */
function getDefaultTaxSummary() {
    return {
        totalVAT: 0, totalWHT: 0, totalStampDuty: 0, totalTaxLiability: 0,
        paidVAT: 0, paidWHT: 0, paidStampDuty: 0, totalPaid: 0,
        unallocatedPaid: 0,
        outstandingVAT: 0, outstandingWHT: 0, outstandingStampDuty: 0, totalOutstanding: 0,
        complianceRate: 0,
        paidFromRemittanceOnly: true,
        accountTypeBreakdown: []
    };
}

/**
 * Get sheet name by year helper
 */
function getSheetNameByYear(year) {
    switch (year) {
        case '2026': return CONFIG.SHEETS.VOUCHERS_2026;
        case '2025': return CONFIG.SHEETS.VOUCHERS_2025;
        case '2024': return CONFIG.SHEETS.VOUCHERS_2024;
        case '2023': return CONFIG.SHEETS.VOUCHERS_2023;
        case '<2023': return CONFIG.SHEETS.VOUCHERS_BEFORE_2023;
        default: return CONFIG.SHEETS.VOUCHERS_2026;
    }
}
