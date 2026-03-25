/**
 * PAYABLE VOUCHER 2026 - Reports Module
 * Comprehensive reporting and analytics
 */

const Reports = {
    // State
    currentYear: '2026',
    summaryData: null,
    allYearsData: null,
    debtProfile: null,
    voucherCache: {},
    voucherStats: null,
    taxMonthlyData: null,
    taxSummary: null,
    categoryChart: null,
    monthlyChart: null,
    statusCountChart: null,
    statusAmountChart: null,
    debtTrendChart: null,
    topDebtorsChart: null,
    categoryParetoChart: null,
    debtConcentrationChart: null,
    agingChart: null,
    cashCommitChart: null,
    accrualsSnapshotChart: null,
    netPayableChart: null,
    voucherDistChart: null,
    taxSplitChart: null,
    revalidatedImpactChart: null,
    categoryShareShiftChart: null,
    cancelledImpactChart: null,
    debtRequestStatus: null,
    debtReportData: null,
    reportAgingChart: null,
    reportCategoryChart: null,
    reportDepartmentChart: null,
    systemConfig: null,
    currentAccountTypeRows: [],
    accountTypeFilters: { accountType: 'ALL', subAccountType: 'ALL' },
    expandedAccountTypeGroups: new Set(),

    /**
     * Initialize reports page
     */
    async init() {
        const isAuth = await Auth.requireAuth();
        if (!isAuth) return;

        const user = Auth.getUser();
        const isPayableStaff = user && (user.role === CONFIG.ROLES.PAYABLE_STAFF);

        if (isPayableStaff) {
            Utils.showToast('Access denied: Reports is not available for Payable Unit Staff.', 'error');
            window.location.replace('dashboard.html');
            return;
        }

        this.setupSidebar();
        this.setupEventListeners();
        await this.loadSystemConfig();
        await this.loadAllReports();
        await this.checkDebtProfileStatus();
    },

    /**
     * Setup sidebar navigation
     */
    setupSidebar() {
        const sidebar = document.getElementById('sidebar');
        if (sidebar) {
            sidebar.innerHTML = Components.getSidebar('reports');
        }
    },

    /**
     * Setup event listeners
     */
    setupEventListeners() {
        // Year selector
        const yearSelector = document.getElementById('yearSelector');
        if (yearSelector) {
            yearSelector.addEventListener('change', (e) => {
                this.currentYear = e.target.value;
                document.getElementById('currentYearLabel').textContent = this.currentYear;
                this.expandedAccountTypeGroups.clear();
                this.loadYearSummary();
            });
        }

        document.getElementById('accountTypeFilter')?.addEventListener('change', () => {
            this.accountTypeFilters.accountType = document.getElementById('accountTypeFilter')?.value || 'ALL';
            this.accountTypeFilters.subAccountType = 'ALL';
            this.populateSubAccountTypeFilter(this.currentAccountTypeRows);
            this.renderAccountTypeTable(this.currentAccountTypeRows);
        });

        document.getElementById('subAccountTypeFilter')?.addEventListener('change', () => {
            this.accountTypeFilters.subAccountType = document.getElementById('subAccountTypeFilter')?.value || 'ALL';
            this.renderAccountTypeTable(this.currentAccountTypeRows);
        });

        // Logout
        document.getElementById('logoutBtn')?.addEventListener('click', () => Auth.logout());

        // Mobile menu
        document.getElementById('menuToggle')?.addEventListener('click', () => {
            document.getElementById('sidebar')?.classList.toggle('active');
        });

        // Print button
        document.getElementById('printBtn')?.addEventListener('click', () => window.print());

        // Export button
        document.getElementById('exportBtn')?.addEventListener('click', () => this.exportToCSV());

        // Debt Profile Request
        document.getElementById('requestDebtProfileBtn')?.addEventListener('click', () => this.handleDebtProfileRequest());
        document.getElementById('reRequestDebtProfileBtn')?.addEventListener('click', () => this.handleDebtProfileRequest());

        // Debt Profile Approval
        document.getElementById('approveDebtProfileBtn')?.addEventListener('click', () => this.handleDebtProfileApproval('approve'));
        document.getElementById('rejectDebtProfileBtn')?.addEventListener('click', () => this.handleDebtProfileApproval('reject'));

        // View Report
        document.getElementById('viewApprovedReportBtn')?.addEventListener('click', () => this.loadFullDebtReport());

        // Debt Profile Events (using delegation for components that might be re-rendered)
        document.addEventListener('click', (e) => {
            const id = e.target.id || e.target.closest('button')?.id;
            
            if (id === 'editNarrativeBtn') {
                console.log('Edit Narrative button clicked');
                this.handleDebtProfileRequest();
            } else if (id === 'createNewReportBtn') {
                console.log('Create New Report button clicked');
                this.handleDebtProfileRequest(true); // pass true for fresh report
            } else if (id === 'downloadFullPDF') {
                this.downloadDebtReportGenerated('pdf');
            } else if (id === 'downloadFullExcel') {
                this.downloadDebtReportGenerated('excel');
            }
        });
        
        // Auto-Generate button (inside modal)
        document.getElementById('autoGenerateBtn')?.addEventListener('click', () => this.generateSmartNarrative());

        // Modal Close logic
        const modalOverlay = document.getElementById('analyticalFormModalOverlay');
        if (modalOverlay) {
            modalOverlay.querySelectorAll('.close-modal, .cancel-modal').forEach(btn => {
                btn.addEventListener('click', () => {
                    console.log('Closing modal');
                    modalOverlay.classList.remove('active');
                });
            });
        } else {
            console.error('analyticalFormModalOverlay not found in DOM');
        }

        // Modal Submit logic
        const analyticalForm = document.getElementById('analyticalReportForm');
        if (analyticalForm) {
            analyticalForm.onsubmit = async (e) => {
                e.preventDefault();
                console.log('Submitting analytical report form');
                await this.submitDebtProfileRequest();
            };
        }

        // Listen for SWR background updates
        document.addEventListener('apiDataUpdated', (e) => {
            const { action, data, params } = e.detail;

            // If the user hasn't switched the year dropdown, re-apply the fresh data
            if (action === 'getSummary' && params.year === this.currentYear) {
                this.summaryData = data;
                this.renderYearSummary(data);
                this.renderFinancialInsights(data.summary);
                this.renderCategoryTable(data.categoryBreakdown);
                this.renderAccountTypeTable(data.accountTypeBreakdown);
                this.renderMonthlyTable(data.monthlyBreakdown);
                this.drawCategoryChart(data.categoryBreakdown);
                this.drawCategoryParetoChart(data.categoryBreakdown);
                this.drawCategoryShareShiftChart(data.categoryBreakdown);
                this.drawMonthlyChart(data.monthlyBreakdown);
                this.drawStatusCharts(data.summary);
                this.drawCashCommitChart(data.summary);
                this.drawAccrualsSnapshot(data.summary);
            }
            if (action === 'getAllYearsSummary') {
                this.allYearsData = data;
                this.renderAllYearsSummary(data);
                this.drawDebtTrendChart(data);
                this.drawNetPayableChart(data);
            }
            if (action === 'getDebtProfile') {
                this.debtProfile = data;
                this.renderDebtProfile(data);
                this.drawDebtConcentration(data);
                this.drawTopDebtorsChart(data);
            }
            if (action === 'getTaxByMonth' && params.year === this.currentYear) {
                this.taxMonthlyData = data;
                this.drawTaxSplitChart(data);
            }
        });
    },

    /**
     * Load all reports data (initial load)
     */
    async loadAllReports() {
        this.showLoading(true);

        try {
            await Promise.all([
                this.loadYearSummary(),
                this.loadAllYearsSummary(),
                this.loadDebtProfile()
            ]);
        } catch (error) {
            console.error('Error loading reports:', error);
            Utils.showToast('Error loading reports', 'error');
        }

        this.showLoading(false);
    },

    async loadSystemConfig() {
        try {
            const result = await API.getSystemConfig();
            if (result.success) {
                this.systemConfig = result.config || result;
            }
        } catch (e) {
            console.warn('System config load failed:', e);
        }
    },

    /**
     * Load summary for selected year
     */
    async loadYearSummary() {
        const result = await API.getSummary(this.currentYear);

        if (result.success) {
            this.summaryData = result;
            this.renderYearSummary(result);
            this.renderFinancialInsights(result.summary);
            this.renderCategoryTable(result.categoryBreakdown);
            this.renderAccountTypeTable(result.accountTypeBreakdown);
            this.renderMonthlyTable(result.monthlyBreakdown);
            this.drawCategoryChart(result.categoryBreakdown);
            this.drawCategoryParetoChart(result.categoryBreakdown);
            this.drawCategoryShareShiftChart(result.categoryBreakdown);
            this.drawMonthlyChart(result.monthlyBreakdown);
            this.drawStatusCharts(result.summary);
            this.drawCashCommitChart(result.summary);
            this.drawAccrualsSnapshot(result.summary);
        } else {
            Utils.showToast(result.error || 'Failed to load summary', 'error');
        }

        await Promise.all([
            this.loadVoucherAnalytics(),
            this.loadTaxByMonth(),
            this.loadTaxSummary()
        ]);
    },

    parseDateFlexible(value) {
        if (!value) return null;
        if (value instanceof Date) return value;
        const d = new Date(value);
        if (!isNaN(d.getTime())) return d;

        // Handle "1/5/2026 20:20:22" style
        const m = String(value).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
        if (!m) return null;

        const a = parseInt(m[1], 10);
        const b = parseInt(m[2], 10);
        const yyyy = parseInt(m[3], 10);
        const hh = parseInt(m[4], 10);
        const mm = parseInt(m[5], 10);
        const ss = parseInt(m[6] || '0', 10);

        // Assume dd/mm unless forced
        let day = a, month = b;
        if (a <= 12 && b <= 12) { day = a; month = b; } // keep dd/mm
        else if (a <= 12 && b > 12) { month = a; day = b; }

        const out = new Date(yyyy, month - 1, day, hh, mm, ss);
        return isNaN(out.getTime()) ? null : out;
    },

    async loadAgingAnalysis2026() {
        const section = document.getElementById('agingSection');
        section?.classList.remove('hidden');

        // Pull UNPAID vouchers. If you have huge data, we can page through.
        // This assumes your getVouchers supports pagination and returns totalPages.
        const all = [];

        this.showLoading(true);
        try {
            // 1. Fetch first page to get total count
            const firstRes = await API.getVouchers('2026', { status: 'Unpaid' }, 1, 200);
            if (!firstRes.success) return;

            all.push(...(firstRes.vouchers || []));
            const totalPages = firstRes.totalPages || 1;

            // 2. Fetch remaining pages in parallel (if any)
            if (totalPages > 1) {
                const promises = [];
                for (let p = 2; p <= totalPages; p++) {
                    promises.push(API.getVouchers('2026', { status: 'Unpaid' }, p, 200));
                }
                const results = await Promise.all(promises);
                results.forEach(res => {
                    if (res.success && res.vouchers) all.push(...res.vouchers);
                });
            }

            const now = new Date();
            const buckets = [
                { label: '0–7 days', min: 0, max: 7, count: 0, amount: 0 },
                { label: '8–14 days', min: 8, max: 14, count: 0, amount: 0 },
                { label: '15–30 days', min: 15, max: 30, count: 0, amount: 0 },
                { label: '31–60 days', min: 31, max: 60, count: 0, amount: 0 },
                { label: '61–90 days', min: 61, max: 90, count: 0, amount: 0 },
                { label: '90+ days', min: 91, max: 99999, count: 0, amount: 0 },
            ];

            all.forEach(v => {
                const created = this.parseDateFlexible(v.createdAt || v.date);
                if (!created) return;

                const ageDays = Math.floor((now - created) / (1000 * 60 * 60 * 24));
                const amt = Number(v.grossAmount || 0);

                const b = buckets.find(x => ageDays >= x.min && ageDays <= x.max);
                if (!b) return;
                b.count += 1;
                b.amount += amt;
            });

            // Draw chart
            const ctx = document.getElementById('agingChart');
            if (ctx) {
                if (this.agingChart) this.agingChart.destroy();

                this.agingChart = new Chart(ctx, {
                    type: 'bar',
                    data: {
                        labels: buckets.map(b => b.label),
                        datasets: [{
                            label: 'Total Amount (₦)',
                            data: buckets.map(b => b.amount),
                            backgroundColor: 'rgba(255,193,7,0.65)'
                        }]
                    },
                    options: {
                        responsive: true,
                        plugins: { legend: { position: 'top' } },
                        scales: {
                            y: { ticks: { callback: v => '₦' + Number(v).toLocaleString() } }
                        }
                    }
                });
            }

            // Render table
            const t = document.getElementById('agingTable');
            if (t) {
                t.innerHTML = `
                <div class="table-container">
                <table>
                    <thead>
                    <tr>
                        <th>Bucket</th>
                        <th class="text-center">Count</th>
                        <th class="text-right">Total Amount</th>
                    </tr>
                    </thead>
                    <tbody>
                    ${buckets.map(b => `
                        <tr>
                        <td><strong>${b.label}</strong></td>
                        <td class="text-center">${Utils.formatNumber(b.count)}</td>
                        <td class="text-right">${Utils.formatCurrency(b.amount)}</td>
                        </tr>
                    `).join('')}
                    </tbody>
                </table>
                </div>
            `;
            }

        } catch (e) {
            console.error('Aging analysis error:', e);
            Utils.showToast('Could not compute aging analysis', 'warning');
        } finally {
            this.showLoading(false);
        }
    },

    getVoucherAmount(voucher) {
        const val = voucher?.grossAmount ?? voucher?.amount ?? voucher?.totalAmount ?? voucher?.contractSum ?? 0;
        const num = Number(val);
        return Number.isFinite(num) ? num : 0;
    },

    isYesValue(value) {
        const raw = String(value || '').trim().toLowerCase();
        return raw === 'yes' || raw === 'y' || raw === 'true';
    },

    isRevalidatedVoucher(voucher) {
        const hasOldNumber = !!String(voucher?.oldVoucherNumber || '').trim();
        return hasOldNumber || this.isYesValue(voucher?.oldVoucherAvailable);
    },

    getRevalidationCounts(vouchers) {
        let revalidatedVouchers = 0;
        let revalidatedWithoutOldNumber = 0;

        (vouchers || []).forEach((v) => {
            const hasOldNumber = !!String(v?.oldVoucherNumber || '').trim();
            const markedAvailable = this.isYesValue(v?.oldVoucherAvailable);
            if (hasOldNumber || markedAvailable) {
                revalidatedVouchers++;
                if (!hasOldNumber && markedAvailable) revalidatedWithoutOldNumber++;
            }
        });

        return { revalidatedVouchers, revalidatedWithoutOldNumber };
    },

    getTotalVoucherAmount(summary) {
        if (!summary) return 0;
        const paid = Number(summary.totalPaidAmount || 0);
        const unpaid = Number(summary.totalUnpaidAmount || 0);
        const cancelled = Number(summary.totalCancelledAmount || 0);
        return paid + unpaid + cancelled;
    },

    drawAgingAnalysisFromVouchers(vouchers) {
        const section = document.getElementById('agingSection');
        section?.classList.remove('hidden');

        const unpaid = (vouchers || []).filter(v => String(v.status || '').toLowerCase() === 'unpaid');
        if (!unpaid.length) {
            document.getElementById('agingTable')?.replaceChildren();
            if (this.agingChart) this.agingChart.destroy();
            return;
        }

        const now = new Date();
        const buckets = [
            { label: '0-7 days', min: 0, max: 7, count: 0, amount: 0 },
            { label: '8-14 days', min: 8, max: 14, count: 0, amount: 0 },
            { label: '15-30 days', min: 15, max: 30, count: 0, amount: 0 },
            { label: '31-60 days', min: 31, max: 60, count: 0, amount: 0 },
            { label: '61-90 days', min: 61, max: 90, count: 0, amount: 0 },
            { label: '90+ days', min: 91, max: 99999, count: 0, amount: 0 },
        ];

        unpaid.forEach(v => {
            const created = this.parseDateFlexible(v.createdAt || v.date);
            if (!created) return;

            const ageDays = Math.floor((now - created) / (1000 * 60 * 60 * 24));
            const amt = this.getVoucherAmount(v);

            const b = buckets.find(x => ageDays >= x.min && ageDays <= x.max);
            if (!b) return;
            b.count += 1;
            b.amount += amt;
        });

        const ctx = document.getElementById('agingChart');
        if (ctx) {
            if (this.agingChart) this.agingChart.destroy();

            this.agingChart = new Chart(ctx, {
                data: {
                    labels: buckets.map(b => b.label),
                    datasets: [
                        {
                            type: 'bar',
                            label: 'Outstanding Amount',
                            data: buckets.map(b => b.amount),
                            backgroundColor: 'rgba(255,193,7,0.65)',
                            yAxisID: 'y'
                        },
                        {
                            type: 'line',
                            label: 'Voucher Count',
                            data: buckets.map(b => b.count),
                            borderColor: 'rgba(23,162,184,1)',
                            backgroundColor: 'rgba(23,162,184,0.2)',
                            tension: 0.3,
                            yAxisID: 'y1'
                        }
                    ]
                },
                options: {
                    responsive: true,
                    plugins: { legend: { position: 'top' } },
                    scales: {
                        y: { ticks: { callback: v => 'â‚¦' + Number(v).toLocaleString() } },
                        y1: {
                            position: 'right',
                            grid: { drawOnChartArea: false },
                            ticks: { callback: v => Number(v).toLocaleString() }
                        }
                    }
                }
            });
        }

        const t = document.getElementById('agingTable');
        if (t) {
            t.innerHTML = `
            <div class="table-container">
            <table>
                <thead>
                <tr>
                    <th>Bucket</th>
                    <th class="text-center">Count</th>
                    <th class="text-right">Total Amount</th>
                </tr>
                </thead>
                <tbody>
                ${buckets.map(b => `
                    <tr>
                    <td><strong>${b.label}</strong></td>
                    <td class="text-center">${Utils.formatNumber(b.count)}</td>
                    <td class="text-right">${Utils.formatCurrency(b.amount)}</td>
                    </tr>
                `).join('')}
                </tbody>
            </table>
            </div>
        `;
        }
    },

    /**
     * Draw category bar chart
     */
    drawCategoryChart(categories) {
        const ctx = document.getElementById('categoryChart');
        if (!ctx || !categories || categories.length === 0) return;

        const labels = categories.map(c => c.category);
        const paidData = categories.map(c => c.amountPaid);
        const balanceData = categories.map(c => c.balance);

        if (this.categoryChart) {
            this.categoryChart.destroy();
        }

        this.categoryChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Amount Paid',
                        data: paidData,
                        backgroundColor: 'rgba(40, 167, 69, 0.7)'
                    },
                    {
                        label: 'Balance (Unpaid)',
                        data: balanceData,
                        backgroundColor: 'rgba(220, 53, 69, 0.7)'
                    }
                ]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'top' },
                    title: { display: false }
                },
                scales: {
                    y: {
                        ticks: {
                            callback: function (value) {
                                return '₦' + Number(value).toLocaleString();
                            }
                        }
                    }
                }
            }
        });
    },

    drawCategoryParetoChart(categories) {
        const ctx = document.getElementById('categoryParetoChart');
        if (!ctx || !categories || !categories.length) return;

        // Sort by unpaid balance desc
        const sorted = [...categories].sort((a, b) => (b.balance || 0) - (a.balance || 0));

        const labels = sorted.map(c => c.category);
        const balances = sorted.map(c => Number(c.balance || 0));
        const total = balances.reduce((s, v) => s + v, 0) || 1;

        // Cumulative %
        let running = 0;
        const cumPct = balances.map(v => {
            running += v;
            return Number(((running / total) * 100).toFixed(2));
        });

        if (this.categoryParetoChart) this.categoryParetoChart.destroy();

        this.categoryParetoChart = new Chart(ctx, {
            data: {
                labels,
                datasets: [
                    {
                        type: 'bar',
                        label: 'Unpaid Balance',
                        data: balances,
                        backgroundColor: 'rgba(220, 53, 69, 0.65)',
                        yAxisID: 'y'
                    },
                    {
                        type: 'line',
                        label: 'Cumulative %',
                        data: cumPct,
                        borderColor: 'rgba(0, 123, 255, 1)',
                        backgroundColor: 'rgba(0, 123, 255, 0.15)',
                        tension: 0.25,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                plugins: { legend: { position: 'top' } },
                scales: {
                    y: {
                        ticks: { callback: v => '₦' + Number(v).toLocaleString() }
                    },
                    y1: {
                        position: 'right',
                        min: 0,
                        max: 100,
                        grid: { drawOnChartArea: false },
                        ticks: { callback: v => v + '%' }
                    }
                }
            }
        });
    },

    /**
     * Draw monthly line chart
     */
    drawMonthlyChart(months) {
        const ctx = document.getElementById('monthlyChart');
        if (!ctx || !months || months.length === 0) return;

        const labels = months.map(m => m.month);
        const paidData = months.map(m => m.paidAmount);
        const unpaidData = months.map(m => m.unpaidAmount);

        if (this.monthlyChart) {
            this.monthlyChart.destroy();
        }

        this.monthlyChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Paid Amount',
                        data: paidData,
                        borderColor: 'rgba(40, 167, 69, 1)',
                        backgroundColor: 'rgba(40, 167, 69, 0.2)',
                        tension: 0.3
                    },
                    {
                        label: 'Unpaid Amount',
                        data: unpaidData,
                        borderColor: 'rgba(255, 193, 7, 1)',
                        backgroundColor: 'rgba(255, 193, 7, 0.2)',
                        tension: 0.3
                    }
                ]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'top' },
                    title: { display: false }
                },
                scales: {
                    y: {
                        ticks: {
                            callback: function (value) {
                                return '₦' + Number(value).toLocaleString();
                            }
                        }
                    }
                }
            }
        });
    },

    renderFinancialInsights(summary) {
        const container = document.getElementById('financialInsightsCards');
        if (!container || !summary) return;

        const totalVoucherAmount = this.getTotalVoucherAmount(summary);
        const paid = Number(summary.totalPaidAmount || 0);
        const unpaid = Number(summary.totalUnpaidAmount || 0);
        const cancelled = Number(summary.totalCancelledAmount || 0);
        const contractSum = Number(summary.totalProcessedContractSum || 0);
        const paymentEfficiency = totalVoucherAmount > 0 ? (paid / totalVoucherAmount) * 100 : 0;
        const accrualRate = totalVoucherAmount > 0 ? (unpaid / totalVoucherAmount) * 100 : 0;
        const commitmentGap = Math.max(contractSum - paid, 0);
        const avgVoucher = this.voucherStats?.average || (summary.totalVouchersRaised ? (totalVoucherAmount / summary.totalVouchersRaised) : 0);

        container.innerHTML = `
            <div class="stat-card info">
                <div class="stat-label">Total Voucher Amount (Raised)</div>
                <div class="stat-value">${Utils.formatCurrency(totalVoucherAmount)}</div>
                <div class="stat-subvalue">Paid + Unpaid + Cancelled</div>
            </div>
            <div class="stat-card paid">
                <div class="stat-label">Cash Outflow (Paid)</div>
                <div class="stat-value">${Utils.formatCurrency(paid)}</div>
                <div class="stat-subvalue">Actual cash impact</div>
            </div>
            <div class="stat-card unpaid">
                <div class="stat-label">Accrued Expenses</div>
                <div class="stat-value">${Utils.formatCurrency(unpaid)}</div>
                <div class="stat-subvalue">${accrualRate.toFixed(1)}% of raised</div>
            </div>
            <div class="stat-card info">
                <div class="stat-label">Payment Efficiency</div>
                <div class="stat-value">${paymentEfficiency.toFixed(1)}%</div>
                <div class="stat-subvalue">
                    Paid / Total Raised
                    <div class="progress-bar-mini" style="margin-top:5px;">
                        <div class="progress-fill success" style="width: ${Math.min(paymentEfficiency, 100)}%"></div>
                    </div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Average Voucher Value</div>
                <div class="stat-value">${Utils.formatCurrency(avgVoucher)}</div>
                <div class="stat-subvalue">Per voucher (amount)</div>
            </div>
            <div class="stat-card cancelled">
                <div class="stat-label">Commitment Gap</div>
                <div class="stat-value">${Utils.formatCurrency(commitmentGap)}</div>
                <div class="stat-subvalue">Contract Sum - Paid</div>
            </div>
        `;
    },

    drawCashCommitChart(summary) {
        const ctx = document.getElementById('cashCommitChart');
        if (!ctx || !summary) return;

        const totalVoucherAmount = this.getTotalVoucherAmount(summary);
        const paid = Number(summary.totalPaidAmount || 0);
        const contractSum = Number(summary.totalProcessedContractSum || 0);

        if (this.cashCommitChart) this.cashCommitChart.destroy();

        this.cashCommitChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Paid (Cash Outflow)', 'Total Voucher Amount', 'Contract Sum'],
                datasets: [{
                    label: 'Amount',
                    data: [paid, totalVoucherAmount, contractSum],
                    backgroundColor: ['rgba(40,167,69,0.7)', 'rgba(0,123,255,0.6)', 'rgba(255,193,7,0.7)']
                }]
            },
            options: {
                responsive: true,
                plugins: { legend: { display: false } },
                scales: {
                    y: { ticks: { callback: v => 'â‚¦' + Number(v).toLocaleString() } }
                }
            }
        });
    },

    drawAccrualsSnapshot(summary) {
        const ctx = document.getElementById('accrualsSnapshotChart');
        if (!ctx || !summary) return;

        const paid = Number(summary.totalPaidAmount || 0);
        const unpaid = Number(summary.totalUnpaidAmount || 0);
        const cancelled = Number(summary.totalCancelledAmount || 0);
        const total = paid + unpaid + cancelled;
        const accrualRate = total > 0 ? (unpaid / total) * 100 : 0;
        const contractSum = Number(summary.totalProcessedContractSum || 0);
        const accrualToContract = contractSum > 0 ? (unpaid / contractSum) * 100 : 0;

        if (this.accrualsSnapshotChart) this.accrualsSnapshotChart.destroy();

        this.accrualsSnapshotChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Total Raised'],
                datasets: [
                    { label: 'Paid', data: [paid], backgroundColor: 'rgba(40,167,69,0.7)' },
                    { label: 'Unpaid (Accrual)', data: [unpaid], backgroundColor: 'rgba(220,53,69,0.7)' },
                    { label: 'Cancelled', data: [cancelled], backgroundColor: 'rgba(108,117,125,0.6)' }
                ]
            },
            options: {
                responsive: true,
                plugins: { legend: { position: 'top' } },
                scales: {
                    x: { stacked: true },
                    y: {
                        stacked: true,
                        ticks: { callback: v => 'â‚¦' + Number(v).toLocaleString() }
                    }
                }
            }
        });

        const note = document.getElementById('accrualsSnapshotNote');
        if (note) {
            note.textContent = `Accrual rate: ${accrualRate.toFixed(1)}% of raised | Accrual vs Contract: ${accrualToContract.toFixed(1)}%`;
        }
    },

    drawCategoryShareShiftChart(categories) {
        const ctx = document.getElementById('categoryShareShiftChart');
        if (!ctx || !categories || !categories.length) return;

        const totalPaid = categories.reduce((s, c) => s + Number(c.amountPaid || 0), 0) || 1;
        const totalUnpaid = categories.reduce((s, c) => s + Number(c.balance || 0), 0) || 1;

        const rows = categories.map(c => ({
            category: c.category,
            paidShare: Number(((Number(c.amountPaid || 0) / totalPaid) * 100).toFixed(2)),
            unpaidShare: Number(((Number(c.balance || 0) / totalUnpaid) * 100).toFixed(2))
        }));

        rows.sort((a, b) => b.unpaidShare - a.unpaidShare);

        if (this.categoryShareShiftChart) this.categoryShareShiftChart.destroy();

        this.categoryShareShiftChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: rows.map(r => r.category),
                datasets: [
                    {
                        label: 'Paid Share %',
                        data: rows.map(r => r.paidShare),
                        backgroundColor: 'rgba(40,167,69,0.7)'
                    },
                    {
                        label: 'Unpaid Share %',
                        data: rows.map(r => r.unpaidShare),
                        backgroundColor: 'rgba(220,53,69,0.7)'
                    }
                ]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                plugins: { legend: { position: 'top' } },
                scales: {
                    x: { ticks: { callback: v => v + '%' } }
                }
            }
        });
    },

    drawNetPayableChart(allYearsData) {
        const ctx = document.getElementById('netPayableChart');
        if (!ctx || !allYearsData || !allYearsData.yearsSummary) return;

        const total = Number(allYearsData.grandTotals?.currentOutstandingBalance || 0);
        const currentRow = allYearsData.yearsSummary.find(y => y.label === this.currentYear);
        const current = Number(currentRow?.currentBalance || 0);
        const prior = Math.max(total - current, 0);

        if (this.netPayableChart) this.netPayableChart.destroy();

        this.netPayableChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Current Year', 'Prior Years'],
                datasets: [{
                    data: [current, prior],
                    backgroundColor: ['rgba(0,123,255,0.7)', 'rgba(108,117,125,0.5)']
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'bottom' },
                    tooltip: {
                        callbacks: {
                            label: c => `${c.label}: â‚¦${Number(c.raw || 0).toLocaleString()}`
                        }
                    }
                }
            }
        });

        const kpi = document.getElementById('netPayableKPI');
        if (kpi) {
            const share = total > 0 ? (current / total) * 100 : 0;
            kpi.innerHTML = `
                <div class="stat-label">Net Payable Position</div>
                <div class="stat-value text-danger">${Utils.formatCurrency(total)}</div>
                <div class="stat-subvalue">Current year share: ${share.toFixed(1)}%</div>
            `;
        }
    },

    async loadTaxSummary() {
        if (!API.getTaxSummary) return;
        const result = await API.getTaxSummary(this.currentYear);
        if (result.success) {
            this.taxSummary = result.summary;
            this.renderTaxSummary(result.summary);
        }
    },

    renderTaxSummary(summary) {
        if (!summary) return;
        
        const liabilityEl = document.getElementById('taxTotalLiability');
        const paidEl = document.getElementById('taxTotalPaid');
        const outstandingEl = document.getElementById('taxTotalOutstanding');
        const complianceEl = document.getElementById('taxComplianceRate');
        
        if (liabilityEl) liabilityEl.textContent = Utils.formatCurrency(summary.totalTaxLiability);
        if (paidEl) paidEl.textContent = Utils.formatCurrency(summary.totalPaid);
        if (outstandingEl) outstandingEl.textContent = Utils.formatCurrency(summary.totalOutstanding);
        if (complianceEl) complianceEl.textContent = (summary.complianceRate || 0).toFixed(1) + '%';
    },

    async loadTaxByMonth() {
        if (!API.getTaxByMonth) return;
        const result = await API.getTaxByMonth(this.currentYear);
        if (result.success) {
            this.taxMonthlyData = result;
            this.drawTaxSplitChart(result);
        } else {
            Utils.showToast(result.error || 'Failed to load tax monthly data', 'error');
        }
    },

    drawTaxSplitChart(data) {
        const ctx = document.getElementById('taxSplitChart');
        if (!ctx || !data) return;
        const months = data.months || data;
        if (!months || !months.length) return;

        const labels = months.map(m => m.month);
        const liability = months.map(m => Number(m.totalTax || 0));
        const paid = months.map(m => Number(m.paidTax || 0));

        if (this.taxSplitChart) this.taxSplitChart.destroy();

        this.taxSplitChart = new Chart(ctx, {
            data: {
                labels,
                datasets: [
                    {
                        type: 'bar',
                        label: 'Total Tax Liability',
                        data: liability,
                        backgroundColor: 'rgba(255,193,7,0.7)'
                    },
                    {
                        type: 'line',
                        label: 'Tax Paid',
                        data: paid,
                        borderColor: 'rgba(40,167,69,1)',
                        backgroundColor: 'rgba(40,167,69,0.2)',
                        tension: 0.3,
                        yAxisID: 'y'
                    }
                ]
            },
            options: {
                responsive: true,
                plugins: { legend: { position: 'top' } },
                scales: {
                    y: { ticks: { callback: v => 'â‚¦' + Number(v).toLocaleString() } }
                }
            }
        });
    },

    async loadVoucherAnalytics() {
        const year = this.currentYear;
        if (this.voucherCache[year]) {
            this.applyVoucherAnalytics(this.voucherCache[year]);
            return;
        }

        const all = [];
        try {
            const firstRes = await API.getVouchers(year, null, 1, 200);
            if (!firstRes.success) return;

            all.push(...(firstRes.vouchers || []));
            const totalPages = firstRes.totalPages || 1;

            if (totalPages > 1) {
                const promises = [];
                for (let p = 2; p <= totalPages; p++) {
                    promises.push(API.getVouchers(year, null, p, 200));
                }
                const results = await Promise.all(promises);
                results.forEach(res => {
                    if (res.success && res.vouchers) all.push(...res.vouchers);
                });
            }
        } catch (e) {
            console.error('Voucher analytics load error:', e);
        }

        this.voucherCache[year] = all;
        this.applyVoucherAnalytics(all);
    },

    getMonthIndexFromVoucher(voucher) {
        const pmtMonth = voucher?.pmtMonth || voucher?.paymentMonth;
        if (pmtMonth) {
            const idx = CONFIG.MONTHS.findIndex(m => m.toLowerCase() === String(pmtMonth).toLowerCase());
            if (idx >= 0) return idx;
        }

        const date = this.parseDateFlexible(voucher?.createdAt || voucher?.date);
        if (date) return date.getMonth();
        return null;
    },

    applyVoucherAnalytics(vouchers) {
        const stats = this.computeVoucherStats(vouchers || []);
        this.voucherStats = stats;
        if (this.summaryData?.summary) {
            const rev = this.getRevalidationCounts(vouchers || []);
            this.summaryData.summary.revalidatedVouchers = rev.revalidatedVouchers;
            this.summaryData.summary.revalidatedWithoutOldNumber = rev.revalidatedWithoutOldNumber;
            this.renderYearSummary(this.summaryData);
        }
        const fromVouchers = this.buildAccountTypeBreakdownFromVouchers(vouchers || []);
        if (fromVouchers.length) {
            this.renderAccountTypeTable(fromVouchers);
        }
        this.renderFinancialInsights(this.summaryData?.summary);
        this.renderVoucherValueCards(stats);
        this.drawVoucherDistributionChart(stats);
        this.drawRevalidatedImpactChart(stats);
        this.drawCancelledImpactChart(stats);
        this.drawAgingAnalysisFromVouchers(vouchers || []);
    },

    computeVoucherStats(vouchers) {
        const amounts = vouchers.map(v => this.getVoucherAmount(v)).filter(v => v > 0);
        const total = amounts.reduce((s, v) => s + v, 0);
        const count = amounts.length;
        const avg = count ? total / count : 0;

        const sorted = [...amounts].sort((a, b) => a - b);
        const pct = (p) => {
            if (!sorted.length) return 0;
            const idx = Math.min(sorted.length - 1, Math.floor(p * (sorted.length - 1)));
            return sorted[idx];
        };

        const revalidated = vouchers.filter(v => this.isRevalidatedVoucher(v));
        const revalidatedAmount = revalidated.reduce((s, v) => s + this.getVoucherAmount(v), 0);

        const cancelled = vouchers.filter(v => String(v.status || '').toLowerCase() === 'cancelled');
        const cancelledAmount = cancelled.reduce((s, v) => s + this.getVoucherAmount(v), 0);

        const cancelledByMonth = Array(12).fill(0);
        const cancelledCountByMonth = Array(12).fill(0);
        cancelled.forEach(v => {
            const idx = this.getMonthIndexFromVoucher(v);
            if (idx === null || idx === undefined) return;
            cancelledByMonth[idx] += this.getVoucherAmount(v);
            cancelledCountByMonth[idx] += 1;
        });

        return {
            totalAmount: total,
            count,
            average: avg,
            median: pct(0.5),
            p75: pct(0.75),
            p90: pct(0.9),
            amounts: sorted,
            revalidatedCount: revalidated.length,
            revalidatedAmount,
            cancelledCount: cancelled.length,
            cancelledAmount,
            cancelledByMonth,
            cancelledCountByMonth
        };
    },

    renderVoucherValueCards(stats) {
        const container = document.getElementById('voucherValueCards');
        if (!container || !stats) return;

        container.innerHTML = `
            <div class="stat-card">
                <div class="stat-label">Average Voucher Value</div>
                <div class="stat-value">${Utils.formatCurrency(stats.average)}</div>
                <div class="stat-subvalue">Mean amount</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Median Voucher Value</div>
                <div class="stat-value">${Utils.formatCurrency(stats.median)}</div>
                <div class="stat-subvalue">P50</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">P75 Voucher Value</div>
                <div class="stat-value">${Utils.formatCurrency(stats.p75)}</div>
                <div class="stat-subvalue">Upper quartile</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">P90 Voucher Value</div>
                <div class="stat-value">${Utils.formatCurrency(stats.p90)}</div>
                <div class="stat-subvalue">High value threshold</div>
            </div>
        `;
    },

    drawVoucherDistributionChart(stats) {
        const ctx = document.getElementById('voucherDistChart');
        if (!ctx || !stats || !stats.amounts || stats.amounts.length === 0) return;

        const values = stats.amounts;
        const min = values[0];
        const max = values[values.length - 1];
        const binCount = 6;
        const step = (max - min) / binCount || 1;

        const bins = Array(binCount).fill(0);
        values.forEach(v => {
            const idx = Math.min(binCount - 1, Math.floor((v - min) / step));
            bins[idx] += 1;
        });

        const labels = bins.map((_, i) => {
            const start = min + (step * i);
            const end = i === binCount - 1 ? max : (start + step);
            return `${Utils.formatCurrency(start)} - ${Utils.formatCurrency(end)}`;
        });

        if (this.voucherDistChart) this.voucherDistChart.destroy();

        this.voucherDistChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [{
                    label: 'Voucher Count',
                    data: bins,
                    backgroundColor: 'rgba(0,123,255,0.6)'
                }]
            },
            options: {
                responsive: true,
                plugins: { legend: { position: 'top' } },
                scales: {
                    y: { ticks: { callback: v => Number(v).toLocaleString() } }
                }
            }
        });
    },

    drawRevalidatedImpactChart(stats) {
        const ctx = document.getElementById('revalidatedImpactChart');
        if (!ctx || !stats) return;

        const totalAmount = this.getTotalVoucherAmount(this.summaryData?.summary || {});
        const revalidatedAmount = Number(stats.revalidatedAmount || 0);
        const newAmount = Math.max(totalAmount - revalidatedAmount, 0);
        const share = totalAmount > 0 ? (revalidatedAmount / totalAmount) * 100 : 0;

        if (this.revalidatedImpactChart) this.revalidatedImpactChart.destroy();

        this.revalidatedImpactChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Revalidated', 'New'],
                datasets: [{
                    data: [revalidatedAmount, newAmount],
                    backgroundColor: ['rgba(0,123,255,0.7)', 'rgba(108,117,125,0.4)']
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'bottom' },
                    tooltip: {
                        callbacks: {
                            label: c => `${c.label}: â‚¦${Number(c.raw || 0).toLocaleString()}`
                        }
                    }
                }
            }
        });

        const kpi = document.getElementById('revalidatedImpactKPI');
        if (kpi) {
            kpi.innerHTML = `
                <div class="stat-label">Revalidated Share</div>
                <div class="stat-value">${share.toFixed(1)}%</div>
                <div class="stat-subvalue">${Utils.formatCurrency(revalidatedAmount)} (${Utils.formatNumber(stats.revalidatedCount)} vouchers)</div>
            `;
        }
    },

    drawCancelledImpactChart(stats) {
        const ctx = document.getElementById('cancelledImpactChart');
        if (!ctx || !stats) return;

        const labels = CONFIG.MONTHS;
        const amountSeries = stats.cancelledByMonth || Array(12).fill(0);
        const countSeries = stats.cancelledCountByMonth || Array(12).fill(0);

        if (this.cancelledImpactChart) this.cancelledImpactChart.destroy();

        this.cancelledImpactChart = new Chart(ctx, {
            data: {
                labels,
                datasets: [
                    {
                        type: 'bar',
                        label: 'Cancelled Amount',
                        data: amountSeries,
                        backgroundColor: 'rgba(220,53,69,0.65)',
                        yAxisID: 'y'
                    },
                    {
                        type: 'line',
                        label: 'Cancelled Count',
                        data: countSeries,
                        borderColor: 'rgba(108,117,125,1)',
                        backgroundColor: 'rgba(108,117,125,0.2)',
                        tension: 0.3,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                plugins: { legend: { position: 'top' } },
                scales: {
                    y: { ticks: { callback: v => 'â‚¦' + Number(v).toLocaleString() } },
                    y1: {
                        position: 'right',
                        grid: { drawOnChartArea: false },
                        ticks: { callback: v => Number(v).toLocaleString() }
                    }
                }
            }
        });

        const totalAmount = this.getTotalVoucherAmount(this.summaryData?.summary || {});
        const cancelledAmount = Number(stats.cancelledAmount || 0);
        const share = totalAmount > 0 ? (cancelledAmount / totalAmount) * 100 : 0;

        const kpi = document.getElementById('cancelledImpactKPI');
        if (kpi) {
            kpi.innerHTML = `
                <div class="stat-label">Cancelled Impact</div>
                <div class="stat-value">${Utils.formatCurrency(cancelledAmount)}</div>
                <div class="stat-subvalue">${share.toFixed(1)}% of total (${Utils.formatNumber(stats.cancelledCount)} vouchers)</div>
            `;
        }
    },

    /**
     * Load all years summary for debt tracking
     */
    async loadAllYearsSummary() {
        const result = await API.getAllYearsSummary();

        if (result.success) {
            this.allYearsData = result;
            this.renderAllYearsSummary(result);
            this.drawDebtTrendChart(result);
            this.drawNetPayableChart(result);
        } else {
            Utils.showToast(result.error || 'Failed to load all-years summary', 'error');
        }
    },

    /**
     * Load debt profile (2026)
     */
    async loadDebtProfile() {
        const result = await API.getDebtProfile();

        if (result.success) {
            this.debtProfile = result;
            this.renderDebtProfile(result);
            this.drawDebtConcentration(result);
            this.drawTopDebtorsChart(result);
        } else {
            Utils.showToast(result.error || 'Failed to load debt profile', 'error');
        }
    },

    drawStatusCharts(summary) {
        if (!summary) return;

        const countCtx = document.getElementById('statusCountChart');
        const amtCtx = document.getElementById('statusAmountChart');
        if (!countCtx || !amtCtx) return;

        const paidCount = Number(summary.paidVouchers || 0);
        const unpaidCount = Number(summary.unpaidVouchers || 0);
        const cancelledCount = Number(summary.cancelledVouchers || 0);

        const paidAmt = Number(summary.totalPaidAmount || 0);
        const unpaidAmt = Number(summary.totalUnpaidAmount || 0);
        const cancelledAmt = Number(summary.totalCancelledAmount || 0);

        // Destroy previous
        if (this.statusCountChart) this.statusCountChart.destroy();
        if (this.statusAmountChart) this.statusAmountChart.destroy();

        const labels = ['Paid', 'Unpaid', 'Cancelled'];
        const colors = ['#28a745', '#ffc107', '#dc3545'];

        this.statusCountChart = new Chart(countCtx, {
            type: 'doughnut',
            data: {
                labels,
                datasets: [{ data: [paidCount, unpaidCount, cancelledCount], backgroundColor: colors }]
            },
            options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
        });

        this.statusAmountChart = new Chart(amtCtx, {
            type: 'doughnut',
            data: {
                labels,
                datasets: [{ data: [paidAmt, unpaidAmt, cancelledAmt], backgroundColor: colors }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'bottom' },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => `${ctx.label}: ₦${Number(ctx.raw || 0).toLocaleString()}`
                        }
                    }
                }
            }
        });
    },

    drawDebtTrendChart(allYearsData) {
        const ctx = document.getElementById('debtTrendChart');
        if (!ctx || !allYearsData || !allYearsData.yearsSummary) return;

        const points = allYearsData.yearsSummary
            .filter(y => !y.error && y.label)
            .map(y => ({ year: y.label, balance: Number(y.currentBalance || 0) }));

        if (this.debtTrendChart) this.debtTrendChart.destroy();

        this.debtTrendChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: points.map(p => p.year),
                datasets: [{
                    label: 'Outstanding Balance',
                    data: points.map(p => p.balance),
                    borderColor: '#dc3545',
                    backgroundColor: 'rgba(220,53,69,0.15)',
                    tension: 0.3,
                    fill: true
                }]
            },
            options: {
                responsive: true,
                plugins: { legend: { position: 'bottom' } },
                scales: {
                    y: {
                        ticks: { callback: (v) => '₦' + Number(v).toLocaleString() }
                    }
                }
            }
        });
    },

    drawTopDebtorsChart(debtProfile) {
        const ctx = document.getElementById('topDebtorsChart');
        if (!ctx || !debtProfile || !debtProfile.topDebtors) return;

        const top = debtProfile.topDebtors.slice(0, 10);
        const labels = top.map(d => Utils.truncate(d.payee || '-', 18));
        const values = top.map(d => Number(d.amount || 0));

        if (this.topDebtorsChart) this.topDebtorsChart.destroy();

        this.topDebtorsChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [{
                    label: 'Amount Owed',
                    data: values,
                    backgroundColor: 'rgba(220,53,69,0.75)'
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                plugins: { legend: { position: 'bottom' } },
                scales: {
                    x: {
                        ticks: { callback: (v) => '₦' + Number(v).toLocaleString() }
                    }
                }
            }
        });
    },

    downloadChart(canvasId, filenameBase) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) {
            Utils.showToast('Chart not found', 'warning');
            return;
        }

        const link = document.createElement('a');
        link.href = canvas.toDataURL('image/png');
        link.download = `${filenameBase}_${this.currentYear}_${new Date().toISOString().slice(0, 10)}.png`;
        link.click();
    },

    drawDebtConcentration(data) {
        const ctx = document.getElementById('debtConcentrationChart');
        if (!ctx || !data) return;

        const totalDebt = Number(data.totalDebt || 0);
        const top10 = (data.topDebtors || []).slice(0, 10);
        const top10Sum = top10.reduce((s, d) => s + Number(d.amount || 0), 0);
        const others = Math.max(totalDebt - top10Sum, 0);

        const pct = totalDebt > 0 ? ((top10Sum / totalDebt) * 100) : 0;

        if (this.debtConcentrationChart) this.debtConcentrationChart.destroy();

        this.debtConcentrationChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Top 10 Payees', 'Others'],
                datasets: [{
                    data: [top10Sum, others],
                    backgroundColor: ['rgba(220,53,69,0.8)', 'rgba(108,117,125,0.4)']
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'bottom' },
                    tooltip: {
                        callbacks: {
                            label: (c) => `${c.label}: ₦${Number(c.raw || 0).toLocaleString()}`
                        }
                    }
                }
            }
        });

        // KPI box update
        const kpi = document.getElementById('debtConcentrationKPI');
        if (kpi) {
            kpi.querySelector('.stat-value').textContent = `${pct.toFixed(1)}%`;
        }
    },

    /**
     * Render year summary cards
     * Uses definitions from backend getSummary:
     * - Total Vouchers Raised: count (ACCOUNT OR MAIL not empty)
     * - Paid/Unpaid/Cancelled: amounts from GROSS AMOUNT
     * - Total Contract Sum: sum(CONTRACT SUM)
     * - Total Debt: Contract Sum - (Paid + Cancelled)
     * - Average Payment Rate: Paid count / Total Raised * 100
     * - Revalidated: count(OLD VOUCHER NUMBER not empty OR OLD VOUCHER NO AVAILABLE? = YES)
     */
    renderYearSummary(data) {
        const container = document.getElementById('yearSummaryCards');
        if (!container || !data.summary) return;

        const stats = data.summary;

        container.innerHTML = `
            <!-- Total Vouchers Raised (COUNT) -->
            <div class="stat-card">
                <div class="stat-label">Total Vouchers Raised</div>
                <div class="stat-value">${Utils.formatNumber(stats.totalVouchersRaised)}</div>
                <div class="stat-subvalue">Count of vouchers with Voucher Number</div>
            </div>
            
            <!-- Paid Vouchers (AMOUNT) -->
            <div class="stat-card paid">
                <div class="stat-label">Paid Vouchers (Amount)</div>
                <div class="stat-value">${Utils.formatCurrency(stats.totalPaidAmount)}</div>
                <div class="stat-subvalue">
                    ${Utils.formatNumber(stats.paidVouchers)} voucher(s) paid
                </div>
            </div>
            
            <!-- Unpaid Vouchers (AMOUNT) -->
            <div class="stat-card unpaid">
                <div class="stat-label">Unpaid Vouchers (Amount)</div>
                <div class="stat-value">${Utils.formatCurrency(stats.totalUnpaidAmount)}</div>
                <div class="stat-subvalue">
                    ${Utils.formatNumber(stats.unpaidVouchers)} voucher(s) unpaid
                </div>
            </div>
            
            <!-- Cancelled Vouchers (AMOUNT) -->
            <div class="stat-card cancelled">
                <div class="stat-label">Cancelled Vouchers (Amount)</div>
                <div class="stat-value">${Utils.formatCurrency(stats.totalCancelledAmount)}</div>
                <div class="stat-subvalue">
                    ${Utils.formatNumber(stats.cancelledVouchers)} voucher(s) cancelled
                </div>
            </div>
            
            <!-- Total Contract Sum (AMOUNT) -->
            <div class="stat-card info">
                <div class="stat-label">Total Contract Sum</div>
                <div class="stat-value">${Utils.formatCurrency(stats.totalProcessedContractSum)}</div>
            <div class="stat-subvalue">From Processed Contracts</div>
                </div>
            
            <!-- Total Debt (AMOUNT) -->
            <div class="stat-card">
                <div class="stat-label">Total Debt</div>
                <div class="stat-value text-danger">${Utils.formatCurrency(stats.totalDebt)}</div>
                <div class="stat-subvalue">
                    = Contract Sum - (Paid + Cancelled)
                </div>
            </div>
            
            <!-- Average Payment Rate -->
            <div class="stat-card info">
                <div class="stat-label">Average Payment Rate</div>
                <div class="stat-value">${stats.averagePaymentPercent}%</div>
                <div class="stat-subvalue">
                    Based on voucher count: Paid / Total Raised
                    <div class="progress-bar-mini" style="margin-top:5px;">
                        <div class="progress-fill" style="width: ${Math.min(stats.averagePaymentPercent, 100)}%"></div>
                    </div>
                </div>
            </div>
            
            <!-- Revalidated Vouchers (COUNT) -->
            <div class="stat-card">
                <div class="stat-label">Revalidated Vouchers</div>
                <div class="stat-value">${Utils.formatNumber(stats.revalidatedVouchers)}</div>
                <div class="stat-subvalue">Old number present or marked available</div>
            </div>

            <!-- Revalidated Without Old Voucher Number -->
            <div class="stat-card">
                <div class="stat-label">Revalidated (No Old Voucher No.)</div>
                <div class="stat-value">${Utils.formatNumber(stats.revalidatedWithoutOldNumber || 0)}</div>
                <div class="stat-subvalue">Old number not available</div>
            </div>
        `;
    },

    /**
     * Render category breakdown table
     */
    renderCategoryTable(categories) {
        const container = document.getElementById('categoryTable');
        if (!container) return;

        if (!categories || categories.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No category data available</p>';
            return;
        }

        // Calculate totals
        let totalVouchers = 0;
        let totalPaid = 0;
        let totalBalance = 0;

        categories.forEach(cat => {
            totalVouchers += cat.vouchersRaised;
            totalPaid += cat.amountPaid;
            totalBalance += cat.balance;
        });

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Category</th>
                            <th class="text-center">Vouchers Raised</th>
                            <th class="text-right">Amount Paid</th>
                            <th class="text-right">Balance</th>
                            <th class="text-center">% Paid</th>
                            <th class="text-center">% of Total Pmt</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        categories.forEach(cat => {
            const percentBar = `
                <div class="progress-bar-mini">
                    <div class="progress-fill ${cat.percentagePaid >= 70 ? 'success' : cat.percentagePaid >= 40 ? 'warning' : 'danger'}" 
                         style="width: ${Math.min(cat.percentagePaid, 100)}%"></div>
                </div>
                <small>${cat.percentagePaid}%</small>
            `;

            html += `
                <tr>
                    <td><strong>${cat.category}</strong></td>
                    <td class="text-center">${Utils.formatNumber(cat.vouchersRaised)}</td>
                    <td class="text-right text-success">${Utils.formatCurrency(cat.amountPaid)}</td>
                    <td class="text-right text-danger">${Utils.formatCurrency(cat.balance)}</td>
                    <td class="text-center">${percentBar}</td>
                    <td class="text-center">${cat.percentOfTotalPayment}%</td>
                </tr>
            `;
        });

        // Totals row
        const totalPercentPaid = totalPaid + totalBalance > 0
            ? ((totalPaid / (totalPaid + totalBalance)) * 100).toFixed(2)
            : 0;

        html += `
            <tr class="totals-row">
                <td><strong>TOTAL</strong></td>
                <td class="text-center"><strong>${Utils.formatNumber(totalVouchers)}</strong></td>
                <td class="text-right text-success"><strong>${Utils.formatCurrency(totalPaid)}</strong></td>
                <td class="text-right text-danger"><strong>${Utils.formatCurrency(totalBalance)}</strong></td>
                <td class="text-center"><strong>${totalPercentPaid}%</strong></td>
                <td class="text-center"><strong>100%</strong></td>
            </tr>
        `;

        html += '</tbody></table></div>';
        container.innerHTML = html;
    },

    getConfiguredAccountTypes() {
        return Object.keys(this.systemConfig?.accountTypes || {});
    },

    getConfiguredSubTypes(baseType) {
        const map = this.systemConfig?.accountTypes || {};
        return Array.isArray(map[baseType]) ? map[baseType] : [];
    },

    buildAccountTypeBreakdownFromVouchers(vouchers) {
        const stats = {};

        (vouchers || []).forEach((v) => {
            const baseType = String(v?.accountType || '').trim() || 'Unspecified';
            const subType = String(v?.subAccountType || '').trim();
            const key = `${baseType}::${subType}`;
            if (!stats[key]) {
                stats[key] = {
                    accountType: baseType,
                    subAccountType: subType,
                    count: 0,
                    totalAmount: 0,
                    paidAmount: 0,
                    unpaidAmount: 0
                };
            }

            const amount = this.getVoucherAmount(v);
            const status = String(v?.status || '').trim().toLowerCase();
            stats[key].count += 1;
            stats[key].totalAmount += amount;
            if (status === 'paid') stats[key].paidAmount += amount;
            if (status === 'unpaid') stats[key].unpaidAmount += amount;
        });

        return Object.values(stats).map((row) => {
            const paymentRate = row.totalAmount > 0 ? (row.paidAmount / row.totalAmount) * 100 : 0;
            return {
                ...row,
                displayName: row.subAccountType ? `${row.accountType} (${row.subAccountType})` : row.accountType,
                paymentRate: Number(paymentRate.toFixed(2))
            };
        });
    },

    parseAccountTypeRow(row) {
        const rawBase = String(row?.accountType || '').trim();
        const rawSub = String(row?.subAccountType || '').trim();

        if (rawSub) {
            return { baseType: rawBase || 'Unspecified', subType: rawSub };
        }

        const m = rawBase.match(/^(.*)\((.*)\)$/);
        if (m) {
            const base = String(m[1] || '').trim();
            const sub = String(m[2] || '').trim();
            return { baseType: base || 'Unspecified', subType: sub };
        }

        return { baseType: rawBase || 'Unspecified', subType: '' };
    },

    populateAccountTypeFilters(accountTypes) {
        const parentSelect = document.getElementById('accountTypeFilter');
        if (!parentSelect) return;

        const previous = this.accountTypeFilters.accountType || 'ALL';
        const baseTypes = Array.from(new Set([
            ...(accountTypes || []).map((row) => this.parseAccountTypeRow(row).baseType),
            ...this.getConfiguredAccountTypes()
        ])).sort();
        parentSelect.innerHTML = '<option value="ALL">All Account Types</option>';
        baseTypes.forEach((base) => {
            parentSelect.innerHTML += `<option value="${base}">${base}</option>`;
        });

        this.accountTypeFilters.accountType = baseTypes.includes(previous) ? previous : 'ALL';
        parentSelect.value = this.accountTypeFilters.accountType;
        this.populateSubAccountTypeFilter(accountTypes);
    },

    populateSubAccountTypeFilter(accountTypes) {
        const subSelect = document.getElementById('subAccountTypeFilter');
        if (!subSelect) return;

        const selectedBase = this.accountTypeFilters.accountType || 'ALL';
        const previousSub = this.accountTypeFilters.subAccountType || 'ALL';
        const fromRows = (accountTypes || [])
            .map((row) => this.parseAccountTypeRow(row))
            .filter((x) => x.subType && (selectedBase === 'ALL' || x.baseType === selectedBase))
            .map((x) => x.subType);
        const fromConfig = selectedBase === 'ALL'
            ? this.getConfiguredAccountTypes().flatMap((base) => this.getConfiguredSubTypes(base))
            : this.getConfiguredSubTypes(selectedBase);
        const subs = Array.from(new Set([...fromRows, ...fromConfig])).sort();

        subSelect.innerHTML = '<option value="ALL">All Sub Account Types</option>';
        subs.forEach((sub) => {
            subSelect.innerHTML += `<option value="${sub}">${sub}</option>`;
        });

        this.accountTypeFilters.subAccountType = subs.includes(previousSub) ? previousSub : 'ALL';
        subSelect.value = this.accountTypeFilters.subAccountType;
    },

    getFilteredAccountTypeRows(accountTypes) {
        const selectedBase = this.accountTypeFilters.accountType || 'ALL';
        const selectedSub = this.accountTypeFilters.subAccountType || 'ALL';

        return (accountTypes || []).filter((row) => {
            const parsed = this.parseAccountTypeRow(row);
            const baseMatches = selectedBase === 'ALL' || parsed.baseType === selectedBase;
            const subMatches = selectedSub === 'ALL' || parsed.subType === selectedSub;
            return baseMatches && subMatches;
        });
    },

    toggleAccountTypeGroup(baseType) {
        if (this.expandedAccountTypeGroups.has(baseType)) {
            this.expandedAccountTypeGroups.delete(baseType);
        } else {
            this.expandedAccountTypeGroups.add(baseType);
        }
        this.renderAccountTypeTable(this.currentAccountTypeRows);
    },

    renderAccountTypeTable(accountTypes) {
        const container = document.getElementById('accountTypeTable');
        if (!container) return;

        this.currentAccountTypeRows = Array.isArray(accountTypes) ? accountTypes : [];
        this.populateAccountTypeFilters(this.currentAccountTypeRows);
        const rows = this.getFilteredAccountTypeRows(this.currentAccountTypeRows);
        const hasConfiguredTypes = this.getConfiguredAccountTypes().length > 0;

        if (!rows.length && !hasConfiguredTypes) {
            container.innerHTML = '<p class="text-muted text-center">No account type data available</p>';
            return;
        }

        const grouped = {};
        rows.forEach((row) => {
            const parsed = this.parseAccountTypeRow(row);
            const baseType = parsed.baseType;
            const subType = parsed.subType;
            if (!grouped[baseType]) {
                grouped[baseType] = {
                    count: 0,
                    totalAmount: 0,
                    paidAmount: 0,
                    unpaidAmount: 0,
                    children: {}
                };
            }

            grouped[baseType].count += Number(row.count || 0);
            grouped[baseType].totalAmount += Number(row.totalAmount || 0);
            grouped[baseType].paidAmount += Number(row.paidAmount || 0);
            grouped[baseType].unpaidAmount += Number(row.unpaidAmount || 0);

            if (subType) {
                if (!grouped[baseType].children[subType]) {
                    grouped[baseType].children[subType] = {
                        name: subType,
                        count: 0,
                        totalAmount: 0,
                        paidAmount: 0,
                        unpaidAmount: 0
                    };
                }
                grouped[baseType].children[subType].count += Number(row.count || 0);
                grouped[baseType].children[subType].totalAmount += Number(row.totalAmount || 0);
                grouped[baseType].children[subType].paidAmount += Number(row.paidAmount || 0);
                grouped[baseType].children[subType].unpaidAmount += Number(row.unpaidAmount || 0);
            }
        });

        // Ensure configured account types/sub-accounts are represented even when no row exists yet.
        this.getConfiguredAccountTypes().forEach((baseType) => {
            if (!grouped[baseType]) {
                grouped[baseType] = {
                    count: 0,
                    totalAmount: 0,
                    paidAmount: 0,
                    unpaidAmount: 0,
                    children: {}
                };
            }
            this.getConfiguredSubTypes(baseType).forEach((subType) => {
                if (!grouped[baseType].children[subType]) {
                    grouped[baseType].children[subType] = {
                        name: subType,
                        count: 0,
                        totalAmount: 0,
                        paidAmount: 0,
                        unpaidAmount: 0
                    };
                }
            });
        });

        const configuredBaseOrder = this.getConfiguredAccountTypes();
        const baseOrderIndex = {};
        configuredBaseOrder.forEach((name, idx) => { baseOrderIndex[name] = idx; });

        let totalCount = 0;
        let totalAmount = 0;
        let totalPaid = 0;
        let totalUnpaid = 0;

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Account Type</th>
                            <th class="text-center">Vouchers</th>
                            <th class="text-right">Total Amount</th>
                            <th class="text-right">Paid Amount</th>
                            <th class="text-right">Unpaid Amount</th>
                            <th class="text-center">Payment Rate</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        Object.keys(grouped)
            .sort((a, b) => {
                const ai = Object.prototype.hasOwnProperty.call(baseOrderIndex, a) ? baseOrderIndex[a] : Number.MAX_SAFE_INTEGER;
                const bi = Object.prototype.hasOwnProperty.call(baseOrderIndex, b) ? baseOrderIndex[b] : Number.MAX_SAFE_INTEGER;
                if (ai !== bi) return ai - bi;
                return grouped[b].totalAmount - grouped[a].totalAmount;
            })
            .forEach((baseType) => {
                const parent = grouped[baseType];
                const paymentRate = parent.totalAmount > 0 ? (parent.paidAmount / parent.totalAmount) * 100 : 0;
                const children = Object.values(parent.children || {});
                const hasChildren = children.length > 0;
                const isExpanded = this.expandedAccountTypeGroups.has(baseType);
                const subOrder = this.getConfiguredSubTypes(baseType);
                const subOrderIndex = {};
                subOrder.forEach((name, idx) => { subOrderIndex[name] = idx; });

                totalCount += parent.count;
                totalAmount += parent.totalAmount;
                totalPaid += parent.paidAmount;
                totalUnpaid += parent.unpaidAmount;

                const safeBase = baseType.replace(/'/g, "\\'");
                html += `
                    <tr class="parent-account-row">
                        <td>
                            ${hasChildren ? `<button class="btn btn-sm btn-secondary" style="margin-right:8px;padding:2px 8px;" onclick="Reports.toggleAccountTypeGroup('${safeBase}')">${isExpanded ? '-' : '+'}</button>` : ''}
                            <strong>${baseType}</strong>
                        </td>
                        <td class="text-center">${Utils.formatNumber(parent.count)}</td>
                        <td class="text-right">${Utils.formatCurrency(parent.totalAmount)}</td>
                        <td class="text-right text-success">${Utils.formatCurrency(parent.paidAmount)}</td>
                        <td class="text-right text-danger">${Utils.formatCurrency(parent.unpaidAmount)}</td>
                        <td class="text-center">${paymentRate.toFixed(2)}%</td>
                    </tr>
                `;

                if (hasChildren && isExpanded) {
                    children
                        .sort((a, b) => {
                            const ai = Object.prototype.hasOwnProperty.call(subOrderIndex, a.name) ? subOrderIndex[a.name] : Number.MAX_SAFE_INTEGER;
                            const bi = Object.prototype.hasOwnProperty.call(subOrderIndex, b.name) ? subOrderIndex[b.name] : Number.MAX_SAFE_INTEGER;
                            if (ai !== bi) return ai - bi;
                            return b.totalAmount - a.totalAmount;
                        })
                        .forEach((child) => {
                            const childRate = child.totalAmount > 0 ? (child.paidAmount / child.totalAmount) * 100 : 0;
                            html += `
                                <tr class="sub-account-row">
                                    <td style="padding-left:42px;">${child.name}</td>
                                    <td class="text-center">${Utils.formatNumber(child.count)}</td>
                                    <td class="text-right">${Utils.formatCurrency(child.totalAmount)}</td>
                                    <td class="text-right text-success">${Utils.formatCurrency(child.paidAmount)}</td>
                                    <td class="text-right text-danger">${Utils.formatCurrency(child.unpaidAmount)}</td>
                                    <td class="text-center">${childRate.toFixed(2)}%</td>
                                </tr>
                            `;
                        });
                }
            });

        const grandRate = totalAmount > 0 ? (totalPaid / totalAmount) * 100 : 0;
        html += `
            <tr class="totals-row">
                <td><strong>TOTAL</strong></td>
                <td class="text-center"><strong>${Utils.formatNumber(totalCount)}</strong></td>
                <td class="text-right"><strong>${Utils.formatCurrency(totalAmount)}</strong></td>
                <td class="text-right text-success"><strong>${Utils.formatCurrency(totalPaid)}</strong></td>
                <td class="text-right text-danger"><strong>${Utils.formatCurrency(totalUnpaid)}</strong></td>
                <td class="text-center"><strong>${grandRate.toFixed(2)}%</strong></td>
            </tr>
        `;

        html += '</tbody></table></div>';
        container.innerHTML = html;
    },

    /**
     * Render monthly breakdown table
     */
    renderMonthlyTable(months) {
        const container = document.getElementById('monthlyTable');
        if (!container) return;

        if (!months || months.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No monthly data available</p>';
            return;
        }

        let totalCount = 0;
        let totalPaid = 0;
        let totalUnpaid = 0;

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Month</th>
                            <th class="text-center">Voucher Count</th>
                            <th class="text-right">Amount Paid</th>
                            <th class="text-right">Amount Unpaid</th>
                            <th class="text-right">Total</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        months.forEach(month => {
            totalCount += month.count;
            totalPaid += month.paidAmount;
            totalUnpaid += month.unpaidAmount;

            const monthTotal = month.paidAmount + month.unpaidAmount;

            html += `
                <tr>
                    <td><strong>${month.month}</strong></td>
                    <td class="text-center">${Utils.formatNumber(month.count)}</td>
                    <td class="text-right text-success">${Utils.formatCurrency(month.paidAmount)}</td>
                    <td class="text-right text-warning">${Utils.formatCurrency(month.unpaidAmount)}</td>
                    <td class="text-right">${Utils.formatCurrency(monthTotal)}</td>
                </tr>
            `;
        });

        // Totals row
        html += `
            <tr class="totals-row">
                <td><strong>TOTAL</strong></td>
                <td class="text-center"><strong>${Utils.formatNumber(totalCount)}</strong></td>
                <td class="text-right text-success"><strong>${Utils.formatCurrency(totalPaid)}</strong></td>
                <td class="text-right text-warning"><strong>${Utils.formatCurrency(totalUnpaid)}</strong></td>
                <td class="text-right"><strong>${Utils.formatCurrency(totalPaid + totalUnpaid)}</strong></td>
            </tr>
        `;

        html += '</tbody></table></div>';
        container.innerHTML = html;
    },

    /**
     * Render all years summary (debt tracking)
     */
    renderAllYearsSummary(data) {
        const container = document.getElementById('allYearsTable');
        if (!container) return;

        if (!data.yearsSummary || data.yearsSummary.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No data available</p>';
            return;
        }

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th class="text-right">Balance B/F</th>
                            <th class="text-center">Total Vouchers</th>
                            <th class="text-right">Total Amount</th>
                            <th class="text-right">Paid</th>
                            <th class="text-center">Revalidated</th>
                            <th class="text-center">Cancelled</th>
                            <th class="text-right">Current Balance</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        data.yearsSummary.forEach(year => {
            if (year.error) {
                html += `
                    <tr>
                        <td><strong>${year.label}</strong></td>
                        <td colspan="7" class="text-muted text-center">${year.error}</td>
                    </tr>
                `;
            } else {
                html += `
                    <tr>
                        <td><strong>${year.label}</strong></td>
                        <td class="text-right">${Utils.formatCurrency(year.balanceBroughtForward)}</td>
                        <td class="text-center">${Utils.formatNumber(year.totalVouchers || 0)}</td>
                        <td class="text-right">${Utils.formatCurrency(year.totalAmount || 0)}</td>
                        <td class="text-right text-success">${Utils.formatCurrency(year.paidAmount || 0)}</td>
                        <td class="text-center">${Utils.formatNumber(year.revalidatedVouchers || 0)}</td>
                        <td class="text-center">${Utils.formatNumber(year.cancelledVouchers || 0)}</td>
                        <td class="text-right text-danger"><strong>${Utils.formatCurrency(year.currentBalance || 0)}</strong></td>
                    </tr>
                `;
            }
        });

        // Grand totals
        if (data.grandTotals) {
            html += `
                <tr class="totals-row">
                    <td><strong>GRAND TOTAL</strong></td>
                    <td class="text-right">-</td>
                    <td class="text-center"><strong>${Utils.formatNumber(data.grandTotals.totalVouchers)}</strong></td>
                    <td class="text-right"><strong>${Utils.formatCurrency(data.grandTotals.totalAmount)}</strong></td>
                    <td class="text-right text-success"><strong>${Utils.formatCurrency(data.grandTotals.totalPaid)}</strong></td>
                    <td class="text-center"><strong>${Utils.formatNumber(data.grandTotals.totalRevalidated)}</strong></td>
                    <td class="text-center"><strong>-</strong></td>
                    <td class="text-right text-danger"><strong>${Utils.formatCurrency(data.grandTotals.currentOutstandingBalance)}</strong></td>
                </tr>
            `;
        }

        html += '</tbody></table></div>';
        container.innerHTML = html;

        // Update grand total card
        const grandTotalCard = document.getElementById('grandTotalDebt');
        if (grandTotalCard && data.grandTotals) {
            grandTotalCard.innerHTML = `
                <div class="stat-label">Total Outstanding Debt (All Years)</div>
                <div class="stat-value text-danger">${Utils.formatCurrency(data.grandTotals.currentOutstandingBalance)}</div>
                <div class="stat-subvalue">From ${data.yearsSummary.length} years</div>
            `;
        }
    },

    /**
     * Render debt profile (2026)
     */
    renderDebtProfile(data) {
        const categoryContainer = document.getElementById('debtByCategory');
        const debtorsContainer = document.getElementById('topDebtors');

        if (categoryContainer && data.debtByCategory) {
            if (data.debtByCategory.length === 0) {
                categoryContainer.innerHTML = '<p class="text-muted text-center">No unpaid vouchers</p>';
            } else {
                let html = '<div class="debt-list">';
                data.debtByCategory.forEach(cat => {
                    const percent = data.totalDebt > 0 ? ((cat.amount / data.totalDebt) * 100).toFixed(1) : 0;
                    html += `
                        <div class="debt-item">
                            <div class="debt-info">
                                <strong>${cat.category}</strong>
                                <span class="text-muted">(${cat.count} vouchers)</span>
                            </div>
                            <div class="debt-amount">
                                ${Utils.formatCurrency(cat.amount)}
                                <small class="text-muted">${percent}%</small>
                            </div>
                            <div class="debt-bar">
                                <div class="debt-bar-fill" style="width: ${percent}%"></div>
                            </div>
                        </div>
                    `;
                });
                html += '</div>';
                categoryContainer.innerHTML = html;
            }
        }

        if (debtorsContainer && data.topDebtors) {
            if (data.topDebtors.length === 0) {
                debtorsContainer.innerHTML = '<p class="text-muted text-center">No unpaid vouchers</p>';
            } else {
                let html = `
                    <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th>#</th>
                                    <th>Payee</th>
                                    <th class="text-center">Vouchers</th>
                                    <th class="text-right">Amount Owed</th>
                                </tr>
                            </thead>
                            <tbody>
                `;

                data.topDebtors.forEach((debtor, index) => {
                    html += `
                        <tr>
                            <td>${index + 1}</td>
                            <td title="${debtor.payee}">${Utils.truncate(debtor.payee, 30)}</td>
                            <td class="text-center">${debtor.count}</td>
                            <td class="text-right text-danger"><strong>${Utils.formatCurrency(debtor.amount)}</strong></td>
                        </tr>
                    `;
                });

                html += '</tbody></table></div>';
                debtorsContainer.innerHTML = html;
            }
        }

        // Update total debt card for 2026
        const totalDebtCard = document.getElementById('totalDebt2026');
        if (totalDebtCard && data.totalDebt !== undefined) {
            totalDebtCard.innerHTML = `
                <div class="stat-label">Total Unpaid (2026)</div>
                <div class="stat-value text-danger">${Utils.formatCurrency(data.totalDebt)}</div>
            `;
        }
    },

    /**
     * Export data to CSV
     */
    exportToCSV() {
        if (!this.summaryData || !this.summaryData.categoryBreakdown) {
            Utils.showToast('No data to export', 'warning');
            return;
        }

        const s = this.summaryData.summary;

        let csv = 'PAYABLE VOUCHER 2026 - REPORT\n';
        csv += `Generated: ${new Date().toLocaleString()}\n\n`;

        // SUMMARY SECTION
        csv += 'SUMMARY\n';
        csv += `Total Vouchers Raised (Count),${s.totalVouchersRaised}\n`;
        csv += `Paid Vouchers (Amount),${s.totalPaidAmount}\n`;
        csv += `Paid Vouchers (Count),${s.paidVouchers}\n`;
        csv += `Unpaid Vouchers (Amount),${s.totalUnpaidAmount}\n`;
        csv += `Unpaid Vouchers (Count),${s.unpaidVouchers}\n`;
        csv += `Cancelled Vouchers (Amount),${s.totalCancelledAmount}\n`;
        csv += `Cancelled Vouchers (Count),${s.cancelledVouchers}\n`;
        csv += `Total Contract Sum,${s.totalProcessedContractSum}\n`;
        csv += `Total Debt,${s.totalDebt}\n`;
        csv += `Average Payment Rate (%),${s.averagePaymentPercent}\n`;
        csv += `Revalidated Vouchers (Count),${s.revalidatedVouchers}\n`;
        csv += `Revalidated Criteria,OLD VOUCHER NUMBER present OR OLD VOUCHER NO AVAILABLE? = YES\n`;
        csv += `Revalidated (No Old Voucher No.),${s.revalidatedWithoutOldNumber || 0}\n\n`;

        // CATEGORY BREAKDOWN
        csv += 'CATEGORY BREAKDOWN\n';
        csv += 'Category,Vouchers Raised,Amount Paid,Balance,% Paid\n';
        this.summaryData.categoryBreakdown.forEach(cat => {
            csv += `${cat.category},${cat.vouchersRaised},${cat.amountPaid},${cat.balance},${cat.percentagePaid}%\n`;
        });
        csv += '\n';

        if (Array.isArray(this.summaryData.accountTypeBreakdown) && this.summaryData.accountTypeBreakdown.length) {
            csv += 'ACCOUNT TYPE BREAKDOWN\n';
            csv += 'Account Type,Vouchers,Total Amount,Paid Amount,Unpaid Amount,Payment Rate %\n';
            this.summaryData.accountTypeBreakdown.forEach(typeRow => {
                const label = typeRow.displayName || typeRow.accountType || 'Unspecified';
                csv += `${label},${typeRow.count || 0},${typeRow.totalAmount || 0},${typeRow.paidAmount || 0},${typeRow.unpaidAmount || 0},${typeRow.paymentRate || 0}\n`;
            });
            csv += '\n';
        }

        // MONTHLY BREAKDOWN
        csv += 'MONTHLY BREAKDOWN\n';
        csv += 'Month,Count,Paid,Unpaid\n';
        this.summaryData.monthlyBreakdown.forEach(month => {
            csv += `${month.month},${month.count},${month.paidAmount},${month.unpaidAmount}\n`;
        });

        // Download CSV
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `PayableVoucher_Report_${this.currentYear}_${new Date().toISOString().split('T')[0]}.csv`;
        link.click();

        Utils.showToast('Report exported successfully', 'success');
    },

    /**
     * Show/hide loading overlay
     */
    showLoading(show) {
        if (window.Components && typeof Components.setLoading === 'function') {
            Components.setLoading(show, 'Loading report data...');
            return;
        }
        const loader = document.getElementById('loadingOverlay');
        if (loader) {
            if (show) {
                loader.classList.remove('hidden');
            } else {
                loader.classList.add('hidden');
            }
        }
    },

    // ==================== DEBT PROFILE WORKFLOW ====================

    async checkDebtProfileStatus() {
        const result = await API.getDebtProfileRequestStatus();
        if (result.success) {
            this.debtRequestStatus = result;
            this.renderDebtProfileStatus();
        }
    },

    renderDebtProfileStatus() {
        const status = this.debtRequestStatus?.status || 'NONE';
        const user = Auth.getUser();
        const isAdmin = user && [CONFIG.ROLES.ADMIN, CONFIG.ROLES.DFA, CONFIG.ROLES.DDFA].includes(user.role);

        // Hide all
        ['Initial', 'Pending', 'Rejected', 'Approved'].forEach(s => {
            document.getElementById(`debtProfileStatus${s}`)?.classList.add('hidden');
        });

        if (status === 'NONE') {
            document.getElementById('debtProfileStatusInitial')?.classList.remove('hidden');
        } else if (status === 'PENDING') {
            document.getElementById('debtProfileStatusPending')?.classList.remove('hidden');
            document.getElementById('pendingRequester').textContent = this.debtRequestStatus.requester;
            document.getElementById('pendingDate').textContent = Utils.formatDateTime(this.debtRequestStatus.timestamp);
            
            if (isAdmin) {
                document.getElementById('approverActions')?.classList.remove('hidden');
            }
        } else if (status === 'REJECTED') {
            document.getElementById('debtProfileStatusRejected')?.classList.remove('hidden');
            document.getElementById('rejectionReason').textContent = `Reason: ${this.debtRequestStatus.comments || 'No comments'}`;
        } else if (status === 'APPROVED') {
            document.getElementById('debtProfileStatusApproved')?.classList.remove('hidden');
            document.getElementById('approvalDateLabel').textContent = Utils.formatDateTime(this.debtRequestStatus.approvalDate);
        }
    },

    async handleDebtProfileRequest(fresh = false) {
        console.log('handleDebtProfileRequest called', { fresh });
        const modal = document.getElementById('analyticalFormModalOverlay');
        if (modal) {
            // Reset fields if fresh
            const titleInput = document.getElementById('reportTitle');
            const summaryInput = document.getElementById('reportSummary');
            const analysisInput = document.getElementById('reportAnalysis');
            const recommendationsInput = document.getElementById('reportRecommendations');

            if (fresh) {
                if (titleInput) titleInput.value = `Debt Profile Report - ${this.currentYear}`;
                if (summaryInput) summaryInput.value = '';
                if (analysisInput) analysisInput.value = '';
                if (recommendationsInput) recommendationsInput.value = '';
            } else if (this.debtRequestStatus?.narrative) {
                // If we have existing narrative, pre-fill it for editing
                const n = this.debtRequestStatus.narrative;
                if (titleInput) titleInput.value = n.title || `Debt Profile Report - ${this.currentYear}`;
                if (summaryInput) summaryInput.value = n.summary || '';
                if (analysisInput) analysisInput.value = n.analysis || '';
                if (recommendationsInput) recommendationsInput.value = n.recommendations || '';
            } else if (titleInput && !titleInput.value) {
                titleInput.value = `Debt Profile Report - ${this.currentYear}`;
            }

            console.log('Opening modal overlay');
            modal.classList.add('active');
        } else {
            console.error('Modal overlay "analyticalFormModalOverlay" not found');
            Utils.showToast('Report configuration error', 'error');
        }
    },

    /**
     * Automatically generates insights based on current summary data
     */
    generateSmartNarrative() {
        if (!this.summaryData) {
            Utils.showToast('Loading data stats...', 'info');
            return;
        }

        const res = this.summaryData;
        const s = res.data?.summary || {};
        const p = res.performance || {};
        const cm = res.currentMonth || {};
        const cats = res.categoryBreakdown || [];
        const topCat = cats.length > 0 ? cats[0] : null;

        const totalDebt = res.totalDebtProfile || s.totalDebt || 0;
        const efficiency = p.efficiency || '0%';
        const growth = p.growth || '0%';
        const balanceBF = res.balanceBF || 0;
        
        const today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

        // SECTION 1: OVERVIEW
        const summary = `This report provides a comprehensive financial overview of the Federal Medical Centre, Abeokuta, as of ${today}. It serves as a dual-purpose handover document and performance analysis, integrating carryover liabilities from 2025 (${Utils.formatCurrency(balanceBF)}) with new 2026 obligations. The report evaluates contract sums, payment efficiency, and debt growth rates to guide management decision-making.`;
        
        // SECTION 5: ANALYSIS
        const analysis = `Financial evaluation for 2026 reveals a total contract sum commitment of ${Utils.formatCurrency(p.contractSum)}. Payment efficiency currently stands at ${efficiency}, indicating the ratio of vouchers settled against total obligations raised. A growth rate of ${growth} is observed relative to the 2025 starting balance. Sectoral analysis identifies "${topCat ? topCat.name : 'Operations'}" as the lead expenditure category. In the current month (${cm.name}), the hospital processed ${Utils.formatCurrency(cm.newObligations)} in new vouchers while effecting ${Utils.formatCurrency(cm.payments)} in payments.`;
        
        // SECTION 6: RECOMMENDATIONS
        const recommendations = `1. Prioritize the settlement of 2025 Balance B/F (${Utils.formatCurrency(balanceBF)}) to reduce legacy interest/inflation risk.\n2. Implement stricter commitment controls in the "${topCat ? topCat.name : 'High-Expenditure'}" sector.\n3. Target a payment efficiency improvement of 15% in the next quarter via automated voucher processing.\n4. Revalidate all outstanding 2026 vouchers to ensure budgetary alignment before the next fiscal cycle.`;

        document.getElementById('reportTitle').value = `DEBT PROFILE & FINANCIAL HANDOVER REPORT - ${this.currentYear}`;
        document.getElementById('reportSummary').value = summary;
        document.getElementById('reportAnalysis').value = analysis;
        document.getElementById('reportRecommendations').value = recommendations;
        
        Utils.showToast('Comprehensive analytical insights generated!', 'success');
    },

    async submitDebtProfileRequest() {
        const reportData = {
            title: document.getElementById('reportTitle').value,
            summary: document.getElementById('reportSummary').value,
            analysis: document.getElementById('reportAnalysis').value,
            recommendations: document.getElementById('reportRecommendations').value,
            filters: { year: this.currentYear, requestedAt: new Date().toISOString() }
        };

        this.showLoading(true);
        const result = await API.requestDebtProfile(reportData);
        this.showLoading(false);

        if (result.success) {
            document.getElementById('analyticalFormModalOverlay')?.classList.remove('active');
            Utils.showToast(result.message || 'Request processed successfully', 'success');
            await this.checkDebtProfileStatus();
            
            // If it was already approved (re-generation), reload the report
            if (this.debtRequestStatus?.status === 'APPROVED') {
                await this.loadFullDebtReport();
            }
        } else {
            Utils.showToast(result.error || 'Request failed', 'error');
        }
        return result;
    },

    async handleDebtProfileApproval(action) {
        const comments = document.getElementById('approvalComments').value;
        if (action === 'reject' && !comments) {
            Utils.showToast('Comments are required for rejection', 'warning');
            return;
        }

        const confirm = await Utils.confirm(`Are you sure you want to ${action} this request?`, `${action.toUpperCase()} Request`);
        if (!confirm) return;

        this.showLoading(true);
        const result = action === 'approve' 
            ? await API.approveDebtProfile(this.debtRequestStatus.requestId, comments)
            : await API.rejectDebtProfile(this.debtRequestStatus.requestId, comments);
        this.showLoading(false);

        if (result.success) {
            Utils.showToast(`Request ${action}d successfully`, 'success');
            await this.checkDebtProfileStatus();
        } else {
            Utils.showToast(result.error || 'Action failed', 'error');
        }
    },

    async loadFullDebtReport() {
        this.showLoading(true);
        const result = await API.getDebtProfileFullData(this.debtRequestStatus.requestId);
        this.showLoading(false);

        if (result.success) {
            this.debtReportData = result.data;
            this.renderFullDebtReport();
            document.getElementById('fullDebtReportContainer').classList.remove('hidden');
            document.getElementById('debtProfileWorkflowCard').classList.add('hidden');
            Utils.showToast('Report loaded successfully', 'success');
        } else {
            Utils.showToast(result.error || 'Failed to load report data', 'error');
        }
    },

    renderFullDebtReport() {
        const data = this.debtReportData;
        const narrative = this.debtRequestStatus?.narrative || {};
        if (!data) return;

        // Header & Narrative
        document.getElementById('reportHeaderTitle').textContent = narrative.title || 'Official Debt Profile Report';
        document.getElementById('reportGeneratedDate').textContent = Utils.formatDateTime(this.debtRequestStatus.timestamp);
        document.getElementById('reportSummaryText').textContent = narrative.summary || 'No summary provided.';
        document.getElementById('reportAnalysisText').textContent = narrative.analysis || 'No detailed analysis provided.';
        document.getElementById('reportRecommendationsText').textContent = narrative.recommendations || 'No recommendations provided.';

        // Stats
        document.getElementById('reportTotalDebt').textContent = Utils.formatCurrency(data.summary.totalDebt);
        document.getElementById('reportOverdueDebt').textContent = Utils.formatCurrency(data.summary.overdueAmount);
        document.getElementById('reportOverdueCount').textContent = `${data.summary.overdueCount} vouchers`;

        // Automated System Insights (Optional element)
        const reportNarrative = document.getElementById('reportNarrative');
        if (reportNarrative) {
            const overduePct = (data.summary.overdueAmount / (data.summary.totalDebt || 1) * 100).toFixed(1);
            reportNarrative.textContent = `Data analysis shows that ${overduePct}% of the outstanding debt is overdue by more than 90 days. The debt is distributed across ${Object.keys(data.byCategory || {}).length} categories.`;
        }

        // Charts
        this.drawReportCharts(data);

        // Details Table
        const tableContainer = document.getElementById('reportDetailsTableContainer');
        if (data.details && data.details.length > 0) {
            tableContainer.innerHTML = `
                <div class="table-container">
                    <table class="table-hover">
                        <thead>
                            <tr>
                                <th>Payee</th>
                                <th>Particulars</th>
                                <th>Category</th>
                                <th>Dept</th>
                                <th>Date</th>
                                <th class="text-right">Amount</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${data.details.map(d => `
                                <tr>
                                    <td>${d.payee}</td>
                                    <td>${d.particular}</td>
                                    <td><span class="badge badge-info">${d.category}</span></td>
                                    <td>${d.department}</td>
                                    <td>${d.date}</td>
                                    <td class="text-right"><strong>${Utils.formatCurrency(d.amount)}</strong></td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            `;
        } else {
            tableContainer.innerHTML = '<p class="text-center text-muted">No detail records found.</p>';
        }
    },

    drawReportCharts(data) {
        // Aging Chart
        const agingCtx = document.getElementById('reportAgingChart');
        if (agingCtx) {
            if (this.reportAgingChart) this.reportAgingChart.destroy();
            this.reportAgingChart = new Chart(agingCtx, {
                type: 'pie',
                data: {
                    labels: Object.keys(data.byAge),
                    datasets: [{
                        data: Object.values(data.byAge),
                        backgroundColor: ['#28a745', '#ffc107', '#dc3545']
                    }]
                },
                options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
            });
        }

        // Category Chart
        const catCtx = document.getElementById('reportCategoryChart');
        if (catCtx) {
            if (this.reportCategoryChart) this.reportCategoryChart.destroy();
            const topCats = Object.entries(data.byCategory)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 8);
            
            this.reportCategoryChart = new Chart(catCtx, {
                type: 'bar',
                data: {
                    labels: topCats.map(c => c[0]),
                    datasets: [{
                        label: 'Amount',
                        data: topCats.map(c => c[1]),
                        backgroundColor: 'rgba(0, 123, 255, 0.7)'
                    }]
                },
                options: { 
                    indexAxis: 'y', 
                    responsive: true,
                    plugins: { legend: { display: false } },
                    scales: { x: { ticks: { callback: v => 'â‚¦' + Number(v).toLocaleString() } } }
                }
            });
        }

        // Department Chart
        const deptCtx = document.getElementById('reportDepartmentChart');
        if (deptCtx) {
            if (this.reportDepartmentChart) this.reportDepartmentChart.destroy();
            const depts = Object.entries(data.byDepartment).sort((a,b) => b[1] - a[1]);
            this.reportDepartmentChart = new Chart(deptCtx, {
                type: 'bar',
                data: {
                    labels: depts.map(d => String(d[0] || '').replace('::', ' - ').replace(/\s+-\s*$/, '')),
                    datasets: [{
                        label: 'Amount',
                        data: depts.map(d => d[1]),
                        backgroundColor: 'rgba(108, 117, 125, 0.7)'
                    }]
                },
                options: { 
                    responsive: true,
                    plugins: { legend: { display: false } },
                    scales: { y: { ticks: { callback: v => 'â‚¦' + Number(v).toLocaleString() } } }
                }
            });
        }
    },

    async downloadDebtReportGenerated(format) {
        if (!this.debtRequestStatus?.requestId) return;

        // Ensure current form details reflect in the report (Auto-save if edited)
        const currentTitle = document.getElementById('reportTitle')?.value?.trim();
        const currentSummary = document.getElementById('reportSummary')?.value?.trim();
        const currentAnalysis = document.getElementById('reportAnalysis')?.value?.trim();
        const currentRecs = document.getElementById('reportRecommendations')?.value?.trim();
        
        const dbNarrative = this.debtRequestStatus.narrative || {};
        
        const hasChanges = (currentTitle && currentTitle !== (dbNarrative.title || '').trim()) || 
                          (currentSummary && currentSummary !== (dbNarrative.summary || '').trim()) ||
                          (currentAnalysis && currentAnalysis !== (dbNarrative.analysis || '').trim()) ||
                          (currentRecs && currentRecs !== (dbNarrative.recommendations || '').trim());
        
        let freshRequestId = this.debtRequestStatus?.requestId;
        
        if (hasChanges) {
            console.log('Detected narrative changes. Current form vs DB:', {
                current: { title: currentTitle, summary: currentSummary, analysis: currentAnalysis, recs: currentRecs },
                db: dbNarrative
            });
            Utils.showToast('Saving latest narrative edits first...', 'info');
            const syncRes = await this.submitDebtProfileRequest();
            if (syncRes && syncRes.success && syncRes.requestId) {
                freshRequestId = syncRes.requestId;
                console.log('Sync successful. New requestId:', freshRequestId);
            } else {
                console.error('Narrative sync failed or no requestId returned:', syncRes);
                Utils.showToast('Could not sync narrative. Exporting previous version.', 'warning');
            }
        }

        if (!freshRequestId) {
            Utils.showToast('No report request found. Generate a profile first.', 'error');
            return;
        }

        Utils.showToast(`Generating high-quality ${format.toUpperCase()}...`, 'info');
        this.showLoading(true);
        
        try {
            if (format === 'pdf') {
                const result = await API.getDebtProfilePDF(freshRequestId);
                if (result && result.success && result.pdfBase64) {
                    const blob = Utils.base64ToBlob(result.pdfBase64, 'application/pdf');
                    const url = URL.createObjectURL(blob);
                    const link = document.createElement('a');
                    link.href = url;
                    link.download = (currentTitle || 'Debt_Profile_Report').replace(/\s+/g, '_') + '.pdf';
                    link.click();
                    Utils.showToast('PDF downloaded successfully', 'success');
                } else {
                    throw new Error(result?.error || 'Failed to generate PDF');
                }
            } else if (format === 'excel') {
                const result = await API.getDebtProfileExcel(freshRequestId);
                if (result && result.success && result.downloadUrl) {
                    window.open(result.downloadUrl, '_blank');
                    Utils.showToast('Excel report opened in new tab', 'success');
                } else {
                    throw new Error(result?.error || 'Failed to generate Excel');
                }
            }
        } catch (err) {
            Utils.showToast(err.message, 'error');
        } finally {
            this.showLoading(false);
        }
    }
};

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => Reports.init());
