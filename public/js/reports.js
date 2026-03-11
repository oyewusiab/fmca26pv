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
    categoryChart: null,
    monthlyChart: null,
    statusCountChart: null,
    statusAmountChart: null,
    debtTrendChart: null,
    topDebtorsChart: null,
    categoryParetoChart: null,
    debtConcentrationChart: null,
    agingChart: null,

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
        await this.loadAllReports();
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
                this.loadYearSummary();
            });
        }

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

        // Listen for SWR background updates
        document.addEventListener('apiDataUpdated', (e) => {
            const { action, data, params } = e.detail;

            // If the user hasn't switched the year dropdown, re-apply the fresh data
            if (action === 'getSummary' && params.year === this.currentYear) {
                this.summaryData = data;
                this.renderYearSummary(data);
                this.renderCategoryTable(data.categoryBreakdown);
                this.renderMonthlyTable(data.monthlyBreakdown);
                this.drawCategoryChart(data.categoryBreakdown);
                this.drawCategoryParetoChart(data.categoryBreakdown);
                this.drawMonthlyChart(data.monthlyBreakdown);
                this.drawStatusCharts(data.summary);
            }
            if (action === 'getAllYearsSummary') {
                this.allYearsData = data;
                this.renderAllYearsSummary(data);
                this.drawDebtTrendChart(data);
            }
            if (action === 'getDebtProfile') {
                this.debtProfile = data;
                this.renderDebtProfile(data);
                this.drawDebtConcentration(data);
                this.drawTopDebtorsChart(data);
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

    /**
     * Load summary for selected year
     */
    async loadYearSummary() {
        const result = await API.getSummary(this.currentYear);

        if (result.success) {
            this.summaryData = result;
            this.renderYearSummary(result);
            this.renderCategoryTable(result.categoryBreakdown);
            this.renderMonthlyTable(result.monthlyBreakdown);
            this.drawCategoryChart(result.categoryBreakdown);
            this.drawCategoryParetoChart(result.categoryBreakdown);
            this.drawMonthlyChart(result.monthlyBreakdown);
            this.drawStatusCharts(result.summary);
        } else {
            Utils.showToast(result.error || 'Failed to load summary', 'error');
        }

        if (this.currentYear === '2026') {
            await this.loadAgingAnalysis2026();
        } else {
            document.getElementById('agingSection')?.classList.add('hidden');
        }
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

    /**
     * Load all years summary for debt tracking
     */
    async loadAllYearsSummary() {
        const result = await API.getAllYearsSummary();

        if (result.success) {
            this.allYearsData = result;
            this.renderAllYearsSummary(result);
            this.drawDebtTrendChart(result);
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
     * - Revalidated: count(OLD VOUCHER NUMBER not empty)
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
                <div class="stat-subvalue">With Old Voucher Number</div>
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
        csv += `Revalidated Vouchers (Count),${s.revalidatedVouchers}\n\n`;
        csv += `Revalidated (No Old Voucher No.),${s.revalidatedWithoutOldNumber || 0}\n\n`;

        // CATEGORY BREAKDOWN
        csv += 'CATEGORY BREAKDOWN\n';
        csv += 'Category,Vouchers Raised,Amount Paid,Balance,% Paid\n';
        this.summaryData.categoryBreakdown.forEach(cat => {
            csv += `${cat.category},${cat.vouchersRaised},${cat.amountPaid},${cat.balance},${cat.percentagePaid}%\n`;
        });
        csv += '\n';

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
        const loader = document.getElementById('loadingOverlay');
        if (loader) {
            if (show) {
                loader.classList.remove('hidden');
            } else {
                loader.classList.add('hidden');
            }
        }
    }
};

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => Reports.init());
