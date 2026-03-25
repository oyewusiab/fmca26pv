/**
 * PAYABLE VOUCHER 2026 - Tax Management Module
 * Comprehensive tax tracking, analysis and reporting
 */

const Tax = {
    // State
    currentYear: '2026',
    taxSummary: null,
    categoryData: null,
    categoryMeta: null,
    monthlyData: null,
    payeeData: null,
    payments: null,
    schedules: null,
    systemConfig: null,
    categoryExpanded: false,
    categoryDefaultLimit: 10,
    monthlyChart: null,

    /**
     * Initialize tax page
     */
    async init() {
        console.log('=== TAX MODULE INIT ===');

        const isAuth = await Auth.requireAuth();
        if (!isAuth) return;

        const user = Auth.getUser();
        if (!user) {
            Utils.showToast('Session error. Please login again.', 'error');
            setTimeout(() => window.location.replace('login.html'), 1500);
            return;
        }

        const userRole = user.role;
        const roleNorm = String(userRole || '').trim().toLowerCase();
        this.setupSidebar();
        this.setupEventListeners();
        await this.loadSystemConfig();
        await this.loadAllData();
    },

    /**
     * Setup sidebar navigation
     */
    setupSidebar() {
        const sidebar = document.getElementById('sidebar');
        if (sidebar) {
            sidebar.innerHTML = Components.getSidebar('tax');
        }
    },

    /**
     * Setup event listeners
     */
    setupEventListeners() {
        // Logout
        document.getElementById('logoutBtn')?.addEventListener('click', () => Auth.logout());

        // Mobile menu
        document.getElementById('menuToggle')?.addEventListener('click', () => {
            document.getElementById('sidebar')?.classList.toggle('active');
        });

        // Record payment button
        document.getElementById('recordPaymentBtn')?.addEventListener('click', () => {
            this.showPaymentModal();
        });

        // Payment modal
        document.getElementById('closePaymentModal')?.addEventListener('click', () => {
            this.hidePaymentModal();
        });

        document.getElementById('cancelPaymentBtn')?.addEventListener('click', () => {
            this.hidePaymentModal();
        });

        // Payment form submit
        document.getElementById('paymentForm')?.addEventListener('submit', (e) => {
            e.preventDefault();
            this.handlePaymentSubmit(e);
        });

        // Export button
        document.getElementById('exportTaxBtn')?.addEventListener('click', () => {
            this.exportTaxData();
        });

        // Refresh payments
        document.getElementById('refreshPaymentsBtn')?.addEventListener('click', async () => {
            await this.loadPaymentHistory();
        });

        // Payee filters
        document.getElementById('applyPayeeFilters')?.addEventListener('click', async () => {
            const category = document.getElementById('payeeCategoryFilter')?.value;
            const month = document.getElementById('payeeMonthFilter')?.value;
            await this.loadPayeeData({ category, month });
        });

        document.getElementById('clearPayeeFilters')?.addEventListener('click', async () => {
            document.getElementById('payeeCategoryFilter').value = '';
            document.getElementById('payeeMonthFilter').value = '';
            await this.loadPayeeData();
        });

        // Expand categories
        document.getElementById('expandCategoriesBtn')?.addEventListener('click', () => {
            this.toggleCategoryExpansion();
        });

        // Schedule events
        document.getElementById('addScheduleBtn')?.addEventListener('click', () => {
            this.showScheduleModal();
        });

        document.getElementById('refreshSchedulesBtn')?.addEventListener('click', async () => {
            await this.loadSchedules();
        });

        document.getElementById('closeScheduleModal')?.addEventListener('click', () => {
            this.hideScheduleModal();
        });

        document.getElementById('cancelScheduleBtn')?.addEventListener('click', () => {
            this.hideScheduleModal();
        });

        document.getElementById('scheduleForm')?.addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleScheduleSubmit(e);
        });

        // Close modal on outside click
        window.addEventListener('click', (e) => {
            const pModal = document.getElementById('paymentModal');
            const sModal = document.getElementById('scheduleModal');
            if (e.target === pModal) this.hidePaymentModal();
            if (e.target === sModal) this.hideScheduleModal();
        });
    },

    /**
     * Load all tax data
     */
    async loadAllData() {
        this.showLoading(true);
        try {
            await Promise.all([
                this.loadTaxSummary(),
                this.loadCategoryData(),
                this.loadMonthlyData(),
                this.loadPayeeData(),
                this.loadPaymentHistory(),
                this.loadSchedules()
            ]);
        } catch (error) {
            console.error('Error loading tax data:', error);
            Utils.showToast('Error loading tax data', 'error');
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
            console.warn('Tax module system config load failed:', e);
        }
    },

    /**
     * Load tax summary
     */
    async loadTaxSummary() {
        const result = await API.getTaxSummary(this.currentYear);
        if (result.success) {
            this.taxSummary = result.summary;
            this.renderTaxSummary(result.summary);
        } else {
            Utils.showToast(result.error || 'Failed to load tax summary', 'error');
        }
    },


    /**
     * Render tax summary
     */
    renderTaxSummary(summary) {
        // Overall totals
        document.getElementById('totalLiability').textContent =
            Utils.formatCurrency(summary.totalTaxLiability);
        document.getElementById('totalPaid').textContent =
            Utils.formatCurrency(summary.totalPaid);
        document.getElementById('totalOutstanding').textContent =
            Utils.formatCurrency(summary.totalOutstanding);

        const complianceRate = summary.complianceRate || 0;
        document.getElementById('complianceRate').textContent =
            complianceRate.toFixed(1) + '%';

        // Compliance badge
        const badgeEl = document.getElementById('complianceBadge');
        let badgeClass = 'compliance-badge ';
        let badgeText = '';

        if (complianceRate >= 90) {
            badgeClass += 'excellent';
            badgeText = 'Excellent';
        } else if (complianceRate >= 75) {
            badgeClass += 'good';
            badgeText = 'Good';
        } else if (complianceRate >= 50) {
            badgeClass += 'fair';
            badgeText = 'Fair';
        } else {
            badgeClass += 'poor';
            badgeText = 'Needs Attention';
        }

        badgeEl.innerHTML = `<span class="${badgeClass}">${badgeText}</span>`;

        // VAT
        document.getElementById('vatLiability').textContent =
            Utils.formatCurrency(summary.totalVAT);
        document.getElementById('vatPaid').textContent =
            Utils.formatCurrency(summary.paidVAT);
        document.getElementById('vatOutstanding').textContent =
            Utils.formatCurrency(summary.outstandingVAT);

        const vatRate = summary.totalVAT > 0 ? (summary.paidVAT / summary.totalVAT * 100) : 0;
        document.getElementById('vatProgress').style.width = vatRate + '%';
        document.getElementById('vatRate').textContent = vatRate.toFixed(1) + '% compliant';

        // WHT
        document.getElementById('whtLiability').textContent =
            Utils.formatCurrency(summary.totalWHT);
        document.getElementById('whtPaid').textContent =
            Utils.formatCurrency(summary.paidWHT);
        document.getElementById('whtOutstanding').textContent =
            Utils.formatCurrency(summary.outstandingWHT);

        const whtRate = summary.totalWHT > 0 ? (summary.paidWHT / summary.totalWHT * 100) : 0;
        document.getElementById('whtProgress').style.width = whtRate + '%';
        document.getElementById('whtRate').textContent = whtRate.toFixed(1) + '% compliant';

        // Stamp Duty
        document.getElementById('stampLiability').textContent =
            Utils.formatCurrency(summary.totalStampDuty);
        document.getElementById('stampPaid').textContent =
            Utils.formatCurrency(summary.paidStampDuty);
        document.getElementById('stampOutstanding').textContent =
            Utils.formatCurrency(summary.outstandingStampDuty);

        const stampRate = summary.totalStampDuty > 0 ?
            (summary.paidStampDuty / summary.totalStampDuty * 100) : 0;
        document.getElementById('stampProgress').style.width = stampRate + '%';
        document.getElementById('stampRate').textContent = stampRate.toFixed(1) + '% compliant';

        this.renderAccountTypeTaxTable(summary.accountTypeBreakdown || []);
    },

    renderAccountTypeTaxTable(rows) {
        const container = document.getElementById('accountTypeTaxTable');
        if (!container) return;

        const sourceRows = Array.isArray(rows) ? rows : [];
        const configured = this.systemConfig?.accountTypes || {};
        if (!sourceRows.length && !Object.keys(configured).length) {
            container.innerHTML = '<p class="text-muted text-center">No account type tax data available</p>';
            return;
        }

        const grouped = {};
        sourceRows.forEach((r) => {
            const base = String(r.baseAccountType || r.accountType || 'Unspecified').trim() || 'Unspecified';
            const sub = String(r.subAccountType || '').trim();
            if (!grouped[base]) {
                grouped[base] = {
                    count: 0,
                    totalTax: 0,
                    paidTax: 0,
                    outstanding: 0,
                    children: {}
                };
            }

            grouped[base].count += Number(r.count || 0);
            grouped[base].totalTax += Number(r.totalTax || 0);
            grouped[base].paidTax += Number(r.paidTax || 0);
            grouped[base].outstanding += Number(r.outstanding || 0);

            if (sub) {
                if (!grouped[base].children[sub]) {
                    grouped[base].children[sub] = {
                        name: sub, count: 0, totalTax: 0, paidTax: 0, outstanding: 0
                    };
                }
                grouped[base].children[sub].count += Number(r.count || 0);
                grouped[base].children[sub].totalTax += Number(r.totalTax || 0);
                grouped[base].children[sub].paidTax += Number(r.paidTax || 0);
                grouped[base].children[sub].outstanding += Number(r.outstanding || 0);
            }
        });

        Object.keys(configured).forEach((base) => {
            if (!grouped[base]) {
                grouped[base] = {
                    count: 0,
                    totalTax: 0,
                    paidTax: 0,
                    outstanding: 0,
                    children: {}
                };
            }
            (configured[base] || []).forEach((sub) => {
                if (!grouped[base].children[sub]) {
                    grouped[base].children[sub] = {
                        name: sub, count: 0, totalTax: 0, paidTax: 0, outstanding: 0
                    };
                }
            });
        });

        const totals = Object.values(grouped).reduce((acc, r) => {
            acc.count += Number(r.count || 0);
            acc.totalTax += Number(r.totalTax || 0);
            acc.paidTax += Number(r.paidTax || 0);
            acc.outstanding += Number(r.outstanding || 0);
            return acc;
        }, { count: 0, totalTax: 0, paidTax: 0, outstanding: 0 });

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Account Type</th>
                            <th class="text-center">Vouchers</th>
                            <th class="text-right">Total Tax</th>
                            <th class="text-right">Paid Tax</th>
                            <th class="text-right">Outstanding</th>
                            <th class="text-center">Compliance</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        const baseOrder = Object.keys(configured);
        const baseOrderIndex = {};
        baseOrder.forEach((name, idx) => { baseOrderIndex[name] = idx; });

        Object.keys(grouped)
            .sort((a, b) => {
                const ai = Object.prototype.hasOwnProperty.call(baseOrderIndex, a) ? baseOrderIndex[a] : Number.MAX_SAFE_INTEGER;
                const bi = Object.prototype.hasOwnProperty.call(baseOrderIndex, b) ? baseOrderIndex[b] : Number.MAX_SAFE_INTEGER;
                if (ai !== bi) return ai - bi;
                return grouped[b].totalTax - grouped[a].totalTax;
            })
            .forEach((base) => {
            const parent = grouped[base];
            const parentCompliance = parent.totalTax > 0 ? (parent.paidTax / parent.totalTax) * 100 : 0;
            html += `
                <tr>
                    <td><strong>${base}</strong></td>
                    <td class="text-center">${Utils.formatNumber(parent.count || 0)}</td>
                    <td class="text-right">${Utils.formatCurrency(parent.totalTax || 0)}</td>
                    <td class="text-right text-success">${Utils.formatCurrency(parent.paidTax || 0)}</td>
                    <td class="text-right text-danger">${Utils.formatCurrency(parent.outstanding || 0)}</td>
                    <td class="text-center">${Number(parentCompliance || 0).toFixed(1)}%</td>
                </tr>
            `;

            const subOrder = Array.isArray(configured[base]) ? configured[base] : [];
            const subOrderIndex = {};
            subOrder.forEach((name, idx) => { subOrderIndex[name] = idx; });
            Object.values(parent.children || {})
                .sort((a, b) => {
                    const ai = Object.prototype.hasOwnProperty.call(subOrderIndex, a.name) ? subOrderIndex[a.name] : Number.MAX_SAFE_INTEGER;
                    const bi = Object.prototype.hasOwnProperty.call(subOrderIndex, b.name) ? subOrderIndex[b.name] : Number.MAX_SAFE_INTEGER;
                    if (ai !== bi) return ai - bi;
                    return b.totalTax - a.totalTax;
                })
                .forEach((child) => {
                    const childCompliance = child.totalTax > 0 ? (child.paidTax / child.totalTax) * 100 : 0;
                    html += `
                        <tr class="sub-account-row">
                            <td style="padding-left:42px;">${child.name}</td>
                            <td class="text-center">${Utils.formatNumber(child.count || 0)}</td>
                            <td class="text-right">${Utils.formatCurrency(child.totalTax || 0)}</td>
                            <td class="text-right text-success">${Utils.formatCurrency(child.paidTax || 0)}</td>
                            <td class="text-right text-danger">${Utils.formatCurrency(child.outstanding || 0)}</td>
                            <td class="text-center">${Number(childCompliance || 0).toFixed(1)}%</td>
                        </tr>
                    `;
                });
        });

        const compliance = totals.totalTax > 0 ? (totals.paidTax / totals.totalTax) * 100 : 0;
        html += `
            <tr class="totals-row">
                <td><strong>TOTAL</strong></td>
                <td class="text-center"><strong>${Utils.formatNumber(totals.count)}</strong></td>
                <td class="text-right"><strong>${Utils.formatCurrency(totals.totalTax)}</strong></td>
                <td class="text-right text-success"><strong>${Utils.formatCurrency(totals.paidTax)}</strong></td>
                <td class="text-right text-danger"><strong>${Utils.formatCurrency(totals.outstanding)}</strong></td>
                <td class="text-center"><strong>${compliance.toFixed(1)}%</strong></td>
            </tr>
        `;

        html += '</tbody></table></div>';
        container.innerHTML = html;
    },

    /**
     * Load category data
     */
    async loadCategoryData() {
        const result = await API.getTaxByCategory(this.currentYear);

        if (result.success) {
            this.categoryData = result.categories;
            this.categoryMeta = {
                paymentAllocationMode: result.paymentAllocationMode || '',
                totalRemitted: Number(result.totalRemitted || 0)
            };
            this.renderCategoryTable(result.categories);
            this.populateCategoryFilter(result.categories);
        } else {
            Utils.showToast(result.error || 'Failed to load category data', 'error');
        }
    },

    /**
     * Render category table
     */
    renderCategoryTable(categories) {
        const container = document.getElementById('categoryTaxTable');
        const expandBtn = document.getElementById('expandCategoriesBtn');

        if (!categories || categories.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No category data available</p>';
            if (expandBtn) {
                expandBtn.disabled = true;
                expandBtn.style.visibility = 'hidden';
            }
            return;
        }

        const canExpand = categories.length > this.categoryDefaultLimit;
        const visibleRows = (this.categoryExpanded || !canExpand)
            ? categories
            : categories.slice(0, this.categoryDefaultLimit);

        if (expandBtn) {
            expandBtn.disabled = !canExpand;
            expandBtn.style.visibility = canExpand ? 'visible' : 'hidden';
            expandBtn.innerHTML = this.categoryExpanded
                ? '<i class="fas fa-compress-alt"></i> Show Top 10'
                : '<i class="fas fa-expand-alt"></i> View All';
        }

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Category</th>
                            <th class="text-center">Vouchers</th>
                            <th class="text-right">VAT</th>
                            <th class="text-right">WHT</th>
                            <th class="text-right">Stamp Duty</th>
                            <th class="text-right">Total Tax</th>
                            <th class="text-right">Paid</th>
                            <th class="text-right">Outstanding</th>
                            <th class="text-center">Compliance</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        visibleRows.forEach(cat => {
            const complianceClass = cat.complianceRate >= 75 ? 'success' :
                cat.complianceRate >= 50 ? 'warning' : 'danger';

            html += `
                <tr class="category-row" data-category="${cat.category}">
                    <td><strong>${cat.category}</strong></td>
                    <td class="text-center">${Utils.formatNumber(cat.count)}</td>
                    <td class="text-right">${Utils.formatCurrency(cat.totalVAT)}</td>
                    <td class="text-right">${Utils.formatCurrency(cat.totalWHT)}</td>
                    <td class="text-right">${Utils.formatCurrency(cat.totalStampDuty)}</td>
                    <td class="text-right"><strong>${Utils.formatCurrency(cat.totalTax)}</strong></td>
                    <td class="text-right text-success">${Utils.formatCurrency(cat.paidTax)}</td>
                    <td class="text-right text-danger">${Utils.formatCurrency(cat.outstanding)}</td>
                    <td class="text-center">
                        <div class="progress-bar-mini">
                            <div class="progress-fill ${complianceClass}" 
                                 style="width: ${cat.complianceRate}%"></div>
                        </div>
                        <small>${cat.complianceRate.toFixed(1)}%</small>
                    </td>
                </tr>
            `;
        });

        // Totals row
        const totals = categories.reduce((acc, cat) => ({
            count: acc.count + cat.count,
            totalVAT: acc.totalVAT + cat.totalVAT,
            totalWHT: acc.totalWHT + cat.totalWHT,
            totalStampDuty: acc.totalStampDuty + cat.totalStampDuty,
            totalTax: acc.totalTax + cat.totalTax,
            paidTax: acc.paidTax + cat.paidTax,
            outstanding: acc.outstanding + cat.outstanding
        }), {
            count: 0, totalVAT: 0, totalWHT: 0, totalStampDuty: 0,
            totalTax: 0, paidTax: 0, outstanding: 0
        });

        const overallCompliance = totals.totalTax > 0 ?
            (totals.paidTax / totals.totalTax * 100) : 0;

        html += `
                        <tr class="totals-row">
                            <td><strong>TOTALS</strong></td>
                            <td class="text-center"><strong>${Utils.formatNumber(totals.count)}</strong></td>
                            <td class="text-right"><strong>${Utils.formatCurrency(totals.totalVAT)}</strong></td>
                            <td class="text-right"><strong>${Utils.formatCurrency(totals.totalWHT)}</strong></td>
                            <td class="text-right"><strong>${Utils.formatCurrency(totals.totalStampDuty)}</strong></td>
                            <td class="text-right"><strong>${Utils.formatCurrency(totals.totalTax)}</strong></td>
                            <td class="text-right"><strong>${Utils.formatCurrency(totals.paidTax)}</strong></td>
                            <td class="text-right"><strong>${Utils.formatCurrency(totals.outstanding)}</strong></td>
                            <td class="text-center"><strong>${overallCompliance.toFixed(1)}%</strong></td>
                        </tr>
                    </tbody>
                </table>
            </div>
        `;

        if (canExpand && !this.categoryExpanded) {
            html += `
                <p class="text-muted" style="margin-top: 10px; font-size: 12px;">
                    Showing top ${this.categoryDefaultLimit} of ${categories.length} categories
                </p>
            `;
        }

        if (this.categoryMeta && this.categoryMeta.paymentAllocationMode === 'UNMAPPED') {
            html += `
                <p class="text-muted" style="margin-top: 6px; font-size: 12px;">
                    Paid/outstanding by category is unmapped unless remittance entries specify category allocation.
                </p>
            `;
        }

        container.innerHTML = html;
    },

    /**
     * Load monthly data
     */
    async loadMonthlyData() {
        const result = await API.getTaxByMonth(this.currentYear);

        if (result.success) {
            this.monthlyData = result.months;
            this.renderMonthlyChart(result.months);
            this.renderMonthlyTable(result.months);
            this.populateMonthFilter(result.months);
        } else {
            Utils.showToast(result.error || 'Failed to load monthly data', 'error');
        }
    },

    /**
     * Render monthly chart
     */
    renderMonthlyChart(months) {
        const ctx = document.getElementById('monthlyTaxChart');
        if (!ctx || !months) return;

        const labels = months.map(m => m.month);
        const vatData = months.map(m => m.totalVAT);
        const whtData = months.map(m => m.totalWHT);
        const stampData = months.map(m => m.totalStampDuty);
        const paidData = months.map(m => m.paidTax);

        if (this.monthlyChart) {
            this.monthlyChart.destroy();
        }

        this.monthlyChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'VAT',
                        data: vatData,
                        backgroundColor: 'rgba(40, 167, 69, 0.7)',
                        stack: 'liability'
                    },
                    {
                        label: 'WHT',
                        data: whtData,
                        backgroundColor: 'rgba(255, 193, 7, 0.7)',
                        stack: 'liability'
                    },
                    {
                        label: 'Stamp Duty',
                        data: stampData,
                        backgroundColor: 'rgba(220, 53, 69, 0.7)',
                        stack: 'liability'
                    },
                    {
                        label: 'Total Paid',
                        data: paidData,
                        backgroundColor: 'rgba(0, 123, 255, 0.5)',
                        type: 'line',
                        borderColor: 'rgba(0, 123, 255, 1)',
                        borderWidth: 2,
                        fill: false
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
                        stacked: true,
                        ticks: {
                            callback: function (value) {
                                return '₦' + Number(value).toLocaleString();
                            }
                        }
                    },
                    x: {
                        stacked: true
                    }
                }
            }
        });
    },

    /**
     * Render monthly table
     */
    renderMonthlyTable(months) {
        const container = document.getElementById('monthlyTaxTable');

        if (!months || months.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No monthly data available</p>';
            return;
        }

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Month</th>
                            <th class="text-center">Vouchers</th>
                            <th class="text-right">VAT</th>
                            <th class="text-right">WHT</th>
                            <th class="text-right">Stamp Duty</th>
                            <th class="text-right">Total Tax</th>
                            <th class="text-right">Paid</th>
                            <th class="text-right">Outstanding</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        months.forEach(month => {
            html += `
                <tr>
                    <td><strong>${month.month}</strong></td>
                    <td class="text-center">${Utils.formatNumber(month.count)}</td>
                    <td class="text-right">${Utils.formatCurrency(month.totalVAT)}</td>
                    <td class="text-right">${Utils.formatCurrency(month.totalWHT)}</td>
                    <td class="text-right">${Utils.formatCurrency(month.totalStampDuty)}</td>
                    <td class="text-right"><strong>${Utils.formatCurrency(month.totalTax)}</strong></td>
                    <td class="text-right text-success">${Utils.formatCurrency(month.paidTax)}</td>
                    <td class="text-right text-danger">${Utils.formatCurrency(month.outstanding)}</td>
                </tr>
            `;
        });

        html += `
                    </tbody>
                </table>
            </div>
        `;

        container.innerHTML = html;
    },

    /**
     * Load payee data
     */
    async loadPayeeData(filters = {}) {
        const result = await API.getTaxByPayee(this.currentYear, filters);

        if (result.success) {
            this.payeeData = result.payees;
            this.renderPayeeTable(result.payees);
        } else {
            Utils.showToast(result.error || 'Failed to load payee data', 'error');
        }
    },

    /**
     * Render payee table
     */
    renderPayeeTable(payees) {
        const container = document.getElementById('payeeTaxTable');

        if (!payees || payees.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No payee data available</p>';
            return;
        }

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Payee</th>
                            <th class="text-center">Vouchers</th>
                            <th class="text-right">VAT</th>
                            <th class="text-right">WHT</th>
                            <th class="text-right">Stamp Duty</th>
                            <th class="text-right">Total Tax</th>
                            <th class="text-right">Paid</th>
                            <th class="text-right">Outstanding</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        // Show top 50 payees
        const displayPayees = payees.slice(0, 50);

        displayPayees.forEach(payee => {
            html += `
                <tr>
                    <td>${Utils.truncate(payee.payee, 40)}</td>
                    <td class="text-center">${Utils.formatNumber(payee.count)}</td>
                    <td class="text-right">${Utils.formatCurrency(payee.totalVAT)}</td>
                    <td class="text-right">${Utils.formatCurrency(payee.totalWHT)}</td>
                    <td class="text-right">${Utils.formatCurrency(payee.totalStampDuty)}</td>
                    <td class="text-right"><strong>${Utils.formatCurrency(payee.totalTax)}</strong></td>
                    <td class="text-right text-success">${Utils.formatCurrency(payee.paidTax)}</td>
                    <td class="text-right text-danger">${Utils.formatCurrency(payee.outstanding)}</td>
                </tr>
            `;
        });

        html += `
                    </tbody>
                </table>
            </div>
        `;

        if (payees.length > 50) {
            html += `<p class="text-muted" style="margin-top: 10px; font-size: 12px;">
                Showing top 50 of ${payees.length} payees
            </p>`;
        }

        if (payees[0] && payees[0].paymentAllocation === 'UNMAPPED') {
            html += `
                <p class="text-muted" style="margin-top: 6px; font-size: 12px;">
                    Paid/outstanding by payee is unmapped unless remittance entries specify payee allocation.
                </p>
            `;
        }

        container.innerHTML = html;
    },

    /**
     * Load payment history
     */
    async loadPaymentHistory() {
        const result = await API.getTaxPayments(this.currentYear);

        if (result.success) {
            this.payments = result.payments;
            this.renderPaymentHistory(result.payments);
        } else {
            Utils.showToast(result.error || 'Failed to load payment history', 'error');
        }
    },

    /**
     * Render payment history
     */
    renderPaymentHistory(payments) {
        const container = document.getElementById('paymentsTable');

        if (!payments || payments.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No payment records found</p>';
            return;
        }

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Date</th>
                            <th>Tax Type</th>
                            <th>Period</th>
                            <th class="text-right">Amount</th>
                            <th>Payment Method</th>
                            <th>Reference</th>
                            <th>Bank</th>
                            <th>Created By</th>
                            <th>Notes</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        payments.forEach(payment => {
            const date = payment.date instanceof Date ? payment.date : new Date(payment.date);
            const formattedDate = date.toLocaleDateString('en-GB');

            html += `
                <tr>
                    <td>${formattedDate}</td>
                    <td><span class="badge badge-info">${payment.taxType}</span></td>
                    <td>${payment.period}</td>
                    <td class="text-right"><strong>${Utils.formatCurrency(payment.amount)}</strong></td>
                    <td>${payment.paymentMethod}</td>
                    <td>${payment.referenceNumber || '-'}</td>
                    <td>${payment.bank || '-'}</td>
                    <td>${payment.createdBy}</td>
                    <td>${Utils.truncate(payment.notes || '-', 30)}</td>
                </tr>
            `;
        });

        // Totals by tax type
        const totals = payments.reduce((acc, p) => {
            acc[p.taxType] = (acc[p.taxType] || 0) + Number(p.amount);
            acc.total += Number(p.amount);
            return acc;
        }, { total: 0 });

        html += `
                        <tr class="totals-row">
                            <td colspan="3"><strong>TOTALS</strong></td>
                            <td class="text-right"><strong>${Utils.formatCurrency(totals.total)}</strong></td>
                            <td colspan="5">
                                ${totals.VAT ? `VAT: ${Utils.formatCurrency(totals.VAT)} | ` : ''}
                                ${totals.WHT ? `WHT: ${Utils.formatCurrency(totals.WHT)} | ` : ''}
                                ${totals['Stamp Duty'] ? `Stamp Duty: ${Utils.formatCurrency(totals['Stamp Duty'])} | ` : ''}
                                ${totals['All tax type'] ? `All tax type: ${Utils.formatCurrency(totals['All tax type'])}` : ''}
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        `;

        container.innerHTML = html;
    },

    /**
     * Show payment modal
     */
    showPaymentModal() {
        const modal = document.getElementById('paymentModal');
        const form = document.getElementById('paymentForm');

        // Reset form
        form.reset();

        // Set default date to today
        const today = new Date().toISOString().split('T')[0];
        form.querySelector('[name="date"]').value = today;

        modal.classList.add('active');
    },

    /**
     * Hide payment modal
     */
    hidePaymentModal() {
        const modal = document.getElementById('paymentModal');
        modal.classList.remove('active');
    },

    /**
     * Handle payment form submit
     */
    async handlePaymentSubmit(e) {
        e.preventDefault();

        const form = e.target;
        const formData = new FormData(form);

        const payment = {
            taxType: formData.get('taxType'),
            date: formData.get('date'),
            period: formData.get('period'),
            year: this.currentYear,
            amount: formData.get('amount'),
            paymentMethod: formData.get('paymentMethod'),
            referenceNumber: formData.get('referenceNumber'),
            bank: formData.get('bank'),
            notes: formData.get('notes')
        };

        // Validate
        if (!payment.taxType || !payment.date || !payment.period || !payment.amount || !payment.paymentMethod) {
            Utils.showToast('Please fill in all required fields', 'warning');
            return;
        }

        this.showLoading(true);

        const result = await API.recordTaxPayment(payment);

        this.showLoading(false);

        if (result.success) {
            Utils.showToast('Tax payment recorded successfully', 'success');
            this.hidePaymentModal();

            // Reload data
            await this.loadPaymentHistory();
            await this.loadTaxSummary();
        } else {
            Utils.showToast(result.error || 'Failed to record payment', 'error');
        }
    },

    /**
     * Export tax data to CSV
     */
    exportTaxData() {
        if (!this.categoryData || !this.monthlyData) {
            Utils.showToast('No data to export', 'warning');
            return;
        }

        let csv = 'TAX REPORT - ' + this.currentYear + '\n\n';

        // Summary
        if (this.taxSummary) {
            csv += 'SUMMARY\n';
            csv += 'Total Tax Liability,' + this.taxSummary.totalTaxLiability + '\n';
            csv += 'Total Paid,' + this.taxSummary.totalPaid + '\n';
            csv += 'Total Outstanding,' + this.taxSummary.totalOutstanding + '\n';
            csv += 'Compliance Rate,' + this.taxSummary.complianceRate.toFixed(2) + '%\n\n';

            if (Array.isArray(this.taxSummary.accountTypeBreakdown) && this.taxSummary.accountTypeBreakdown.length) {
                csv += 'ACCOUNT TYPE TAX BREAKDOWN\n';
                csv += 'Account Type,Vouchers,Total Tax,Paid Tax,Outstanding,Compliance %\n';
                this.taxSummary.accountTypeBreakdown.forEach(row => {
                    csv += `${row.accountType || 'Unspecified'},${row.count || 0},${row.totalTax || 0},${row.paidTax || 0},${row.outstanding || 0},${Number(row.complianceRate || 0).toFixed(2)}\n`;
                });
                csv += '\n';
            }
        }

        // Category breakdown
        csv += 'CATEGORY BREAKDOWN\n';
        csv += 'Category,Vouchers,VAT,WHT,Stamp Duty,Total Tax,Paid,Outstanding,Compliance %\n';

        this.categoryData.forEach(cat => {
            csv += `${cat.category},${cat.count},${cat.totalVAT},${cat.totalWHT},${cat.totalStampDuty},${cat.totalTax},${cat.paidTax},${cat.outstanding},${cat.complianceRate.toFixed(2)}\n`;
        });

        csv += '\n';

        // Monthly breakdown
        csv += 'MONTHLY BREAKDOWN\n';
        csv += 'Month,Vouchers,VAT,WHT,Stamp Duty,Total Tax,Paid,Outstanding\n';

        this.monthlyData.forEach(month => {
            csv += `${month.month},${month.count},${month.totalVAT},${month.totalWHT},${month.totalStampDuty},${month.totalTax},${month.paidTax},${month.outstanding}\n`;
        });

        // Download
        const blob = new Blob([csv], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Tax_Report_${this.currentYear}_${new Date().toISOString().slice(0, 10)}.csv`;
        a.click();
        window.URL.revokeObjectURL(url);

        Utils.showToast('Tax report exported successfully', 'success');
    },

    /**
     * Populate category filter dropdown
     */
    populateCategoryFilter(categories) {
        const select = document.getElementById('payeeCategoryFilter');
        if (!select) return;

        const currentValue = select.value;
        select.innerHTML = '<option value="">All Categories</option>';

        categories.forEach(cat => {
            const option = document.createElement('option');
            option.value = cat.category;
            option.textContent = cat.category;
            select.appendChild(option);
        });

        if (currentValue) {
            select.value = currentValue;
        }
    },

    /**
     * Populate month filter dropdown
     */
    populateMonthFilter(months) {
        const select = document.getElementById('payeeMonthFilter');
        if (!select) return;

        const currentValue = select.value;
        select.innerHTML = '<option value="">All Months</option>';

        const monthsList = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ];

        monthsList.forEach(month => {
            const option = document.createElement('option');
            option.value = month;
            option.textContent = month;
            select.appendChild(option);
        });

        if (currentValue) {
            select.value = currentValue;
        }
    },

    /**
     * Toggle category expansion
     */
    toggleCategoryExpansion() {
        if (!Array.isArray(this.categoryData) || this.categoryData.length === 0) return;
        if (this.categoryData.length <= this.categoryDefaultLimit) {
            this.renderCategoryTable(this.categoryData);
            return;
        }

        this.categoryExpanded = !this.categoryExpanded;
        this.renderCategoryTable(this.categoryData);
    },

    /**
     * Load tax schedules
     */
    async loadSchedules() {
        const result = await API.getTaxSchedule(this.currentYear);
        if (result.success) {
            this.schedules = result.schedules;
            this.renderSchedules(result.schedules);
        } else {
            Utils.showToast(result.error || 'Failed to load schedules', 'error');
        }
    },

    /**
     * Render schedules table
     */
    renderSchedules(schedules) {
        const container = document.getElementById('schedulesTable');
        if (!schedules || schedules.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No tax schedules planned yet.</p>';
            return;
        }

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Month</th>
                            <th>Tax Type</th>
                            <th>Due Date</th>
                            <th class="text-right">Planned Amount</th>
                            <th>Status</th>
                            <th>Creator</th>
                            <th>Notes</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        schedules.forEach(s => {
            const statusClass = s.status.toLowerCase().replace(' ', '-');
            const dueDate = s.dueDate ? new Date(s.dueDate).toLocaleDateString('en-GB') : '-';
            
            html += `
                <tr>
                    <td><strong>${s.month}</strong></td>
                    <td><span class="badge badge-info">${s.taxType}</span></td>
                    <td>${dueDate}</td>
                    <td class="text-right"><strong>${Utils.formatCurrency(s.amount)}</strong></td>
                    <td><span class="badge badge-${statusClass}">${s.status}</span></td>
                    <td>${s.createdBy || '-'}</td>
                    <td>${Utils.truncate(s.notes || '-', 25)}</td>
                    <td>
                        <button class="btn btn-icon text-primary" onclick="Tax.editSchedule('${s.id}')" title="Edit">
                            <i class="fas fa-edit"></i>
                        </button>
                    </td>
                </tr>
            `;
        });

        html += `</tbody></table></div>`;
        container.innerHTML = html;
    },

    /**
     * Show schedule modal
     */
    showScheduleModal(scheduleId = null) {
        const modal = document.getElementById('scheduleModal');
        const form = document.getElementById('scheduleForm');
        const title = document.getElementById('scheduleModalTitle');
        
        form.reset();
        document.getElementById('scheduleIdField').value = '';
        
        if (scheduleId) {
            const s = this.schedules.find(x => x.id === scheduleId);
            if (s) {
                title.textContent = 'Edit Tax Schedule';
                document.getElementById('scheduleIdField').value = s.id;
                form.querySelector('[name="month"]').value = s.month;
                form.querySelector('[name="dueDate"]').value = s.dueDate ? s.dueDate.split('T')[0] : '';
                form.querySelector('[name="taxType"]').value = s.taxType;
                form.querySelector('[name="amount"]').value = s.amount;
                form.querySelector('[name="status"]').value = s.status;
                form.querySelector('[name="notes"]').value = s.notes || '';
            }
        } else {
            title.textContent = 'Add Tax Schedule';
            const nextMonth = new Date();
            nextMonth.setMonth(nextMonth.getMonth() + 1);
            form.querySelector('[name="dueDate"]').value = nextMonth.toISOString().split('T')[0];
        }

        modal.classList.add('active');
    },

    hideScheduleModal() {
        document.getElementById('scheduleModal').classList.remove('active');
    },

    async handleScheduleSubmit(e) {
        const form = e.target;
        const formData = new FormData(form);
        const scheduleId = formData.get('id');
        
        const scheduleData = {
            month: formData.get('month'),
            year: this.currentYear, // Ensure year is passed
            dueDate: formData.get('dueDate'),
            taxType: formData.get('taxType'),
            amount: formData.get('amount'),
            status: formData.get('status'),
            notes: formData.get('notes')
        };

        this.showLoading(true);
        let result;
        
        if (scheduleId) {
            result = await API.updateTaxSchedule(scheduleId, scheduleData);
        } else {
            result = await API.createTaxSchedule(scheduleData);
        }

        this.showLoading(false);
        
        if (result.success) {
            Utils.showToast(result.message || 'Schedule saved successfully', 'success');
            this.hideScheduleModal();
            await this.loadSchedules();
        } else {
            Utils.showToast(result.error || 'Failed to save schedule', 'error');
        }
    },

    editSchedule(id) {
        this.showScheduleModal(id);
    },

    /**
     * Show/hide loading overlay
     */
    showLoading(show) {
        if (window.Components && typeof Components.setLoading === 'function') {
            Components.setLoading(show, 'Loading tax data...');
            return;
        }
        const overlay = document.getElementById('loadingOverlay');
        if (overlay) {
            if (show) overlay.classList.remove('hidden');
            else overlay.classList.add('hidden');
        }
    }
};

// Global access for onclick handlers
window.Tax = Tax;

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => Tax.init());
} else {
    Tax.init();
}
