/**
 * PAYABLE VOUCHER 2026 - Dashboard Module
 * With "eye" icon masking and 2‑minute auto‑mask
 */

const Dashboard = {
    stats: null,
    permissions: null,
    masked: true,
    maskTimeout: null,

    async init() {
        console.log('Dashboard.init() started');

        // Check authentication
        const isAuth = await Auth.requireAuth();
        if (!isAuth) {
            console.log('Auth failed, redirecting...');
            return;
        }

        console.log('Auth passed');

        // Initialize activity tracker
        if (window.ActivityTracker && typeof ActivityTracker.init === 'function') {
            ActivityTracker.init();
        }

        // Get permissions
        this.permissions = await Auth.getPermissions();
        console.log('Permissions loaded:', this.permissions);

        // Setup UI components - ORDER MATTERS!
        this.setupSidebar();
        this.setupUserInfo();
        this.setupEventListeners();
        this.initMasking();

        // Load data
        await this.loadDashboardData();

        // Initialize notification bell after sidebar is ready
        if (window.Notifications && typeof Notifications.initBell === 'function') {
            setTimeout(() => Notifications.initBell(), 100);
        }

        console.log('Dashboard.init() completed');
    },

    // ---------- SIDEBAR ----------

    setupSidebar() {
        console.log('setupSidebar() called');

        const sidebar = document.getElementById('sidebar');
        if (!sidebar) {
            console.error('ERROR: Sidebar element not found!');
            return;
        }

        // Check if Components is available
        if (typeof Components === 'undefined') {
            console.error('ERROR: Components is not defined!');
            return;
        }

        if (typeof Components.getSidebar !== 'function') {
            console.error('ERROR: Components.getSidebar is not a function!');
            return;
        }

        // Check if user is available
        const user = Auth.getUser();
        console.log('User for sidebar:', user);

        if (!user) {
            console.error('ERROR: No user found for sidebar!');
            return;
        }

        // Generate and set sidebar HTML
        const sidebarHtml = Components.getSidebar('dashboard');
        console.log('Sidebar HTML length:', sidebarHtml.length);

        if (sidebarHtml && sidebarHtml.length > 0) {
            sidebar.innerHTML = sidebarHtml;
            console.log('Sidebar HTML set successfully');
        } else {
            console.error('ERROR: getSidebar returned empty string!');
        }
    },

    setupUserInfo() {
        const user = Auth.getUser();
        if (!user) return;

        const avatarEl = document.getElementById('userAvatar');
        const nameEl = document.getElementById('userName');
        const roleEl = document.getElementById('userRole');

        if (avatarEl) avatarEl.textContent = Utils.getInitials(user.name);
        if (nameEl) nameEl.textContent = user.name;
        if (roleEl) roleEl.textContent = user.role;
    },

    // ---------- EVENT LISTENERS ----------

    setupEventListeners() {
        // Logout button
        const logoutBtn = document.getElementById('logoutBtn');
        if (logoutBtn) {
            logoutBtn.addEventListener('click', () => Auth.logout());
        }

        // Refresh button
        const refreshBtn = document.getElementById('refreshBtn');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => this.loadDashboardData());
        }

        // Mobile menu toggle
        const menuToggle = document.getElementById('menuToggle');
        const sidebar = document.getElementById('sidebar');
        if (menuToggle && sidebar) {
            menuToggle.addEventListener('click', () => {
                sidebar.classList.toggle('active');
            });
        }

        // Primary button setup
        this.setupPrimaryButtons();

        // Eye icon toggle for masking
        const maskBtn = document.getElementById('maskToggleBtn');
        if (maskBtn) {
            maskBtn.addEventListener('click', () => this.setMasked(!this.masked));
        }
    },

    setupPrimaryButtons() {
        const perms = this.permissions || {};
        const headerBtn = document.getElementById('primaryVoucherBtn');

        if (headerBtn) {
            if (perms.canCreateVoucher) {
                headerBtn.innerHTML = '<i class="fas fa-plus"></i> New Voucher';
                headerBtn.onclick = () => { window.location.href = 'vouchers.html?new=true'; };
            } else {
                headerBtn.innerHTML = '<i class="fas fa-file-invoice-dollar"></i> View Vouchers';
                headerBtn.onclick = () => { window.location.href = 'vouchers.html'; };
            }
        }
    },

    // ---------- MASKING ----------

    initMasking() {
        this.masked = true;
        this.applyMasking();

        const resetTimer = () => this.resetMaskTimer();

        ['mousemove', 'keydown', 'click', 'scroll', 'touchstart'].forEach(evt => {
            document.addEventListener(evt, resetTimer, { passive: true });
        });

        this.resetMaskTimer();
    },

    resetMaskTimer() {
        if (this.maskTimeout) clearTimeout(this.maskTimeout);
        this.maskTimeout = setTimeout(() => {
            if (!this.masked) {
                this.setMasked(true);
            }
        }, 2 * 60 * 1000);
    },

    setMasked(mask) {
        this.masked = mask;
        this.applyMasking();
        this.resetMaskTimer();
    },

    applyMasking() {
        if (this.stats) {
            this.renderStats(this.stats.stats || this.stats);
            this.renderCategoryBreakdown(this.stats.categoryBreakdown);
            this.renderMonthlyBreakdown(this.stats.monthlyBreakdown);
            this.renderRecentVouchers(this.stats.recentVouchers);
        }

        const btn = document.getElementById('maskToggleBtn');
        if (btn) {
            if (this.masked) {
                btn.innerHTML = '<i class="fas fa-eye"></i>';
                btn.title = 'Show figures';
            } else {
                btn.innerHTML = '<i class="fas fa-eye-slash"></i>';
                btn.title = 'Hide figures';
            }
        }
    },

    // ---------- DATA LOADING ----------

    async loadDashboardData() {
        // 1. Try to load from cache immediately for instant render
        const cached = sessionStorage.getItem('pv2026_dashboard_stats');
        if (cached) {
            this.stats = JSON.parse(cached);
            this.applyMasking();
        }

        this.showLoading(true);

        try {
            const result = await API.getDashboardStats();

            if (!result.success) {
                Utils.showToast(result.error || 'Failed to load dashboard data', 'error');
                this.showLoading(false);
                return;
            }

            this.stats = result;
            sessionStorage.setItem('pv2026_dashboard_stats', JSON.stringify(result));
            this.applyMasking();

        } catch (error) {
            console.error('Dashboard load error:', error);
            Utils.showToast('Error loading dashboard.', 'error');
        }

        this.showLoading(false);
    },

    // ---------- RENDERING ----------

    renderStats(stats) {
        const container = document.getElementById('statsGrid');
        if (!container || !stats) return;

        const m = this.masked;

        const tvCount = Utils.formatNumber(stats.totalVouchersRaised || 0);
        const paidAmt = Utils.formatCurrency(stats.totalPaidAmount || 0);
        const unpaidAmt = Utils.formatCurrency(stats.totalUnpaidAmount || 0);
        const cancAmt = Utils.formatCurrency(stats.totalCancelledAmount || 0);
        const tcs = Utils.formatCurrency(stats.totalProcessedContractSum || 0);
        const debt = Utils.formatCurrency(stats.totalDebt || 0);
        const pRate = `${stats.averagePaymentPercent || 0}%`;
        const revalCnt = Utils.formatNumber(stats.revalidatedVouchers || 0);

        container.innerHTML = `
            <div class="stat-card">
                <div class="stat-label">Total Vouchers Raised</div>
                <div class="stat-value">${m ? '***' : tvCount}</div>
            </div>

            <div class="stat-card paid">
                <div class="stat-label">Paid Vouchers (Amount)</div>
                <div class="stat-value">${m ? '***' : paidAmt}</div>
                <div class="stat-subvalue">
                    ${m ? '***' : Utils.formatNumber(stats.paidVouchers || 0)} voucher(s) paid
                </div>
            </div>

            <div class="stat-card unpaid">
                <div class="stat-label">Unpaid Vouchers (Amount)</div>
                <div class="stat-value">${m ? '***' : unpaidAmt}</div>
                <div class="stat-subvalue">
                    ${m ? '***' : Utils.formatNumber(stats.unpaidVouchers || 0)} voucher(s) unpaid
                </div>
            </div>

            <div class="stat-card cancelled">
                <div class="stat-label">Cancelled Vouchers (Amount)</div>
                <div class="stat-value">${m ? '***' : cancAmt}</div>
                <div class="stat-subvalue">
                    ${m ? '***' : Utils.formatNumber(stats.cancelledVouchers || 0)} voucher(s) cancelled
                </div>
            </div>

            <div class="stat-card info">
                <div class="stat-label">Total Processed Contract Sum</div>
                <div class="stat-value">${m ? '***' : tcs}</div>
            </div>

            <div class="stat-card">
                <div class="stat-label">Total Debt</div>
                <div class="stat-value text-danger">${m ? '***' : debt}</div>
            </div>

            <div class="stat-card info">
                <div class="stat-label">Average Payment Rate</div>
                <div class="stat-value">${m ? '***' : pRate}</div>
            </div>

            <div class="stat-card">
                <div class="stat-label">Revalidated Vouchers</div>
                <div class="stat-value">${m ? '***' : revalCnt}</div>
            </div>
        `;
    },

    renderCategoryBreakdown(categories) {
        const container = document.getElementById('categoryBreakdown');
        if (!container) return;

        if (!categories || categories.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No category data available</p>';
            return;
        }

        const m = this.masked;

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
                        </tr>
                    </thead>
                    <tbody>
        `;

        categories.forEach(cat => {
            html += `
                <tr>
                    <td><strong>${cat.category}</strong></td>
                    <td class="text-center">${m ? '***' : Utils.formatNumber(cat.vouchersRaised)}</td>
                    <td class="text-right text-success">${m ? '***' : Utils.formatCurrency(cat.amountPaid)}</td>
                    <td class="text-right text-danger">${m ? '***' : Utils.formatCurrency(cat.balance)}</td>
                    <td class="text-center">${m ? '***' : `${cat.percentagePaid}%`}</td>
                </tr>
            `;
        });

        html += '</tbody></table></div>';
        container.innerHTML = html;
    },

    renderMonthlyBreakdown(months) {
        const container = document.getElementById('monthlyBreakdown');
        if (!container) return;

        if (!months || months.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No monthly data available</p>';
            return;
        }

        const m = this.masked;

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Month</th>
                            <th class="text-center">Count</th>
                            <th class="text-right">Paid</th>
                            <th class="text-right">Unpaid</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        months.forEach(mon => {
            if (mon.count > 0) {
                html += `
                    <tr>
                        <td><strong>${mon.month}</strong></td>
                        <td class="text-center">${m ? '***' : Utils.formatNumber(mon.count)}</td>
                        <td class="text-right text-success">${m ? '***' : Utils.formatCurrency(mon.paidAmount)}</td>
                        <td class="text-right text-warning">${m ? '***' : Utils.formatCurrency(mon.unpaidAmount)}</td>
                    </tr>
                `;
            }
        });

        html += '</tbody></table></div>';
        container.innerHTML = html;
    },

    renderRecentVouchers(vouchers) {
        const container = document.getElementById('recentVouchers');
        if (!container) return;

        if (!vouchers || vouchers.length === 0) {
            container.innerHTML = '<p class="text-muted text-center">No recent vouchers</p>';
            return;
        }

        const m = this.masked;

        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Voucher No.</th>
                            <th>Payee</th>
                            <th>Amount</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        vouchers.forEach(v => {
            html += `
                <tr onclick="window.location.href='vouchers.html?edit=${v.rowIndex}'" style="cursor:pointer;">
                    <td><strong>${m ? '***' : (v.accountOrMail || '-')}</strong></td>
                    <td>${m ? '***' : Utils.truncate(v.payee, 20)}</td>
                    <td>${m ? '***' : Utils.formatCurrency(v.grossAmount)}</td>
                    <td>${m ? '***' : Utils.getStatusBadge(v.status)}</td>
                </tr>
            `;
        });

        html += '</tbody></table></div>';
        container.innerHTML = html;
    },

    showLoading(show) {
        const overlay = document.getElementById('loadingOverlay');
        if (overlay) {
            overlay.classList.toggle('hidden', !show);
        }
    }
};

async function loadDashboardActionItems() {
    try {
        const res = await API.getActionItemCount();

        if (!res.success) return;

        const card = document.getElementById("actionItemsMiniCard");
        const text = document.getElementById("actionItemsMiniText");

        if (!card || !text) return;

        const count = res.count || 0;

        if (count > 0) {
            text.textContent = `${count} action item${count === 1 ? '' : 's'}`;
            card.style.display = "block";
        } else {
            card.style.display = "none";
        }

    } catch (e) {
        console.error("Failed to load action item count", e);
    }
}

document.addEventListener("DOMContentLoaded", loadDashboardActionItems);

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    console.log('DOMContentLoaded fired for dashboard');
    Dashboard.init();
});