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
        await Components.initPage('dashboard');
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
            this.renderAgedPayables(this.stats.aging);
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
            this.loadPendingActions();
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

            await this.loadPendingActions();
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
        if (window.Components && typeof Components.setLoading === 'function') {
            Components.setLoading(show, 'Loading dashboard...');
            return;
        }
        const overlay = document.getElementById('loadingOverlay');
        if (overlay) {
            overlay.classList.toggle('hidden', !show);
        }
    },

    renderAgedPayables(aging) {
        const container = document.getElementById('agedPayables');
        if (!container) return;

        if (!aging) {
            container.innerHTML = '<p class="text-muted text-center">No aging data available</p>';
            return;
        }

        const m = this.masked;

        const u30Amt = aging.under30 ? aging.under30.amount : 0;
        const u30Cnt = aging.under30 ? aging.under30.count : 0;
        const t60Amt = aging.thirtyToSixty ? aging.thirtyToSixty.amount : 0;
        const t60Cnt = aging.thirtyToSixty ? aging.thirtyToSixty.count : 0;
        const o60Amt = aging.overSixty ? aging.overSixty.amount : 0;
        const o60Cnt = aging.overSixty ? aging.overSixty.count : 0;

        const totalAmt = u30Amt + t60Amt + o60Amt;
        const u30Pct = totalAmt > 0 ? (u30Amt / totalAmt) * 100 : 0;
        const t60Pct = totalAmt > 0 ? (t60Amt / totalAmt) * 100 : 0;
        const o60Pct = totalAmt > 0 ? (o60Amt / totalAmt) * 100 : 0;

        container.innerHTML = `
            <div class="aging-bar-container">
                <div class="aging-bar-label">
                    <span>0 - 30 Days</span>
                    <span>${m ? '***' : Utils.formatCurrency(u30Amt)}</span>
                </div>
                <div class="aging-bar-progress">
                    <div class="aging-bar-fill under-30" style="width: ${u30Pct}%"></div>
                </div>
                <div class="aging-bar-meta">${m ? '***' : u30Cnt} voucher(s) outstanding</div>
            </div>

            <div class="aging-bar-container">
                <div class="aging-bar-label">
                    <span>31 - 60 Days</span>
                    <span>${m ? '***' : Utils.formatCurrency(t60Amt)}</span>
                </div>
                <div class="aging-bar-progress">
                    <div class="aging-bar-fill thirty-to-sixty" style="width: ${t60Pct}%"></div>
                </div>
                <div class="aging-bar-meta">${m ? '***' : t60Cnt} voucher(s) outstanding</div>
            </div>

            <div class="aging-bar-container">
                <div class="aging-bar-label">
                    <span>61+ Days</span>
                    <span>${m ? '***' : Utils.formatCurrency(o60Amt)}</span>
                </div>
                <div class="aging-bar-progress">
                    <div class="aging-bar-fill over-60" style="width: ${o60Pct}%"></div>
                </div>
                <div class="aging-bar-meta">${m ? '***' : o60Cnt} voucher(s) outstanding</div>
            </div>
        `;
    },

    /**
     * Pending Actions — unified widget pulling from 3 lightweight sources in parallel:
     *   1. Action Items count (unit-filtered, no sync triggered)
     *   2. Unread notifications count
     *   3. Pending deletions list (approver roles only)
     *
     * The card auto-hides when all counts reach zero.
     */
    async loadPendingActions() {
        const card = document.getElementById('pendingActionsCard');
        const container = document.getElementById('pendingActionsList');
        const middleRow = document.getElementById('middleRowGrid');
        if (!card || !container) return;

        const user = Auth.getUser();
        const perms = this.permissions || {};
        const role = String(user?.role || '').toUpperCase();

        const isApprover = user &&
            ([CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.CPO, CONFIG.ROLES.ADMIN].includes(user.role) || perms.canApproveDelete);
        const isPayable = [CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.PAYABLE_STAFF].includes(user?.role);
        const isCPO = user?.role === CONFIG.ROLES.CPO;
        const isManager = ['ADMIN', 'DDFA', 'DFA'].includes(role);

        // Fetch all 3 sources concurrently — none triggers a spreadsheet scan
        const [actionRes, notifRes, deletionRes] = await Promise.allSettled([
            API.getActionItemCount({ status: 'PENDING' }),
            API.getNotifications(false),
            isApprover ? API.getPendingDeletions() : Promise.resolve({ success: true, vouchers: [] })
        ]);

        const actionCount = actionRes.status === 'fulfilled' && actionRes.value?.success
            ? (actionRes.value.count || 0) : 0;
        
        const notifList = (notifRes.status === 'fulfilled' && notifRes.value?.success && Array.isArray(notifRes.value.notifications))
            ? notifRes.value.notifications
            : [];

        const unreadMentions = notifList.filter(n => !n.read && (n.type === 'mention' || String(n.title || '').includes('Mentioned') || String(n.message || '').includes('commented')));
        const unreadOthers = notifList.filter(n => !n.read && !(n.type === 'mention' || String(n.title || '').includes('Mentioned') || String(n.message || '').includes('commented')));

        const mentionCount = unreadMentions.length;
        const otherNotifCount = unreadOthers.length;

        const deletionCount = deletionRes.status === 'fulfilled' && deletionRes.value?.success
            ? (deletionRes.value.vouchers?.length || 0) : 0;

        const totalCount = actionCount + mentionCount + otherNotifCount + deletionCount;

        // Update badge
        const badge = document.getElementById('pendingActionsTotalBadge');
        if (badge) badge.textContent = totalCount > 99 ? '99+' : String(totalCount);

        if (totalCount === 0) {
            card.style.display = 'none';
            if (middleRow) middleRow.style.gridTemplateColumns = '1fr';
            return;
        }

        // Show card and restore 2-column grid
        card.style.display = 'block';
        if (middleRow) middleRow.style.gridTemplateColumns = '1fr 1fr';

        let html = '<div class="pending-actions-list">';

        // ── Action Items row ──
        if (actionCount > 0) {
            let label, desc, link;
            if (isPayable) {
                label = 'Payable unit action items';
                desc = `${actionCount} voucher${actionCount !== 1 ? 's' : ''} need attention from Payable unit. Consider reviewing and releasing or correcting records.`;
            } else if (isCPO) {
                label = 'CPO unit action items';
                desc = `${actionCount} voucher${actionCount !== 1 ? 's' : ''} need attention from CPO unit. Consider requesting or updating payment status.`;
            } else if (isManager) {
                label = 'Action items requiring attention';
                desc = `${actionCount} pending item${actionCount !== 1 ? 's' : ''} across units require review.`;
            } else {
                label = 'Pending action items';
                desc = `${actionCount} item${actionCount !== 1 ? 's' : ''} need your attention.`;
            }
            link = 'notifications.html?tab=actionitems';
            html += `
              <div class="pending-action-row warning">
                <div class="pending-action-icon"><i class="fas fa-exclamation-triangle"></i></div>
                <div class="pending-action-content">
                  <div class="pending-action-title">${label}</div>
                  <div class="pending-action-desc">${desc}</div>
                </div>
                <div class="pending-action-count warning">${actionCount > 99 ? '99+' : actionCount}</div>
                <a href="${link}" class="pending-action-link">Take Action <i class="fas fa-arrow-right"></i></a>
              </div>`;
        }

        // ── Mentions & Tagged Comments row ──
        if (mentionCount > 0) {
            html += `
              <div class="pending-action-row info" style="border-left: 4px solid #0284c7;">
                <div class="pending-action-icon" style="color: #0284c7;"><i class="fas fa-comment-dots"></i></div>
                <div class="pending-action-content">
                  <div class="pending-action-title">Mentions & tagged comments</div>
                  <div class="pending-action-desc">${mentionCount} unread comment${mentionCount !== 1 ? 's' : ''} tagging your account.</div>
                </div>
                <div class="pending-action-count info" style="background: #e0f2fe; color: #0369a1;">${mentionCount > 99 ? '99+' : mentionCount}</div>
                <a href="notifications.html" class="pending-action-link">View Discussions <i class="fas fa-arrow-right"></i></a>
              </div>`;
        }

        // ── Other Unread Notifications row ──
        if (otherNotifCount > 0) {
            html += `
              <div class="pending-action-row info">
                <div class="pending-action-icon"><i class="fas fa-bell"></i></div>
                <div class="pending-action-content">
                  <div class="pending-action-title">Unread notifications</div>
                  <div class="pending-action-desc">${otherNotifCount} unread notification${otherNotifCount !== 1 ? 's' : ''} await your review.</div>
                </div>
                <div class="pending-action-count info">${otherNotifCount > 99 ? '99+' : otherNotifCount}</div>
                <a href="notifications.html" class="pending-action-link">View <i class="fas fa-arrow-right"></i></a>
              </div>`;
        }

        // ── Pending Deletions row (approvers only) ──
        if (deletionCount > 0 && isApprover) {
            html += `
              <div class="pending-action-row danger">
                <div class="pending-action-icon"><i class="fas fa-trash-alt"></i></div>
                <div class="pending-action-content">
                  <div class="pending-action-title">Voucher deletion requests</div>
                  <div class="pending-action-desc">${deletionCount} voucher${deletionCount !== 1 ? 's' : ''} awaiting deletion approval.</div>
                </div>
                <div class="pending-action-count danger">${deletionCount > 99 ? '99+' : deletionCount}</div>
                <a href="vouchers.html?filter=pending_delete" class="pending-action-link">Review <i class="fas fa-arrow-right"></i></a>
              </div>`;
        }

        html += '</div>';
        container.innerHTML = html;
    }
};

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    console.log('DOMContentLoaded fired for dashboard');
    Dashboard.init();
    
    // Listen for background API updates (SWR)
    document.addEventListener('apiDataUpdated', (e) => {
        const { action, data } = e.detail;
        if (action === 'getDashboardStats') {
            console.log('Dashboard stats updated in background:', data);
            Dashboard.stats = data;
            sessionStorage.setItem('pv2026_dashboard_stats', JSON.stringify(data));
            Dashboard.applyMasking();
        }
    });
});
        const card = document.getElementById('criticalComplianceActionsCard');
