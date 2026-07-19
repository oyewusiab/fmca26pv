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
            this.loadCriticalComplianceActions();
            this.loadApprovalsQueue();
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

            await this.loadCriticalComplianceActions();
            await this.loadApprovalsQueue();

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

    async loadCriticalComplianceActions() {
        const card = document.getElementById('criticalComplianceActionsCard');
        const list = document.getElementById('criticalComplianceActionsList');
        if (!card || !list) return;

        try {
            // Fetch PENDING action items
            const res = await API.getActionItems({ status: 'PENDING' });
            if (res.success && res.items && res.items.length > 0) {
                // Filter and sort by severity (critical first)
                const sorted = res.items.sort((a, b) => {
                    const sevA = String(a.severity || '').toLowerCase();
                    const sevB = String(b.severity || '').toLowerCase();
                    if (sevA === 'danger' && sevB !== 'danger') return -1;
                    if (sevA !== 'danger' && sevB === 'danger') return 1;
                    if (sevA === 'warning' && sevB !== 'warning' && sevB !== 'danger') return -1;
                    if (sevA !== 'warning' && sevB === 'warning' && sevA !== 'danger') return 1;
                    return 0;
                });

                // Pick top 3
                const top3 = sorted.slice(0, 3);
                
                let html = '';
                top3.forEach(item => {
                    const isDanger = String(item.severity || '').toLowerCase() === 'danger';
                    const iconClass = isDanger ? 'fa-exclamation-circle text-danger' : 'fa-exclamation-triangle text-warning';
                    const bgStyle = isDanger ? 'background: #fff8f8; border-color: #fbd5d5;' : 'background: #fffdf5; border-color: #fef3c7;';
                    
                    html += `
                        <div class="mini-compliance-item" style="${bgStyle}">
                            <div class="mini-compliance-text">
                                <i class="fas ${iconClass}" style="margin-right: 8px;"></i>
                                <strong>[${item.voucherNumber || 'General'}]</strong> ${item.message || item.title}
                            </div>
                            <a href="vouchers.html?lookup=true&voucher=${encodeURIComponent(item.voucherNumber)}" class="btn btn-sm btn-secondary" style="margin-left: 10px;">
                                <i class="fas fa-arrow-right"></i> Fix
                            </a>
                        </div>
                    `;
                });
                
                list.innerHTML = html;
                card.style.display = 'block';
            } else {
                card.style.display = 'none';
            }
        } catch (e) {
            console.error('Error loading critical compliance actions:', e);
            card.style.display = 'none';
        }
    },

    async loadApprovalsQueue() {
        const card = document.getElementById('myApprovalsQueueCard');
        const container = document.getElementById('myApprovalsQueue');
        const middleRow = document.getElementById('middleRowGrid');
        if (!card || !container) return;

        // Check if user is an authorized approver
        const user = Auth.getUser();
        const perms = this.permissions || {};
        
        // Approver roles: Payable Head, CPO, Admin
        const approverRoles = [CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.CPO, CONFIG.ROLES.ADMIN];
        const isApprover = user && (approverRoles.includes(user.role) || perms.canApproveDelete);

        if (!isApprover) {
            // Hide approvals queue and make Aged Payables span 100% full-width
            card.style.display = 'none';
            if (middleRow) middleRow.style.gridTemplateColumns = '1fr';
            return;
        }

        // Show approvals queue card and restore 2-column grid
        card.style.display = 'block';
        if (middleRow) middleRow.style.gridTemplateColumns = '1fr 1fr';

        try {
            const res = await API.getPendingDeletions();
            if (res.success && res.vouchers && res.vouchers.length > 0) {
                // Show top 5 recent pending deletions
                const top5 = res.vouchers.slice(0, 5);
                
                let html = '';
                top5.forEach(v => {
                    const rowIdx = v.rowIndex;
                    html += `
                        <div class="approvals-queue-item">
                            <div class="approvals-queue-details">
                                <strong>Voucher:</strong> ${v.accountOrMail || v.voucherNo || '-'}<br>
                                <span class="text-muted" style="font-size: 13px;">
                                    <strong>Payee:</strong> ${Utils.truncate(v.payee, 22)} | 
                                    <strong>Category:</strong> ${v.categories || '-'}
                                </span>
                                <div class="approvals-queue-amount">${Utils.formatCurrency(v.grossAmount || 0)}</div>
                            </div>
                            <div class="approvals-queue-actions">
                                <button class="btn btn-sm btn-success" onclick="Dashboard.approveVoucherDelete(${rowIdx}, '${v.accountOrMail || v.voucherNo || ''}')" title="Approve Deletion" style="background-color: #2ecc71; border-color: #2ecc71; color: white;">
                                    <i class="fas fa-check"></i>
                                </button>
                                <button class="btn btn-sm btn-danger" onclick="Dashboard.rejectVoucherDelete(${rowIdx}, '${v.accountOrMail || v.voucherNo || ''}')" title="Reject Deletion" style="background-color: #e74c3c; border-color: #e74c3c; color: white;">
                                    <i class="fas fa-times"></i>
                                </button>
                            </div>
                        </div>
                    `;
                });
                
                container.innerHTML = html;
            } else {
                container.innerHTML = `
                    <div class="text-center py-4 text-muted">
                        <i class="fas fa-check-double fa-2x mb-2 text-success"></i>
                        <p class="mb-0">No deletion requests awaiting your approval.</p>
                    </div>
                `;
            }
        } catch (e) {
            console.error('Error loading approvals queue:', e);
            container.innerHTML = '<p class="text-danger text-center">Failed to load approvals queue.</p>';
        }
    },

    async approveVoucherDelete(rowIndex, voucherNo) {
        const confirmed = await Utils.confirm(`Approve permanent deletion of voucher ${voucherNo}? This cannot be undone.`, 'Approve Deletion');
        if (!confirmed) return;

        this.showLoading(true);
        try {
            const res = await API.approveDelete(rowIndex);
            if (res.success) {
                Utils.showToast('Voucher successfully deleted', 'success');
                // Reload dashboard data
                await this.loadDashboardData();
            } else {
                Utils.showToast(res.error || 'Failed to approve deletion', 'error');
            }
        } catch (e) {
            console.error('approveVoucherDelete error:', e);
            Utils.showToast('Error approving deletion', 'error');
        }
        this.showLoading(false);
    },

    async rejectVoucherDelete(rowIndex, voucherNo) {
        const reason = prompt('Reason for rejecting deletion request:', 'Rejected from Dashboard');
        if (reason === null) return; // Cancelled prompt

        this.showLoading(true);
        try {
            const res = await API.rejectDelete(rowIndex, reason || 'Rejected from Dashboard');
            if (res.success) {
                Utils.showToast('Deletion request rejected. Voucher restored.', 'success');
                // Reload dashboard data
                await this.loadDashboardData();
            } else {
                Utils.showToast(res.error || 'Failed to reject deletion request', 'error');
            }
        } catch (e) {
            console.error('rejectVoucherDelete error:', e);
            Utils.showToast('Error rejecting deletion request', 'error');
        }
        this.showLoading(false);
    }
};

async function loadDashboardActionItems() {
    try {
        const card = document.getElementById("actionItemsMiniCard");
        const text = document.getElementById("actionItemsMiniText");
        const countBubble = document.getElementById("widgetCount");

        if (!card || !text) return;

        const res = await API.getActionItemCount();

        if (!res.success) {
            card.style.display = "none";
            return;
        }

        const count = res.count || 0;

        if (count > 0) {
            text.textContent = `${count} action item${count === 1 ? '' : 's'}`;
            if (countBubble) {
                countBubble.textContent = count > 99 ? '99+' : String(count);
            }
            card.style.display = "flex";
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
