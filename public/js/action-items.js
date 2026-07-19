/**
 * PAYABLE VOUCHER 2026 - Action Items Module
 */

const ActionItems = {
  initialized: false,
  items: [],
  settings: null,
  user: null,

  RULE_DEFS: [
    {
      key: 'PAID_NO_CN',
      label: 'Paid voucher not released',
      description: 'Voucher is marked PAID but has not been released to CPO (no control number).'
    },
    {
      key: 'UNPAID_NO_CN_30D',
      label: 'Unpaid over 30 days',
      description: 'Voucher is still UNPAID with no control number after 30 days.'
    },
    {
      key: 'RELEASED_UNPAID_15D',
      label: 'Released but still unpaid after 15 days',
      description: 'Voucher has control number but remains UNPAID for more than 15 days.'
    }
  ],

  normalize(value) {
    return String(value || '').trim().toUpperCase();
  },

  isManagerRole() {
    const role = this.normalize(this.user?.role);
    return role === 'ADMIN' || role === 'DDFA' || role === 'DFA';
  },

  async init() {
    if (this.initialized) return;
    this.user = Auth.getUser();
    this.bindEvents();
    this.configureRoleView();
    this.initialized = true;
  },

  configureRoleView() {
    const canManage = this.isManagerRole();
    const settingsBtn = document.getElementById('actionItemsSettingsBtn');
    const unitFilter = document.getElementById('actionItemsUnitFilter');
    const statusFilter = document.getElementById('actionItemsStatusFilter');

    if (settingsBtn) settingsBtn.style.display = canManage ? 'inline-flex' : 'none';
    if (unitFilter) unitFilter.style.display = canManage ? 'inline-block' : 'none';
    if (statusFilter) statusFilter.style.display = canManage ? 'inline-block' : 'none';
  },

  bindEvents() {
    document.getElementById('refreshActionItemsBtn')?.addEventListener('click', () => this.refreshFull());
    document.getElementById('actionItemsUnitFilter')?.addEventListener('change', () => this.refreshFull());
    document.getElementById('actionItemsStatusFilter')?.addEventListener('change', () => this.refreshFull());
    document.getElementById('actionItemsSettingsBtn')?.addEventListener('click', () => this.openSettings());
    document.getElementById('closeActionSettingsBtn')?.addEventListener('click', () => this.closeSettings());
    document.getElementById('cancelActionSettingsBtn')?.addEventListener('click', () => this.closeSettings());
    document.getElementById('saveActionSettingsBtn')?.addEventListener('click', () => this.saveSettings());
  },

  buildFilters() {
    const unitFilter = document.getElementById('actionItemsUnitFilter');
    const statusFilter = document.getElementById('actionItemsStatusFilter');

    return {
      unit: unitFilter ? unitFilter.value : 'ALL',
      status: statusFilter ? statusFilter.value : 'PENDING'
    };
  },

  async refreshFull() {
    if (this.isLoading) {
      console.log('Action items refresh already in progress, skipping duplicate call.');
      return;
    }
    this.isLoading = true;
    await this.init();

    const list = document.getElementById('actionItemsList');
    const card = document.getElementById('actionItemsCard');
    if (!list || !card) {
      this.isLoading = false;
      return;
    }

    // ── Phase 1: Instant render from existing log (no spreadsheet scan) ──
    list.innerHTML = `
      <div class="text-center p-3">
        <div class="spinner"></div>
        <p class="text-muted mt-2" style="font-size:13px;">Loading…</p>
      </div>
    `;

    let hasPhase1Data = false;
    try {
      const quickRes = await API.getActionItems({ ...this.buildFilters(), skipSync: true });
      if (quickRes.success && Array.isArray(quickRes.items)) {
        this.items = quickRes.items;
        this.renderList();
        this.updateCounters();
        card.style.display = 'block';
        hasPhase1Data = true;
        // Append a "Syncing…" banner below the list
        const banner = document.createElement('div');
        banner.id = 'actionItemsSyncBanner';
        banner.style.cssText = [
          'display:flex', 'align-items:center', 'gap:8px',
          'padding:9px 16px', 'background:#fffde7',
          'border-top:1px solid #ffe082',
          'font-size:12.5px', 'color:#795548',
          'border-radius:0 0 var(--radius-sm,6px) var(--radius-sm,6px)'
        ].join(';');
        banner.innerHTML = `
          <span style="width:13px;height:13px;border:2px solid #ffe082;border-top-color:#f9a825;border-radius:50%;animation:spin 0.7s linear infinite;display:inline-block;flex-shrink:0;"></span>
          <span>Syncing latest data from spreadsheet — results will refresh automatically…</span>
        `;
        list.parentElement?.appendChild(banner);
      }
    } catch (_) { /* ignore phase-1 failure — phase 2 will handle it */ }

    if (!hasPhase1Data) {
      list.innerHTML = `
        <div class="text-center p-5">
          <div class="spinner"></div>
          <p class="text-muted mt-3">Syncing action items — this may take up to 30 seconds on first load.</p>
        </div>
      `;
    }

    // ── Phase 2: Background sync → then refresh with fresh data ──
    try {
      await API.syncActionItemsOnly();
      const res = await API.getActionItems({ ...this.buildFilters(), skipSync: true });

      // Remove sync banner
      document.getElementById('actionItemsSyncBanner')?.remove();

      if (!res.success) {
        if (!hasPhase1Data) {
          list.innerHTML = `<div class="alert alert-danger">${res.error || 'Failed to load action items.'}</div>`;
        }
        this.updateCounters();
        return;
      }

      this.items = Array.isArray(res.items) ? res.items : [];
      this.renderList();
      this.updateCounters();
      card.style.display = 'block';
    } catch (error) {
      console.error('Action items sync error:', error);
      document.getElementById('actionItemsSyncBanner')?.remove();
      if (!hasPhase1Data) {
        list.innerHTML = `<div class="alert alert-danger">Failed to sync action items. Please try refreshing.</div>`;
        this.items = [];
        this.updateCounters();
      }
      // If phase-1 data exists, keep it displayed — just silently fail the background sync
    } finally {
      this.isLoading = false;
    }
  },

  renderList() {
    const list = document.getElementById('actionItemsList');
    if (!list) return;

    if (!this.items.length) {
      list.innerHTML = `
        <div class="empty-state">
          <i class="fas fa-check-circle"></i>
          <h3>No pending action items</h3>
          <p class="text-muted">You are up to date.</p>
        </div>
      `;
      return;
    }

    list.innerHTML = this.items.map((item) => this.renderRow(item)).join('');
  },

  renderRow(item) {
    const severity = this.normalize(item.severity).toLowerCase() || 'info';
    const iconClass = severity === 'danger'
      ? 'icon-wrap-danger'
      : severity === 'warning'
        ? 'icon-wrap-warning'
        : 'icon-wrap-info';

    const voucher = item.voucherNumber || '-';
    const payee = item.payee || '-';
    const amount = Utils.formatCurrency(item.amount || 0);
    const unit = this.normalize(item.unit) || '-';
    const year = item.year || '2026';
    const voucherStatus = this.normalize(item.status) || '-';
    const itemStatus = this.normalize(item.itemStatus) || '-';
    const accountType = item.accountType || 'Unspecified';
    const pmtMonth = item.pmtMonth || '-';
    const category = item.category || 'Uncategorized';

    const rowIndex = item.rowIndex;

    return `
      <div class="action-item-row ${severity}">
        <div class="action-item-icon-wrap ${iconClass}">
          <i class="fas fa-exclamation-circle"></i>
        </div>
        <div class="action-item-content">
          <div class="action-item-title">${item.title || 'Action required'}</div>
          <div class="action-item-msg">${item.message || ''}</div>
          <div class="action-item-badges">
            <span class="small-badge"><strong>Voucher:</strong> ${voucher}</span>
            <span class="small-badge"><strong>Payee:</strong> ${payee}</span>
            <span class="small-badge"><strong>Amount:</strong> ${amount}</span>
            <span class="small-badge"><strong>Year:</strong> ${year}</span>
            <span class="small-badge"><strong>Unit:</strong> ${unit}</span>
            <span class="small-badge"><strong>Voucher Status:</strong> ${voucherStatus}</span>
            <span class="small-badge"><strong>Item Status:</strong> ${itemStatus}</span>
            <span class="small-badge"><strong>Account Type:</strong> ${accountType}</span>
            <span class="small-badge"><strong>PMT Month:</strong> ${pmtMonth}</span>
            <span class="small-badge"><strong>Category:</strong> ${category}</span>
          </div>
        </div>
        <div class="action-item-actions" style="display: flex; flex-direction: column; gap: 8px;">
          <a class="btn-view" onclick="ActionItems.handleViewAction('${String(voucher).replace(/'/g, "\\'")}')" style="display: inline-flex; align-items: center; justify-content: center; gap: 5px; cursor: pointer; text-decoration: none;">
            <i class="fas fa-eye"></i> View
          </a>
          <a class="btn-resolve" onclick="ActionItems.handleEditAction(${rowIndex})" style="display: inline-flex; align-items: center; justify-content: center; gap: 5px; cursor: pointer; text-decoration: none; background-color: var(--primary-color, #10b981); color: white; padding: 6px 12px; border-radius: 4px; font-weight: 500;">
            <i class="fas fa-tools"></i> Fix Discrepancy
          </a>
        </div>
      </div>
    `;
  },

  updateCounters() {
    const total = this.items.length;
    const critical = this.items.filter((item) => this.normalize(item.severity) === 'DANGER').length;

    const totalEl = document.getElementById('actionItemsCount');
    const criticalEl = document.getElementById('actionItemsCritical');
    if (totalEl) totalEl.textContent = total;
    if (criticalEl) criticalEl.textContent = critical;

    const tabBadge = document.getElementById('tabActionBadge');
    if (tabBadge) {
      tabBadge.textContent = total;
      tabBadge.style.display = total > 0 ? 'inline-block' : 'none';
    }

    const hubAction = document.getElementById('hubActionCount');
    const hubCritical = document.getElementById('hubCriticalCount');
    if (hubAction) hubAction.textContent = total;
    if (hubCritical) hubCritical.textContent = critical;

    const miniCard = document.getElementById('actionItemsMiniCard');
    const miniText = document.getElementById('actionItemsMiniText');
    const miniCount = document.getElementById('widgetCount');
    if (miniCard && miniText) {
      if (total > 0) {
        miniCard.style.display = 'flex';
        miniText.textContent = `${total} action item${total === 1 ? '' : 's'}`;
        if (miniCount) miniCount.textContent = total;
      } else {
        miniCard.style.display = 'none';
      }
    }
  },

  handleViewAction(voucherNumber) {
    const voucher = encodeURIComponent(voucherNumber || '');
    window.location.href = `vouchers.html?lookup=true&voucher=${voucher}`;
  },

  handleEditAction(rowIndex) {
    if (!rowIndex) return;
    window.location.href = `vouchers.html?edit=${rowIndex}`;
  },

  openSettings() {
    if (!this.isManagerRole()) return;
    this.loadSettings();
  },

  closeSettings() {
    document.getElementById('actionItemsSettingsModal')?.classList.remove('active');
  },

  async loadSettings() {
    const body = document.getElementById('actionItemsSettingsBody');
    if (!body) return;

    body.innerHTML = `
      <div class="text-center p-4">
        <div class="spinner"></div>
        <p class="text-muted mt-2">Loading settings...</p>
      </div>
    `;
    document.getElementById('actionItemsSettingsModal')?.classList.add('active');

    try {
      const res = await API.getActionItemSettings();
      if (!res.success) {
        body.innerHTML = `<div class="alert alert-danger">${res.error || 'Failed to load settings.'}</div>`;
        return;
      }

      this.settings = res.settings || { unit: { PAYABLE: {}, CPO: {} } };
      body.innerHTML = this.renderSettingsForm(this.settings);
    } catch (error) {
      console.error('Action settings load error:', error);
      body.innerHTML = `<div class="alert alert-danger">Failed to load settings.</div>`;
    }
  },

  renderSettingsForm(settings) {
    const unitDefs = [
      {
        key: 'PAYABLE',
        label: 'Payable Unit',
        rules: ['PAID_NO_CN', 'MISSING_DATA', 'DUPLICATE_VOUCHER']
      },
      {
        key: 'CPO',
        label: 'CPO Unit',
        rules: ['PAID_NO_CN', 'UNPAID_NO_CN_30D', 'RELEASED_UNPAID_15D']
      },
      {
        key: 'ADMIN',
        label: 'Admin / DDFA / DFA',
        rules: ['MISSING_DATA', 'DUPLICATE_VOUCHER']
      }
    ];

    let html = '<div class="action-settings-grid">';

    unitDefs.forEach(({ key: unit, label, rules: unitRules }) => {
      html += `<div class="card"><div class="card-header"><h4 class="card-title">${label}</h4></div><div class="card-body">`;
      unitRules.forEach(ruleKey => {
        const ruleDef = this.RULE_DEFS.find(r => r.key === ruleKey);
        if (!ruleDef) return;
        const enabled = settings?.unit?.[unit]?.[ruleKey] !== false;
        html += `
          <div class="rule-toggle-row">
            <div>
              <div class="rule-title">${ruleDef.label}</div>
              <div class="rule-desc">${ruleDef.description}</div>
            </div>
            <label class="switch-toggle">
              <input type="checkbox" data-unit="${unit}" data-rule="${ruleKey}" ${enabled ? 'checked' : ''}>
              <span class="switch-slider"></span>
            </label>
          </div>
        `;
      });
      html += '</div></div>';
    });

    html += '</div>';
    return html;
  },

  buildSettingsPayloadFromForm() {
    const payload = { unit: { PAYABLE: {}, CPO: {}, ADMIN: {} } };
    const toggles = document.querySelectorAll('#actionItemsSettingsBody input[type="checkbox"][data-unit][data-rule]');

    toggles.forEach((checkbox) => {
      const unit = checkbox.getAttribute('data-unit');
      const rule = checkbox.getAttribute('data-rule');
      if (!payload.unit[unit]) payload.unit[unit] = {};
      payload.unit[unit][rule] = checkbox.checked;
    });

    return payload;
  },

  async saveSettings() {
    if (!this.isManagerRole()) return;
    const payload = this.buildSettingsPayloadFromForm();

    try {
      const res = await API.saveActionItemSettings(payload);
      if (!res.success) {
        Utils.showToast(res.error || 'Failed to save action item settings.', 'error');
        return;
      }

      Utils.showToast(res.message || 'Action item settings saved.', 'success');
      this.closeSettings();
      await this.refreshFull();
    } catch (error) {
      console.error('Save action settings error:', error);
      Utils.showToast('Failed to save action item settings.', 'error');
    }
  }
};

async function loadActionItemsFullPage() {
  return ActionItems.refreshFull();
}

function handleViewAction(voucherNumber) {
  ActionItems.handleViewAction(voucherNumber);
}

function handleEditAction(rowIndex) {
  ActionItems.handleEditAction(rowIndex);
}

window.ActionItems = ActionItems;
window.loadActionItemsFullPage = loadActionItemsFullPage;
window.handleViewAction = handleViewAction;
window.handleEditAction = handleEditAction;
