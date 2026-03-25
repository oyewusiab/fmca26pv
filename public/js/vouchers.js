/**
 * PAYABLE VOUCHER 2026 - 
 */
//
if (typeof Utils !== 'undefined' && !Utils.formatNumber) {
  Utils.formatNumber = function (num) {
    return new Intl.NumberFormat('en-NG', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(num || 0);
  };
}

const Vouchers = {
  // ===== State =====
  vouchers: [],
  currentPage: 1,
  pageSize: 50,
  totalCount: 0,
  totalPages: 0,

  selectedVoucher: null,
  selectedVouchers: [],

  categoriesLoaded: false,
  systemConfig: null,
  systemConfigLoaded: false,
  paymentTypeHandlersReady: false,
  amountFormattingReady: false,

  permissions: null,
  isEditMode: false,

  filters: {
    status: 'All',
    category: 'All',
    searchTerm: '',
    pmtMonth: 'All',
    dateFrom: '',
    dateTo: '',
    amountMin: '',
    amountMax: '',
    release: 'All'
  },

  // Global Search
  isGlobalSearchMode: false,
  globalSearchYears: ['2026', '2025', '2024', '2023', '<2023'],

  // Pending deletions
  pendingDeletionsLoaded: false,
  pendingDeletions: [],

  // Delete workflow targets
  deleteTargetVoucher: null,
  rejectTargetVoucher: null,

  // Enhanced release workflow
  isPayableUnit: false,
  isCPO: false,
  releaseSearchResults: [],
  selectedForRelease: [],
  _loadSeq: 0,
  _searchSeq: 0,
  _releaseSearchSeq: 0,
  _lastFiltersKey: '',

  // ===== Init =====
  async init() {
    const isAuth = await Auth.requireAuth();
    if (!isAuth) return;

    const rawPermissions = await Auth.getPermissions();
    this.permissions = this.normalizePermissions(rawPermissions);

    this.setupUI();

    const preloadTasks = [];
    if (!this.categoriesLoaded) {
      preloadTasks.push(
        this.loadCategories().finally(() => { this.categoriesLoaded = true; })
      );
    }

    if (!this.systemConfigLoaded) {
      preloadTasks.push(
        this.loadSystemConfig().finally(() => { this.systemConfigLoaded = true; })
      );
    }

    await Promise.all(preloadTasks);
    await this.loadVouchers();
    this.loadPendingDeletions();

    this.setupEventListeners();
    this.handleUrlParams();
    this.setupUppercaseInputs();
    this.bindAmountFormattingHandlers();
  },

  normalizePermissions(rawPermissions) {
    const user = Auth.getUser() || {};
    const role = user.role || '';

    let perms = rawPermissions || {};

    // Support wrapped API response
    if (perms && perms.success && perms.permissions) {
      perms = perms.permissions;
    }

    const fallback = this.getRoleBasedPermissions(role);

    return {
      ...fallback,
      ...(perms || {})
    };
  },

  getRoleBasedPermissions(role) {
    return {
      canCreateVoucher: [
        CONFIG.ROLES.PAYABLE_STAFF,
        CONFIG.ROLES.PAYABLE_HEAD,
        CONFIG.ROLES.ADMIN
      ].includes(role),

      canLookup: [
        CONFIG.ROLES.PAYABLE_STAFF,
        CONFIG.ROLES.PAYABLE_HEAD,
        CONFIG.ROLES.CPO,
        CONFIG.ROLES.ADMIN
      ].includes(role),

      canEditVoucher: [
        CONFIG.ROLES.PAYABLE_STAFF,
        CONFIG.ROLES.PAYABLE_HEAD,
        CONFIG.ROLES.ADMIN
      ].includes(role),

      canUpdateStatus: [
        CONFIG.ROLES.PAYABLE_STAFF,
        CONFIG.ROLES.PAYABLE_HEAD,
        CONFIG.ROLES.CPO,
        CONFIG.ROLES.ADMIN
      ].includes(role),

      canRequestDelete: [
        CONFIG.ROLES.PAYABLE_STAFF,
        CONFIG.ROLES.PAYABLE_HEAD,
        CONFIG.ROLES.CPO,
        CONFIG.ROLES.AUDIT,
        CONFIG.ROLES.ADMIN
      ].includes(role),

      canApproveDelete: [
        CONFIG.ROLES.PAYABLE_HEAD,
        CONFIG.ROLES.CPO,
        CONFIG.ROLES.ADMIN
      ].includes(role),

      canBatchUpdateStatus: [
        CONFIG.ROLES.CPO,
        CONFIG.ROLES.ADMIN
      ].includes(role),

      canReleaseToUnit: [
        CONFIG.ROLES.PAYABLE_STAFF,
        CONFIG.ROLES.PAYABLE_HEAD,
        CONFIG.ROLES.CPO,
        CONFIG.ROLES.ADMIN
      ].includes(role)
    };
  },

  getEffectivePermissions() {
    if (!this.permissions) {
      this.permissions = this.normalizePermissions({});
    }
    return this.permissions;
  },

  setupUI() {
    const user = Auth.getUser();
    const perms = this.getEffectivePermissions();

    const sidebar = document.getElementById('sidebar');
    if (sidebar) sidebar.innerHTML = Components.getSidebar('vouchers');

    const newVoucherBtn = document.getElementById('newVoucherBtn');
    const lookupBtn = document.getElementById('lookupBtn');
    const cpoActions = document.getElementById('cpoActions');
    const batchBtn = document.getElementById('batchStatusBtn');
    const releaseBtn = document.getElementById('releaseToUnitBtn');

    if (newVoucherBtn) {
      newVoucherBtn.classList.toggle('hidden', !perms.canCreateVoucher);
    }

    if (lookupBtn) {
      lookupBtn.classList.toggle('hidden', !perms.canLookup);
    }

    if (cpoActions) {
      cpoActions.style.display = perms.canBatchUpdateStatus || perms.canReleaseToUnit ? 'block' : 'none';
    }

    if (batchBtn) {
      batchBtn.style.display = perms.canBatchUpdateStatus ? 'inline-flex' : 'none';
    }

    if (releaseBtn) {
      releaseBtn.style.display = perms.canReleaseToUnit ? 'inline-flex' : 'none';
    }
  },

  handleUrlParams() {
    const urlParams = new URLSearchParams(window.location.search);

    if (urlParams.get('new') === 'true') this.openVoucherForm();
    if (urlParams.get('lookup') === 'true') this.openLookupModal();

    const editRow = urlParams.get('edit');
    if (editRow) this.editVoucher(parseInt(editRow, 10));

    const filter = urlParams.get('filter');
    if (filter) {
      const map = { pending: 'Pending Deletion', unpaid: 'Unpaid', paid: 'Paid' };
      if (map[filter]) {
        const statusFilter = document.getElementById('statusFilter');
        if (statusFilter) statusFilter.value = map[filter];
        this.filters.status = map[filter];
        this.currentPage = 1;
        this.loadVouchers();
      }
    }
  },

  // Auto-convert inputs to UPPERCASE
  setupUppercaseInputs() {
    // List of input IDs to convert to uppercase
    const uppercaseFields = [
      'formOldVoucherNumber',
      'formPayee',
      'formAccountOrMail',
      'formParticular'
    ];

    uppercaseFields.forEach(fieldId => {
      const field = document.getElementById(fieldId);
      if (field) {
        field.addEventListener('input', function () {
          // Get cursor position
          const start = this.selectionStart;
          const end = this.selectionEnd;

          // Convert to uppercase
          this.value = this.value.toUpperCase();

          // Restore cursor position
          this.setSelectionRange(start, end);
        });
      }
    });

    // Listen for SWR background data arrivals
    document.addEventListener('apiDataUpdated', (e) => {
      const { action, data, params } = e.detail;

      // If categories arrived fresh in the background, update UI silently
      if (action === 'getCategories') {
        this.categories = data.categories || [];
        this.populateCategoryDropdowns();
      }

      // If the vouchers list arrived fresh in the background, only re-render if we are still on the same page/filters
      if (action === 'getVouchers') {
        // Very simple check to ensure we only apply if the user hasn't heavily navigated away
        // In a robust app, we would deep equal the params, but this is a solid start for 700kbps connections
        if (params.page === this.currentPage && !this.isGlobalSearchMode) {
          this.vouchers = data.vouchers || [];
          this.totalCount = data.totalCount || 0;
          this.totalPages = data.totalPages || 0;
          this.renderVoucherList();
        }
      }
    });
  },

  getAmountFieldIds() {
    return ['formContractSum', 'formGrossAmount', 'formVat', 'formWht', 'formStampDuty'];
  },

  parseCurrencyInputValue(value) {
    const cleaned = String(value ?? '')
      .replace(/,/g, '')
      .replace(/[^\d.]/g, '');
    if (!cleaned) return 0;
    const parsed = parseFloat(cleaned);
    return Number.isFinite(parsed) ? parsed : 0;
  },

  formatCurrencyInputValue(value) {
    const raw = String(value ?? '').replace(/,/g, '').trim();
    if (!raw) return '';
    const num = this.parseCurrencyInputValue(raw);
    if (!Number.isFinite(num)) return '';
    return num.toLocaleString('en-NG', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  },

  setAmountFieldValue(fieldId, value) {
    const el = document.getElementById(fieldId);
    if (!el) return;
    const num = Number(value);
    if (!Number.isFinite(num) || num === 0) {
      el.value = '';
      return;
    }
    el.value = num.toLocaleString('en-NG', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  },

  getAmountFieldValue(fieldId) {
    const el = document.getElementById(fieldId);
    return el ? this.parseCurrencyInputValue(el.value) : 0;
  },

  recalculateVoucherNet() {
    const gross = this.getAmountFieldValue('formGrossAmount');
    const vat = this.getAmountFieldValue('formVat');
    const wht = this.getAmountFieldValue('formWht');
    const stamp = this.getAmountFieldValue('formStampDuty');
    const net = gross - (vat + wht + stamp);

    const netField = document.getElementById('formNet');
    if (netField) netField.value = net.toFixed(2);

    const netDisplay = document.getElementById('formNetDisplay');
    if (netDisplay) netDisplay.textContent = Utils.formatNumber(net);
  },

  bindAmountFormattingHandlers() {
    if (this.amountFormattingReady) return;
    this.amountFormattingReady = true;

    this.getAmountFieldIds().forEach((fieldId) => {
      const el = document.getElementById(fieldId);
      if (!el) return;

      el.addEventListener('focus', () => {
        const num = this.parseCurrencyInputValue(el.value);
        el.value = num ? String(num) : '';
      });

      el.addEventListener('input', () => {
        el.value = String(el.value || '').replace(/[^\d.,]/g, '');
        this.recalculateVoucherNet();
      });

      el.addEventListener('blur', () => {
        el.value = this.formatCurrencyInputValue(el.value);
        this.recalculateVoucherNet();
      });
    });
  },

  // ===== Categories =====
  async loadCategories() {
    // Check cache first
    const cached = sessionStorage.getItem('pv2026_categories');
    if (cached) {
      this.categories = JSON.parse(cached);
      this.populateCategoryDropdowns();
      return;
    }

    const result = await API.getCategories();
    if (result.success) {
      this.categories = result.categories || [];
      sessionStorage.setItem('pv2026_categories', JSON.stringify(this.categories));
      this.populateCategoryDropdowns();
    }
  },

  async loadSystemConfig() {
    const result = await API.getSystemConfig();
    if (result.success) {
      this.systemConfig = result.config;
      this.populateAccountTypeDropdown();
    }
  },

  populateAccountTypeDropdown() {
    const formAccountType = document.getElementById('formAccountType');
    if (!formAccountType || !this.systemConfig) return;
    
    // Check if event listener is already attached to prevent duplicates
    if (!this._accountTypeListenerAdded) {
      formAccountType.addEventListener('change', () => this.handleAccountTypeChange());
      this._accountTypeListenerAdded = true;
    }
    
    const cur = formAccountType.value;
    formAccountType.innerHTML = '<option value="">Select Type</option>';
    
    const accountTypes = Object.keys(this.systemConfig.accountTypes || {}).sort();
    accountTypes.forEach(type => {
      formAccountType.innerHTML += `<option value="${type}">${type}</option>`;
    });
    
    if (cur) {
      formAccountType.value = cur;
    }
    this.handleAccountTypeChange();
  },

  handleAccountTypeChange() {
    const formAccountType = document.getElementById('formAccountType');
    const formSubAccountType = document.getElementById('formSubAccountType');
    
    if (!formAccountType || !formSubAccountType || !this.systemConfig) return;
    
    const selectedType = formAccountType.value;
    const subTypes = (this.systemConfig.accountTypes && this.systemConfig.accountTypes[selectedType]) || [];
    
    const curSub = formSubAccountType.value;
    formSubAccountType.innerHTML = '<option value="">Select Sub-Type</option>';
    
    if (subTypes.length > 0) {
      subTypes.forEach(sub => {
        formSubAccountType.innerHTML += `<option value="${sub}">${sub}</option>`;
      });
      formSubAccountType.disabled = false;
      formSubAccountType.required = true;
      const ast = document.getElementById('subAccountTypeAst');
      if (ast) ast.style.display = 'inline';
      
      // Keep old selection if valid
      if (subTypes.includes(curSub)) {
        formSubAccountType.value = curSub;
      }
    } else {
      formSubAccountType.disabled = true;
      formSubAccountType.required = false;
      formSubAccountType.value = '';
      const ast = document.getElementById('subAccountTypeAst');
      if (ast) ast.style.display = 'none';
    }
  },

  accountTypeRequiresSubType(accountType) {
    if (!this.systemConfig || !this.systemConfig.accountTypes) return false;
    const subTypes = this.systemConfig.accountTypes[accountType] || [];
    return Array.isArray(subTypes) && subTypes.length > 0;
  },

  populateCategoryDropdowns() {
    // Dropdowns on forms
    const dropdowns = document.querySelectorAll('.category-select');
    dropdowns.forEach(dd => {
      const cur = dd.value;
      dd.innerHTML = '<option value="">Select Category</option>';
      this.categories.forEach(cat => dd.innerHTML += `<option value="${cat}">${cat}</option>`);
      if (cur) dd.value = cur;
    });

    // Filter dropdown
    const filterDropdown = document.getElementById('categoryFilter');
    if (filterDropdown) {
      filterDropdown.innerHTML = '<option value="All">All Categories</option>';
      this.categories.forEach(cat => filterDropdown.innerHTML += `<option value="${cat}">${cat}</option>`);
    }

    // Release dropdown category filter
    const releaseCat = document.getElementById('releaseCategoryFilter');
    if (releaseCat) {
      // keep "All"
      const cur = releaseCat.value;
      releaseCat.innerHTML = '<option value="All">All Categories</option>';
      this.categories.forEach(cat => releaseCat.innerHTML += `<option value="${cat}">${cat}</option>`);
      if (cur) releaseCat.value = cur;
    }
  },

  // ===== Vouchers list =====
  async loadVouchers(options = {}) {
    const requestId = ++this._loadSeq;
    if (!options.skipLoading) this.showLoading(true);

    try {
      const result = await API.getVouchers('2026', this.filters, this.currentPage, this.pageSize);
      if (requestId !== this._loadSeq) return;

      if (!result.success) {
        Utils.showToast(result.error || 'Failed to load vouchers', 'error');
        return;
      }

      this.vouchers = result.vouchers || [];
      this.totalCount = result.totalCount || 0;
      this.totalPages = result.totalPages || 0;

      this.vouchers.sort((a, b) => {
        const ar = Number(a?.rowIndex || 0);
        const br = Number(b?.rowIndex || 0);
        if (br !== ar) return br - ar;
        const ad = a?.date ? new Date(a.date).getTime() : 0;
        const bd = b?.date ? new Date(b.date).getTime() : 0;
        return bd - ad;
      });

      this.renderVoucherList();
    } catch (e) {
      console.error('loadVouchers error:', e);
      Utils.showToast('Error loading vouchers', 'error');
    } finally {
      if (!options.skipLoading) this.showLoading(false);
    }
  },

  renderVoucherList() {
    const container = document.getElementById('vouchersList');
    if (!container) return;

    const countEl = document.getElementById('voucherCount');
    if (countEl) {
      countEl.textContent = `${this.totalCount} voucher(s) (Page ${this.currentPage} of ${this.totalPages || 1})`;
    }

    if (!this.vouchers.length) {
      container.innerHTML = Components.getEmptyState('No vouchers found', 'fa-file-invoice');
      const pg = document.getElementById('paginationContainer');
      if (pg) pg.innerHTML = '';
      return;
    }

    const user = Auth.getUser();
    const canSelect = user && [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.ADMIN,
      CONFIG.ROLES.CPO
    ].includes(user.role);

    // Use DocumentFragment for 10x faster DOM insertion
    const fragment = document.createDocumentFragment();

    const tableContainer = document.createElement('div');
    tableContainer.className = 'table-container';

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    thead.innerHTML = `
        <tr>
            ${canSelect ? '<th><input type="checkbox" id="selectAll"></th>' : ''}
            <th>S/N</th>
            <th>Voucher No.</th>
            <th>Payee</th>
            <th>Particular</th>
            <th>Gross Amount</th>
            <th>Category</th>
            <th>Account Type</th>
            <th>Control No.</th>
            <th>Status</th>
            <th>Pmt Month</th>
            <th>Actions</th>
        </tr>
    `;
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    const startSN = (this.currentPage - 1) * this.pageSize;

    this.vouchers.forEach((v, idx) => {
      const sn = startSN + idx + 1;
      const tr = document.createElement('tr');
      tr.dataset.row = v.rowIndex;

      tr.innerHTML = `
          ${canSelect ? `<td><input type="checkbox" class="voucher-checkbox" value="${v.rowIndex}"></td>` : ''}
          <td>${sn}</td>
          <td><strong>${v.accountOrMail || '-'}</strong></td>
          <td title="${v.payee || ''}">${Utils.truncate(v.payee || '-', 20)}</td>
          <td title="${v.particular || ''}"><div class="particular-cell">${v.particular || '-'}</div></td>
          <td>${Utils.formatCurrency(v.grossAmount || 0)}</td>
          <td>${v.categories || '-'}</td>
          <td title="${v.subAccountType ? `${v.accountType} - ${v.subAccountType}` : v.accountType || ''}">${v.subAccountType ? `${v.accountType} (${v.subAccountType})` : (v.accountType || '-')}</td>
          <td>${v.controlNumber || '<span class="text-muted">-</span>'}</td>
          <td>${Utils.getStatusBadge(v.status || '')}</td>
          <td>${v.pmtMonth || '-'}</td>
          <td>
            <div class="action-buttons">
              <button class="btn btn-sm btn-secondary" onclick="Vouchers.viewVoucher(${v.rowIndex})" title="View">
                <i class="fas fa-eye"></i>
              </button>
              ${this.getActionButtons(v)}
            </div>
          </td>
      `;
      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableContainer.appendChild(table);
    fragment.appendChild(tableContainer);

    // Clear and append all at once
    container.innerHTML = '';
    container.appendChild(fragment);

    // Reattach listeners securely
    this.updateSelection?.();
    this.renderPagination?.();

    if (canSelect) {
      const selectAll = document.getElementById('selectAll');
      if (selectAll) {
        selectAll.addEventListener('change', () => {
          document.querySelectorAll('.voucher-checkbox').forEach(cb => cb.checked = selectAll.checked);
          this.updateSelection();
        });
      }
      document.querySelectorAll('.voucher-checkbox').forEach(cb => {
        cb.addEventListener('change', () => this.updateSelection());
      });
    }

    this.renderPagination();
  },

  updateSelection() {
    const checked = document.querySelectorAll('.voucher-checkbox:checked');
    this.selectedVouchers = Array.from(checked).map(cb => parseInt(cb.value, 10));
  },

  getActionButtons(voucher) {
    const perms = this.getEffectivePermissions();
    const user = Auth.getUser();
    if (!user) return '';

    let buttons = '';

    // Edit
    if (perms.canEditVoucher && user.role !== CONFIG.ROLES.CPO && voucher.status !== 'Pending Deletion') {
      const canEdit =
        voucher.status === 'Unpaid' ||
        user.role === CONFIG.ROLES.ADMIN ||
        user.role === CONFIG.ROLES.PAYABLE_HEAD;

      if (canEdit) {
        buttons += `
        <button class="btn btn-sm btn-primary" onclick="Vouchers.editVoucher(${voucher.rowIndex})" title="Edit">
          <i class="fas fa-edit"></i>
        </button>
      `;
      }
    }

    // Status update
    if (perms.canUpdateStatus && voucher.status !== 'Pending Deletion') {
      buttons += `
      <button class="btn btn-sm btn-success" onclick="Vouchers.openStatusModal(${voucher.rowIndex})" title="Update Status">
        <i class="fas fa-check-circle"></i>
      </button>
    `;
    }

    // Deletion workflow
    if (voucher.status === 'Pending Deletion') {
      if (perms.canApproveDelete) {
        buttons += `
        <button class="btn btn-sm btn-success" onclick="Vouchers.approveDelete(${voucher.rowIndex})" title="Approve Delete">
          <i class="fas fa-check"></i>
        </button>
        <button class="btn btn-sm btn-warning" onclick="Vouchers.rejectDelete(${voucher.rowIndex})" title="Reject Delete">
          <i class="fas fa-times"></i>
        </button>
      `;
      }

      if (perms.canRequestDelete) {
        buttons += `
        <button class="btn btn-sm btn-secondary" onclick="Vouchers.cancelDeleteRequest(${voucher.rowIndex})" title="Undo Request">
          <i class="fas fa-undo"></i>
        </button>
      `;
      }
    } else {
      if (perms.canRequestDelete) {
        buttons += `
        <button class="btn btn-sm btn-danger" onclick="Vouchers.requestDelete(${voucher.rowIndex})" title="Request Deletion">
          <i class="fas fa-trash"></i>
        </button>
      `;
      }
    }

    return buttons;
  },

  renderPagination() {
    const container = document.getElementById('paginationContainer');
    if (!container) return;

    if (this.totalPages <= 1) {
      container.innerHTML = '';
      return;
    }

    container.innerHTML = `
      <div class="pagination">
        <button class="btn btn-sm btn-secondary" ${this.currentPage === 1 ? 'disabled' : ''} onclick="Vouchers.goToPage(1)">
          <i class="fas fa-angle-double-left"></i>
        </button>
        <button class="btn btn-sm btn-secondary" ${this.currentPage === 1 ? 'disabled' : ''} onclick="Vouchers.goToPage(${this.currentPage - 1})">
          <i class="fas fa-chevron-left"></i>
        </button>

        <span class="pagination-info" style="padding: 5px 15px;">Page ${this.currentPage} of ${this.totalPages}</span>

        <button class="btn btn-sm btn-secondary" ${this.currentPage === this.totalPages ? 'disabled' : ''} onclick="Vouchers.goToPage(${this.currentPage + 1})">
          <i class="fas fa-chevron-right"></i>
        </button>
        <button class="btn btn-sm btn-secondary" ${this.currentPage === this.totalPages ? 'disabled' : ''} onclick="Vouchers.goToPage(${this.totalPages})">
          <i class="fas fa-angle-double-right"></i>
        </button>

        <select class="form-control" style="width: auto; margin-left: 15px;" onchange="Vouchers.changePageSize(this.value)">
          <option value="25" ${this.pageSize === 25 ? 'selected' : ''}>25 per page</option>
          <option value="50" ${this.pageSize === 50 ? 'selected' : ''}>50 per page</option>
          <option value="100" ${this.pageSize === 100 ? 'selected' : ''}>100 per page</option>
        </select>
      </div>
    `;
  },

  goToPage(page) {
    if (page < 1 || page > this.totalPages || page === this.currentPage) return;
    this.currentPage = page;
    this.loadVouchers();
  },

  changePageSize(size) {
    this.pageSize = parseInt(size, 10);
    this.currentPage = 1;
    this.loadVouchers();
  },

  applyFilters(options = {}) {
    this.filters.status = document.getElementById('statusFilter')?.value || 'All';
    this.filters.category = document.getElementById('categoryFilter')?.value || 'All';
    this.filters.searchTerm = document.getElementById('searchInput')?.value.trim() || '';
    this.filters.pmtMonth = document.getElementById('pmtMonthFilter')?.value || 'All';
    this.filters.dateFrom = document.getElementById('dateFromFilter')?.value || '';
    this.filters.dateTo = document.getElementById('dateToFilter')?.value || '';
    this.filters.amountMin = document.getElementById('amountMinFilter')?.value || '';
    this.filters.amountMax = document.getElementById('amountMaxFilter')?.value || '';
    this.filters.release = document.getElementById('releaseFilter')?.value || 'All';

    const filtersKey = JSON.stringify(this.filters);
    if (!options.force && this._lastFiltersKey === filtersKey) return;
    this._lastFiltersKey = filtersKey;
    this.currentPage = 1;

    this.updateActiveFiltersText();

    if (this.filters.searchTerm && options.forceGlobal) {
      this.globalSearch();
      return;
    }

    this.isGlobalSearchMode = false;
    this.loadVouchers({ skipLoading: !!options.skipLoading });
  },

  runSearch() {
    const term = document.getElementById('searchInput')?.value.trim() || '';
    if (!term) {
      this.applyFilters();
      return;
    }
    this.applyFilters({ forceGlobal: true, force: true });
  },

  // ===== View voucher (with payee unpaid total) =====
  async viewVoucher(rowIndex) {
    let voucher;

    // If global search mode, fetch from 2026
    if (this.isGlobalSearchMode) {
      this.showLoading(true);
      const r = await this.get2026VoucherForAction(rowIndex);
      this.showLoading(false);
      if (!r.success) {
        Utils.showToast(r.error || 'Voucher not found', 'error');
        return;
      }
      voucher = r.voucher;
    } else {
      // Normal mode
      voucher = this.vouchers.find(v => v.rowIndex === rowIndex);
      if (!voucher) {
        Utils.showToast('Voucher not found', 'error');
        return;
      }
    }

    this.selectedVoucher = voucher;
    const modal = document.getElementById('viewVoucherModal');
    const content = document.getElementById('viewVoucherContent');
    if (!modal || !content) return;

    content.innerHTML = `
      <div class="voucher-details">
        <!-- Header -->
        <div class="voucher-header-print">
          <h2>PAYMENT VOUCHER</h2>
          <div class="voucher-number">${voucher.accountOrMail || '-'}</div>
          <div class="voucher-date">Date: ${Utils.formatDate(voucher.date)}</div>
        </div>

        <!-- Status Section -->
        <div class="detail-section">
          <div class="detail-section-title">Status Information</div>
          <div class="detail-row">
            <div class="detail-group">
              <label>Current Status</label>
              <div>${Utils.getStatusBadge(voucher.status)}</div>
            </div>
            <div class="detail-group">
              <label>Payment Month</label>
              <div>${voucher.pmtMonth || '-'}</div>
            </div>
            <div class="detail-group">
              <label>Control Number</label>
              <div>${voucher.controlNumber || '<span class="text-muted">Not Released</span>'}</div>
            </div>
          </div>
        </div>

        <!-- Payee Section -->
        <div class="detail-section">
          <div class="detail-section-title">Payee Information</div>
          <div class="detail-group full-width">
            <label>Payee Name</label>
            <div><strong>${voucher.payee || '-'}</strong></div>
          </div>
          <div class="detail-group full-width" style="margin-top: 12px;">
            <label>Particular / Description</label>
            <div>${voucher.particular || '-'}</div>
          </div>
        </div>

        <!-- Financial Section -->
        <div class="detail-section">
          <div class="detail-section-title">Financial Details</div>
          ${voucher.contractSum ? `
            <div class="detail-group full-width" style="margin-bottom: 15px;">
              <label>Contract Sum</label>
              <div><strong>${Utils.formatCurrency(voucher.contractSum)}</strong></div>
            </div>
          ` : ''}
          
          <table class="financial-table">
            <thead>
              <tr>
                <th>Description</th>
                <th class="amount">Amount (₦)</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>Gross Amount</td>
                <td class="amount">${Utils.formatCurrency(voucher.grossAmount || 0)}</td>
              </tr>
              <tr>
                <td>Less: VAT</td>
                <td class="amount">(${Utils.formatCurrency(voucher.vat || 0)})</td>
              </tr>
              <tr>
                <td>Less: WHT</td>
                <td class="amount">(${Utils.formatCurrency(voucher.wht || 0)})</td>
              </tr>
              <tr>
                <td>Less: Stamp Duty</td>
                <td class="amount">(${Utils.formatCurrency(voucher.stampDuty || 0)})</td>
              </tr>
              <tr class="net-row">
                <td><strong>NET Amount Payable</strong></td>
                <td class="amount"><strong>${Utils.formatCurrency(voucher.net || 0)}</strong></td>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- Classification Section -->
        <div class="detail-section">
          <div class="detail-section-title">Classification</div>
          <div class="detail-row">
            <div class="detail-group">
              <label>Category</label>
              <div>${voucher.categories || '-'}</div>
            </div>
            <div class="detail-group">
              <label>Account Type</label>
              <div>${voucher.accountType || '-'} ${voucher.subAccountType ? `(${voucher.subAccountType})` : ''}</div>
            </div>
            <div class="detail-group">
              <label>Old Voucher No.</label>
              <div>${voucher.oldVoucherNumber || '-'}</div>
            </div>
          </div>
        </div>

        <!-- Payee Unpaid Total -->
        <div id="payeeUnpaidContainer" class="payee-unpaid-warning" style="display:none;">
           <!-- Content loaded asynchronously -->
        </div>
      </div>
    `;

    modal.classList.add('active');

    // Load unpaid stats asynchronously (non-blocking)
    this.loadPayeeUnpaidStats(voucher.payee);
  },

  async loadPayeeUnpaidStats(payee) {
    const container = document.getElementById('payeeUnpaidContainer');
    if (!container) return;

    try {
      const result = await API.getVouchers('2026', { searchTerm: payee }, 1, 100);
      if (result.success) {
        let total = 0;
        let count = 0;
        (result.vouchers || []).forEach(v => {
          if (String(v.payee || '').trim() === String(payee || '').trim() && v.status === 'Unpaid') {
            total += Number(v.grossAmount || 0);
            count++;
          }
        });

        if (count > 0) {
          container.innerHTML = `
            <div class="warning-title">
              <i class="fas fa-exclamation-triangle"></i>
              Total Unpaid for "${payee || ''}"
            </div>
            <div class="warning-amount">${Utils.formatCurrency(total)}</div>
            <div class="warning-note">${count} unpaid voucher(s) found</div>
          `;
          container.style.display = 'block';
        }
      }
    } catch (e) {
      console.error('payee unpaid calc error:', e);
    }
  },

  // ===== Voucher form =====
  openVoucherForm(voucher = null) {
    this.isEditMode = !!voucher;
    this.selectedVoucher = voucher;

    const modal = document.getElementById('voucherFormModal');
    const title = document.getElementById('voucherFormTitle');
    const form = document.getElementById('voucherForm');
    if (!modal || !title || !form) return;

    const formatDateForInput = (value) => {
      if (!value) return '';
      if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
      const d = new Date(value);
      return isNaN(d.getTime()) ? '' : d.toISOString().split('T')[0];
    };

    title.textContent = this.isEditMode ? 'Edit Voucher' : 'Create New Voucher';
    form.reset();

    const lumpSumRadio = document.getElementById('paymentLumpSum');
    if (lumpSumRadio) lumpSumRadio.checked = true;

    const otherPartGroup = document.getElementById('otherPartPaymentGroup');
    if (otherPartGroup) otherPartGroup.classList.add('hidden');

    const contractSumGroup = document.getElementById('contractSumGroup');
    const contractSumRequired = document.getElementById('contractSumRequired');
    if (contractSumGroup) contractSumGroup.classList.remove('required-field');
    if (contractSumRequired) contractSumRequired.style.display = 'none';

    this.updateParticularHint('lumpsum');
    this.populateCategoryDropdowns();

    if (voucher) {
      document.getElementById('formOldVoucherNumber').value = voucher.oldVoucherNumber || '';
      document.getElementById('formPayee').value = voucher.payee || '';
      document.getElementById('formAccountOrMail').value = voucher.accountOrMail || '';
      document.getElementById('formParticular').value = voucher.particular || '';
      this.setAmountFieldValue('formContractSum', voucher.contractSum);
      this.setAmountFieldValue('formGrossAmount', voucher.grossAmount);
      this.setAmountFieldValue('formVat', voucher.vat);
      this.setAmountFieldValue('formWht', voucher.wht);
      this.setAmountFieldValue('formStampDuty', voucher.stampDuty);
      document.getElementById('formNet').value = voucher.net || '';
      document.getElementById('formNetDisplay').textContent = Utils.formatNumber(voucher.net || 0);
      document.getElementById('formCategories').value = voucher.categories || '';
      document.getElementById('formAccountType').value = voucher.accountType || '';
      this.handleAccountTypeChange();
      document.getElementById('formSubAccountType').value = voucher.subAccountType || '';
      document.getElementById('formDate').value = formatDateForInput(voucher.date);

      this.detectPaymentType(voucher.particular || '');
    } else {
      document.getElementById('formDate').value = new Date().toISOString().split('T')[0];
      document.getElementById('formNetDisplay').textContent = '0.00';
      document.getElementById('formNet').value = '0';
    }

    const oldVoucher = (voucher && voucher.oldVoucherNumber) ? String(voucher.oldVoucherNumber).trim() : '';
    const oldAvailRaw = voucher ? voucher.oldVoucherAvailable : '';
    const oldAvail = (oldAvailRaw === undefined || oldAvailRaw === null)
      ? ''
      : String(oldAvailRaw).trim().toLowerCase();

    this.setOldVoucherAvailability('');
    if (oldVoucher) {
      this.setOldVoucherAvailability('yes');
    } else if (oldAvail === 'yes' || oldAvail === 'true') {
      this.setOldVoucherAvailability('yes');
    } else if (oldAvail === 'no' || oldAvail === 'false') {
      this.setOldVoucherAvailability('no');
    }
    this.applyOldVoucherAvailability();

    this.setupPaymentTypeHandlers();
    this.recalculateVoucherNet();
    modal.classList.add('active');
  },

  getOldVoucherAvailability() {
    return document.querySelector('input[name="oldVoucherAvailable"]:checked')?.value || '';
  },

  setOldVoucherAvailability(value) {
    const yes = document.getElementById('oldVoucherAvailableYes');
    const no = document.getElementById('oldVoucherAvailableNo');
    if (!yes || !no) return;

    if (value === 'yes') {
      yes.checked = true;
      no.checked = false;
    } else if (value === 'no') {
      yes.checked = false;
      no.checked = true;
    } else {
      yes.checked = false;
      no.checked = false;
    }
  },

  applyOldVoucherAvailability() {
    const choice = this.getOldVoucherAvailability();
    const input = document.getElementById('formOldVoucherNumber');
    const required = document.getElementById('oldVoucherRequired');
    const lookupBtn = document.getElementById('oldVoucherLookupBtn');
    if (!input) return;

    if (choice === 'yes') {
      input.disabled = false;
      input.required = true;
      if (required) required.style.display = 'inline';
      if (lookupBtn) lookupBtn.disabled = false;
    } else if (choice === 'no') {
      input.value = '';
      input.disabled = true;
      input.required = false;
      if (required) required.style.display = 'none';
      if (lookupBtn) lookupBtn.disabled = true;
    } else {
      input.disabled = true;
      input.required = false;
      if (required) required.style.display = 'none';
      if (lookupBtn) lookupBtn.disabled = true;
    }
  },

  detectPaymentType(particular) {
    const lower = particular.toLowerCase();

    if (/^(first|1st)\s*part[-\s]?p(ay)?m(en)?t/i.test(lower)) {
      document.getElementById('paymentFirstPart').checked = true;
      this.updateParticularHint('firstPart');
    } else if (/^(bal(ance)?|final|fnl)\s*p(ay)?m(en)?t/i.test(lower)) {
      document.getElementById('paymentBalance').checked = true;
      this.updateParticularHint('balance');
    } else if (/^(2nd|second|3rd|third|4th|fourth|5th|fifth)\s*part[-\s]?p(ay)?m(en)?t/i.test(lower)) {
      document.getElementById('paymentOtherPart').checked = true;
      document.getElementById('otherPartPaymentGroup').classList.remove('hidden');
      this.updateParticularHint('otherPart');
    }
  },

  async editVoucher(rowIndex) {
    const perms = this.getEffectivePermissions();
    const user = Auth.getUser();
    if (!perms.canEditVoucher || (user && user.role === CONFIG.ROLES.CPO)) {
      Utils.showToast('You are not authorized to edit vouchers', 'error');
      return;
    }

    // If global search mode, fetch from 2026 to avoid wrong-year match
    if (this.isGlobalSearchMode) {
      this.showLoading(true);
      const r = await this.get2026VoucherForAction(rowIndex);
      this.showLoading(false);
      if (!r.success) {
        Utils.showToast(r.error || 'Voucher not found', 'error');
        return;
      }
      return this.openVoucherForm(r.voucher);
    }

    // Normal mode - existing logic
    let voucher = this.vouchers.find(v => v.rowIndex === rowIndex);

    if (!voucher) {
      this.showLoading(true);
      const result = await API.getVoucherByRow(rowIndex, '2026');
      this.showLoading(false);
      if (!result.success) {
        Utils.showToast(result.error || 'Voucher not found', 'error');
        return;
      }
      voucher = result.voucher;
    }

    this.openVoucherForm(voucher);
  },

  async saveVoucher() {
    const perms = this.getEffectivePermissions();
    const user = Auth.getUser();
    if (this.isEditMode && (!perms.canEditVoucher || (user && user.role === CONFIG.ROLES.CPO))) {
      Utils.showToast('You are not authorized to edit vouchers', 'error');
      return;
    }

    const particularValidation = this.validateParticular();
    if (!particularValidation.valid) {
      Utils.showToast(particularValidation.error, 'error');
      return;
    }

    const paymentType = document.querySelector('input[name="paymentType"]:checked')?.value;
    const contractSum = this.getAmountFieldValue('formContractSum');

    if (paymentType === 'firstPart' && !contractSum) {
      Utils.showToast('Contract Sum is required for First Part-Payment', 'error');
      return;
    }

    const payee = document.getElementById('formPayee').value.trim();
    const accountOrMail = document.getElementById('formAccountOrMail').value.trim();
    const gross = this.getAmountFieldValue('formGrossAmount');
    const vat = this.getAmountFieldValue('formVat');
    const wht = this.getAmountFieldValue('formWht');
    const stampDuty = this.getAmountFieldValue('formStampDuty');
    const net = gross - (vat + wht + stampDuty);
    const oldVoucherChoice = this.getOldVoucherAvailability();
    let oldVoucherNumber = document.getElementById('formOldVoucherNumber').value.trim();

    document.getElementById('formNet').value = net.toFixed(2);

    if (!payee) return Utils.showToast('Payee name is required', 'error');
    if (!accountOrMail) return Utils.showToast('Voucher Number is required', 'error');
    if (!gross) return Utils.showToast('Gross amount is required', 'error');
    
    // Validate Account Type and Category
    const categories = String(document.getElementById('formCategories').value || '').trim();
    const accountType = String(document.getElementById('formAccountType').value || '').trim();
    const subAccountType = String(document.getElementById('formSubAccountType').value || '').trim();
    const subTypeRequiredByConfig = this.accountTypeRequiresSubType(accountType);

    if (!categories) return Utils.showToast('Category is required', 'error');
    if (!accountType) return Utils.showToast('Account Type is required', 'error');
    if (subTypeRequiredByConfig && !subAccountType) {
      return Utils.showToast('Sub Account Type is required for the selected Account Type', 'error');
    }

    if (oldVoucherChoice === 'yes' && !oldVoucherNumber) {
      return Utils.showToast('Old voucher number is required when "Yes" is selected', 'error');
    }
    if (oldVoucherChoice === 'no') {
      oldVoucherNumber = '';
    }

    const voucherData = {
      payee,
      accountOrMail,
      particular: document.getElementById('formParticular').value.trim(),
      contractSum: contractSum,
      grossAmount: gross,
      vat,
      wht,
      stampDuty,
      net,
      categories: document.getElementById('formCategories').value,
      accountType: document.getElementById('formAccountType').value,
      subAccountType: document.getElementById('formSubAccountType').value || '',
      date: document.getElementById('formDate').value,
      totalGross: gross,
      oldVoucherNumber: oldVoucherNumber,
      oldVoucherAvailable: oldVoucherChoice ? (oldVoucherChoice === 'yes' ? 'Yes' : 'No') : ''
    };

    this.showLoading(true);

    try {
      let result;

      if (this.isEditMode && this.selectedVoucher) {
        voucherData.controlNumber = this.selectedVoucher.controlNumber || '';
        result = await API.updateVoucher(this.selectedVoucher.rowIndex, voucherData);
      } else {
        voucherData.controlNumber = '';
        result = await API.createVoucher(voucherData);
      }

      if (result.success) {
        Utils.showToast(result.message || 'Voucher saved successfully', 'success');
        this.closeModal('voucherFormModal');
        this.selectedVoucher = null;
        this.isEditMode = false;
        await this.loadVouchers();
      } else {
        Utils.showToast(result.error || 'Failed to save voucher', 'error');
      }
    } catch (e) {
      console.error('saveVoucher error:', e);
      Utils.showToast('Error saving voucher', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  // ===== Lookup =====
  openLookupModal() {
    const modal = document.getElementById('lookupModal');
    if (!modal) return;
    document.getElementById('lookupVoucherNumber').value = '';
    const res = document.getElementById('lookupResult');
    if (res) {
      res.innerHTML = '';
      res.classList.add('hidden');
    }
    modal.classList.add('active');
  },

  async performLookup() {
    const voucherNumber = document.getElementById('lookupVoucherNumber').value.trim();
    if (!voucherNumber) return Utils.showToast('Please enter a voucher number', 'error');

    this.showLoading(true);

    try {
      const result = await API.lookupVoucher(voucherNumber);
      const container = document.getElementById('lookupResult');
      if (!container) return;

      container.classList.remove('hidden');

      if (!result.success) {
        container.innerHTML = `<div class="alert alert-error">${result.error || 'Lookup failed'}</div>`;
        return;
      }

      if (!result.found) {
        container.innerHTML = `<div class="alert alert-warning">${result.message || 'Not found'}</div>`;
        return;
      }

      // show found
      this.lookupResult = result;
      const v = result.voucher;

      // Already revalidated?
      if (result.alreadyRevalidated) {
        container.innerHTML = `
          <div class="alert alert-warning">
            <strong>${result.message}</strong><br>
            Existing Voucher: ${result.existingVoucherNumber || '-'}
          </div>
        `;
        return;
      }

      // Paid cannot revalidate
      if (!result.canRevalidate) {
        container.innerHTML = `
          <div class="alert alert-error">
            <strong>${result.message}</strong><br>
            ${result.reason || ''}
          </div>
        `;
        return;
      }

      container.innerHTML = `
        <div class="alert alert-success">
          <strong>${result.message}</strong>
          ${result.requiresAuthorization ? `<br><span class="text-warning">${result.warning || ''}</span>` : ''}
        </div>

        <div class="lookup-voucher-details">
          <h4>Voucher Details</h4>
          <p><strong>Payee:</strong> ${v.payee || '-'}</p>
          <p><strong>Voucher No:</strong> ${v.accountOrMail || '-'}</p>
          <p><strong>Particular:</strong> ${v.particular || '-'}</p>
          <p><strong>Amount:</strong> ${Utils.formatCurrency(v.grossAmount || 0)}</p>
          <p><strong>Status:</strong> ${Utils.getStatusBadge(v.status || '')}</p>
        </div>

        <button class="btn btn-primary" onclick="Vouchers.createFromLookup()" style="width:100%;margin-top:15px;">
          <i class="fas fa-plus"></i> Revalidate This Voucher
        </button>
      `;
    } catch (e) {
      console.error('performLookup error:', e);
      Utils.showToast('Error performing lookup', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  createFromLookup() {
    if (!this.lookupResult || !this.lookupResult.voucher) {
      Utils.showToast('No lookup result available', 'error');
      return;
    }

    this.closeModal('lookupModal');

    const v = this.lookupResult.voucher;
    this.openVoucherForm();

    setTimeout(() => {
      this.setOldVoucherAvailability('yes');
      this.applyOldVoucherAvailability();
      document.getElementById('formOldVoucherNumber').value = v.accountOrMail || '';
      document.getElementById('formPayee').value = v.payee || '';
      document.getElementById('formParticular').value = v.particular || '';
      this.setAmountFieldValue('formContractSum', v.contractSum);
      this.setAmountFieldValue('formGrossAmount', v.grossAmount);
      this.setAmountFieldValue('formVat', v.vat);
      this.setAmountFieldValue('formWht', v.wht);
      this.setAmountFieldValue('formStampDuty', v.stampDuty);
      document.getElementById('formCategories').value = v.categories || '';
      this.recalculateVoucherNet();
    }, 100);
  },

  async inlineLookup() {
    const oldVoucherChoice = this.getOldVoucherAvailability();
    if (oldVoucherChoice !== 'yes') {
      return Utils.showToast('Select "Yes" to enter an old voucher number', 'warning');
    }
    const oldVN = document.getElementById('formOldVoucherNumber').value.trim();
    if (!oldVN) return Utils.showToast('Enter an old voucher number first', 'warning');

    this.showLoading(true);

    try {
      const result = await API.lookupVoucher(oldVN);
      if (result.success && result.found && result.canRevalidate) {
        const v = result.voucher;

        document.getElementById('formPayee').value = v.payee || '';
        document.getElementById('formParticular').value = v.particular || '';
        this.setAmountFieldValue('formContractSum', v.contractSum);
        this.setAmountFieldValue('formGrossAmount', v.grossAmount);
        this.setAmountFieldValue('formVat', v.vat);
        this.setAmountFieldValue('formWht', v.wht);
        this.setAmountFieldValue('formStampDuty', v.stampDuty);
        document.getElementById('formCategories').value = v.categories || '';
        this.recalculateVoucherNet();

        Utils.showToast(`Found in ${result.sourceYear}. Filled fields.`, 'success');
      } else {
        Utils.showToast(result.message || 'Voucher not found / cannot revalidate', 'warning');
      }
    } catch (e) {
      console.error('inlineLookup error:', e);
      Utils.showToast('Lookup failed', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  // ===== Status =====
  async openStatusModal(rowIndex) {
    let voucher;

    // If global search mode, fetch from 2026
    if (this.isGlobalSearchMode) {
      this.showLoading(true);
      const r = await this.get2026VoucherForAction(rowIndex);
      this.showLoading(false);
      if (!r.success) {
        Utils.showToast(r.error || 'Voucher not found', 'error');
        return;
      }
      voucher = r.voucher;
    } else {
      // Normal mode
      voucher = this.vouchers.find(v => v.rowIndex === rowIndex);
      if (!voucher) {
        Utils.showToast('Voucher not found', 'error');
        return;
      }
    }

    this.selectedVoucher = voucher;

    document.getElementById('statusVoucherInfo').innerHTML = `
      <p><strong>Voucher:</strong> ${voucher.accountOrMail || '-'}</p>
      <p><strong>Payee:</strong> ${voucher.payee || '-'}</p>
      <p><strong>Amount:</strong> ${Utils.formatCurrency(voucher.grossAmount || 0)}</p>
    `;

    const user = Auth.getUser();
    const canSetMonth = user && (user.role === CONFIG.ROLES.CPO || user.role === CONFIG.ROLES.ADMIN);

    // Always show the group
    const pmtMonthGroup = document.getElementById('pmtMonthGroup');
    if (pmtMonthGroup) pmtMonthGroup.style.display = 'block';

    // Set values
    document.getElementById('newStatus').value = voucher.status || 'Unpaid';
    document.getElementById('newPmtMonth').value = voucher.pmtMonth || '';

    // Only CPO/Admin can edit payment month
    document.getElementById('newPmtMonth').disabled = !canSetMonth;

    // Require month only when setting Paid (and only if canSetMonth)
    const statusEl = document.getElementById('newStatus');
    const monthEl = document.getElementById('newPmtMonth');

    const updateMonthRequirement = () => {
      const st = statusEl.value;
      monthEl.required = (canSetMonth && st === 'Paid');
    };

    statusEl.onchange = updateMonthRequirement;
    updateMonthRequirement();

    document.getElementById('statusModal').classList.add('active');
  },



  async saveBatchStatus() {
    const cn = document.getElementById('batchControlNumber').value.trim();
    if (!cn) return Utils.showToast('Please enter a control number', 'error');

    const status = document.getElementById('batchStatus').value;
    const pmtMonth = document.getElementById('batchPmtMonth').value;

    this.showLoading(true);

    try {
      // Optimistic Update
      const oldStates = [];
      this.vouchers.forEach(v => {
        if ((v.controlNumber || '').trim() === cn) {
          oldStates.push({ v, oldStatus: v.status, oldPmtMonth: v.pmtMonth });
          v.status = status;
          if (pmtMonth) v.pmtMonth = pmtMonth;
        }
      });
      if (oldStates.length > 0) this.renderVoucherList();
      this.closeModal('batchStatusModal');

      const result = await API.batchUpdateStatus(cn, status, pmtMonth);
      if (result.success) {
        Utils.showToast(result.message, 'success');
      } else {
        // Revert Optimistic Update
        oldStates.forEach(s => {
          s.v.status = s.oldStatus;
          s.v.pmtMonth = s.oldPmtMonth;
        });
        if (oldStates.length > 0) this.renderVoucherList();
        Utils.showToast(result.error || 'Batch update failed', 'error');
      }
    } catch (e) {
      console.error('saveBatchStatus error:', e);
      Utils.showToast('Error updating batch status', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  // ===== Assign control number (manual) =====
  openAssignControlModal() {
    if (!this.selectedVouchers.length) return Utils.showToast('Please select vouchers first', 'warning');
    document.getElementById('selectedCount').textContent = this.selectedVouchers.length;
    document.getElementById('newControlNumber').value = '';
    document.getElementById('assignControlModal').classList.add('active');
  },

  async assignControlNumber() {
    const cn = document.getElementById('newControlNumber').value.trim();
    if (!cn) return Utils.showToast('Please enter a control number', 'error');
    if (!this.selectedVouchers.length) return Utils.showToast('No vouchers selected', 'error');

    this.showLoading(true);

    try {
      // Optimistic update
      const oldStates = [];
      this.selectedVouchers.forEach(idx => {
        const v = this.vouchers.find(x => x.rowIndex === idx);
        if (v) {
          oldStates.push({ v, oldCn: v.controlNumber });
          v.controlNumber = cn;
        }
      });
      if (oldStates.length > 0) this.renderVoucherList();
      this.closeModal('assignControlModal');

      const result = await API.assignControlNumber(this.selectedVouchers, cn);
      if (result.success) {
        Utils.showToast(result.message, 'success');
        this.selectedVouchers = [];
        // Background SWR handles sync
      } else {
        // Revert setup
        oldStates.forEach(s => s.v.controlNumber = s.oldCn);
        if (oldStates.length > 0) this.renderVoucherList();
        Utils.showToast(result.error || 'Assignment failed', 'error');
      }
    } catch (e) {
      console.error('assignControlNumber error:', e);
      Utils.showToast('Error assigning control number', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  // ===== Deletion workflow (MODAL-BASED) =====
  async requestDelete(rowIndex) {
    let voucher;

    // If global search mode, fetch from 2026
    if (this.isGlobalSearchMode) {
      this.showLoading(true);
      const r = await this.get2026VoucherForAction(rowIndex);
      this.showLoading(false);
      if (!r.success) {
        Utils.showToast(r.error || 'Voucher not found', 'error');
        return;
      }
      voucher = r.voucher;
    } else {
      // Normal mode
      voucher = this.vouchers.find(v => v.rowIndex === rowIndex);
      if (!voucher) {
        Utils.showToast('Voucher not found', 'error');
        return;
      }
    }

    this.deleteTargetVoucher = voucher;

    const isReleased = voucher.controlNumber && String(voucher.controlNumber).trim() !== '';

    const info = document.getElementById('deleteVoucherInfo');
    const reasonEl = document.getElementById('deleteReason');
    const approvalText = document.getElementById('deleteApprovalText');

    if (!info || !reasonEl || !approvalText) {
      Utils.showToast('Delete modal elements missing in vouchers.html', 'error');
      return;
    }

    info.innerHTML = `
    <p><strong>Voucher No:</strong> ${voucher.accountOrMail || '-'}</p>
    <p><strong>Payee:</strong> ${voucher.payee || '-'}</p>
    <p><strong>Amount:</strong> ${Utils.formatCurrency(voucher.grossAmount || 0)}</p>
    <p><strong>Status:</strong> ${Utils.getStatusBadge(voucher.status || '')}</p>
    ${isReleased ? `<p><strong>Control No:</strong> ${voucher.controlNumber}</p>` : ''}
  `;

    approvalText.innerHTML = isReleased
      ? '<strong>⚠️ This voucher has been RELEASED.</strong> Deletion requires approval from <strong>CPO or Admin</strong>.'
      : 'This will require approval from <strong>Payable Unit Head or Admin</strong>.';

    reasonEl.value = '';
    document.getElementById('deleteRequestModal').classList.add('active');
  },

  async submitDeleteRequest() {
    const voucher = this.deleteTargetVoucher;
    if (!voucher) return Utils.showToast('No voucher selected', 'error');

    const reason = document.getElementById('deleteReason').value.trim();
    if (!reason) return Utils.showToast('Reason for deletion is required', 'error');

    const confirmed = await Utils.confirm(
      `Submit deletion request?\n\nVoucher: ${voucher.accountOrMail}\nPayee: ${voucher.payee}\nReason: ${reason}`,
      'Confirm Deletion Request'
    );
    if (!confirmed) return;

    this.showLoading(true);
    this.closeModal('deleteRequestModal');

    try {
      // Optimistic update
      const originalStatus = voucher.status;
      voucher.status = 'Pending Deletion';
      this.renderVoucherList();

      const result = await API.requestDelete(voucher.rowIndex, reason, originalStatus || 'Unpaid');
      if (result.success) {
        Utils.showToast(result.message || 'Deletion request submitted', 'success');
        this.deleteTargetVoucher = null;
        if (this.pendingDeletionsLoaded) await this.loadPendingDeletions();
      } else {
        // Revert 
        voucher.status = originalStatus;
        this.renderVoucherList();
        Utils.showToast(result.error || 'Failed to request deletion', 'error');
      }
    } catch (e) {
      console.error('submitDeleteRequest error:', e);
      Utils.showToast('Error requesting deletion', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  async rejectDelete(rowIndex) {
    let v;

    // Add global search mode check
    if (this.isGlobalSearchMode) {
      this.showLoading(true);
      const r = await this.get2026VoucherForAction(rowIndex);
      this.showLoading(false);
      if (!r.success) {
        Utils.showToast(r.error || 'Voucher not found', 'error');
        return;
      }
      v = r.voucher;
    } else {
      // Normal mode - check pending deletions first, then vouchers
      v = this.pendingDeletions.find(x => x.rowIndex === rowIndex) ||
        this.vouchers.find(x => x.rowIndex === rowIndex);
      if (!v) {
        Utils.showToast('Voucher not found', 'error');
        return;
      }
    }

    this.rejectTargetVoucher = v;

    const info = document.getElementById('rejectVoucherInfo');
    const reasonEl = document.getElementById('rejectReason');
    if (!info || !reasonEl) {
      Utils.showToast('Reject modal elements missing', 'error');
      return;
    }

    info.innerHTML = `
    <p><strong>Voucher No:</strong> ${v.accountOrMail || '-'}</p>
    <p><strong>Payee:</strong> ${v.payee || '-'}</p>
    <p><strong>Amount:</strong> ${Utils.formatCurrency(v.grossAmount || 0)}</p>
  `;
    reasonEl.value = '';

    document.getElementById('rejectDeleteModal').classList.add('active');
  },

  async approveDelete(rowIndex) {
    // Add this check at the very beginning
    let voucher;
    if (this.isGlobalSearchMode) {
      this.showLoading(true);
      const r = await this.get2026VoucherForAction(rowIndex);
      this.showLoading(false);
      if (!r.success) {
        Utils.showToast(r.error || 'Voucher not found', 'error');
        return;
      }
      voucher = r.voucher;
    }

    // Then continue with your existing code
    const confirmed = await Utils.confirm('Approve permanent deletion? This cannot be undone.', 'Approve Deletion');
    if (!confirmed) return;

    this.showLoading(true);

    try {
      const result = await API.approveDelete(rowIndex);
      if (result.success) {
        Utils.showToast(result.message || 'Voucher deleted', 'success');
        await this.loadVouchers();
        if (this.pendingDeletionsLoaded) {
          await this.loadPendingDeletions();
          this.renderPendingDeletions();
        }
      } else {
        Utils.showToast(result.error || 'Approval failed', 'error');
      }
    } catch (e) {
      console.error('approveDelete error:', e);
      Utils.showToast('Error approving deletion', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  async submitRejectDelete() {
    const v = this.rejectTargetVoucher;
    if (!v) return Utils.showToast('No voucher selected', 'error');

    const reason = document.getElementById('rejectReason').value.trim();

    this.showLoading(true);
    this.closeModal('rejectDeleteModal');

    try {
      const result = await API.rejectDelete(v.rowIndex, reason);
      if (result.success) {
        Utils.showToast(result.message || 'Request rejected', 'success');
        this.rejectTargetVoucher = null;
        await this.loadVouchers();
        if (this.pendingDeletionsLoaded) {
          await this.loadPendingDeletions();
          this.renderPendingDeletions();
        }
      } else {
        Utils.showToast(result.error || 'Rejection failed', 'error');
      }
    } catch (e) {
      console.error('submitRejectDelete error:', e);
      Utils.showToast('Error rejecting deletion', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  async cancelDeleteRequest(rowIndex) {
    const confirmed = await Utils.confirm(
      'Undo/cancel this deletion request?\n\nThe voucher will be restored to its previous status.',
      'Cancel Deletion Request'
    );
    if (!confirmed) return;

    this.showLoading(true);

    try {
      const result = await API.cancelDeleteRequest(rowIndex);
      if (result.success) {
        Utils.showToast(result.message || 'Request cancelled', 'success');
        await this.loadVouchers();
        if (this.pendingDeletionsLoaded) await this.loadPendingDeletions();
      } else {
        Utils.showToast(result.error || 'Failed to cancel request', 'error');
      }
    } catch (e) {
      console.error('cancelDeleteRequest error:', e);
      Utils.showToast('Error cancelling request', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  // ===== Pending deletions (Approvers) =====
  async loadPendingDeletions() {
    const user = Auth.getUser();
    if (!user) return;

    const approverRoles = [CONFIG.ROLES.PAYABLE_HEAD, CONFIG.ROLES.CPO, CONFIG.ROLES.ADMIN];
    if (!approverRoles.includes(user.role)) return;

    try {
      const result = await API.getPendingDeletions();
      if (result.success) {
        this.pendingDeletions = result.vouchers || [];

        const card = document.getElementById('pendingDeletionsCard');
        const countEl = document.getElementById('pendingCount');

        if (card && this.pendingDeletions.length > 0) {
          card.style.display = 'block';
          if (countEl) countEl.textContent = this.pendingDeletions.length;
        } else if (card) {
          card.style.display = 'none';
        }

        this.pendingDeletionsLoaded = true;
      }
    } catch (e) {
      console.error('loadPendingDeletions error:', e);
    }
  },

  togglePendingDeletions() {
    const list = document.getElementById('pendingDeletionsList');
    const toggleText = document.getElementById('pendingToggleText');
    if (!list || !toggleText) return;

    if (list.classList.contains('hidden')) {
      list.classList.remove('hidden');
      toggleText.textContent = 'Hide';
      this.renderPendingDeletions();
    } else {
      list.classList.add('hidden');
      toggleText.textContent = 'Show';
    }
  },

  renderPendingDeletions() {
    const container = document.getElementById('pendingDeletionsList');
    if (!container) return;

    if (!this.pendingDeletions.length) {
      container.innerHTML = '<p class="text-muted text-center">No pending deletion requests</p>';
      return;
    }

    let html = `
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Voucher No.</th>
              <th>Payee</th>
              <th>Amount</th>
              <th>Control No.</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
    `;

    this.pendingDeletions.forEach(v => {
      const isReleased = v.controlNumber && String(v.controlNumber).trim() !== '';
      html += `
        <tr>
          <td><strong>${v.accountOrMail || '-'}</strong></td>
          <td title="${v.payee || ''}">${Utils.truncate(v.payee || '', 25)}</td>
          <td>${Utils.formatCurrency(v.grossAmount || 0)}</td>
          <td>${v.controlNumber || '-'} ${isReleased ? '<span class="badge badge-pending">Released</span>' : ''}</td>
          <td>
            <div class="action-buttons">
              <button class="btn btn-sm btn-success" onclick="Vouchers.approveDelete(${v.rowIndex})"><i class="fas fa-check"></i> Approve</button>
              <button class="btn btn-sm btn-warning" onclick="Vouchers.rejectDelete(${v.rowIndex})"><i class="fas fa-times"></i> Reject</button>
              <button class="btn btn-sm btn-secondary" onclick="Vouchers.viewVoucher(${v.rowIndex})"><i class="fas fa-eye"></i> View</button>
            </div>
          </td>
        </tr>
      `;
    });

    html += '</tbody></table></div>';
    container.innerHTML = html;
  },

  // ===== Enhanced Release Workflow =====
  openReleaseModal() {
    const user = Auth.getUser();
    const perms = this.getEffectivePermissions();

    if (!user) {
      Utils.showToast('User session not found', 'error');
      return;
    }

    if (!perms.canReleaseToUnit) {
      Utils.showToast('You are not authorized to release vouchers', 'error');
      return;
    }

    this.isPayableUnit = [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.ADMIN
    ].includes(user.role);

    this.isCPO = (user.role === CONFIG.ROLES.CPO);

    this.releaseSearchResults = [];

    this.selectedForRelease = (this.selectedVouchers || [])
      .map(rowIndex => this.vouchers.find(v => v.rowIndex === rowIndex))
      .filter(Boolean)
      .map(v => ({ ...v, sourceYear: '2026' }));

    const releaseSearchInput = document.getElementById('releaseSearchInput');
    const releaseStatusFilter = document.getElementById('releaseStatusFilter');
    const releaseSearchResults = document.getElementById('releaseSearchResults');
    const releaseControlNumber = document.getElementById('releaseControlNumber');
    const releaseTargetUnit = document.getElementById('releaseTargetUnit');
    const customTargetUnit = document.getElementById('customTargetUnit');
    const releasePurpose = document.getElementById('releasePurpose');
    const controlNumberGroup = document.getElementById('controlNumberGroup');
    const releasePurposeGroup = document.getElementById('releasePurposeGroup');
    const releaseStep1 = document.getElementById('releaseStep1');
    const releaseStep2 = document.getElementById('releaseStep2');
    const releaseModal = document.getElementById('releaseModal');

    if (!releaseModal) {
      Utils.showToast('Release modal not found in page', 'error');
      return;
    }

    if (releaseSearchInput) releaseSearchInput.value = '';
    if (releaseStatusFilter) releaseStatusFilter.value = 'Unpaid';
    if (releaseSearchResults) {
      releaseSearchResults.innerHTML = '<p class="text-muted text-center">Enter search criteria and click Search</p>';
    }

    if (releaseControlNumber) releaseControlNumber.value = '';
    if (releaseTargetUnit) releaseTargetUnit.value = '';
    if (customTargetUnit) customTargetUnit.value = '';
    if (releasePurpose) releasePurpose.value = '';

    if (this.isCPO) {
      controlNumberGroup?.classList.add('hidden');
      releasePurposeGroup?.classList.remove('hidden');
    } else {
      controlNumberGroup?.classList.remove('hidden');
      releasePurposeGroup?.classList.add('hidden');
    }

    releaseStep1?.classList.remove('hidden');
    releaseStep2?.classList.add('hidden');

    this.updateReleaseSelectedDisplay();
    releaseModal.classList.add('active');
  },

  handleTargetUnitChange() {
    const target = document.getElementById('releaseTargetUnit').value;
    const customGroup = document.getElementById('customTargetUnitGroup');
    if (!customGroup) return;

    if (target === 'Others') {
      customGroup.classList.remove('hidden');
    } else {
      customGroup.classList.add('hidden');
      document.getElementById('customTargetUnit').value = '';
    }
  },

  async searchForRelease() {
    const requestId = ++this._releaseSearchSeq;
    const searchTerm = document.getElementById('releaseSearchInput').value.trim();
    const statusFilter = document.getElementById('releaseStatusFilter').value;
    const categoryFilter = document.getElementById('releaseCategoryFilter').value;

    if (!searchTerm && statusFilter === 'All' && categoryFilter === 'All') {
      Utils.showToast('Please enter search criteria', 'warning');
      return;
    }

    const container = document.getElementById('releaseSearchResults');
    container.innerHTML = '<p class="text-muted text-center">Searching...</p>';

    this.showLoading(true);

    try {
      // Required order
      const years = ['2026', '2025', '2024', '2023', '<2023'];
      let all = [];

      // Parallelize requests for speed
      const promises = years.map(year => {
        const filters = { searchTerm, status: statusFilter, category: categoryFilter };
        return API.getVouchers(year, filters, 1, 100)
          .then(res => ({ year, res }))
          .catch(e => ({ year, res: { success: false } }));
      });

      const results = await Promise.all(promises);
      if (requestId !== this._releaseSearchSeq) return;

      results.forEach(({ year, res }) => {
        if (res.success && res.vouchers) {
          res.vouchers.forEach(v => all.push({ ...v, sourceYear: year }));
        }
      });

      this.releaseSearchResults = all;
      this.renderReleaseSearchResults();
      this.updateReleaseSelectedDisplay();

    } catch (e) {
      console.error('searchForRelease error:', e);
      container.innerHTML = '<p class="text-danger text-center">Search failed.</p>';
    } finally {
      this.showLoading(false);
    }
  },

  async globalSearch() {
    const requestId = ++this._searchSeq;
    const term = (this.filters.searchTerm || '').trim();
    if (!term) {
      this.isGlobalSearchMode = false;
      return this.loadVouchers();
    }

    this.isGlobalSearchMode = true;
    this.showLoading(true);

    try {
      const years = this.globalSearchYears || ['2026', '2025', '2024', '2023', '<2023'];
      const allResults = [];

      const filters = {
        searchTerm: term,
        status: this.filters.status,
        category: this.filters.category,
        pmtMonth: this.filters.pmtMonth,
        dateFrom: this.filters.dateFrom,
        dateTo: this.filters.dateTo,
        amountMin: this.filters.amountMin,
        amountMax: this.filters.amountMax,
        release: this.filters.release
      };

      // Fetch current-year first for fast first paint, then append archives.
      const [firstYear, ...archiveYears] = years;
      const firstRes = await API.getVouchers(firstYear, filters, 1, 200);
      if (requestId !== this._searchSeq) return;
      if (firstRes.success && Array.isArray(firstRes.vouchers)) {
        allResults.push(...firstRes.vouchers.map(v => ({ ...v, sourceYear: firstYear })));
      }

      const archiveResults = await Promise.all(
        archiveYears.map(year =>
          API.getVouchers(year, filters, 1, 200)
            .then(res => res.success && Array.isArray(res.vouchers)
              ? res.vouchers.map(v => ({ ...v, sourceYear: year }))
              : []
            )
            .catch(() => [])
        )
      );
      if (requestId !== this._searchSeq) return;
      archiveResults.forEach(batch => allResults.push(...batch));

      this.vouchers = allResults;
      this.totalCount = allResults.length;
      this.totalPages = 1;
      this.currentPage = 1;

      this.renderGlobalSearchResults();
    } catch (e) {
      console.error('Global search error:', e);
      Utils.showToast('Global search failed', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  renderGlobalSearchResults() {
    const container = document.getElementById('vouchersList');
    if (!container) return;

    const countEl = document.getElementById('voucherCount');
    if (countEl) countEl.textContent = `Found ${this.vouchers.length} voucher(s) across all years`;

    if (!this.vouchers.length) {
      container.innerHTML = Components.getEmptyState('No vouchers found across all years', 'fa-search');
      document.getElementById('paginationContainer').innerHTML = '';
      return;
    }

    const user = Auth.getUser();
    const canSelect = user && [
      CONFIG.ROLES.PAYABLE_STAFF,
      CONFIG.ROLES.PAYABLE_HEAD,
      CONFIG.ROLES.ADMIN,
      CONFIG.ROLES.CPO
    ].includes(user.role);

    let html = `
      <div class="table-container">
        <table>
          <thead>
            <tr>
              ${canSelect ? `<th><input type="checkbox" id="selectAll" onchange="Vouchers.toggleSelectAll()"></th>` : ''}
              <th>Year</th>
              <th>Voucher No.</th>
              <th>Payee</th>
              <th>Particular</th>
              <th>Gross Amount</th>
              <th>Category</th>
              <th>Control No.</th>
              <th>Status</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
    `;

    this.vouchers.forEach(v => {
      const year = v.sourceYear || '2026';
      const is2026 = year === '2026';
      const yearBadgeClass = is2026 ? 'badge-paid' : 'badge-unpaid';

      html += `
        <tr>
          ${canSelect ? `
            <td>
              ${is2026
            ? `<input type="checkbox" class="voucher-checkbox" value="${v.rowIndex}" onchange="Vouchers.updateSelection()">`
            : `<input type="checkbox" disabled title="Archive records are not selectable">`
          }
            </td>
          ` : ''}

          <td><span class="badge ${yearBadgeClass}">${year}</span></td>
          <td><strong>${v.accountOrMail || '-'}</strong></td>
          <td title="${v.payee || ''}">${Utils.truncate(v.payee || '-', 20)}</td>
          <td title="${v.particular || ''}"><div class="particular-cell">${v.particular || '-'}</div></td>
          <td>${Utils.formatCurrency(v.grossAmount || 0)}</td>
          <td>${v.categories || '-'}</td>
          <td>${v.controlNumber || '-'}</td>
          <td>${Utils.getStatusBadge(v.status || '')}</td>

          <td>
            <div class="action-buttons">
              <button class="btn btn-sm btn-secondary"
                      onclick="Vouchers.viewVoucherByYear(${v.rowIndex}, '${year}')"
                      title="View">
                <i class="fas fa-eye"></i>
              </button>

              ${is2026 ? this.getActionButtons(v) : ''}
            </div>
          </td>
        </tr>
      `;
    });

    html += `</tbody></table></div>`;
    container.innerHTML = html;

    // No pagination in global mode
    const pag = document.getElementById('paginationContainer');
    if (pag) pag.innerHTML = '';
  },

  async viewVoucherByYear(rowIndex, year) {
    // If year is 2026, you can reuse your existing viewVoucher
    if (!year || year === '2026') {
      return this.viewVoucher(rowIndex);
    }

    this.showLoading(true);
    try {
      const result = await API.getVoucherByRow(rowIndex, year);
      if (!result.success || !result.voucher) {
        Utils.showToast('Voucher not found', 'error');
        return;
      }

      // Temporarily show using the same modal layout
      const voucher = result.voucher;

      const modal = document.getElementById('viewVoucherModal');
      const content = document.getElementById('viewVoucherContent');

      content.innerHTML = `
        <div class="voucher-details">
          <div class="detail-group full-width">
            <label>Source Year</label>
            <div><strong>${year}</strong></div>
          </div>

          <div class="detail-row">
            <div class="detail-group">
              <label>Status</label>
              <div>${Utils.getStatusBadge(voucher.status)}</div>
            </div>
            <div class="detail-group">
              <label>Payment Month</label>
              <div>${voucher.pmtMonth || '-'}</div>
            </div>
          </div>

          <div class="detail-row">
            <div class="detail-group">
              <label>Voucher Number</label>
              <div><strong>${voucher.accountOrMail || '-'}</strong></div>
            </div>
            <div class="detail-group">
              <label>Control Number</label>
              <div>${voucher.controlNumber || '-'}</div>
            </div>
          </div>

          <div class="detail-group full-width">
            <label>Payee</label>
            <div><strong>${voucher.payee || '-'}</strong></div>
          </div>

          <div class="detail-group full-width">
            <label>Particular</label>
            <div>${voucher.particular || '-'}</div>
          </div>

          <div class="detail-row">
            <div class="detail-group">
              <label>Gross Amount</label>
              <div class="total-amount">${Utils.formatCurrency(voucher.grossAmount)}</div>
            </div>
            <div class="detail-group">
              <label>Date</label>
              <div>${Utils.formatDate(voucher.date)}</div>
            </div>
          </div>
        </div>
      `;

      modal.classList.add('active');
    } catch (e) {
      console.error('viewVoucherByYear error:', e);
      Utils.showToast('Error loading voucher', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  // Add this NEW function after viewVoucherByYear()
  async get2026VoucherForAction(rowIndex) {
    // In normal mode, use current page cache first
    if (!this.isGlobalSearchMode) {
      const local = this.vouchers.find(v => v.rowIndex === rowIndex);
      if (local) return { success: true, voucher: local };
    }

    // In global mode, always fetch the authoritative 2026 row
    try {
      const res = await API.getVoucherByRow(rowIndex, '2026');
      return res.success
        ? { success: true, voucher: res.voucher }
        : { success: false, error: res.error || 'Voucher not found' };
    } catch (e) {
      console.error('get2026VoucherForAction error:', e);
      return { success: false, error: 'Error fetching voucher' };
    }
  },


  clearFilters() {
    const statusFilter = document.getElementById('statusFilter');
    const categoryFilter = document.getElementById('categoryFilter');
    const searchInput = document.getElementById('searchInput');
    const pmtMonthFilter = document.getElementById('pmtMonthFilter');
    const dateFromFilter = document.getElementById('dateFromFilter');
    const dateToFilter = document.getElementById('dateToFilter');
    const amountMinFilter = document.getElementById('amountMinFilter');
    const amountMaxFilter = document.getElementById('amountMaxFilter');
    const releaseFilter = document.getElementById('releaseFilter');

    if (statusFilter) statusFilter.value = 'All';
    if (categoryFilter) categoryFilter.value = 'All';
    if (searchInput) searchInput.value = '';
    if (pmtMonthFilter) pmtMonthFilter.value = 'All';
    if (dateFromFilter) dateFromFilter.value = '';
    if (dateToFilter) dateToFilter.value = '';
    if (amountMinFilter) amountMinFilter.value = '';
    if (amountMaxFilter) amountMaxFilter.value = '';
    if (releaseFilter) releaseFilter.value = 'All';

    this.filters = {
      status: 'All',
      category: 'All',
      searchTerm: '',
      pmtMonth: 'All',
      dateFrom: '',
      dateTo: '',
      amountMin: '',
      amountMax: '',
      release: 'All'
    };

    this.isGlobalSearchMode = false;
    this.currentPage = 1;
    this.loadVouchers();
  },

  clearSearch() {
    const searchInput = document.getElementById('searchInput');
    if (searchInput) searchInput.value = '';
    this.filters.searchTerm = '';
    this.isGlobalSearchMode = false;
    this.applyFilters();
  },

  renderReleaseSearchResults() {
    const container = document.getElementById('releaseSearchResults');
    if (!container) return;

    const results = this.releaseSearchResults || [];
    if (!results.length) {
      container.innerHTML = '<p class="text-muted text-center">No vouchers found</p>';
      return;
    }

    let html = `
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th style="width:30px;"></th>
              <th>Voucher No.</th>
              <th>Payee</th>
              <th>Amount</th>
              <th>Status</th>
              <th>Year</th>
              <th>Control No.</th>
            </tr>
          </thead>
          <tbody>
    `;

    results.forEach((v, i) => {
      const isSelected = this.selectedForRelease.some(x => x.rowIndex === v.rowIndex && x.sourceYear === v.sourceYear);
      const hasCN = v.controlNumber && String(v.controlNumber).trim() !== '';

      // rule: only 2026 selectable
      let disabled = (v.sourceYear !== '2026');

      // payable unit cannot release already released vouchers; CPO can proceed
      if (!this.isCPO && hasCN) disabled = true;

      const selectedStyle = isSelected ? 'background-color: rgba(40, 167, 69, 0.1); border-left: 3px solid var(--success-color);' : '';
      html += `
    <tr style="${disabled ? 'opacity:0.6;cursor:not-allowed;' : 'cursor:pointer;'} ${selectedStyle}"
        onclick="${disabled ? '' : `Vouchers.toggleReleaseSelection(${i})`}">
          <td>
            <input type="checkbox"
              ${isSelected ? 'checked' : ''}
              ${disabled ? 'disabled' : ''}
              onclick="event.stopPropagation(); ${disabled ? '' : `Vouchers.toggleReleaseSelection(${i})`}">
          </td>
          <td><strong>${v.accountOrMail || '-'}</strong></td>
          <td title="${v.payee || ''}">${Utils.truncate(v.payee || '', 20)}</td>
          <td>${Utils.formatCurrency(v.grossAmount || 0)}</td>
          <td>${Utils.getStatusBadge(v.status || '')}</td>
          <td>${v.sourceYear}</td>
          <td>${v.controlNumber || '-'}</td>
        </tr>
      `;
    });

    html += '</tbody></table></div>';
    container.innerHTML = html;

  },

  toggleReleaseSelection(index) {
    const v = this.releaseSearchResults[index];
    if (!v) return;

    if (v.sourceYear !== '2026') return;
    const hasCN = v.controlNumber && String(v.controlNumber).trim() !== '';
    if (!this.isCPO && hasCN) {
      Utils.showToast('Already released voucher cannot be re-released without approval.', 'warning');
      return;
    }

    const pos = this.selectedForRelease.findIndex(x => x.rowIndex === v.rowIndex && x.sourceYear === v.sourceYear);
    if (pos >= 0) this.selectedForRelease.splice(pos, 1);
    else this.selectedForRelease.push(v);

    this.renderReleaseSearchResults();
    this.updateReleaseSelectedDisplay();
  },

  updateReleaseSelectedDisplay() {
    const wrap = document.getElementById('releaseSelectedList');
    const countEl = document.getElementById('releaseSelectedCount');
    const tagsEl = document.getElementById('releaseSelectedTags');
    const totalEl = document.getElementById('releaseSelectedTotal');

    if (!wrap || !countEl || !tagsEl || !totalEl) return;

    if (!this.selectedForRelease.length) {
      wrap.classList.add('hidden');
      countEl.textContent = '0';
      tagsEl.innerHTML = '';
      totalEl.textContent = '₦0.00';
      return;
    }

    wrap.classList.remove('hidden');
    countEl.textContent = this.selectedForRelease.length;

    let total = 0;
    tagsEl.innerHTML = this.selectedForRelease.map((v, idx) => {
      total += Number(v.grossAmount || 0);
      return `
        <span class="release-tag">
          ${v.accountOrMail || v.payee || 'Voucher'}
          <span class="remove-tag" onclick="event.stopPropagation(); Vouchers.removeFromReleaseSelection(${idx})">
            <i class="fas fa-times"></i>
          </span>
        </span>
      `;
    }).join('');

    totalEl.textContent = Utils.formatCurrency(total);
  },

  removeFromReleaseSelection(idx) {
    this.selectedForRelease.splice(idx, 1);
    this.renderReleaseSearchResults();
    this.updateReleaseSelectedDisplay();
  },

  proceedToReleaseStep2() {
    if (!this.selectedForRelease.length) {
      Utils.showToast('Please select at least one voucher to release', 'warning');
      return;
    }

    // Payable unit cannot proceed with already released vouchers
    if (!this.isCPO) {
      const already = this.selectedForRelease.filter(v => v.controlNumber && String(v.controlNumber).trim() !== '');
      if (already.length) return Utils.showToast('Some selected vouchers are already released. Remove them from selections.', 'error');
    }

    // summary
    let total = 0;
    const list = this.selectedForRelease.map(v => {
      total += Number(v.grossAmount || 0);
      return `<span class="release-tag">${v.accountOrMail || '-'}</span>`;
    }).join('');

    const summaryBox = document.getElementById('releaseSummaryBox');
    if (summaryBox) {
      summaryBox.innerHTML = `
        <h5><i class="fas fa-clipboard-list"></i> Release Summary</h5>
        <p><strong>Number of Vouchers:</strong> ${this.selectedForRelease.length}</p>
        <p><strong>Total Amount:</strong> ${Utils.formatCurrency(total)}</p>
        <p><strong>Voucher Numbers:</strong></p>
        <div class="release-tags">${list}</div>
      `;
    }

    document.getElementById('releaseStep1').classList.add('hidden');
    document.getElementById('releaseStep2').classList.remove('hidden');
  },

  backToReleaseStep1() {
    document.getElementById('releaseStep1').classList.remove('hidden');
    document.getElementById('releaseStep2').classList.add('hidden');
  },

  async submitRelease() {
    let targetUnit = document.getElementById('releaseTargetUnit').value;
    if (!targetUnit) return Utils.showToast('Please select a target unit', 'error');

    if (targetUnit === 'Others') {
      targetUnit = document.getElementById('customTargetUnit').value.trim();
      if (!targetUnit) return Utils.showToast('Please specify the target unit', 'error');
    }

    if (!this.selectedForRelease.length) return Utils.showToast('No vouchers selected', 'error');

    let controlNumber = '';
    let purpose = '';

    if (this.isCPO) {
      purpose = document.getElementById('releasePurpose').value.trim();
      if (!purpose) return Utils.showToast('Purpose is required for CPO release', 'error');

      if (this.isCPO) {
        const cnSet = new Set(this.selectedForRelease.map(v => (v.controlNumber || '').trim()).filter(Boolean));
        if (cnSet.size !== 1) {
          Utils.showToast('CPO release requires all selected vouchers to have the SAME Control Number.', 'error');
          return;
        }
        controlNumber = [...cnSet][0];
      }

      // CPO uses existing CN from vouchers
      if (this.selectedForRelease[0].controlNumber) controlNumber = this.selectedForRelease[0].controlNumber;
    } else {
      controlNumber = document.getElementById('releaseControlNumber').value.trim();
      if (!controlNumber) return Utils.showToast('Please enter or generate a control number', 'error');
    }

    const total = this.selectedForRelease.reduce((sum, v) => sum + Number(v.grossAmount || 0), 0);

    const confirmed = await Utils.confirm(
      `Release ${this.selectedForRelease.length} voucher(s) to ${targetUnit}?\n\nTotal: ${Utils.formatCurrency(total)}\n${this.isCPO ? 'Purpose: ' + purpose : 'Control No: ' + controlNumber}`,
      'Confirm Release'
    );
    if (!confirmed) return;

    this.showLoading(true);

    try {
      const rowIndexes = this.selectedForRelease.map(v => v.rowIndex);

      // Call backend release router (recommended)
      const result = await API.post('releaseVouchers', {
        rowIndexes,
        controlNumber,
        targetUnit,
        purpose,
        isCPORelease: this.isCPO
      });

      if (result.success) {
        Utils.showToast(result.message || 'Released successfully', 'success');
        this.closeModal('releaseModal');
        this.selectedForRelease = [];
        this.selectedVouchers = [];
        await this.loadVouchers();
      } else {
        Utils.showToast(result.error || 'Release failed', 'error');
      }
    } catch (e) {
      console.error('submitRelease error:', e);
      Utils.showToast('Error releasing vouchers', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  // ===== Filter Toggle =====
  toggleFilters() {
    const filtersBar = document.getElementById('filtersBar');
    const toggleBtn = document.getElementById('filterToggleBtn');
    const toggleText = document.getElementById('filterToggleText');

    if (!filtersBar) return;

    const isCollapsed = filtersBar.classList.contains('collapsed');

    if (isCollapsed) {
      filtersBar.classList.remove('collapsed');
      toggleBtn.classList.add('expanded');
      toggleText.textContent = 'Hide Filters';
    } else {
      filtersBar.classList.add('collapsed');
      toggleBtn.classList.remove('expanded');
      toggleText.textContent = 'Show Filters';
    }
  },

  updateActiveFiltersText() {
    const textEl = document.getElementById('activeFiltersText');
    if (!textEl) return;

    const activeFilters = [];

    if (this.filters.status && this.filters.status !== 'All') {
      activeFilters.push(this.filters.status);
    }
    if (this.filters.category && this.filters.category !== 'All') {
      activeFilters.push(this.filters.category);
    }
    if (this.filters.searchTerm) {
      activeFilters.push(`"${this.filters.searchTerm}"`);
    }
    if (this.filters.pmtMonth && this.filters.pmtMonth !== 'All') {
      activeFilters.push(`Pmt: ${this.filters.pmtMonth}`);
    }
    if (this.filters.dateFrom) {
      activeFilters.push(`From: ${this.filters.dateFrom}`);
    }
    if (this.filters.dateTo) {
      activeFilters.push(`To: ${this.filters.dateTo}`);
    }
    if (this.filters.amountMin) {
      activeFilters.push(`Min: ₦${this.filters.amountMin}`);
    }
    if (this.filters.amountMax) {
      activeFilters.push(`Max: ₦${this.filters.amountMax}`);
    }
    if (this.filters.release && this.filters.release !== 'All') {
      activeFilters.push(this.filters.release);
    }

    if (activeFilters.length === 0) {
      textEl.innerHTML = 'Filters: <span class="active-filter">None active</span>';
    } else {
      textEl.innerHTML = 'Active: <span class="active-filter">' + activeFilters.join(', ') + '</span>';
    }
  },

  // ===== Utilities =====
  setupEventListeners() {
    const debouncedFilter = Utils.debounce(() => this.applyFilters({ skipLoading: true }), 320);
    const debouncedLocalSearch = Utils.debounce(
      () => this.applyFilters({ skipGlobal: true, skipLoading: true }),
      450
    );

    document.getElementById('statusFilter')?.addEventListener('change', () => this.applyFilters());
    document.getElementById('categoryFilter')?.addEventListener('change', () => this.applyFilters());
    document.getElementById('pmtMonthFilter')?.addEventListener('change', () => this.applyFilters());
    document.getElementById('dateFromFilter')?.addEventListener('change', () => this.applyFilters());
    document.getElementById('dateToFilter')?.addEventListener('change', () => this.applyFilters());
    document.getElementById('releaseFilter')?.addEventListener('change', () => this.applyFilters());

    document.getElementById('amountMinFilter')?.addEventListener('input', debouncedFilter);
    document.getElementById('amountMaxFilter')?.addEventListener('input', debouncedFilter);

    const searchInput = document.getElementById('searchInput');
    if (searchInput) {
      searchInput.addEventListener('input', debouncedLocalSearch);
      searchInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') this.runSearch();
      });
    }

    document.getElementById('newVoucherBtn')?.addEventListener('click', () => this.openVoucherForm());
    document.getElementById('lookupBtn')?.addEventListener('click', () => this.openLookupModal());
    document.getElementById('batchStatusBtn')?.addEventListener('click', () => this.openBatchStatusModal());
    document.getElementById('releaseToUnitBtn')?.addEventListener('click', () => this.openReleaseModal());

    document.getElementById('logoutBtn')?.addEventListener('click', () => Auth.logout());
    document.getElementById('menuToggle')?.addEventListener('click', () => {
      document.getElementById('sidebar')?.classList.toggle('active');
    });

    document.querySelectorAll('.modal-overlay').forEach(m => {
      m.addEventListener('click', (e) => {
        if (e.target === m) m.classList.remove('active');
      });
    });

    document.getElementById('voucherForm')?.addEventListener('submit', (e) => {
      e.preventDefault();
      this.saveVoucher();
    });

    document.querySelectorAll('input[name="oldVoucherAvailable"]').forEach(radio => {
      radio.addEventListener('change', () => this.applyOldVoucherAvailability());
    });

    const recalcNet = () => this.recalculateVoucherNet();

    ['formGrossAmount', 'formVat', 'formWht', 'formStampDuty'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('input', recalcNet);
    });
  },

  async saveStatus() {
    if (!this.selectedVoucher) {
      Utils.showToast('No voucher selected', 'error');
      return;
    }

    const status = document.getElementById('newStatus')?.value || 'Unpaid';
    const pmtMonth = document.getElementById('newPmtMonth')?.value || '';
    const user = Auth.getUser();
    const canSetMonth = user && (user.role === CONFIG.ROLES.CPO || user.role === CONFIG.ROLES.ADMIN);

    if (canSetMonth && status === 'Paid' && !pmtMonth) {
      Utils.showToast('Payment month is required when setting status to Paid', 'error');
      return;
    }

    this.showLoading(true);

    try {
      // Optimistic update
      const originalStatus = this.selectedVoucher.status;
      const originalPmtMonth = this.selectedVoucher.pmtMonth;

      this.selectedVoucher.status = status;
      this.selectedVoucher.pmtMonth = pmtMonth;
      this.renderVoucherList();
      this.closeModal('statusModal');

      const result = typeof API.updateStatus === 'function'
        ? await API.updateStatus(this.selectedVoucher.rowIndex, status, pmtMonth)
        : await API.post('updateStatus', {
          rowIndex: this.selectedVoucher.rowIndex,
          status,
          pmtMonth
        });

      if (result.success) {
        Utils.showToast(result.message || 'Status updated successfully', 'success');
        // Background SWR handles fresh data sync, no block on UI needed
      } else {
        // Revert 
        this.selectedVoucher.status = originalStatus;
        this.selectedVoucher.pmtMonth = originalPmtMonth;
        this.renderVoucherList();
        Utils.showToast(result.error || 'Failed to update status', 'error');
      }
    } catch (e) {
      console.error('saveStatus error:', e);
      Utils.showToast('Error updating status', 'error');
    } finally {
      this.showLoading(false);
    }
  },

  async generateControlNumber() {
    const targetSelect = document.getElementById('releaseTargetUnit');
    const target = targetSelect?.value || '';

    if (!target) {
      Utils.showToast('Please select a target unit first', 'warning');
      return;
    }

    if (target === 'Others') {
      Utils.showToast('Please specify the target unit name for "Others"', 'warning');
      return;
    }

    this.showLoading(true);

    try {
      console.log('=== generateControlNumber START ===');
      console.log('Target Unit:', target);

      // Try GET first
      let result;

      if (typeof API.get === 'function') {
        console.log('Calling API.get("getNextControlNumber", { targetUnit: "' + target + '" })');
        result = await API.get('getNextControlNumber', { targetUnit: target });
      } else if (typeof API.getNextControlNumber === 'function') {
        console.log('Calling API.getNextControlNumber("' + target + '")');
        result = await API.getNextControlNumber(target);
      } else {
        // Fallback to POST
        console.log('Calling API.post("getNextControlNumber", { targetUnit: "' + target + '" })');
        result = await API.post('getNextControlNumber', { targetUnit: target });
      }

      console.log('API Response:', JSON.stringify(result));
      console.log('=== generateControlNumber END ===');

      if (result && result.success && result.controlNumber) {
        document.getElementById('releaseControlNumber').value = result.controlNumber;
        Utils.showToast('Control number generated: ' + result.controlNumber, 'success');
      } else {
        const errorMsg = (result && result.error) ? result.error : 'Failed to generate control number';
        Utils.showToast(errorMsg, 'error');
        console.error('generateControlNumber failed:', errorMsg);
      }
    } catch (e) {
      console.error('generateControlNumber exception:', e);
      Utils.showToast('Error generating control number: ' + e.message, 'error');
    } finally {
      this.showLoading(false);
    }
  },

  openBatchStatusModal() {
    const perms = this.getEffectivePermissions();
    if (!perms.canBatchUpdateStatus) {
      Utils.showToast('You are not authorized to batch update status', 'error');
      return;
    }

    const modal = document.getElementById('batchStatusModal');
    if (!modal) return;

    const cnInput = document.getElementById('batchControlNumber');
    const statusInput = document.getElementById('batchStatus');
    const monthInput = document.getElementById('batchPmtMonth');

    if (cnInput) cnInput.value = '';
    if (statusInput) statusInput.value = 'Paid';
    if (monthInput) monthInput.value = '';

    modal.classList.add('active');
  },

  toggleSelectAll() {
    const selectAll = document.getElementById('selectAll');
    const isChecked = !!(selectAll && selectAll.checked);

    document.querySelectorAll('.voucher-checkbox:not(:disabled)').forEach(cb => {
      cb.checked = isChecked;
    });

    this.updateSelection();
  },

  // ===== Payment Type Handling =====
  setupPaymentTypeHandlers() {
    if (this.paymentTypeHandlersReady) return;
    this.paymentTypeHandlersReady = true;

    const paymentTypes = document.querySelectorAll('input[name="paymentType"]');
    const otherPartGroup = document.getElementById('otherPartPaymentGroup');
    const contractSumGroup = document.getElementById('contractSumGroup');
    const contractSumRequired = document.getElementById('contractSumRequired');

    paymentTypes.forEach(radio => {
      radio.addEventListener('change', () => {
        const value = radio.value;

        if (value === 'otherPart') {
          otherPartGroup?.classList.remove('hidden');
        } else {
          otherPartGroup?.classList.add('hidden');
        }

        if (value === 'firstPart') {
          contractSumGroup?.classList.add('required-field');
          if (contractSumRequired) contractSumRequired.style.display = 'inline';
          const contractSumInput = document.getElementById('formContractSum');
          if (contractSumInput) contractSumInput.required = true;
        } else {
          contractSumGroup?.classList.remove('required-field');
          if (contractSumRequired) contractSumRequired.style.display = 'none';
          const contractSumInput = document.getElementById('formContractSum');
          if (contractSumInput) contractSumInput.required = false;
        }

        this.updateParticularHint(value);
        this.autoFillParticular(value);
      });
    });

    const otherPartType = document.getElementById('otherPartPaymentType');
    if (otherPartType) {
      otherPartType.addEventListener('change', () => {
        this.autoFillParticular('otherPart');
      });
    }
  },

  updateParticularHint(paymentType) {
    const hint = document.getElementById('particularHint');
    const note = document.getElementById('particularNote');

    const hints = {
      'lumpsum': '',
      'firstPart': '(Must start with "First Part-Payment of" or similar)',
      'balance': '(Must start with "Balance Payment of" or similar)',
      'otherPart': '(Must indicate which part-payment, e.g., "Second Part-Payment of")'
    };

    if (hint) hint.textContent = hints[paymentType] || '';

    if (note) {
      note.textContent = paymentType !== 'lumpsum'
        ? 'The particular will be validated based on your payment type selection.'
        : '';
    }
  },

  autoFillParticular(paymentType) {
    const particular = document.getElementById('formParticular');
    if (!particular || particular.value.trim() !== '') return;

    const prefixes = {
      'firstPart': 'First Part-Payment of ',
      'balance': 'Balance Payment of ',
      'otherPart': this.getOtherPartPrefix()
    };

    if (prefixes[paymentType]) {
      particular.value = prefixes[paymentType];
      particular.focus();
      particular.setSelectionRange(particular.value.length, particular.value.length);
    }
  },

  getOtherPartPrefix() {
    const type = document.getElementById('otherPartPaymentType')?.value;
    const prefixes = {
      'second': 'Second Part-Payment of ',
      'third': 'Third Part-Payment of ',
      'fourth': 'Fourth Part-Payment of ',
      'fifth': 'Fifth Part-Payment of '
    };
    return prefixes[type] || '';
  },

  validateParticular() {
    const paymentType = document.querySelector('input[name="paymentType"]:checked')?.value;
    const particular = document.getElementById('formParticular').value.trim().toLowerCase();

    if (paymentType === 'lumpsum' || !paymentType) return { valid: true };

    // First Part-Payment validation
    if (paymentType === 'firstPart') {
      const validPatterns = [
        /^first\s*part[-\s]?p(ay)?m(en)?t/i,
        /^1st\s*part[-\s]?p(ay)?m(en)?t/i,
        /^first\s*pt[-\s]?pmt/i,
        /^1st\s*pt[-\s]?pmt/i
      ];

      const isValid = validPatterns.some(pattern => pattern.test(particular));
      if (!isValid) {
        return {
          valid: false,
          error: 'For First Part-Payment, the particular must begin with "First Part-Payment of" (or similar like "1st pt-pmt of")'
        };
      }
    }

    // Balance Payment validation
    if (paymentType === 'balance') {
      const validPatterns = [
        /^bal(ance)?\s*p(ay)?m(en)?t/i,
        /^final\s*p(ay)?m(en)?t/i,
        /^fnl\s*p(ay)?m(en)?t/i,
        /^bal\s*pmt/i
      ];

      const isValid = validPatterns.some(pattern => pattern.test(particular));
      if (!isValid) {
        return {
          valid: false,
          error: 'For Balance Payment, the particular must begin with "Balance Payment of" (or similar like "Bal pmt of", "Final payment of")'
        };
      }
    }

    // Other Part-Payment validation
    if (paymentType === 'otherPart') {
      const otherType = document.getElementById('otherPartPaymentType')?.value;
      if (!otherType) {
        return { valid: false, error: 'Please select the part-payment type' };
      }

      if (otherType !== 'custom') {
        const typePatterns = {
          'second': /^(2nd|second)\s*part[-\s]?p(ay)?m(en)?t/i,
          'third': /^(3rd|third)\s*part[-\s]?p(ay)?m(en)?t/i,
          'fourth': /^(4th|fourth)\s*part[-\s]?p(ay)?m(en)?t/i,
          'fifth': /^(5th|fifth)\s*part[-\s]?p(ay)?m(en)?t/i
        };

        const pattern = typePatterns[otherType];
        if (pattern && !pattern.test(particular)) {
          return {
            valid: false,
            error: `For ${otherType} part-payment, the particular must begin with "${otherType} Part-Payment of" (or similar)`
          };
        }
      }
    }

    return { valid: true };
  },

  printVoucher() {
    window.print();
  },

  closeModal(id) {
    document.getElementById(id)?.classList.remove('active');
  },

  showLoading(show) {
    if (window.Components && typeof Components.setLoading === 'function') {
      Components.setLoading(show, this.isGlobalSearchMode ? 'Searching all years...' : 'Loading vouchers...');
      return;
    }
    const loader = document.getElementById('loadingOverlay');
    if (loader) loader.classList.toggle('hidden', !show);
  }

}; // <-- This closes the Vouchers object

window.Vouchers = Vouchers;
window.openVoucherModal = (voucher, rowIndex) => Vouchers.openVoucherForm(voucher);

document.addEventListener('DOMContentLoaded', () => Vouchers.init());
