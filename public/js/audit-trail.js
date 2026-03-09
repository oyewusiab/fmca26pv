const AuditTrailPage = {
  limit: 50,
  offset: 0,
  total: 0,

  async init() {
    try {
      this.showLoading(true);

      const ok = await Auth.requireAuth();
      if (!ok) return;

      const user = Auth.getUser();
      if (!user || user.role !== CONFIG.ROLES.ADMIN) {
        this.showLoading(false);
        alert('Unauthorized: Admin only');
        window.location.href = 'dashboard.html';
        return;
      }

      // Sidebar
      const sidebar = document.getElementById('sidebar');
      if (sidebar) sidebar.innerHTML = Components.getSidebar('auditTrail');

      // Mobile menu
      document.getElementById('menuToggle')?.addEventListener('click', () => {
        document.getElementById('sidebar')?.classList.toggle('active');
      });

      // Refresh
      document.getElementById('refreshAuditBtn')?.addEventListener('click', () => {
        this.offset = 0;
        this.load();
      });

      await this.load();

    } catch (e) {
      console.error('AuditTrail init error:', e);
      this.renderError('Audit page crashed: ' + (e.message || e));
    } finally {
      this.showLoading(false);
    }
  },

  async load() {
    try {
      this.showLoading(true);

      if (!API.getAuditTrail) {
        this.renderError('API.getAuditTrail is not defined. Add it to api.js.');
        return;
      }

      const res = await API.getAuditTrail(this.limit, this.offset);

      if (!res || typeof res !== 'object') {
        this.renderError('Invalid response from server.');
        return;
      }

      if (!res.success) {
        this.renderError(res.error || 'Failed to load audit trail.');
        return;
      }

      this.total = res.total || 0;
      this.renderTable(res.records || []);
      this.renderPagination();

    } catch (e) {
      console.error('AuditTrail load error:', e);
      this.renderError('Failed to load audit trail: ' + (e.message || e));
    } finally {
      this.showLoading(false);
    }
  },

  renderTable(records) {
    const container = document.getElementById('auditTrailContainer');
    if (!container) return;

    if (!records.length) {
      container.innerHTML = `<p class="text-muted text-center">No audit records yet.</p>`;
      return;
    }

    let html = `
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Timestamp</th>
              <th>User Email</th>
              <th>User Name</th>
              <th>Role</th>
              <th>Action</th>
              <th>Description</th>
              <th>Sheet</th>
              <th>Row</th>
            </tr>
          </thead>
          <tbody>
    `;

    records.forEach(r => {
      html += `
        <tr>
          <td>${r.timestamp || '-'}</td>
          <td>${r.email || '-'}</td>
          <td>${r.name || '-'}</td>
          <td>${r.role || '-'}</td>
          <td><strong>${r.action || '-'}</strong></td>
          <td>${r.description || '-'}</td>
          <td>${r.sheet || '-'}</td>
          <td>${r.rowIndex || '-'}</td>
        </tr>
      `;
    });

    html += `</tbody></table></div>`;
    container.innerHTML = html;
  },

  renderPagination() {
    const el = document.getElementById('auditPagination');
    if (!el) return;

    const currentPage = Math.floor(this.offset / this.limit) + 1;
    const totalPages = Math.max(1, Math.ceil(this.total / this.limit));

    if (totalPages <= 1) {
      el.innerHTML = '';
      return;
    }

    el.innerHTML = `
      <div class="pagination">
        <button class="btn btn-sm btn-secondary" ${currentPage === 1 ? 'disabled' : ''} onclick="AuditTrailPage.goPage(${currentPage - 1})">
          <i class="fas fa-chevron-left"></i>
        </button>

        <span style="padding:6px 12px;">Page ${currentPage} of ${totalPages} (Total: ${this.total})</span>

        <button class="btn btn-sm btn-secondary" ${currentPage === totalPages ? 'disabled' : ''} onclick="AuditTrailPage.goPage(${currentPage + 1})">
          <i class="fas fa-chevron-right"></i>
        </button>
      </div>
    `;
  },

  goPage(page) {
    this.offset = (page - 1) * this.limit;
    if (this.offset < 0) this.offset = 0;
    this.load();
  },

  renderError(msg) {
    const container = document.getElementById('auditTrailContainer');
    if (container) container.innerHTML = `<p class="text-danger text-center">${msg}</p>`;
  },

  showLoading(show) {
    document.getElementById('loadingOverlay')?.classList.toggle('hidden', !show);
  }
};

document.addEventListener('DOMContentLoaded', () => AuditTrailPage.init());