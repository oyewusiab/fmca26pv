/**
 * PAYABLE VOUCHER 2026 - Announcements Module
 */

const Announcements = {
  initializedComposer: false,
  availableUsers: [],
  activeAnnouncements: [],
  popupQueue: [],
  popupOpen: false,

  normalize(value) {
    return String(value || '').trim().toUpperCase();
  },

  isAdmin() {
    const role = this.normalize(Auth.getUser()?.role);
    return role === this.normalize(CONFIG.ROLES.ADMIN);
  },

  initComposer() {
    const editor = document.getElementById('announcementEditor');
    if (!editor || this.initializedComposer) return;
    this.initializedComposer = true;

    document.getElementById('annToAll')?.addEventListener('change', () => this.handleRecipientToggle('all'));
    document.getElementById('annToSpecific')?.addEventListener('change', () => this.handleRecipientToggle('specific'));
    document.getElementById('annUserSearch')?.addEventListener('input', () => this.renderUserOptions());
    document.getElementById('annExpiryPreset')?.addEventListener('change', () => this.handleExpiryPresetChange());
    document.getElementById('sendAnnouncementBtn')?.addEventListener('click', () => this.submitAnnouncement());
    document.getElementById('clearAnnouncementBtn')?.addEventListener('click', () => this.clearComposer());
    document.getElementById('refreshAnnouncementsBtn')?.addEventListener('click', () => this.loadAnnouncements());

    document.querySelectorAll('#announcementEditorToolbar .toolbar-btn').forEach((button) => {
      button.addEventListener('click', () => {
        const cmd = button.getAttribute('data-cmd');
        const value = button.getAttribute('data-value');
        editor.focus();
        document.execCommand(cmd, false, value || null);
      });
    });

    this.handleRecipientToggle();
    this.handleExpiryPresetChange();
    this.loadUserOptions();
  },

  handleRecipientToggle(source = null) {
    const toAll = document.getElementById('annToAll');
    const toSpecific = document.getElementById('annToSpecific');
    const specificWrap = document.getElementById('annSpecificUsersWrap');
    if (!toAll || !toSpecific || !specificWrap) return;

    if (source === 'all' && toAll.checked) {
      toSpecific.checked = false;
    }
    if (source === 'specific' && toSpecific.checked) {
      toAll.checked = false;
    }

    if (!toAll.checked && !toSpecific.checked) {
      toAll.checked = true;
    }

    specificWrap.style.display = toSpecific.checked ? 'block' : 'none';
  },

  handleExpiryPresetChange() {
    const preset = document.getElementById('annExpiryPreset');
    const custom = document.getElementById('annExpiryCustom');
    if (!preset || !custom) return;

    const useCustom = preset.value === 'custom';
    custom.style.display = useCustom ? 'block' : 'none';

    if (useCustom && !custom.value) {
      const nextHour = new Date(Date.now() + (60 * 60 * 1000));
      nextHour.setMinutes(0, 0, 0);
      custom.value = nextHour.toISOString().slice(0, 16);
    }
  },

  async loadUserOptions() {
    const list = document.getElementById('annSpecificUsersList');
    if (!list) return;

    list.innerHTML = '<p class="text-muted">Loading users...</p>';

    try {
      const res = await API.getUsers();
      if (!res.success) {
        list.innerHTML = '<p class="text-muted">Unable to load users. You can still send to all users.</p>';
        this.availableUsers = [];
        return;
      }

      this.availableUsers = Array.isArray(res.users) ? res.users : [];
      this.renderUserOptions();
    } catch (error) {
      console.error('Failed to load users for announcement selection:', error);
      list.innerHTML = '<p class="text-muted">Unable to load users.</p>';
      this.availableUsers = [];
    }
  },

  renderUserOptions() {
    const list = document.getElementById('annSpecificUsersList');
    if (!list) return;

    const search = this.normalize(document.getElementById('annUserSearch')?.value || '');
    const selected = new Set(
      Array.from(list.querySelectorAll('input[type="checkbox"]:checked')).map((input) => input.value)
    );

    const filtered = this.availableUsers.filter((user) => {
      if (!search) return true;
      const target = this.normalize(`${user.name} ${user.email} ${user.role}`);
      return target.includes(search);
    });

    if (!filtered.length) {
      list.innerHTML = '<p class="text-muted">No users match your search.</p>';
      return;
    }

    list.innerHTML = filtered.map((user) => {
      const checked = selected.has(user.email) ? 'checked' : '';
      const safeName = this.escapeHtml(user.name || '-');
      const safeEmail = this.escapeHtml(user.email || '-');
      const safeRole = this.escapeHtml(user.role || '-');
      return `
        <label class="announcement-user-option ${checked ? 'is-selected' : ''}">
          <input type="checkbox" value="${safeEmail}" ${checked}>
          <span>
            <strong>${safeName}</strong>
            <div class="meta">${safeEmail} - ${safeRole}</div>
          </span>
        </label>
      `;
    }).join('');

    list.querySelectorAll('.announcement-user-option input[type="checkbox"]').forEach((input) => {
      input.addEventListener('change', () => {
        input.closest('.announcement-user-option')?.classList.toggle('is-selected', input.checked);
      });
    });
  },

  escapeHtml(text) {
    return String(text || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  },

  collectComposerData() {
    const editor = document.getElementById('announcementEditor');
    const locations = [];
    if (document.getElementById('annLocationLogin')?.checked) locations.push('LOGIN');
    if (document.getElementById('annLocationDashboard')?.checked) locations.push('DASHBOARD');

    if (!locations.length) {
      throw new Error('Please select where the message should be displayed.');
    }

    const html = (editor?.innerHTML || '').trim();
    const text = (editor?.textContent || '').trim();
    if (!text) {
      throw new Error('Please enter the announcement message.');
    }

    let targets = ['ALL'];
    const toSpecific = document.getElementById('annToSpecific')?.checked;
    if (toSpecific) {
      const selected = Array.from(
        document.querySelectorAll('#annSpecificUsersList input[type="checkbox"]:checked')
      ).map((input) => input.value);

      if (!selected.length) {
        throw new Error('Select at least one user or switch to "All Users".');
      }
      targets = selected;
    }

    const preset = document.getElementById('annExpiryPreset')?.value || '24h';
    let expiresAt;
    if (preset === 'custom') {
      const custom = document.getElementById('annExpiryCustom')?.value;
      if (!custom) throw new Error('Please specify custom expiration date/time.');
      const dt = new Date(custom);
      if (Number.isNaN(dt.getTime())) throw new Error('Invalid custom expiration date/time.');
      expiresAt = dt.toISOString();
    } else {
      const hours = parseInt(preset.replace('h', ''), 10);
      expiresAt = new Date(Date.now() + (hours * 3600000)).toISOString();
    }

    const allowDismiss = (document.getElementById('annAllowDismiss')?.value || 'yes') === 'yes';

    return {
      message: html,
      locations,
      targets,
      expiresAt,
      allowDismiss
    };
  },

  async submitAnnouncement() {
    if (!this.isAdmin()) {
      Utils.showToast('Only ADMIN can send announcements.', 'error');
      return;
    }

    try {
      const payload = this.collectComposerData();
      const res = await API.createAnnouncement(payload);
      if (!res.success) {
        Utils.showToast(res.error || 'Failed to create announcement.', 'error');
        return;
      }

      Utils.showToast(res.message || 'Announcement sent successfully.', 'success');
      this.clearComposer();
      await this.loadAnnouncements();
    } catch (error) {
      Utils.showToast(error.message || 'Failed to create announcement.', 'error');
    }
  },

  clearComposer() {
    const editor = document.getElementById('announcementEditor');
    if (editor) editor.innerHTML = '';

    const toAll = document.getElementById('annToAll');
    const toSpecific = document.getElementById('annToSpecific');
    if (toAll) toAll.checked = true;
    if (toSpecific) toSpecific.checked = false;

    const annLocationLogin = document.getElementById('annLocationLogin');
    const annLocationDashboard = document.getElementById('annLocationDashboard');
    if (annLocationLogin) annLocationLogin.checked = false;
    if (annLocationDashboard) annLocationDashboard.checked = true;

    const preset = document.getElementById('annExpiryPreset');
    if (preset) preset.value = '24h';
    const allowDismiss = document.getElementById('annAllowDismiss');
    if (allowDismiss) allowDismiss.value = 'yes';

    document.querySelectorAll('#annSpecificUsersList input[type="checkbox"]').forEach((input) => {
      input.checked = false;
    });

    this.handleRecipientToggle();
    this.handleExpiryPresetChange();
    this.renderUserOptions();
  },

  async loadAnnouncements() {
    const container = document.getElementById('announcementsListContent');

    try {
      const res = await API.getActiveAnnouncements();
      if (!res.success) {
        if (container) {
          container.innerHTML = `<div class="alert alert-danger">${this.escapeHtml(res.error || 'Failed to load announcements.')}</div>`;
        }
        return;
      }

      this.activeAnnouncements = Array.isArray(res.announcements) ? res.announcements : [];
      const countEl = document.getElementById('hubAnnouncementCount');
      if (countEl) countEl.textContent = this.activeAnnouncements.length;

      if (!container) return;

      if (!this.activeAnnouncements.length) {
        container.innerHTML = '<p class="text-muted">No active announcements.</p>';
        return;
      }

      container.innerHTML = this.activeAnnouncements.map((announcement) => this.renderAnnouncementCard(announcement)).join('');
    } catch (error) {
      console.error('Failed to load announcements:', error);
      if (container) container.innerHTML = '<div class="alert alert-danger">Failed to load announcements.</div>';
    }
  },

  renderAnnouncementCard(announcement) {
    const locations = (announcement.locations || []).map((loc) => this.escapeHtml(loc)).join(', ') || 'N/A';
    const targets = (announcement.targets || []).map((target) => this.escapeHtml(target)).join(', ') || 'ALL';
    const expires = announcement.expiresAt ? Utils.formatDateTime(announcement.expiresAt) : '-';
    const createdBy = announcement.createdBy ? this.escapeHtml(announcement.createdBy) : '-';

    return `
      <article class="announcement-admin-card">
        <div class="announcement-admin-title">
          <strong>Announcement</strong>
          <span class="badge badge-info">${announcement.allowDismiss ? 'Dismissible' : 'Mandatory'}</span>
        </div>
        <div>${announcement.message || ''}</div>
        <div class="announcement-admin-meta">
          <span><i class="fas fa-location-dot"></i> ${locations}</span>
          <span><i class="fas fa-users"></i> ${targets}</span>
          <span><i class="fas fa-clock"></i> Expires: ${this.escapeHtml(expires)}</span>
          <span><i class="fas fa-user-shield"></i> ${createdBy}</span>
        </div>
      </article>
    `;
  },

  async displayGlobalAnnouncements(location) {
    try {
      const res = await API.getActiveAnnouncements(location);
      if (!res.success) return;

      const list = Array.isArray(res.announcements) ? res.announcements : [];
      const filtered = list.filter((announcement) => {
        const locations = announcement.locations || [];
        return locations.includes(location);
      });
      if (!filtered.length) return;

      const pending = filtered.filter((announcement) => {
        const seenKey = `pv2026_seen_announcement_${announcement.id}`;
        const dismissedPublicKey = `pv2026_public_dismissed_${announcement.id}`;
        return sessionStorage.getItem(seenKey) !== '1' && localStorage.getItem(dismissedPublicKey) !== '1';
      });

      if (!pending.length) return;

      this.popupQueue.push(...pending);
      this.showNextPopup();
    } catch (error) {
      console.error('Failed to display announcements:', error);
    }
  },

  showNextPopup() {
    if (this.popupOpen) return;
    const next = this.popupQueue.shift();
    if (!next) return;

    this.popupOpen = true;
    const host = document.getElementById('announcementPopupHost') || document.body;

    const overlay = document.createElement('div');
    overlay.className = 'announcement-popup-overlay';
    overlay.innerHTML = `
      <div class="announcement-popup">
        <div class="announcement-popup-header">
          <h4><i class="fas fa-bullhorn"></i> Announcement</h4>
          <button class="btn btn-secondary btn-sm" id="announcementPopupCloseBtn">
            <i class="fas fa-times"></i>
          </button>
        </div>
        <div class="announcement-popup-body">${next.message || ''}</div>
        <div class="announcement-popup-footer">
          ${next.allowDismiss ? `
            <label class="announcement-popup-dismiss">
              <input type="checkbox" id="announcementDontShowAgain"> Do not show again
            </label>
          ` : '<span></span>'}
          <button class="btn btn-primary btn-sm" id="announcementPopupOkBtn">Close</button>
        </div>
      </div>
    `;

    host.appendChild(overlay);

    const close = async () => {
      const dontShow = Boolean(overlay.querySelector('#announcementDontShowAgain')?.checked);
      if (dontShow && next.allowDismiss) {
        if (Auth.isLoggedIn()) {
          await this.dismiss(next.id);
        } else {
          localStorage.setItem(`pv2026_public_dismissed_${next.id}`, '1');
        }
      }

      sessionStorage.setItem(`pv2026_seen_announcement_${next.id}`, '1');
      overlay.remove();
      this.popupOpen = false;
      this.showNextPopup();
    };

    overlay.querySelector('#announcementPopupOkBtn')?.addEventListener('click', close);
    overlay.querySelector('#announcementPopupCloseBtn')?.addEventListener('click', close);
  },

  async dismiss(announcementId) {
    try {
      await API.dismissAnnouncement(announcementId);
    } catch (error) {
      console.error('Failed to dismiss announcement:', error);
    }
  }
};

window.Announcements = Announcements;

document.addEventListener('DOMContentLoaded', () => {
  const path = String(window.location.pathname || '').toLowerCase();
  const isDashboard = path.endsWith('dashboard.html');
  const isLogin = path.endsWith('index.html') || path.endsWith('login.html') || path === '/';

  if (isDashboard) {
    setTimeout(() => {
      if (Auth.isLoggedIn()) Announcements.displayGlobalAnnouncements('DASHBOARD');
    }, 700);
  }

  if (isLogin) {
    setTimeout(() => {
      Announcements.displayGlobalAnnouncements('LOGIN');
    }, 700);
  }
});
