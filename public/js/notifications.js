/**
 * PAYABLE VOUCHER 2026 - Notifications Module
 */

const Notifications = {
  unreadCount: 0,
  notifications: [],
  bellInitialized: false,
  pageInitialized: false,

  normalize(value) {
    return String(value || '').trim().toUpperCase();
  },

  escapeHtml(text) {
    return String(text || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  },

  async initBell() {
    if (!Auth.isLoggedIn()) return;

    await this.waitForElement('notificationBadge', 1200);
    await this.updateBellCount();

    if (!this.bellInitialized) {
      setInterval(() => this.updateBellCount(), 30000); // 30 seconds polling
      this.bellInitialized = true;
    }
  },

  waitForElement(id, timeout = 1000) {
    return new Promise((resolve) => {
      const element = document.getElementById(id);
      if (element) {
        resolve(element);
        return;
      }

      const start = Date.now();
      const interval = setInterval(() => {
        const target = document.getElementById(id);
        if (target || Date.now() - start >= timeout) {
          clearInterval(interval);
          resolve(target);
        }
      }, 50);
    });
  },

  async initPage() {
    if (this.pageInitialized) return;

    const isAuth = await Auth.requireAuth();
    if (!isAuth) return;

    this.pageInitialized = true;

    const sidebar = document.getElementById('sidebar');
    if (sidebar) sidebar.innerHTML = Components.getSidebar('notifications');

    const menuToggle = document.getElementById('menuToggle');
    if (menuToggle && sidebar) {
      const closeSidebar = () => sidebar.classList.remove('active');

      menuToggle.addEventListener('click', (event) => {
        event.stopPropagation();
        sidebar.classList.toggle('active');
      });

      document.addEventListener('click', (event) => {
        if (window.innerWidth > 992) return;
        if (!sidebar.contains(event.target) && !menuToggle.contains(event.target)) {
          closeSidebar();
        }
      });

      document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') closeSidebar();
      });

      window.addEventListener('resize', () => {
        if (window.innerWidth > 992) closeSidebar();
      });

      sidebar.querySelectorAll('.nav-item').forEach((link) => {
        link.addEventListener('click', () => {
          if (window.innerWidth <= 992) closeSidebar();
        });
      });
    }

    document.getElementById('refreshHubBtn')?.addEventListener('click', () => this.refreshHub());
    document.getElementById('markAllReadBtn')?.addEventListener('click', () => this.markAllAsRead());

    // Search and filter listeners
    document.getElementById('notificationSearchInput')?.addEventListener('input', () => this.renderNotifications());
    document.getElementById('notificationReadFilter')?.addEventListener('change', () => this.renderNotifications());
    document.getElementById('notificationDateFrom')?.addEventListener('change', () => this.renderNotifications());
    document.getElementById('notificationDateTo')?.addEventListener('change', () => this.renderNotifications());
    document.getElementById('clearNotificationFiltersBtn')?.addEventListener('click', () => {
      const searchInput = document.getElementById('notificationSearchInput');
      const readFilter = document.getElementById('notificationReadFilter');
      const dateFrom = document.getElementById('notificationDateFrom');
      const dateTo = document.getElementById('notificationDateTo');
      if (searchInput) searchInput.value = '';
      if (readFilter) readFilter.value = 'ALL';
      if (dateFrom) dateFrom.value = '';
      if (dateTo) dateTo.value = '';
      this.renderNotifications();
    });

    this.setupRoleView();
    this.setupTabs();

    await this.loadNotifications();
    await this.updateBellCount();

    if (window.ActionItems && typeof ActionItems.refreshFull === 'function') {
      ActionItems.refreshFull();
    }

    const isAdmin = this.normalize(Auth.getUser()?.role) === this.normalize(CONFIG.ROLES.ADMIN);
    if (window.Announcements && typeof Announcements.initComposer === 'function' && isAdmin) {
      Announcements.initComposer();
      Announcements.loadAnnouncements();
    } else if (window.Announcements && typeof Announcements.loadAnnouncements === 'function') {
      Announcements.loadAnnouncements();
    }
  },

  setupRoleView() {
    const user = Auth.getUser();
    const isAdmin = this.normalize(user?.role) === this.normalize(CONFIG.ROLES.ADMIN);

    const adminTabButton = document.getElementById('tabAnnouncementsBtn');
    const adminTab = document.getElementById('tab-announcements');
    if (adminTabButton) adminTabButton.style.display = isAdmin ? 'inline-flex' : 'none';
    if (adminTab) adminTab.style.display = isAdmin ? '' : 'none';
  },

  setupTabs() {
    const tabButtons = document.querySelectorAll('#notificationHubTabs .tab-btn');
    tabButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const targetId = button.getAttribute('data-target');
        this.activateTab(targetId);
      });
    });
  },

  activateTab(targetId) {
    if (!targetId) return;

    document.querySelectorAll('#notificationHubTabs .tab-btn').forEach((btn) => {
      btn.classList.toggle('active', btn.getAttribute('data-target') === targetId);
    });

    document.querySelectorAll('.hub-tab-content').forEach((section) => {
      section.classList.toggle('active', section.id === targetId);
    });

    if (targetId === 'tab-actionitems' && window.ActionItems) {
      ActionItems.refreshFull();
    }

    if (targetId === 'tab-announcements' && window.Announcements) {
      if (this.normalize(Auth.getUser()?.role) === this.normalize(CONFIG.ROLES.ADMIN)) {
        Announcements.initComposer();
      }
      Announcements.loadAnnouncements();
    }
  },

  async refreshHub() {
    await this.loadNotifications();
    await this.updateBellCount();

    if (window.ActionItems && typeof ActionItems.refreshFull === 'function') {
      await ActionItems.refreshFull();
    }

    if (window.Announcements && typeof Announcements.loadAnnouncements === 'function') {
      await Announcements.loadAnnouncements();
    }
  },

  async updateBellCount() {
    try {
      const result = await API.getNotifications(true);
      if (!result.success) return;

      const newUnreadCount = result.unreadCount || 0;

      if (this.bellInitialized && newUnreadCount > this.unreadCount) {
        const fullRes = await API.getNotifications(false);
        if (fullRes.success && fullRes.notifications) {
          const newItems = fullRes.notifications.filter(n => !n.read);
          const newlyReceived = newItems.slice(0, newUnreadCount - this.unreadCount);
          newlyReceived.reverse().forEach(item => {
            Utils.showToast(`🔔 ${item.title}: ${item.message}`, item.type || 'info');
          });
        }
      }

      this.unreadCount = newUnreadCount;
      const badge = document.getElementById('notificationBadge');
      if (badge) {
        if (this.unreadCount > 0) {
          badge.textContent = this.unreadCount > 99 ? '99+' : String(this.unreadCount);
          badge.style.display = 'inline-block';
        } else {
          badge.style.display = 'none';
        }
      }

      this.updateNotificationStats();
    } catch (error) {
      console.error('Failed to update notification count:', error);
    }
  },

  updateNotificationStats() {
    const unread = this.unreadCount || 0;

    const hubUnread = document.getElementById('hubUnreadCount');
    if (hubUnread) hubUnread.textContent = unread;

    const tabBadge = document.getElementById('tabNotificationBadge');
    if (tabBadge) {
      tabBadge.textContent = unread > 99 ? '99+' : String(unread);
      tabBadge.style.display = unread > 0 ? 'inline-block' : 'none';
    }

    const markAllButton = document.getElementById('markAllReadBtn');
    if (markAllButton) {
      markAllButton.style.display = this.notifications.length > 0 ? 'inline-flex' : 'none';
    }

    const countEl = document.getElementById('notificationCount');
    if (countEl) {
      countEl.textContent = `${this.notifications.length} notification(s), ${unread} unread`;
    }
  },

  async loadNotifications() {
    const container = document.getElementById('notificationsList');
    if (!container) return;

    container.innerHTML = '<p class="text-muted text-center">Loading notifications...</p>';

    try {
      const result = await API.getNotifications(false);
      if (!result.success) {
        container.innerHTML = `<div class="alert alert-danger">${this.escapeHtml(result.error || 'Failed to load notifications')}</div>`;
        return;
      }

      this.notifications = Array.isArray(result.notifications) ? result.notifications : [];
      this.unreadCount = result.unreadCount || 0;
      this.updateNotificationStats();

      if (!this.notifications.length) {
        container.innerHTML = `
          <div class="empty-state">
            <i class="fas fa-bell-slash"></i>
            <h3>No Notifications</h3>
            <p class="text-muted">You're all caught up.</p>
          </div>
        `;
        return;
      }

      this.renderNotifications();
    } catch (error) {
      console.error('Load notifications error:', error);
      container.innerHTML = '<div class="alert alert-danger">Failed to load notifications.</div>';
    }
  },

  renderNotifications() {
    const container = document.getElementById('notificationsList');
    if (!container) return;

    const query = document.getElementById('notificationSearchInput')?.value?.toLowerCase() || '';
    const readFilter = document.getElementById('notificationReadFilter')?.value || 'UNREAD';
    const dateFromVal = document.getElementById('notificationDateFrom')?.value || '';
    const dateToVal = document.getElementById('notificationDateTo')?.value || '';

    const dateFrom = dateFromVal ? new Date(dateFromVal + 'T00:00:00') : null;
    const dateTo = dateToVal ? new Date(dateToVal + 'T23:59:59') : null;

    const filtered = this.notifications.filter(notification => {
      const isRead = Boolean(notification.read);
      if (readFilter === 'UNREAD' && isRead) return false;
      if (readFilter === 'READ' && !isRead) return false;

      if (query) {
        const title = (notification.title || '').toLowerCase();
        const msg = (notification.message || '').toLowerCase();
        if (!title.includes(query) && !msg.includes(query)) return false;
      }

      if (dateFrom || dateTo) {
        const ts = new Date(notification.timestamp);
        if (dateFrom && ts < dateFrom) return false;
        if (dateTo && ts > dateTo) return false;
      }

      return true;
    });

    const countEl = document.getElementById('notificationCount');
    if (countEl) {
      countEl.textContent = `${filtered.length} matching of ${this.notifications.length} notification(s), ${this.unreadCount} unread`;
    }

    if (!filtered.length) {
      container.innerHTML = `
        <div class="empty-state">
          <i class="fas fa-filter"></i>
          <h3>No matching notifications</h3>
          <p class="text-muted">Adjust your filters to see more results.</p>
        </div>
      `;
      return;
    }

    let html = '<div class="notifications-list">';
    filtered.forEach((notification) => {
      const rowIndex = Number(notification.rowIndex || 0);
      const isRead = Boolean(notification.read);
      const typeClass = this.getTypeClass(notification.type);
      const iconClass = this.getTypeIcon(notification.type);
      const readClass = isRead ? 'notification-read' : 'notification-unread';
      const safeTitle = this.escapeHtml(notification.title || 'Notification');
      const safeMessage = this.escapeHtml(notification.message || '');
      const safeTime = this.escapeHtml(Utils.formatDateTime(notification.timestamp));

      html += `
        <div class="notification-item ${typeClass} ${readClass}" data-row="${rowIndex}" ${notification.link ? `style="cursor: pointer;" onclick="Notifications.handleNotificationClick(${rowIndex}, '${notification.link}')"` : ''}>
          <div class="notification-icon">
            <i class="fas ${iconClass}"></i>
          </div>
          <div class="notification-content">
            <div class="notification-title">${safeTitle}</div>
            <div class="notification-message">${safeMessage}</div>
            <div class="notification-time"><i class="fas fa-clock"></i> ${safeTime}</div>
          </div>
          <div class="notification-actions" onclick="event.stopPropagation();">
            ${!isRead ? `
              <button class="btn btn-sm btn-secondary" onclick="Notifications.markAsRead(${rowIndex})">
                <i class="fas fa-check"></i>
              </button>
            ` : ''}
          </div>
        </div>
      `;
    });
    html += '</div>';

    container.innerHTML = html;
  },

  getTypeIcon(type) {
    switch ((type || '').toLowerCase()) {
      case 'success':
        return 'fa-check-circle';
      case 'warning':
        return 'fa-exclamation-triangle';
      case 'danger':
      case 'error':
        return 'fa-times-circle';
      default:
        return 'fa-info-circle';
    }
  },

  getTypeClass(type) {
    switch ((type || '').toLowerCase()) {
      case 'success':
        return 'notification-success';
      case 'warning':
        return 'notification-warning';
      case 'danger':
      case 'error':
        return 'notification-danger';
      default:
        return 'notification-info';
    }
  },

  async markAsRead(rowIndex) {
    if (!rowIndex) return;

    try {
      const result = await API.markNotificationRead(rowIndex);
      if (!result.success) return;

      const target = this.notifications.find((item) => Number(item.rowIndex) === Number(rowIndex));
      if (target) target.read = true;
      this.unreadCount = Math.max(0, this.unreadCount - 1);

      this.renderNotifications();
      await this.updateBellCount();
    } catch (error) {
      console.error('Mark notification read error:', error);
    }
  },

  async handleNotificationClick(rowIndex, link) {
    if (!rowIndex) return;
    try {
      await this.markAsRead(rowIndex);
      if (link) {
        window.location.href = link;
      }
    } catch (error) {
      console.error('Handle notification click error:', error);
      if (link) window.location.href = link;
    }
  },

  async markAllAsRead() {
    try {
      const result = await API.markAllNotificationsRead();
      if (!result.success) {
        Utils.showToast(result.error || 'Failed to mark all notifications as read.', 'error');
        return;
      }

      Utils.showToast(result.message || 'All notifications marked as read.', 'success');
      await this.loadNotifications();
      await this.updateBellCount();
    } catch (error) {
      console.error('Mark all notifications error:', error);
      Utils.showToast('Failed to mark all notifications as read.', 'error');
    }
  },
};

window.Notifications = Notifications;

document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    if (Auth.isLoggedIn()) Notifications.initBell();
  }, 300);
});
