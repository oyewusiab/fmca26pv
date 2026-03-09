/**
 * PAYABLE VOUCHER 2026 - Notifications Module
 */

const Notifications = {
    unreadCount: 0,
    notifications: [],
    bellInitialized: false,
    
    /**
     * Initialize notification bell (call on every page)
     */
    async initBell() {
        console.log('Notifications.initBell() called');
        
        if (!Auth.isLoggedIn()) {
            console.log('User not logged in, skipping bell init');
            return;
        }
        
        // Wait a bit for sidebar to be fully rendered
        await this.waitForElement('notificationBadge', 1000);
        
        // Load unread count
        await this.updateBellCount();
        
        // Refresh count every 2 minutes
        if (!this.bellInitialized) {
            setInterval(() => this.updateBellCount(), 120000);
            this.bellInitialized = true;
        }
        
        console.log('Notification bell initialized');
    },
    
    /**
     * Wait for an element to exist in DOM
     */
    waitForElement(id, timeout = 1000) {
        return new Promise((resolve) => {
            const element = document.getElementById(id);
            if (element) {
                resolve(element);
                return;
            }
            
            const startTime = Date.now();
            const interval = setInterval(() => {
                const el = document.getElementById(id);
                if (el || Date.now() - startTime > timeout) {
                    clearInterval(interval);
                    resolve(el);
                }
            }, 50);
        });
    },
    
    /**
     * Update bell badge count
     */
    async updateBellCount() {
        try {
            const result = await API.getNotifications(true); // Only unread
            
            if (result.success) {
                this.unreadCount = result.unreadCount || 0;
                
                const badge = document.getElementById('notificationBadge');
                console.log('Notification badge element:', badge);
                console.log('Unread count:', this.unreadCount);
                
                if (badge) {
                    if (this.unreadCount > 0) {
                        badge.textContent = this.unreadCount > 99 ? '99+' : this.unreadCount;
                        badge.style.display = 'inline-block';
                        console.log('Badge updated:', badge.textContent);
                    } else {
                        badge.style.display = 'none';
                    }
                } else {
                    console.warn('Notification badge element not found');
                }
            }
        } catch (e) {
            console.error('Failed to update notification count:', e);
        }
    },
    
    /**
     * Initialize notifications page
     */
    async initPage() {
        const isAuth = await Auth.requireAuth();
        if (!isAuth) return;
        
        // Setup sidebar
        const sidebar = document.getElementById('sidebar');
        if (sidebar) {
            sidebar.innerHTML = Components.getSidebar('notifications');
        }
        
        // Setup mobile menu
        document.getElementById('menuToggle')?.addEventListener('click', () => {
            document.getElementById('sidebar')?.classList.toggle('active');
        });
        
        // Load notifications
        await this.loadNotifications();
        
        // Setup mark all as read button
        document.getElementById('markAllReadBtn')?.addEventListener('click', () => {
            this.markAllAsRead();
        });
        
        // Update bell count
        await this.updateBellCount();
    },
    
    /**
     * Load all notifications
     */
    async loadNotifications() {
        const container = document.getElementById('notificationsList');
        if (!container) return;
        
        container.innerHTML = '<p class="text-muted text-center">Loading notifications...</p>';
        
        try {
            const result = await API.getNotifications(false); // All notifications
            
            if (!result.success) {
                container.innerHTML = `<p class="text-danger text-center">${result.error}</p>`;
                return;
            }
            
            this.notifications = result.notifications || [];
            this.unreadCount = result.unreadCount || 0;
            
            // Update count display
            const countEl = document.getElementById('notificationCount');
            if (countEl) {
                countEl.textContent = `${this.notifications.length} notification(s), ${this.unreadCount} unread`;
            }
            
            if (this.notifications.length === 0) {
                container.innerHTML = `
                    <div class="empty-state" style="text-align: center; padding: 60px 20px;">
                        <i class="fas fa-bell-slash" style="font-size: 64px; color: var(--text-muted); margin-bottom: 20px;"></i>
                        <h3 style="color: var(--text-muted);">No Notifications</h3>
                        <p class="text-muted">You're all caught up!</p>
                    </div>
                `;
                return;
            }
            
            this.renderNotifications();
            
        } catch (e) {
            console.error('Load notifications error:', e);
            container.innerHTML = '<p class="text-danger text-center">Failed to load notifications</p>';
        }
    },
    
    /**
     * Render notifications list
     */
    renderNotifications() {
        const container = document.getElementById('notificationsList');
        if (!container) return;
        
        let html = '<div class="notifications-list">';
        
        this.notifications.forEach(notif => {
            const typeIcon = this.getTypeIcon(notif.type);
            const typeClass = this.getTypeClass(notif.type);
            const readClass = notif.read ? 'notification-read' : 'notification-unread';
            
            html += `
                <div class="notification-item ${readClass} ${typeClass}" data-row="${notif.rowIndex}">
                    <div class="notification-icon">
                        <i class="fas ${typeIcon}"></i>
                    </div>
                    <div class="notification-content">
                        <div class="notification-title">${notif.title}</div>
                        <div class="notification-message">${notif.message}</div>
                        <div class="notification-time">
                            <i class="fas fa-clock"></i> ${Utils.formatDateTime(notif.timestamp)}
                        </div>
                    </div>
                    <div class="notification-actions">
                        ${notif.link ? `
                            <a href="${notif.link}" class="btn btn-sm btn-primary" onclick="Notifications.markAsRead(${notif.rowIndex})">
                                <i class="fas fa-external-link-alt"></i> View
                            </a>
                        ` : ''}
                        ${!notif.read ? `
                            <button class="btn btn-sm btn-secondary" onclick="Notifications.markAsRead(${notif.rowIndex})">
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
    
    /**
     * Get icon for notification type
     */
    getTypeIcon(type) {
        switch(type) {
            case 'success': return 'fa-check-circle';
            case 'warning': return 'fa-exclamation-triangle';
            case 'danger': return 'fa-times-circle';
            default: return 'fa-info-circle';
        }
    },
    
    /**
     * Get CSS class for notification type
     */
    getTypeClass(type) {
        switch(type) {
            case 'success': return 'notification-success';
            case 'warning': return 'notification-warning';
            case 'danger': return 'notification-danger';
            default: return 'notification-info';
        }
    },
    
    /**
     * Mark single notification as read
     */
    async markAsRead(rowIndex) {
        try {
            const result = await API.markNotificationRead(rowIndex);
            
            if (result.success) {
                // Update local state
                const notif = this.notifications.find(n => n.rowIndex === rowIndex);
                if (notif && !notif.read) {
                    notif.read = true;
                    this.unreadCount = Math.max(0, this.unreadCount - 1);
                }
                
                // Update UI
                const item = document.querySelector(`.notification-item[data-row="${rowIndex}"]`);
                if (item) {
                    item.classList.remove('notification-unread');
                    item.classList.add('notification-read');
                }
                
                // Update badge
                await this.updateBellCount();
            }
        } catch (e) {
            console.error('Mark as read error:', e);
        }
    },
    
    /**
     * Mark all notifications as read
     */
    async markAllAsRead() {
        try {
            const result = await API.markAllNotificationsRead();
            
            if (result.success) {
                Utils.showToast(result.message || 'All notifications marked as read', 'success');
                await this.loadNotifications();
                await this.updateBellCount();
            } else {
                Utils.showToast(result.error || 'Failed to mark as read', 'error');
            }
        } catch (e) {
            console.error('Mark all as read error:', e);
            Utils.showToast('Error marking notifications as read', 'error');
        }
    }
};

// Initialize bell on every page after DOM and Auth are ready
document.addEventListener('DOMContentLoaded', () => {
    // Small delay to ensure sidebar is rendered
    setTimeout(() => {
        if (Auth.isLoggedIn()) {
            Notifications.initBell();
        }
    }, 300);
});