/**
 * PAYABLE VOUCHER 2026 - Shared UI Components
 */

const Components = {
    /**
     * Generates the sidebar HTML
     * @param {string} activePage - Current active page name
     * @returns {string} Sidebar HTML
     */
    getSidebar(activePage = 'dashboard') {
        const user = Auth.getUser();
        if (!user) return '';

        const isActive = (page) => (activePage === page ? 'active' : '');
        const isPayableUnit = [
            CONFIG.ROLES.PAYABLE_STAFF,
            CONFIG.ROLES.PAYABLE_HEAD
            ].includes(user.role);

        // Base navigation
        let navItems = `
            <a href="dashboard.html" class="nav-item ${isActive('dashboard')}">
            <i class="fas fa-tachometer-alt"></i> Dashboard
            </a>
            <a href="vouchers.html" class="nav-item ${isActive('vouchers')}">
            <i class="fas fa-file-invoice-dollar"></i> Vouchers
            </a>
            ${!isPayableUnit ? `
            <a href="reports.html" class="nav-item ${isActive('reports')}">
                <i class="fas fa-chart-bar"></i> Reports
            </a>
            ` : ''}
            <a href="notifications.html" class="nav-item ${isActive('notifications')}" id="notificationBell">
            <i class="fas fa-bell"></i> Notifications
            <span class="nav-badge" id="notificationBadge" style="display:none;">0</span>
            </a>
        `;

        // Admin-only Audit Trail
        if (user.role === CONFIG.ROLES.ADMIN) {
            navItems += `
            <a href="audit-trail.html" class="nav-item ${isActive('auditTrail')}">
                <i class="fas fa-clipboard-list"></i> Audit Trail
            </a>
            `;
        }

        // Admin-only User Management
        if (user.role === CONFIG.ROLES.ADMIN) {
            navItems += `
            <a href="users.html" class="nav-item ${isActive('users')}">
                <i class="fas fa-users-cog"></i> User Management
            </a>
            `;
        }

        return `
            <div class="sidebar-header">
            <img src="images/fmc-logo.png" alt="FMC Logo" class="sidebar-logo"
                onerror="this.src='https://via.placeholder.com/60?text=FMC'">
            <h2 class="sidebar-title">PAYABLE VOUCHERS</h2>
            <p class="sidebar-subtitle">FMC Abeokuta • Finance Dept</p>
            </div>

            <nav class="sidebar-nav" id="sidebarNav">
            ${navItems}
            </nav>

            <div class="sidebar-footer">
            <div class="user-info"
                style="cursor:pointer;"
                role="button"
                tabindex="0"
                onclick="window.location.href='profile.html'"
                onkeydown="if(event.key==='Enter'){window.location.href='profile.html'}">
                <div class="user-avatar" id="userAvatar">${Utils.getInitials(user.name)}</div>
                <div class="user-details">
                <div class="user-name" id="userName">${user.name}</div>
                <div class="user-role" id="userRole">${user.role}</div>
                </div>
            </div>
            <button class="btn-logout" id="logoutBtn" onclick="Auth.logout()">
                <i class="fas fa-sign-out-alt"></i> Logout
            </button>
            </div>
        `;    
    },
    
    /**
     * Generates loading overlay HTML
     */
    getLoadingOverlay(message = 'Loading...') {
        return `
            <div id="loadingOverlay" class="loading-overlay hidden">
                <div class="spinner"></div>
                <p class="loading-text">${message}</p>
            </div>
        `;
    },
    
    /**
     * Generates empty state HTML
     */
    getEmptyState(message = 'No data found', icon = 'fa-inbox') {
        return `
            <div class="empty-state" style="text-align: center; padding: 40px 20px; color: var(--text-muted);">
                <i class="fas ${icon}" style="font-size: 48px; margin-bottom: 15px; opacity: 0.5;"></i>
                <p>${message}</p>
            </div>
        `;
    },
    
    /**
     * Generates pagination HTML
     */
    getPagination(currentPage, totalPages, containerId = 'pagination') {
        if (totalPages <= 1) return '';
        
        let html = '<div class="pagination" style="display: flex; justify-content: center; gap: 5px; margin-top: 20px;">';
        
        html += `
            <button class="btn btn-sm btn-secondary" 
                    onclick="handlePageChange(${currentPage - 1})"
                    ${currentPage === 1 ? 'disabled' : ''}>
                <i class="fas fa-chevron-left"></i>
            </button>
        `;
        
        for (let i = 1; i <= totalPages; i++) {
            if (i === currentPage) {
                html += `<button class="btn btn-sm btn-primary">${i}</button>`;
            } else if (i === 1 || i === totalPages || (i >= currentPage - 2 && i <= currentPage + 2)) {
                html += `<button class="btn btn-sm btn-secondary" onclick="handlePageChange(${i})">${i}</button>`;
            } else if (i === currentPage - 3 || i === currentPage + 3) {
                html += `<span style="padding: 5px;">...</span>`;
            }
        }
        
        html += `
            <button class="btn btn-sm btn-secondary" 
                    onclick="handlePageChange(${currentPage + 1})"
                    ${currentPage === totalPages ? 'disabled' : ''}>
                <i class="fas fa-chevron-right"></i>
            </button>
        `;
        
        html += '</div>';
        return html;
    },
    
    /**
     * Opens a modal
     */
    openModal(id) {
        const modal = document.getElementById(id);
        if (modal) {
            modal.classList.add('active');
        }
    },
    
    /**
     * Closes a modal
     */
    closeModal(id) {
        const modal = document.getElementById(id);
        if (modal) {
            modal.classList.remove('active');
        }
    },
    
    /**
     * Initializes common page elements
     */
    async initPage(activePage) {
        const isAuth = await Auth.requireAuth();
        if (!isAuth) return false;
        
        const sidebar = document.getElementById('sidebar');
        if (sidebar) {
            sidebar.innerHTML = this.getSidebar(activePage);
        }
        
        const menuToggle = document.getElementById('menuToggle');
        if (menuToggle && sidebar) {
            menuToggle.addEventListener('click', () => {
                sidebar.classList.toggle('active');
            });
        }
        
        document.addEventListener('click', (e) => {
            if (window.innerWidth <= 992 && sidebar && menuToggle) {
                if (!sidebar.contains(e.target) && !menuToggle.contains(e.target)) {
                    sidebar.classList.remove('active');
                }
            }
        });
        
        return true;
    }
};

console.log('Components loaded successfully');