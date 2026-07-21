/**
 * PAYABLE VOUCHER 2026 - User Management Module
 * Admin only functionality
 */

const Users = {
    // State
    users: [],
    selectedUser: null,
    isEditMode: false,
    currentPage: 1,
    itemsPerPage: 10,
    credentialsData: null,
    
    /**
     * Initialize users page
     */
    async init() {
        // Check authentication
        const isAuth = await Auth.requireAuth();
        if (!isAuth) return;
        
        // Check if user is admin
        const user = Auth.getUser();
        if (!user || user.role !== CONFIG.ROLES.ADMIN) {
            Utils.showToast('Unauthorized: Admin access required', 'error');
            window.location.href = 'dashboard.html';
            return;
        }
        
        // Setup sidebar
        this.setupSidebar();
        
        // Populate filter roles
        this.setupRoleFilters();

        // Load users
        await this.loadUsers();
        
        // Setup event listeners
        this.setupEventListeners();
    },

    setupRoleFilters() {
        const filterRole = document.getElementById('filterRole');
        if (filterRole && CONFIG.ROLES) {
            filterRole.innerHTML = '<option value="">All Roles</option>';
            Object.values(CONFIG.ROLES).forEach(role => {
                filterRole.innerHTML += `<option value="${role}">${role}</option>`;
            });
        }
    },
    
    /**
     * Setup sidebar navigation
     */
    setupSidebar() {
        const sidebar = document.getElementById('sidebar');
        if (sidebar) {
            sidebar.innerHTML = Components.getSidebar('users');
        }
    },
    
    /**
     * Setup event listeners
     */
    setupEventListeners() {
        // Logout
        document.getElementById('logoutBtn')?.addEventListener('click', () => Auth.logout());
        
        // Mobile menu
        document.getElementById('menuToggle')?.addEventListener('click', () => {
            document.getElementById('sidebar')?.classList.toggle('active');
        });
        
        // Search
        const searchInput = document.getElementById('searchUsers');
        if (searchInput) {
            searchInput.addEventListener('input', Utils.debounce(() => this.filterUsers(), 300));
        }
        
        // Form submission
        document.getElementById('userForm')?.addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveUser();
        });
        
        // Close modal on outside click
        document.querySelectorAll('.modal-overlay').forEach(modal => {
            modal.addEventListener('click', (e) => {
                if (e.target === modal) {
                    modal.classList.remove('active');
                }
            });
        });
    },

    onEmailInput() {
        if (this.isEditMode) return;
        const email = (document.getElementById('formEmail')?.value || '').trim();
        const usernameInput = document.getElementById('formUsername');
        if (email && usernameInput && (!usernameInput.value || usernameInput.dataset.autoDerived === 'true')) {
            const prefix = email.split('@')[0] || '';
            usernameInput.value = prefix.toLowerCase().replace(/[^a-z0-9._-]/g, '');
            usernameInput.dataset.autoDerived = 'true';
        }
    },
    
    /**
     * Load users from backend
     */
    async loadUsers() {
        this.showLoading(true);
        
        try {
            const result = await API.getUsers();
            
            if (result.success) {
                this.users = result.users || [];
                this.renderUsersList();
            } else {
                Utils.showToast(result.error || 'Failed to load users', 'error');
            }
        } catch (error) {
            console.error('Load users error:', error);
            Utils.showToast('Error loading users', 'error');
        }
        
        this.showLoading(false);
    },
    
    /**
     * Filter users based on search & dropdowns
     */
    filterUsers() {
        this.currentPage = 1;
        this.renderUsersList();
    },

    changeItemsPerPage() {
        const select = document.getElementById('itemsPerPage');
        if (select) {
            this.itemsPerPage = parseInt(select.value, 10) || 10;
            this.currentPage = 1;
            this.renderUsersList();
        }
    },

    goToPage(page) {
        this.currentPage = page;
        this.renderUsersList();
    },
    
    escapeHtml(str) {
        if (str === null || str === undefined) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    },

    /**
     * Render users list
     */
    renderUsersList() {
        const container = document.getElementById('usersList');
        if (!container) return;
        
        const searchTerm = (document.getElementById('searchUsers')?.value || '').toLowerCase().trim();
        const statusFilter = (document.getElementById('filterStatus')?.value || '').toLowerCase();
        const roleFilter = (document.getElementById('filterRole')?.value || '').toLowerCase();
        
        // Filter users
        let filtered = this.users.filter(user => {
            const name = String(user.name || '').toLowerCase();
            const email = String(user.email || '').toLowerCase();
            const username = String(user.username || '').toLowerCase();
            const role = String(user.role || '').toLowerCase();
            const dept = String(user.department || '').toLowerCase();
            const isActive = user.active === true || user.active === 'TRUE' || user.active === 'true';

            // Text search
            if (searchTerm) {
                const matchesText = name.includes(searchTerm) ||
                                    email.includes(searchTerm) ||
                                    username.includes(searchTerm) ||
                                    role.includes(searchTerm) ||
                                    dept.includes(searchTerm);
                if (!matchesText) return false;
            }

            // Status filter
            if (statusFilter === 'active' && !isActive) return false;
            if (statusFilter === 'inactive' && isActive) return false;

            // Role filter
            if (roleFilter && role !== roleFilter) return false;

            return true;
        });
        
        // Update count
        document.getElementById('userCount').textContent = `${filtered.length} user(s) found`;

        if (filtered.length === 0) {
            container.innerHTML = Components.getEmptyState('No users match the criteria', 'fa-users');
            document.getElementById('usersPaginationContainer').innerHTML = '';
            return;
        }

        // Pagination slice
        const totalItems = filtered.length;
        const totalPages = Math.ceil(totalItems / this.itemsPerPage);
        if (this.currentPage > totalPages) this.currentPage = Math.max(1, totalPages);

        const startIndex = (this.currentPage - 1) * this.itemsPerPage;
        const paginatedUsers = filtered.slice(startIndex, startIndex + this.itemsPerPage);
        
        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>User Profile</th>
                            <th>Email & Username</th>
                            <th>Department & Contact</th>
                            <th>Role</th>
                            <th>Status</th>
                            <th style="text-align: right;">Actions</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        paginatedUsers.forEach(user => {
            const isActive = user.active === true || user.active === 'TRUE' || user.active === 'true';
            const statusBadge = isActive
                ? '<span class="badge badge-paid">Active</span>'
                : '<span class="badge badge-cancelled">Inactive</span>';

            const roleClass = (user.role || '').toLowerCase().replace(/\s+/g, '-');
            const deptText = user.department || '-';
            const phoneText = user.phone || '';
            const usernameText = user.username ? `@${user.username}` : '';
            
            html += `
                <tr>
                    <td>
                        <div style="display: flex; align-items: center; gap: 10px;">
                            <div class="user-avatar-small">${Utils.getInitials(user.name)}</div>
                            <div>
                                <strong>${this.escapeHtml(user.name)}</strong>
                            </div>
                        </div>
                    </td>
                    <td>
                        <div>${this.escapeHtml(user.email)}</div>
                        ${usernameText ? `<div style="font-size:11px; color:#666;">${this.escapeHtml(usernameText)}</div>` : ''}
                    </td>
                    <td>
                        <div>${this.escapeHtml(deptText)}</div>
                        ${phoneText ? `<div style="font-size:11px; color:#666;"><i class="fas fa-phone-alt" style="font-size:10px;"></i> ${this.escapeHtml(phoneText)}</div>` : ''}
                    </td>
                    <td><span class="role-badge role-${roleClass}">${this.escapeHtml(user.role)}</span></td>
                    <td>${statusBadge}</td>
                    <td style="text-align: right;">
                        <div class="action-buttons" style="justify-content: flex-end;">
                            <button class="btn btn-sm btn-primary" onclick="Users.editUser(${user.rowIndex})" title="Edit User">
                                <i class="fas fa-edit"></i>
                            </button>
                            <button class="btn btn-sm btn-secondary" onclick="Users.resetPassword(${user.rowIndex})" title="Reset Password">
                                <i class="fas fa-key"></i>
                            </button>
                            <button class="btn btn-sm btn-danger" onclick="Users.deleteUser(${user.rowIndex})" title="Delete User">
                                <i class="fas fa-trash"></i>
                            </button>
                        </div>
                    </td>
                </tr>
            `;
        });
        
        html += '</tbody></table></div>';
        container.innerHTML = html;

        // Render pagination
        const pageContainer = document.getElementById('usersPaginationContainer');
        if (pageContainer) {
            if (totalPages > 1) {
                pageContainer.innerHTML = Components.getPagination(this.currentPage, totalPages, totalItems);
                // Attach click handlers
                pageContainer.querySelectorAll('.pagination-btn').forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        const targetPage = parseInt(e.currentTarget.dataset.page, 10);
                        if (targetPage) this.goToPage(targetPage);
                    });
                });
            } else {
                pageContainer.innerHTML = '';
            }
        }
    },
    
    /**
     * Open user form modal
     */
    openUserForm(user = null) {
        this.isEditMode = !!user;
        this.selectedUser = user;
        
        const modal = document.getElementById('userFormModal');
        const title = document.getElementById('userFormTitle');
        const form = document.getElementById('userForm');
        
        title.textContent = this.isEditMode ? 'Edit User Profile' : 'Add New User';
        form.reset();
        
        // Populate role dropdown
        const roleSelect = document.getElementById('formRole');
        roleSelect.innerHTML = '<option value="">Select Role</option>';
        Object.values(CONFIG.ROLES).forEach(role => {
            roleSelect.innerHTML += `<option value="${role}">${role}</option>`;
        });
        
        const activeCheck = document.getElementById('formActive');
        const selfNotice = document.getElementById('selfEditNotice');
        const currentUser = Auth.getUser();

        if (user) {
            document.getElementById('formName').value = user.name || '';
            document.getElementById('formEmail').value = user.email || '';
            document.getElementById('formUsername').value = user.username || '';
            document.getElementById('formDepartment').value = user.department || '';
            document.getElementById('formPhone').value = user.phone || '';
            document.getElementById('formRole').value = user.role || '';
            activeCheck.checked = user.active === true || user.active === 'TRUE' || user.active === 'true';

            // Self edit protection
            if (currentUser && user.email && user.email.toLowerCase() === currentUser.email.toLowerCase()) {
                activeCheck.disabled = true;
                roleSelect.disabled = true;
                if (selfNotice) selfNotice.classList.remove('hidden');
            } else {
                activeCheck.disabled = false;
                roleSelect.disabled = false;
                if (selfNotice) selfNotice.classList.add('hidden');
            }
        } else {
            activeCheck.checked = true;
            activeCheck.disabled = false;
            roleSelect.disabled = false;
            if (selfNotice) selfNotice.classList.add('hidden');
        }
        
        modal.classList.add('active');
    },
    
    /**
     * Edit user
     */
    editUser(rowIndex) {
        const user = this.users.find(u => u.rowIndex === rowIndex);
        if (user) {
            this.openUserForm(user);
        }
    },
    
    /**
     * Save user (create or update)
     */
    async saveUser() {
        const form = document.getElementById('userForm');
        
        if (!form.checkValidity()) {
            form.reportValidity();
            return;
        }

        const name = document.getElementById('formName').value.trim();
        const email = document.getElementById('formEmail').value.trim();
        let username = document.getElementById('formUsername').value.trim();
        const department = document.getElementById('formDepartment').value;
        const phone = document.getElementById('formPhone').value.trim();
        const role = document.getElementById('formRole').value;
        const active = document.getElementById('formActive').checked;
        
        if (!username && email) {
            username = email.split('@')[0].toLowerCase().replace(/[^a-z0-9._-]/g, '');
        }

        const userData = {
            name,
            email,
            username,
            department,
            phone,
            role,
            active
        };
        
        if (!userData.name || !userData.email || !userData.role || !userData.username) {
            Utils.showToast('Please fill all required fields (Name, Email, Username, Role)', 'error');
            return;
        }
        
        this.showLoading(true);
        
        try {
            let result;
            
            if (this.isEditMode && this.selectedUser) {
                result = await API.updateUser(this.selectedUser.rowIndex, userData);
            } else {
                result = await API.createUser(userData);
            }
            
            if (result.success) {
                Utils.showToast(result.message || 'User saved successfully', 'success');
                this.closeModal('userFormModal');
                await this.loadUsers();

                // On new user creation, present credentials modal
                if (!this.isEditMode) {
                    this.showCredentialsModal({
                        name: userData.name,
                        email: userData.email,
                        username: userData.username,
                        role: userData.role,
                        password: 'Welcome123',
                        message: 'New user account created successfully! Provide these credentials to the staff member:'
                    });
                }
            } else {
                Utils.showToast(result.error || 'Failed to save user', 'error');
            }
        } catch (error) {
            console.error('Save user error:', error);
            Utils.showToast('Error saving user', 'error');
        }
        
        this.showLoading(false);
    },
    
    /**
     * Reset user password
     */
    async resetPassword(rowIndex) {
        const user = this.users.find(u => u.rowIndex === rowIndex);
        if (!user) return;
        
        const confirm = await Utils.confirm(
            `Reset password for ${user.name}?\n\nThe temporary password will be set to: Welcome123`,
            'Reset Password'
        );
        
        if (!confirm) return;
        
        this.showLoading(true);
        
        try {
            const result = await API.updateUser(rowIndex, { resetPassword: true });
            
            if (result.success) {
                Utils.showToast(result.message || 'Password reset successfully', 'success');
                this.showCredentialsModal({
                    name: user.name,
                    email: user.email,
                    username: user.username || user.email.split('@')[0],
                    role: user.role,
                    password: 'Welcome123',
                    message: `Password reset successfully for ${user.name}. Here are the temporary credentials:`
                });
            } else {
                Utils.showToast(result.error || 'Failed to reset password', 'error');
            }
        } catch (error) {
            console.error('Reset password error:', error);
            Utils.showToast('Error resetting password', 'error');
        }
        
        this.showLoading(false);
    },

    showCredentialsModal(data) {
        this.credentialsData = data;
        const modal = document.getElementById('credentialsModal');
        const msg = document.getElementById('credentialsMessage');
        const credName = document.getElementById('credName');
        const credEmail = document.getElementById('credEmail');
        const credUsername = document.getElementById('credUsername');
        const credRole = document.getElementById('credRole');
        const credPass = document.getElementById('credPassword');

        if (!modal) return;

        if (msg) msg.textContent = data.message || 'User login details:';
        if (credName) credName.textContent = data.name || '-';
        if (credEmail) credEmail.textContent = data.email || '-';
        if (credUsername) credUsername.textContent = data.username || '-';
        if (credRole) credRole.textContent = data.role || '-';
        if (credPass) credPass.textContent = data.password || 'Welcome123';

        modal.classList.add('active');
    },

    copyCredentials() {
        if (!this.credentialsData) return;
        const d = this.credentialsData;
        const text = `PAYABLE VOUCHER 2026 LOGIN CREDENTIALS:\n` +
                     `Name: ${d.name}\n` +
                     `Email: ${d.email}\n` +
                     `Username: ${d.username}\n` +
                     `Role: ${d.role}\n` +
                     `Temporary Password: ${d.password}\n\n` +
                     `Please change your password upon logging in.`;

        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).then(() => {
                Utils.showToast('Credentials copied to clipboard!', 'success');
            }).catch(() => {
                this.fallbackCopyText(text);
            });
        } else {
            this.fallbackCopyText(text);
        }
    },

    fallbackCopyText(text) {
        const textArea = document.createElement("textarea");
        textArea.value = text;
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        try {
            document.execCommand('copy');
            Utils.showToast('Credentials copied to clipboard!', 'success');
        } catch (err) {
            Utils.showToast('Failed to copy text', 'error');
        }
        document.body.removeChild(textArea);
    },
    
    /**
     * Delete user
     */
    async deleteUser(rowIndex) {
        const user = this.users.find(u => u.rowIndex === rowIndex);
        if (!user) return;
        
        const currentUser = Auth.getUser();
        if (user.email.toLowerCase() === currentUser.email.toLowerCase()) {
            Utils.showToast('You cannot delete your own account', 'error');
            return;
        }
        
        const confirm = await Utils.confirm(
            `Are you sure you want to delete user:\n\n${user.name}\n${user.email}\n\nThis action cannot be undone!`,
            'Delete User'
        );
        
        if (!confirm) return;
        
        this.showLoading(true);
        
        try {
            const result = await API.deleteUser(rowIndex);
            
            if (result.success) {
                Utils.showToast(result.message || 'User deleted successfully', 'success');
                await this.loadUsers();
            } else {
                Utils.showToast(result.error || 'Failed to delete user', 'error');
            }
        } catch (error) {
            console.error('Delete user error:', error);
            Utils.showToast('Error deleting user', 'error');
        }
        
        this.showLoading(false);
    },
    
    /**
     * Close modal
     */
    closeModal(modalId) {
        const modal = document.getElementById(modalId);
        if (modal) {
            modal.classList.remove('active');
        }
    },
    
    /**
     * Show/hide loading
     */
    showLoading(show) {
        if (window.Components && typeof Components.setLoading === 'function') {
            Components.setLoading(show, 'Loading users...');
            return;
        }
        const loader = document.getElementById('loadingOverlay');
        if (loader) {
            if (show) {
                loader.classList.remove('hidden');
            } else {
                loader.classList.add('hidden');
            }
        }
    }
};

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => Users.init());
