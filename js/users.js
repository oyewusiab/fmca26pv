/**
 * PAYABLE VOUCHER 2026 - User Management Module
 * Admin only functionality
 */

const Users = {
    // State
    users: [],
    selectedUser: null,
    isEditMode: false,
    
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
        
        // Load users
        await this.loadUsers();
        
        // Setup event listeners
        this.setupEventListeners();
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
    
    /**
     * Load users from backend
     */
    async loadUsers() {
        this.showLoading(true);
        
        try {
            const result = await API.getUsers();
            
            if (result.success) {
                this.users = result.users;
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
     * Render users list
     */
    renderUsersList() {
        const container = document.getElementById('usersList');
        if (!container) return;
        
        // Get search term
        const searchTerm = document.getElementById('searchUsers')?.value.toLowerCase() || '';
        
        // Filter users
        let filteredUsers = this.users;
        if (searchTerm) {
            filteredUsers = this.users.filter(user => 
                user.name.toLowerCase().includes(searchTerm) ||
                user.email.toLowerCase().includes(searchTerm) ||
                user.role.toLowerCase().includes(searchTerm)
            );
        }
        
        // Update count
        document.getElementById('userCount').textContent = `${filteredUsers.length} user(s)`;
        
        if (filteredUsers.length === 0) {
            container.innerHTML = Components.getEmptyState('No users found', 'fa-users');
            return;
        }
        
        let html = `
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Name</th>
                            <th>Email</th>
                            <th>Role</th>
                            <th>Status</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        filteredUsers.forEach(user => {
            const statusBadge = user.active === true || user.active === 'TRUE' || user.active === 'true'
                ? '<span class="badge badge-paid">Active</span>'
                : '<span class="badge badge-cancelled">Inactive</span>';
            
            html += `
                <tr>
                    <td>
                        <div style="display: flex; align-items: center; gap: 10px;">
                            <div class="user-avatar-small">${Utils.getInitials(user.name)}</div>
                            <strong>${user.name}</strong>
                        </div>
                    </td>
                    <td>${user.email}</td>
                    <td><span class="role-badge role-${user.role.toLowerCase().replace(/\s+/g, '-')}">${user.role}</span></td>
                    <td>${statusBadge}</td>
                    <td>
                        <div class="action-buttons">
                            <button class="btn btn-sm btn-primary" onclick="Users.editUser(${user.rowIndex})" title="Edit">
                                <i class="fas fa-edit"></i>
                            </button>
                            <button class="btn btn-sm btn-secondary" onclick="Users.resetPassword(${user.rowIndex})" title="Reset Password">
                                <i class="fas fa-key"></i>
                            </button>
                            <button class="btn btn-sm btn-danger" onclick="Users.deleteUser(${user.rowIndex})" title="Delete">
                                <i class="fas fa-trash"></i>
                            </button>
                        </div>
                    </td>
                </tr>
            `;
        });
        
        html += '</tbody></table></div>';
        container.innerHTML = html;
    },
    
    /**
     * Filter users based on search
     */
    filterUsers() {
        this.renderUsersList();
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
        
        title.textContent = this.isEditMode ? 'Edit User' : 'Add New User';
        form.reset();
        
        // Populate role dropdown
        const roleSelect = document.getElementById('formRole');
        roleSelect.innerHTML = '<option value="">Select Role</option>';
        Object.values(CONFIG.ROLES).forEach(role => {
            roleSelect.innerHTML += `<option value="${role}">${role}</option>`;
        });
        
        if (user) {
            document.getElementById('formName').value = user.name || '';
            document.getElementById('formEmail').value = user.email || '';
            document.getElementById('formRole').value = user.role || '';
            document.getElementById('formActive').checked = user.active === true || user.active === 'TRUE' || user.active === 'true';
        } else {
            document.getElementById('formActive').checked = true;
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
        
        const userData = {
            name: document.getElementById('formName').value.trim(),
            email: document.getElementById('formEmail').value.trim(),
            role: document.getElementById('formRole').value,
            active: document.getElementById('formActive').checked
        };
        
        if (!userData.name || !userData.email || !userData.role) {
            Utils.showToast('Please fill all required fields', 'error');
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
            `Reset password for ${user.name}?\n\nThe password will be reset to: Welcome123`,
            'Reset Password'
        );
        
        if (!confirm) return;
        
        this.showLoading(true);
        
        try {
            const result = await API.updateUser(rowIndex, { resetPassword: true });
            
            if (result.success) {
                Utils.showToast(result.message || 'Password reset successfully', 'success');
            } else {
                Utils.showToast(result.error || 'Failed to reset password', 'error');
            }
        } catch (error) {
            console.error('Reset password error:', error);
            Utils.showToast('Error resetting password', 'error');
        }
        
        this.showLoading(false);
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