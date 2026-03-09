/**
 * PAYABLE VOUCHER 2026 - Authentication Module
 */

const Auth = {
    /**
     * Stores session data in localStorage
     */
    saveSession(token, user) {
        const sessionData = {
            token: token,
            user: user,
            timestamp: new Date().getTime()
        };
        localStorage.setItem(CONFIG.SESSION_KEY, JSON.stringify(sessionData));
    },

    /**
     * Gets current session from localStorage
     */
    getSession() {
        const sessionStr = localStorage.getItem(CONFIG.SESSION_KEY);
        if (!sessionStr) return null;

        try {
            return JSON.parse(sessionStr);
        } catch (e) {
            return null;
        }
    },

    /**
     * Gets current token
     */
    getToken() {
        const session = this.getSession();
        return session ? session.token : null;
    },

    /**
     * Gets current user
     */
    getUser() {
        const session = this.getSession();
        return session ? session.user : null;
    },

    /**
     * Checks if user is logged in
     */
    isLoggedIn() {
        return this.getToken() !== null;
    },

    /**
     * Clears session (logout)
     */
    clearSession() {
        localStorage.removeItem(CONFIG.SESSION_KEY);
        sessionStorage.removeItem('pv2026_permissions'); // Clear permissions cache
        localStorage.removeItem('pv2026_auth_validated');
    },

    /**
     * Login function
     */
    async login(email, password) {
        try {
            const url = `${CONFIG.API_URL}?action=login&email=${encodeURIComponent(email)}&password=${encodeURIComponent(password)}`;

            const response = await fetch(url);
            const result = await response.json();

            if (result.success) {
                this.saveSession(result.token, result.user);
            }

            return result;
        } catch (error) {
            console.error('Login error:', error);
            return { success: false, error: 'Network error. Please check your connection.' };
        }
    },

    /**
     * Validates current session with server
     */
    async validateSession() {
        const token = this.getToken();
        if (!token) return { success: false };

        // Check if we have a recent validation cache (15 minutes = 900000 ms)
        const lastValidated = localStorage.getItem('pv2026_auth_validated');
        if (lastValidated && (Date.now() - parseInt(lastValidated, 10) < 900000)) {
            return { success: true, cached: true };
        }

        try {
            const url = `${CONFIG.API_URL}?action=validateSession&token=${encodeURIComponent(token)}`;
            const response = await fetch(url);
            const result = await response.json();

            if (!result.success) {
                this.clearSession();
            } else {
                localStorage.setItem('pv2026_auth_validated', Date.now().toString());
            }

            return result;
        } catch (error) {
            console.error('Session validation error:', error);
            return { success: false, error: 'Network error' };
        }
    },

    /**
     * Logout function
     */
    async logout() {
        const token = this.getToken();

        if (token) {
            try {
                const url = `${CONFIG.API_URL}?action=logout&token=${encodeURIComponent(token)}`;
                await fetch(url);
            } catch (error) {
                console.error('Logout error:', error);
            }
        }

        this.clearSession();
        window.location.href = 'index.html';
    },

    /**
     * Changes user password
     */
    async changePassword(oldPassword, newPassword) {
        const token = this.getToken();
        if (!token) return { success: false, error: 'Not logged in' };

        try {
            const response = await fetch(CONFIG.API_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    action: 'changePassword',
                    token: token,
                    oldPassword: oldPassword,
                    newPassword: newPassword
                })
            });

            return await response.json();
        } catch (error) {
            console.error('Change password error:', error);
            return { success: false, error: 'Network error' };
        }
    },

    /**
     * Checks if user has specific role
     */
    hasRole(roles) {
        const user = this.getUser();
        if (!user) return false;

        if (user.role === CONFIG.ROLES.ADMIN) return true;

        if (Array.isArray(roles)) {
            return roles.includes(user.role);
        }

        return user.role === roles;
    },

    /**
     * Gets role permissions
     */
    async getPermissions() {
        const token = this.getToken();
        if (!token) return null;

        // Check cache first
        const cached = sessionStorage.getItem('pv2026_permissions');
        if (cached) {
            return JSON.parse(cached);
        }

        try {
            const url = `${CONFIG.API_URL}?action=getRolePermissions&token=${encodeURIComponent(token)}`;
            const response = await fetch(url);
            const result = await response.json();

            if (result.success) {
                sessionStorage.setItem('pv2026_permissions', JSON.stringify(result.permissions));
            }
            return result.success ? result.permissions : null;
        } catch (error) {
            console.error('Get permissions error:', error);
            return null;
        }
    },

    /**
     * Protects a page - redirects to login if not authenticated
     */
    async requireAuth() {
        if (!this.isLoggedIn()) {
            window.location.href = 'index.html';
            return false;
        }

        // Validate session with server
        const validation = await this.validateSession();
        if (!validation.success) {
            window.location.href = 'index.html';
            return false;
        }

        return true;
    }
};