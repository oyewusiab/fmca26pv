import { defineConfig } from 'vite';
import { resolve } from 'path';

export default defineConfig({
    build: {
        rollupOptions: {
            input: {
                main: resolve(__dirname, 'index.html'),
                dashboard: resolve(__dirname, 'dashboard.html'),
                vouchers: resolve(__dirname, 'vouchers.html'),
                reports: resolve(__dirname, 'reports.html'),
                users: resolve(__dirname, 'users.html'),
                profile: resolve(__dirname, 'profile.html'),
                auditTrail: resolve(__dirname, 'audit-trail.html'),
                forgotPassword: resolve(__dirname, 'forgot-password.html'),
                notifications: resolve(__dirname, 'notifications.html')
            }
        }
    }
});
