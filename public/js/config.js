/**
 * PAYABLE VOUCHER 2026 - Configuration
 * Federal Medical Centre, Abeokuta
 */

const CONFIG = {
    // IMPORTANT: Replace this with your actual Apps Script Web App URL
    API_URL: 'https://script.google.com/macros/s/AKfycbwqIVfk4jUGA5MsTD2r5Ny8IXe34dQl3AOtBEy-aBHWBfP-jwTp8X02UZVQKt19XkxyYQ/exec',

    // App settings
    APP_NAME: 'PAYABLE VOUCHER 2026',
    ORGANIZATION: 'Federal Medical Centre, Abeokuta',
    DEPARTMENT: 'Finance & Accounts Department',

    // Session settings
    SESSION_KEY: 'pv2026_session',

    // Valid statuses
    STATUSES: ['Unpaid', 'Paid', 'Cancelled', 'Pending Deletion'],

    // Months
    MONTHS: [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ],

    // Roles
    ROLES: {
        PAYABLE_STAFF: 'Payable Unit Staff',
        PAYABLE_HEAD: 'Payable Unit Head',
        CPO: 'CPO',
        AUDIT: 'Audit Unit',
        DDFA: 'DDFA',
        DFA: 'DFA',
        ADMIN: 'ADMIN'
    },

    // Years for lookup
    YEARS: ['2026', '2025', '2024', '2023', '<2023']
};

// Make sure to update the API_URL above with your actual deployment URL!