/**
 * Legacy notifications page shim.
 * Real page logic now lives in:
 * - public/js/notifications.js
 * - public/js/action-items.js
 * - public/js/announcements.js
 */

document.addEventListener('DOMContentLoaded', () => {
  if (window.Notifications && typeof window.Notifications.initPage === 'function') {
    window.Notifications.initPage();
  }
});

