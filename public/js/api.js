/**
 * PAYABLE VOUCHER 2026 - API Module
 * Handles communication with the Google Apps Script backend
 *
 * Safe & compatible rewrite:
 * - No duplicate method names
 * - POST uses text/plain to avoid CORS preflight (Apps Script friendly)
 * - Backward-compatible wrapper signatures where earlier code variants differed
 */

(function () {
  // Prevent accidental double-load
  if (window.API) return;

  const API = {
    // ==================== CACHING & STATE ====================
    _cachePrefix: 'pv2026_api_',
    _inflightRequests: new Map(),

    // Time-to-live in seconds for various endpoints
    _ttls: {
      'getDashboardStats': 300,      // 5 minutes
      'getSummary': 300,             // 5 minutes
      'getAllYearsSummary': 3600,    // 1 hour
      'getDebtProfile': 600,         // 10 minutes
      'getQuickStats': 300,          // 5 minutes
      'getVouchers': 60,             // 1 minute (list views)
      'getUsers': 3600,              // 1 hour
      'getCategories': 86400,        // 24 hours
      'getActionItems': 60,          // 1 minute
      'getRolePermissions': 3600,    // 1 hour
      'getTaxSummary': 600,          // 10 minutes
      'getTaxByCategory': 900,       // 15 minutes
      'getTaxByMonth': 900,          // 15 minutes
      'getTaxByPayee': 600,          // 10 minutes
      'getTaxPayments': 300,         // 5 minutes
      'getTaxSchedule': 300          // 5 minutes
    },

    _getCacheKey(action, params) {
      return `${action}:${JSON.stringify(params)}`;
    },

    _getFromCache(key) {
      try {
        const item = localStorage.getItem(this._cachePrefix + key);
        if (!item) return null;
        const record = JSON.parse(item);
        if (Date.now() > record.expiry) {
          localStorage.removeItem(this._cachePrefix + key);
          return null;
        }
        return record.data;
      } catch (e) { return null; }
    },

    _setCache(key, data, ttlSeconds) {
      if (!ttlSeconds) return;
      try {
        const record = { data, expiry: Date.now() + (ttlSeconds * 1000) };
        localStorage.setItem(this._cachePrefix + key, JSON.stringify(record));
      } catch (e) { console.warn('Cache quota exceeded'); }
    },

    _invalidateCache(pattern = '') {
      Object.keys(localStorage).forEach(key => {
        if (key.startsWith(this._cachePrefix) && key.includes(pattern)) {
          localStorage.removeItem(key);
        }
      });
    },

    // ==================== CORE REQUESTS ====================

    /**
     * Makes a GET request to the API with SWR (Stale-While-Revalidate)
     * @param {string} action
     * @param {Object} params
     * @param {number} [customTtl] Optional override for TTL
     */
    async get(action, params = {}, customTtl = null) {
      try {
        params = params || {};
        const token = Auth.getToken?.();
        let url = `${CONFIG.API_URL}?action=${encodeURIComponent(action)}`;

        if (token) url += `&token=${encodeURIComponent(token)}`;

        for (const key in params) {
          const val = params[key];
          if (val !== undefined && val !== null) {
            url += `&${encodeURIComponent(key)}=${encodeURIComponent(val)}`;
          }
        }

        const cacheKey = this._getCacheKey(action, params);
        const ttl = customTtl !== null ? customTtl : (this._ttls[action] || 0);

        // 1. SWR: Immediately return cached data if available
        let cachedData = null;
        if (ttl > 0) {
          cachedData = this._getFromCache(cacheKey);
        }

        // 2. Fetch fresh data in the background
        const fetchPromise = (async () => {
          if (this._inflightRequests.has(cacheKey)) {
            return await this._inflightRequests.get(cacheKey);
          }

          const requestPromise = fetch(url, { method: "GET", redirect: "follow" });

          this._inflightRequests.set(cacheKey, requestPromise.then(async r => {
            try {
              const text = await r.clone().text();
              return JSON.parse(text);
            } catch {
              return {};
            }
          }));

          try {
            const response = await requestPromise;
            const text = await response.text();
            let result;
            try {
              result = JSON.parse(text);
            } catch (e) {
              console.error(`API GET non-JSON response (${action}):`, text);
              return { success: false, error: "Invalid response from server" };
            }

            if (result?.error && String(result.error).includes("Session expired")) {
              Auth.clearSession?.();
              window.location.href = "index.html";
              return { success: false, error: "Session expired" };
            }

            // Set Cache
            if (result.success && ttl > 0) {
              this._setCache(cacheKey, result, ttl);

              // Dispatch event to let UI know fresh data arrived (if they want to re-render without reloading)
              document.dispatchEvent(new CustomEvent('apiDataUpdated', {
                detail: { action, params, data: result }
              }));
            }

            return result;
          } finally {
            this._inflightRequests.delete(cacheKey);
          }
        })();

        // If we have stale cache, return it instantly! The background promise will silently update cache.
        if (cachedData) {
          // Let the caller know this is cached data so they can optionally show a loader
          cachedData._isCached = true;
          return cachedData;
        }

        // Otherwise, wait for the network (first load)
        return await fetchPromise;

      } catch (error) {
        this._inflightRequests.delete(this._getCacheKey(action, params));
        console.error(`API GET error (${action}):`, error);
        return { success: false, error: "Network error. Please check your connection." };
      }
    },

    /**
     * Makes a POST request to the API
     * Uses text/plain to avoid CORS preflight issues with Google Apps Script
     * @param {string} action
     * @param {Object} data
     */
    async post(action, data = {}) {
      try {
        const token = Auth.getToken?.();

        const payload = {
          action,
          token,
          ...data,
        };

        const response = await fetch(CONFIG.API_URL, {
          method: "POST",
          redirect: "follow",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify(payload),
        });

        const text = await response.text();
        let result;
        try {
          result = JSON.parse(text);
        } catch (e) {
          console.error(`API POST non-JSON response (${action}):`, text);
          return { success: false, error: "Invalid response from server" };
        }

        if (result?.error && String(result.error).includes("Session expired")) {
          Auth.clearSession?.();
          window.location.href = "index.html";
          return { success: false, error: "Session expired" };
        }

        // Invalidate cache on successful write
        if (result.success) {
          this._invalidateCache(); // Simple strategy: clear all API cache to ensure freshness
        }

        return result;
      } catch (error) {
        console.error(`API POST error (${action}):`, error);
        return { success: false, error: "Network error. Please check your connection." };
      }
    },

    // ==================== VOUCHERS ====================

    /**
     * Gets vouchers (supports filters + optional pagination if backend supports it)
     */
    async getVouchers(year = "2026", filters = null, page = 1, pageSize = 50) {
      const params = { year, page, pageSize };
      if (filters) params.filters = JSON.stringify(filters);
      return await this.get("getVouchers", params);
    },

    async getVoucherByRow(rowIndex, year = "2026") {
      return await this.get("getVoucherByRow", { rowIndex, year });
    },

    async createVoucher(voucher) {
      return await this.post("createVoucher", { voucher });
    },

    async updateVoucher(rowIndex, voucher) {
      return await this.post("updateVoucher", { rowIndex, voucher });
    },

    async updateStatus(rowIndex, status, pmtMonth = null) {
      return await this.post("updateStatus", { rowIndex, status, pmtMonth });
    },

    async batchUpdateStatus(controlNumber, status, pmtMonth = null) {
      return await this.post("batchUpdateStatus", { controlNumber, status, pmtMonth });
    },

    async assignControlNumber(rowIndexes, controlNumber) {
      return await this.post("assignControlNumber", { rowIndexes, controlNumber });
    },

    /**
     * releaseSelectedVouchers backend signature is (token, rowIndexes, targetUnit)
     * Some older front-end code called it with 3 args (rowIndexes, controlNumber, targetUnit).
     * This wrapper supports BOTH safely.
     */
    async releaseSelectedVouchers(rowIndexes, arg2, arg3) {
      let targetUnit;

      // Called as (rowIndexes, targetUnit)
      if (arg3 === undefined) {
        targetUnit = arg2;
      } else {
        // Called as (rowIndexes, controlNumber, targetUnit) — ignore controlNumber for this endpoint
        targetUnit = arg3;
      }

      return await this.post("releaseSelectedVouchers", {
        rowIndexes,
        targetUnit,
      });
    },

    async releaseVouchersWithNotification(rowIndexes, controlNumber, targetUnit) {
      return await this.post("releaseVouchersWithNotification", {
        rowIndexes,
        controlNumber,
        targetUnit,
      });
    },

    /**
     * Enhanced release flow (your Vouchers module uses API.post('releaseVouchers', ...))
     */
    async releaseVouchers(payload) {
      return await this.post("releaseVouchers", payload);
    },

    // ==================== DELETION WORKFLOW ====================

    async requestDelete(rowIndex, reason, previousStatus) {
      return await this.post("requestDelete", { rowIndex, reason, previousStatus });
    },

    async cancelDeleteRequest(rowIndex) {
      return await this.post("cancelDeleteRequest", { rowIndex });
    },

    async approveDelete(rowIndex) {
      return await this.post("approveDelete", { rowIndex });
    },

    async rejectDelete(rowIndex, reason = "") {
      return await this.post("rejectDelete", { rowIndex, reason });
    },

    async deleteVoucher(rowIndex) {
      return await this.post("deleteVoucher", { rowIndex });
    },

    async getPendingDeletions() {
      return await this.get("getPendingDeletions");
    },

    // ==================== LOOKUP ====================

    async lookupVoucher(voucherNumber) {
      return await this.get("lookupVoucher", { voucherNumber });
    },

    async createVoucherFromLookup(lookupResult, additionalData = {}) {
      return await this.post("createVoucherFromLookup", { lookupResult, additionalData });
    },

    // ==================== NOTIFICATIONS ====================

    async getNotifications(onlyUnread = false) {
      return await this.get("getNotifications", { onlyUnread });
    },

    async markNotificationRead(rowIndex) {
      return await this.post("markNotificationRead", { rowIndex });
    },

    async markAllNotificationsRead() {
      return await this.post("markAllNotificationsRead", {});
    },

    // ==================== AUDIT ====================

    async getAuditTrail(limit = 50, offset = 0) {
      return await this.get("getAuditTrail", { limit, offset });
    },

    // ==========================================
    // ANNOUNCEMENTS / MESSAGES endpoints
    // ==========================================
    async getActiveAnnouncements(location = null) {
        const params = {};
        if (location) params.location = location;
        return await this.get('getActiveAnnouncements', params, 30);
    },

    async dismissAnnouncement(announcementId) {
        return this.post('dismissAnnouncement', { announcementId });
    },

    async createAnnouncement(announcement) {
        // announcement expects { message, targets, locations, expiresAt, allowDismiss }
        return this.post('createAnnouncement', { announcement });
    },

    // ==================== SUMMARY / REPORTS ====================

    async getDashboardStats() {
      return await this.get("getDashboardStats");
    },

    async getSummary(year = "2026") {
      return await this.get("getSummary", { year });
    },

    async getAllYearsSummary() {
      return await this.get("getAllYearsSummary");
    },

    async getDebtProfile() {
      return await this.get("getDebtProfile");
    },

    async getTaxSummary(year = "2026") {
      return await this.get("getTaxSummary", { year });
    },

    async getTaxByCategory(year = "2026") {
      return await this.get("getTaxByCategory", { year });
    },

    async getTaxByMonth(year = "2026") {
      return await this.get("getTaxByMonth", { year });
    },

    async getTaxByPayee(year = "2026", filters = null) {
      return await this.get("getTaxByPayee", { year, filters });
    },

    async getTaxPayments(year = "2026") {
      return await this.get("getTaxPayments", { year });
    },

    async recordTaxPayment(payment) {
      const result = await this.post("recordTaxPayment", { payment });
      if (result && result.success) {
        this._invalidateCache("getTax");
      }
      return result;
    },

    async updateTaxPayment(paymentId, updates) {
      const result = await this.post("updateTaxPayment", { paymentId, updates });
      if (result && result.success) {
        this._invalidateCache("getTax");
      }
      return result;
    },

    async deleteTaxPayment(paymentId) {
      const result = await this.post("deleteTaxPayment", { paymentId });
      if (result && result.success) {
        this._invalidateCache("getTax");
      }
      return result;
    },

    async getTaxSchedule(year = "2026") {
      return await this.get("getTaxSchedule", { year });
    },

    async createTaxSchedule(schedule) {
      return await this.post("createTaxSchedule", { schedule });
    },

    async updateTaxSchedule(scheduleId, updates) {
      return await this.post("updateTaxSchedule", { scheduleId, updates });
    },

    async getQuickStats() {
      return await this.get("getQuickStats");
    },

    // ==================== ACTION ITEMS ====================

    async getActionItems(params = {}) {
      return await this.get("getActionItems", { ...params });
    },

    async getActionItemCount(params = {}) {
      return await this.get("getActionItemCount", params);
    },

    async getActionItemSettings() {
      return await this.get("getActionItemSettings");
    },

    async saveActionItemSettings(settings) {
      return await this.post("saveActionItemSettings", { settings });
    },

    // ==================== USERS / SYSTEM ====================

    async getUsers() {
      return await this.get("getUsers");
    },

    async createUser(user) {
      return await this.post("createUser", { user });
    },

    async updateUser(rowIndex, user) {
      return await this.post("updateUser", { rowIndex, user });
    },

    async deleteUser(rowIndex) {
      return await this.post("deleteUser", { rowIndex });
    },

    async getCategories() {
      return await this.get("getCategories");
    },

    async getRolePermissions() {
      return await this.get("getRolePermissions");
    },
  };

  window.API = API;
})();

// ===== API Add-ons (safe) =====
(function () {
  const API = window.API;
  if (!API) return;

  // Password reset
  if (!API.requestPasswordReset) {
    API.requestPasswordReset = async function (identifier) {
      return await this.post('requestPasswordReset', { identifier });
    };
  }

  if (!API.resetPasswordWithOtp) {
    API.resetPasswordWithOtp = async function (identifier, otp, newPassword) {
      return await this.post('resetPasswordWithOtp', { identifier, otp, newPassword });
    };
  }

  // Profile
  if (!API.getMyProfile) {
    API.getMyProfile = async function () {
      return await this.get('getMyProfile');
    };
  }

  if (!API.updateMyProfile) {
    API.updateMyProfile = async function (profile) {
      return await this.post('updateMyProfile', { profile });
    };
  }
})();

// ==================== UTILS (SAFE MERGE) ====================
// This avoids breaking your existing pages if Utils is defined elsewhere.
(function () {
  window.Utils = window.Utils || {};
  const Utils = window.Utils;

  Utils.formatCurrency =
    Utils.formatCurrency ||
    function (amount) {
      if (amount === null || amount === undefined || isNaN(amount)) return "₦0.00";
      return (
        "₦" +
        Number(amount).toLocaleString("en-NG", {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        })
      );
    };

  Utils.formatNumber =
    Utils.formatNumber ||
    function (num) {
      if (num === null || num === undefined || isNaN(num)) return "0";
      return Number(num).toLocaleString("en-NG");
    };

  Utils.parseCurrency =
    Utils.parseCurrency ||
    function (str) {
      if (typeof str === "number") return str;
      if (!str) return 0;
      return parseFloat(String(str).replace(/[₦,\s]/g, "")) || 0;
    };

  Utils.formatDate =
    Utils.formatDate ||
    function (dateStr) {
      if (!dateStr) return "-";
      try {
        const date = new Date(dateStr);
        return date.toLocaleDateString("en-NG", { year: "numeric", month: "short", day: "numeric" });
      } catch {
        return dateStr;
      }
    };

  Utils.formatDateTime =
    Utils.formatDateTime ||
    function (dateStr) {
      if (!dateStr) return "-";
      try {
        const date = new Date(dateStr);
        return date.toLocaleString("en-NG", {
          year: "numeric",
          month: "short",
          day: "numeric",
          hour: "2-digit",
          minute: "2-digit",
        });
      } catch {
        return dateStr;
      }
    };

  Utils.getStatusBadge =
    Utils.getStatusBadge ||
    function (status) {
      const statusLower = (status || "unpaid").toLowerCase();
      let badgeClass = "badge-unpaid";
      switch (statusLower) {
        case "paid":
          badgeClass = "badge-paid";
          break;
        case "cancelled":
          badgeClass = "badge-cancelled";
          break;
        case "pending deletion":
          badgeClass = "badge-pending";
          break;
      }
      return `<span class="badge ${badgeClass}">${status || "Unpaid"}</span>`;
    };

  Utils.truncate =
    Utils.truncate ||
    function (text, length = 30) {
      if (!text) return "";
      if (text.length <= length) return text;
      return text.substring(0, length) + "...";
    };

  Utils.getInitials =
    Utils.getInitials ||
    function (name) {
      if (!name) return "??";
      const parts = name.trim().split(" ");
      if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
      return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    };

  Utils.debounce =
    Utils.debounce ||
    function (func, wait = 300) {
      let timeout;
      return function (...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
      };
    };

  Utils.getToastIcon =
    Utils.getToastIcon ||
    function (type) {
      switch (type) {
        case "success":
          return "fa-check-circle";
        case "error":
          return "fa-exclamation-circle";
        case "warning":
          return "fa-exclamation-triangle";
        default:
          return "fa-info-circle";
      }
    };

  Utils.showToast =
    Utils.showToast ||
    function (message, type = "info", duration = 3000) {
      document.querySelectorAll(".toast").forEach((t) => t.remove());

      const toast = document.createElement("div");
      toast.className = `toast toast-${type}`;
      toast.innerHTML = `
        <div class="toast-content">
          <i class="fas ${Utils.getToastIcon(type)}"></i>
          <span>${message}</span>
        </div>
      `;

      if (!document.getElementById("toastStyles")) {
        const style = document.createElement("style");
        style.id = "toastStyles";
        style.textContent = `
          .toast{position:fixed;top:20px;right:20px;padding:15px 20px;border-radius:8px;color:#fff;font-size:14px;z-index:9999;animation:slideIn .3s ease;box-shadow:0 4px 12px rgba(0,0,0,.15);max-width:400px}
          .toast-content{display:flex;align-items:center;gap:10px}
          .toast-success{background:#28a745}
          .toast-error{background:#dc3545}
          .toast-warning{background:#ffc107;color:#333}
          .toast-info{background:#17a2b8}
          @keyframes slideIn{from{transform:translateX(100%);opacity:0}to{transform:translateX(0);opacity:1}}
          @keyframes slideOut{from{transform:translateX(0);opacity:1}to{transform:translateX(100%);opacity:0}}
        `;
        document.head.appendChild(style);
      }

      document.body.appendChild(toast);

      setTimeout(() => {
        toast.style.animation = "slideOut 0.3s ease";
        setTimeout(() => toast.remove(), 300);
      }, duration);
    };

  Utils.confirm =
    Utils.confirm ||
    async function (message, title = "Confirm") {
      return new Promise((resolve) => {
        const modal = document.createElement("div");
        modal.className = "modal-overlay active";
        modal.innerHTML = `
          <div class="modal" style="max-width:400px;">
            <div class="modal-header">
              <h3 class="modal-title">${title}</h3>
            </div>
            <div class="modal-body">
              <p style="white-space:pre-line;">${message}</p>
            </div>
            <div class="modal-footer">
              <button class="btn btn-secondary" id="confirmCancel">Cancel</button>
              <button class="btn btn-primary" id="confirmOk">Confirm</button>
            </div>
          </div>
        `;
        document.body.appendChild(modal);

        modal.querySelector("#confirmOk").addEventListener("click", () => {
          modal.remove();
          resolve(true);
        });

        modal.querySelector("#confirmCancel").addEventListener("click", () => {
          modal.remove();
          resolve(false);
        });
      });
    };
})();
