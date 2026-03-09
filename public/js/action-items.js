/**
 * Action Items Logic
 * Handles fetching, rendering, and interactions for the Action Items page.
 */

async function loadActionItemsFullPage() {

  const list = document.getElementById("actionItemsList");
  const card = document.getElementById("actionItemsCard");
  const countBadge = document.getElementById("actionItemsCount");

  if (!list) return; // not on notifications page

  list.innerHTML = `<div class="text-center p-4">
      <div class="spinner"></div>
      <p class="text-muted mt-2">Loading action items...</p>
    </div>`;

  try {

    const res = await API.getActionItems({
      unit: document.getElementById("actionItemsUnitFilter")?.value,
      status: document.getElementById("actionItemsStatusFilter")?.value
    });

    if (!res.success) {
      list.innerHTML = `<div class="alert alert-danger">${res.error}</div>`;
      return;
    }

    const items = res.items || [];
    const count = res.count || 0;

    countBadge.textContent = count;

    if (count === 0) {
      card.style.display = "none";
      return;
    }

    card.style.display = "block";

    list.innerHTML = items.map(item => renderPremiumRow(item)).join("");

  } catch (err) {
    console.error(err);
    list.innerHTML = `<div class="alert alert-danger">Failed to load Action Items.</div>`;
  }
}

async function refreshSidebarBadge() {
  const badge = document.getElementById("navActionBadge");
  if (!badge) return;

  const res = await API.getActionItemCount();
  if (!res.success) return;

  badge.textContent = res.count;
  badge.style.display = res.count > 0 ? "inline-block" : "none";
}

function renderPremiumRow(item) {

  const iconClass =
    item.severity === "danger" ? "icon-wrap-danger" :
      item.severity === "warning" ? "icon-wrap-warning" :
        "icon-wrap-info";

  const dotClass =
    item.severity === "danger" ? "dot-danger" :
      item.severity === "warning" ? "dot-warning" :
        "dot-info";

  return `
    <div class="action-item-row">
      <div class="action-item-icon-wrap ${iconClass}">
        <i class="fas fa-exclamation-circle"></i>
      </div>

      <div class="action-item-content">

        <div class="action-item-header">
          <div class="action-item-title">${item.title}</div>
        </div>

        <div class="action-item-msg">${item.message}</div>

        <div class="action-item-badges">
          <span class="small-badge"><strong>Voucher:</strong> ${item.voucherNumber}</span>
          <span class="small-badge"><strong>Payee:</strong> ${item.payee}</span>
          <span class="small-badge"><strong>Amount:</strong> ${Utils.formatCurrency(item.amount)}</span>
        </div>

      </div>

      <div class="action-item-actions">
        <a class="btn-view" onclick="handleViewAction('${item.voucherNumber}')">
          <i class="fas fa-eye"></i> View
        </a>
      </div>
    </div>
  `;
}

document.addEventListener("DOMContentLoaded", () => {
  setTimeout(loadActionItemsFullPage, 150);
});
