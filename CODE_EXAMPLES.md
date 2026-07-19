# Code Examples & Reference

## Sorting Implementation

### 1. State Management

```javascript
// In Vouchers object state (line ~30)
sortConfig: {
    field: 'date',      // date, voucherNo, payee, grossAmount, netAmount
    direction: 'desc'   // asc or desc
},
```

### 2. Sorting Method - getSortedVouchers()

```javascript
getSortedVouchers() {
    const sorted = [...this.vouchers];  // Create copy
    const { field, direction } = this.sortConfig;
    
    sorted.sort((a, b) => {
        let aVal, bVal;
        
        // Map field names to actual values
        switch(field) {
            case 'voucherNo':
                aVal = (a.accountOrMail || '').toLowerCase();
                bVal = (b.accountOrMail || '').toLowerCase();
                break;
            case 'payee':
                aVal = (a.payee || '').toLowerCase();
                bVal = (b.payee || '').toLowerCase();
                break;
            case 'grossAmount':
                aVal = parseFloat(a.grossAmount || 0);
                bVal = parseFloat(b.grossAmount || 0);
                break;
            case 'netAmount':
                aVal = parseFloat((a.grossAmount || 0) - ((a.vat || 0) + (a.wht || 0) + (a.stampDuty || 0)));
                bVal = parseFloat((b.grossAmount || 0) - ((b.vat || 0) + (b.wht || 0) + (b.stampDuty || 0)));
                break;
            case 'date':
            default:
                aVal = a.date ? new Date(a.date).getTime() : 0;
                bVal = b.date ? new Date(b.date).getTime() : 0;
        }
        
        // Handle different data types
        if (typeof aVal === 'string') {
            return direction === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
        } else {
            return direction === 'asc' ? aVal - bVal : bVal - aVal;
        }
    });
    
    return sorted;
}
```

### 3. Update Sort Method - updateSort()

```javascript
updateSort(field) {
    // If same field clicked, toggle direction
    if (this.sortConfig.field === field) {
        this.sortConfig.direction = this.sortConfig.direction === 'asc' ? 'desc' : 'asc';
    } else {
        // New field selected
        this.sortConfig.field = field;
        // Amounts default to descending (highest first)
        this.sortConfig.direction = field === 'grossAmount' || field === 'netAmount' ? 'desc' : 'asc';
    }
    
    // Re-render table with new sort
    this.renderVoucherList();
}
```

---

## Table HTML Changes

### 1. Table Headers with Sort Icons

```html
<thead>
    <tr>
        ${canSelect ? '<th><input type="checkbox" id="selectAll"></th>' : ''}
        <th>S/N</th>
        <th class="sortable" onclick="Vouchers.updateSort('voucherNo')">
            Voucher No. ${sortIcon('voucherNo')}
        </th>
        <th class="sortable" onclick="Vouchers.updateSort('payee')">
            Payee ${sortIcon('payee')}
        </th>
        <th>Particular</th>
        <th class="sortable" onclick="Vouchers.updateSort('grossAmount')">
            Gross Amount ${sortIcon('grossAmount')}
        </th>
        <th class="sortable" onclick="Vouchers.updateSort('netAmount')">
            Net Amount ${sortIcon('netAmount')}
        </th>
        <th>Category</th>
        <th>Control No.</th>
        <th>Status</th>
        <th>Pmt Month</th>
        <th>Actions</th>
    </tr>
</thead>
```

### 2. Sort Icon Helper Function

```javascript
const sortIcon = (field) => {
    if (this.sortConfig.field !== field) 
        return '<i class="fas fa-arrows-alt-v" style="opacity: 0.4;"></i>';
    return this.sortConfig.direction === 'asc' 
        ? '<i class="fas fa-arrow-up"></i>' 
        : '<i class="fas fa-arrow-down"></i>';
};
```

### 3. Table Body with New Columns

```javascript
sortedVouchers.forEach((v, idx) => {
    const sn = startSN + idx + 1;
    const tr = document.createElement('tr');
    tr.dataset.row = v.rowIndex;
    
    // Calculate net amount
    const netAmount = (v.grossAmount || 0) - ((v.vat || 0) + (v.wht || 0) + (v.stampDuty || 0));

    tr.innerHTML = `
        // ... checkbox cell if needed ...
        <td class="sn-cell">${sn}</td>
        <td class="voucher-no-cell"><strong>${v.accountOrMail || '-'}</strong></td>
        <td class="payee-cell" title="${v.payee || ''}">${Utils.truncate(v.payee || '-', 25)}</td>
        <td class="particular-cell" title="${v.particular || ''}">
            <div class="text-truncate">${v.particular || '-'}</div>
        </td>
        <td class="amount-cell" style="text-align: right;">
            ${Utils.formatCurrency(v.grossAmount || 0)}
        </td>
        <td class="amount-cell" style="text-align: right;">
            ${Utils.formatCurrency(netAmount || 0)}
        </td>
        <td class="category-cell">${v.categories || '-'}</td>
        <td class="control-no-cell">${v.controlNumber || '<span class="text-muted">-</span>'}</td>
        <td class="status-cell">${Utils.getStatusBadge(v.status || '')}</td>
        <td>${v.pmtMonth || '-'}</td>
        <td>
            <div class="action-buttons">
                <button class="btn btn-sm btn-secondary" onclick="Vouchers.viewVoucher(${v.rowIndex})" title="View">
                    <i class="fas fa-eye"></i>
                </button>
                ${this.getActionButtons(v)}
            </div>
        </td>
    `;
    tbody.appendChild(tr);
});
```

---

## CSS Styling

### 1. Sortable Headers Styling

```css
table th.sortable {
    cursor: pointer;
    user-select: none;
    transition: all 0.2s ease;
}

table th.sortable:hover {
    background: #f0f1f3;
    color: var(--primary-color);
}
```

### 2. Table Cell Styling

```css
table td.sn-cell {
    font-weight: 600;
    color: #6c757d;
    width: 50px;
}

table td.voucher-no-cell {
    color: var(--primary-color);
    font-weight: 600;
}

table td.amount-cell {
    font-family: 'Courier New', monospace;
    font-weight: 500;
    color: var(--primary-color);
}

table td.category-cell {
    background: linear-gradient(135deg, rgba(26, 95, 42, 0.08) 0%, transparent 100%);
    border-radius: 4px;
    padding: 8px 10px;
}

table td.control-no-cell {
    color: #007bff;
    font-family: monospace;
    font-weight: 600;
}
```

### 3. Row Hover Effects

```css
table tbody tr {
    transition: all 0.2s ease;
}

table tbody tr:hover {
    background-color: rgba(26, 95, 42, 0.04);
    box-shadow: inset 0 0 0 1px rgba(26, 95, 42, 0.08);
}
```

---

## Sidebar Toggle Implementation

### 1. HTML Structure

```html
<button class="menu-toggle" id="menuToggle">
    <i class="fas fa-bars"></i>
</button>

<div class="app-container">
    <aside class="sidebar" id="sidebar"></aside>
    <main class="main-content">
        <!-- Content here -->
    </main>
</div>
```

### 2. JavaScript Event Listener

```javascript
document.getElementById('menuToggle')?.addEventListener('click', () => {
    document.getElementById('sidebar')?.classList.toggle('active');
});
```

### 3. CSS Transitions

```css
.sidebar {
    transform: translateX(-100%);  /* Hidden by default on mobile */
    transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
}

.sidebar.active {
    transform: translateX(0);  /* Visible when active */
}

.menu-toggle {
    position: fixed;
    top: 15px;
    left: 15px;
    background: var(--primary-color);
    border: none;
    padding: 10px 14px;
    border-radius: 8px;
    cursor: pointer;
    transition: all 0.3s ease;
    z-index: 1200;
}

.menu-toggle:hover {
    background: var(--primary-dark);
    transform: translateY(-2px);
}
```

---

## Net Amount Calculation

### Formula

```
Net Amount = Gross Amount - (VAT + WHT + Stamp Duty)

Example:
Gross Amount: 100,000.00
VAT:           10,000.00
WHT:            5,000.00
Stamp Duty:       500.00
                 ─────────
Net Amount:     84,500.00 ← Displayed in table
```

### Code Implementation

```javascript
const netAmount = (v.grossAmount || 0) - 
                  ((v.vat || 0) + (v.wht || 0) + (v.stampDuty || 0));

// Then format it
tr.innerHTML += `<td class="amount-cell" style="text-align: right;">
    ${Utils.formatCurrency(netAmount || 0)}
</td>`;
```

---

## Usage Examples

### Example 1: Sort by Payee Name

```javascript
// User clicks "Payee" column header
Vouchers.updateSort('payee');

// Result: Vouchers sorted A-Z by payee name

// Click again to reverse
Vouchers.updateSort('payee');  // Now Z-A
```

### Example 2: Sort by Amount (Highest First)

```javascript
// User clicks "Gross Amount" column header
Vouchers.updateSort('grossAmount');

// Result: Vouchers sorted by amount (highest → lowest)
// Amount defaults to descending for typical use case
```

### Example 3: Sort by Voucher Number

```javascript
// User clicks "Voucher No." column header
Vouchers.updateSort('voucherNo');

// Result: Vouchers sorted alphabetically A-Z
// Click again for Z-A
```

### Example 4: Toggle Sidebar

```javascript
// User clicks ☰ menu button
document.getElementById('menuToggle').click();

// Result: Sidebar slides in from left with smooth animation
// Click again to hide it
```

---

## Checking Sort State

```javascript
// Get current sort configuration
console.log(Vouchers.sortConfig);

// Output examples:
// { field: 'payee', direction: 'asc' }
// { field: 'grossAmount', direction: 'desc' }
// { field: 'netAmount', direction: 'asc' }

// Check which field is currently sorted
if (Vouchers.sortConfig.field === 'payee') {
    console.log('Currently sorted by payee');
}

// Check sort direction
if (Vouchers.sortConfig.direction === 'desc') {
    console.log('Descending order (Z→A or High→Low)');
}
```

---

## Performance Tips

### Efficient Sorting
```javascript
// getSortedVouchers() uses spread operator for non-destructive sort
const sorted = [...this.vouchers];  // ✅ Creates copy
sorted.sort(...);                    // ✅ Sorts copy
return sorted;                       // ✅ Returns sorted copy

// Original this.vouchers remains unchanged
```

### Rendering Optimization
```javascript
// Uses DocumentFragment for batch DOM updates
const fragment = document.createDocumentFragment();
// ... append all elements to fragment ...
container.appendChild(fragment);  // ✅ Single DOM update
```

---

## Troubleshooting

### "Sort icons not showing"
**Check**: Font Awesome CSS is loaded
```html
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
```

### "Sidebar not toggling"
**Check**: JavaScript is enabled and event listener is attached
```javascript
// Verify in console
console.log(document.getElementById('menuToggle')); // Should return element
console.log(document.getElementById('sidebar'));     // Should return element
```

### "Sort not working on new column"
**Check**: Column class is marked as sortable
```html
<th class="sortable" onclick="Vouchers.updateSort('payee')">
    <!-- Needs class="sortable" -->
</th>
```

---

## References

- **Font Awesome Icons**: https://fontawesome.com/icons
- **CSS Grid**: https://developer.mozilla.org/en-US/docs/Web/CSS/CSS_Grid_Layout
- **Array.sort()**: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/sort
- **localStorage.collapse states**: Not used (state is in memory)

---

*Documentation Last Updated: March 31, 2026*
