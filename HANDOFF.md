# Stockr — Session Handoff

_Last updated: 2026-06-19_

---

## 1. Goal

Build **Stockr** — an internal stock ordering web app for the Shakedown restaurant group. Staff at multiple sites submit prep and stock orders via a mobile-friendly web form. Orders are logged to Google Sheets and sent to site-specific Telegram groups (Prep and Stock). A manager dashboard provides spend analytics, order lookup, and credit management.

The app is a single-page HTML app served from **GitHub Pages** (`go2umar/Shakedown-Orders`). The backend is a **Google Apps Script** web app (`stockr.js` → `Code.gs`). Data lives in Google Sheets.

---

## 2. Current State

All features are fully implemented in both files. Frontend auto-deploys on push. **Backend requires a manual redeploy** (paste `stockr.js` into `Code.gs`, deploy with a new version).

### Live features

- **Order submission** — site/type selection, product search, qty steppers, delivery date, custom items, Telegram send
- **Recall** — within 30 min, delete+resend Telegram, auto-credit removed items, basket cleared on success
- **Add-to-order** — manager adds items post-submission; TG message is deleted and resent with full updated order
- **Batch credit system** — manager stages multiple credits inline, then saves all in one call; Credits sheet updated, Orders Summary deducted, Telegram silently annotated
- **Cancel staged credit** — ✕ button removes a staged credit before saving
- **Change Qty** — immediate bidirectional qty correction; updates Order Log, site sheet, and Orders Summary; rebuilds TG (delete+resend so group is notified)
- **Reverse Credit** — deletes all credit entries for a fully-credited item, restores Orders Summary, rebuilds TG silently
- **Change Delivery Date** — updates Order Log, site sheet, and Orders Summary; rebuilds TG (delete+resend)
- **Manager dashboard** — stats, charts, spend by site/supplier, order lookup with site+date filter, top items, credits log, unpriced items, failed Telegram log
- **Item Cost Breakdown** — dashboard section showing all items grouped by category with VAT treatment, VAT amount per item, subtotals per category, and grand totals (ex-VAT / VAT / inc-VAT). Exports to CSV, Excel (.xls), and PDF
- **Item Override** — manager can manually fix category and VAT treatment per item via inline edit button (✎) in the cost breakdown table; saved to "Item Overrides" sheet; applied on every dashboard load
- **PIN** — cross-device sync via Apps Script `PropertiesService`; Change PIN modal in dashboard
- **Branding** — Stockr name, dark theme (`#0D0D0D` bg + `#FF9F1A` orange), `header.png`, splash screen, PWA manifest (`icon.png`)

---

## 3. Files in Flight

| File | Location | How it deploys |
|------|----------|----------------|
| `stockr.js` | `C:\Users\Shake\OneDrive\Desktop\orders\stockr.js` | **Manual** — paste into `Code.gs`, redeploy with new version |
| `index.html` | `C:\Users\Shake\OneDrive\Desktop\orders\index.html` | **Automatic** — GitHub Pages on every push |
| `manifest.json` | `C:\Users\Shake\OneDrive\Desktop\orders\manifest.json` | Automatic |
| `header.png` / `icon.png` | `C:\Users\Shake\OneDrive\Desktop\orders\` | Automatic |

---

## 4. Backend — Key Routes & Handlers

### `doGet` actions
| `action=` | Handler |
|-----------|---------|
| `products` (default) | `handleProductsGet` |
| `dashboard` | `handleDashboardGet` |
| `get_orders` | `handleGetOrders` |
| `recent_order` | `handleRecentOrderGet` |
| `get_pin` | `handleGetPin` |
| `summary` | `handleSummaryGet` |

### `doPost` actions (payload.action)
| `action` | Handler |
|----------|---------|
| `submit` | `handleOrderPost` |
| `recall` | `handleRecallPost` |
| `add_to_order` | `handleAddToOrder` |
| `batch_credit` | `handleBatchCredit` |
| `change_qty` | `handleChangeQty` |
| `reverse_credit` | `handleReverseCredit` |
| `change_delivery_date` | `handleChangeDeliveryDate` |
| `save_item_override` | `handleSaveItemOverride` |
| `credit` | `handleCreditPost` (legacy, single-item) |
| `set_pin` | `handleSetPin` |

### Utility functions (run manually from Apps Script editor)
| Function | Purpose |
|----------|---------|
| `syncOrdersSummary()` | Recalculates every order's total from the Order Log and subtracts batch credits — fixes all historical Orders Summary data in one run. Safe to run multiple times. |
| `setupTrigger()` | Run once after initial deploy to set up price edit trigger and monthly archive trigger |

---

## 5. Key Technical Decisions

### Consuming credit map pattern
Re-added items (same name + price) were getting incorrectly flagged as credited. The fix: consume the credit map sequentially as items appear in the Order Log. Applied in all four Telegram-rebuilding handlers and in `handleGetOrders` (HTML display).

```javascript
const remainingCredits = Object.assign({}, buildCreditMap(ss, orderId));
// Per item:
const key   = name + '|' + price.toFixed(4);
const avail = remainingCredits[key] || 0;
const used  = Math.min(avail, qty);
remainingCredits[key] = Math.max(0, avail - qty); // consume
```

### `fmtTimestamp(raw)` helper
Google Sheets cells containing timestamps are Date objects, not strings. Calling `.toString()` on them produces a long JS date string. `fmtTimestamp` handles both cases and always outputs `dd/MM/yy HH:mm`.

### `buildCreditMap(ss, orderId)` helper
Returns `{ "name|price.toFixed(4)": totalCreditedQty }` for an order. Used anywhere the Telegram message is rebuilt so credited items can be excluded or annotated.

### Telegram strategy
- **Credits (batch_credit, reverse_credit)** → edit-first (silent update, no new group notification). Falls back to delete+resend, then fresh send if edits fail.
- **Order changes (change_qty, change_delivery_date, add_to_order)** → delete+resend (fires a new notification so the group sees the correction).

### Order lookup table — 6 columns
Supplier column was removed to prevent overflow on mobile. All `colSpan` values reference `const colCount = 6;`.

### Credits keyed by `name|price`
Prevents bleed between two rows for the same item at different price points (e.g. item repriced mid-order).

### Manager PIN
Stored in `PropertiesService` (cross-device). `get_pin` and `set_pin` endpoints. Default is `1234`.

### Active-only filtering
All four Price List lookups skip rows where col F (`active`) is not `yes`. Inactive items are rejected at submission.

### All POST requests use form-encoded body
`Content-Type: application/x-www-form-urlencoded` with `body: 'data=' + encodeURIComponent(JSON.stringify(payload))`. Using `application/json` triggers a CORS preflight OPTIONS request that Google Apps Script cannot handle.

### Item Cost Breakdown — data sources
- Built in `handleDashboardGet` using **order IDs from Orders Summary** as the filter (not date-filtering the Order Log directly). This guarantees itemBreakdown totals match Orders Summary exactly.
- **VAT treatment**: Price List col J (index 9). "No VAT" = no uplift. Contains "20%" = multiply spend × 1.2 for total cost.
- **Category**: Price List col H (index 7). All unique categories sent in `categories` array in the dashboard response — frontend uses this dynamically so new categories appear automatically.
- **Manual overrides**: "Item Overrides" sheet (col A: Item Name, col B: Category, col C: VAT Treatment). Read after Price List lookup; takes precedence. Written by `handleSaveItemOverride`.
- **Batch credits subtracted**: Credits sheet entries with reason ≠ `'Order Recalled'` are subtracted from itemSpend/itemTotals for orders in the filtered set. Recall credits are skipped because they already zeroed the Order Log qty.

### `handleSetItemPrice` — must update Orders Summary
When a manager sets a price for an unpriced item, the handler updates Order Log and site sheet AND recalculates Orders Summary from the updated Order Log. Earlier versions skipped the Orders Summary update, causing a permanent gap. Fixed. Historical data corrected by running `syncOrdersSummary()`.

### `syncOrdersSummary()` — how it works
1. Sums Order Log `Total (£)` per order ID (skipping qty ≤ 0 rows)
2. Subtracts batch credits (Credits sheet entries where reason ≠ `'Order Recalled'`)
3. Writes the corrected value to Orders Summary col G for any row where the difference is ≥ £0.01
4. Logs every change to Apps Script Logs (View → Logs)

### `parseLogDate` helper (in `handleDashboardGet`)
Google Sheets reads col M (dd/MM/yyyy date string) back as a Date object. `parseLogDate` handles both cases for the supplier breakdown loop.

---

## 6. Google Sheets Structure

| Sheet | Purpose |
|-------|---------|
| `Order Log` | One row per line item; source of truth |
| `Orders Summary` | One row per order (for dashboard + Looker Studio) |
| `Credits` | One row per credit event |
| `TG Messages` | Maps orderId → prep/stock Telegram message IDs + sent time |
| `Price List` | Product catalogue with site columns, active flag, order type, category (col H), VAT treatment (col J) |
| `Price History` | Auto-logged by `onPriceEdit` trigger whenever Price List col E changes |
| `Item Overrides` | Manual category/VAT overrides per item name (created on first save) |
| `[Site name]` | Per-site copy of Order Log rows (e.g. "SD Withington") |

---

## 7. Price List Column Map

| Col | Index | Field |
|-----|-------|-------|
| A | 0 | Item Name |
| B | 1 | Pkg Type |
| C | 2 | Unit |
| D | 3 | Supplier |
| E | 4 | Price |
| F | 5 | Active (yes/no) |
| G | 6 | — |
| H | 7 | Category |
| I | 8 | Order Type (prep/stock/both) |
| J | 9 | VAT Treatment |
| K+ | 10+ | Site columns (SD Withington = 10, etc.) |

---

## 8. Failed Approaches (Cumulative — Do Not Retry)

| Attempt | Why it failed |
|---------|---------------|
| SVG-drawn logo in header | Too small/complex at 36px; user wanted actual brand images |
| Inline base64 manifest | `short_name: "Orders"` didn't match brand; replaced with external `manifest.json` |
| Sticky dashboard header with image | Image overlapped nav text; separated into scrollable hero + separate nav |
| `border-radius` / `margin` on header image | User wants full-width, perfect rectangle |
| Hamburger button `color: transparent` | User saw blue button; changed to white on dark glass background |
| `localStorage` for PIN | Device-specific; replaced with `PropertiesService` |
| Box-drawing chars in Telegram messages | Caused "4-" on some WhatsApp clients; use ASCII `-----` dashes |
| "Increase Qty" as a separate button | User removed it immediately; replaced with bidirectional "Change Qty" |
| `editOrSend` delete-first in `batch_credit` | If `oldMsgId` was null it fell through to sending a new message (double message); changed to edit-first |
| Static credit map in `handleGetOrders` | Re-added items showed as credited in HTML; fixed with consuming map |
| `origTime = (row[0] || '').toString()` | Produced long JS Date string in Telegram; fixed with `fmtTimestamp()` |
| `colSpan = 7` with Supplier column | Action column overflowed screen; removed Supplier column, now 6 columns |
| `Content-Type: application/json` for POST | Triggers CORS preflight that GAS can't handle; always use form-encoded |
| Date-filtering Order Log by col M string | Sheets returns col M as a Date object, `parseDDMMYYYY` returns null, filter never fires — all historical rows included. Fixed by filtering Order Log via order IDs from Orders Summary instead |
| `handleSetItemPrice` not updating Orders Summary | Priced items updated Order Log but left Orders Summary stale, causing a permanent spend discrepancy. Fixed + historical data corrected via `syncOrdersSummary()` |
| Orange category header rows in cost breakdown | User preferred grey; changed to slate grey `#4B5563` |
| Hardcoded category list in dropdown | New categories silently missing; now reads `data.categories` from backend dynamically |

---

## 9. Next Session — Start Here

Backend and frontend are fully in sync. All historical Orders Summary totals have been corrected by running `syncOrdersSummary()`. Numbers match across all dashboard panels.

### Likely next feature candidates
- WhatsApp message formatting improvements based on real-world use
- Edge cases in recall / credit flows found during daily use
- Performance / UX improvements to the manager dashboard
