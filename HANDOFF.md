# Stockr — Session Handoff

---

## 1. Goal

Build **Stockr** — an internal stock ordering web app for the Shakedown restaurant group. Staff at multiple sites submit prep and stock orders via a mobile-friendly web form. Orders are logged to Google Sheets and sent to site-specific Telegram groups (Prep and Stock). A manager dashboard provides spend analytics, order lookup, and credit management.

The app is a single-page HTML app served from **GitHub Pages** (`go2umar/Shakedown-Orders`). The backend is a **Google Apps Script** web app (`stockr.js` → `Code.gs`). Data lives in Google Sheets.

---

## 2. Current State

The app is fully functional and branded as **Stockr**. Core ordering, recall, credit, and manager dashboard flows are working. The following are live:

- **Order submission** — site/type selection, product search, qty steppers, delivery date, custom items, Telegram send
- **Recall** — within 30 min, delete+resend Telegram, auto-credit removed items, add new items, basket cleared on success
- **Add to order** — manager can add items post-submission; Telegram message is updated (delete+resend)
- **Credits** — manager credits from order lookup; removed items can be credited for past orders
- **Manager dashboard** — stats, charts, spend by site/supplier, order lookup, filtering by site + date range
- **PIN** — cross-device sync via Apps Script `PropertiesService`; change PIN UI in dashboard
- **Branding** — Stockr name, dark theme (`#0D0D0D` + orange `#FF9F1A`), custom brand header image, splash screen, PWA manifest

**Pending manual steps (not yet deployed):**
- Several recent `stockr.js` changes need to be pasted into `Code.gs` and redeployed as a new version
- The corrupt header row in the **Orders Summary** sheet needs to be manually deleted (row containing "Order ID" in column A, between ORD-0056 and ORD-0057)

---

## 3. Files in Flight

| File | Location | How it deploys |
|------|----------|----------------|
| `stockr.js` | `C:\Users\Shake\OneDrive\Desktop\orders\stockr.js` | **Manual** — paste into `Code.gs` in Apps Script, redeploy with new version |
| `index.html` | `C:\Users\Shake\OneDrive\Desktop\orders\index.html` | **Automatic** — GitHub Pages serves it on every push |
| `manifest.json` | `C:\Users\Shake\OneDrive\Desktop\orders\manifest.json` | **Automatic** — GitHub Pages |
| `header.png` | `C:\Users\Shake\OneDrive\Desktop\orders\header.png` | **Automatic** — GitHub Pages |
| `icon.png` | `C:\Users\Shake\OneDrive\Desktop\orders\icon.png` | **Automatic** — GitHub Pages (also used as PWA icon) |

---

## 4. Changed This Session

### Backend (`stockr.js`)
- **Recall flow** — zero-qty items now filtered from Telegram messages; new items stamped with recall timestamp not original order time; TG Messages read with `break` after first match
- **Add-to-order** — now edits (delete+resend) the existing Telegram message instead of sending a separate addition; triggers proper notification
- **Credits** — matched by `name + price` to prevent cross-row bleed when same item exists at two price points
- **Active item filtering** — all four Price List lookups now skip inactive rows; inactive items are rejected on submission
- **Auto-credit on recall** — removed/reduced items during recall are automatically credited in Credits sheet
- **Telegram format** — `*bold*` WhatsApp markers; `📌` pin emoji; site name as bold all-caps header; ASCII `-----` separator (box-drawing chars caused "4-" on some WhatsApp clients); category grouping
- **Manager PIN** — stored in `PropertiesService` (cross-device); `get_pin` and `set_pin` endpoints added
- **Dashboard date filter** — `row[10]` is a Date object in Sheets; now handled with `instanceof Date` check instead of `.toString()`; corrupt header rows skipped via `ORD-\d+` orderId validation
- **Empty state** — dashboard shows "No orders found for this period" when filtered result is zero

### Frontend (`index.html`)
- **Stockr branding** — dark header (`header.png`), splash screen, orange accent throughout, section labels white on dark bg, status bar dark
- **Manager dashboard** — fully scrollable (no sticky elements); dark theme; content constrained to 640px on desktop; header image shared with main app
- **PIN overlay** — dark theme, orange button; Change PIN modal added; PIN synced from server on load
- **Item filter** — search bar in order detail view (>8 items); credit form rows use `data-form-row` so filter no longer opens all credit forms
- **Credit form layout** — consistent column layout (title / inputs / buttons) regardless of item name length
- **Add-item search** — fixed in manager view (was empty because `_searchIndex` only populated in recall mode); now builds local pool from `SITE_DATA`
- **Basket clear** — after recall, `resetForm()` called so basket is empty on return
- **PWA** — `manifest.json` with `name: "Stockr"`, `icon.png` as app icon; `apple-touch-icon` updated

---

## 5. Failed Attempts

| Attempt | What happened | Why it failed |
|---------|---------------|---------------|
| SVG-drawn logo icons | Created SVG cart+person in header | Too small/complex at 36px; user wanted actual brand images instead |
| Inline base64 manifest | Previously used base64-encoded manifest with `short_name: "Orders"` | Name didn't match brand; replaced with external `manifest.json` |
| Sticky manager dashboard header with image | Combined header image + nav bar in one sticky element | Image overlapped nav text; separated into scrollable hero + separate nav bar |
| `border-radius` on header | Added rounded corners to make header look contained | User wanted a perfect rectangle; removed |
| `margin: 12px` on header | Added side margins to align with card layout | User wanted full-width header; reverted |
| `color: transparent` on hamburger button | Made hamburger transparent (click target over image icon) | User saw blue button; changed to white with dark glass background |
| `localStorage` for PIN | Stored manager PIN in localStorage | Device-specific; didn't sync to phone. Replaced with `PropertiesService` |

---

## 6. Next Step

**Deploy the pending backend changes** — paste `stockr.js` into `Code.gs` in the Apps Script editor and redeploy with a new version. This activates:
- Cross-device PIN sync (`get_pin` / `set_pin`)
- Dashboard date filter fix (Date object handling)
- Corrupt row filtering (`ORD-\d+` check)
- WhatsApp-friendly separators (`-----`)

Then **delete the corrupt header row** from the Orders Summary sheet (the row with "Order ID" in column A between ORD-0056 and ORD-0057) to clean up the monthly trend chart permanently.
