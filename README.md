# Petty Cash Manager (Offline-First PWA)

A complete, installable Progressive Web App for managing a local petty cash
fund — ledger, reconciliation, controls, analytics, and Excel/backup
import-export — with **zero backend**. Everything is stored on-device in
IndexedDB.

## Quick Start

No build step, no npm install, no server required.

**Option A — Double-click**
Open `index.html` directly in Chrome, Edge, or Firefox. The app loads
Tailwind, Dexie, SheetJS, and Chart.js from CDN and runs immediately.
(Full offline support and installability require Option B, since Service
Workers need `http://` or `https://`, not `file://`.)

**Option B — VS Code Live Server (recommended)**
1. Open the `petty-cash-pwa` folder in VS Code.
2. Install the "Live Server" extension if you don't have it.
3. Right-click `index.html` → "Open with Live Server".
4. The app opens at `http://127.0.0.1:5500` (or similar). The Service
   Worker registers, the app becomes installable (look for the install
   icon in the address bar), and it will keep working offline afterward.

**Option C — Any static file server**
```bash
cd petty-cash-pwa
python3 -m http.server 8080
# then open http://localhost:8080
```

Once loaded over `http`/`https` at least once with a network connection,
the app is fully cached and continues to work with no internet connection.

## Directory Structure

```
petty-cash-pwa/
├── index.html          # App shell — CDN script tags + mount point
├── manifest.json        # PWA manifest (installable app metadata)
├── sw.js                 # Service worker — offline caching + update detection
├── js/
│   ├── db.js             # Dexie/IndexedDB schema and data-access functions
│   ├── app.js             # Main SPA controller — routing, views, forms
│   ├── excel.js            # SheetJS import/export + printable vouchers
│   ├── charts.js            # Chart.js dashboard visualizations
│   └── crypto-helper.js      # Optional AES-GCM encryption for receipt images
└── icons/                       # App icons (SVG, installable-PWA compliant)
```

## Core Features (as specified)

- **Ledger**: Inflows, Outflows, and Advance/IOU transactions with full field
  set (voucher ID, payment method, category/subcategory, GL code, receipt
  image, status, approver/handler).
- **Auto-generated Voucher IDs**: `PC-YYYYMM-XXXX`, sequential per month.
- **Float & Reconciliation Engine**: live balances per channel (cash /
  mobile money / bank), interactive denomination counter with automatic
  variance highlighting against the system cash balance.
- **Controls**: configurable soft/hard spend caps, low-float reorder-point
  alert, same-payee frequency/split-transaction flag, custody handover log.
- **Import/Export**: styled `.xlsx` export (auto-filter, frozen header,
  totals row), filtered exports (date/category/payee/method), full
  JSON database backup & restore, printable HTML→PDF vouchers, and an
  Excel/CSV import wizard with column mapping and voucher-ID de-duplication.
- **Analytics**: KPI cards (balance, monthly spend, pending receipts, net
  variance) plus category breakdown, daily spend velocity, and payment
  channel charts.
- **UI**: Material Design 3–styled Tailwind interface, responsive sidebar
  (desktop) / bottom nav (mobile), FAB for instant transaction entry, dark
  mode, toast notifications.

## Autonomously Added Features

The brief granted explicit engineering autonomy to add anything that would
make the app more robust or complete. The following were added beyond the
literal spec, and why:

1. **Full audit trail** (`auditTrail` table in `db.js`) — every create,
   update, delete, import, restore, settings change, and custody handover is
   timestamped and logged. Petty cash is a common fraud/error surface;
   an audit trail is standard financial-controls practice and costs almost
   nothing to add.
2. **Receipt image compression** — uploaded receipt photos are downscaled
   and re-encoded via an in-browser canvas before being stored (max 1024px,
   JPEG quality 0.72). Raw phone photos can be several MB each; without
   this, IndexedDB storage would balloon quickly on a device logging dozens
   of receipts a day.
3. **Optional encryption-at-rest for receipts** (`crypto-helper.js`) —
   AES-GCM with a key derived (PBKDF2, 100k iterations) from an app PIN.
   The key lives only in memory for the session and is never persisted.
   This is opt-in because it requires the PIN to always be entered
   correctly to recover images — appropriate for sensitive expense
   documentation, but not forced on users who just want quick offline
   logging.
4. **App PIN lock screen** — a lightweight local lock so a phone left
   unattended doesn't expose the ledger. Pairs with the encryption feature.
5. **Dark mode toggle** — persisted per-install, respects the Material You
   guidance already requested for the light theme.
6. **Keyboard shortcuts** — `N` opens a new transaction, `1`–`5` jump
   between views, `Esc` closes modals. Meaningful time-saver for someone
   logging many vouchers a day at a desk.
7. **PWA update toast** — when a new Service Worker version is detected,
   users get a non-blocking "reload to update" toast instead of silently
   running stale cached code indefinitely.
8. **Live, in-form spend-cap warning** — the amount field warns (or blocks,
   in hard-cap mode) *while typing*, before the voucher is even submitted,
   rather than only after the fact.
9. **Auto-column-guessing on Excel import** — when mapping an uploaded
   sheet's columns, the importer pre-selects likely matches by header name
   (e.g. a column literally called "Amount" auto-maps to the Amount field),
   so the user usually only needs to confirm rather than map from scratch.

## Data & Privacy

All data — transactions, settings, custody logs, audit trail, and receipt
images — is stored **only** in this browser's IndexedDB on this device.
Nothing is transmitted anywhere except to the public CDNs used to load the
Tailwind/Dexie/SheetJS/Chart.js libraries themselves. Use the full-backup
export regularly, since clearing browser data will remove the ledger.
