# Tally-Style Offline Accounting Base

A Node.js + Express accounting foundation that follows a base-to-enhanced roadmap:

- One-time company setup with optional linked source file path
- Ledger and voucher engine with double-entry validation
- Trial Balance, P&L, Balance Sheet, Dashboard, Outstanding Aging
- Excel export for reports/vouchers
- File watcher sync for linked JSON/XLSX sources
- Tax config, audit trail, backup/restore, multi-company support

## Project structure

```text
/app
  /backups
  /companies
  /linked-sources
  /config
    companies.json
    users.json
    tax.json
  /data
    /audit
  /lib
    accounting.js
    excel.js
    storage.js
  /public
    index.html
    styles.css
    app.js
  server.js
```

## Run locally

```bash
npm install
npm run start
```

Open: `http://localhost:3000`

Default credentials:

- user: `admin`
- password: `admin123`

## Scripts

- `npm run start` – start app server
- `npm run test` – run node tests

## Linked data sync

At company creation/update, set `sourceFilePath` to a file inside `app/linked-sources` (`.json` or `.xlsx`).

For `.xlsx`, optional sheets:

- `Ledgers` columns: `name,type,group,openingBalance,openingType,isParty`
- `Vouchers` columns: `date,type,narration,debitLedger,creditLedger,amount`

Sync behavior:

- auto sync on startup
- auto sync on linked file changes
- manual sync from UI/API

## API highlights

- `GET /api/companies`
- `POST /api/companies`
- `POST /api/companies/:companyId/ledgers`
- `POST /api/companies/:companyId/vouchers`
- `GET /api/companies/:companyId/reports/trial-balance`
- `GET /api/companies/:companyId/reports/pnl`
- `GET /api/companies/:companyId/reports/balance-sheet`
- `GET /api/companies/:companyId/export/trial-balance.xlsx`
- `POST /api/companies/:companyId/backup`
- `POST /api/companies/:companyId/restore`

All `/api/*` endpoints (except `/api/health`) require headers:

- `x-user`
- `x-password`
