const express = require('express');
const path = require('path');
const fsp = require('fs/promises');
const chokidar = require('chokidar');
const { randomUUID } = require('crypto');
const {
  ensureAppFiles,
  loadRegistry,
  createCompany,
  getCompany,
  updateCompany,
  loadCompanyData,
  saveCompanyData,
  validateLedgerPayload,
  loadUsers,
  loadTaxConfig,
  saveTaxConfig,
  appendAudit,
  loadAudit,
  createBackup,
  restoreBackup,
  listBackups,
} = require('./lib/storage');
const {
  validateVoucher,
  computeLedgerBalances,
  generateTrialBalance,
  generateProfitAndLoss,
  generateBalanceSheet,
  generateOutstandingAging,
  generateDashboardSnapshot,
} = require('./lib/accounting');
const { parseLinkedWorkbook, writeExcelReport } = require('./lib/excel');

const app = express();
const PORT = process.env.PORT || 3000;
const watchers = new Map();

app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'public')));

async function authMiddleware(req, res, next) {
  if (!req.path.startsWith('/api')) return next();
  if (req.path === '/api/health') return next();

  const username = String(req.header('x-user') || '').trim();
  const password = String(req.header('x-password') || '').trim();
  const users = await loadUsers();
  const user = users.find((u) => u.username === username && u.password === password);

  if (!user) return res.status(401).json({ error: 'Unauthorized. Provide x-user and x-password headers.' });

  req.user = user;
  return next();
}

function requireRole(...roles) {
  return (req, res, next) => {
    if (!roles.includes(req.user?.role)) {
      return res.status(403).json({ error: 'Forbidden for this role.' });
    }
    return next();
  };
}

function ledgerNames(data) {
  return new Set((data.ledgers || []).map((x) => String(x.name || '').toLowerCase()));
}

async function syncLinkedFile(companyId) {
  const company = await getCompany(companyId);
  if (!company || !company.sourceFilePath) return;

  const data = await loadCompanyData(companyId);
  try {
    const linked = await parseLinkedWorkbook(company.sourceFilePath);
    if (linked.ledgers.length) {
      const normalized = [];
      for (const ledger of linked.ledgers) {
        normalized.push(validateLedgerPayload(ledger, normalized));
      }
      data.ledgers = normalized;
    }

    if (linked.vouchers.length) {
      const names = new Set(data.ledgers.map((l) => l.name.toLowerCase()));
      data.vouchers = linked.vouchers
        .map((voucher) => ({
          ...voucher,
          id: randomUUID(),
          createdAt: new Date().toISOString(),
          createdBy: 'system-sync',
        }))
        .filter((voucher) => validateVoucher(names, voucher).valid);
    }

    data.linkedSource = {
      ...(data.linkedSource || {}),
      sourceFilePath: company.sourceFilePath,
      lastSyncedAt: new Date().toISOString(),
      lastError: null,
    };
    await saveCompanyData(companyId, data);
    await appendAudit(companyId, 'LINKED_SYNC', 'system-sync', { sourceFilePath: company.sourceFilePath });
  } catch (error) {
    data.linkedSource = {
      ...(data.linkedSource || {}),
      sourceFilePath: company.sourceFilePath,
      lastError: error.message,
    };
    await saveCompanyData(companyId, data);
  }
}

function registerCompanyWatcher(company) {
  if (!company.sourceFilePath) return;
  if (watchers.has(company.id)) {
    watchers.get(company.id).close();
    watchers.delete(company.id);
  }

  const watcher = chokidar.watch(company.sourceFilePath, { ignoreInitial: true });
  watcher.on('change', async () => {
    await syncLinkedFile(company.id);
  });
  watcher.on('error', () => {
    // no-op
  });
  watchers.set(company.id, watcher);
}

async function hydrateWatchers() {
  const companies = await loadRegistry();
  for (const company of companies) {
    registerCompanyWatcher(company);
    await syncLinkedFile(company.id);
  }
}

async function companyContext(req, res, next) {
  try {
    const companyId = req.params.companyId;
    const company = await getCompany(companyId);
    if (!company) return res.status(404).json({ error: 'Company not found.' });
    const data = await loadCompanyData(companyId);
    req.company = company;
    req.companyData = data;
    return next();
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}

function deriveReports(data, from, to) {
  const balances = computeLedgerBalances(data.ledgers || [], data.vouchers || [], from, to);
  const trialBalance = generateTrialBalance(balances);
  const pnl = generateProfitAndLoss(balances);
  const balanceSheet = generateBalanceSheet(balances, pnl);
  const aging = generateOutstandingAging(balances);
  const dashboard = generateDashboardSnapshot(data.vouchers || [], trialBalance, pnl);
  return { balances, trialBalance, pnl, balanceSheet, aging, dashboard };
}

app.use(authMiddleware);

app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok', at: new Date().toISOString() });
});

app.get('/api/companies', async (_req, res) => {
  res.json(await loadRegistry());
});

app.post('/api/companies', requireRole('admin'), async (req, res) => {
  try {
    const company = await createCompany(req.body || {});
    registerCompanyWatcher(company);
    await appendAudit(company.id, 'COMPANY_CREATE', req.user.username, company);
    if (company.sourceFilePath) await syncLinkedFile(company.id);
    res.status(201).json(company);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.patch('/api/companies/:companyId', requireRole('admin'), companyContext, async (req, res) => {
  try {
    const patch = {};
    if (Object.prototype.hasOwnProperty.call(req.body, 'sourceFilePath')) patch.sourceFilePath = req.body.sourceFilePath;
    if (Object.prototype.hasOwnProperty.call(req.body, 'name')) patch.name = String(req.body.name || '').trim();

    const updated = await updateCompany(req.company.id, patch);
    registerCompanyWatcher(updated);
    await appendAudit(updated.id, 'COMPANY_UPDATE', req.user.username, patch);
    if (updated.sourceFilePath) await syncLinkedFile(updated.id);
    res.json(updated);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.post('/api/companies/:companyId/sync', requireRole('admin'), companyContext, async (req, res) => {
  await syncLinkedFile(req.company.id);
  res.json({ status: 'synced' });
});

app.get('/api/companies/:companyId/data', companyContext, (req, res) => {
  res.json(req.companyData);
});

app.get('/api/companies/:companyId/ledgers', companyContext, (req, res) => {
  res.json(req.companyData.ledgers || []);
});

app.post('/api/companies/:companyId/ledgers', requireRole('admin'), companyContext, async (req, res) => {
  try {
    const ledgers = req.companyData.ledgers || [];
    const ledger = validateLedgerPayload(req.body, ledgers);
    ledgers.push(ledger);
    req.companyData.ledgers = ledgers;
    await saveCompanyData(req.company.id, req.companyData);
    await appendAudit(req.company.id, 'LEDGER_CREATE', req.user.username, ledger);
    res.status(201).json(ledger);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.put('/api/companies/:companyId/ledgers/:ledgerId', requireRole('admin'), companyContext, async (req, res) => {
  try {
    const ledgers = req.companyData.ledgers || [];
    const index = ledgers.findIndex((x) => x.id === req.params.ledgerId);
    if (index === -1) return res.status(404).json({ error: 'Ledger not found.' });

    const nextLedger = validateLedgerPayload({ ...ledgers[index], ...req.body, id: req.params.ledgerId }, ledgers, req.params.ledgerId);
    ledgers[index] = nextLedger;
    req.companyData.ledgers = ledgers;
    await saveCompanyData(req.company.id, req.companyData);
    await appendAudit(req.company.id, 'LEDGER_UPDATE', req.user.username, nextLedger);
    return res.json(nextLedger);
  } catch (error) {
    return res.status(400).json({ error: error.message });
  }
});

app.get('/api/companies/:companyId/vouchers', companyContext, (req, res) => {
  const { from, to, type } = req.query;
  const rows = (req.companyData.vouchers || []).filter((voucher) => {
    if (type && voucher.type !== type) return false;
    if (from && voucher.date < from) return false;
    if (to && voucher.date > to) return false;
    return true;
  });
  res.json(rows);
});

app.post('/api/companies/:companyId/vouchers', requireRole('admin'), companyContext, async (req, res) => {
  try {
    const taxConfig = await loadTaxConfig();
    const payload = { ...req.body };
    payload.entries = Array.isArray(payload.entries) ? payload.entries : [];

    if (payload.tax?.enabled && Number(payload.tax.rate || 0) > 0) {
      const taxableAmount = payload.entries.reduce((sum, e) => sum + Number(e.credit || 0), 0);
      const taxAmount = (taxableAmount * Number(payload.tax.rate)) / 100;
      if (taxAmount > 0) {
        const taxLedgerName = payload.tax.ledgerName || taxConfig.defaultTaxLedger || 'GST Output';
        payload.entries.push({ ledgerName: taxLedgerName, debit: 0, credit: taxAmount });
        const debitLine = payload.entries.find((x) => Number(x.debit || 0) > 0);
        if (debitLine) debitLine.debit = Number(debitLine.debit || 0) + taxAmount;
      }
    }

    const names = ledgerNames(req.companyData);
    const validation = validateVoucher(names, payload);
    if (!validation.valid) return res.status(400).json({ error: validation.errors.join(' | ') });

    const voucher = {
      id: randomUUID(),
      date: payload.date,
      type: payload.type || 'Journal',
      narration: payload.narration || '',
      entries: payload.entries.map((entry) => ({
        ledgerName: entry.ledgerName,
        debit: Number(entry.debit || 0),
        credit: Number(entry.credit || 0),
      })),
      createdAt: new Date().toISOString(),
      createdBy: req.user.username,
    };

    req.companyData.vouchers = req.companyData.vouchers || [];
    req.companyData.vouchers.push(voucher);
    await saveCompanyData(req.company.id, req.companyData);
    await appendAudit(req.company.id, 'VOUCHER_CREATE', req.user.username, voucher);
    res.status(201).json(voucher);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.get('/api/companies/:companyId/reports/trial-balance', companyContext, (req, res) => {
  const { from, to } = req.query;
  const { trialBalance } = deriveReports(req.companyData, from, to);
  res.json(trialBalance);
});

app.get('/api/companies/:companyId/reports/pnl', companyContext, (req, res) => {
  const { from, to } = req.query;
  const { pnl } = deriveReports(req.companyData, from, to);
  res.json(pnl);
});

app.get('/api/companies/:companyId/reports/balance-sheet', companyContext, (req, res) => {
  const { from, to } = req.query;
  const { balanceSheet } = deriveReports(req.companyData, from, to);
  res.json(balanceSheet);
});

app.get('/api/companies/:companyId/reports/dashboard', companyContext, (req, res) => {
  const { from, to } = req.query;
  const { dashboard } = deriveReports(req.companyData, from, to);
  res.json(dashboard);
});

app.get('/api/companies/:companyId/reports/outstanding-aging', companyContext, (req, res) => {
  const { from, to } = req.query;
  const { aging } = deriveReports(req.companyData, from, to);
  res.json(aging);
});

app.get('/api/companies/:companyId/export/:reportName.xlsx', companyContext, async (req, res) => {
  try {
    const { from, to } = req.query;
    const reportName = req.params.reportName;
    const reports = deriveReports(req.companyData, from, to);

    let headers = [];
    let rows = [];

    if (reportName === 'trial-balance') {
      headers = ['Ledger', 'Type', 'Group', 'Debit', 'Credit'];
      rows = reports.trialBalance.rows.map((row) => [row.ledger, row.type, row.group, row.debit, row.credit]);
    } else if (reportName === 'pnl-income') {
      headers = ['Ledger', 'Amount'];
      rows = reports.pnl.incomeRows.map((row) => [row.ledger, row.amount]);
    } else if (reportName === 'pnl-expense') {
      headers = ['Ledger', 'Amount'];
      rows = reports.pnl.expenseRows.map((row) => [row.ledger, row.amount]);
    } else if (reportName === 'vouchers') {
      headers = ['Date', 'Type', 'Narration', 'Ledger', 'Debit', 'Credit'];
      rows = (req.companyData.vouchers || []).flatMap((voucher) =>
        voucher.entries.map((entry) => [voucher.date, voucher.type, voucher.narration, entry.ledgerName, entry.debit, entry.credit])
      );
    } else {
      return res.status(400).json({ error: 'Unsupported report export name.' });
    }

    const exportDir = path.join('/tmp', 'exports');
    await fsp.mkdir(exportDir, { recursive: true });
    const filePath = path.join(exportDir, `${req.company.id}-${reportName}-${Date.now()}.xlsx`);

    await writeExcelReport(filePath, reportName, headers, rows);
    res.download(filePath, `${reportName}.xlsx`, async () => {
      try {
        await fsp.unlink(filePath);
      } catch {
        // no-op
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/companies/:companyId/audit', companyContext, async (req, res) => {
  res.json(await loadAudit(req.company.id));
});

app.get('/api/tax-config', async (_req, res) => {
  res.json(await loadTaxConfig());
});

app.put('/api/tax-config', requireRole('admin'), async (req, res) => {
  const body = req.body || {};
  const config = {
    defaultTaxLedger: body.defaultTaxLedger || 'GST Output',
    rates: Array.isArray(body.rates)
      ? body.rates.map((x) => ({ name: String(x.name || '').trim(), rate: Number(x.rate || 0) })).filter((x) => x.name && x.rate >= 0)
      : [{ name: 'GST18', rate: 18 }],
    mode: body.mode === 'inclusive' ? 'inclusive' : 'exclusive',
  };
  await saveTaxConfig(config);
  res.json(config);
});

app.post('/api/companies/:companyId/backup', requireRole('admin'), companyContext, async (req, res) => {
  try {
    const fileName = await createBackup(req.company.id);
    await appendAudit(req.company.id, 'BACKUP_CREATE', req.user.username, { fileName });
    res.status(201).json({ fileName });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.get('/api/companies/:companyId/backups', requireRole('admin'), companyContext, async (req, res) => {
  res.json(await listBackups(req.company.id));
});

app.post('/api/companies/:companyId/restore', requireRole('admin'), companyContext, async (req, res) => {
  try {
    const backupFileName = String(req.body?.backupFileName || '');
    await restoreBackup(req.company.id, backupFileName);
    await appendAudit(req.company.id, 'BACKUP_RESTORE', req.user.username, { backupFileName });
    res.json({ status: 'restored', backupFileName });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.use((_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

async function start() {
  await ensureAppFiles();
  await hydrateWatchers();
  app.listen(PORT, () => {
    // eslint-disable-next-line no-console
    console.log(`Accounting app running on http://localhost:${PORT}`);
  });
}

start();
