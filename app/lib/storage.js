const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const { randomUUID } = require('crypto');

const APP_DIR = path.resolve(__dirname, '..');
const CONFIG_DIR = path.join(APP_DIR, 'config');
const COMPANIES_DIR = path.join(APP_DIR, 'companies');
const AUDIT_DIR = path.join(APP_DIR, 'data', 'audit');
const BACKUP_DIR = path.join(APP_DIR, 'backups');
const REGISTRY_FILE = path.join(CONFIG_DIR, 'companies.json');
const USERS_FILE = path.join(CONFIG_DIR, 'users.json');
const TAX_FILE = path.join(CONFIG_DIR, 'tax.json');

async function ensureDir(dir) {
  await fsp.mkdir(dir, { recursive: true });
}

async function ensureFile(filePath, defaultValue) {
  if (!fs.existsSync(filePath)) {
    await fsp.writeFile(filePath, JSON.stringify(defaultValue, null, 2));
  }
}

async function ensureAppFiles() {
  await Promise.all([ensureDir(CONFIG_DIR), ensureDir(COMPANIES_DIR), ensureDir(AUDIT_DIR), ensureDir(BACKUP_DIR)]);
  await ensureFile(REGISTRY_FILE, []);
  await ensureFile(USERS_FILE, [{ username: 'admin', password: 'admin123', role: 'admin' }]);
  await ensureFile(TAX_FILE, {
    defaultTaxLedger: 'GST Output',
    rates: [{ name: 'GST18', rate: 18 }],
    mode: 'exclusive',
  });
}

async function readJson(filePath, fallback = null) {
  try {
    const content = await fsp.readFile(filePath, 'utf-8');
    return JSON.parse(content);
  } catch {
    return fallback;
  }
}

async function writeJson(filePath, data) {
  await fsp.writeFile(filePath, JSON.stringify(data, null, 2));
}

function normalizeLedgerName(name) {
  return String(name || '').trim();
}

function companyFilePath(companyId) {
  return path.join(COMPANIES_DIR, `${companyId}.json`);
}

function companyBackupFilePath(companyId) {
  return path.join(BACKUP_DIR, `${companyId}-${Date.now()}.json`);
}

async function loadRegistry() {
  return (await readJson(REGISTRY_FILE, [])) || [];
}

async function saveRegistry(registry) {
  await writeJson(REGISTRY_FILE, registry);
}

async function loadUsers() {
  return (await readJson(USERS_FILE, [])) || [];
}

async function loadTaxConfig() {
  return (await readJson(TAX_FILE, {})) || {};
}

async function saveTaxConfig(config) {
  await writeJson(TAX_FILE, config);
}

async function appendAudit(companyId, action, actor, payload) {
  const filePath = path.join(AUDIT_DIR, `${companyId}.json`);
  const data = (await readJson(filePath, [])) || [];
  data.push({
    id: randomUUID(),
    action,
    actor,
    payload,
    at: new Date().toISOString(),
  });
  await writeJson(filePath, data);
}

async function loadAudit(companyId) {
  return (await readJson(path.join(AUDIT_DIR, `${companyId}.json`), [])) || [];
}

function createCompanyTemplate(company) {
  return {
    company,
    meta: {
      templateVersion: '1.0.0',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      namingRule: 'Ledger names are unique and immutable by meaning. Do not rename randomly.',
    },
    ledgers: [
      { id: randomUUID(), name: 'Cash', type: 'Asset', group: 'Current Assets', openingBalance: 0, openingType: 'debit', isParty: false },
      { id: randomUUID(), name: 'Capital', type: 'Equity', group: 'Capital Account', openingBalance: 0, openingType: 'credit', isParty: false },
      { id: randomUUID(), name: 'Sales', type: 'Income', group: 'Revenue', openingBalance: 0, openingType: 'credit', isParty: false },
      { id: randomUUID(), name: 'Purchase', type: 'Expense', group: 'Direct Expenses', openingBalance: 0, openingType: 'debit', isParty: false },
      { id: randomUUID(), name: 'GST Output', type: 'Liability', group: 'Duties & Taxes', openingBalance: 0, openingType: 'credit', isParty: false },
      { id: randomUUID(), name: 'GST Input', type: 'Asset', group: 'Duties & Taxes', openingBalance: 0, openingType: 'debit', isParty: false },
    ],
    vouchers: [],
    linkedSource: {
      sourceFilePath: company.sourceFilePath || null,
      lastSyncedAt: null,
      lastError: null,
    },
  };
}

async function createCompany(payload) {
  const registry = await loadRegistry();
  const company = {
    id: randomUUID(),
    name: String(payload.name || '').trim(),
    code: String(payload.code || '').trim().toUpperCase(),
    financialYear: String(payload.financialYear || '').trim() || '2026-2027',
    baseCurrency: String(payload.baseCurrency || 'INR').trim().toUpperCase(),
    sourceFilePath: payload.sourceFilePath || null,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  if (!company.name) throw new Error('Company name is required.');
  if (registry.some((x) => x.code === company.code && company.code)) {
    throw new Error('Company code must be unique.');
  }

  registry.push(company);
  await saveRegistry(registry);
  await writeJson(companyFilePath(company.id), createCompanyTemplate(company));
  return company;
}

async function getCompany(companyId) {
  const registry = await loadRegistry();
  return registry.find((x) => x.id === companyId) || null;
}

async function updateCompany(companyId, patch) {
  const registry = await loadRegistry();
  const index = registry.findIndex((x) => x.id === companyId);
  if (index === -1) throw new Error('Company not found.');
  registry[index] = { ...registry[index], ...patch, updatedAt: new Date().toISOString() };
  await saveRegistry(registry);
  return registry[index];
}

async function loadCompanyData(companyId) {
  const filePath = companyFilePath(companyId);
  const data = await readJson(filePath, null);
  if (!data) throw new Error('Company data file not found.');
  return data;
}

async function saveCompanyData(companyId, data) {
  data.meta = { ...(data.meta || {}), updatedAt: new Date().toISOString() };
  await writeJson(companyFilePath(companyId), data);
}

function validateLedgerPayload(data, existingLedgers = [], editingLedgerId = null) {
  const allowedTypes = new Set(['Asset', 'Liability', 'Income', 'Expense', 'Equity']);
  const name = normalizeLedgerName(data.name);
  const type = String(data.type || '').trim();
  const group = String(data.group || '').trim() || 'General';
  const openingType = String(data.openingType || 'debit').trim().toLowerCase();
  const openingBalance = Number(data.openingBalance || 0);

  if (!name) throw new Error('Ledger name is required.');
  if (!allowedTypes.has(type)) throw new Error('Ledger type must be one of Asset/Liability/Income/Expense/Equity.');
  if (!['debit', 'credit'].includes(openingType)) throw new Error('Opening type must be debit or credit.');
  if (!Number.isFinite(openingBalance) || openingBalance < 0) throw new Error('Opening balance must be a non-negative number.');

  const duplicate = existingLedgers.some(
    (ledger) => ledger.id !== editingLedgerId && normalizeLedgerName(ledger.name).toLowerCase() === name.toLowerCase()
  );
  if (duplicate) throw new Error('Ledger name must be unique.');

  return {
    id: data.id || randomUUID(),
    name,
    type,
    group,
    openingBalance,
    openingType,
    isParty: Boolean(data.isParty),
  };
}

async function createBackup(companyId) {
  const sourcePath = companyFilePath(companyId);
  if (!fs.existsSync(sourcePath)) throw new Error('Company file not found.');
  const backupPath = companyBackupFilePath(companyId);
  await fsp.copyFile(sourcePath, backupPath);
  return path.basename(backupPath);
}

async function restoreBackup(companyId, backupFileName) {
  const sourcePath = path.join(BACKUP_DIR, backupFileName);
  if (!fs.existsSync(sourcePath)) throw new Error('Backup file not found.');
  await fsp.copyFile(sourcePath, companyFilePath(companyId));
}

async function listBackups(companyId) {
  const files = await fsp.readdir(BACKUP_DIR);
  return files.filter((file) => file.startsWith(`${companyId}-`)).sort().reverse();
}

module.exports = {
  APP_DIR,
  CONFIG_DIR,
  COMPANIES_DIR,
  REGISTRY_FILE,
  ensureAppFiles,
  readJson,
  writeJson,
  loadRegistry,
  saveRegistry,
  loadUsers,
  loadTaxConfig,
  saveTaxConfig,
  createCompany,
  getCompany,
  updateCompany,
  loadCompanyData,
  saveCompanyData,
  validateLedgerPayload,
  appendAudit,
  loadAudit,
  createBackup,
  restoreBackup,
  listBackups,
};
