function normalizeNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function toTime(dateStr) {
  if (!dateStr) return null;
  const time = new Date(dateStr).getTime();
  return Number.isNaN(time) ? null : time;
}

function inDateRange(date, from, to) {
  const d = toTime(date);
  if (d === null) return false;
  const f = toTime(from);
  const t = toTime(to);
  if (f !== null && d < f) return false;
  if (t !== null && d > t) return false;
  return true;
}

function validateVoucher(ledgerNames, voucher) {
  const errors = [];
  if (!voucher || typeof voucher !== 'object') errors.push('Voucher payload is required.');
  if (!voucher?.date || Number.isNaN(new Date(voucher.date).getTime())) errors.push('Valid voucher date is required.');
  if (!Array.isArray(voucher?.entries) || voucher.entries.length < 2) errors.push('At least two voucher entries are required.');

  let debitTotal = 0;
  let creditTotal = 0;
  for (const entry of voucher?.entries || []) {
    if (!entry.ledgerName || !ledgerNames.has(entry.ledgerName.toLowerCase())) {
      errors.push(`Unknown ledger: ${entry.ledgerName || 'N/A'}`);
      continue;
    }
    const debit = normalizeNumber(entry.debit);
    const credit = normalizeNumber(entry.credit);
    if ((debit > 0 && credit > 0) || (debit === 0 && credit === 0)) {
      errors.push(`Entry for ${entry.ledgerName} must have either debit or credit amount.`);
    }
    debitTotal += debit;
    creditTotal += credit;
  }

  if (Math.abs(debitTotal - creditTotal) > 0.0001) {
    errors.push(`Voucher is unbalanced. Debit=${debitTotal} Credit=${creditTotal}`);
  }

  return { valid: errors.length === 0, errors, debitTotal, creditTotal };
}

function computeLedgerBalances(ledgers, vouchers, from, to) {
  const balances = new Map();

  for (const ledger of ledgers) {
    const key = ledger.name;
    const openingBalance = normalizeNumber(ledger.openingBalance);
    const openingType = (ledger.openingType || 'debit').toLowerCase();

    balances.set(key, {
      ledgerName: key,
      type: ledger.type,
      group: ledger.group,
      isParty: Boolean(ledger.isParty),
      debitTotal: openingType === 'debit' ? openingBalance : 0,
      creditTotal: openingType === 'credit' ? openingBalance : 0,
      lastEntryDate: null,
    });
  }

  for (const voucher of vouchers) {
    if (!inDateRange(voucher.date, from, to)) continue;
    for (const entry of voucher.entries || []) {
      const balance = balances.get(entry.ledgerName);
      if (!balance) continue;

      balance.debitTotal += normalizeNumber(entry.debit);
      balance.creditTotal += normalizeNumber(entry.credit);
      const dateTs = toTime(voucher.date);
      if (dateTs !== null && (!balance.lastEntryDate || dateTs > toTime(balance.lastEntryDate))) {
        balance.lastEntryDate = voucher.date;
      }
    }
  }

  for (const balance of balances.values()) {
    const net = balance.debitTotal - balance.creditTotal;
    balance.net = net;
    balance.closingSide = net >= 0 ? 'debit' : 'credit';
    balance.closingAmount = Math.abs(net);
  }

  return [...balances.values()].sort((a, b) => a.ledgerName.localeCompare(b.ledgerName));
}

function generateTrialBalance(balances) {
  const rows = balances.map((item) => ({
    ledger: item.ledgerName,
    type: item.type,
    group: item.group,
    debit: item.closingSide === 'debit' ? item.closingAmount : 0,
    credit: item.closingSide === 'credit' ? item.closingAmount : 0,
  }));

  const totals = rows.reduce(
    (acc, row) => {
      acc.debit += row.debit;
      acc.credit += row.credit;
      return acc;
    },
    { debit: 0, credit: 0 }
  );

  return { rows, totals };
}

function generateProfitAndLoss(balances) {
  const incomeRows = [];
  const expenseRows = [];

  for (const item of balances) {
    if (item.type === 'Income') {
      incomeRows.push({ ledger: item.ledgerName, amount: item.creditTotal - item.debitTotal });
    } else if (item.type === 'Expense') {
      expenseRows.push({ ledger: item.ledgerName, amount: item.debitTotal - item.creditTotal });
    }
  }

  const totalIncome = incomeRows.reduce((sum, row) => sum + normalizeNumber(row.amount), 0);
  const totalExpense = expenseRows.reduce((sum, row) => sum + normalizeNumber(row.amount), 0);
  const netProfit = totalIncome - totalExpense;

  return { incomeRows, expenseRows, totalIncome, totalExpense, netProfit };
}

function generateBalanceSheet(balances, pnl) {
  const assets = [];
  const liabilities = [];

  for (const item of balances) {
    if (item.type === 'Asset') {
      assets.push({ ledger: item.ledgerName, amount: item.debitTotal - item.creditTotal });
    }
    if (item.type === 'Liability' || item.type === 'Equity') {
      liabilities.push({ ledger: item.ledgerName, amount: item.creditTotal - item.debitTotal });
    }
  }

  if (pnl.netProfit >= 0) {
    liabilities.push({ ledger: 'Current Year Profit', amount: pnl.netProfit });
  } else {
    assets.push({ ledger: 'Current Year Loss', amount: Math.abs(pnl.netProfit) });
  }

  const totalAssets = assets.reduce((sum, row) => sum + normalizeNumber(row.amount), 0);
  const totalLiabilities = liabilities.reduce((sum, row) => sum + normalizeNumber(row.amount), 0);

  return { assets, liabilities, totalAssets, totalLiabilities };
}

function generateOutstandingAging(balances, now = new Date()) {
  const todayTs = now.getTime();
  const buckets = [
    { label: '0-30', min: 0, max: 30 },
    { label: '31-60', min: 31, max: 60 },
    { label: '61-90', min: 61, max: 90 },
    { label: '90+', min: 91, max: Number.MAX_SAFE_INTEGER },
  ];

  const rows = [];
  for (const item of balances.filter((x) => x.isParty)) {
    const balanceAmount = item.debitTotal - item.creditTotal;
    if (Math.abs(balanceAmount) < 0.0001) continue;

    const lastTs = toTime(item.lastEntryDate);
    const ageDays = lastTs ? Math.floor((todayTs - lastTs) / (1000 * 60 * 60 * 24)) : 0;
    const bucket = buckets.find((b) => ageDays >= b.min && ageDays <= b.max)?.label || '90+';

    rows.push({
      ledger: item.ledgerName,
      amount: balanceAmount,
      side: balanceAmount >= 0 ? 'Receivable' : 'Payable',
      ageDays,
      bucket,
    });
  }

  return rows;
}

function generateDashboardSnapshot(vouchers, trialBalance, pnl) {
  return {
    totalVouchers: vouchers.length,
    totalLedgers: trialBalance.rows.length,
    totalDebit: trialBalance.totals.debit,
    totalCredit: trialBalance.totals.credit,
    totalIncome: pnl.totalIncome,
    totalExpense: pnl.totalExpense,
    netProfit: pnl.netProfit,
  };
}

module.exports = {
  validateVoucher,
  computeLedgerBalances,
  generateTrialBalance,
  generateProfitAndLoss,
  generateBalanceSheet,
  generateOutstandingAging,
  generateDashboardSnapshot,
};
