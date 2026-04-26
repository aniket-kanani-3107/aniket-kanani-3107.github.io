const test = require('node:test');
const assert = require('node:assert/strict');
const {
  validateVoucher,
  computeLedgerBalances,
  generateTrialBalance,
  generateProfitAndLoss,
  generateBalanceSheet,
} = require('../app/lib/accounting');

test('voucher validation enforces balanced entries', () => {
  const names = new Set(['cash', 'sales']);
  const bad = validateVoucher(names, {
    date: '2026-04-26',
    entries: [
      { ledgerName: 'Cash', debit: 100, credit: 0 },
      { ledgerName: 'Sales', debit: 0, credit: 80 },
    ],
  });
  assert.equal(bad.valid, false);

  const good = validateVoucher(names, {
    date: '2026-04-26',
    entries: [
      { ledgerName: 'Cash', debit: 100, credit: 0 },
      { ledgerName: 'Sales', debit: 0, credit: 100 },
    ],
  });
  assert.equal(good.valid, true);
});

test('core reports generate expected summary', () => {
  const ledgers = [
    { name: 'Cash', type: 'Asset', group: 'Current Assets', openingBalance: 0, openingType: 'debit', isParty: false },
    { name: 'Sales', type: 'Income', group: 'Revenue', openingBalance: 0, openingType: 'credit', isParty: false },
    { name: 'Purchase', type: 'Expense', group: 'Direct', openingBalance: 0, openingType: 'debit', isParty: false },
    { name: 'Capital', type: 'Equity', group: 'Capital', openingBalance: 0, openingType: 'credit', isParty: false },
  ];

  const vouchers = [
    {
      date: '2026-04-01',
      entries: [
        { ledgerName: 'Cash', debit: 1000, credit: 0 },
        { ledgerName: 'Capital', debit: 0, credit: 1000 },
      ],
    },
    {
      date: '2026-04-02',
      entries: [
        { ledgerName: 'Cash', debit: 500, credit: 0 },
        { ledgerName: 'Sales', debit: 0, credit: 500 },
      ],
    },
    {
      date: '2026-04-03',
      entries: [
        { ledgerName: 'Purchase', debit: 200, credit: 0 },
        { ledgerName: 'Cash', debit: 0, credit: 200 },
      ],
    },
  ];

  const balances = computeLedgerBalances(ledgers, vouchers);
  const tb = generateTrialBalance(balances);
  const pnl = generateProfitAndLoss(balances);
  const bs = generateBalanceSheet(balances, pnl);

  assert.equal(Number(tb.totals.debit.toFixed(2)), Number(tb.totals.credit.toFixed(2)));
  assert.equal(pnl.netProfit, 300);
  assert.equal(Number(bs.totalAssets.toFixed(2)), Number(bs.totalLiabilities.toFixed(2)));
});
