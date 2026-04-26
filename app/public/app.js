const state = {
  companies: [],
  currentCompanyId: null,
  chart: null,
};

function authHeaders() {
  return {
    'x-user': document.getElementById('auth-user').value,
    'x-password': document.getElementById('auth-password').value,
  };
}

async function api(path, options = {}) {
  const response = await fetch(path, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      ...authHeaders(),
      ...(options.headers || {}),
    },
  });

  const contentType = response.headers.get('content-type') || '';
  const payload = contentType.includes('application/json') ? await response.json() : await response.text();
  if (!response.ok) throw new Error(payload.error || payload || response.statusText);
  return payload;
}

function setStatus(message) {
  document.getElementById('status').textContent = `[${new Date().toLocaleString()}] ${message}`;
}

function table(elId, headers, rows) {
  const el = document.getElementById(elId);
  el.innerHTML = `<thead><tr>${headers.map((h) => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows
    .map((row) => `<tr>${row.map((cell) => `<td>${cell ?? ''}</td>`).join('')}</tr>`)
    .join('')}</tbody>`;
}

async function loadCompanies() {
  state.companies = await api('/api/companies');
  const select = document.getElementById('company-select');
  select.innerHTML = state.companies.map((c) => `<option value="${c.id}">${c.name} (${c.code || 'N/A'})</option>`).join('');

  if (state.companies.length && !state.currentCompanyId) {
    state.currentCompanyId = state.companies[0].id;
  }

  if (state.currentCompanyId) {
    select.value = state.currentCompanyId;
    await refreshCompanyData();
  } else {
    document.getElementById('company-meta').textContent = 'No companies available.';
  }
}

async function refreshCompanyData() {
  if (!state.currentCompanyId) return;
  const data = await api(`/api/companies/${state.currentCompanyId}/data`);
  document.getElementById('company-meta').textContent = JSON.stringify(data.meta, null, 2);

  table(
    'ledger-table',
    ['Name', 'Type', 'Group', 'Opening', 'Party'],
    (data.ledgers || []).map((l) => [l.name, l.type, l.group, `${l.openingType} ${l.openingBalance}`, l.isParty ? 'Yes' : 'No'])
  );

  table(
    'voucher-table',
    ['Date', 'Type', 'Narration', 'Entries'],
    (data.vouchers || []).slice(-20).reverse().map((v) => [
      v.date,
      v.type,
      v.narration || '',
      (v.entries || []).map((e) => `${e.ledgerName}: D${e.debit || 0}/C${e.credit || 0}`).join(' | '),
    ])
  );
}

function reportFilters() {
  const from = document.getElementById('filter-from').value;
  const to = document.getElementById('filter-to').value;
  const params = new URLSearchParams();
  if (from) params.set('from', from);
  if (to) params.set('to', to);
  return params.toString() ? `?${params.toString()}` : '';
}

function renderChart(labels, values, title) {
  const ctx = document.getElementById('report-chart');
  if (state.chart) state.chart.destroy();
  state.chart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label: title, data: values }],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: true } },
    },
  });
}

async function runReport(reportName) {
  if (!state.currentCompanyId) throw new Error('Select a company first.');
  const query = reportFilters();
  const data = await api(`/api/companies/${state.currentCompanyId}/reports/${reportName}${query}`);

  if (reportName === 'trial-balance') {
    table('report-table', ['Ledger', 'Type', 'Group', 'Debit', 'Credit'], data.rows.map((r) => [r.ledger, r.type, r.group, r.debit, r.credit]));
    renderChart(data.rows.map((r) => r.ledger), data.rows.map((r) => r.debit - r.credit), 'Trial Balance (Net)');
  } else if (reportName === 'pnl') {
    const rows = [
      ...data.incomeRows.map((r) => ['Income', r.ledger, r.amount]),
      ...data.expenseRows.map((r) => ['Expense', r.ledger, r.amount]),
      ['Summary', 'Net Profit', data.netProfit],
    ];
    table('report-table', ['Category', 'Ledger', 'Amount'], rows);
    renderChart(['Income', 'Expense', 'Net'], [data.totalIncome, data.totalExpense, data.netProfit], 'P&L');
  } else if (reportName === 'balance-sheet') {
    const rows = [
      ...data.assets.map((r) => ['Asset', r.ledger, r.amount]),
      ...data.liabilities.map((r) => ['Liability/Equity', r.ledger, r.amount]),
      ['Total', 'Assets', data.totalAssets],
      ['Total', 'Liabilities', data.totalLiabilities],
    ];
    table('report-table', ['Side', 'Ledger', 'Amount'], rows);
    renderChart(['Assets', 'Liabilities'], [data.totalAssets, data.totalLiabilities], 'Balance Sheet');
  } else if (reportName === 'dashboard') {
    table('report-table', ['Metric', 'Value'], Object.entries(data).map(([k, v]) => [k, v]));
    renderChart(['Income', 'Expense', 'Profit'], [data.totalIncome, data.totalExpense, data.netProfit], 'Dashboard');
  } else {
    table('report-table', ['Ledger', 'Amount', 'Side', 'Age Days', 'Bucket'], data.map((r) => [r.ledger, r.amount, r.side, r.ageDays, r.bucket]));
    renderChart(
      ['0-30', '31-60', '61-90', '90+'],
      ['0-30', '31-60', '61-90', '90+'].map((bucket) => data.filter((x) => x.bucket === bucket).reduce((s, x) => s + Math.abs(Number(x.amount || 0)), 0)),
      'Outstanding Aging'
    );
  }
}

async function exportVouchers() {
  if (!state.currentCompanyId) throw new Error('Select a company first.');
  setStatus('Use an API client for authenticated exports with x-user/x-password headers.');
  window.open(`/api/companies/${state.currentCompanyId}/export/vouchers.xlsx${reportFilters()}`, '_blank');
}

document.getElementById('company-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  try {
    const form = new FormData(event.target);
    await api('/api/companies', {
      method: 'POST',
      body: JSON.stringify(Object.fromEntries(form.entries())),
    });
    event.target.reset();
    await loadCompanies();
    setStatus('Company created.');
  } catch (error) {
    setStatus(error.message);
  }
});

document.getElementById('company-select').addEventListener('change', async (event) => {
  state.currentCompanyId = event.target.value;
  await refreshCompanyData();
  setStatus('Company changed.');
});

document.getElementById('sync-company').addEventListener('click', async () => {
  try {
    await api(`/api/companies/${state.currentCompanyId}/sync`, { method: 'POST' });
    await refreshCompanyData();
    setStatus('Linked source synced.');
  } catch (error) {
    setStatus(error.message);
  }
});

document.getElementById('backup-company').addEventListener('click', async () => {
  try {
    const out = await api(`/api/companies/${state.currentCompanyId}/backup`, { method: 'POST' });
    setStatus(`Backup created: ${out.fileName}`);
  } catch (error) {
    setStatus(error.message);
  }
});

document.getElementById('ledger-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  try {
    const form = new FormData(event.target);
    const payload = Object.fromEntries(form.entries());
    payload.isParty = form.get('isParty') === 'on';
    await api(`/api/companies/${state.currentCompanyId}/ledgers`, {
      method: 'POST',
      body: JSON.stringify(payload),
    });
    event.target.reset();
    await refreshCompanyData();
    setStatus('Ledger added.');
  } catch (error) {
    setStatus(error.message);
  }
});

document.getElementById('voucher-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  try {
    const form = new FormData(event.target);
    const amount = Number(form.get('amount') || 0);
    const taxRate = Number(form.get('taxRate') || 0);
    const payload = {
      date: form.get('date'),
      type: form.get('type'),
      narration: form.get('narration'),
      entries: [
        { ledgerName: form.get('debitLedger'), debit: amount, credit: 0 },
        { ledgerName: form.get('creditLedger'), debit: 0, credit: amount },
      ],
      tax: { enabled: taxRate > 0, rate: taxRate },
    };

    await api(`/api/companies/${state.currentCompanyId}/vouchers`, {
      method: 'POST',
      body: JSON.stringify(payload),
    });
    event.target.reset();
    await refreshCompanyData();
    setStatus('Voucher posted.');
  } catch (error) {
    setStatus(error.message);
  }
});

document.querySelectorAll('.report-btn').forEach((button) => {
  button.addEventListener('click', async () => {
    try {
      await runReport(button.dataset.report);
      setStatus(`${button.dataset.report} report loaded.`);
    } catch (error) {
      setStatus(error.message);
    }
  });
});

document.getElementById('export-vouchers').addEventListener('click', exportVouchers);

loadCompanies().catch((error) => setStatus(error.message));
