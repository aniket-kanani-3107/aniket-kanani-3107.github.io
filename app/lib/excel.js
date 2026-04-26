const fs = require('fs');
const path = require('path');
const readXlsxFile = require('read-excel-file/node');
const writeXlsxFile = require('write-excel-file/node');

function headerIndexMap(headerRow) {
  const map = new Map();
  headerRow.forEach((value, index) => {
    map.set(String(value || '').trim().toLowerCase(), index);
  });
  return map;
}

function rowValue(row, map, key) {
  const idx = map.get(key.toLowerCase());
  if (idx === undefined) return null;
  return row[idx];
}

async function parseLinkedWorkbook(filePath) {
  if (!fs.existsSync(filePath)) throw new Error('Linked file not found.');
  const stat = fs.statSync(filePath);
  if (stat.size > 10 * 1024 * 1024) throw new Error('Linked file is too large (max 10MB).');
  const ext = path.extname(filePath).toLowerCase();

  if (ext === '.json') {
    let content;
    try {
      content = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    } catch {
      throw new Error('Invalid linked JSON format.');
    }
    return {
      ledgers: Array.isArray(content.ledgers) ? content.ledgers : [],
      vouchers: Array.isArray(content.vouchers) ? content.vouchers : [],
    };
  }

  if (ext !== '.xlsx') {
    throw new Error('Only .json or .xlsx linked files are supported.');
  }

  const result = { ledgers: [], vouchers: [] };

  try {
    const ledgerRows = await readXlsxFile(filePath, { sheet: 'Ledgers' });
    if (ledgerRows.length > 1) {
      const map = headerIndexMap(ledgerRows[0]);
      result.ledgers = ledgerRows.slice(1).filter((row) => row.length).map((row) => ({
        name: String(rowValue(row, map, 'name') || '').trim(),
        type: String(rowValue(row, map, 'type') || 'Asset').trim(),
        group: String(rowValue(row, map, 'group') || 'General').trim(),
        openingBalance: Number(rowValue(row, map, 'openingBalance') || 0),
        openingType: String(rowValue(row, map, 'openingType') || 'debit').toLowerCase(),
        isParty: ['true', '1', 'yes'].includes(String(rowValue(row, map, 'isParty') || '').toLowerCase()),
      })).filter((x) => x.name);
    }
  } catch {
    // optional sheet
  }

  try {
    const voucherRows = await readXlsxFile(filePath, { sheet: 'Vouchers' });
    if (voucherRows.length > 1) {
      const map = headerIndexMap(voucherRows[0]);
      result.vouchers = voucherRows.slice(1).filter((row) => row.length).map((row) => {
        const amount = Number(rowValue(row, map, 'amount') || 0);
        const debitLedger = String(rowValue(row, map, 'debitLedger') || '').trim();
        const creditLedger = String(rowValue(row, map, 'creditLedger') || '').trim();
        return {
          date: String(rowValue(row, map, 'date') || new Date().toISOString().slice(0, 10)),
          type: String(rowValue(row, map, 'type') || 'Journal').trim(),
          narration: String(rowValue(row, map, 'narration') || '').trim(),
          entries: [
            { ledgerName: debitLedger, debit: amount, credit: 0 },
            { ledgerName: creditLedger, debit: 0, credit: amount },
          ],
        };
      }).filter((x) => x.entries[0].ledgerName && x.entries[1].ledgerName && x.entries[0].debit > 0);
    }
  } catch {
    // optional sheet
  }

  return result;
}

async function writeExcelReport(filePath, sheetName, headers, rows) {
  const data = [
    headers.map((h) => ({ value: h, fontWeight: 'bold' })),
    ...rows.map((row) => row.map((cell) => ({ value: cell ?? '' }))),
  ];

  await writeXlsxFile(data, {
    filePath,
    sheet: sheetName,
  });
}

module.exports = {
  parseLinkedWorkbook,
  writeExcelReport,
};
