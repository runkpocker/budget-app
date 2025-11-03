// Bunny Balut Budget — "Button Bunny" (single page, validated month setter)
const SPREADSHEET_ID = '1-KZorvFCZi-Ic7a54hx4p2s0fqM6LHt_wVcSQkxAS6c';

// Serve single-page app (index.html)
function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Bunny Balut Budget')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    const msg = '<pre style="color:#ff8080;background:#1b1b1b;padding:12px;border-radius:8px;">'
      + 'Renderer failed\n' + String(err) + '</pre>';
    return HtmlService.createHtmlOutput(msg).setTitle('Bunny Balut Budget');
  }
}

// HTML include helper
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// Spreadsheet helpers
function _ss() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function _safeNamed(name) {
  try {
    const r = _ss().getRangeByName(name);
    return r ? r.getDisplayValues() : [];
  } catch (_) { return []; }
}

/** Return flattened, trimmed, unique, sorted strings */
function _uniqClean(arr) {
  return Array.from(
    new Set((arr || []).flat().map(v => String(v || '').trim()).filter(Boolean))
  ).sort();
}

/** Safe read of a named range (by exact name). */
function _getNamed(name) {
  try {
    const rng = _ss().getRangeByName(name);
    return rng ? _uniqClean(rng.getValues()) : [];
  } catch (_) {
    return [];
  }
}

/** Try multiple candidate names, then regex-match over all named ranges. */
function _resolveNamedValues(candidates, regexes) {
  // 1) Strong candidates (exact names you probably use)
  for (const name of candidates) {
    const v = _getNamed(name);
    if (v.length) return v;
  }
  // 2) Heuristic search across all named ranges
  try {
    const all = _ss().getNamedRanges();
    for (const r of all) {
      const nm = r.getName();
      if (regexes.some(re => re.test(nm))) {
        const v = _uniqClean(r.getRange().getValues());
        if (v.length) return v;
      }
    }
  } catch (_) {}
  return [];
}

/** Fallback readers from sheets (Type/Category on Categories, col A on Accounts) */
function _catsFromCategoriesSheet(kind /* 'income' | 'expense' */) {
  const sh = _ss().getSheetByName('Categories');
  if (!sh) return [];
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return [];
  const header = (sh.getRange(1,1,1,lastCol).getValues()[0] || []).map(v => String(v||'').trim().toLowerCase());
  const typeIdx = header.indexOf('type');
  const catIdx  = header.indexOf('category');
  if (typeIdx < 0 || catIdx < 0) return [];
  const out = [];
  const rows = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  for (const r of rows) {
    const ty = String(r[typeIdx] || '').trim().toLowerCase();
    const ca = String(r[catIdx]  || '').trim();
    if (ca && ty === kind) out.push(ca);
  }
  return _uniqClean(out);
}

function _accountsFromSheet() {
  const sh = _ss().getSheetByName('Accounts');
  if (!sh) return [];
  const n = Math.max(0, sh.getLastRow() - 1);
  return n ? _uniqClean(sh.getRange(2,1,n,1).getValues()) : [];
}

/** === MAIN: categories per Type with discovery ===
 *  - Income:   look for Income category named ranges (many common spellings)
 *  - Expense:  same
 *  - Transfer: try account lists by name; else Accounts sheet
 *  - Blank:    union Income+Expense
 */
function listCategoriesByType(type) {
  const t = String(type || '').trim().toLowerCase();

  if (t === 'transfer') {
    const accounts = _resolveNamedValues(
      [
        'Accounts', 'AccountList', 'AccountsList', 'TransferAccounts',
        'Transfer_Accounts', 'Accounts_Names'
      ],
      [/^accounts?$/i, /transfer.*(acct|account)/i, /(acct|account).*(list|names)/i]
    );
    return accounts.length ? accounts : _accountsFromSheet();
  }

  if (t === 'income' || t === 'expense') {
    const isIncome = t === 'income';
    const vals = _resolveNamedValues(
      isIncome
        ? ['IncomeCats','Income_Categories','Categories_Income','Cat_Income','IncomeCategoryList']
        : ['ExpenseCats','Expense_Categories','Categories_Expense','Cat_Expense','ExpenseCategoryList'],
      isIncome
        ? [/income.*cat/i, /cat.*income/i, /income.*categor/i]
        : [/expense.*cat/i, /cat.*expense/i, /expense.*categor/i]
    );
    if (vals.length) return vals;
    // Per-type fallback from Categories sheet
    return _catsFromCategoriesSheet(isIncome ? 'income' : 'expense');
  }

  // Unknown/blank → union of income+expense
  const inc = listCategoriesByType('Income');
  const exp = listCategoriesByType('Expense');
  return _uniqClean([inc, exp]);
}


/** Fallback: derive categories by scanning the Categories sheet (Type, Category) */
function deriveCatsFromSheet_() {
  const ss  = _ss();
  const sh  = ss.getSheetByName('Categories');
  if (!sh) return { income: [], expense: [] };

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return { income: [], expense: [] };

  const header = (sh.getRange(1,1,1,lastCol).getValues()[0] || []).map(String);
  const typeIdx = header.findIndex(h => h.trim().toLowerCase() === 'type');
  const catIdx  = header.findIndex(h => h.trim().toLowerCase() === 'category');
  if (typeIdx < 0 || catIdx < 0) return { income: [], expense: [] };

  const rows = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  const inc = [], exp = [];
  for (const r of rows) {
    const t = String(r[typeIdx] || '').trim().toLowerCase();
    const c = String(r[catIdx]  || '').trim();
    if (!c) continue;
    if (t === 'income')  inc.push(c);
    if (t === 'expense') exp.push(c);
  }
  const uniq = (arr) => Array.from(new Set(arr)).sort();
  return { income: uniq(inc), expense: uniq(exp) };
}

// --- Month picker endpoint that RESPECTS Summary!B4 data validation ---
function setSummaryMonthFromIso(isoYYYYMM) {
  if (!/^\d{4}-\d{2}$/.test(String(isoYYYYMM || ''))) {
    throw new Error('Expected YYYY-MM (e.g., 2025-10)');
  }
  const [yStr, mStr] = isoYYYYMM.split('-');
  const y = Number(yStr), m = Number(mStr);
  if (!(y > 1900 && m >= 1 && m <= 12)) throw new Error('Invalid month');

  const sh = _ss().getSheetByName('Summary');
  if (!sh) throw new Error('Summary sheet not found');

  const b4 = sh.getRange('B4');
  const dv = b4.getDataValidation();

  const monthKey = (v) => {
    if (v instanceof Date && !isNaN(v)) {
      return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2,'0')}`;
    }
    const t = String(v || '').trim();
    if (/^\d{4}-\d{2}$/.test(t)) return t;
    let m1 = t.match(/^(\d{4})[-\/\. ](\d{1,2})$/);
    if (m1) return `${m1[1]}-${String(m1[2]).padStart(2,'0')}`;
    let m2 = t.match(/^([A-Za-z]+)\s+(\d{4})$/);
    if (m2) {
      const names = ['january','february','march','april','may','june','july','august','september','october','november','december'];
      const idx = names.indexOf(m2[1].slice(0,3).toLowerCase());
      if (idx >= 0) return `${m2[2]}-${String(idx + 1).padStart(2,'0')}`;
    }
    const d = new Date(t);
    if (!isNaN(d)) return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2,'0')}`;
    return null;
  };
  const targetKey = `${y}-${String(m).padStart(2,'0')}`;

  if (dv) {
    const type = dv.getCriteriaType();
    const args = dv.getCriteriaValues();
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      const rng = args[0];
      if (rng) {
        const vals = rng.getValues();
        const disps = rng.getDisplayValues();
        for (let i = 0; i < vals.length; i++) {
          if (monthKey(vals[i][0]) === targetKey) {
            b4.setValue(vals[i][0]);
            return _returnSummary_(b4);
          }
        }
        for (let i = 0; i < disps.length; i++) {
          if (monthKey(disps[i][0]) === targetKey) {
            b4.setValue(vals[i][0]);
            return _returnSummary_(b4);
          }
        }
      }
      throw new Error(`Selected month (${isoYYYYMM}) is not in the allowed list for Summary!B4`);
    }

    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      const list = (args[0] || []).map(String);
      for (const s of list) {
        if (monthKey(s) === targetKey) {
          b4.setValue(s);
          return _returnSummary_(b4);
        }
      }
      throw new Error(`Selected month (${isoYYYYMM}) is not in the allowed list for Summary!B4`);
    }
  }

  b4.setValue(new Date(y, m - 1, 1, 12, 0, 0, 0));
  return _returnSummary_(b4);
}

// Build response payload for the UI
function _returnSummary_(b4Range) {
  const actVsBud = _safeNamed('ActVsBud');
  const inc      = _safeNamed('IncActVsBud');
  const exp      = _safeNamed('ExpActVsBud');
  return { b4: b4Range.getDisplayValue(), actVsBud, income: inc, expense: exp };
}

function getCurrentMonthIso() {
  const sh = _ss().getSheetByName('Summary');
  if (!sh) return { iso: '' };
  const raw = sh.getRange('B4').getValue();
  const disp = sh.getRange('B4').getDisplayValue();
  const toIso = (v) => {
    if (v instanceof Date && !isNaN(v)) {
      return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2,'0')}`;
    }
    const t = String(v || '').trim();
    if (/^\d{4}-\d{2}$/.test(t)) return t;
    let m = t.match(/^(\d{4})[-\/\. ](\d{1,2})$/);
    if (m) return `${m[1]}-${String(m[2]).padStart(2,'0')}`;
    m = t.match(/^([A-Za-z]+)\s+(\d{4})$/);
    if (m) {
      const names = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
      const idx = names.indexOf(m[1].slice(0,3).toLowerCase());
      if (idx >= 0) return `${m[2]}-${String(idx + 1).padStart(2,'0')}`;
    }
    return '';
  };
  return { iso: toIso(raw) || toIso(disp) || '' };
}

/***** === Data Explorer API === *****/
const SHEET_BUDGETS = 'Budget';
const SHEET_TRANSACTIONS = 'Transactions';
const BUDGET_HEADERS = ['Month','Type','Category','Budget'];
const TX_HEADERS     = ['Date','Account','Type','Category','Description','Amount'];

function readSheetObjects_(sheetName) {
  const sh = _ss().getSheetByName(sheetName);
  if (!sh) throw new Error('Missing sheet: ' + sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).map(row => {
    const o = {};
    headers.forEach((h, i) => o[h] = row[i]);
    return o;
  });
}

function asYYYYMM_(v) {
  if (!v) return '';
  if (v instanceof Date && !isNaN(v)) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2,'0');
    return `${y}-${m}`;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-\/. ](\d{1,2})$/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2,'0')}`;
  const d = new Date(s);
  if (!isNaN(d)) return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
  return '';
}

function listTypes() {
  return ['Income', 'Expense', 'Transfer']; // order matters for your defaults
}

function getBudgetView(args) {
  args = args || {};
  const month = asYYYYMM_(args && args.month ? args.month : '');
  const rows = readSheetObjects_(SHEET_BUDGETS);
  const headers = Object.keys(rows[0] || {});
  const missing = BUDGET_HEADERS.filter(h => headers.indexOf(h) < 0);
  if (missing.length) throw new Error('Budget sheet missing headers: ' + missing.join(', '));
  const mapped = rows.map(r => {
    const budgetStr = String(r.Budget || '').replace(/,/g, '');
    return {
      Month: asYYYYMM_(r.Month),
      Type: String(r.Type||''),
      Category: String(r.Category||''),
      Budget: Number(budgetStr)
    };
  });
  const filtered = mapped.filter(r =>
    (!month || r.Month === month) &&
    r.Month && r.Type && Number.isFinite(r.Budget)
  );
  return { rows: filtered };
}

function getTransactionView(args) {
  args = args || {};
  const month = asYYYYMM_(args && args.month ? args.month : '');
  const type = args && args.type ? String(args.type) : '';
  const category = args && args.category ? String(args.category) : '';
  const page = Math.max(1, Number(args && args.page || 1));
  const pageSize = Math.min(200, Math.max(10, Number(args && args.pageSize || 50)));

  const rows = readSheetObjects_(SHEET_TRANSACTIONS);
  const headers = Object.keys(rows[0] || {});
  const missing = TX_HEADERS.filter(h => headers.indexOf(h) < 0);
  if (missing.length) throw new Error('Transactions sheet missing headers: ' + missing.join(', '));
  const normalized = rows.map(r => {
    let yyyy = '1970', mm = '01', dd = '01', yyyymm = '1970-01';
    let d = r.Date;

    if (typeof d === 'string' && d.includes('-')) {
      const parts = d.split('T')[0].split('-');
      if (parts.length === 3) {
        yyyy = parts[0];
        mm = parts[1].padStart(2, '0');
        dd = parts[2].padStart(2, '0');
        yyyymm = `${yyyy}-${mm}`;
      }
    } else {
      if (typeof d === 'number') d = new Date(Math.round((d - 25569) * 86400 * 1000));
      const dateObj = d instanceof Date ? d : new Date(d);
      if (!isNaN(dateObj)) {
        yyyy = dateObj.getFullYear();
        mm = String(dateObj.getMonth() + 1).padStart(2,'0');
        dd = String(dateObj.getDate()).padStart(2,'0');
        yyyymm = `${yyyy}-${mm}`;
      }
    }

    const amtStr = String(r.Amount || '').replace(/,/g, '');
    const amt = Number(amtStr);

    return {
      DateISO: `${yyyy}-${mm}-${dd}`,
      Month: yyyymm,
      Account: String(r.Account || ''),
      Type: String(r.Type || ''),
      Category: String(r.Category || ''),
      Amount: Number.isFinite(amt) ? amt : null,
      Description: String(r.Description || r.Note || '')
    };
  });
  const filtered = normalized.filter(r =>
    (!month || r.Month === month) &&
    (!type || r.Type === type) &&
    (!category || r.Category === category) &&
    r.DateISO && r.Account && r.Amount !== null
  );
  const total = filtered.length;
  const start = (page - 1) * pageSize;
  const slice = filtered.slice(start, start + pageSize);
  return { rows: slice, page, pageSize, total, pages: Math.max(1, Math.ceil(total / pageSize)) };
}

function normalizeBudgetMonthCell_(v) {
  const tz = _ss().getSpreadsheetTimeZone();
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    return Utilities.formatDate(d, tz, 'yyyy-MM');
  }
  m = s.match(/^([A-Za-z]+)[\s\-]+(\d{4})$/);
  if (m) {
    const names = ['january','february','march','april','may','june','july','august','september','october','november','december'];
    const token = m[1].toLowerCase().replace(/\./g,'');
    let idx = names.indexOf(token);
    if (idx < 0) idx = names.map(n => n.slice(0,3)).indexOf(token.slice(0,3));
    if (idx >= 0) {
      const d = new Date(Number(m[2]), idx, 1);
      return Utilities.formatDate(d, tz, 'yyyy-MM');
    }
  }
  const d2 = new Date(s);
  return isNaN(d2) ? '' : Utilities.formatDate(d2, tz, 'yyyy-MM');
}

/** Utility: A1 helpers */
function colToA1(col){
  let s = "";
  while (col > 0) { let m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = (col - m - 1) / 26; }
  return s;
}

/** Inspect Transactions sheet */
function inspectTransactionsSheet(){
  const ss  = _ss();
  const sh  = ss.getSheetByName('Transactions');
  if (!sh) throw new Error('Missing "Transactions" sheet');

  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error('Transactions sheet has no columns');

  const headers  = (sh.getRange(1, 1, 1, lastCol).getValues()[0] || []).map(v => String(v || '').trim());
  const tmplRow  = 2;
  const formulas = sh.getRange(tmplRow, 1, 1, lastCol).getFormulas()[0];
  const values   = sh.getRange(tmplRow, 1, 1, lastCol).getValues()[0];

  const cols = [];
  for (let c = 1; c <= lastCol; c++) {
    const header  = headers[c - 1] || '';
    const formula = formulas[c - 1] || '';
    let hidden = false;
    try { hidden = sh.isColumnHiddenByUser(c); } catch (_) {}
    cols.push({
      index: c,
      letter: colToA1(c),
      header,
      isHidden: !!hidden,
      hasFormulaInTemplate: formula !== '',
      sampleFormula: formula || null,
      sampleValue: values[c - 1]
    });
  }

  return {
    sheetName: sh.getName(),
    lastCol,
    templateRow: tmplRow,
    columns: cols,
    inputColumns: cols.filter(col => !col.hasFormulaInTemplate).map(c => c.header)
  };
}

/** Quick Add meta (accounts + types). Categories will be fetched live by type. */
function getQuickAddMeta(){
  const ss   = _ss();
  const accts= ss.getSheetByName('Accounts');
  if (!accts) throw new Error('Missing "Accounts" sheet');

  const accounts = (accts.getRange(2,1,Math.max(0, accts.getLastRow()-1), 1).getValues() || [])
    .map(r => String(r[0] || '').trim())
    .filter(Boolean);

  const types = listTypes(); // ['Income','Expense','Transfer']
  return { accounts, types };
}

/** SAFETY-FIRST Quick Add **/
function quickAddTransaction(payload) {
  const ss = _ss();
  const sh = ss.getSheetByName('Transactions');
  if (!sh) throw new Error('Missing "Transactions" sheet');

  const meta = inspectTransactionsSheet();       // has header map + which cols are formula
  const lastCol = meta.lastCol;
  const tmplRow = meta.templateRow || 2;         // template row is row 2

  // --- Validate inputs
  const d = new Date(payload.date);
  if (isNaN(d)) throw new Error('Invalid date');

  const type = String(payload.type || '').trim();
  const cat  = String(payload.category || '').trim();
  const acct = String(payload.account || '').trim();
  const desc = String(payload.description || '').trim();
  const amt  = Number(String(payload.amount || '').replace(/,/g, ''));
  if (!type || !cat || !acct || !Number.isFinite(amt)) {
    throw new Error('Missing required fields');
  }

  // --- 1) Snapshot the TEMPLATE FORMULAS BEFORE we shift anything
  const tmplFormulas = sh.getRange(tmplRow, 1, 1, lastCol).getFormulas()[0]; // array of strings ('' if no formula)

  // --- 2) Insert a NEW ROW at row 2 so the new record becomes row 2
  sh.insertRowsBefore(tmplRow, 1);               // new row is now at row 2; old template moved to row 3

  // --- 3) Copy ONLY format + data validation from the (shifted) old template (now row 3) to the new row 2
  sh.getRange(tmplRow + 1, 1, 1, lastCol).copyTo(
    sh.getRange(tmplRow, 1, 1, lastCol),
    { formatOnly: true }
  );

  // --- 4) Restore the template formulas into the new row 2 (exactly the same formulas as template)
  for (let c = 1; c <= lastCol; c++) {
    const f = tmplFormulas[c - 1] || '';
    if (f) {
      sh.getRange(tmplRow, c).setFormula(f);
    }
  }

  // --- Helper: header -> column index
  const byHeader = {};
  meta.columns.forEach(col => byHeader[col.header] = col.index);

  // --- Helper: write to an input column ONLY (skip formula columns)
  const safeWrite = (header, value) => {
    const col = byHeader[header];
    if (!col) return;                                    // header missing
    const colMeta = meta.columns[col - 1];
    if (colMeta.hasFormulaInTemplate) return;            // DO NOT touch formula cols
    sh.getRange(tmplRow, col).setValue(value);
  };

  // --- 5) Fill inputs (Date, Account, Type, Category, Amount, Description)
  safeWrite('Date',        d);
  safeWrite('Account',     acct);
  safeWrite('Type',        type);
  safeWrite('Category',    cat);
  safeWrite('Amount',      amt);
  safeWrite('Description', desc);

  // Done. New record is row 2, all hidden/formula columns intact.
  return { ok: true, row: tmplRow };
}



/***** DEBUG (optional) *****/
function quickAddDiagnostics(){
  const meta = inspectTransactionsSheet();
  return {
    templateRow: meta.templateRow,
    inputColumns: meta.inputColumns,
    hiddenColumns: meta.columns.filter(c => c.isHidden).map(c => ({index:c.index, letter:c.letter, header:c.header})),
    formulaColumns: meta.columns.filter(c => c.hasFormulaInTemplate).map(c => ({
      index: c.index, letter: c.letter, header: c.header, formula: c.sampleFormula
    }))
  };
}
