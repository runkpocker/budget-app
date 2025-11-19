// Bunny Balut Budget — "Button Bunny" (single page, validated month setter)
const SPREADSHEET_ID = '1-KZorvFCZi-Ic7a54hx4p2s0fqM6LHt_wVcSQkxAS6c';

// Returns the currently deployed web app base URL (dev/exec).
function getWebAppBaseUrl() {
  var url = ScriptApp.getService().getUrl();
  if (!url) throw new Error('This script is not deployed as a web app.');
  return url;
}

// Convenience: base + ?view=...
function getWebAppUrl(view) {
  var base = getWebAppBaseUrl();
  var v = String(view || 'index');
  return base + (base.indexOf('?') > -1 ? '&' : '?') + 'view=' + encodeURIComponent(v);
}

function doGet(e) {
  try {
    var view = (e && e.parameter && e.parameter.view) || 'landing';
    var file = (view === 'index') ? 'index' : 'landing';
    
    // Inject the canonical deployment URL
    var t = HtmlService.createTemplateFromFile(file);
    t.BASE_URL = ScriptApp.getService().getUrl();
    
    return t
      .evaluate()
      .setTitle('Bunny Balut Budget')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    var msg = '<pre style="color:#ff8080;background:#1b1b1b;padding:12px;border-radius:8px;">'
            + 'Renderer failed\n' + String(err) + '</pre>';
    return HtmlService.createHtmlOutput(msg).setTitle('Bunny Balut Budget');
  }
}

// HTML include helper
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// Spreadsheet helpers
function _ss() { 
  // Optimization: If container-bound, you could use SpreadsheetApp.getActiveSpreadsheet();
  return SpreadsheetApp.openById(SPREADSHEET_ID); 
}

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

/** === MAIN: categories per Type with discovery === */
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
    
    // Helper to check against list or range
    const checkList = (list) => {
      for (let i = 0; i < list.length; i++) {
        // list[i] might be an array [val] or just val
        const val = Array.isArray(list[i]) ? list[i][0] : list[i];
        if (monthKey(val) === targetKey) {
          b4.setValue(val);
          return true;
        }
      }
      return false;
    };

    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      const rng = args[0];
      if (rng) {
        if (checkList(rng.getValues())) return _returnSummary_(b4);
        if (checkList(rng.getDisplayValues())) return _returnSummary_(b4);
      }
      throw new Error(`Selected month (${isoYYYYMM}) is not in the allowed list for Summary!B4`);
    }

    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      if (checkList(args[0] || [])) return _returnSummary_(b4);
      throw new Error(`Selected month (${isoYYYYMM}) is not in the allowed list for Summary!B4`);
    }
  }

  // Fallback if no validation or validation passed manually
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
  return ['Income', 'Expense', 'Transfer']; 
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

/** Utility: A1 helpers */
function colToA1(col){
  let s = "";
  while (col > 0) { 
    let m = (col - 1) % 26; 
    s = String.fromCharCode(65 + m) + s;
    col = (col - m - 1) / 26; 
  }
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
  const types = listTypes(); 
  return { accounts, types };
}

// Parse "YYYY-MM-DD" as a LOCAL midnight date (no UTC shift).
function parseLocalDate_(yyyy_mm_dd) {
  if (!yyyy_mm_dd) return null;
  const [y, m, d] = String(yyyy_mm_dd).split('-').map(Number);
  if (!y || !m || !d) return null;
  return new Date(y, m - 1, d); // local midnight in project timezone
}

/** SAFETY-FIRST Quick Add **/
function quickAddTransaction(payload) {
  const ss = _ss();
  const sh = ss.getSheetByName('Transactions');
  if (!sh) throw new Error('Missing "Transactions" sheet');

  const meta = inspectTransactionsSheet();
  const lastCol = meta.lastCol;
  const tmplRow = meta.templateRow || 2;

  // --- Validate inputs
  const d = parseLocalDate_(payload.date);
  if (isNaN(d)) throw new Error('Invalid date'); // 'd' is used for validation only
  const type = String(payload.type || '').trim();
  const cat  = String(payload.category || '').trim();
  const acct = String(payload.account || '').trim();
  const desc = String(payload.description || '').trim();
  const amt  = Number(String(payload.amount || '').replace(/,/g, ''));
  
  if (!type || !cat || !acct || !Number.isFinite(amt)) {
    throw new Error('Missing required fields');
  }

  // --- 1) Snapshot the TEMPLATE FORMULAS BEFORE we shift anything
  const tmplFormulas = sh.getRange(tmplRow, 1, 1, lastCol).getFormulas()[0];

  // --- 2) Insert a NEW ROW at row 2 so the new record becomes row 2
  sh.insertRowsBefore(tmplRow, 1);

  // --- 3) Copy ONLY format + data validation from the (shifted) old template (now row 3) to the new row 2
  sh.getRange(tmplRow + 1, 1, 1, lastCol).copyTo(
    sh.getRange(tmplRow, 1, 1, lastCol),
    { formatOnly: true }
  );

  // --- 4) Restore the template formulas into the new row 2
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
    if (!col) return;
    const colMeta = meta.columns[col - 1];
    if (colMeta.hasFormulaInTemplate) return; // DO NOT touch formula cols
    sh.getRange(tmplRow, col).setValue(value);
  };

  // --- 5) Fill inputs
  // FIX: Write the date as a simple String (YYYY-MM-DD) to prevent timezone shifting.
  // Since 'd' was created in the Script Timezone, formatting it back to string in the same timezone preserves the input.
  safeWrite('Date',        Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  
  safeWrite('Account',     acct);
  safeWrite('Type',        type);
  safeWrite('Category',    cat);
  safeWrite('Amount',      amt);
  safeWrite('Description', desc);

  return { ok: true, row: tmplRow };
}

/** Debug Helper */
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