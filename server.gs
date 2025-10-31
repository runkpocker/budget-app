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
    // Failsafe error surface instead of blank screen
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
function _ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function _safeNamed(name) {
  try {
    const r = _ss().getRangeByName(name);
    return r ? r.getDisplayValues() : [];
  } catch (err) {
    return [];
  }
}

// --- Month picker endpoint that RESPECTS Summary!B4 data validation ---
// --- This is the ORIGINAL, V1 working code ---
function setSummaryMonthFromIso(isoYYYYMM) {
  if (!/^\d{4}-\d{2}$/.test(String(isoYYYYMM || ''))) {
    throw new Error('Expected YYYY-MM (e.g., 2025-10)');
  }
  const [yStr, mStr] = isoYYYYMM.split('-');
  const y = Number(yStr), m = Number(mStr);
  // 1..12
  if (!(y > 1900 && m >= 1 && m <= 12)) throw new Error('Invalid month');
  const sh = _ss().getSheetByName('Summary');
  if (!sh) throw new Error('Summary sheet not found');

  const b4 = sh.getRange('B4');
  const dv = b4.getDataValidation();
  // Normalize value to YYYY-MM for matching (handles dates, strings, "Oct 2025", etc.)
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
    // Validation: allowed range
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      const rng = args[0];
      if (rng) {
        const vals = rng.getValues();
        const disps = rng.getDisplayValues();
        // Prefer raw values
        for (let i = 0; i < vals.length; i++) {
          if (monthKey(vals[i][0]) === targetKey) {
            b4.setValue(vals[i][0]);
            return _returnSummary_(b4);
          }
        }
        // Fallback to display values
        for (let i = 0; i < disps.length; i++) {
          if (monthKey(disps[i][0]) === targetKey) {
            b4.setValue(vals[i][0]);
            return _returnSummary_(b4);
          }
        }
      }
      throw new Error(`Selected month (${isoYYYYMM}) is not in the allowed list for Summary!B4`);
    }

    // Validation: allowed list
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

  // No validation on B4: write first-of-month at local noon (avoid DST backshift)
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

// Kept for future use (UI no longer auto-reads it on load)
// --- This is the ORIGINAL, V1 working code ---
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

// --- This is the ORIGINAL, V1 working code ---
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

// --- START: This is the ONLY change from V1 (our successful fix) ---
function listCategories() {
  try {
    const ss = _ss();
    const inc = ss.getRangeByName('IncomeCats').getValues().flat();
    const exp = ss.getRangeByName('ExpenseCats').getValues().flat();
    
    // Combine, filter out blanks, get unique, and sort
    const set = new Set([...inc, ...exp].filter(Boolean));
    return Array.from(set).sort();
    
  } catch (e) {
    Logger.log('Error in listCategories: ' + e.message);
    // Fallback to old method if named ranges fail
    const tx = readSheetObjects_(SHEET_TRANSACTIONS);
    const b  = readSheetObjects_(SHEET_BUDGETS);
    const set = new Set([...tx.map(r => r.Category), ...b.map(r => r.Category)].filter(Boolean));
    return Array.from(set).sort();
  }
}

function listTypes() {
  // Your setup script hardcodes these validation lists.
  return ['Expense', 'Income', 'Transfer'];
}
// --- END: Successful fix ---

// --- This is the ORIGINAL, V1 working code ---
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

// --- This is the ORIGINAL, V1 working code ---
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
    // --- THIS IS THE ORIGINAL, TIMEZONE-SAFE LOGIC ---
    let yyyy = '1970', mm = '01', dd = '01', yyyymm = '1970-01';
    let d = r.Date;

    // Handle string dates like '2025-11-01' manually to avoid timezone bugs
    if (typeof d === 'string' && d.includes('-')) {
      const parts = d.split('T')[0].split('-'); // Get YYYY-MM-DD
      if (parts.length === 3) {
        yyyy = parts[0];
       
        mm = parts[1].padStart(2, '0');
        dd = parts[2].padStart(2, '0');
        yyyymm = `${yyyy}-${mm}`;
      }
    } else {
      // Fallback for Sheets numeric dates or other formats
      if (typeof d === 'number') d = new Date(Math.round((d - 25569) * 86400 * 1000));
      const dateObj = d instanceof Date ? d : new Date(d);
      if (!isNaN(dateObj)) {
     
           yyyy = dateObj.getFullYear();
        mm = String(dateObj.getMonth() + 1).padStart(2,'0');
        dd = String(dateObj.getDate()).padStart(2,'0');
        yyyymm = `${yyyy}-${mm}`;
      }
    }

    const amtStr = String(r.Amount ||
 '').replace(/,/g, '');
    const amt = Number(amtStr);
    // --- END ORIGINAL LOGIC ---

    return {
      DateISO: `${yyyy}-${mm}-${dd}`,
      Month: yyyymm, // Use the manually parsed month
      Account: String(r.Account || ''),
      Type: String(r.Type || ''),
      Category: String(r.Category || ''),
      Amount: Number.isFinite(amt) ?
 amt : null,
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

// --- This is the ORIGINAL, V1 working code ---
function normalizeBudgetMonthCell_(v) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM');
  const s = String(v || '').trim();
  // dd/mm/yyyy or d/m/yyyy (or '-' separators)
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    return Utilities.formatDate(d, tz, 'yyyy-MM');
  }
  // "October 2025" / "Oct 2025" / "Oct-2025"
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
  return isNaN(d2) ?
 '' : Utilities.formatDate(d2, tz, 'yyyy-MM');
}

/***** === DEBUGGING FUNCTION === *****/

function debugDataExplorer() {
  Logger.log('--- STARTING DEBUG ---');
  Logger.log('Testing with hardcoded month: 2025-11');
  const testMonth = '2025-11';
  
  try {
    // --- Test Budget Sheet ---
    Logger.log('--- TESTING BUDGET SHEET ---');
    const budgetRows = readSheetObjects_(SHEET_BUDGETS);
    Logger.log('Raw Budget rows found: ' + budgetRows.length);
    if (budgetRows.length > 0) {
      Logger.log('Raw Budget Headers: ' + JSON.stringify(Object.keys(budgetRows[0])));
      Logger.log('Raw Budget Row 1: ' + JSON.stringify(budgetRows[0]));
      
      const budgetResult = getBudgetView({ month: testMonth });
      Logger.log('Filtered Budget rows: ' + budgetResult.rows.length);
      if (budgetResult.rows.length > 0) {
        Logger.log('Filtered Budget Row 1: ' + JSON.stringify(budgetResult.rows[0]));
      }
    } else {
      Logger.log('Budget sheet appears empty or has no data rows.');
    }

    // --- Test Transactions Sheet ---
    Logger.log('--- TESTING TRANSACTIONS SHEET ---');
    const txRows = readSheetObjects_(SHEET_TRANSACTIONS);
    Logger.log('Raw Transaction rows found: ' + txRows.length);
    if (txRows.length > 0) {
      Logger.log('Raw TX Headers: ' + JSON.stringify(Object.keys(txRows[0])));
      Logger.log('Raw TX Row 1: ' + JSON.stringify(txRows[0]));
      
      const txResult = getTransactionView({ month: testMonth, page: 1, pageSize: 50 });
      Logger.log('Filtered TX rows: ' + txResult.rows.length);
      Logger.log('TX Total Rows: ' + txResult.total);
      if (txResult.rows.length > 0) {
        Logger.log('Filtered TX Row 1: ' + JSON.stringify(txResult.rows[0]));
      }
    } else {
      Logger.log('Transactions sheet appears empty or has no data rows.');
    }
    
  } catch (e) {
    Logger.log('!!! DEBUG FAILED WITH ERROR !!!');
    Logger.log(e.message);
    Logger.log(e.stack);
  }
  Logger.log('--- DEBUG COMPLETE ---');
}