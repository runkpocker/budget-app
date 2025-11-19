// BBL_CONFIG.gs — central, read-only config (ADD-ONLY)
var BBL_CONFIG = Object.freeze({
  SPREADSHEET_ID: '1-KZorvFCZi-Ic7a54hx4p2s0fqM6LHt_wVcSQkxAS6c',
  SHEETS: Object.freeze({
    Accounts: 'Accounts',
    Categories: 'Categories',
    Budget: 'Budget',
    Transactions: 'Transactions',
    Summary: 'Summary'
  }),
  NAMED_RANGES: Object.freeze({
    IncomeCats: 'IncomeCats',
    ExpenseCats: 'ExpenseCats',
    AccAct: 'AccAct'
  }),
  FEATURE_FLAGS: Object.freeze({
    splashEgg: true
  })
});
