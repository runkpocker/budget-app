// BBL_Schema.gs — single source of truth (ADD-ONLY)
var BBL_Schema = Object.freeze({
  Accounts: { name: BBL_CONFIG.SHEETS.Accounts, cols: { Account:1, Type:2, EstBal:3 } },
  Categories: { name: BBL_CONFIG.SHEETS.Categories, cols: { Type:1, Name:2 } },
  Budget: { name: BBL_CONFIG.SHEETS.Budget },
  Transactions: { name: BBL_CONFIG.SHEETS.Transactions },
  Summary: { name: BBL_CONFIG.SHEETS.Summary, cols: { Account:1, Net:4, EstBal:5 } },
  Ranges: Object.freeze({
    IncomeCats: BBL_CONFIG.NAMED_RANGES.IncomeCats,
    ExpenseCats: BBL_CONFIG.NAMED_RANGES.ExpenseCats,
    AccAct: BBL_CONFIG.NAMED_RANGES.AccAct
  })
});
