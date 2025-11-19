// BBL_Repo.gs — tiny helpers (ADD-ONLY)
function BBL_getSS_(){ return SpreadsheetApp.openById(BBL_CONFIG.SPREADSHEET_ID); }
function BBL_getSheet_(name){
  var sh = BBL_getSS_().getSheetByName(name);
  if (!sh) throw new Error('BBL: Missing sheet ' + name);
  return sh;
}
function BBL_readNamedRange_(name){
  var r = BBL_getSS_().getRangeByName(name);
  return r ? r.getDisplayValues() : null;
}
