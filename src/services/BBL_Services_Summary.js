// BBL_Services_Summary.gs — misc helpers (ADD-ONLY)
var BBL_SummaryService = (function(){
  function isValidMonth(isoMonth){ return /^\d{4}-(0[1-9]|1[0-2])$/.test(String(isoMonth||'')); }
  return Object.freeze({ isValidMonth:isValidMonth });
})();
