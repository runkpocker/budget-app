// BBL_Services_Accounts.gs — read-only (ADD-ONLY)
var BBL_AccountsService = (function(){
  function listEstimates(){
    var sh = BBL_getSheet_(BBL_Schema.Accounts.name);
    var last = sh.getLastRow();
    if (last < 2) return [];
    var rows = sh.getRange(2,1,last-1,3).getValues();
    return rows.map(function(r){
      return { account: r[0], type: r[1], estBal: Number(r[2]) || 0 };
    }).filter(function(o){ return o.account; });
  }
  function getEstimateMap(){
    var map = {}; listEstimates().forEach(function(o){ map[String(o.account).trim()] = o.estBal; });
    return map;
  }
  return Object.freeze({ listEstimates:listEstimates, getEstimateMap:getEstimateMap });
})();
