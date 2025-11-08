// BBL_Services_Categories.gs — cached list by type (ADD-ONLY)
var BBL_CategoriesService = (function(){
  var TTL = 120; // seconds
  function key(t){ return 'bbl:cats:'+t; }
  function listByType(type){
    type = String(type || '').toLowerCase();
    var cache = CacheService.getScriptCache();
    try {
      var c = cache.get(key(type));
      if (c) return JSON.parse(c);
    } catch(e){}

    var values = null;
    if (type === 'income') values = BBL_readNamedRange_(BBL_Schema.Ranges.IncomeCats);
    else if (type === 'expense') values = BBL_readNamedRange_(BBL_Schema.Ranges.ExpenseCats);

    var items = [];
    if (values && values.length){
      items = values.map(function(row){ return (row[0]||'').toString().trim(); })
                    .filter(function(s){ return s; });
    } else {
      var sh = BBL_getSheet_(BBL_Schema.Categories.name);
      var last = sh.getLastRow();
      if (last >= 2){
        var rows = sh.getRange(2,1,last-1,2).getValues();
        items = rows.filter(function(r){ return String(r[0]).toLowerCase() === type; })
                    .map(function(r){ return (r[1]||'').toString().trim(); })
                    .filter(function(s){ return s; });
      }
    }
    try { cache.put(key(type), JSON.stringify(items), TTL); } catch(e){}
    return items;
  }
  return Object.freeze({ listByType:listByType });
})();
