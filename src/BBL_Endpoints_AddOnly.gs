// BBL_Endpoints_AddOnly.gs — new, non-colliding endpoints (ADD-ONLY)
function BBL_ping(){ return { ok:true, data:{ now:(new Date()).toISOString() } }; }
function BBL_getAccountEstimates(){
  try { return { ok:true, data: BBL_AccountsService.getEstimateMap() }; }
  catch(e){ BBL_log('BBL_getAccountEstimates_error', String(e)); return { ok:false, message:String(e) }; }
}
function BBL_listCategoriesByType(type){
  try { return { ok:true, data: BBL_CategoriesService.listByType(type) }; }
  catch(e){ BBL_log('BBL_listCategoriesByType_error', { type:type, err:String(e) }); return { ok:false, message:String(e) }; }
}
