// BBL_Logging.gs — optional structured logger (ADD-ONLY)
function BBL_log(tag, payload){
  try {
    var entry = { ts:(new Date()).toISOString(), tag:String(tag||'') };
    if (payload !== undefined) entry.payload = payload;
    var line = JSON.stringify(entry);
    Logger.log(line);
    try { console.log(line); } catch (e) {}
  } catch (e) {
    Logger.log('BBL_log failed: ' + e);
  }
}
