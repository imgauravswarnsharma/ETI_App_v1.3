/*
-------------------------------------
Utility Function — Read Switch Table
-------------------------------------
*/
function getAutomationSwitchMap_(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Automation_Control");

  if(!sheet) return {};

  const lastCol = sheet.getLastColumn();
  if(lastCol === 0) return {};

  const header = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const values = sheet.getRange(2,1,1,lastCol).getValues()[0];

  const map = {};

  for(let i=0;i<header.length;i++){

    const name = header[i];
    let value = values[i];

    if(!name) continue;

    if(value === "TRUE") value = true;
    if(value === "FALSE") value = false;

    map[name] = value;
  }

  return map;
}


/*
-------------------------------------
LOGGING MODE SWITCHES (SINGLE SOURCE)
-------------------------------------
// ACTION LOG
*/
function isActionLogEnabled_(){
  const switches = getAutomationSwitchMap_();
  return switches["Enable_Action_Log"] === true;
}


// EXECUTION LOG
function isExecutionLogEnabled_(){
  const switches = getAutomationSwitchMap_();
  return switches["Enable_Execution_Log"] === true;
}

// CONSOLE LOG
function isConsoleLogEnabled_(){
  const switches = getAutomationSwitchMap_();
  return switches["Enable_Console_Log"] === true;
}

// DEBUG MODE
function isDebugModeEnabled_(){
  const switches = getAutomationSwitchMap_();
  const mode = switches["Access_Mode"];
  return mode === "GOD" || mode === "DEV_L2" || mode === "DEV_L1";
}


/*
-------------------------------------
Utility — Get Execution Status
-------------------------------------
*/
function getExecutionStatusMap_(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Automation_Control");

  if(!sheet) return {};

  const header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(3,1,1,sheet.getLastColumn()).getValues()[0];

  const map = {};

  for(let i=0;i<header.length;i++){
    if(header[i]){
      map[header[i]] = values[i];
    }
  }

  return map;
}


/*
-------------------------------------
Utility — Check Executable Switch
-------------------------------------
*/
function isExecutableSwitch_(name){

  return (
    name.startsWith("Run_") ||
    name.startsWith("Populate_") ||
    name.startsWith("Promote_")
  );
}


/*
-------------------------------------
Utility — Reset Switch
-------------------------------------
*/
function resetSwitch_(switchName){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Automation_Control");

  if(!sheet) return;

  const header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const index = header.indexOf(switchName);

  if(index !== -1){
    sheet.getRange(2,index+1).setValue(false);
  }
}


/*
-------------------------------------
Utility — Set Status
-------------------------------------
*/
function setExecutionStatus_(switchName, status){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Automation_Control");

  if(!sheet) return;

  const header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const index = header.indexOf(switchName);

  if(index !== -1){
    sheet.getRange(3,index+1).setValue(status);
  }
}


/*
-------------------------------------
Utility — Timestamp
-------------------------------------
*/
function setExecutionTimestamp_(switchName, timestamp){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Automation_Control");

  if(!sheet) return;

  const header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const index = header.indexOf(switchName);

  if(index !== -1){
    sheet.getRange(4,index+1).setValue(timestamp);
  }
}


/*
-------------------------------------
Utility — Duration
-------------------------------------
*/
function setExecutionDuration_(switchName, durationMs){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Automation_Control");

  if(!sheet) return;

  const header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const index = header.indexOf(switchName);

  if(index !== -1){
    sheet.getRange(5,index+1).setValue(formatDuration_(durationMs));
  }
}


/*
-------------------------------------
Helper — Format Duration
-------------------------------------
*/
function formatDuration_(ms){

  const totalSeconds = Math.floor(ms / 1000);

  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  if(hours > 0) return `${hours}h ${minutes}m ${seconds}s`;
  if(minutes > 0) return `${minutes}m ${seconds}s`;

  return `${seconds}s`;
}


/*
-------------------------------------
Utility — Dashboard Execution Tracking: Message
-------------------------------------
*/
function setLogMessage_(switchName, message){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Automation_Control");

  if(!sheet) return;

  const header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const index = header.indexOf(switchName);

  if(index !== -1){
    sheet.getRange(6,index+1).setValue(message);
  }
}


/*
-------------------------------------
Controller (QUEUE ENABLED)
-------------------------------------
*/
function automationController_onChange(e){

  const switches = getAutomationSwitchMap_();
  const statusMap = getExecutionStatusMap_();

  /*
  -------------------------------------
  PHASE 1 — MARK WAITING (NO LOCK)
  -------------------------------------
  */
  for(const name in switches){

    if(!isExecutableSwitch_(name)) continue;

    if(switches[name] === true){

      const status = statusMap[name];

      if(status !== "RUNNING" && status !== "WAITING"){
        setExecutionStatus_(name, "WAITING");
        setLogMessage_(name, "Queued");
      }
    }
  }

  /*
  -------------------------------------
  PHASE 2 — LOCK CONTROL
  -------------------------------------
  */
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    return;
  }

  try {

    /*
    -------------------------------------
    PREVENT PARALLEL RUNNING
    -------------------------------------
    */
    const updatedStatusMap = getExecutionStatusMap_();

    for(const key in updatedStatusMap){
      if(updatedStatusMap[key] === "RUNNING"){
        return;
      }
    }

    /*
    -------------------------------------
    FUNCTION MAP
    -------------------------------------
    */
    const functionMap = {

      "Run_Transaction_Pipeline": pipeline_transactions_,
      "Run_Item_Pipeline": pipeline_items_,
      "Run_Brand_Pipeline": pipeline_brands_,
      "Run_Product_Pipeline": pipeline_products_,
      "Run_Item_Brand_Mapping_Pipeline": pipeline_item_brand_mapping_,
      "Run_Item_Brand_Product_Mapping_Pipeline": pipeline_item_brand_product_mapping_,

      "Populate_Items_Staging": populateStagingLookupItems_FromTransactionResolution,
      "Populate_Brands_Staging": populateStagingLookupBrands_FromTransactionResolution,
      "Populate_Products_Staging": populateStagingLookupProducts_FromTransactionResolution,

      "Promote_Items_To_Lookup": promoteApprovedItems_FromStaging_ToLookup,
      "Promote_Brands_To_Lookup": promoteApprovedBrands_FromStaging_ToLookup,
      "Promote_Products_To_Lookup": promoteApprovedProducts_FromStaging_ToLookup,

      "Run_Sheets_Metadata_Pipeline": sheets_metadata_pipeline_,
      "Run_Scripts_Metadata_Pipeline": scripts_metadata_pipeline_,
      "Run_Full_Metadata_Pipeline": full_metadata_pipeline_,

      "Run_Access_Mode_Pipeline": pipeline_access_mode_
    };

    /*
    -------------------------------------
    LOOP — PROCESS QUEUE
    -------------------------------------
    */
    const MAX_ITERATIONS = 20;
    let iteration = 0;

    while(iteration < MAX_ITERATIONS){

      iteration++;

      const currentStatusMap = getExecutionStatusMap_();

      let nextSwitch = null;

      for(const key in currentStatusMap){
        if(currentStatusMap[key] === "WAITING"){
          nextSwitch = key;
          break;
        }
      }

      if(!nextSwitch) break;

      const fn = functionMap[nextSwitch];

      if(!fn){
        setExecutionStatus_(nextSwitch, "FAILED");
        setLogMessage_(nextSwitch, "No function mapped");
        resetSwitch_(nextSwitch);
        continue;
      }

      const startTime = new Date();

      setExecutionStatus_(nextSwitch, "RUNNING");
      setExecutionTimestamp_(nextSwitch, startTime);
      setLogMessage_(nextSwitch, nextSwitch);

      try{

        ETI_executeControlledFunction_(nextSwitch, fn);
        
        /*-------------------------------------
            CHECK IF SCHEDULER WILL CONTINUE
        -------------------------------------*/
        if (getExecutionContext_()?.incomplete_step === true) {

          /*
          -------------------------------------
          PAUSE STATE (SCHEDULER HANDOVER)
          -------------------------------------
          */
          setExecutionStatus_(nextSwitch, "RUNNING");
          setLogMessage_(nextSwitch, "Execution paused");

          return;         // EXIT CONTROLLER LOOP

        }

        /*-------------------------------------
            SUCCESS PATH (NON-SCHEDULER)
         -------------------------------------*/
        const durationMs = new Date() - startTime;

        setExecutionDuration_(nextSwitch, durationMs);
        setExecutionStatus_(nextSwitch, "SUCCESS");
        setLogMessage_(nextSwitch, "Completed successfully");

      } catch(err){

        const durationMs = new Date() - startTime;

        setExecutionDuration_(nextSwitch, durationMs);
        setExecutionStatus_(nextSwitch, "FAILED");

        const msg = err && err.message ? err.message : err;
        setLogMessage_(nextSwitch, "FAILED: " + msg);

        console.error(nextSwitch, err);
      }

      // -------------------------------------
      // ALWAYS RESET SWITCH AFTER FINAL STATE
      // -------------------------------------
      resetSwitch_(nextSwitch);
      }

  }
  finally{
    lock.releaseLock();
  }
}


/*
-------------------------------------
Execution Wrapper — Controller Layer
-------------------------------------
*/
function ETI_executeControlledFunction_(switchName, fn){
  
  const functionName = fn.name || '';

  /*
  -------------------------------------
  LOGGER EXECUTION CONTEXT INITIALIZER
  -------------------------------------
  */
  initExecutionContext_({
    run_context: 'STANDALONE',
    trigger_type: 'CONTROLLER',
    function_name: fn.name || null,
    switch_name: switchName
  });
  
    saveExecutionContext_();   // ensure persistence before execution

  try {
    fn(); // Execute Function
  } catch (err) {

    /* -------------------------------------
       CRITICAL: LOG ERROR TO LOGGER
      ------------------------------------- */
    ETI_logError_(
      'Controller',
      functionName,
      '',                  // no sheet
      err,
      'CONTROLLED_EXECUTION'
    );

    throw err;  // preserve original flow

  } finally {
    flushLogs_(); // Log Flush for Batch Writing (Critical)
  }
}


/*
-------------------------------------
FINALIZE EXECUTION (CONTROLLER OWNED)
-------------------------------------
*/
function finalizeExecutionFromScheduler_(status = 'SUCCESS', message = ''){

  const ctx = getExecutionContext_();
  if (!ctx) return;

  const switchName = ctx?.switch_name;
  if (!switchName) return;

  try {

    const durationMs = new Date() - new Date(ctx.started_at);

    /*
    -------------------------------------
    APPLY FINAL STATE (FROM SCHEDULER)
    -------------------------------------
    */
    setExecutionDuration_(switchName, durationMs);
    setExecutionStatus_(switchName, status);
    setLogMessage_(switchName, message || status);

    resetSwitch_(switchName);

  } catch (err) {
    console.error('finalizeExecutionFromScheduler_ failed', err);
  }

  /*
  -------------------------------------
  CLEAR EXECUTION CONTEXT
  -------------------------------------
  */
  clearExecutionContext_();
}