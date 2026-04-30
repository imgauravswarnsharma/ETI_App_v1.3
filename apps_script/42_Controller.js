/**
 * =========================================================
 * SCRIPT: EXECUTION CONTROLLER
 * =========================================================
 *
 * Layer:
 * - Execution Control Layer (Controller)
 *
 * Purpose:
 * - Serve as the primary execution entry point
 * - Route execution based on automation switches
 * - Initialize execution context for all runs
 * - Control lifecycle of function and pipeline execution
 *
 * System Role:
 * - Entry layer for all triggered executions
 * - Bridges UI / trigger events to backend execution
 * - Coordinates with Scheduler for controlled continuation
 * - Ensures consistent execution initialization and completion
 *
 * Core Responsibilities:
 * - Read automation switches and determine execution target
 * - Initialize execution context (trigger_type, switch_name)
 * - Route execution to:
 *   • Pipeline functions
 *   • Standalone functions
 * - Wrap execution inside controlled environment
 * - Handle execution success / failure / exit states
 * - Integrate with scheduler for continuation-based execution
 *
 * Input Dependencies:
 * - Execution Context Layer:
 *   - initExecutionContext_
 *   - getExecutionContext_
 *   - saveExecutionContext_
 * - Scheduler Layer:
 *   - exitAndScheduleContinuation_
 *   - finalizeExecutionFromScheduler_
 * - Logger Layer:
 *   - ETI_log_
 *   - flushLogs_
 * - Switch System:
 *   - getAutomationSwitchMap_
 *
 * Output Targets:
 * - Execution Context (initialized + updated state)
 * - Logger (execution lifecycle logs)
 * - Scheduler (continuation control when required)
 *
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 *
 * 1. Trigger event occurs:
 *    - Manual execution OR automation trigger (onChange / button)
 *
 * 2. Controller entry function executes:
 *    - Reads automation switch map
 *
 * 3. Determine execution target:
 *    - Identify active switch
 *    - Map switch → function or pipeline
 *
 * 4. Initialize execution context:
 *    - trigger_type = CONTROLLER / MANUAL
 *    - switch_name = selected switch
 *    - execution_id generated
 *
 * 5. Controlled execution wrapper:
 *    - ETI_executeControlledFunction_(targetFunction)
 *
 * 6. Inside controlled execution:
 *    a. Execute target function
 *    b. Monitor for:
 *       - timeout exit (scheduler)
 *       - error conditions
 *
 * 7. If scheduler exit triggered:
 *    - exitAndScheduleContinuation_ handles continuation
 *    - controller stops execution
 *
 * 8. If execution completes:
 *    - finalizeExecutionFromScheduler_ OR direct finalize
 *
 * 9. Finalization:
 *    - Update execution state
 *    - Log completion status
 *    - Flush logs
 *
 *
 * =========================================================
 * ALGORITHM (ACTUAL IMPLEMENTATION LOGIC)
 * =========================================================
 *
 * 1. Load switch map using getAutomationSwitchMap_
 *
 * 2. Identify active switch:
 *    - Iterate switches
 *    - Select enabled switch
 *
 * 3. Resolve execution target:
 *    - Map switch → function name or pipeline
 *    - Validate existence in global scope
 *
 * 4. Initialize execution context:
 *    - Set:
 *      • execution_id
 *      • trigger_type
 *      • switch_name
 *      • run_context (default STANDALONE)
 *
 * 5. Controlled execution:
 *    - Wrap target call inside try-catch
 *    - Use ETI_executeControlledFunction_
 *
 * 6. Error handling:
 *    - Capture error
 *    - Log via ETI_logError_
 *    - Propagate failure state
 *
 * 7. Scheduler interaction:
 *    - If ctx.incomplete_step === true:
 *      → exit (scheduler will resume)
 *
 * 8. Execution completion:
 *    - Clear execution markers
 *    - Finalize via finalizeExecutionFromScheduler_
 *
 * 9. Log flushing:
 *    - Ensure flushLogs_ is called in finally block
 *
 *
 * =========================================================
 * DESIGN PRINCIPLES
 * =========================================================
 *
 * - Single entry point for all executions
 * - Switch-driven execution routing
 * - Context-first execution initialization
 * - Controlled execution wrapper (no direct calls)
 * - Scheduler-compatible design (resume-safe)
 * - Separation of concerns (controller does not contain business logic)
 *
 *
 * =========================================================
 * IDENTITY & SAFETY
 * =========================================================
 *
 * - Only layer allowed to initialize execution context
 * - Does NOT contain business logic
 * - Safe for repeated trigger execution
 * - Prevents uncontrolled execution paths
 * - Ensures consistent execution lifecycle across system
 *
 * =========================================================
 */


/*
=========================================================
MODULE: SWITCH READING & MAPPING
=========================================================*/
/*
-------------------------------------
READ AUTOMATION SWITCH MAP
-------------------------------------*/
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
=========================================================
MODULE: LOGGING MODE SWITCHES
=========================================================*/
/* ---| ACTION LOG ENABLED |--- */
function isActionLogEnabled_(){

  const switches = getAutomationSwitchMap_();
  return switches["Enable_Action_Log"] === true;
}

/* ---| EXECUTION LOG ENABLED |--- */
function isExecutionLogEnabled_(){

  const switches = getAutomationSwitchMap_();
  return switches["Enable_Execution_Log"] === true;
}

/* ---| CONSOLE LOG ENABLED |--- */
function isConsoleLogEnabled_(){

  const switches = getAutomationSwitchMap_();
  return switches["Enable_Console_Log"] === true;
}

/* ---| DEBUG MODE ENABLED |--- */
function isDebugModeEnabled_(){

  const switches = getAutomationSwitchMap_();
  const mode = switches["Access_Mode"];
  return mode === "GOD" || mode === "DEV_L2" || mode === "DEV_L1";
}


/*
=========================================================
MODULE: EXECUTION STATE UTILITIES
=========================================================*/
/*
-------------------------------------
GET EXECUTION STATUS MAP
-------------------------------------*/
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
=========================================================
MODULE: SWITCH CONTROL UTILITIES
=========================================================*/
/* 
-------------------------------------
CHECK EXECUTABLE SWITCH
-------------------------------------*/
function isExecutableSwitch_(name){

  return (
    name.startsWith("Run_") ||
    name.startsWith("Populate_") ||
    name.startsWith("Promote_")
  );
}


/*
-------------------------------------
RESET SWITCH
-------------------------------------*/
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
SET EXECUTION STATUS 
-------------------------------------*/
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
SET EXECUTION TIMESTAMP 
-------------------------------------*/
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
SET EXECUTION DURATION 
-------------------------------------*/
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
SET LOG MESSAGE
-------------------------------------*/
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
=========================================================
MODULE: HELPER UTILITIES
=========================================================*/
/* 
-------------------------------------
FORMAT DURATION 
-------------------------------------*/
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
=========================================================
MODULE: CONTROLLER ENGINE
=========================================================*/
/*
-------------------------------------
AUTOMATION CONTROLLER
-------------------------------------*/
function automationController_onChange(e){
  const switches = getAutomationSwitchMap_();
  const statusMap = getExecutionStatusMap_();

  /* ---| PHASE 1 — MARK WAITING (NO LOCK) |--- */
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

  /* ---| PHASE 2 — LOCK CONTROL |--- */
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    return;
  }

  try {
    /* ---| PREVENT PARALLEL RUNNING |--- */
    const updatedStatusMap = getExecutionStatusMap_();

    for(const key in updatedStatusMap){
      if(updatedStatusMap[key] === "RUNNING"){
        return;
      }
    }

    /* ---| FUNCTION MAP |--- */
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


    /* ---| LOOP — PROCESS QUEUE |--- */
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

        /* ---| CHECK IF SCHEDULER WILL CONTINUE |--- */
        if (getExecutionContext_()?.incomplete_step === true) {

          setExecutionStatus_(nextSwitch, "RUNNING");
          setLogMessage_(nextSwitch, "Execution paused");

          return;
        }

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

      resetSwitch_(nextSwitch);
    }

  } finally {
    lock.releaseLock();
  }
}


/*
=========================================================
MODULE: EXECUTION WRAPPER
=========================================================*/
/*
-----------------------------
CONTROLLED EXECUTION WRAPPER
-----------------------------*/
function ETI_executeControlledFunction_(switchName, fn){
  
  const functionName = fn.name || '';

  initExecutionContext_({
    run_context: 'STANDALONE',
    trigger_type: 'CONTROLLER',
    function_name: fn.name || null,
    switch_name: switchName
  });

  saveExecutionContext_();

  try {
    fn();
  } catch (err) {

    ETI_logError_(
      'Controller',
      functionName,
      '',
      err,
      'CONTROLLED_EXECUTION'
    );

    throw err;

  } finally {
    flushLogs_();
  }
}


/*
=========================================================
MODULE: SCHEDULER FINALIZATION
=========================================================*/
/*
-------------------
FINALIZE EXECUTION
-------------------*/
function finalizeExecutionFromScheduler_(status = 'SUCCESS', message = ''){

  const ctx = getExecutionContext_();
  if (!ctx) return;

  const switchName = ctx?.switch_name;
  if (!switchName) return;

  try {

    const durationMs = new Date() - new Date(ctx.started_at);

    setExecutionDuration_(switchName, durationMs);
    setExecutionStatus_(switchName, status);
    setLogMessage_(switchName, message || status);

    resetSwitch_(switchName);

  } catch (err) {
    console.error('finalizeExecutionFromScheduler_ failed', err);
  }

  clearExecutionContext_();
}