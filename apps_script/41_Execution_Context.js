/**
 * =========================================================
 * SCRIPT: EXECUTION CONTEXT
 * =========================================================
 *
 * Layer:
 * - Core Execution State Layer
 *
 * PURPOSE:
 * - Maintain execution state across system
 * - Provide initialization + access
 * - SINGLE SOURCE OF TRUTH for execution metadata
 *
 * DESIGN RULES:
 * In-memory primary state with controlled persistence
 * Persistence handled via dedicated functions
 * =========================================================
 */


/*
-------------------------------------
GLOBAL EXECUTION CONTEXT (SINGLETON)
-------------------------------------
*/
var EXECUTION_CONTEXT = null;


/*
-------------------------------------
INIT EXECUTION CONTEXT
-------------------------------------
*/
function initExecutionContext_(options = {}) {

  EXECUTION_CONTEXT = {
    execution_id: Utilities.getUuid(),

    /* 
    =========================
     EXECUTION IDENTITY (ENTRY FILLS)
    =========================*/
    pipeline_name: options.pipeline_name || null,
    function_name: options.function_name || null,
    switch_name: options.switch_name || null,      // Controller

    run_context: options.run_context || "STANDALONE",
    trigger_type: options.trigger_type || "MANUAL",

    started_at: new Date(),

    is_resumed: false,                            // Scheduler
    resume_count: 0,                              // Scheduler
    incomplete_step: false,                       // Scheduler    

    function_index: 0,                            // Pipeline

    last_flush_log_execution_id: null,            // Logger


  };

  return EXECUTION_CONTEXT;
}


/*
-------------------------------------
GET EXECUTION CONTEXT
-------------------------------------
*/
function getExecutionContext_(){
  return EXECUTION_CONTEXT;
}



/*
-------------------------------------
GET OR INIT (SAFE FALLBACK)
-------------------------------------
*/
function getOrInitExecutionContext_(options = null){

  let ctx = getExecutionContext_();

  if (ctx) {

    // Optional enrichment (only if provided)
    if (options){
      if (options.pipeline_name && !ctx.pipeline_name) {
        ctx.pipeline_name = options.pipeline_name;
      }
      if (options.function_name && !ctx.function_name) {
        ctx.function_name = options.function_name;
      }
      if (options.switch_name && !ctx.switch_name) {
        ctx.switch_name = options.switch_name;
      }
    }

    return ctx;
  }

  return initExecutionContext_(options || {
    run_context: 'STANDALONE',
    trigger_type: 'MANUAL'
  });
}


/*
-------------------------------------
INTERNAL STORAGE ACCESS (ABSTRACTION)
-------------------------------------
*/
function getExecutionContextStore_(){
  return PropertiesService.getScriptProperties();
}


/*
-------------------------------------
SAVE EXECUTION CONTEXT (PERSIST)
-------------------------------------
*/
function saveExecutionContext_(){

  const ctx = getExecutionContext_();
  if (!ctx) return false;

  try {

    const store = getExecutionContextStore_();

    store.setProperty(
      'ETI_EXECUTION_CONTEXT',
      JSON.stringify(ctx)
    );

    return true;

  } catch (err) {

    console.error('saveExecutionContext_ failed', err);
    return false;
  }
}


/*
-------------------------------------
RESTORE EXECUTION CONTEXT (FROM STORE)
-------------------------------------
*/
function restoreExecutionContext_(){

  try {

    const store = getExecutionContextStore_();

    const saved = store.getProperty('ETI_EXECUTION_CONTEXT');

    if (!saved) return false;

    const parsed = JSON.parse(saved);

    if (!parsed || !parsed.execution_id) return false;

    /*
    -------------------------------------
    DEFAULT SAFETY (ENSURE FIELDS)
    -------------------------------------
    */
    if (parsed.function_index === undefined) parsed.function_index = 0;
    if (parsed.incomplete_step === undefined) parsed.incomplete_step = false;

    EXECUTION_CONTEXT = parsed;

    return true;

  } catch (err) {

    console.error('restoreExecutionContext_ failed', err);
    return false;
  }
}


/*
-------------------------------------
CLEAR EXECUTION CONTEXT (POST COMPLETE)
-------------------------------------
*/
function clearExecutionContext_(){

  try {

    const store = getExecutionContextStore_();
    store.deleteProperty('ETI_EXECUTION_CONTEXT');

  } catch (err) {
    console.error('clearExecutionContext_ failed', err);
  }

  EXECUTION_CONTEXT = null;
}


/*
-------------------------------------
RESTORE OR INIT EXECUTION CONTEXT
-------------------------------------
*/
function restoreOrInitExecutionContext_(options = null){

  const restored = restoreExecutionContext_();

  if (restored) return getExecutionContext_();

  return getOrInitExecutionContext_(options);
}