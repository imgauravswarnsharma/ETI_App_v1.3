/**
 * =========================================================
 * SCRIPT: EXECUTION CONTEXT
 * =========================================================
 *
 * Layer:
 * - Core Execution State Layer
 *
 * Purpose:
 * - Maintain execution state across the system
 * - Provide a single source of truth for execution metadata
 * - Enable persistence and restoration for scheduler-based continuation
 *
 * System Role:
 * - Central state container shared across all layers
 * - Used by:
 *   - Controller (initialization)
 *   - Scheduler (resume + continuation)
 *   - Logger (read-only metadata enrichment)
 *   - Pipelines / Scripts (read-only usage)
 *
 * Core Responsibilities:
 * - Initialize execution context at entry points
 * - Store execution metadata (execution_id, trigger_type, etc.)
 * - Persist execution state for continuation
 * - Restore execution state during scheduler resume
 * - Provide access to current execution context
 *
 * Input Dependencies:
 * - Apps Script Services:
 *   - PropertiesService (state persistence)
 * - Utility Functions:
 *   - generateUUID_ (execution_id generation)
 *
 * Output Targets:
 * - Script Properties (persistent execution state)
 * - In-memory context (runtime access)
 *
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 *
 * 1. Execution entry (Controller / Manual run):
 *    - initExecutionContext_() is invoked
 *
 * 2. Initialization:
 *    - Generate execution_id
 *    - Set trigger_type (CONTROLLER / MANUAL)
 *    - Set run_context (default STANDALONE)
 *    - Set started_at timestamp
 *
 * 3. Context propagation:
 *    - Context stored in memory
 *    - Accessible via getExecutionContext_()
 *
 * 4. Pipeline execution:
 *    - Context is enriched (pipeline_name, run_context = PIPELINE)
 *
 * 5. During execution:
 *    - Context is read (not modified) by:
 *      • Logger
 *      • Business scripts
 *
 * 6. Scheduler interaction:
 *    - saveExecutionContext_() persists context before exit
 *
 * 7. Resume flow:
 *    - restoreExecutionContext_() reloads context
 *    - Execution continues using restored state
 *
 * 8. Execution completion:
 *    - Context may be cleared or reset
 *
 *
 * =========================================================
 * ALGORITHM (ACTUAL IMPLEMENTATION LOGIC)
 * =========================================================
 *
 * 1. Initialization:
 *    - Create context object with:
 *      • execution_id (UUID)
 *      • trigger_type
 *      • run_context
 *      • started_at
 *
 * 2. Context Storage:
 *    - Store context in global variable (runtime)
 *    - Persist to PropertiesService when required
 *
 * 3. Context Access:
 *    - getExecutionContext_() returns current context
 *    - getOrInitExecutionContext_() ensures availability
 *
 * 4. Persistence:
 *    - saveExecutionContext_() serializes context
 *    - Writes to Script Properties
 *
 * 5. Restoration:
 *    - restoreExecutionContext_() reads stored state
 *    - Rehydrates runtime context object
 *
 * 6. Context Enrichment:
 *    - Downstream layers update:
 *      • pipeline_name
 *      • function_name
 *      • switch_name
 *      • resume_count
 *
 * 7. Resume Handling:
 *    - Maintain flags:
 *      • is_resumed
 *      • incomplete_step
 *
 * 8. Execution Integrity:
 *    - Ensure single active context per execution
 *    - Prevent re-initialization in lower layers
 *
 *
 * =========================================================
 * DESIGN PRINCIPLES
 * =========================================================
 *
 * - Single source of truth for execution state
 * - Context initialized only at entry points
 * - Downstream layers enrich, not recreate
 * - Persistence-first design for continuation safety
 * - Lightweight and globally accessible
 * - No business logic embedded
 *
 *
 * =========================================================
 * IDENTITY & SAFETY
 * =========================================================
 *
 * - Only Controller (or entry layer) initializes context
 * - Lower layers MUST NOT reinitialize context
 * - Safe for resume-based execution cycles
 * - Prevents state fragmentation across system
 * - Ensures deterministic execution tracking
 *
 * =========================================================
 */



/* --- GLOBAL STATE --- */
var EXECUTION_CONTEXT = null;


/*
=========================================================
MODULE: INITIALIZATION
=========================================================*/
/* 
-------------------------------------
INIT EXECUTION CONTEXT
-------------------------------------*/
function initExecutionContext_(options = {}) {

  EXECUTION_CONTEXT = {
    execution_id: Utilities.getUuid(),

    // EXECUTION IDENTITY 
    pipeline_name: options.pipeline_name || null,
    function_name: options.function_name || null,
    switch_name: options.switch_name || null,

    run_context: options.run_context || "STANDALONE",
    trigger_type: options.trigger_type || "MANUAL",

    started_at: new Date(),

    // SCHEDULER STATE
    is_resumed: false,
    resume_count: 0,
    incomplete_step: false,

    // PIPELINE STATE
    function_index: 0,

    // LOGGER STATE
    last_flush_log_execution_id: null,
  };

  return EXECUTION_CONTEXT;
}


/*
=========================================================
MODULE: ACCESS & SAFE RETRIEVAL
=========================================================*/
/* 
-------------------------------------
GET EXECUTION CONTEXT
-------------------------------------*/
function getExecutionContext_(){
  return EXECUTION_CONTEXT;
}


/*
-------------------------------------
GET OR INIT (SAFE FALLBACK)
-------------------------------------*/
function getOrInitExecutionContext_(options = null){

  let ctx = getExecutionContext_();
  if (ctx) {
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
=========================================================
MODULE: PERSISTENCE LAYER
=========================================================*/
/*
-------------------------------------
STORAGE ACCESS 
-------------------------------------*/
function getExecutionContextStore_(){
  return PropertiesService.getScriptProperties();
}


/* 
-------------------------------------
SAVE CONTEXT 
-------------------------------------*/
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
  } 

  catch (err) {
    console.error('saveExecutionContext_ failed', err);
    return false;
  }
}


/* 
-------------------------------------
RESTORE CONTEXT 
-------------------------------------*/
function restoreExecutionContext_(){

  try {
    const store = getExecutionContextStore_();
    const saved = store.getProperty('ETI_EXECUTION_CONTEXT');

    if (!saved) return false;
    const parsed = JSON.parse(saved);
    if (!parsed || !parsed.execution_id) return false;

    // DEFAULT SAFETY
    if (parsed.function_index === undefined) parsed.function_index = 0;
    if (parsed.incomplete_step === undefined) parsed.incomplete_step = false;

    EXECUTION_CONTEXT = parsed;
    return true;
  } 

  catch (err) {
    console.error('restoreExecutionContext_ failed', err);
    return false;
  }
}


/*
=========================================================
MODULE: LIFECYCLE MANAGEMENT
=========================================================*/
/*
-------------------------------------
 CLEAR CONTEXT 
-------------------------------------*/
function clearExecutionContext_(){

  try {
    const store = getExecutionContextStore_();
    store.deleteProperty('ETI_EXECUTION_CONTEXT');
  } 

  catch (err) {
    console.error('clearExecutionContext_ failed', err);
  }
  EXECUTION_CONTEXT = null;
}


/* 
-------------------------------------
RESTORE OR INIT 
-------------------------------------*/
function restoreOrInitExecutionContext_(options = null){
  const restored = restoreExecutionContext_();

  if (restored) return getExecutionContext_();
  return getOrInitExecutionContext_(options);
}