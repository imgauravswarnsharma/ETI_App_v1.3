/**
   * =========================================================
   * SCRIPT: PIPELINE ORCHESTRATION ENGINE
   * =========================================================
   *
   * Layer:
   * - Execution Orchestration Layer (Pipelines)
   *
   * Purpose:
   * - Execute a sequence of functions in a controlled, ordered manner
   * - Enable grouped execution of related processes
   * - Support resume-aware execution using execution context
   *
   * System Role:
   * - Sits between Controller and individual script functions
   * - Orchestrates multiple functions as a single execution unit
   * - Works with Scheduler for continuation handling
   *
   * Core Responsibilities:
   * - Define ordered execution steps (function arrays)
   * - Execute functions sequentially
   * - Maintain execution pointer for resume
   * - Persist execution state between steps
   * - Handle controlled exit for scheduler continuation
   *
   * Input Dependencies:
   * - Execution Context Layer:
   *   - getExecutionContext_
   *   - getOrInitExecutionContext_
   *   - saveExecutionContext_
   * - Scheduler Layer:
   *   - exitAndScheduleContinuation_
   * - Logger Layer:
   *   - ETI_log_
   *   - flushLogs_
   *
   * Output Targets:
   * - Execution Context (step pointer, pipeline_name)
   * - Logger (pipeline-level logs)
   *
   *
   * =========================================================
   * EXECUTION FLOW
   * =========================================================
   *
   * 1. Pipeline function invoked by Controller
   *
   * 2. Execution context initialized (if not present)
   *
   * 3. Pipeline metadata setup:
   *    - ctx.pipeline_name = current pipeline
   *    - ctx.run_context = PIPELINE
   *
   * 4. Define execution steps:
   *    - Array of functions (ordered execution)
   *
   * 5. Resume handling:
   *    - Determine start index using ctx.function_index
   *
   * 6. Loop execution:
   *    For each step:
   *    a. Update context pointer:
   *       - function_index
   *       - function_name
   *    b. Persist context (saveExecutionContext_)
   *    c. Execute function
   *
   * 7. Controlled exit handling:
   *    - If function returns 'EXIT' → stop execution
   *    - If ctx.incomplete_step === true → stop execution
   *
   * 8. Continue execution:
   *    - Move to next function until all steps completed
   *
   * 9. Pipeline completion:
   *    - Log completion
   *    - Clear execution pointer
   *
   * 10. Final flush:
   *    - flushLogs_() executed in finally block
   *
   *
   * =========================================================
   * ALGORITHM (ACTUAL IMPLEMENTATION LOGIC)
   * =========================================================
   *
   * 1. Initialize execution context (if not already initialized)
   *
   * 2. Set pipeline-level metadata:
   *    - pipeline_name
   *    - run_context = PIPELINE
   *
   * 3. Define steps array:
   *    - [fn1, fn2, fn3, ...]
   *
   * 4. Determine start index:
   *    - startIndex = ctx.function_index || 0
   *
   * 5. Iterative execution:
   *    For i = startIndex → steps.length:
   *      a. Set:
   *         ctx.function_index = i
   *         ctx.function_name = steps[i].name
   *      b. Persist context
   *      c. Execute function
   *
   * 6. Exit conditions:
   *    - If function returns 'EXIT' → return immediately
   *    - If ctx.incomplete_step === true → return immediately
   *
   * 7. Completion handling:
   *    - All steps executed successfully
   *
   * 8. Logging:
   *    - Log pipeline start and completion
   *
   * 9. Cleanup:
   *    - Context pointer may be reset downstream
   *
   *
   * =========================================================
   * DESIGN PRINCIPLES
   * =========================================================
   *
   * - Deterministic execution order
   * - Resume-safe execution (state-driven)
   * - Non-blocking continuation (scheduler compatible)
   * - Separation of orchestration and logic
   * - Lightweight control layer (no business logic)
   *
   *
   * =========================================================
   * IDENTITY & SAFETY
   * =========================================================
   *
   * - Does NOT initialize execution context independently
   * - Only enriches context (pipeline_name, pointers)
   * - Safe for partial execution and resume cycles
   * - Prevents step re-execution via pointer tracking
   * - Ensures idempotent pipeline progression
   *
   * =========================================================
   * */



/* 
=========================================================
PIPELINE: Transactions
=========================================================*/
function pipeline_transactions_(){

  /* --- EXECUTION CONTEXT --- */
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_transactions_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    /* --- PIPELINE START LOG (ONLY IF FRESH RUN) --- */
    if (!ctx.is_resumed) {
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        level: 'INFO',
        action: 'PIPELINE: START',
        stepName: 'PIPELINE: START',
        details: 'Transaction pipeline execution started'
      });
    }

    /* --- STEP DEFINITIONS (ORDERED EXECUTION) --- */
    const steps = [
      backfillTxnIDs_TransactionRaw,
      cleanupInvalidTransactions_TransactionRaw
    ];

    const startIndex = ctx?.function_index || 0;

    /* --- MAIN EXECUTION LOOP (RESUME-AWARE) --- */
    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /* --- SAVE RESUME POINTER (CRITICAL) --- */
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();

      /* --- EXECUTE STEP FUNCTION --- */
      const result = fn();

      /* --- EXIT HANDLING (SCHEDULER) --- */
      if (result === 'EXIT') return;

      /* --- INCOMPLETE STEP (TIMEOUT CASE) --- */
      if (ctx?.incomplete_step) return;
    }

    /* --- PIPELINE COMPLETION METRICS --- */
    const durationMs = new Date().getTime() - t0.getTime();

    /* --- PIPELINE END LOG --- */
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* --- ERROR LOGGING --- */
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, err, 'PIPELINE');
    throw err;

  } finally {
    /* --- FINAL LOG FLUSH --- */
    flushLogs_();
  }
}


/* 
=========================================================
PIPELINE: Items
=========================================================*/
function pipeline_items_(){

  /* --- EXECUTION CONTEXT --- */
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_items_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    /* --- PIPELINE START LOG (ONLY IF FRESH RUN) --- */
    if (!ctx.is_resumed) {
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        level: 'INFO',
        action: 'PIPELINE: START',
        stepName: 'PIPELINE: START',
        details: 'Item pipeline execution started'
      });
    }

    /* --- STEP DEFINITIONS (ORDERED EXECUTION) --- */
    const steps = [
      populateStagingLookupItems_FromTransactionResolution,
      processStagingItems_StateMachine,
      promoteApprovedItems_FromStaging_ToLookup,
      backfill_ItemIDs_Machine_LookupItems,
      cleanupOrphan_ItemIDs_Machine_LookupItems
    ];

    const startIndex = ctx?.function_index || 0;

    /* --- MAIN EXECUTION LOOP (RESUME-AWARE) --- */
    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /* --- SAVE RESUME POINTER (CRITICAL) --- */
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();

      /* --- EXECUTE STEP FUNCTION --- */
      const result = fn();

      /* --- EXIT HANDLING (SCHEDULER) --- */
      if (result === 'EXIT') return;

      /* --- INCOMPLETE STEP (TIMEOUT CASE) --- */
      if (ctx?.incomplete_step) return;
    }

    /* --- PIPELINE COMPLETION METRICS --- */
    const durationMs = new Date().getTime() - t0.getTime();

    /* --- PIPELINE END LOG --- */
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* --- ERROR LOGGING --- */
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, err, 'PIPELINE');
    throw err;

  } finally {
    /* --- FINAL LOG FLUSH --- */
    flushLogs_();
  }
}


/* =========================================================
   PIPELINE: Brands
   ========================================================= */
function pipeline_brands_(){

  /* --- EXECUTION CONTEXT --- */
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_brands_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    /* --- PIPELINE START LOG (ONLY IF FRESH RUN) --- */
    if (!ctx.is_resumed) {
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        level: 'INFO',
        action: 'PIPELINE: START',
        stepName: 'PIPELINE: START',
        details: 'Brand pipeline execution started'
      });
    }

    /* --- STEP DEFINITIONS (ORDERED EXECUTION) --- */
    const steps = [
      populateStagingLookupBrands_FromTransactionResolution,
      processStagingBrands_StateMachine,
      promoteApprovedBrands_FromStaging_ToLookup,
      backfill_BrandIDs_Machine_LookupBrands,
      cleanupOrphan_BrandIDs_Machine_LookupBrands
    ];

    const startIndex = ctx?.function_index || 0;

    /* --- MAIN EXECUTION LOOP (RESUME-AWARE) --- */
    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /* --- SAVE RESUME POINTER (CRITICAL) --- */
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();

      /* --- EXECUTE STEP FUNCTION --- */
      const result = fn();

      /* --- EXIT HANDLING (SCHEDULER) --- */
      if (result === 'EXIT') return;

      /* --- INCOMPLETE STEP (TIMEOUT CASE) --- */  
      if (ctx?.incomplete_step) return;
    }

    /* --- PIPELINE COMPLETION METRICS --- */
    const durationMs = new Date().getTime() - t0.getTime();

    /* --- PIPELINE END LOG --- */
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* --- ERROR LOGGING --- */
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, err, 'PIPELINE');

    throw err;

  } finally {

    /* --- FINAL LOG FLUSH --- */
    flushLogs_();
  }
}


/* =========================================================
   PIPELINE: Products
   ========================================================= */
function pipeline_products_(){

  /* --- EXECUTION CONTEXT --- */
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_products_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    /* --- PIPELINE START LOG (ONLY IF FRESH RUN) --- */
    if (!ctx.is_resumed) {
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        level: 'INFO',
        action: 'PIPELINE: START',
        stepName: 'PIPELINE: START',
        details: 'Product pipeline execution started'
      });
    }

    /* --- STEP DEFINITIONS (ORDERED EXECUTION) --- */
    const steps = [
      populateStagingLookupProducts_FromTransactionResolution,
      processStagingProducts_StateMachine,
      promoteApprovedProducts_FromStaging_ToLookup,
      backfill_ProductIDs_Machine_LookupProducts,
      cleanupOrphan_ProductIDs_Machine_LookupProducts
    ];

    const startIndex = ctx?.function_index || 0;

    /* --- MAIN EXECUTION LOOP (RESUME-AWARE) --- */
    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /* --- SAVE RESUME POINTER (CRITICAL) --- */
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();

      /* --- EXECUTE STEP FUNCTION --- */
      const result = fn();

      /* --- EXIT HANDLING (SCHEDULER) --- */
      if (result === 'EXIT') return;

      /* --- INCOMPLETE STEP (TIMEOUT CASE) --- */
      if (ctx?.incomplete_step) return;
    }

    /* --- PIPELINE COMPLETION METRICS --- */
    const durationMs = new Date().getTime() - t0.getTime();

    /* --- PIPELINE END LOG --- */
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* --- ERROR LOGGING --- */
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, err, 'PIPELINE');

    throw err;

  } finally {

    /* --- FINAL LOG FLUSH --- */
    flushLogs_();
  }
}


/* =========================================================
   PIPELINE: Items-Brands Mapping
   ========================================================= */
function pipeline_item_brand_mapping_(){

  /* --- EXECUTION CONTEXT --- */
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_item_brand_mapping_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    /* --- PIPELINE START LOG (ONLY IF FRESH RUN) --- */
    if (!ctx.is_resumed) {
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        level: 'INFO',
        action: 'PIPELINE: START',
        stepName: 'PIPELINE: START',
        details: 'Item-Brand mapping pipeline execution started'
      });
    }

    /* --- STEP DEFINITIONS (ORDERED EXECUTION) --- */
    const steps = [
      populateMapping_Item_Brand_FromTransactionResolution,
      processMapping_Item_Brand_StateMachine,
      cleanupMapping_Item_Brand_InvalidRows
    ];

    const startIndex = ctx?.function_index || 0;

    /* --- MAIN EXECUTION LOOP (RESUME-AWARE) --- */
    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /* --- SAVE RESUME POINTER (CRITICAL) ---*/
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();

      /* --- EXECUTE STEP FUNCTION --- */
      const result = fn();

      /* --- EXIT HANDLING (SCHEDULER) --- */
      if (result === 'EXIT') return;

      /* --- INCOMPLETE STEP (TIMEOUT CASE) --- */
      if (ctx?.incomplete_step) return;
    }

    /* --- PIPELINE COMPLETION METRICS --- */
    const durationMs = new Date().getTime() - t0.getTime();

    /* --- PIPELINE END LOG --- */
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* --- ERROR LOGGING --- */
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, err, 'PIPELINE');

    throw err;

  } finally {

    /* --- FINAL LOG FLUSH --- */
    flushLogs_();
  }
}


/* =========================================================
   PIPELINE: Items-Brands-Products Mapping
   ========================================================= */
function pipeline_item_brand_product_mapping_(){

  /* --- EXECUTION CONTEXT --- */
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_item_brand_product_mapping_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    /* --- PIPELINE START LOG (ONLY IF FRESH RUN) --- */
    if (!ctx.is_resumed) {
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        level: 'INFO',
        action: 'PIPELINE: START',
        stepName: 'PIPELINE: START',
        details: 'Item-Brand-Product mapping pipeline execution started'
      });
    }

    /* --- STEP DEFINITIONS (ORDERED EXECUTION) --- */
    const steps = [
      populateMapping_Item_Brand_Product_FromTransactionResolution,
      processMapping_Item_Brand_Product_StateMachine,
      cleanupMapping_Item_Brand_Product_InvalidRows
    ];

    const startIndex = ctx?.function_index || 0;

    /* --- MAIN EXECUTION LOOP (RESUME-AWARE) --- */
    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /* --- SAVE RESUME POINTER (CRITICAL) --- */
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();

      /* --- EXECUTE STEP FUNCTION --- */
      const result = fn();

      /* --- EXIT HANDLING (SCHEDULER) --- */
      if (result === 'EXIT') return;

      /* --- INCOMPLETE STEP (TIMEOUT CASE) --- */
      if (ctx?.incomplete_step) return;
    }

    /* --- PIPELINE COMPLETION METRICS --- */
    const durationMs = new Date().getTime() - t0.getTime();

    /* --- PIPELINE END LOG --- */
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* --- ERROR LOGGING --- */
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, err, 'PIPELINE');

    throw err;

  } finally {

    /* --- FINAL LOG FLUSH --- */
    flushLogs_();
  }
}


/* =========================================================
   PIPELINE: Sheets Metadata
   ========================================================= */
function sheets_metadata_pipeline_(){

  /* --- EXECUTION CONTEXT --- */
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'sheets_metadata_pipeline_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    /* --- PIPELINE START LOG (ONLY IF FRESH RUN) --- */
    if (!ctx.is_resumed) {
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        level: 'INFO',
        action: 'PIPELINE: START',
        stepName: 'PIPELINE: START',
        details: 'Sheets metadata pipeline execution started'
      });
    }

    /* --- STEP DEFINITIONS (ORDERED EXECUTION) --- */
    const steps = [
      exportSchemaSnapshot,
      exportFormulaInventory,
      classifyColumns_fromManifest,
      generateDerivedColumnLogic
    ];

    const startIndex = ctx?.function_index || 0;

    /* --- MAIN EXECUTION LOOP (RESUME-AWARE) --- */
    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /* --- SAVE RESUME POINTER (CRITICAL) --- */
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();

      /* --- EXECUTE STEP FUNCTION --- */
      const result = fn();

      /* --- EXIT HANDLING (SCHEDULER) --- */
      if (result === 'EXIT') return;

      /* --- INCOMPLETE STEP (TIMEOUT CASE) --- */
      if (ctx?.incomplete_step) return;
    }

    /* --- PIPELINE COMPLETION METRICS --- */
    const durationMs = new Date().getTime() - t0.getTime();

    /* --- PIPELINE END LOG --- */
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* --- ERROR LOGGING --- */
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, err, 'PIPELINE');

    throw err;

  } finally {

    /* --- FINAL LOG FLUSH --- */
    flushLogs_();
  }
}

/* =========================================================
   PIPELINE: Scripts Metadata 
   ========================================================= */
function scripts_metadata_pipeline_(){

 /* --- EXECUTION CONTEXT --- */ 
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'scripts_metadata_pipeline_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    /* --- PIPELINE START LOG (ONLY IF FRESH RUN) --- */
    if (!ctx.is_resumed) {
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        level: 'INFO',
        action: 'PIPELINE: START',
        stepName: 'PIPELINE: START',
        details: 'Scripts metadata pipeline execution started'
      });
    }

    /* --- STEP DEFINITIONS (ORDERED EXECUTION) --- */
    const steps = [
      extractScriptFunctionInventory_,
      generateScriptCallMap_RAW_,
      generateScriptCallMap_INTERNAL_,
      generateScriptPipelineMap_,
      generateScriptArchitectureLogic_,
      generateScriptArchitectureDiagram_,
      extractScriptFunctionCodeSummary_,
      extractScriptDataFlowMap_,
      extractScriptPerformanceMap_,
      extractScriptSheetInteractionMap_,
      generateAIContext_,
      exportAIContextMarkdown_
    ];

    const startIndex = ctx?.function_index || 0;

    /* --- MAIN EXECUTION LOOP (RESUME-AWARE) --- */
    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /* --- SAVE RESUME POINTER (CRITICAL) --- */
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();

      /* --- EXECUTE STEP FUNCTION --- */
      const result = fn();

      /* --- EXIT HANDLING (SCHEDULER) --- */
      if (result === 'EXIT') return;

      /* --- INCOMPLETE STEP (TIMEOUT CASE) --- */
      if (ctx?.incomplete_step) return;
    }

    /* --- PIPELINE COMPLETION METRICS --- */
    const durationMs = new Date().getTime() - t0.getTime();

    /* --- PIPELINE END LOG --- */
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* --- ERROR LOGGING --- */
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, err, 'PIPELINE');

    throw err;

  } finally {

      /* --- FINAL LOG FLUSH --- */
    flushLogs_();
  }
}

/* =========================================================
   PIPELINE: Full Metadata
   ========================================================= */
function full_metadata_pipeline_(){

  /* --- EXECUTION CONTEXT --- */
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'full_metadata_pipeline_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    /* --- PIPELINE START LOG (ONLY IF FRESH RUN) --- */
    if (!ctx.is_resumed) {
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        level: 'INFO',
        action: 'PIPELINE: START',
        stepName: 'PIPELINE: START',
        details: 'Full metadata pipeline execution started'
      });
    }

    /* --- STEP DEFINITIONS (ORDERED EXECUTION) --- */
    const steps = [
      sheets_metadata_pipeline_,
      scripts_metadata_pipeline_
    ];

    const startIndex = ctx?.function_index || 0;

    /* --- MAIN EXECUTION LOOP (RESUME-AWARE) --- */
    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /* --- SAVE RESUME POINTER (CRITICAL) --- */
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();

      /* --- EXECUTE SUB-PIPELINE --- */
      const result = fn();

      /* --- EXIT HANDLING (SCHEDULER) --- */
      if (result === 'EXIT') return;

      /* --- INCOMPLETE STEP (TIMEOUT CASE) --- */
      if (ctx?.incomplete_step) return;
    }

    /* --- PIPELINE COMPLETION METRICS --- */
    const durationMs = new Date().getTime() - t0.getTime();

    /* --- PIPELINE END LOG --- */
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* --- ERROR LOGGING --- */
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, err, 'PIPELINE');

    throw err;

  } finally {

    /* --- FINAL LOG FLUSH --- */
    flushLogs_();
  }
}

/* =========================================================
   PIPELINE: Access Governance
   ========================================================= */
function pipeline_access_mode_(){

  /*
  -------------------------------------
  EXECUTION CONTEXT
  -------------------------------------*/
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_access_mode_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    /*
    -------------------------------------
    PIPELINE START LOG (ONLY IF FRESH RUN)
    -------------------------------------*/
    if (!ctx.is_resumed) {
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        level: 'INFO',
        action: 'PIPELINE: START',
        stepName: 'PIPELINE: START',
        details: 'Access governance pipeline execution started'
      });
    }

    /*
    -------------------------------------
    STEP DEFINITIONS (ORDERED EXECUTION)
    -------------------------------------*/
    const steps = [
      reconcile_access_control_metadata_,
      apply_access_governance_
    ];

    const startIndex = ctx?.function_index || 0;

    /*
    -------------------------------------
    MAIN EXECUTION LOOP (RESUME-AWARE)
    -------------------------------------*/
    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /*
      -------------------------------------
      SAVE RESUME POINTER (CRITICAL)
      -------------------------------------*/
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();

      /*
      -------------------------------------
      EXECUTE STEP FUNCTION
      -------------------------------------*/
      const result = fn();

      /*
      -------------------------------------
      EXIT HANDLING (SCHEDULER)
      -------------------------------------*/
      if (result === 'EXIT') return;

      /*
      -------------------------------------
      INCOMPLETE STEP (TIMEOUT CASE)
      -------------------------------------*/
      if (ctx?.incomplete_step) return;
    }

    /*
    -------------------------------------
    PIPELINE COMPLETION METRICS
    -------------------------------------*/
    const durationMs = new Date().getTime() - t0.getTime();

    /*
    -------------------------------------
    PIPELINE END LOG
    -------------------------------------*/
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /*
    -------------------------------------
    ERROR LOGGING
    -------------------------------------*/
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, err, 'PIPELINE');

    throw err;

  } finally {

    /*
    -------------------------------------
    FINAL LOG FLUSH (MANDATORY)
    -------------------------------------*/
    flushLogs_();
  }
}