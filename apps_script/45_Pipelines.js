/**
 * =========================================================
 * PIPELINE: <pipeline_name_>
 * =========================================================
 *
 * LAYER:
 * - Execution Orchestration Layer
 *
 * PURPOSE:
 * - Execute ordered sequence of functions
 * - Maintain checkpoint-based resumable execution
 *
 * DESIGN RULES:
 * ✔ Sequential execution
 * ✔ Checkpoint before each step
 * ✔ Scheduler-compatible
 * ✔ Idempotent step execution
 *
 * =========================================================
 */


/*
-------------------------------------
EXECUTION CONTEXT
-------------------------------------

Rule:

If context already exists → DO NOT reinitialize
Only enhance (pipeline_name + run_context)

If no context → auto handled via getOrInitExecutionContext_
*/

/* =========================
   Transaction Pipeline
   ========================= */
function pipeline_transactions_(){

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_transactions_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

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

    const steps = [
      backfillTxnIDs_TransactionRaw,
      cleanupInvalidTransactions_TransactionRaw
    ];

    const startIndex = ctx?.function_index || 0;

    for (let i = startIndex; i < steps.length; i++){

      const fn = steps[i];

      /*
      -------------------------------------
      SAVE RESUME POINTER (CRITICAL)
      -------------------------------------
      */
      ctx.function_index = i;
      ctx.function_name = fn.name;
      saveExecutionContext_();   // CRITICAL FIX

      const result = fn();

      /*
      -------------------------------------
      EXIT HANDLING (SCHEDULER)
      -------------------------------------
      */
      if (result === 'EXIT') {
        return;
      }

      /*
      -------------------------------------
      INCOMPLETE STEP (TIMEOUT CASE)
      -------------------------------------
      */
      if (ctx?.incomplete_step) {
        return;
      }
    }

    const durationMs = new Date().getTime() - t0.getTime();

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE: END',
      stepName: 'PIPELINE: END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    flushLogs_();

  }
}
/* =========================
   Item Pipeline
   ========================= */
function pipeline_items_(){

/* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_items_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE START',
      details: 'Item pipeline execution started'
    });

/*-------------------------------------
  ACTUAL PIPELINE FUNCTIONS 
-------------------------------------*/
    populateStagingLookupItems_FromTransactionResolution();
    processStagingItems_StateMachine();
    promoteApprovedItems_FromStaging_ToLookup();
    backfill_ItemIDs_Machine_LookupItems();
    cleanupOrphan_ItemIDs_Machine_LookupItems();

    const durationMs = new Date().getTime() - t0.getTime();

/*-------------------------------------
  LOGGING
-------------------------------------*/
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

/*-------------------------------------
  ERROR LOGGING
-------------------------------------*/
    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    flushLogs_(); // CRITICAL: Flush buffered logs once
  
  }
}


/* =========================
   Brand Pipeline
   ========================= */
function pipeline_brands_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */
  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_brands_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE START',
      details: 'Brand pipeline execution started'
    });

    // -------------------------------------
    // ACTUAL PIPELINE FUNCTIONS
    // -------------------------------------
    populateStagingLookupBrands_FromTransactionResolution();
    processStagingBrands_StateMachine();
    promoteApprovedBrands_FromStaging_ToLookup();
    backfill_BrandIDs_Machine_LookupBrands();
    cleanupOrphan_BrandIDs_Machine_LookupBrands();

    const durationMs = new Date().getTime() - t0.getTime();

    // -------------------------------------
    // LOGGING
    // -------------------------------------
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    // -------------------------------------
    // ERROR LOGGING
    // -------------------------------------
    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    flushLogs_();

  }
}


/* =========================
   Product Pipeline
   ========================= */
function pipeline_products_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_products_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE START',
      details: 'Product pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    populateStagingLookupProducts_FromTransactionResolution();
    processStagingProducts_StateMachine();
    promoteApprovedProducts_FromStaging_ToLookup();
    backfill_ProductIDs_Machine_LookupProducts();
    cleanupOrphan_ProductIDs_Machine_LookupProducts();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}


/* =========================
   Item-Brand Mapping Pipeline
   ========================= */
function pipeline_item_brand_mapping_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_item_brand_mapping_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE START',
      details: 'Item-Brand mapping pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    populateMapping_Item_Brand_FromTransactionResolution();
    processMapping_Item_Brand_StateMachine();
    cleanupMapping_Item_Brand_InvalidRows();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}


/* =========================
   Item-Brand-Product Mapping Pipeline
   ========================= */
function pipeline_item_brand_product_mapping_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_item_brand_product_mapping_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE START',
      details: 'Item-Brand-Product mapping pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    populateMapping_Item_Brand_Product_FromTransactionResolution();
    processMapping_Item_Brand_Product_StateMachine();
    cleanupMapping_Item_Brand_Product_InvalidRows();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}


/* =========================
   Sheets Metadata Pipeline
   ========================= */
function sheets_metadata_pipeline_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'sheets_metadata_pipeline_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE START',
      details: 'Sheets metadata pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    STEP3_exportSchemaSnapshot();
    exportFormulaInventory_v2_manifest();
    classifyColumns_fromManifest();
    generateDerivedColumnLogic();
    //reconcile_access_control_metadata_();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}


/* =========================
   Scripts Metadata Pipeline
   ========================= */
function scripts_metadata_pipeline_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'scripts_metadata_pipeline_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE START',
      details: 'Scripts metadata pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    extractScriptFunctionInventory_();
    generateScriptCallMap_RAW_();
    generateScriptCallMap_INTERNAL_();
    generateScriptPipelineMap_();
    generateScriptArchitectureLogic_();
    generateScriptArchitectureDiagram_();
    extractScriptFunctionCodeSummary_();
    extractScriptDataFlowMap_();
    extractScriptPerformanceMap_();
    extractScriptSheetInteractionMap_();
    generateAIContext_();
    exportAIContextMarkdown_();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}


/* =========================
   Full Metadata Pipeline
   ========================= */
function full_metadata_pipeline_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'full_metadata_pipeline_';

  let ctx = getOrInitExecutionContext_();
  ctx.pipeline_name = FUNCTION_NAME;
  ctx.run_context = "PIPELINE";

  const t0 = new Date();

  try {

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE START',
      details: 'Full metadata pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    sheets_metadata_pipeline_();
    scripts_metadata_pipeline_();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}


/* =========================
   Access Governance Pipeline
   ========================= */
function pipeline_access_mode_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */
  let ctx = getOrInitExecutionContext_();

  ctx.pipeline_name = "pipeline_access_mode_";
  ctx.run_context = "PIPELINE";

  /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
  ------------------------------------- */
  reconcile_access_control_metadata_();
  apply_access_governance_();
}
