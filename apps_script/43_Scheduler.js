/**
 * =========================================================
 * SCRIPT: SCHEDULER & CONTINUATION ENGINE
 * =========================================================
 *
 * Layer:
 * - Execution Control Layer (Scheduler)
 *
 * Purpose:
 * - Handle time-bound execution limits in Apps Script
 * - Enable safe continuation of long-running processes
 * - Manage controlled exit and resume using triggers
 *
 * System Role:
 * - Works between Controller and Execution Context
 * - Ensures pipelines and functions resume safely after timeout
 * - Prevents execution loss due to Apps Script runtime limits
 *
 * Core Responsibilities:
 * - Detect timeout conditions during execution
 * - Persist execution state before exit
 * - Schedule continuation using time-based triggers
 * - Restore execution context during resume
 * - Resume pipeline or function execution deterministically
 *
 * Input Dependencies:
 * - Execution Context Layer:
 *   - getExecutionContext_
 *   - saveExecutionContext_
 *   - restoreExecutionContext_
 * - Controller Layer:
 *   - finalizeExecutionFromScheduler_
 *   - automationController_onChange
 * - Logger Layer:
 *   - ETI_log_
 *   - flushLogs_
 * - Apps Script Services:
 *   - ScriptApp (trigger management)
 *   - PropertiesService (runtime flags)
 *
 * Output Targets:
 * - Trigger Queue (time-based continuation triggers)
 * - Execution Context (updated state for resume)
 *
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 *
 * 1. Execution starts via Controller or Pipeline
 *
 * 2. During execution:
 *    - shouldExitForTimeout_(startTime) is evaluated
 *
 * 3. If timeout threshold reached:
 *    a. exitAndScheduleContinuation_() is invoked
 *    b. Mark current step as incomplete (ctx.incomplete_step = true)
 *    c. Persist execution state (function / pipeline pointer)
 *    d. Log scheduler pause event
 *    e. Schedule continuation trigger (scheduleContinuation_)
 *    f. Exit execution safely
 *
 * 4. Trigger fires → ETI_schedulerResume_()
 *
 * 5. Resume flow:
 *    a. Reset scheduler runtime flags
 *    b. Restore execution context
 *    c. Increment resume_count
 *    d. Reset incomplete_step flag
 *    e. Log resume event
 *
 * 6. Execution continuation:
 *    - If pipeline_name exists → resume pipeline
 *    - Else if function_name exists → resume function
 *    - Else → fallback to controller
 *
 * 7. Post execution:
 *    - If still incomplete → exit again (loop)
 *    - Else:
 *      a. Clear execution pointers
 *      b. Finalize execution via controller
 *
 *
 * =========================================================
 * ALGORITHM (ACTUAL IMPLEMENTATION LOGIC)
 * =========================================================
 *
 * 1. Timeout Detection:
 *    - Compare elapsed time vs TIME_LIMIT_MS - BUFFER_MS
 *
 * 2. Exit Strategy:
 *    - Mark ctx.incomplete_step = true
 *    - Persist execution metadata:
 *      • function_name
 *      • pipeline_name
 *
 * 3. Retry Control:
 *    - Maintain ctx.resume_count
 *    - Enforce MAX_RESUME limit
 *    - If exceeded:
 *      → log error
 *      → finalize execution (EXIT)
 *
 * 4. Trigger Scheduling:
 *    - Use PropertiesService for active trigger guard
 *    - Prevent duplicate trigger creation
 *    - Clear existing triggers before scheduling
 *
 * 5. Continuation Flag:
 *    - __ETI_CONTINUATION_SCHEDULED__ ensures:
 *      → no false success overwrite in resume flow
 *
 * 6. Resume Engine:
 *    - Restore execution context
 *    - Increment resume count
 *    - Reset incomplete_step flag
 *
 * 7. Dynamic Execution Routing:
 *    - globalThis[pipeline_name]() → pipeline resume
 *    - globalThis[function_name]() → function resume
 *
 * 8. Execution Completion:
 *    - If no incomplete step:
 *      → clear pointers
 *      → finalize via controller
 *
 * 9. Fallback Handling:
 *    - If no function/pipeline found:
 *      → fallback to controller trigger
 *
 *
 * =========================================================
 * DESIGN PRINCIPLES
 * =========================================================
 *
 * - Time-bound execution safety (prevent hard termination)
 * - Deterministic continuation (resume-safe design)
 * - Context-driven execution (no local state dependency)
 * - Controlled retry mechanism (bounded resume attempts)
 * - Non-blocking scheduling (trigger-based continuation)
 * - Separation of concerns (scheduler does not execute logic)
 *
 *
 * =========================================================
 * IDENTITY & SAFETY
 * =========================================================
 *
 * - Does NOT execute business logic directly
 * - Only controls execution flow and continuation
 * - Safe for repeated resume cycles
 * - Prevents duplicate trigger execution via guard
 * - Execution context is single source of truth
 *
 * =========================================================
 */




/* --- SCHEDULER CONFIG --- */

const ETI_SCHEDULER_CONFIG = {
  TIME_LIMIT_MS: 1 * 295 * 1000,   // 5 min (max hard limit)
  BUFFER_MS: 2000,              // Safety buffer
  TRIGGER_DELAY_MS: 2000,       // Restart delay
  MAX_RESUME: 5                // Resume Limit
};


/*
=========================================================
MODULE: RUNTIME FLAGS
=========================================================*/
/*
----------------------------
/* --- CONTINUATION TRACKING FLAG --- */
var __ETI_CONTINUATION_SCHEDULED__ = false;


/*
=========================================================
MODULE: TIMEOUT MANAGEMENT
=========================================================*/
/*
-------------------------------------
CHECK IF SHOULD EXIT
-------------------------------------*/
function shouldExitForTimeout_(startTime){

  const now = new Date().getTime();
  const elapsed = now - startTime;

  return elapsed >= (ETI_SCHEDULER_CONFIG.TIME_LIMIT_MS - ETI_SCHEDULER_CONFIG.BUFFER_MS);
}


/*
=========================================================
MODULE: TRIGGER MANAGEMENT
=========================================================*/
/*
-------------------------------------
CLEAR SCHEDULER TRIGGERS
-------------------------------------*/
function clearExistingSchedulerTriggers_(){

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'ETI_schedulerResume_') {
      ScriptApp.deleteTrigger(t);
    }
  });
}


/*
-------------------------------------
SCHEDULE CONTINUATION TRIGGER
-------------------------------------*/
function scheduleContinuation_(){

  const props = PropertiesService.getScriptProperties();

  /* --- SAFE ACTIVE GUARD (NO HARD BLOCK) --- */
  const active = props.getProperty('ETI_SCHEDULER_ACTIVE');
  const lastTs = props.getProperty('ETI_SCHEDULER_TS');
  const now = Date.now();

  if (active === 'TRUE' && lastTs){
    const diff = now - Number(lastTs);

    
    if (diff < 500){   // Prevent rapid duplicate triggers. 500ms currently, can adjust as needed
      return;
    }
  }

  /* --- SET ACTIVE + TIMESTAMP --- */
  props.setProperty('ETI_SCHEDULER_ACTIVE', 'TRUE');
  props.setProperty('ETI_SCHEDULER_TS', String(now));

  /* --- CLEAN OLD TRIGGERS --- */
  clearExistingSchedulerTriggers_();

  /* --- MARK CONTINUATION --- */
  __ETI_CONTINUATION_SCHEDULED__ = true;

  ScriptApp.newTrigger('ETI_schedulerResume_')
    .timeBased()
    .after(ETI_SCHEDULER_CONFIG.TRIGGER_DELAY_MS)
    .create();
}


/*
=========================================================
MODULE: STATE PERSISTENCE
=========================================================*/
/*
-------------------------------------
SAVE EXECUTION STATE
-------------------------------------*/
function saveExecutionState_(meta = {}){

  const ctx = getExecutionContext_();
  if (!ctx) return;

  /* --- RESUME METADATA --- */
  if (meta.functionName){
    ctx.function_name = meta.functionName;
  }

  if (meta.pipelineName){
    ctx.pipeline_name = meta.pipelineName;
  }

  /* --- PERSIST CONTEXT --- */
  saveExecutionContext_();
}


/*
=========================================================
MODULE: EXIT & CONTINUATION HANDLING
=========================================================*/
/*
-------------------------------------
EXIT + SCHEDULE CONTINUATION
-------------------------------------*/
function exitAndScheduleContinuation_(scriptName, functionName, meta = {}){

  const ctx = getExecutionContext_();

  /* --- MARK STEP INCOMPLETE --- */
  if (!ctx) return 'EXIT';

  ctx.incomplete_step = true;

  /* --- SAVE STATE FIRST --- */
  saveExecutionState_({
    functionName,
    ...meta
  });

  const currentCount = ctx.resume_count || 0;
  const MAX_RESUME = ETI_SCHEDULER_CONFIG.MAX_RESUME;

  /* --- RETRY EXHAUSTED → FINALIZE EXIT --- */
  if (currentCount >= MAX_RESUME){

    ETI_log_({
      scriptName,
      functionName,
      level: 'ERROR',
      action: 'EXIT',
      details: `Max retry limit reached (${ctx.resume_count})`
    });

    flushLogs_();

    /* --- CRITICAL: MARK CONTINUATION TO PREVENT SUCCESS OVERWRITE IN RESUME FLOW --- */
    __ETI_CONTINUATION_SCHEDULED__ = true;

    const switchName = ctx?.switch_name;

    if (switchName){
      finalizeExecutionFromScheduler_('EXIT', 'Stopped after retry limit');
    }

    return 'EXIT';
  }

  /* --- SCHEDULER PAUSE LOG (TIMEOUT)--- */
  ETI_log_({
    scriptName,
    functionName,
    level: 'WARN',
    action: 'SCHEDULER',
    stepName: 'PAUSE (TIMEOUT)',
    details: `Execution paused due to time limit (scheduler)`
  });

  flushLogs_();

  /* --- SCHEDULE CONTINUATION --- */
  scheduleContinuation_();

  return 'EXIT';
}


/*
=========================================================
MODULE: RESUME ENGINE
=========================================================*/
/*
-------------------------
SCHEDULER RESUME ENTRY
-------------------------*/
function ETI_schedulerResume_(){

  const props = PropertiesService.getScriptProperties();

  try {

    /* --- RESET CONTINUATION FLAG --- */
    __ETI_CONTINUATION_SCHEDULED__ = false;
    props.deleteProperty('ETI_SCHEDULER_ACTIVE');

    /* --- RESTORE CONTEXT --- */
    const restored = restoreExecutionContext_();
    if (!restored) {
      console.error('RESUME FAILED: No execution context found');
      return;
    }

    const ctx = getExecutionContext_();
    if (!ctx) return;

    /* --- INCREMENT RESUME COUNT --- */
    ctx.resume_count = (ctx.resume_count || 0) + 1;
    ctx.is_resumed = true;
    saveExecutionContext_();

    /* --- RESET INCOMPLETE STATE --- */
    if (ctx?.incomplete_step === true) {
      ctx.incomplete_step = false;
      saveExecutionContext_();
    }

    /* --- LOG RESUME --- */
    ETI_log_({
      scriptName: 'Scheduler',
      functionName: 'ETI_schedulerResume_',
      level: 'INFO',
      action: 'SCHEDULER',
      stepName: 'RESUME',
      details: `Resuming execution (#${ctx.resume_count})`
    });

    flushLogs_();


    /* --- RESUME POINTERS --- */
    const resumeFn = ctx?.function_name;
    const resumePipeline = ctx?.pipeline_name;
    const switchName = ctx?.switch_name;


    /* --- RESUME PIPELINE --- */
    if (resumePipeline && typeof globalThis[resumePipeline] === 'function') {

      setLogMessage_(switchName, "Execution resumed");

      globalThis[resumePipeline]();

      restoreExecutionContext_();

      const ctxAfter = getExecutionContext_();

      if (ctxAfter?.incomplete_step === true){
        return;
      }

      ctxAfter.function_name = null;
      ctxAfter.pipeline_name = null;
      ctxAfter.incomplete_step = false;

      saveExecutionContext_();

      finalizeExecutionFromScheduler_('SUCCESS','Completed successfully');
      return;
    }


    /* --- RESUME FUNCTION --- */
    if (resumeFn && typeof globalThis[resumeFn] === 'function') {

      setLogMessage_(switchName, "Execution resumed");

      globalThis[resumeFn]();

      restoreExecutionContext_();

      const ctxAfter = getExecutionContext_();

      if (ctxAfter?.incomplete_step === true){
        return;
      }

      ctxAfter.function_name = null;
      ctxAfter.pipeline_name = null;
      ctxAfter.incomplete_step = false;

      saveExecutionContext_();

      finalizeExecutionFromScheduler_('SUCCESS','Completed successfully');
      return;
    }

    
    /* --- FALLBACK --- */
    if (switchName){
      automationController_onChange();
      return;
    }

  } catch (err) {

    console.error('Scheduler Resume Error:', err);

    const ctx = getExecutionContext_();
    const switchName = ctx?.switch_name;

    if (switchName){
      finalizeExecutionFromScheduler_('FAILED', err.message);
    }
  }
}