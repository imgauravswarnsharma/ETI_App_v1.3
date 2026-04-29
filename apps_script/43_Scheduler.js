/*
-------------------------------------
SCHEDULER CONFIG
-------------------------------------
*/
const ETI_SCHEDULER_CONFIG = {
  TIME_LIMIT_MS: 1 * 5000 * 1000,     // 4 min (hard limit)
  BUFFER_MS: 2000,                  // Safety buffer
  TRIGGER_DELAY_MS: 2000,           // Restart delay
  MAX_RESUME: 3                     // Resume Limit
};


/*
-------------------------------------
RUNTIME FLAG: CONTINUATION TRACKING
-------------------------------------
*/
var __ETI_CONTINUATION_SCHEDULED__ = false;



/*
-------------------------------------
CHECK IF SHOULD EXIT
-------------------------------------
*/
function shouldExitForTimeout_(startTime){

  const now = new Date().getTime();
  const elapsed = now - startTime;

  return elapsed >= (ETI_SCHEDULER_CONFIG.TIME_LIMIT_MS - ETI_SCHEDULER_CONFIG.BUFFER_MS);
}


/*
-------------------------------------
CLEAR EXISTING SCHEDULER TRIGGERS
-------------------------------------
*/
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
-------------------------------------
*/
function scheduleContinuation_(){

  const props = PropertiesService.getScriptProperties();

  /*
  -------------------------------------
  SAFE ACTIVE GUARD (NO HARD BLOCK)
  -------------------------------------
  */
  const active = props.getProperty('ETI_SCHEDULER_ACTIVE');
  const lastTs = props.getProperty('ETI_SCHEDULER_TS');
  const now = Date.now();

  if (active === 'TRUE' && lastTs){
    const diff = now - Number(lastTs);

    // Prevent rapid duplicate triggers (<5 sec)
    if (diff < 1500){
      return;
    }
  }

  /*
  -------------------------------------
  SET ACTIVE + TIMESTAMP
  -------------------------------------
  */
  props.setProperty('ETI_SCHEDULER_ACTIVE', 'TRUE');
  props.setProperty('ETI_SCHEDULER_TS', String(now));

  /*
  -------------------------------------
  CLEAN OLD TRIGGERS (CRITICAL)
  -------------------------------------
  */
  clearExistingSchedulerTriggers_();

  /*
  -------------------------------------
  MARK CONTINUATION SCHEDULED
  -------------------------------------
  */
  __ETI_CONTINUATION_SCHEDULED__ = true;

  ScriptApp.newTrigger('ETI_schedulerResume_')
    .timeBased()
    .after(ETI_SCHEDULER_CONFIG.TRIGGER_DELAY_MS)
    .create();
}


/*
-------------------------------------
SAVE STATE FOR RESUME
-------------------------------------
*/
function saveExecutionState_(meta = {}){

  const ctx = getExecutionContext_();
  if (!ctx) return;

  /*
  -------------------------------------
  RESUME METADATA (SCHEDULER STATE ONLY)
  -------------------------------------
  */
  if (meta.functionName){
    ctx.function_name = meta.functionName;
  }

  if (meta.pipelineName){
    ctx.pipeline_name = meta.pipelineName;
  }

  /*
  -------------------------------------
  PERSIST CONTEXT (CRITICAL)
  -------------------------------------
  */
  saveExecutionContext_();
}


/*
-------------------------------------
EXIT + SCHEDULE CONTINUATION
-------------------------------------
*/
function exitAndScheduleContinuation_(scriptName, functionName, meta = {}){

  const ctx = getExecutionContext_();
  
  /*
  -------------------------------------
  MARK STEP INCOMPLETE
  -------------------------------------
  */
  if (!ctx) return 'EXIT';

  ctx.incomplete_step = true;

  /*
  -------------------------------------
  SAVE STATE FIRST
  -------------------------------------
  */
  saveExecutionState_({
    functionName,
    ...meta
  });

  /*
  -------------------------------------
  CURRENT RESUME COUNT (NO INCREMENT HERE)
  -------------------------------------
  */
  const currentCount = ctx.resume_count || 0;
  const MAX_RESUME = ETI_SCHEDULER_CONFIG.MAX_RESUME;

  /*
  -------------------------------------
  RETRY EXHAUSTED → FINALIZE EXIT
  -------------------------------------
  */
  if (currentCount >= MAX_RESUME){

    ETI_log_({
      scriptName,
      functionName,
      level: 'ERROR',
      action: 'EXIT',
      details: `Max retry limit reached (${ctx.resume_count})`
    });

    flushLogs_();

    /*
    -------------------------------------
    CRITICAL: MARK CONTINUATION TO PREVENT
    SUCCESS OVERWRITE IN RESUME FLOW
    -------------------------------------
    */
    __ETI_CONTINUATION_SCHEDULED__ = true;

    const switchName = ctx?.switch_name;

    if (switchName){
      finalizeExecutionFromScheduler_('EXIT', 'Stopped after retry limit');
    }

    return 'EXIT';
  }


  /*
  -------------------------------------
  SCHEDULER LOG (TIMEOUT)
  -------------------------------------
  */
  ETI_log_({
    scriptName,
    functionName,
    level: 'WARN',
    action: 'SCHEDULER',
    stepName: 'PAUSE (TIMEOUT)',
    details: `Execution paused due to time limit (scheduler)`
  });

  flushLogs_();

  /*
  -------------------------------------
  SCHEDULE CONTINUATION
  -------------------------------------
  */
  scheduleContinuation_();

  return 'EXIT';
}


/*
-------------------------------------
SCHEDULER RESUME ENTRY
-------------------------------------
*/
function ETI_schedulerResume_(){

  const props = PropertiesService.getScriptProperties();

  try {

    /*
    -------------------------------------
    RESET CONTINUATION FLAG
    -------------------------------------
    */
    __ETI_CONTINUATION_SCHEDULED__ = false;
    props.deleteProperty('ETI_SCHEDULER_ACTIVE');

    /*
    -------------------------------------
    RESTORE CONTEXT (CRITICAL)
    -------------------------------------
    */
    const restored = restoreExecutionContext_();
    if (!restored) {
      console.error('RESUME FAILED: No execution context found');
      return;
    }

    const ctx = getExecutionContext_();
    if (!ctx) return;

    /*
    -------------------------------------
    INCREMENT RESUME COUNT
    -------------------------------------
    */
    ctx.resume_count = (ctx.resume_count || 0) + 1;
    ctx.is_resumed = true;
    saveExecutionContext_();


    /*
    -------------------------------------
    RESET PREVIOUS INCOMPLETE STATE
    -------------------------------------
    */
    if (ctx?.incomplete_step === true) {
      ctx.incomplete_step = false;
      saveExecutionContext_();
    }

    /*
    -------------------------------------
    LOG RESUME
    -------------------------------------
    */
    ETI_log_({
      scriptName: 'Scheduler',
      functionName: 'ETI_schedulerResume_',
      level: 'INFO',
      action: 'SCHEDULER',
      stepName: 'RESUME',
      details: `Resuming execution (#${ctx.resume_count})`
    });

    flushLogs_();


    /*
    -------------------------------------
    RESUME POINTERS
    -------------------------------------
    */
    const resumeFn = ctx?.function_name;
    const resumePipeline = ctx?.pipeline_name;
    const switchName = ctx?.switch_name;

    /*
    -------------------------------------
    RESUME PIPELINE
    -------------------------------------
    */
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

    /*
    -------------------------------------
    RESUME FUNCTION
    -------------------------------------
    */
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

    /*
    -------------------------------------
    FALLBACK (CRITICAL)
    -------------------------------------
    */
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