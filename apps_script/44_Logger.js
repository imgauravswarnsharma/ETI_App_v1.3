/**
 * =========================================================
 * SYSTEM CONTEXT: ETI STRUCTURED LOGGER
 * =========================================================
 *
 * POSITION IN ARCHITECTURE:
 * ---------------------------------------------------------
 * Logger = Core Infrastructure Layer
 *
 * Used by:
 * - Controller (entry point)
 * - Pipelines (execution grouping)
 * - All scripts (business logic)
 *
 * It is the SINGLE SOURCE OF TRUTH for:
 * - Debug logging
 * - Action logging (persistent audit)
 * - Execution context propagation
 *
 *
 * =========================================================
 * 1. CORE COMPONENTS
 * =========================================================
 *
 * 1. EXECUTION CONTEXT (GLOBAL STATE)
 * ---------------------------------------------------------
 * Defined via:
 *   initExecutionContext_()
 *   getExecutionContext_()
 *
 * Structure:
 * {
 *   execution_id   → unique UUID per run
 *   pipeline_name  → set at pipeline level
 *   run_context    → STANDALONE / PIPELINE
 *   trigger_type   → MANUAL / CONTROLLER
 *   started_at     → timestamp
 * }
 *
 * Behavior:
 * - Initialized ONLY at execution entry points
 *   (Controller OR manual pipeline/script)
 *
 * - Pipeline layer may MODIFY context (not recreate)
 *
 * - Logger READS context (never mutates it)
 *
 *
 * =========================================================
 * 2. LOGGING FLOW (END-TO-END)
 * =========================================================
 *
 * ENTRY POINTS:
 * ---------------------------------------------------------
 * A. Controller Execution
 *    → initExecutionContext_({ trigger_type: CONTROLLER })
 *
 * B. Manual Execution (Apps Script UI)
 *    → initExecutionContext_() OR implicit defaults
 *
 * C. Pipeline Execution
 *    → Enhances context:
 *       pipeline_name
 *       run_context = PIPELINE
 *
 *
 * LOGGING EXECUTION:
 * ---------------------------------------------------------
 * Step 1: Script calls ETI_log_(payload)
 *
 * Step 2: Logger builds:
 *   - Debug line (console)
 *   - Structured row (buffer)
 *
 * Step 3: Buffer accumulates logs
 *
 * Step 4: flushLogs_() writes to sheet (batch)
 *
 *
 * =========================================================
 * 3. LOG TYPES
 * =========================================================
 *
 * A. DEBUG LOG (REAL-TIME)
 * ---------------------------------------------------------
 * Output:
 *   Apps Script console
 *
 * Format:
 *   [Script - Function - Sheet] ACTION ⇒ Details
 *
 * Example:
 *   [Items - populate... - Staging_Lookup_Items] SUMMARY ⇒ Scanned=1999 | ...
 *
 * Purpose:
 * - Fast visual debugging
 * - Execution trace
 *
 *
 * B. ACTION LOG (PERSISTENT)
 * ---------------------------------------------------------
 * Output:
 *   Action_Logs sheet
 *
 * Stored as structured rows
 *
 * Includes:
 * - Execution metadata
 * - Debug message snapshot
 * - Error details
 *
 *
 * =========================================================
 * 4. ACTION LOG SCHEMA
 * =========================================================
 *
 * Columns:
 *
 * Timestamp
 * Level
 *
 * Trigger_Type
 * Run_Context
 *
 * Action
 * Debug_Message
 * Error_Message
 *
 * Execution_ID
 *
 * Pipeline_Name
 * Script_Name
 * Function_Name
 * Sheet_Name
 * Switch_Name
 *
 * Step_Name
 * Row_Number
 *
 * Details
 *
 *
 * Design Intent:
 * ---------------------------------------------------------
 * - Debug_Message = primary scan column
 * - Other columns = structured filtering / audit
 *
 *
 * =========================================================
 * 5. LOG FORMATTING SYSTEM
 * =========================================================
 *
 * Centralized via:
 *   buildLogComponents_()
 *
 * Sub-components:
 *
 * 1. formatAction_
 *    → Converts ACTION_NAME → "ACTION NAME"
 *
 * 2. formatStep_
 *    → Converts STEP_NAME → "STEP NAME"
 *
 * 3. sanitizeLogText_
 *    → Removes:
 *       - new lines
 *       - extra spaces
 *       - unsafe characters
 *
 * 4. Header Builder:
 *    → [Script - Function - Sheet]
 *
 *
 * RULE:
 * ---------------------------------------------------------
 * Logger controls:
 *   ✔ Structure
 *   ✔ Safety
 *
 * Script controls:
 *   ✔ Meaning
 *   ✔ Metrics
 *
 *
 * =========================================================
 * 6. BUFFERED WRITE SYSTEM
 * =========================================================
 *
 * Mechanism:
 * ---------------------------------------------------------
 * - Logs are NOT written immediately
 * - Stored in ETI_LOG_BUFFER
 * - Written in batch via flushLogs_()
 *
 * Benefits:
 * ---------------------------------------------------------
 * ✔ Performance optimized
 * ✔ Reduces API calls
 * ✔ Prevents partial writes
 *
 *
 * CRITICAL RULE:
 * ---------------------------------------------------------
 * flushLogs_() MUST be called:
 * - At pipeline end
 * - At controller wrapper end
 * - In finally blocks
 *
 *
 * =========================================================
 * 7. EXECUTION CONTEXT PROPAGATION
 * =========================================================
 *
 * FLOW:
 *
 * Controller
 *   ↓
 * initExecutionContext_
 *   ↓
 * Pipeline (enhances context)
 *   ↓
 * Script functions (read-only)
 *   ↓
 * Logger uses context
 *
 *
 * RULES:
 * ---------------------------------------------------------
 * ✔ Only ENTRY POINT initializes context
 * ✔ Lower layers MUST NOT reinitialize
 * ✔ Context can only be ENRICHED downstream
 *
 *
 * =========================================================
 * 8. ERROR HANDLING
 * =========================================================
 *
 * WRAPPER:
 *   ETI_logError_()
 *
 * Captures:
 * - error.message → Details
 * - error.stack   → Error_Message
 *
 *
 * =========================================================
 * 9. STEP LOGGING (STANDARDIZED)
 * =========================================================
 *
 * Utilities:
 * - ETI_logStepStart_
 * - ETI_logStepEnd_
 *
 * Output:
 *   PROCESS ⇒ LOAD STAGING started
 *   PROCESS ⇒ LOAD STAGING completed
 *
 *
 * =========================================================
 * 10. DESIGN PRINCIPLES
 * =========================================================
 *
 * ✔ Single Source of Truth (Logger only)
 * ✔ No logging logic in business scripts
 * ✔ Context-driven logging
 * ✔ Plug-and-play across system
 * ✔ Performance-first (batch writes)
 * ✔ Readable debug-first design
 *
 *
 * =========================================================
 * 11. CONTROLLER + PIPELINE + LOGGER INTERPLAY
 * =========================================================
 *
 * Controller:
 *   → Defines trigger_type
 *   → Starts execution
 *
 * Pipeline:
 *   → Defines run_context
 *   → Defines pipeline_name
 *
 * Script:
 *   → Emits logs (no context logic)
 *
 * Logger:
 *   → Formats + persists logs
 *
 *
 * =========================================================
 * 12. FINAL ARCHITECTURE SUMMARY
 * =========================================================
 *
 * Controller  → WHEN to run
 * Pipeline    → WHAT group to run
 * Script      → HOW logic runs
 * Logger      → WHAT happened (audit + debug)
 *
 * =========================================================
 */


/* 
-------------------------------------
  GLOBAL LOG BUFFER (BATCH WRITE)
-------------------------------------*/
var ETI_LOG_BUFFER = [];


/*
-------------------------------------
LOG FLUSH CONFIG (ADDED)
-------------------------------------*/
const ETI_LOG_FLUSH_SIZE = 60;


/*
-------------------------------------
SCHEMA (SINGLE SOURCE OF TRUTH)
-------------------------------------*/
const ETI_LOG_SCHEMA = [
  'Timestamp',
  'Level',

  'Trigger_Type',
  'Run_Context',

  'Action',
  'Debug_Message',
  'Error_Message',

  'Execution_ID',

  'Pipeline_Name',
  'Script_Name',
  'Function_Name',
  'Sheet_Name',
  'Switch_Name',

  'Step_Name',
  'Row_Number',

  'Details'
];


/*
-------------------------------------
SCHEMA INDEX MAP (POSITION AGNOSTIC)
-------------------------------------*/
const ETI_LOG_INDEX = (() => {
  const map = {};
  ETI_LOG_SCHEMA.forEach((col, idx) => {
    map[col] = idx;
  });
  return map;
})();


/*
-------------------------------------
LOG SHEET CONSTANTS
-------------------------------------*/
const ACTION_LOG_SHEET = 'Action_Logs';
const EXECUTION_LOG_SHEET = 'Execution_Logs';


/*
-------------------------------------
LOG FORMATTER (CENTRALIZED)
-------------------------------------*/

/* -------- Semantic Formatting -------- */
function formatAction_(action){
  return (action || '').replace(/_/g, ' ');
}

function formatStep_(step){
  return (step || '')
    .replace(/_/g, ' ')
    .toUpperCase();
}

/* -------- Safety Formatting -------- */
function sanitizeLogText_(text){
  if (!text) return '';
  return String(text)
    .replace(/\n/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/* -------- Orchestrator -------- */
function buildLogComponents_(payload){

  const actionDisplay = formatAction_(payload.action);

  const cleanDetails = sanitizeLogText_(payload.details);
  const cleanError = sanitizeLogText_(payload.errorMessage);

  const header = `[${payload.scriptName || ''} - ${payload.functionName || ''}${payload.sheetName ? ' - ' + payload.sheetName : ''}]`;

  const logLine = [
    header,
    actionDisplay,
    '⇒',
    payload.rowNumber ? `ROW ${payload.rowNumber}` : '',
    cleanDetails
  ].join(' ').replace(/\s+/g, ' ').trim();

  return {
    logLine,
    cleanDetails,
    cleanError
  };
}


/*
------------------------------------- 
  LOGGER SWITCH BUILDER (EXPLICIT)
-------------------------------------*/
let LOGGER_SWITCHES = null;

function buildLoggerSwitchMap_(){
  const switches = getAutomationSwitchMap_();

  return {
    action:    switches["Enable_Action_Log"]    === true,
    execution: switches["Enable_Execution_Log"] === true,
    console:   switches["Enable_Console_Log"]   === true
  };
}


/*
------------------------------------
Logger's  Switches Map
------------------------------------*/
function getLoggerSwitches_(){
  if (!LOGGER_SWITCHES){
    LOGGER_SWITCHES = buildLoggerSwitchMap_();
  }
  return LOGGER_SWITCHES;
}



/*
-------------------------------------
CORE LOGGER (ETI LOG)
-------------------------------------*/
function ETI_log_(payload) {
  if (!payload) return;

  if (!payload.functionName && payload.scriptName) {
    payload.functionName = payload.scriptName;
  }

  /*
  -------------------------------------
  EXECUTION CONTEXT FALLBACK
  -------------------------------------*/
  let ctx = getOrInitExecutionContext_();
  const switches = getLoggerSwitches_();

  /*
  -------------------------------------
  ACTION ENHANCEMENT (RESUME VISIBILITY)
  -------------------------------------*/
  const actionLabel = payload.action || '';

  /*
  -------------------------------------
  FORMATTED LOG OUTPUT
  -------------------------------------*/
  const { logLine, cleanDetails, cleanError } = buildLogComponents_(payload);

  /*
  -------------------------------------
  CONSOLE MODE
  -------------------------------------*/
  if (switches.console) {
    if (payload.level === 'ERROR') console.error(logLine);
    else console.log(logLine);
  }

  /*
  -------------------------------------
  CLASSIFY LOG TYPE
  -------------------------------------*/
  const isRowLog = !!payload.rowNumber;

  // Drop only if both disabled
  if (!switches.action && !switches.execution) return;

  // If action is OFF → restrict logs
  if (!switches.action) {
    if (isRowLog) return;           // skip row logs
    if (!switches.execution) return;
  }


  /*
  -------------------------------------
  ACTION LOG (BUFFERED)
  -------------------------------------*/
  ETI_LOG_BUFFER.push([
    new Date(),
    payload.level || 'INFO',

    /*
    -------------------------------------
    FIX 1: CONTEXT PRIORITY (CRITICAL)
    -------------------------------------*/
    ctx?.trigger_type || payload.triggerType || 'MANUAL',
    ctx?.run_context || 'STANDALONE',

    actionLabel,
    logLine,
    cleanError,

    ctx?.execution_id || '',

    ctx?.pipeline_name || '',
    payload.scriptName || '',
    payload.functionName || '',
    payload.sheetName || '',

    /*
    -------------------------------------
    FIX 2: SWITCH NAME FROM CONTEXT
    -------------------------------------*/
    payload.switchName || ctx?.switch_name || '',

    payload.stepName ? formatStep_(payload.stepName) : '',
    payload.rowNumber || '',

    /*
    -------------------------------------
    FIX 3: RESUME VISIBILITY
    -------------------------------------*/
    cleanDetails
  ]);

  /*
  -------------------------------------
  AUTO FLUSH
  -------------------------------------*/
  if (ETI_LOG_BUFFER.length >= ETI_LOG_FLUSH_SIZE) {
    flushLogs_();
  }
}


/*
-------------------------------------
LOG SHEET INITIALIZER (CENTRALIZED)
-------------------------------------*/
function ensureLogSheet_(logSS, sheetName, schema){

  let sh = logSS.getSheetByName(sheetName);

  // Ensure sheet + base schema
  if (!sh) {
    sh = logSS.insertSheet(sheetName);
    sh.appendRow(schema);
  }


  /*
  -------------------------------
  POSITION-AGNOSTIC HEADER SYSTEM
  -------------------------------*/
  const existingHeader = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

  const headerMap = {};
  existingHeader.forEach((col, idx) => {
    headerMap[col] = idx + 1;
  });

  // Identify missing columns
  const missingColumns = schema.filter(col => !headerMap[col]);

  // Append missing columns
  if (missingColumns.length > 0) {
    sh.getRange(1, existingHeader.length + 1, 1, missingColumns.length)
      .setValues([missingColumns]);

    missingColumns.forEach((col, i) => {
      headerMap[col] = existingHeader.length + i + 1;
    });
  }

  return {
    sheet: sh,
    headerMap
  };
}


/*
-------------------------------------
FLUSH LOGS (BATCH WRITE)
-------------------------------------*/
function flushLogs_(){

  if (!ETI_LOG_BUFFER || ETI_LOG_BUFFER.length === 0) return;

  const switches = getLoggerSwitches_();

  /*
  -------------------------------------
  HARD EXIT (NO LOGGING ENABLED)
  -------------------------------------*/
  if (!switches.action && !switches.execution) return;

  const logSS = getLogsSpreadsheet_();

  let sh = null;
  let execSh = null;
  let headerMap = null;
  let execHeaderMap = null;

  /*
  -------------------------------------
  CONDITIONAL SHEET INIT
  -------------------------------------*/
  if (switches.action) {
    const res = ensureLogSheet_(logSS, ACTION_LOG_SHEET, ETI_LOG_SCHEMA);
    sh = res.sheet;
    headerMap = res.headerMap;
  }

  if (switches.execution) {
    const res = ensureLogSheet_(logSS, EXECUTION_LOG_SHEET, ETI_LOG_SCHEMA);
    execSh = res.sheet;
    execHeaderMap = res.headerMap;
  }

  let startRow = sh ? sh.getLastRow() + 1 : 0;
  let execStartRow = execSh ? execSh.getLastRow() + 1 : 0;

  const ctx = getExecutionContext_();
  const executionId = ctx?.execution_id;

  /*
  -------------------------------------
  EXECUTION-BOUNDARY SEPARATOR (SYNCED)
  -------------------------------------*/
  if (ctx) {

    if (!ctx.last_flush_log_execution_id) {
      ctx.last_flush_log_execution_id = null;
    }

    if (executionId && ctx.last_flush_log_execution_id !== executionId) {

      if (switches.action && sh && sh.getLastRow() > 1) {
        sh.insertRowBefore(startRow);
        sh.getRange(startRow, 1, 1, sh.getLastColumn()).setBackground('#fbbc04');
        startRow++;
      }

      if (switches.execution && execSh && execSh.getLastRow() > 1) {
        execSh.insertRowBefore(execStartRow);
        execSh.getRange(execStartRow, 1, 1, execSh.getLastColumn()).setBackground('#fbbc04');
        execStartRow++;
      }

      ctx.last_flush_log_execution_id = executionId;
      saveExecutionContext_();
    }
  }

  /*
  -------------------------------------
  CLASSIFICATION (ROW vs EXECUTION)
  -------------------------------------*/
  const classifiedLogs = ETI_LOG_BUFFER.map(row => ({
    raw: row,
    isRowLog: !!row[ETI_LOG_INDEX['Row_Number']]
  }));


/*
-------------------------------------
POSITION-AWARE WRITE (DUAL LOG)
-------------------------------------*/
  const actionRows = [];
  const executionRows = [];

  classifiedLogs.forEach(entry => {

    const rawRow = entry.raw;

    const obj = {};
    ETI_LOG_SCHEMA.forEach((col, i) => obj[col] = rawRow[i]);

    /*
    -------------------------------------
    BUILD ACTION ROW
    -------------------------------------*/
    if (switches.action && sh) {
      const actionOutput = new Array(sh.getLastColumn()).fill('');

      Object.keys(obj).forEach(col => {
        const colIndex = headerMap[col];
        if (colIndex) {
          actionOutput[colIndex - 1] = obj[col];
        }
      });

      actionRows.push(actionOutput);
    }

    /*
    -------------------------------------
    BUILD EXECUTION ROW (HEADER MAP)
    -------------------------------------*/
    if (!entry.isRowLog && switches.execution && execSh) {

      const execOutput = new Array(execSh.getLastColumn()).fill('');

      Object.keys(obj).forEach(col => {
        const colIndex = execHeaderMap[col];
        if (colIndex) {
          execOutput[colIndex - 1] = obj[col];
        }
      });

      executionRows.push(execOutput);
    }
  });

  /*
  -------------------------------------
  ACTION LOG WRITE
  -------------------------------------*/
  if (switches.action && actionRows.length > 0) {
    sh.getRange(startRow, 1, actionRows.length, sh.getLastColumn())
      .setValues(actionRows);
  }

  /*
  -------------------------------------
  EXECUTION LOG WRITE
  -------------------------------------*/
  if (switches.execution && executionRows.length > 0) {
    execSh.getRange(execStartRow, 1, executionRows.length, execSh.getLastColumn())
      .setValues(executionRows);
  }


/*
-------------------------------------
CLEAR BUFFER
-------------------------------------*/
ETI_LOG_BUFFER = [];
}


/*
-------------------------------------
WRAPPER: START LOGGER
-------------------------------------*/
function ETI_logStart_(scriptName, functionName, sheetName){
  ETI_log_({
    scriptName,
    functionName,
    sheetName,
    level: 'INFO',
    action: 'START',
    stepName: 'START',
    details: 'Execution started'
  });
}


/*
-------------------------------------
WRAPPER: STEP LOGGER
-------------------------------------*/
function ETI_logStepStart_(scriptName, functionName, sheetName, stepName=''){
  ETI_log_({
    scriptName,
    functionName,
    sheetName,
    level: 'INFO',
    action: 'PROCESS',
    stepName: stepName ? `${stepName} (START)` : '',
    details: stepName ? `${formatStep_(stepName)} started` : 'Step started'
  });
}


function ETI_logNotice_(scriptName, functionName, sheetName, stepName, details=''){
  ETI_log_({
    scriptName,
    functionName,
    sheetName,
    level: 'INFO',
    action: 'NOTICE',
    stepName,
    details
  });
}


function ETI_logStepEnd_(scriptName, functionName, sheetName, stepName='', details=''){
  ETI_log_({
    scriptName,
    functionName,
    sheetName,
    level: 'INFO',
    action: 'PROCESS',
    stepName: stepName ? `${stepName} (END)` : '',
    details: details || (stepName ? `${formatStep_(stepName)} completed` : 'Step completed')
  });
}


/*
-------------------------------------
WRAPPER: SUMMARY LOGGER
-------------------------------------*/
function ETI_logSummary_(scriptName, functionName, sheetName, details){
  ETI_log_({
    scriptName,
    functionName,
    sheetName,
    level: 'INFO',
    action: 'SUMMARY',
    stepName: 'SUMMARY',
    details
  });
}


/*
-------------------------------------
WRAPPER: SKIPPED LOGGER
-------------------------------------*/
function ETI_logSkip_(scriptName, functionName, sheetName, details){
  ETI_log_({
    scriptName,
    functionName,
    sheetName,
    level: 'INFO',
    action: 'SKIP',
    stepName: 'SKIP',
    details
  });
}


/*
-------------------------------------
WRAPPER: EXIT LOGGER
-------------------------------------*/
function ETI_logExit_(scriptName, functionName, sheetName, details){
  ETI_log_({
    scriptName,
    functionName,
    sheetName,
    level: 'ERROR',
    action: 'EXIT',
    stepName: 'EXIT',
    details
  });
}


/*
-------------------------------------
WRAPPER: END LOGGER
-------------------------------------*/
function ETI_logEnd_(scriptName, functionName, sheetName){
  ETI_log_({
    scriptName,
    functionName,
    sheetName,
    level: 'INFO',
    action: 'END',
    stepName:'END',
    details: 'Execution completed'
  });
}


/*
-------------------------------------
WRAPPER: ERROR LOGGER
-------------------------------------*/
function ETI_logError_(scriptName, functionName, sheetName, error, stepName='ERROR'){
  ETI_log_({
    scriptName,
    functionName,
    sheetName,
    level: 'ERROR',
    action: 'ERROR',
    stepName,
    details: error?.message || '',
    errorMessage: error?.stack || ''
  });
}