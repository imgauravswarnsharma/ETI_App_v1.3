/**
 * =========================================================
 * SCRIPT: ETI STRUCTURED LOGGER
 * =========================================================
 *
 * Layer:
 * - Core Infrastructure Layer (Logging System)
 *
 * Purpose:
 * - Provide centralized logging across entire ETI system
 * - Capture structured logs with execution context
 * - Support both debug (console) and persistent (sheet) logging
 *
 * System Role:
 * - Single source of truth for logging
 * - Consumed by:
 *   - Controller (entry logging)
 *   - Pipelines (execution grouping logs)
 *   - All scripts (event emission only)
 *
 * Core Responsibilities:
 * - Format log messages (standardized structure)
 * - Enrich logs using execution context
 * - Buffer logs in memory for performance
 * - Persist logs into Action_Logs and Execution_Logs
 *
 * Input Dependencies:
 * - Execution Context Layer:
 *   - getExecutionContext_
 *   - getOrInitExecutionContext_
 * - Automation Switch System:
 *   - getAutomationSwitchMap_
 * - Logs Spreadsheet:
 *   - getLogsSpreadsheet_
 *
 * Output Targets:
 * - Action_Logs sheet (all logs)
 * - Execution_Logs sheet (non-row logs only)
 *
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 *
 * 1. Script invokes ETI_log_(payload)
 *
 * 2. Logger performs:
 *    a. Payload normalization
 *    b. Execution context retrieval
 *    c. Switch evaluation (console / action / execution)
 *
 * 3. Log formatting:
 *    a. formatAction_
 *    b. formatStep_
 *    c. sanitizeLogText_
 *    d. buildLogComponents_
 *
 * 4. Console logging (if enabled)
 *
 * 5. Log classification:
 *    - Row-level log (has rowNumber)
 *    - Execution-level log (no rowNumber)
 *
 * 6. Buffer push:
 *    - Append structured row to ETI_LOG_BUFFER
 *
 * 7. Auto flush trigger:
 *    - If buffer size ≥ ETI_LOG_FLUSH_SIZE → flushLogs_()
 *
 * 8. flushLogs_ execution:
 *    a. Validate switches
 *    b. Initialize sheets (ensureLogSheet_)
 *    c. Insert execution boundary separator (if needed)
 *    d. Transform rows using header map
 *    e. Write Action logs (all rows)
 *    f. Write Execution logs (non-row only)
 *    g. Clear buffer
 *
 *
 * =========================================================
 * ALGORITHM (ACTUAL IMPLEMENTATION LOGIC)
 * =========================================================
 *
 * 1. Build ETI_LOG_SCHEMA → canonical column definition
 *
 * 2. Build ETI_LOG_INDEX:
 *    - Map column name → index (position agnostic access)
 *
 * 3. Logger Switch Resolution:
 *    - Build once via buildLoggerSwitchMap_
 *    - Cache in LOGGER_SWITCHES
 *
 * 4. Log Processing:
 *    - Construct header: [Script - Function - Sheet]
 *    - Normalize text (remove newlines, extra spaces)
 *    - Generate final logLine
 *
 * 5. Buffer System:
 *    - Push structured array into ETI_LOG_BUFFER
 *    - Maintain execution metadata inside row
 *
 * 6. Flush Mechanism:
 *    - Initialize sheets if missing
 *    - Ensure schema consistency (append missing columns)
 *    - Map row → sheet structure using headerMap
 *
 * 7. Execution Boundary Handling:
 *    - Compare execution_id with last_flush_log_execution_id
 *    - Insert separator row if new execution detected
 *
 * 8. Classification Logic:
 *    - Row logs → Action_Logs only
 *    - Execution logs → Action_Logs + Execution_Logs
 *
 * 9. Batch Write:
 *    - Write all rows using setValues()
 *    - Avoid per-row operations
 *
 * 10. Cleanup:
 *     - Reset ETI_LOG_BUFFER after write
 *
 *
 * =========================================================
 * DESIGN PRINCIPLES
 * =========================================================
 *
 * - Centralized logging (no distributed logging logic)
 * - Context-driven (no local state mutation)
 * - Non-blocking logging system
 * - Performance-first (batch writes)
 * - Schema-driven structure (position agnostic)
 * - Plug-and-play across all scripts
 *
 *
 * =========================================================
 * IDENTITY & SAFETY
 * =========================================================
 *
 * - Does NOT initialize execution context
 * - Only reads and enriches from context
 * - Safe for repeated invocation
 * - Buffer cleared after flush → no duplication
 *
 * =========================================================
 */


/* 
-------------------------------------
GLOBAL LOG BUFFER (BATCH WRITE)
-------------------------*/
var ETI_LOG_BUFFER = [];


/*
-------------------------
LOG FLUSH CONFIG
-------------------------*/
const ETI_LOG_FLUSH_SIZE = 60;


/*
-------------------------
SCHEMA (SINGLE SOURCE OF TRUTH)
-------------------------*/
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
-------------------------
SCHEMA INDEX MAP
-------------------------*/
const ETI_LOG_INDEX = (() => {
  const map = {};
  ETI_LOG_SCHEMA.forEach((col, idx) => {
    map[col] = idx;
  });
  return map;
})();


/*
-------------------------
LOG SHEET CONSTANTS
-------------------------*/
const ACTION_LOG_SHEET = 'Action_Logs';
const EXECUTION_LOG_SHEET = 'Execution_Logs';


/*
---------------------------------------------------------
SUB-MODULE: LOG FORMATTER (CENTRALIZED)
---------------------------------------------------------*/

/* --- Semantic Formatting --- */
function formatAction_(action){
  return (action || '').replace(/_/g, ' ');
}

function formatStep_(step){
  return (step || '')
    .replace(/_/g, ' ')
    .toUpperCase();
}

/* --- Safety Formatting --- */
function sanitizeLogText_(text){
  if (!text) return '';
  return String(text)
    .replace(/\n/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/* --- Orchestrator --- */
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
---------------------------------------------------------
SUB-MODULE: LOGGER SWITCH HANDLING
---------------------------------------------------------*/
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
-------------------------
GET LOGGER SWITCHES
-------------------------*/
function getLoggerSwitches_(){
  if (!LOGGER_SWITCHES){
    LOGGER_SWITCHES = buildLoggerSwitchMap_();
  }
  return LOGGER_SWITCHES;
}


/*
-------------------------
CORE LOGGER (ETI LOG)
-------------------------*/
function ETI_log_(payload) {
  if (!payload) return;

  if (!payload.functionName && payload.scriptName) {
    payload.functionName = payload.scriptName;
  }

  /* --- EXECUTION CONTEXT --- */
  let ctx = getOrInitExecutionContext_();
  const switches = getLoggerSwitches_();

  /* --- ACTION ENHANCEMENT (RESUME VISIBILITY) --- */
  const actionLabel = payload.action || '';

/* --- FORMATTED LOG OUTPUT --- */
  const { logLine, cleanDetails, cleanError } = buildLogComponents_(payload);

  /* --- CONSOLE MODE --- */
  if (switches.console) {
    if (payload.level === 'ERROR') console.error(logLine);
    else console.log(logLine);
  }

  /* ---  CLASSIFY LOG TYPE --- */
  const isRowLog = !!payload.rowNumber;

  if (!switches.action && !switches.execution) return;

  if (!switches.action) {
    if (isRowLog) return;
    if (!switches.execution) return;
  }

  /* --- BUFFER PUSH --- */
  ETI_LOG_BUFFER.push([
    new Date(),
    payload.level || 'INFO',

    /* --- CONTEXT PRIORITY (CRITICAL) --- */
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

    /* --- SWITCH NAME FROM CONTEXT (IF AVAILABLE) --- */
    payload.switchName || ctx?.switch_name || '',

    payload.stepName ? formatStep_(payload.stepName) : '',
    payload.rowNumber || '',

    /* --- RESUME VISIBILITY --- */
    cleanDetails
  ]);

  /* --- AUTO FLUSH --- */
  if (ETI_LOG_BUFFER.length >= ETI_LOG_FLUSH_SIZE) {
    flushLogs_();
  }
}


/*
-------------------------
ENSURE GOOGLE SHEET FOR LOGS
-------------------------*/
function ensureLogSheet_(logSS, sheetName, schema){

  let sh = logSS.getSheetByName(sheetName);

  if (!sh) {
    sh = logSS.insertSheet(sheetName);
    sh.appendRow(schema);
  }

  /* --- POSITION-AGNOSTIC HEADER SYSTEM --- */
  const existingHeader = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

  const headerMap = {};
  existingHeader.forEach((col, idx) => {
    headerMap[col] = idx + 1;
  });

  /* --- Identify missing columns --- */
  const missingColumns = schema.filter(col => !headerMap[col]);

  /* --- Append missing columns --- */
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
-------------------------
FLUSH LOGS (BATCH)
-------------------------*/
function flushLogs_(){

  if (!ETI_LOG_BUFFER || ETI_LOG_BUFFER.length === 0) return;

  const switches = getLoggerSwitches_();

  /* --- HARD EXIT IF BOTH LOGS DISABLED --- */
  if (!switches.action && !switches.execution) return;

  const logSS = getLogsSpreadsheet_();

  let sh = null;
  let execSh = null;
  let headerMap = null;
  let execHeaderMap = null;

  /* --- ENSURE SHEETS & HEADERS --- */
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

  /* ---EXECUTION BOUNDARY SEPARATION (ROW) --- */
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

  /* --- CLASSIFY LOGS INTO ACTION VS EXECUTION --- */
  const classifiedLogs = ETI_LOG_BUFFER.map(row => ({
    raw: row,
    isRowLog: !!row[ETI_LOG_INDEX['Row_Number']]
  }));

  /* --- PREPARE BATCH OUTPUT ARRAYS --- */
  const actionRows = [];
  const executionRows = [];

  classifiedLogs.forEach(entry => {

    const rawRow = entry.raw;

    const obj = {};
    ETI_LOG_SCHEMA.forEach((col, i) => obj[col] = rawRow[i]);

    /* --- ACTION LOGS (ALL ROWS) --- */
    if (switches.action && sh) {
      const actionOutput = new Array(sh.getLastColumn()).fill('');

      Object.keys(obj).forEach(col => {
        const colIndex = headerMap[col];
        if (colIndex) actionOutput[colIndex - 1] = obj[col];
      });

      actionRows.push(actionOutput);
    }


    if (!entry.isRowLog && switches.execution && execSh) {

      const execOutput = new Array(execSh.getLastColumn()).fill('');

      Object.keys(obj).forEach(col => {
        const colIndex = execHeaderMap[col];
        if (colIndex) execOutput[colIndex - 1] = obj[col];
      });

      executionRows.push(execOutput);
    }
  });

  /* --- BATCH WRITE ACTION LOGS --- */
  if (switches.action && actionRows.length > 0) {
    sh.getRange(startRow, 1, actionRows.length, sh.getLastColumn())
      .setValues(actionRows);
  }

  /* --- BATCH WRITE EXECUTION LOGS --- */
  if (switches.execution && executionRows.length > 0) {
    execSh.getRange(execStartRow, 1, executionRows.length, execSh.getLastColumn())
      .setValues(executionRows);
  }

  /* --- CLEAR BUFFER --- */
  ETI_LOG_BUFFER = [];
}



/*
===============================================
MODULE: LOGGER WRAPPERS (STANDARDIZED LOGGING)
===============================================/*

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
/* --- STEP START --- */
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

/* --- STEP NOTICE (IN-STEP UPDATES) --- */
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

/* --- STEP END --- */
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
/* --- SUMMARY (FUNCTION-LEVEL SYNOPSIS) --- */
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
/* --- SKIP NOTICE (FUNCTION-LEVEL) --- */
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
/* --- EXIT NOTICE (EXECUTION-LEVEL) --- */
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
/* --- END NOTICE (FUNCTION-LEVEL) --- */
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
/* --- ERROR NOTICE (FUNCTION-LEVEL) --- */
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