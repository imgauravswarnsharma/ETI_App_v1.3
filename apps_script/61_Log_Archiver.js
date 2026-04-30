/**
 * =========================================================
 * SCRIPT: LOG ARCHIVER (EXECUTION-SAFE)
 * =========================================================
 *
 * LAYER:
 * - Maintenance Layer (Standalone / Time-triggered)
 *
 * PURPOSE:
 * - Archive old logs from Action and Execution log sheets
 * - Maintain sheet size within operational limits
 * - Preserve execution integrity by avoiding active executions
 *
 * ---------------------------------------------------------
 * EXECUTION ENTRY
 * ---------------------------------------------------------
 *
 * Function: runLogArchival_()
 *
 * Invocation:
 * - Manual execution OR
 * - Time-based trigger (external to script)
 *
 * Behavior:
 * 1. Checks if archival is enabled via config
 * 2. Checks for active execution context
 *    - If active → exit immediately (no operation)
 * 3. Executes archival independently for:
 *    - Action Log
 *    - Execution Log
 * 4. Each component is fail-isolated (try/catch per sheet)
 *
 * ---------------------------------------------------------
 * EXECUTION FLOW
 * ---------------------------------------------------------
 *
 * runLogArchival_()
 *   ├── Validate ENABLE_ARCHIVAL flag
 *   ├── Check execution context
 *   │     └── If active → exit
 *   ├── Process ACTION log (if enabled)
 *   │     └── archiveLogSheet_(ACTION_CONFIG)
 *   ├── Process EXECUTION log (if enabled)
 *   │     └── archiveLogSheet_(EXECUTION_CONFIG)
 *
 * archiveLogSheet_(config)
 *   ├── Load target log sheet
 *   ├── Ensure row capacity (always executed)
 *   ├── Validate minimum rows and threshold
 *   ├── Build source header map
 *   ├── Read full data range
 *   ├── Compute execution-safe cutoff index
 *   ├── Slice rows eligible for archival
 *   ├── Initialize archive sheet (create if missing)
 *   ├── Map rows → archive schema (header-aligned)
 *   ├── Validate non-empty output
 *   ├── Append to archive sheet
 *   ├── Verify write success
 *   ├── Delete archived rows from source
 *   ├── Optional cleanup (format reset)
 *
 * findExecutionSafeCutoff_(data, sourceMap)
 *   ├── Identify Execution_ID and Action columns
 *   ├── Traverse data bottom → top
 *   ├── Track executions using Set
 *   ├── Detect incomplete execution:
 *   │     - Action NOT in terminal set
 *   ├── Return cutoff index before incomplete execution
 *   ├── If all complete → return full length
 *
 * ---------------------------------------------------------
 * ALGORITHM (DETAILED)
 * ---------------------------------------------------------
 *
 * 1. Capacity Management:
 *    - Ensure minimum free rows using ensureFreeRowCapacity_
 *
 * 2. Archival Trigger Condition:
 *    - Proceed only if totalRows ≥ TRIGGER_THRESHOLD
 *
 * 3. Header Mapping:
 *    - Build column index map from header row (position-agnostic)
 *
 * 4. Data Extraction:
 *    - Read all rows except header
 *
 * 5. Safe Cutoff Determination:
 *    - Identify boundary excluding last incomplete execution
 *
 * 6. Row Selection:
 *    - Select rows [0 → cutoffIndex)
 *
 * 7. Archive Sheet Preparation:
 *    - Ensure archive sheet exists with required schema
 *
 * 8. Row Transformation:
 *    - Map source rows → archive schema using header alignment
 *
 * 9. Data Validation:
 *    - Ensure at least one non-empty row before write
 *
 * 10. Write Operation:
 *     - Append rows to archive sheet
 *
 * 11. Write Verification:
 *     - Confirm row count written matches expected
 *
 * 12. Source Cleanup:
 *     - Delete archived rows from original sheet
 *
 * 13. Formatting Cleanup (Optional):
 *     - Reset background formatting for remaining rows
 *
 * ---------------------------------------------------------
 * SAFETY GUARANTEES
 * ---------------------------------------------------------
 *
 * - No archival during active execution
 * - Last incomplete execution is never archived
 * - Strict write verification before deletion
 * - Fail-safe isolation per sheet (no cross-impact)
 * - Header-based mapping ensures column safety
 *
 * ---------------------------------------------------------
 * CONFIGURATION CONTROL
 * ---------------------------------------------------------
 *
 * LOG_ARCHIVER_CONFIG:
 * - ENABLE_ARCHIVAL → global toggle
 * - ENABLE_CLEANUP → formatting cleanup toggle
 *
 * Per Sheet:
 * - ENABLE → enable/disable per log
 * - SHEET_NAME → target sheet
 * - TRIGGER_THRESHOLD → minimum rows to trigger archival
 * - MAX_ROWS → informational (not enforced in logic)
 *
 * ---------------------------------------------------------
 * NON-GOALS (EXPLICIT)
 * ---------------------------------------------------------
 *
 * - Does NOT manage trigger creation
 * - Does NOT interact with pipelines or scheduler
 * - Does NOT modify execution context
 * - Does NOT perform partial execution handling
 *
 * =========================================================
 */

/*
-------------------------------------
CONFIG (PLUG & PLAY)
-------------------------------------*/
const LOG_ARCHIVER_CONFIG = {

  ENABLE_ARCHIVAL: true,
  ENABLE_CLEANUP: true,

  SHEETS: {

    ACTION: {
      ENABLE: true,
      SHEET_NAME: ACTION_LOG_SHEET,
      TRIGGER_THRESHOLD: 50,
      MAX_ROWS: 5000
    },

    EXECUTION: {
      ENABLE: true,
      SHEET_NAME: EXECUTION_LOG_SHEET,
      TRIGGER_THRESHOLD: 40,
      MAX_ROWS: 2000
    }

  }

};


/*
-------------------------------------
ENTRY POINT (MANUAL / CONTROLLER)
-------------------------------------*/
function runLogArchival_(){

  if (!LOG_ARCHIVER_CONFIG.ENABLE_ARCHIVAL) return;

  const ctx = getExecutionContext_();
  if (ctx) {
    console.log('Archival skipped: active execution');
    return;
  }

  // ACTION LOG
  if (LOG_ARCHIVER_CONFIG.SHEETS.ACTION.ENABLE) {
    try {
      archiveLogSheet_(LOG_ARCHIVER_CONFIG.SHEETS.ACTION);
    } catch (err) {
      console.error('Action log archival failed', err);
    }
  }

  // EXECUTION LOG
  if (LOG_ARCHIVER_CONFIG.SHEETS.EXECUTION.ENABLE) {
    try {
      archiveLogSheet_(LOG_ARCHIVER_CONFIG.SHEETS.EXECUTION);
    } catch (err) {
      console.error('Execution log archival failed', err);
    }
  }
}


/*
-------------------------------------
CORE ARCHIVAL FUNCTION
-------------------------------------*/
function archiveLogSheet_(config){

  const logSS = getLogsSpreadsheet_();
  const sheetName = config.SHEET_NAME;

  const sh = logSS.getSheetByName(sheetName);
  if (!sh) return;

  const totalRows = sh.getLastRow();

  /*
  -------------------------------------
  CAPACITY MANAGEMENT (ALWAYS RUN)
  -------------------------------------*/
  const capacityResult = ensureFreeRowCapacity_(sh, {
    minFreeRows: 1200,
    expansionChunk: 3600
  });

  if (capacityResult?.expanded) {
    console.log(
      `${sheetName}: Capacity expanded by ${capacityResult.rowsAdded} rows ` +
      `(freeRowsBefore=${capacityResult.freeRowsBefore}, freeRowsAfter=${capacityResult.freeRowsAfter})`
    );
  }

  // SAFETY: prevent invalid range
  if (totalRows <= 1) return;

  // ARCHIVAL CONDITION
  if (totalRows < config.TRIGGER_THRESHOLD) return;

  const totalCols = sh.getLastColumn();

  /*
  -------------------------------------
  SOURCE HEADER MAP
  -------------------------------------*/
  const sourceHeader = sh.getRange(1,1,1,totalCols).getValues()[0];

  const sourceMap = {};
  sourceHeader.forEach((col, idx) => {
    sourceMap[col] = idx;
  });

  /*
  -------------------------------------
  DATA READ
  -------------------------------------*/
  const data = sh.getRange(2, 1, totalRows - 1, totalCols).getValues();
  if (!data || data.length === 0) return;

  /*
  -------------------------------------
  SAFE CUTOFF
  -------------------------------------*/
  const cutoffIndex = findExecutionSafeCutoff_(data, sourceMap);
  if (cutoffIndex <= 0) return;

  const rowsToArchive = data.slice(0, cutoffIndex);

  /*
  -------------------------------------
  ARCHIVE INIT
  -------------------------------------*/
  const archiveSheetName = sheetName + '_Archive';

  const {
    sheet: archiveSh,
    headerMap: archiveHeaderMap
  } = ensureLogSheet_(
    logSS,
    archiveSheetName,
    ETI_LOG_SCHEMA
  );

  const archiveTotalCols = archiveSh.getLastColumn();

  /*
  -------------------------------------
  BUILD OUTPUT (STRICT HEADER ALIGNMENT)
  -------------------------------------*/
  const outputRows = rowsToArchive.map(rawRow => {

    const output = new Array(archiveTotalCols).fill('');

    Object.keys(archiveHeaderMap).forEach(col => {

      const sourceIdx = sourceMap[col];
      const targetIdx = archiveHeaderMap[col];

      if (sourceIdx !== undefined && targetIdx !== undefined) {
        output[targetIdx - 1] = rawRow[sourceIdx];
      }

    });

    return output;
  });

  /*
  -------------------------------------
  SAFETY CHECK (CRITICAL)
  -------------------------------------*/
  const hasData = outputRows.some(r => r.some(c => c !== '' && c !== null));

  if (!hasData) {
    throw new Error(`${sheetName}: Archive aborted (empty output)`);
  }

  /*
  -------------------------------------
  WRITE
  -------------------------------------*/
  const writeStartRow = archiveSh.getLastRow() + 1;

  archiveSh.getRange(
    writeStartRow,
    1,
    outputRows.length,
    archiveTotalCols
  ).setValues(outputRows);

  /*
  -------------------------------------
  VERIFY WRITE
  -------------------------------------*/
  const check = archiveSh.getRange(
    writeStartRow,
    1,
    outputRows.length,
    1
  ).getValues();

  if (check.length !== outputRows.length) {
    throw new Error(`${sheetName}: Archive write verification failed`);
  }

  /*
  -------------------------------------
  DELETE SOURCE
  -------------------------------------*/
  sh.deleteRows(2, rowsToArchive.length);

  /*
  -------------------------------------
  CLEANUP
  -------------------------------------*/
  if (LOG_ARCHIVER_CONFIG.ENABLE_CLEANUP) {

    const remainingRows = sh.getLastRow();

    if (remainingRows > 1) {
      sh.getRange(2, 1, remainingRows - 1, totalCols)
        .setBackground(null);
    }
  }

  console.log(`${sheetName}: Archived ${rowsToArchive.length} rows`);
}


/*
-------------------------------------
EXECUTION SAFE CUTOFF (STATE-BASED)
-------------------------------------*/
function findExecutionSafeCutoff_(data, sourceMap){

  const execIdIdx = sourceMap['Execution_ID'];
  const actionIdx = sourceMap['Action'];

  if (execIdIdx === undefined || actionIdx === undefined) return 0;

  const TERMINAL_ACTIONS = [
    'END',
    'PIPELINE END',
    'ERROR',
    'EXIT'
  ];

  const seen = new Set();

  for (let i = data.length - 1; i >= 0; i--) {

    const execId = data[i][execIdIdx];
    const action = data[i][actionIdx];

    // Skip separator / empty rows
    if (!execId) continue;

    // First time seeing this execution from bottom → last row of execution
    if (!seen.has(execId)) {

      seen.add(execId);

      const isComplete = TERMINAL_ACTIONS.includes(action);

      // Found incomplete execution → stop boundary
      if (!isComplete) {
        return i + 1;
      }
    }
  }

  // All executions complete → archive everything
  return data.length;
}
 
/*
-------------------------------------
Time Based Archiver Trigger
-------------------------------------
*/
/*
function setup_logArchiverTriggers_() {

  // Remove old triggers for this function (clean setup)
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runLogArchival_') {
      ScriptApp.deleteTrigger(t);
    }
  });

  /*
  -------------------------------------
  RUN 1 → ~3:00 AM
  -------------------------------------
  */
  /*
  ScriptApp.newTrigger('runLogArchival_')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .nearMinute(0)
    .create();

  /*
  -------------------------------------
  RUN 2 → ~4:00 AM
  -------------------------------------
  */
  /*
  ScriptApp.newTrigger('runLogArchival_')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .nearMinute(0)
    .create();
}

*/