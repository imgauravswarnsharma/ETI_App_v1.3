/**
 * =========================================================
 * SCRIPT: LOG ARCHIVER (EXECUTION-SAFE)
 * =========================================================
 *
 * PURPOSE:
 * - Archive old logs from Action + Execution sheets
 * - Preserve execution integrity
 * - Prevent sheet overflow
 *
 * DESIGN:
 * ✔ Execution-boundary aware (excludes LAST execution)
 * ✔ No mid-execution archival
 * ✔ Config-driven
 * ✔ Fail-safe per component
 * ✔ FULLY POSITION-AGNOSTIC (same as logger)
 *
 * =========================================================
 */


/*
-------------------------------------
CONFIG (PLUG & PLAY)
-------------------------------------
*/
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
-------------------------------------
*/
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
-------------------------------------
*/
function archiveLogSheet_(config){

  const logSS = getLogsSpreadsheet_();
  const sheetName = config.SHEET_NAME;

  const sh = logSS.getSheetByName(sheetName);
  if (!sh) return;

  const totalRows = sh.getLastRow();

  /*
  -------------------------------------
  CAPACITY MANAGEMENT (ALWAYS RUN)
  -------------------------------------
  */
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
  -------------------------------------
  */
  const sourceHeader = sh.getRange(1,1,1,totalCols).getValues()[0];

  const sourceMap = {};
  sourceHeader.forEach((col, idx) => {
    sourceMap[col] = idx;
  });

  /*
  -------------------------------------
  DATA READ
  -------------------------------------
  */
  const data = sh.getRange(2, 1, totalRows - 1, totalCols).getValues();
  if (!data || data.length === 0) return;

  /*
  -------------------------------------
  SAFE CUTOFF
  -------------------------------------
  */
  const cutoffIndex = findExecutionSafeCutoff_(data, sourceMap);
  if (cutoffIndex <= 0) return;

  const rowsToArchive = data.slice(0, cutoffIndex);

  /*
  -------------------------------------
  ARCHIVE INIT
  -------------------------------------
  */
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
  -------------------------------------
  */
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
  -------------------------------------
  */
  const hasData = outputRows.some(r => r.some(c => c !== '' && c !== null));

  if (!hasData) {
    throw new Error(`${sheetName}: Archive aborted (empty output)`);
  }

  /*
  -------------------------------------
  WRITE
  -------------------------------------
  */
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
  -------------------------------------
  */
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
  -------------------------------------
  */
  sh.deleteRows(2, rowsToArchive.length);

  /*
  -------------------------------------
  CLEANUP
  -------------------------------------
  */
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
-------------------------------------
*/
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