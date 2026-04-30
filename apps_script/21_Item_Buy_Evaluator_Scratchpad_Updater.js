/**
 * =========================================================
 * SCRIPT: ITEM BUY EVALUATOR (SCRATCHPAD PROCESSOR)
 * =========================================================
 *
 * LAYER:
 * - Business Logic Layer (Event-driven, single-row evaluator)
 *
 * TRIGGER:
 * - Installable onChange trigger
 * - Executes ONLY when:
 *   ✔ changeType = EDIT
 *   ✔ Active sheet = Item_Buy_Evaluate
 *   ✔ Edited row = 2
 *
 * PURPOSE:
 * - Evaluate a single scratchpad row
 * - Compare against Item_Evaluation_Log
 * - Perform deterministic mutation:
 *     → INSERT new record (new variant)
 *     → UPDATE existing record (same variant)
 *     → SKIP (no change)
 *
 * ---------------------------------------------------------
 * EXECUTION PRECONDITIONS
 * ---------------------------------------------------------
 *
 * Execution proceeds ONLY if ALL conditions are TRUE:
 *
 * 1. Input_Ready_For_Comparison = TRUE
 * 2. Compare_ID = 1
 * 3. Evaluation_ID is present
 * 4. Summary_UI is NOT empty
 *
 * If any condition fails:
 * → Execution exits silently (no logging)
 *
 * ---------------------------------------------------------
 * EXECUTION FLOW (STEP-BY-STEP)
 * ---------------------------------------------------------
 *
 * 1. Trigger Filtering
 *    - Validate event type, sheet, and row
 *
 * 2. Gating Check
 *    - Validate input readiness + evaluation identifiers
 *
 * 3. Lock Acquisition
 *    - Acquire ScriptLock (wait up to 5 seconds)
 *
 * 4. Data Stabilization
 *    - flush → sleep → flush (ensure formula resolution)
 *
 * 5. Row Read
 *    - Read evaluator row (rowIndex = 2)
 *    - Build header map for column access
 *
 * 6. Snapshot Construction
 *    - Build snapshot from evaluator row
 *    - Map:
 *        Planned_* → Evaluated_*
 *        Current_* → Evaluated fields
 *
 * 7. Load Existing Log Data
 *    - Read Item_Evaluation_Log (all rows)
 *
 * 8. Match Detection
 *    - Find existing row using variant key:
 *        (Evaluated_Item
 *         Evaluated_Brand
 *         Evaluated_Product
 *         Evaluated_Platform
 *         Evaluated_Qty
 *         Evaluated_Qty_Unit)
 *
 * 9. Decision Logic
 *
 *    CASE A: Match Found
 *
 *        - Compare:
 *            existingPrice vs newPrice
 *            existingDate vs newDate
 *
 *        A1. No Change
 *            → Log SKIP
 *            → End execution
 *
 *        A2. Only Date Changed
 *            → Preserve existing price
 *            → UPDATE row (date only)
 *
 *        A3. Price Changed
 *            → UPDATE row (price + date)
 *
 *    CASE B: No Match Found
 *        → INSERT new row
 *        → Copy format from row 2 (anchor row)
 *
 * 10. Optional Reset
 *     - Clears evaluator row if PERSIST_MODE = false
 *
 * 11. Logging
 *     - START → PROCESS → END
 *     - Only meaningful actions logged:
 *         ✔ INSERT
 *         ✔ UPDATE
 *         ✔ SKIP (post-match only)
 *
 * 12. Finalization
 *     - Release lock
 *     - Flush log buffer
 *
 * ---------------------------------------------------------
 * DECISION ALGORITHM
 * ---------------------------------------------------------
 *
 * IF match exists:
 *
 *     IF price unchanged AND date unchanged:
 *         → SKIP
 *
 *     ELSE IF price unchanged AND date changed:
 *         → UPDATE (date only)
 *
 *     ELSE:
 *         → UPDATE (price + date)
 *
 * ELSE:
 *     → INSERT
 *
 * ---------------------------------------------------------
 * DATA INTERACTION
 * ---------------------------------------------------------
 *
 * READ:
 * - Item_Buy_Evaluate (row 2 only)
 * - Item_Evaluation_Log (full read)
 *
 * WRITE:
 * - Item_Evaluation_Log
 *     → UPDATE existing row
 *     → INSERT new row
 *
 * OPTIONAL:
 * - Reset evaluator row (if persist disabled)
 *
 * ---------------------------------------------------------
 * LOGGING BEHAVIOR
 * ---------------------------------------------------------
 *
 * Logged:
 * ✔ START (execution begins)
 * ✔ PROCESS (INSERT / UPDATE)
 * ✔ SKIP (match found but no change)
 * ✔ END (execution completes)
 *
 * NOT Logged:
 * ✘ Trigger events
 * ✘ Gating failures
 * ✘ Non-execution paths
 *
 * Logging uses:
 * - ETI_log_ (PROCESS)
 * - ETI_logSkip_
 * - ETI_logStart_
 * - ETI_logEnd_
 *
 * ---------------------------------------------------------
 * CONCURRENCY CONTROL
 * ---------------------------------------------------------
 *
 * - Uses ScriptLock
 * - waitLock(5000 ms)
 *
 * Purpose:
 * - Prevent concurrent updates to Item_Evaluation_Log
 * - Avoid duplicate insert/update race conditions
 *
 * ---------------------------------------------------------
 * DESIGN CONSTRAINTS
 * ---------------------------------------------------------
 *
 * 1. Variant Identity Definition:
 *    - Based on:
 *        Item + Brand + Product + Platform + Qty + Qty Unit
 *    - Price is NOT part of identity
 *
 * 2. Evaluation Log Behavior:
 *    - Stores latest state per variant
 *    - NOT a historical log
 *
 * 3. Single-Row Execution:
 *    - Only processes rowIndex = 2
 *
 * 4. Deterministic Execution:
 *    - No partial updates
 *    - Full row overwrite on update
 *
 * 5. Logging Discipline:
 *    - No trigger noise
 *    - Only mutation-level logging
 *
 * =========================================================
 */


/* --- GLOBAL CONFIGS --- */
const EVALUATOR_PERSIST_MODE = true // IF TRUE THEN EVALUATOR WILL NOT OUTPUTS CLEAR POST EXECUTION. HELPFUL FOR DEBUG.
const SCRIPT_NAME  = 'Evaluator';
const FUNCTION_NAME = 'processEvaluationRow_';
const SHEET_NAME   = 'Item_Buy_Evaluate';

/* 
=========================================================
MODULE: ON CHANGE AUTO-TRIGGER
=========================================================*/
function buy_Evaluator_onChange(e) {

  // ---- FILTER 1: EVENT TYPE ----
  if (!e || e.changeType !== 'EDIT') return;

  const sheet = e.source.getActiveSheet();
  if (!sheet) return;

  // ---- FILTER 2: SHEET ----
  if (sheet.getName() !== "Item_Buy_Evaluate") return;

  const range = e.source.getActiveRange();
  if (!range) return;

  // ---- FILTER 3: ROW ----
  if (range.getRow() !== 2) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName("Item_Buy_Evaluate");
  if (!evalSheet) return;

  const data = evalSheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headerMap = getHeaderMap_(evalSheet);
  const row = data[1];

  const evaluationId = getCell_(row, headerMap, "Evaluation_ID");
  const inputReady = getCell_(row, headerMap, "Input_Ready_For_Comparison");

  if (inputReady === true && evaluationId) {
    processEvaluationRow_(evalSheet, 2);
  }
}
/*
=========================================================
MODULE: CORE EVALUATOR LOGIC
=========================================================*/
/*
-------------------------------------
PROCESS EVALUATION ROW
-------------------------------------*/
function processEvaluationRow_(evalSheet, rowIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    /* --- START --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

    const headerMap = getHeaderMap_(evalSheet);

    SpreadsheetApp.flush();
    Utilities.sleep(250);
    SpreadsheetApp.flush();

    const row = evalSheet
      .getRange(rowIndex, 1, 1, evalSheet.getLastColumn())
      .getValues()[0];

    const get = (col) => getCell_(row, headerMap, col);

    if (get("Input_Ready_For_Comparison") !== true) return;
    if (Number(get("Compare_ID")) !== 1) return;
    if (!get("Evaluation_ID")) return;

    if (!get("Summary_UI") || get("Summary_UI").toString().trim() === "") {
      return;
    }

    const logSheet = SpreadsheetApp.getActive().getSheetByName("Item_Evaluation_Log");
    if (!logSheet) return;

    const logHeaderMap = getHeaderMap_(logSheet);

    const lastRow = logSheet.getLastRow();
    const logData =
      lastRow > 1
        ? logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).getValues()
        : [];

    const snapshot = buildSnapshot_(headerMap, row);
    const matchIndex = findMatchRow_(snapshot, logHeaderMap, logData);

    if (matchIndex !== -1) {

      const existingRow = logData[matchIndex];

      const existingPrice = Number(getCell_(existingRow, logHeaderMap, "Evaluated_Price"));
      const newPrice = Number(snapshot["Evaluated_Price"]);

      const existingDate = getCell_(existingRow, logHeaderMap, "Evaluation_Date");
      const newDate = snapshot["Evaluation_Date"];

      const priceChanged = existingPrice !== newPrice;
      const dateChanged =
        new Date(existingDate).getTime() !== new Date(newDate).getTime();
      
      /* --- STEP: UPDATE EXISTING EVALUATION LOG RECORD --- */

      // CASE A: Nothing changed (price and date are the same) —> SKIP
      if (!priceChanged && !dateChanged) {

        ETI_logSkip_(
          SCRIPT_NAME,
          FUNCTION_NAME,
          SHEET_NAME,
          'Match found but no change in price/date'
        );

        ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

        return;
      }

      // CASE B: Only date changed (price is the same) —> UPDATE (date only, preserve price)
      if (!priceChanged && dateChanged) {
    snapshot["Evaluated_Price"] = existingPrice; // preserve
      }
      // CASE C: Price changed —> UPDATE (date will naturally update along with price)

      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: 'Item_Evaluation_Log',
        level: 'INFO',
        action: 'PROCESS',
        rowNumber: matchIndex + 2,
        details: `UPDATE | PriceChanged=${priceChanged} | DateChanged=${dateChanged}`
      });

      logSheet
        .getRange(matchIndex + 2, 1, 1, logSheet.getLastColumn())
        .setValues([buildLogRow_(snapshot, logHeaderMap)]);

    } else {

      /* --- STEP: INSERT NEW RECORD INTO EVALUATION LOG] --- */
      const newRowValues = buildLogRow_(snapshot, logHeaderMap);
      const insertRowIndex = logSheet.getLastRow() + 1;

      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: 'Item_Evaluation_Log',
        level: 'INFO',
        action: 'PROCESS',
        rowNumber: insertRowIndex,
        details: `INSERT | Item=${snapshot["Evaluated_Item"]} | Platform=${snapshot["Evaluated_Platform"]} | Qty=${snapshot["Evaluated_Qty"]} ${snapshot["Evaluated_Qty_Unit"]}`
      });

      logSheet
        .getRange(insertRowIndex, 1, 1, logSheet.getLastColumn())
        .setValues([newRowValues]);

      /* ===== FORMAT COPY FROM ROW 2 ANCHOR ===== */
      if (logSheet.getLastRow() >= 2) {
        logSheet
          .getRange(2, 1, 1, logSheet.getLastColumn())
          .copyTo(
            logSheet.getRange(insertRowIndex, 1, 1, logSheet.getLastColumn()),
            { formatOnly: true }
          );
      }
    }

    if (!EVALUATOR_PERSIST_MODE) {
      resetEvaluatorRow_(evalSheet, rowIndex, headerMap);
    }
    
    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

  } catch (err) {

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SHEET_NAME,
      err,
      'MAIN'
    );

    throw err;

  } finally {
    lock.releaseLock();
    flushLogs_();
  }
}

/*
=========================================================
MODULE: UTILITY FUNCTIONS
=========================================================*/
/*
-------------------------------------
   SNAPSHOT BUILDER
-------------------------------------*/
function buildSnapshot_(headerMap, row) {

  const snapshot = {};

  for (const key in headerMap) {
    snapshot[key] = row[headerMap[key]];
  }

  snapshot["Evaluated_Item"] = snapshot["Planned_Item"];
  snapshot["Evaluated_Brand"] = snapshot["Planned_Brand"];
  snapshot["Evaluated_Product"] = snapshot["Planned_Product"];
  snapshot["Evaluated_Platform"] = snapshot["Current_Platform"];
  snapshot["Evaluated_Qty"] = snapshot["Planned_Qty"];
  snapshot["Evaluated_Qty_Unit"] = snapshot["Planned_Qty_Unit"];
  snapshot["Recorded_Normalised_Qty"] = snapshot["Planned_Normalised_Qty"];
  snapshot["Evaluated_Price"] = snapshot["Current_Price"];

  snapshot["Logged_At"] = snapshot["Evaluated_At"];
  snapshot["Eval_Ready_For_Logging"] = true;

  return snapshot;
}


/*
-------------------------------------
FIND MATCH ROW
-------------------------------------*/
function findMatchRow_(snapshot, logHeaderMap, logData) {

  const norm = (v) => (v || "").toString().trim().toLowerCase();
  const num  = (v) => Number(v);

  for (let i = 0; i < logData.length; i++) {
    const row = logData[i];

    const match =
      norm(getCell_(row, logHeaderMap, "Evaluated_Item")) === norm(snapshot["Evaluated_Item"]) &&
      norm(getCell_(row, logHeaderMap, "Evaluated_Brand")) === norm(snapshot["Evaluated_Brand"]) &&
      norm(getCell_(row, logHeaderMap, "Evaluated_Product")) === norm(snapshot["Evaluated_Product"]) &&
      norm(getCell_(row, logHeaderMap, "Evaluated_Platform")) === norm(snapshot["Evaluated_Platform"]) &&
      num(getCell_(row, logHeaderMap, "Evaluated_Qty")) === num(snapshot["Evaluated_Qty"]) &&
      norm(getCell_(row, logHeaderMap, "Evaluated_Qty_Unit")) === norm(snapshot["Evaluated_Qty_Unit"]);

    if (match) return i;
  }

  return -1;
}

/*
-------------------------------------
   BUILD LOG ROW
-------------------------------------*/
function buildLogRow_(snapshot, logHeaderMap) {

  const row = new Array(Object.keys(logHeaderMap).length).fill("");

  for (const logCol in logHeaderMap) {

    const snapshotKey = Object.keys(snapshot).find(
      k => k.toLowerCase().trim() === logCol.toLowerCase().trim()
    );

    if (snapshotKey !== undefined) {
      row[logHeaderMap[logCol]] = snapshot[snapshotKey];
    }
  }

  return row;
}


/*
-------------------------------------
   RESET EVALUATOR
-------------------------------------*/
function resetEvaluatorRow_(sheet, rowIndex, headerMap) {

  const clearCols = [
    "Planned_Item",
    "Planned_Brand",
    "Planned_Product",
    "Current_Platform",
    "Planned_Qty",
    "Planned_Qty_Unit",
    "Current_Price",
    "Evaluation_Date",
    "Evaluated_At"
  ];

  clearCols.forEach(col => {
    const idx = headerMap[col];
    if (idx !== undefined) {
      sheet.getRange(rowIndex, idx + 1).clearContent();
    }
  });

  const evalIdx = headerMap["Evaluation_ID"];
  if (evalIdx !== undefined) {
    sheet.getRange(rowIndex, evalIdx + 1).clearContent();
  }
}


/*
-------------------------------------
   SAFE HEADER MAP
-------------------------------------*/
function getHeaderMap_(sheet) {

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};

  headers.forEach((h, i) => {
    if (!h) return;
    const clean = h.toString().trim();
    map[clean] = i;
  });

  return map;
}

/*
-------------------------------------
   SAFE CELL ACCESSOR
-------------------------------------*/
function getCell_(row, headerMap, colName) {

  const key = Object.keys(headerMap).find(
    k => k.toLowerCase() === colName.toLowerCase()
  );

  if (!key) return undefined;

  return row[headerMap[key]];
}



