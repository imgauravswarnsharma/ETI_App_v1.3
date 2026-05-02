/* 
=========================================================
  FUNCTION: TXN ID BACKFILL
========================================================= */
/**
 * Script Name: backfillTxnIDs_TransactionRaw
 * Script Language: JavaScript (Google Apps Script)
 * App Version: v1.3
 * 
 * PURPOSE:
 * - Assign Txn_ID_Machine to valid transaction rows
 * - Ensure idempotent behavior (skip already processed rows)
 * - Support scheduler-driven execution (pause + resume safe)
 * - Use buffered batch writes for performance and data safety
 *
 * INPUT DEPENDENCIES:
 * - Sheet: Transaction_Raw
 * - Required Columns:
 *   - Trx_Date_Entered
 *   - Item_Name_Entered
 *   - Qty_Value_Entered
 *   - Qty_Unit_Entered
 *   - Price_Entered
 *   - Txn_ID_Machine
 *
 * OUTPUT TARGETS:
 * - Sheet: Transaction_Raw
 * - Mutation:
 *   - Update Txn_ID_Machine column (row-level assignment)
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Initialize execution (logging + timer)
 *
 * 2. Load full dataset from Transaction_Raw into memory
 *
 * 3. Validate dataset presence
 *    - Exit if no data rows
 *
 * 4. Resolve column indices from header
 *
 * 5. Initialize:
 *    - Txn_ID buffer (column-level)
 *    - Execution counters
 *    - Dirty range tracking for batch writes
 *
 * 6. Start processing loop (row-by-row):
 *    a. Validate required fields
 *    b. Skip invalid rows
 *    c. Skip rows with existing Txn_ID_Machine
 *    d. Check scheduler timeout:
 *       - If triggered:
 *         - Flush pending buffered writes
 *         - Exit loop safely
 *    e. Generate Txn_ID_Machine (UUID)
 *    f. Store in buffer (no direct write)
 *    g. Update dirty range
 *    h. Log mutation
 *    i. Periodically flush buffer (batch write)
 *
 * 7. After loop:
 *    - Flush remaining buffered writes
 *
 * 8. Close processing step and emit execution summary
 *
 * 9. If scheduler exit triggered:
 *    - Persist state and schedule continuation
 *
 * 10. End execution and flush logs
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Iterate all transaction rows
 *
 * - Process only rows where:
 *     - All required fields are present
 *     - Txn_ID_Machine is empty
 *
 * - For each eligible row:
 *     - Generate UUID using Utilities.getUuid()
 *     - Store value in in-memory buffer
 *
 * - Maintain contiguous dirty range for efficient batch writes
 *
 * - Periodically flush buffered values to sheet
 *   (prevents memory overflow and maintains write consistency)
 *
 * - On timeout:
 *     - Flush any pending buffered writes before exit
 *     - Ensure no partial or inconsistent writes occur
 *
 * - Idempotency guarantee:
 *     - Rows with existing Txn_ID_Machine are always skipped
 *     - Re-runs do not modify already processed rows
 *
 * - Write strategy:
 *     - Column-level batch updates (Txn_ID_Machine only)
 *     - No full dataset overwrite
 * 
 */


function backfillTxnIDs_TransactionRaw() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME = 'Transactions';
  const FUNCTION_NAME = 'backfillTxnIDs_TransactionRaw';
  const SHEET_NAME = 'Transaction_Raw';

  const t0 = new Date(); // Execution timer

  let shouldExit = false; // Scheduler-controlled exit flag (loop-safe, no return inside loop)

  /*
  ============================================
  CORE EXECUTION BLOCK
  ============================================*/
  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) throw new Error(`Sheet ${SHEET_NAME} not found`);

    const data = sh.getDataRange().getValues(); // Load full dataset into memory


    /* --- INPUT VALIDATION (FUNCTION LEVEL) --- */
    if (data.length < 2) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'No data rows found');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);
      
      return;   // Idempotent exit on empty dataset (only header or completely empty)
    }


    /* --- HEADER RESOLUTION --- */
    // Map column names to indices for easy reference
    const header = data[0];
    const col = name => header.indexOf(name);   

    // Required columns and their indices
    const IDX = {
      trxDate: col('Trx_Date_Entered'),
      item: col('Item_Name_Entered'),
      qtyVal: col('Qty_Value_Entered'),
      qtyUnit: col('Qty_Unit_Entered'),
      price: col('Price_Entered'),
      txnId: col('Txn_ID_Machine')
    };

    // Enforce schema contract (hard fail if broken)
    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Transaction_Raw missing column: ${k}`);
    }


    /* --- BUFFER SETUP --- */
    const numRows = data.length - 1;

    // Working in-memory buffer for Txn_ID column (avoids per-row writes)
    const txnIdCol = 
      sh.getRange(2, IDX.txnId + 1, numRows, 1)
      .getValues();

    // Execution counters (for logging summary only)
    let scanned = 0;
    let skipInvalid = 0;
    let skipHasTxnId = 0;
    let generated = 0;

    // Dirty range = contiguous block to flush in batch
    let dirtyStart = null;
    let dirtyEnd = null;


    /*
    ---------------------------------------------------------
    PROCESS LOOP [GENERATE_ID]
    ---------------------------------------------------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'GENERATE_ID');

    // Iterate over data rows (starting from 1 to skip header)
    for (let i = 1; i < data.length; i++) { 

      scanned++;

      const rowNum = i + 1;
      const r = data[i];


      /* --- VALIDATION: REQUIRED FIELDS --- */
      // Transaction is valid only if all required fields exist
      const isValidTxn =      
        r[IDX.trxDate] &&
        r[IDX.item] &&
        r[IDX.qtyVal] &&
        r[IDX.qtyUnit] &&
        r[IDX.price];

      if (!isValidTxn) {  
        skipInvalid++;      // Skip invalid transactions but count them
        continue;
      }

      
      if (r[IDX.txnId]) {
        skipHasTxnId++;     // If Txn_ID already exists, skip but count it
        continue;
      }

      /* --- SCHEDULER: TIMEOUT CHECK --- */
      // Check for timeout at the start of each iteration to allow graceful exit
      if (shouldExitForTimeout_(t0)) {

        if (dirtyStart !== null) {

          // Flush pending buffered writes before exit
          const startRow = dirtyStart + 1;
          const numDirtyRows = dirtyEnd - dirtyStart + 1;

          // Flush the dirty range of Txn_IDs to the sheet
          sh.getRange(startRow, IDX.txnId + 1, numDirtyRows, 1)
            .setValues(txnIdCol.slice(dirtyStart - 1, dirtyEnd));

          // Reset buffer state after flush
          dirtyStart = null;
          dirtyEnd = null;
        }

        shouldExit = true;
        break; // Controlled loop exit
      }

      /* --- GENERATE: TXN_ID --- */
      const machineId = Utilities.getUuid();    // Generate a unique Txn_ID (using UUID for simplicity)
      txnIdCol[i - 1][0] = machineId;           // Buffer the generated Txn_ID in memory for batch writing


      /* --- BUFFER: DIRTY RANGE TRACKING --- */
      // Mark the current row as dirty for later flushing
      if (dirtyStart === null) dirtyStart = i;
      dirtyEnd = i;

      generated++; // Increment generated count for logging

      // Logging the mutated row with new Txn_ID for traceability
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: SHEET_NAME,
        level: 'INFO',
        rowNumber: rowNum,
        action: 'PROCESS',
        stepName: 'GENERATE_ID',
        details: `Generated Txn_ID_Machine: ${machineId}`
      });


      /* --- BUFFER: PERIODIC FLUSH CONTROL --- */
      // Write buffered Txn_IDs back to the sheet once batch limit is reached
      if (i % 200 === 0 && dirtyStart !== null) {

        const startRow = dirtyStart + 1;
        const numDirtyRows = dirtyEnd - dirtyStart + 1;

        // Flush the dirty range of Txn_IDs to the sheet
        sh.getRange(startRow, IDX.txnId + 1, numDirtyRows, 1)
          .setValues(txnIdCol.slice(dirtyStart - 1, dirtyEnd));

        // Reset buffer state after flush
        dirtyStart = null;
        dirtyEnd = null;

        flushLogs_(); // keep logs consistent with partial writes
      }
    }


    /* --- FINAL BUFFER FLUSH --- */
    // After loop completion, flush any remaining buffered Txn_IDs to the sheet
    if (dirtyStart !== null) {
       
      // Write buffered Txn_IDs back to the sheet for the final dirty range
      const startRow = dirtyStart + 1;
      const numDirtyRows = dirtyEnd - dirtyStart + 1;

      sh.getRange(startRow, IDX.txnId + 1, numDirtyRows, 1)
        .setValues(txnIdCol.slice(dirtyStart - 1, dirtyEnd));
    }


    /* --- INSTRUMENTATION: STEP CLOSE + SUMMARY --- */
    if (generated === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SHEET_NAME,
        'GENERATE_ID',
        'No Txn_ID generated (all rows invalid or already processed)'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'GENERATE_ID');

    // Logging Summary
    const durationMs = new Date().getTime() - t0.getTime();
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SHEET_NAME,
      `Scanned=${scanned}, Invalid=${skipInvalid}, Existing=${skipHasTxnId}, Generated=${generated}, DurationMs=${durationMs}`
    );

    
    /* --- SCHEDULER: EXIT AFTER LOOP (IF TIMEOUT) --- */
      if (shouldExit) {
      return exitAndScheduleContinuation_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        {
          pipelineName: getExecutionContext_()?.pipeline_name
        }
      );
    }

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

  }


  /*
  ============================================
  MODULE: ERROR HANDLING BLOCK
  ============================================*/
  catch (err) {

    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, err, 'MAIN');
    throw err;
  }


  /*
  ============================================
  MODULE: GUARANTEED FINALIZATION BLOCK
  ============================================*/
  finally {
    
    flushLogs_(); // guarantees log persistence in all paths
  }

}


/* 
=========================================================
  FUNCTION: TXN CLEANUP (INVALID TRANSACTIONS)
========================================================= */
/**
 * 
 * Script Language: JavaScript (Google Apps Script)
 * Script Name: cleanupInvalidTransactions_TransactionRaw
 * App Version: v1.3
 *
 * PURPOSE:
 * - Enforce transaction completeness invariant
 * - Identify partially filled transactions
 * - Clear entire row for invalid (partial) transactions
 * - Maintain audit trace via row snapshot logging
 *
 * INPUT DEPENDENCIES:
 * - Sheet: Transaction_Raw
 * - Required Columns:
 *   - Trx_Date_Entered
 *   - Item_Name_Entered
 *   - Qty_Value_Entered
 *   - Qty_Unit_Entered
 *   - Price_Entered
 *   - Txn_ID_Machine
 *
 * OUTPUT TARGETS:
 * - Sheet: Transaction_Raw
 * - Mutation:
 *   - Full row overwrite (invalid rows cleared to empty values)
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Initialize execution (logging + timer)
 *
 * 2. Load full dataset from Transaction_Raw into memory
 *
 * 3. Validate dataset presence
 *    - Exit if no data rows
 *
 * 4. Resolve column indices from header
 *
 * 5. Initialize:
 *    - Working copy of dataset (full in-memory clone)
 *    - Execution counters
 *
 * 6. Start processing loop (row-by-row):
 *    a. Extract prerequisite fields
 *    b. Compute filled field count
 *    c. Identify partially filled rows:
 *       - Condition: 0 < filledCount < totalRequiredFields
 *    d. If row is valid (fully empty OR fully filled):
 *       - Skip
 *    e. If row is invalid (partially filled):
 *       - Check scheduler timeout:
 *           - If triggered:
 *               - Exit loop safely (no writes performed yet)
 *       - Capture row snapshot
 *       - Clear entire row in working dataset
 *       - Log mutation
 *
 * 7. After loop:
 *    - Write full dataset back in a single batch operation
 *
 * 8. Close processing step and emit execution summary
 *
 * 9. If scheduler exit triggered:
 *    - Persist state and schedule continuation
 *
 * 10. End execution and flush logs
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Iterate all transaction rows
 *
 * - For each row:
 *     - Extract prerequisite fields
 *     - Compute filledCount:
 *         count of non-empty prerequisite fields
 *
 * - Identify invalid rows:
 *     - Condition:
 *         0 < filledCount < totalRequiredFields
 *
 * - For each invalid row:
 *     - Capture row snapshot (JSON)
 *     - Replace entire row with empty values
 *
 * - Perform all mutations in-memory using working dataset copy
 *
 * - After loop:
 *     - Write entire dataset back using single setValues()
 *
 * - Scheduler behavior:
 *     - Timeout check occurs only before mutation
 *     - No partial writes occur before exit
 *
 * - Idempotency guarantee:
 *     - Re-running produces same output
 *     - Already cleaned rows remain unchanged
 *
 * - Write strategy:
 *     - Full dataset overwrite (single batch write)
 */

function cleanupInvalidTransactions_TransactionRaw() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME  = 'Transactions';
  const FUNCTION_NAME = 'cleanupInvalidTransactions_TransactionRaw';
  const SHEET_NAME   = 'Transaction_Raw';

  const t0 = new Date(); // Execution timer

  let shouldExit = false; // Scheduler-controlled exit flag

  /*
  ============================================
  CORE EXECUTION BLOCK
  ============================================*/
  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) throw new Error(`Sheet ${SHEET_NAME} not found`);

    const range = sh.getDataRange();
    const data = range.getValues(); // Load full dataset into memory


    /* --- INPUT VALIDATION (FUNCTION LEVEL) --- */
    if (data.length < 2) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'No data rows found');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);
      
      return;   // Idempotent exit on empty dataset (only header or completely empty)
    }


    /* --- HEADER RESOLUTION --- */
    // Map column names to indices for easy reference
    const header = data[0];
    const col = name => header.indexOf(name);

    // Required columns and their indices
    const IDX = {
      trxDate: col('Trx_Date_Entered'),
      item:    col('Item_Name_Entered'),
      qtyVal:  col('Qty_Value_Entered'),
      qtyUnit: col('Qty_Unit_Entered'),
      price:   col('Price_Entered'),
      txnId:   col('Txn_ID_Machine')
    };

    // Enforce schema contract (hard fail if broken)
    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Transaction_Raw missing column: ${k}`);
    }


    /* --- BUFFER SETUP --- */
    let scanned = 0;
    let deleted = 0;

    // Create full working copy for safe mutation
    const output = data.map(r => r.slice());


    /*
    ---------------------------------------------------------
    PROCESS LOOP [DELETE INVALID TXN]
    ---------------------------------------------------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'DELETE INVALID TXN');

    // Iterate over data rows (starting from 1 to skip header)
    for (let i = 1; i < output.length; i++) {

      scanned++;

      const rowNum = i + 1;
      const r = output[i];


      /* --- VALIDATION: PARTIAL TRANSACTION CHECK --- */
      // Count how many prerequisite fields are filled
      const prereqValues = [
        r[IDX.trxDate],
        r[IDX.item],
        r[IDX.qtyVal],
        r[IDX.qtyUnit],
        r[IDX.price]
      ];

      const filledCount =
        prereqValues.filter(v => v !== '' && v !== null).length;


      // Partially filled = invalid → must delete entire row
      if (filledCount > 0 && filledCount < prereqValues.length)
        {
          /* --- SCHEDULER: TIMEOUT CHECK --- */
          // Ensure no partial mutation/logging occurs before exit
          if (shouldExitForTimeout_(t0)) {
            shouldExit = true;
            break;
          }
          
        const snapshot = JSON.stringify(r); // capture for audit

        /* --- CLEAR: INVALID ROW --- */
        output[i] = new Array(r.length).fill('');
        deleted++;

        // Logging the mutated row for traceability
        ETI_log_({
          scriptName: SCRIPT_NAME,
          functionName: FUNCTION_NAME,
          sheetName: SHEET_NAME,
          level: 'WARN',
          rowNumber: rowNum,
          action: 'PROCESS',
          stepName: 'DELETE INVALID TXN',
          details: `Invalid partial transaction deleted. Snapshot=${snapshot}`
        });
      }
    }


    /* --- WRITE BACK: APPLY ALL MUTATIONS --- */
    // Single batch write (no partial flushing needed)
    range.setValues(output);


    /* --- INSTRUMENTATION: STEP CLOSE + SUMMARY --- */
    if (deleted === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SHEET_NAME,
        'DELETE INVALID TXN',
        'No invalid transactions found'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'DELETE INVALID TXN');
    
    const durationMs = new Date().getTime() - t0.getTime(); // Execution timer ended
    // Logging Summary
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SHEET_NAME,
      `Scanned=${scanned}, Deleted=${deleted}, DurationMs=${durationMs}`
    );


    /* --- SCHEDULER: EXIT AFTER LOOP (IF TIMEOUT) --- */
    if (shouldExit) {
      return exitAndScheduleContinuation_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        {
          pipelineName: getExecutionContext_()?.pipeline_name
        }
      );
    }

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

  }


  /*
  ============================================
  MODULE: ERROR HANDLING BLOCK
  ============================================*/
  catch (err) {

    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, err, 'MAIN');
    throw err;
  }


  /*
  ============================================
  MODULE: GUARANTEED FINALIZATION BLOCK
  ============================================*/
  finally {

    flushLogs_(); // guarantees log persistence in all paths
  }
}