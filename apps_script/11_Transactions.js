// TRANSACTION ID - BACKFILLING
/**
 * Script Name: backfillTxnIDs_TransactionRaw
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Backfill Txn_ID_Machine for valid transaction rows in Transaction_Raw
 * - Preserve idempotency and non-blocking behavior
 * - Emit both console logs (debug) and structured sheet logs (audit)
 *
 * Preconditions:
 * - Sheet must exist: Transaction_Raw
 * - Header row present in row 1
 * - Required columns:
 *   - Trx_Date_Entered
 *   - Item_Name_Entered
 *   - Qty_Value_Entered
 *   - Qty_Unit_Entered
 *   - Price_Entered
 *   - Txn_ID_Machine
 *
 * Algorithm (Step-by-Step):
 * 1. Generate Execution_ID
 * 2. Load Transaction_Raw into memory
 * 3. Resolve column indexes from header
 * 4. Iterate rows:
 *    a. Skip invalid transactions
 *    b. Skip rows with existing Txn_ID_Machine
 *    c. Generate and write Txn_ID_Machine for valid rows
 * 5. Emit execution summary
 *
 * Failure Modes:
 * - Transaction_Raw sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */
function backfillTxnIDs_TransactionRaw() {

  const SCRIPT_NAME = 'Transactions';
  const FUNCTION_NAME = 'backfillTxnIDs_TransactionRaw';
  const SHEET_NAME = 'Transaction_Raw';

  const t0 = new Date(); // used for timeout

  // =========================
  // CONTROL FLAG (ADDED)
  // =========================
  let shouldExit = false;

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) throw new Error(`Sheet ${SHEET_NAME} not found`);

    const data = sh.getDataRange().getValues();

    /* =========================
       SKIP: NO DATA (FUNCTION LEVEL)
    ========================= */
    if (data.length < 2) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'No data rows found');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);
      return;
    }

    const header = data[0];
    const col = name => header.indexOf(name);

    const IDX = {
      trxDate: col('Trx_Date_Entered'),
      item: col('Item_Name_Entered'),
      qtyVal: col('Qty_Value_Entered'),
      qtyUnit: col('Qty_Unit_Entered'),
      price: col('Price_Entered'),
      txnId: col('Txn_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Transaction_Raw missing column: ${k}`);
    }

    const numRows = data.length - 1;

    // Column buffer (Txn_ID_Machine only)
    const txnIdCol = sh
      .getRange(2, IDX.txnId + 1, numRows, 1)
      .getValues();

    let scanned = 0;
    let skipInvalid = 0;
    let skipHasTxnId = 0;
    let generated = 0;

    // Dirty Range Tracking
    let dirtyStart = null;
    let dirtyEnd = null;

    /* =================================
       STEP — GENERATE_ID
    ================================= */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'GENERATE_ID');

    for (let i = 1; i < data.length; i++) {

      scanned++;

      const rowNum = i + 1;
      const r = data[i];

      const isValidTxn =
        r[IDX.trxDate] &&
        r[IDX.item] &&
        r[IDX.qtyVal] &&
        r[IDX.qtyUnit] &&
        r[IDX.price];

      if (!isValidTxn) {
        skipInvalid++;
        continue;
      }

      if (r[IDX.txnId]) {
        skipHasTxnId++;
        continue;
      }

      /* =========================
         TIMEOUT CHECK (FIXED)
      ========================= */
      if (shouldExitForTimeout_(t0)) {

        if (dirtyStart !== null) {
          const startRow = dirtyStart + 1;
          const numDirtyRows = dirtyEnd - dirtyStart + 1;

          sh.getRange(startRow, IDX.txnId + 1, numDirtyRows, 1)
            .setValues(txnIdCol.slice(dirtyStart - 1, dirtyEnd));

          dirtyStart = null;
          dirtyEnd = null;
        }

        shouldExit = true;
        break;   // 🔥 CRITICAL FIX (no hard return)
      }

      const machineId = Utilities.getUuid();

      txnIdCol[i - 1][0] = machineId;

      if (dirtyStart === null) dirtyStart = i;
      dirtyEnd = i;

      generated++;

      /* =========================
         LOG PER ROW (MUTATION)
      ========================= */
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

      /* =========================
         PERIODIC COMMIT
      ========================= */
      if (i % 200 === 0 && dirtyStart !== null) {

        const startRow = dirtyStart + 1;
        const numDirtyRows = dirtyEnd - dirtyStart + 1;

        sh.getRange(startRow, IDX.txnId + 1, numDirtyRows, 1)
          .setValues(txnIdCol.slice(dirtyStart - 1, dirtyEnd));

        dirtyStart = null;
        dirtyEnd = null;

        flushLogs_();
      }
    }

    /* =========================
       FINAL COMMIT
    ========================= */
    if (dirtyStart !== null) {

      const startRow = dirtyStart + 1;
      const numDirtyRows = dirtyEnd - dirtyStart + 1;

      sh.getRange(startRow, IDX.txnId + 1, numDirtyRows, 1)
        .setValues(txnIdCol.slice(dirtyStart - 1, dirtyEnd));
    }

    /* =========================
       NOTICE (STEP LEVEL — NO MUTATION)
    ========================= */
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

    const durationMs = new Date().getTime() - t0.getTime();

    /* =========================
       SUMMARY
    ========================= */
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SHEET_NAME,
      `Scanned=${scanned}, Invalid=${skipInvalid}, Existing=${skipHasTxnId}, Generated=${generated}, DurationMs=${durationMs}`
    );

    /* =========================
       CONTROLLED EXIT (FIXED)
    ========================= */
    if (shouldExit) {
      return exitAndScheduleContinuation_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        {
          pipelineName: getExecutionContext_()?.pipeline_name
        }
      );
    }

    /* =========================
       END
    ========================= */
    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

  } catch (err) {

    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, err, 'MAIN');
    throw err;

  } finally {

    flushLogs_();
  }
}






// TRANSACTION ID - CLEANUP
/**
 * Script Name: cleanupInvalidTransactions_TransactionRaw
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Clean up invalid transaction rows after Txn_ID backfill
 * - Enforce prerequisite completeness invariant
 * - If transaction is partially filled → invalidate entire row
 * - Log full row snapshot for admin audit
 *
 * Preconditions:
 * - Sheet must exist: Transaction_Raw
 * - Header row present in row 1
 * - Required columns:
 *   - Trx_Date_Entered
 *   - Item_Name_Entered
 *   - Qty_Value_Entered
 *   - Qty_Unit_Entered
 *   - Price_Entered
 *   - Txn_ID_Machine
 *
 * Algorithm (Step-by-Step):
 * 1. Generate Execution_ID
 * 2. Load Transaction_Raw into memory
 * 3. Resolve column indexes from header
 * 4. Iterate rows:
 *    a. If prerequisites partially filled → INVALID
 *    b. Clear entire row (all columns)
 *    c. Log full row snapshot
 *    d. Fully empty or fully valid rows are untouched
 * 5. Write all mutations back in a single batch
 * 6. Emit execution summary
 *
 * Failure Modes:
 * - Transaction_Raw sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 * - Remains ACTIVE for ETI v1.3
 */

function cleanupInvalidTransactions_TransactionRaw() {

  const SCRIPT_NAME  = 'Transactions';
  const FUNCTION_NAME = 'cleanupInvalidTransactions_TransactionRaw';
  const SHEET_NAME   = 'Transaction_Raw';
  const t0 = new Date();

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error(`Sheet ${SHEET_NAME} not found`);

    const range = sh.getDataRange();
    const data  = range.getValues();

    /* =========================
       SKIP: NO DATA (FUNCTION LEVEL)
    ========================= */
    if (data.length < 2) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'No data rows found');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);
      return;
    }

    const header = data[0];
    const col = name => header.indexOf(name);

    const IDX = {
      trxDate: col('Trx_Date_Entered'),
      item:    col('Item_Name_Entered'),
      qtyVal:  col('Qty_Value_Entered'),
      qtyUnit: col('Qty_Unit_Entered'),
      price:   col('Price_Entered'),
      txnId:   col('Txn_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Transaction_Raw missing column: ${k}`);
    }

    let scanned = 0;
    let deleted = 0;
    const output = data.map(r => r.slice());

    /* =========================
       STEP — DELETE INVALID TXN
    ========================= */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'DELETE INVALID TXN');

    for (let i = 1; i < output.length; i++) {

      scanned++;

      const rowNum = i + 1;
      const r = output[i];

      const prereqValues = [
        r[IDX.trxDate],
        r[IDX.item],
        r[IDX.qtyVal],
        r[IDX.qtyUnit],
        r[IDX.price]
      ];

      const filledCount =
        prereqValues.filter(v => v !== '' && v !== null).length;

      if (filledCount > 0 && filledCount < prereqValues.length) {

        const snapshot = JSON.stringify(r);

        output[i] = new Array(r.length).fill('');
        deleted++;

        /* =========================
           LOG PER ROW (MUTATION)
        ========================= */
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

    /* =========================
       NOTICE (STEP LEVEL — NO MUTATION)
    ========================= */
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

    /* =========================
       WRITE BACK
    ========================= */
    range.setValues(output);

    const durationMs = new Date().getTime() - t0.getTime();

    /* =========================
       SUMMARY
    ========================= */
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SHEET_NAME,
      `Scanned=${scanned}, Deleted=${deleted}, DurationMs=${durationMs}`
    );

    /* =========================
       END
    ========================= */
    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

  } catch (err) {

    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, err, 'MAIN');
    throw err;

  } finally {

    flushLogs_();
  }
}