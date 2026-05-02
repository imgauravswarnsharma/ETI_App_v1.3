/* 
=========================================================
FUNCTION: ITEM-BRAND MAPPING DISCOVERY
=========================================================*/
/**
 * Script Name: populateMapping_Item_Brand_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Discover Item ↔ Brand relationships from transaction data
 * - Create mapping rows for new relationships only
 * - Maintain first-seen transaction traceability
 *
 * PRECONDITIONS:
 * - Sheets exist:
 *   - Transaction_Resolution
 *   - Mapping_Item_Brand
 * - Required columns exist in both sheets
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load existing mapping data and build identity set
 * 2. Load transaction data
 * 3. Iterate transaction rows:
 *    - Validate required identifiers
 *    - Skip duplicates
 *    - Construct mapping row
 * 4. Batch append new mappings
 * 5. Emit execution summary
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Process rows where:
 *     - Txn_ID_Machine exists
 *     - Item_ID_Machine exists
 *     - Brand_ID_Machine exists
 * - Deduplicate using (Item_ID + Brand_ID)
 * - Assign:
 *     - First_Seen_Txn_Date (priority: entered → created → now)
 *     - First_Seen_Txn_ID
 * - Default mapping state:
 *     - Active = TRUE
 *     - Analytics = TRUE
 *     - Archived = FALSE
 *
 * FAILURE MODES:
 * - Required sheet missing
 * - Required column missing
 */


function populateMapping_Item_Brand_FromTransactionResolution() {

  /* --- FUNCTION-LEVEL CONSTANTS --- */
  const SCRIPT_NAME  = 'Mapping_Item_Brand';
  const FUNCTION_NAME = 'populateMapping_Item_Brand_FromTransactionResolution';
  const TGT_SHEET   = 'Mapping_Item_Brand';

  const TXN_SHEET = 'Transaction_Resolution';
  const MAP_SHEET = 'Mapping_Item_Brand';

  const t0 = new Date();

  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txSh = ss.getSheetByName(TXN_SHEET);
    const mpSh = ss.getSheetByName(MAP_SHEET);

    if (!txSh || !mpSh) throw new Error('Required sheet not found');


    /* --- STEP: LOAD_MAPPING --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_MAPPING');

    const mpData = mpSh.getDataRange().getValues();
    const mpHdr  = mpData[0];
    const mpCol  = n => mpHdr.indexOf(n);

    const IDX_MAP = {
      itemCanon: mpCol('Item_Name_Canonical'),
      brandCanon: mpCol('Brand_Name_Canonical'),
      itemStatus: mpCol('Item_Status_Snapshot'),
      brandStatus: mpCol('Brand_Status_Snapshot'),
      mapActive: mpCol('Is_Mapping_Active'),
      analytics: mpCol('Is_Analytics_Enabled'),
      archived: mpCol('Is_Archived'),
      createdAt: mpCol('Created_At'),
      notes: mpCol('Notes'),
      firstSeenDate: mpCol('First_Seen_Txn_Date'),
      firstSeenTxn: mpCol('First_Seen_Txn_ID'),
      itemId: mpCol('Item_ID_Machine'),
      brandId: mpCol('Brand_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX_MAP)) {
      if (v === -1) throw new Error(`Mapping_Item_Brand missing column: ${k}`);
    }

    // Build existing mapping identity set (Item_ID + Brand_ID)
    const existingSet = new Set();

    for (let i = 1; i < mpData.length; i++) {

      const itemId = mpData[i][IDX_MAP.itemId];
      const brandId = mpData[i][IDX_MAP.brandId];

      if (!itemId || !brandId) continue;

      existingSet.add(`${itemId}||${brandId}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_MAPPING');


    /* --- STEP: LOAD_TXN --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');

    const txData = txSh.getDataRange().getValues();
    const txHdr  = txData[0];
    const txCol  = n => txHdr.indexOf(n);

    const IDX_TX = {
      txnId: txCol('Txn_ID_Machine'),
      txnDateEntered: txCol('Txn_Date_Entered'),
      createdAt: txCol('Created_At'),
      itemId: txCol('Item_ID_Machine'),
      brandId: txCol('Brand_ID_Machine'),
      itemCanon: txCol('Item_Name_Canonical'),
      brandCanon: txCol('Brand_Name_Canonical')
    };

    for (const [k,v] of Object.entries(IDX_TX)) {
      if (v === -1) throw new Error(`Transaction_Resolution missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');


    /* --- VALIDATION: SOURCE DATA --- */
    if (txData.length <= 1) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'SRC_Table contains no data rows');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }


    /*
    ---------------------------------------------------------
    PROCESS LOOP
    ---------------------------------------------------------
    // Scan transactions → discover Item ↔ Brand relationships
    // Deduplicate using mapping identity
    // Construct mapping rows
    */
    let scanned = 0;
    let skipNoTxn = 0;
    let skipNoItem = 0;
    let skipNoBrand = 0;
    let skipDuplicate = 0;

    const rowsToAppend = [];

    for (let i = 1; i < txData.length; i++) {

      scanned++;

      const r = txData[i];

      const txnId = r[IDX_TX.txnId];
      const itemId = r[IDX_TX.itemId];
      const brandId = r[IDX_TX.brandId];

      // Skip rows without transaction identity
      if (!txnId) { skipNoTxn++; continue; }

      // Skip rows without item
      if (!itemId) { skipNoItem++; continue; }

      // Skip rows without brand
      if (!brandId) { skipNoBrand++; continue; }

      const key = `${itemId}||${brandId}`;

      // Skip duplicate mapping
      if (existingSet.has(key)) { skipDuplicate++; continue; }

      const row = new Array(mpHdr.length).fill('');

      row[IDX_MAP.itemCanon] = r[IDX_TX.itemCanon];
      row[IDX_MAP.brandCanon] = r[IDX_TX.brandCanon];

      row[IDX_MAP.itemStatus] = '';
      row[IDX_MAP.brandStatus] = '';

      row[IDX_MAP.mapActive] = true;
      row[IDX_MAP.analytics] = true;
      row[IDX_MAP.archived] = false;

      row[IDX_MAP.createdAt] = new Date();
      row[IDX_MAP.notes] = 'Discovered from Transaction_Resolution';

      // Resolve first seen date (priority: entered → created → now)
      let firstSeenDate = r[IDX_TX.txnDateEntered] || r[IDX_TX.createdAt] || new Date();

      row[IDX_MAP.firstSeenDate] = firstSeenDate;
      row[IDX_MAP.firstSeenTxn] = txnId;

      row[IDX_MAP.itemId] = itemId;
      row[IDX_MAP.brandId] = brandId;

      rowsToAppend.push(row);
      existingSet.add(key);

      // Log mapping creation
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'INFO',
        action: 'PROCESS',
        stepName: 'WRITE_OUTPUT',
        details: `Item_ID=${itemId}, Brand_ID=${brandId}`
      });
    }


    /* --- STEP: WRITE_OUTPUT --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if (rowsToAppend.length > 0) {

      mpSh.getRange(
        mpSh.getLastRow() + 1,
        1,
        rowsToAppend.length,
        mpHdr.length
      ).setValues(rowsToAppend);

    } else {

      /* --- NOTICE — NO INSERT --- */
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'WRITE_OUTPUT',
        'No new mappings to insert'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');


    /* --- SUMMARY --- */
    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Scanned=${scanned} | Inserted=${rowsToAppend.length} | ` +
      `Skipped: NoTxn=${skipNoTxn}, NoItem=${skipNoItem}, NoBrand=${skipNoBrand}, Duplicate=${skipDuplicate} | ` +
      `DurationMs=${durationMs}`
    );

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

  } catch (err) {

    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, err, 'MAIN');
    throw err;

  } finally {

    flushLogs_();
  }
}



/* 
=========================================================
FUNCTION: ITEM-BRAND MAPPING STATE RECONCILIATION
=========================================================*/
/**
 * Script Name: processMapping_Item_Brand_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Reconcile mapping rows with current Item and Brand states
 * - Derive status snapshots for both entities
 * - Maintain mapping activation consistency
 *
 * PRECONDITIONS:
 * - Sheets exist:
 *   - Mapping_Item_Brand
 *   - Lookup_Items
 *   - Lookup_Brands
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load mapping, item, and brand data
 * 2. Build Item and Brand state maps
 * 3. Iterate mapping rows:
 *    - Resolve entity states
 *    - Derive snapshot statuses
 *    - Update mapping activation
 *    - Detect and log drift
 * 4. Write updated rows back
 * 5. Emit summary
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Derive Item / Brand status from:
 *     - Is_Approved, Is_Active, Is_Archived
 * - Mapping active if:
 *     - Neither Item nor Brand is Archived
 * - Always enforce:
 *     - Is_Analytics_Enabled = TRUE
 * - Detect drift when:
 *     - Previous activation ≠ derived activation
 *
 * FAILURE MODES:
 * - Required sheet missing
 * - Required column missing
 */
function processMapping_Item_Brand_StateMachine() {

  /* --- FUNCTION-LEVEL CONSTANTS --- */
  const SCRIPT_NAME = 'Mapping_Item_Brand';
  const FUNCTION_NAME = 'processMapping_Item_Brand_StateMachine';
  const TGT_SHEET = 'Mapping_Item_Brand';

  const MAP_SHEET = 'Mapping_Item_Brand';
  const ITEM_SHEET = 'Lookup_Items';
  const BRAND_SHEET = 'Lookup_Brands';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t0 = new Date();

  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);


    /* --- STEP: LOAD_DATA --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DATA');

    const mapSh = ss.getSheetByName(MAP_SHEET);
    const itemSh = ss.getSheetByName(ITEM_SHEET);
    const brandSh = ss.getSheetByName(BRAND_SHEET);

    if (!mapSh || !itemSh || !brandSh) {
      throw new Error('Required sheet missing');
    }

    const mapData = mapSh.getDataRange().getValues();
    const mapHdr = mapData[0];

    const col = n => mapHdr.indexOf(n);

    const IDX = {
      itemCanon: col('Item_Name_Canonical'),
      brandCanon: col('Brand_Name_Canonical'),
      itemStatus: col('Item_Status_Snapshot'),
      brandStatus: col('Brand_Status_Snapshot'),
      mapActive: col('Is_Mapping_Active'),
      analytics: col('Is_Analytics_Enabled'),
      archived: col('Is_Archived'),
      notes: col('Notes'),
      itemId: col('Item_ID_Machine'),
      brandId: col('Brand_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing column: ${k}`);
    }


    /* --- BUILD ITEM STATE MAP --- */
    const itemData = itemSh.getDataRange().getValues();
    const itemHdr = itemData[0];

    const ic = n => itemHdr.indexOf(n);

    const IDX_ITEM = {
      id: ic('Item_ID_Machine'),
      approved: ic('Is_Approved'),
      active: ic('Is_Active'),
      archived: ic('Is_Archived')
    };

    const itemState = {};

    for (let i = 1; i < itemData.length; i++) {

      const r = itemData[i];
      const id = r[IDX_ITEM.id];

      // Skip rows without Item_ID
      if (!id) continue;

      // Store current governance state snapshot
      itemState[id] = {
        approved: r[IDX_ITEM.approved],
        active: r[IDX_ITEM.active],
        archived: r[IDX_ITEM.archived]
      };
    }


    /* --- BUILD BRAND STATE MAP --- */
    const brandData = brandSh.getDataRange().getValues();
    const brandHdr = brandData[0];

    const bc = n => brandHdr.indexOf(n);

    const IDX_BRAND = {
      id: bc('Brand_ID_Machine'),
      approved: bc('Is_Approved'),
      active: bc('Is_Active'),
      archived: bc('Is_Archived')
    };

    const brandState = {};

    for (let i = 1; i < brandData.length; i++) {

      const r = brandData[i];
      const id = r[IDX_BRAND.id];

      // Skip rows without Brand_ID
      if (!id) continue;

      // Store current governance state snapshot
      brandState[id] = {
        approved: r[IDX_BRAND.approved],
        active: r[IDX_BRAND.active],
        archived: r[IDX_BRAND.archived]
      };
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DATA');


    /* --- STEP: STATE_MACHINE --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'DRIFT_REPAIR');

    let repaired = 0;
    let valid = 0;


    /*
    ---------------------------------------------------------
    PROCESS LOOP
    ---------------------------------------------------------
    // Reconcile mapping rows with current Item + Brand state
    // Derive status snapshots
    // Update mapping activation flag
    */
    for (let i = 1; i < mapData.length; i++) {

      const row = mapData[i];

      const itemId = row[IDX.itemId];
      const brandId = row[IDX.brandId];

      const item = itemState[itemId];
      const brand = brandState[brandId];

      let itemStatus = 'Unknown';
      let brandStatus = 'Unknown';


      // Derive Item status from governance flags
      if (item) {
        if (item.archived) itemStatus = 'Archived';
        else if (item.active) itemStatus = 'Active';
        else if (item.approved) itemStatus = 'Approved (Hidden Dropdown)';
      }

      // Derive Brand status from governance flags
      if (brand) {
        if (brand.archived) brandStatus = 'Archived';
        else if (brand.active) brandStatus = 'Active';
        else if (brand.approved) brandStatus = 'Approved (Hidden Dropdown)';
      }


      const prevActive = row[IDX.mapActive];

      // Mapping active only if BOTH entities are not archived
      const newActive =
        !(itemStatus === 'Archived' || brandStatus === 'Archived');


      // Update snapshot columns
      row[IDX.itemStatus] = itemStatus;
      row[IDX.brandStatus] = brandStatus;

      row[IDX.mapActive] = newActive;
      row[IDX.analytics] = true; // Always enabled


      // Detect drift (activation state change)
      if (prevActive !== newActive) {

        repaired++;

        row[IDX.notes] =
          `Mapping state updated due to entity status change`;

        // Log drift repair
        ETI_log_({
          scriptName: SCRIPT_NAME,
          functionName: FUNCTION_NAME,
          sheetName: TGT_SHEET,
          level: 'WARN',
          rowNumber: i + 1,
          action: 'PROCESS',
          stepName: 'DRIFT_REPAIR',
          details: `Item_ID=${itemId}, Brand_ID=${brandId}, Active changed from ${prevActive} to ${newActive}`
        });

      } else {

        valid++;
      }
    }


    /* --- NOTICE: NO DRIFT --- */
    if (repaired === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'DRIFT_REPAIR',
        'No drift detected; all mappings already aligned'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'DRIFT_REPAIR');


    /* --- STEP: WRITE_BACK --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');

    mapSh
      .getRange(2, 1, mapData.length - 1, mapHdr.length)
      .setValues(mapData.slice(1));

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');


    /* --- SUMMARY --- */
    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Valid=${valid}, Repaired=${repaired}, DurationMs=${durationMs}`
    );

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

  } catch (err) {

    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, err, 'MAIN');
    throw err;

  } finally {

    flushLogs_();
  }
}



/* 
=========================================================
FUNCTION: ITEM-BRAND MAPPING CLEANUP
=========================================================*/
/**
 * Script Name: cleanupMapping_Item_Brand_InvalidRows
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Remove invalid mapping rows
 * - Ensure mapping table integrity
 *
 * PRECONDITIONS:
 * - Sheet exists: Mapping_Item_Brand
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load mapping data
 * 2. Iterate rows:
 *    - Identify invalid rows
 *    - Collect row indices
 * 3. Delete rows in reverse order
 * 4. Emit summary
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Row is invalid if:
 *     - Item_ID_Machine missing OR
 *     - Brand_ID_Machine missing OR
 *     - Item_Name_Canonical missing OR
 *     - Brand_Name_Canonical missing
 * - Perform deletion after scan (reverse order)
 *
 * FAILURE MODES:
 * - Sheet missing
 * - Column missing
 */
function cleanupMapping_Item_Brand_InvalidRows() {

  /* --- FUNCTION-LEVEL CONSTANTS --- */
  const SCRIPT_NAME = 'Mapping_Item_Brand';
  const FUNCTION_NAME = 'cleanupMapping_Item_Brand_InvalidRows';
  const TGT_SHEET = 'Mapping_Item_Brand';

  const SHEET_NAME = 'Mapping_Item_Brand';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t0 = new Date();

  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);


    /* --- STEP: LOAD_DATA --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DATA');

    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error('Mapping_Item_Brand sheet not found');

    const data = sh.getDataRange().getValues();
    const hdr = data[0];
    const col = n => hdr.indexOf(n);

    const IDX = {
      itemId: col('Item_ID_Machine'),
      brandId: col('Brand_ID_Machine'),
      itemCanon: col('Item_Name_Canonical'),
      brandCanon: col('Brand_Name_Canonical')
    };

    for (const [k,v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DATA');


    /*
    ---------------------------------------------------------
    PROCESS LOOP
    ---------------------------------------------------------
    // Identify invalid mapping rows
    // Condition: missing Item_ID OR missing Brand_ID
    // OR missing canonical values
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'CLEANUP');

    let scanned = 0;
    let removed = 0;

    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {

      scanned++;

      const r = data[i];
      const rowNum = i + 1;

      const itemId = r[IDX.itemId];
      const brandId = r[IDX.brandId];
      const itemCanon = r[IDX.itemCanon];
      const brandCanon = r[IDX.brandCanon];

      // Invalid if critical identifiers are missing
      const isInvalid =
        (!itemId || !brandId) ||
        (!itemCanon || !brandCanon);

      if (!isInvalid) continue;

      rowsToDelete.push(rowNum);
      removed++;

      // Log deletion candidate
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'WARN',
        rowNumber: rowNum,
        action: 'DELETE',
        stepName: 'CLEANUP',
        details: `Invalid mapping row detected`
      });
    }


    // No-op notice (no invalid rows)
    if (removed === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'CLEANUP',
        'No invalid mapping rows found'
      );
    }


    // Delete rows (reverse order to avoid index shift)
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sh.deleteRow(rowsToDelete[i]);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'CLEANUP');


    /* --- SUMMARY --- */
    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Scanned=${scanned}, Removed=${removed}, DurationMs=${durationMs}`
    );

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

  } catch (err) {

    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, err, 'MAIN');
    throw err;

  } finally {

    flushLogs_();
  }
}