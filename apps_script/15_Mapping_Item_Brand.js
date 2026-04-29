// MAPPING ITEM-BRAND: POPULATE
/**
 * Script Name: populateMapping_Item_Brand_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Discover Item ↔ Brand relationships observed in Transaction_Resolution
 * - Insert mapping rows into Mapping_Item_Brand
 * - Preserve first-seen transaction evidence
 * - Capture canonical snapshots for audit visibility
 *
 * Mapping Identity:
 * (Item_ID_Machine, Brand_ID_Machine)
 *
 * Preconditions:
 * - Sheet must exist: Transaction_Resolution
 * - Sheet must exist: Mapping_Item_Brand
 * - Header row must exist in row 1
 *
 * Required columns in Transaction_Resolution:
 * - Txn_ID_Machine
 * - Created_At
 * - Item_ID_Machine
 * - Brand_ID_Machine
 * - Item_Name_Canonical
 * - Brand_Name_Canonical
 *
 * Required columns in Mapping_Item_Brand:
 * - Item_Name_Canonical
 * - Brand_Name_Canonical
 * - Item_Status_Snapshot
 * - Brand_Status_Snapshot
 * - Is_Mapping_Active
 * - Is_Analytics_Enabled
 * - Is_Archived
 * - Created_At
 * - Notes
 * - First_Seen_Txn_Date
 * - First_Seen_Txn_ID
 * - Item_ID_Machine
 * - Brand_ID_Machine
 *
 * Algorithm (Step-by-Step):
 *
 * 1. Load Mapping_Item_Brand and resolve column indexes.
 * 2. Build in-memory Set of existing mapping identities:
 *      Item_ID_Machine + Brand_ID_Machine
 *
 * 3. Load Transaction_Resolution rows.
 *
 * 4. Iterate each transaction:
 *
 *    Skip if:
 *      - Txn_ID_Machine missing
 *      - Item_ID_Machine missing
 *      - Brand_ID_Machine missing
 *
 * 5. Construct mapping identity key.
 *
 *    If identity already exists:
 *       → skip row
 *
 * 6. Create new mapping row:
 *
 *      Item_Name_Canonical
 *      Brand_Name_Canonical
 *
 *      Item_Status_Snapshot  = ""
 *      Brand_Status_Snapshot = ""
 *
 *      Is_Mapping_Active     = TRUE
 *      Is_Analytics_Enabled  = TRUE
 *      Is_Archived           = FALSE
 *
 *      Created_At = NOW()
 *      Notes      = "Discovered from Transaction_Resolution"
 *
 *      First_Seen_Txn_Date
 *      First_Seen_Txn_ID
 *
 *      Item_ID_Machine
 *      Brand_ID_Machine
 *
 * 7. Append rows in batch.
 *
 * 8. Emit execution summary and completion logs.
 *
 * Failure Modes:
 * - Required sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function populateMapping_Item_Brand_FromTransactionResolution() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME  = 'Mapping_Item_Brand';
  const FUNCTION_NAME = 'populateMapping_Item_Brand_FromTransactionResolution';
  const TGT_SHEET   = 'Mapping_Item_Brand';

  const TXN_SHEET = 'Transaction_Resolution';
  const MAP_SHEET = 'Mapping_Item_Brand';

  const t0 = new Date();

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txSh = ss.getSheetByName(TXN_SHEET);
    const mpSh = ss.getSheetByName(MAP_SHEET);

    if (!txSh || !mpSh) {
      throw new Error('Required sheet not found');
    }

    /* =========================
       STEP — LOAD_MAPPING
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_MAPPING');

    /* =========================
       READ MAPPING TABLE
    ========================= */

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

    /* =========================
       BUILD EXISTING IDENTITY SET
    ========================= */

    const existingSet = new Set();

    for (let i = 1; i < mpData.length; i++) {

      const itemId = mpData[i][IDX_MAP.itemId];
      const brandId = mpData[i][IDX_MAP.brandId];

      if (!itemId || !brandId) continue;

      const key = `${itemId}||${brandId}`;

      existingSet.add(key);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_MAPPING');

    /* =========================
       STEP — LOAD_TXN
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');

    /* =========================
       READ TRANSACTION RESOLUTION
    ========================= */

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
      if (v === -1) {
        throw new Error(`Transaction_Resolution missing column: ${k}`);
      }
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');

    /* =========================
       SKIP — NO DATA
    ========================= */

    if (txData.length <= 1) {

      ETI_logSkip_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'SRC_Table contains no data rows'
      );

      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }

    /* =========================
       COUNTERS
    ========================= */

    let scanned = 0;
    let skipNoTxn = 0;
    let skipNoItem = 0;
    let skipNoBrand = 0;
    let skipDuplicate = 0;

    const rowsToAppend = [];

    /* =========================
       DISCOVERY LOOP
    ========================= */

    for (let i = 1; i < txData.length; i++) {

      scanned++;

      const r = txData[i];

      const txnId = r[IDX_TX.txnId];
      const itemId = r[IDX_TX.itemId];
      const brandId = r[IDX_TX.brandId];

      if (!txnId) {
        skipNoTxn++;
        continue;
      }

      if (!itemId) {
        skipNoItem++;
        continue;
      }

      if (!brandId) {
        skipNoBrand++;
        continue;
      }

      const key = `${itemId}||${brandId}`;

      if (existingSet.has(key)) {
        skipDuplicate++;
        continue;
      }

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

      let firstSeenDate = r[IDX_TX.txnDateEntered];

      if (!firstSeenDate) {
        firstSeenDate = r[IDX_TX.createdAt];
      }

      if (!firstSeenDate) {
        firstSeenDate = new Date();
      }

      row[IDX_MAP.firstSeenDate] = firstSeenDate;
      row[IDX_MAP.firstSeenTxn] = txnId;

      row[IDX_MAP.itemId] = itemId;
      row[IDX_MAP.brandId] = brandId;

      rowsToAppend.push(row);

      existingSet.add(key);
    }

    /* =========================
       STEP — WRITE_OUTPUT
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if (rowsToAppend.length > 0) {

      for (let i = 0; i < rowsToAppend.length; i++) {

        ETI_log_({
          scriptName: SCRIPT_NAME,
          functionName: FUNCTION_NAME,
          sheetName: TGT_SHEET,
          level: 'INFO',
          action: 'PROCESS',
          stepName: 'WRITE_OUTPUT',
          details: 'Mapping row appended'
        });
      }

      mpSh.getRange(
        mpSh.getLastRow() + 1,
        1,
        rowsToAppend.length,
        mpHdr.length
      ).setValues(rowsToAppend);

    } else {

      /* =========================
         STEP: NOTICE — NO INSERT
      ========================= */

      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'WRITE_OUTPUT',
        'No new mappings to insert'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    /* =========================
       SUMMARY
    ========================= */

    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Scanned=${scanned} | Inserted=${rowsToAppend.length} | ` +
      `Skipped: NoTxn=${skipNoTxn}, NoItem=${skipNoItem}, NoBrand=${skipNoBrand}, Duplicate=${skipDuplicate} | ` +
      `DurationMs=${durationMs}`
    );

    /* =========================
       END
    ========================= */

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

  } catch (err) {

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      err,
      'MAIN'
    );

    throw err;

  } finally {

    flushLogs_();

  }
}





// MAPPING ITEM-BRAND: PROCESS
/**
 * Script Name: processMapping_Item_Brand_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Reconcile mapping rows against current entity governance state
 * - Update status snapshots
 * - Derive mapping flags deterministically
 * - Repair drift caused by entity lifecycle changes
 *
 * Preconditions:
 * - Sheet exists: Mapping_Item_Brand
 * - Sheet exists: Lookup_Items
 * - Sheet exists: Lookup_Brands
 * - Sheet exists: Automation_Control
 */

function processMapping_Item_Brand_StateMachine() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME = 'Mapping_Item_Brand';
  const FUNCTION_NAME = 'processMapping_Item_Brand_StateMachine';
  const TGT_SHEET = 'Mapping_Item_Brand';

  const MAP_SHEET = 'Mapping_Item_Brand';
  const ITEM_SHEET = 'Lookup_Items';
  const BRAND_SHEET = 'Lookup_Brands';
  const CTRL_SHEET = 'Automation_Control';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t0 = new Date();

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    /* 
    ======================================
       LOAD SHEETS
    ====================================== 
    */
    const mapSh = ss.getSheetByName(MAP_SHEET);
    const itemSh = ss.getSheetByName(ITEM_SHEET);
    const brandSh = ss.getSheetByName(BRAND_SHEET);

    if (!mapSh || !itemSh || !brandSh) {
      throw new Error('Required sheet missing');
    }

    /* =========================
       STEP — LOAD_DATA
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DATA');

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

    /* =========================
       BUILD ITEM STATE MAP
    ========================= */

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

      if (!id) continue;

      itemState[id] = {

        approved: r[IDX_ITEM.approved],
        active: r[IDX_ITEM.active],
        archived: r[IDX_ITEM.archived]
      };
    }

    /* =========================
       BUILD BRAND STATE MAP
    ========================= */

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

      if (!id) continue;

      brandState[id] = {

        approved: r[IDX_BRAND.approved],
        active: r[IDX_BRAND.active],
        archived: r[IDX_BRAND.archived]
      };
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DATA');

    /* =========================
       STEP — DRIFT_REPAIR
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'DRIFT_REPAIR');

    /* =========================
       PROCESS MAPPINGS
    ========================= */

    let repaired = 0;
    let valid = 0;

    for (let i = 1; i < mapData.length; i++) {

      const row = mapData[i];

      const itemId = row[IDX.itemId];
      const brandId = row[IDX.brandId];

      const item = itemState[itemId];
      const brand = brandState[brandId];

      let itemStatus = 'Unknown';
      let brandStatus = 'Unknown';

      if (item) {

        if (item.archived) itemStatus = 'Archived';
        else if (item.active) itemStatus = 'Active';
        else if (item.approved) itemStatus = 'Approved (Hidden Dropdown)';
      }

      if (brand) {

        if (brand.archived) brandStatus = 'Archived';
        else if (brand.active) brandStatus = 'Active';
        else if (brand.approved) brandStatus = 'Approved (Hidden Dropdown)';
      }

      const prevActive = row[IDX.mapActive];

      const newActive =
        !(itemStatus === 'Archived' || brandStatus === 'Archived');

      row[IDX.itemStatus] = itemStatus;
      row[IDX.brandStatus] = brandStatus;

      row[IDX.mapActive] = newActive;

      row[IDX.analytics] = true;

      if (prevActive !== newActive) {

        repaired++;

        row[IDX.notes] =
          `Mapping state updated due to entity status change`;

        /* =========================
           LOG PER ROW
        ========================= */
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

    /* =========================
       NOTICE — NO DRIFT
    ========================= */

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

    /* =========================
       STEP — WRITE_BACK
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');

    mapSh
      .getRange(2,1,mapData.length-1,mapHdr.length)
      .setValues(mapData.slice(1));

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');

    /* =========================
       SUMMARY
    ========================= */

    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Valid=${valid}, Repaired=${repaired}, DurationMs=${durationMs}`
    );

    /* =========================
       END
    ========================= */

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

  } catch (err) {

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      err,
      'MAIN'
    );

    throw err;

  } finally {

    flushLogs_();

  }
}





// MAPPING ITEM-BRAND: CLEANUP
/**
 * Script Name: cleanupMapping_Item_Brand_InvalidRows
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Permanently DELETE invalid rows from Mapping_Item_Brand
 * - A row is invalid if:
 *   - Item_ID_Machine is missing OR
 *   - Brand_ID_Machine is missing
 * - Maintain a clean, authoritative mapping table
 *
 * Preconditions:
 * - Sheet must exist: Mapping_Item_Brand
 * - Header row present in row 1
 * - Required columns (header-based):
 *   - Item_ID_Machine
 *   - Brand_ID_Machine
 *
 * Algorithm (Step-by-Step):
 *
 * 1. Generate Execution_ID for traceability.
 * 2. Load Mapping_Item_Brand sheet.
 * 3. Resolve column indexes using header names.
 * 4. Scan all rows and identify invalid records:
 *      Item_ID_Machine missing
 *      OR
 *      Brand_ID_Machine missing
 * 5. Collect row numbers to delete.
 * 6. Delete rows in reverse order (bottom-up).
 * 7. Emit execution summary and completion logs.
 *
 * Failure Modes:
 * - Mapping_Item_Brand sheet not found
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function cleanupMapping_Item_Brand_InvalidRows() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME  = 'Mapping_Item_Brand';
  const FUNCTION_NAME = 'cleanupMapping_Item_Brand_InvalidRows';
  const TGT_SHEET    = 'Mapping_Item_Brand';

  const MAP_SHEET    = 'Mapping_Item_Brand';

  const t0 = new Date();

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSh = ss.getSheetByName(MAP_SHEET);

    if (!mapSh) {
      throw new Error(`Sheet not found: ${MAP_SHEET}`);
    }

    const data = mapSh.getDataRange().getValues();

    /* =========================
       SKIP: NO DATA
    ========================= */
    if (data.length < 2) {

      ETI_logSkip_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'No data rows found'
      );

      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }

    /* =========================
       HEADER MAPPING
    ========================= */
    const header = data[0];
    const col = n => header.indexOf(n);

    const IDX = {
      itemId: col('Item_ID_Machine'),
      brandId: col('Brand_ID_Machine')
    };

    for (const [key, idx] of Object.entries(IDX)) {

      if (idx === -1) {
        throw new Error(
          `Missing required column: ${
            key === 'itemId'
              ? 'Item_ID_Machine'
              : 'Brand_ID_Machine'
          }`
        );
      }
    }

    /* =========================
       STEP — CLEANUP_INVALID_ROWS
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'CLEANUP_INVALID_ROWS');

    /* =========================
       IDENTIFY INVALID ROWS
    ========================= */

    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {

      const rowNum = i + 1;

      const itemId  = data[i][IDX.itemId];
      const brandId = data[i][IDX.brandId];

      if (!itemId || !brandId) {
        rowsToDelete.push(rowNum);
      }
    }

    /* =========================
       DELETE ROWS (BOTTOM-UP)
    ========================= */

    let deletedCount = 0;

    for (let i = rowsToDelete.length - 1; i >= 0; i--) {

      mapSh.deleteRow(rowsToDelete[i]);
      deletedCount++;

      /* =========================
         LOG PER ROW
      ========================= */
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'WARN',
        rowNumber: rowsToDelete[i],
        action: 'PROCESS',
        stepName: 'CLEANUP_INVALID_ROWS',
        details: 'Missing Item_ID_Machine or Brand_ID_Machine'
      });
    }

    /* =========================
       NOTICE — NO CLEANUP
    ========================= */

    if (deletedCount === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'CLEANUP_INVALID_ROWS',
        'No invalid rows found'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'CLEANUP_INVALID_ROWS');

    /* =========================
       SUMMARY
    ========================= */

    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Deleted=${deletedCount}, DurationMs=${durationMs}`
    );

    /* =========================
       END
    ========================= */

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

  } catch (err) {

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      err,
      'MAIN'
    );

    throw err;

  } finally {

    flushLogs_();

  }
}