// MAPPING ITEM-BRAND-PRODUCT: POPULATE
/**
 * Script Name: populateMapping_Item_Brand_Product_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Populate Mapping_Item_Brand_Product with unique
 *   (Item_ID_Machine, Brand_ID_Machine, Product_ID_Machine) relationships
 *   discovered from Transaction_Resolution.
 * - Each mapping is recorded exactly once (first-seen semantics).
 * - Transaction rows serve strictly as discovery evidence.
 *
 * Preconditions:
 * - Sheets must exist:
 *     Transaction_Resolution
 *     Mapping_Item_Brand_Product
 * - Header row present in row 1
 * - Required columns must exist (header-based)
 *
 * Algorithm (Step-by-Step):
 *
 * 1. Generate Execution_ID for traceability.
 * 2. Load Mapping_Item_Brand_Product and build an in-memory set of existing
 *    identity keys:
 *
 *        Item_ID | Brand_ID | Product_ID
 *
 * 3. Load Transaction_Resolution.
 *
 * 4. Iterate rows sequentially:
 *
 *      Skip if:
 *        - Txn_ID_Machine missing
 *        - Item_ID_Machine missing
 *        - Brand_ID_Machine missing
 *        - Product_ID_Machine missing
 *
 *      Build identity key.
 *
 *      Capture only the FIRST occurrence of each identity key
 *      within the current execution batch.
 *
 * 5. Construct mapping rows using:
 *
 *      Canonical snapshots
 *      Evidence metadata
 *      Default governance flags
 *
 * 6. Append rows in batch write.
 *
 * 7. Emit execution summary logs.
 *
 * Failure Modes:
 * - Required sheet missing
 * - Required column missing
 *
 * Notes:
 * - Discovery script NEVER modifies existing rows.
 * - Governance reconciliation handled by processing script.
 */
function populateMapping_Item_Brand_Product_FromTransactionResolution() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME  = 'Mapping_Item_Brand_Product';
  const FUNCTION_NAME = 'populateMapping_Item_Brand_Product_FromTransactionResolution';
  const TGT_SHEET   = 'Mapping_Item_Brand_Product';

  const TXN_SHEET = 'Transaction_Resolution';
  const MAP_SHEET = 'Mapping_Item_Brand_Product';

  const t0 = new Date();

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tsSh  = ss.getSheetByName(TXN_SHEET);
    const mapSh = ss.getSheetByName(MAP_SHEET);

    if (!tsSh || !mapSh) {
      throw new Error('Required sheet not found');
    }

    /* =========================
       STEP — LOAD_MAPPING
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_MAPPING');

    const mapData = mapSh.getDataRange().getValues();
    const mapHdr  = mapData[0];

    const mapCol = n => mapHdr.indexOf(n);

    const IDX_MAP = {
      itemId: mapCol('Item_ID_Machine'),
      brandId: mapCol('Brand_ID_Machine'),
      productId: mapCol('Product_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX_MAP)) {
      if (v === -1) {
        throw new Error(`Mapping_Item_Brand_Product missing column: ${k}`);
      }
    }

    const existingSet = new Set();

    for (let i = 1; i < mapData.length; i++) {

      const iId = mapData[i][IDX_MAP.itemId];
      const bId = mapData[i][IDX_MAP.brandId];
      const pId = mapData[i][IDX_MAP.productId];

      if (iId && bId && pId) {
        existingSet.add(iId + '|' + bId + '|' + pId);
      }
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_MAPPING');

    /* =========================
       STEP — LOAD_TXN
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');

    const tsData = tsSh.getDataRange().getValues();
    const tsHdr  = tsData[0];

    const txCol = n => tsHdr.indexOf(n);

    const IDX_TX = {

      txnId: txCol('Txn_ID_Machine'),
      txnDateEntered: txCol('Txn_Date_Entered'),
      createdAt: txCol('Created_At'),

      itemId: txCol('Item_ID_Machine'),
      brandId: txCol('Brand_ID_Machine'),
      productId: txCol('Product_ID_Machine'),

      itemCanon: txCol('Item_Name_Canonical'),
      brandCanon: txCol('Brand_Name_Canonical'),
      productCanon: txCol('Product_Name_Canonical')
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

    if (tsData.length <= 1) {

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
       DISCOVERY LOOP
    ========================= */

    const firstSeenMap = new Map();

    let scanned = 0;

    for (let i = 1; i < tsData.length; i++) {

      scanned++;

      const r = tsData[i];

      if (!r[IDX_TX.txnId]) continue;

      if (!r[IDX_TX.itemId] ||
          !r[IDX_TX.brandId] ||
          !r[IDX_TX.productId]) continue;

      const key =
        r[IDX_TX.itemId] + '|' +
        r[IDX_TX.brandId] + '|' +
        r[IDX_TX.productId];

      if (firstSeenMap.has(key)) continue;

      let firstSeenDate = r[IDX_TX.txnDateEntered];

      if (!firstSeenDate) {
        firstSeenDate = r[IDX_TX.createdAt];
      }

      if (!firstSeenDate) {
        firstSeenDate = new Date();
      }

      firstSeenMap.set(key, {

        itemId: r[IDX_TX.itemId],
        brandId: r[IDX_TX.brandId],
        productId: r[IDX_TX.productId],

        txnId: r[IDX_TX.txnId],
        txnDate: firstSeenDate,

        itemCanon: r[IDX_TX.itemCanon] || '',
        brandCanon: r[IDX_TX.brandCanon] || '',
        productCanon: r[IDX_TX.productCanon] || ''
      });
    }

    /* =========================
       BUILD ROWS
    ========================= */

    const rowsToAppend = [];

    for (const [key,v] of firstSeenMap.entries()) {

      if (existingSet.has(key)) continue;

      const row = new Array(mapHdr.length).fill('');

      row[mapCol('Item_Name_Canonical')]    = v.itemCanon;
      row[mapCol('Brand_Name_Canonical')]   = v.brandCanon;
      row[mapCol('Product_Name_Canonical')] = v.productCanon;

      row[mapCol('Is_Mapping_Active')]      = true;
      row[mapCol('Is_Analytics_Enabled')]   = true;
      row[mapCol('Is_Archived')]            = false;

      row[mapCol('Created_At')]             = new Date();
      row[mapCol('Notes')]                  = 'Discovered from Transaction_Resolution';

      row[mapCol('First_Seen_Txn_Date')]    = v.txnDate;
      row[mapCol('First_Seen_Txn_ID')]      = v.txnId;

      row[mapCol('Item_ID_Machine')]        = v.itemId;
      row[mapCol('Brand_ID_Machine')]       = v.brandId;
      row[mapCol('Product_ID_Machine')]     = v.productId;

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

      mapSh.getRange(
        mapSh.getLastRow() + 1,
        1,
        rowsToAppend.length,
        rowsToAppend[0].length
      ).setValues(rowsToAppend);

    } else {

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
      `Scanned=${scanned} | Inserted=${rowsToAppend.length} | DurationMs=${durationMs}`
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






// MAPPING ITEM-BRAND-PRODUCT: PROCESS
/**
 * Script Name: processMapping_Item_Brand_Product_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Reconcile Item–Brand–Product mappings with current entity states
 * - Update snapshot status columns
 * - Derive mapping flags deterministically
 * - Repair drift caused by entity lifecycle changes
 *
 * Preconditions:
 * - Sheets must exist:
 *     Mapping_Item_Brand_Product
 *     Lookup_Items
 *     Lookup_Brands
 *     Lookup_Products
 *     Automation_Control
 */

function processMapping_Item_Brand_Product_StateMachine() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME  = 'Mapping_Item_Brand_Product';
  const FUNCTION_NAME = 'processMapping_Item_Brand_Product_StateMachine';
  const TGT_SHEET   = 'Mapping_Item_Brand_Product';

  const MAP_SHEET   = 'Mapping_Item_Brand_Product';
  const ITEM_SHEET  = 'Lookup_Items';
  const BRAND_SHEET = 'Lookup_Brands';
  const PROD_SHEET  = 'Lookup_Products';

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
    const mapSh   = ss.getSheetByName(MAP_SHEET);
    const itemSh  = ss.getSheetByName(ITEM_SHEET);
    const brandSh = ss.getSheetByName(BRAND_SHEET);
    const prodSh  = ss.getSheetByName(PROD_SHEET);

    if (!mapSh || !itemSh || !brandSh || !prodSh) {
      throw new Error('Required sheet missing');
    }

    /* =========================
       STEP — LOAD_MAPPING
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_MAPPING');

    const mapData = mapSh.getDataRange().getValues();
    const mapHdr  = mapData[0];

    const col = n => mapHdr.indexOf(n);

    const IDX = {

      itemCanon: col('Item_Name_Canonical'),
      brandCanon: col('Brand_Name_Canonical'),
      productCanon: col('Product_Name_Canonical'),

      itemStatus: col('Item_Status_Snapshot'),
      brandStatus: col('Brand_Status_Snapshot'),
      productStatus: col('Product_Status_Snapshot'),

      mapActive: col('Is_Mapping_Active'),
      analytics: col('Is_Analytics_Enabled'),
      archived: col('Is_Archived'),

      notes: col('Notes'),

      itemId: col('Item_ID_Machine'),
      brandId: col('Brand_ID_Machine'),
      productId: col('Product_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing column: ${k}`);
    }

    /* ======================================
       BUILD ENTITY STATE MAPS
    ====================================== */

    function buildStateMap(sheet, idColName) {

      const data = sheet.getDataRange().getValues();
      const hdr  = data[0];
      const c = n => hdr.indexOf(n);

      const IDX_STATE = {
        id: c(idColName),
        approved: c('Is_Approved'),
        active: c('Is_Active'),
        archived: c('Is_Archived')
      };

      const stateMap = {};

      for (let i = 1; i < data.length; i++) {

        const r = data[i];
        const id = r[IDX_STATE.id];

        if (!id) continue;

        stateMap[id] = {

          approved: r[IDX_STATE.approved],
          active: r[IDX_STATE.active],
          archived: r[IDX_STATE.archived]
        };
      }

      return stateMap;
    }

    const itemState  = buildStateMap(itemSh,  'Item_ID_Machine');
    const brandState = buildStateMap(brandSh, 'Brand_ID_Machine');
    const prodState  = buildStateMap(prodSh,  'Product_ID_Machine');

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_MAPPING');

    /* =========================
       STEP — DRIFT_REPAIR
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'DRIFT_REPAIR');

    let repaired = 0;
    let valid = 0;

    function resolveStatus(state) {

      if (!state) return 'Unknown';

      if (state.archived) return 'Archived';
      if (state.active)   return 'Active';
      if (state.approved) return 'Approved (Hidden Dropdown)';

      return 'Rejected';
    }

    for (let i = 1; i < mapData.length; i++) {

      const row = mapData[i];

      const itemId    = row[IDX.itemId];
      const brandId   = row[IDX.brandId];
      const productId = row[IDX.productId];

      const itemStatus    = resolveStatus(itemState[itemId]);
      const brandStatus   = resolveStatus(brandState[brandId]);
      const productStatus = resolveStatus(prodState[productId]);

      const prevActive = row[IDX.mapActive];

      const newActive =
        !(itemStatus === 'Archived' ||
          brandStatus === 'Archived' ||
          productStatus === 'Archived');

      row[IDX.itemStatus]    = itemStatus;
      row[IDX.brandStatus]   = brandStatus;
      row[IDX.productStatus] = productStatus;

      row[IDX.mapActive] = newActive;
      row[IDX.analytics] = true;

      if (prevActive !== newActive) {

        repaired++;

        row[IDX.notes] =
          `Mapping state updated due to entity lifecycle change`;

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
          details: `Item_ID=${itemId}, Brand_ID=${brandId}, Product_ID=${productId}, Active changed from ${prevActive} to ${newActive}`
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
       STEP — WRITE_OUTPUT
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    mapSh
      .getRange(2,1,mapData.length-1,mapHdr.length)
      .setValues(mapData.slice(1));

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    /* =========================
       SUMMARY
    ========================= */

    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Valid=${valid} | Repaired=${repaired} | DurationMs=${durationMs}`
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




// MAPPING ITEM-BRAND-PRODUCT: CLEANUP
/**
 * Script Name: cleanupMapping_Item_Brand_Product_InvalidRows
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Permanently DELETE invalid rows from Mapping_Item_Brand_Product
 * - A row is invalid if:
 *     Item_ID_Machine missing
 *     OR Brand_ID_Machine missing
 *     OR Product_ID_Machine missing
 *
 * Preconditions:
 * - Sheet must exist: Mapping_Item_Brand_Product
 * - Header row present in row 1
 * - Required columns exist
 *
 * Algorithm:
 * 1. Generate Execution_ID
 * 2. Read Mapping_Item_Brand_Product
 * 3. Resolve column indexes
 * 4. Identify rows missing identity columns
 * 5. Delete rows bottom-up
 * 6. Emit execution logs
 */
function cleanupMapping_Item_Brand_Product_InvalidRows() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME  = 'Mapping_Item_Brand_Product';
  const FUNCTION_NAME = 'cleanupMapping_Item_Brand_Product_InvalidRows';
  const TGT_SHEET    = 'Mapping_Item_Brand_Product';

  const MAP_SHEET    = 'Mapping_Item_Brand_Product';

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
      brandId: col('Brand_ID_Machine'),
      productId: col('Product_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
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

      const itemId    = data[i][IDX.itemId];
      const brandId   = data[i][IDX.brandId];
      const productId = data[i][IDX.productId];

      if (!itemId || !brandId || !productId) {
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
        details: 'Missing Item_ID_Machine, Brand_ID_Machine, or Product_ID_Machine'
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

