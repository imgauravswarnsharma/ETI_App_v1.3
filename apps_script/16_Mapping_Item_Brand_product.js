/* 
=========================================================
FUNCTION: ITEM-BRAND-PRODUCT MAPPING DISCOVERY
=========================================================*/
/**
 * Script Name: populateMapping_Item_Brand_Product_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Derive Item–Brand–Product relationships from transactions
 * - Insert only new (first-seen) mapping identities
 * - Maintain transaction-based evidence for mappings
 *
 * PRECONDITIONS:
 * - Sheets exist:
 *   - Transaction_Resolution
 *   - Mapping_Item_Brand_Product
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load existing mapping and build identity set
 * 2. Load transaction data
 * 3. Iterate transactions:
 *    - Validate identity fields
 *    - Capture first occurrence per identity
 * 4. Construct mapping rows
 * 5. Append new rows
 * 6. Emit summary
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Identity key = Item_ID | Brand_ID | Product_ID
 * - Process only rows with all IDs present
 * - Capture first occurrence per execution batch
 * - Skip existing mappings
 * - Default:
 *     Active = TRUE
 *     Analytics = TRUE
 *     Archived = FALSE
 *
 * FAILURE MODES:
 * - Sheet missing
 * - Column missing
 */
function populateMapping_Item_Brand_Product_FromTransactionResolution() {

  /* --- CONSTANTS --- */
  const SCRIPT_NAME  = 'Mapping_Item_Brand_Product';
  const FUNCTION_NAME = 'populateMapping_Item_Brand_Product_FromTransactionResolution';
  const TGT_SHEET   = 'Mapping_Item_Brand_Product';

  const TXN_SHEET = 'Transaction_Resolution';
  const MAP_SHEET = 'Mapping_Item_Brand_Product';

  const t0 = new Date();

  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tsSh  = ss.getSheetByName(TXN_SHEET);
    const mapSh = ss.getSheetByName(MAP_SHEET);

    if (!tsSh || !mapSh) throw new Error('Required sheet not found');


    /* --- STEP: LOAD_MAPPING --- */
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
      if (v === -1) throw new Error(`Missing column: ${k}`);
    }

    // Build identity set from existing mappings
    const existingSet = new Set();

    for (let i = 1; i < mapData.length; i++) {
      const iId = mapData[i][IDX_MAP.itemId];
      const bId = mapData[i][IDX_MAP.brandId];
      const pId = mapData[i][IDX_MAP.productId];

      if (iId && bId && pId) {
        existingSet.add(`${iId}|${bId}|${pId}`);
      }
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_MAPPING');


    /* --- STEP: LOAD_TXN --- */
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

    // Validate required columns
    for (const [k, v] of Object.entries(IDX_TX)) {
      if (v === -1) {
        throw new Error(`Transaction_Resolution missing column: ${k}`);
      }
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');


    /* --- VALIDATION --- */
    if (tsData.length <= 1) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'No data rows');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }


    /*
    ---------------------------------------------------------
    PROCESS LOOP // Scan transactions → extract unique identity triples
    ---------------------------------------------------------*/
    const firstSeenMap = new Map();
    let scanned = 0;

    for (let i = 1; i < tsData.length; i++) {

      scanned++;
      const r = tsData[i];

      // Skip invalid transaction rows
      if (!r[IDX_TX.txnId]) continue;

      // Skip incomplete identity rows
      if (!r[IDX_TX.itemId] || !r[IDX_TX.brandId] || !r[IDX_TX.productId]) continue;

      const key = `${r[IDX_TX.itemId]}|${r[IDX_TX.brandId]}|${r[IDX_TX.productId]}`;

      // Capture only first occurrence
      if (firstSeenMap.has(key)) continue;

      let firstSeenDate =
        r[IDX_TX.txnDateEntered] ||
        r[IDX_TX.createdAt] ||
        new Date();

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


    /* --- BUILD OUTPUT --- */
    const rowsToAppend = [];

    for (const [key, v] of firstSeenMap.entries()) {

      // Skip already existing mappings
      if (existingSet.has(key)) continue;

      const row = new Array(mapHdr.length).fill('');

      row[mapCol('Item_Name_Canonical')]    = v.itemCanon;
      row[mapCol('Brand_Name_Canonical')]   = v.brandCanon;
      row[mapCol('Product_Name_Canonical')] = v.productCanon;

      row[mapCol('Is_Mapping_Active')]    = true;
      row[mapCol('Is_Analytics_Enabled')] = true;
      row[mapCol('Is_Archived')]          = false;

      row[mapCol('Created_At')] = new Date();
      row[mapCol('Notes')]      = 'Discovered from Transaction_Resolution';

      row[mapCol('First_Seen_Txn_Date')] = v.txnDate;
      row[mapCol('First_Seen_Txn_ID')]   = v.txnId;

      row[mapCol('Item_ID_Machine')]    = v.itemId;
      row[mapCol('Brand_ID_Machine')]   = v.brandId;
      row[mapCol('Product_ID_Machine')] = v.productId;

      rowsToAppend.push(row);
      existingSet.add(key);
    }


    /* --- STEP: WRITE_OUTPUT --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if (rowsToAppend.length > 0) {
      mapSh.getRange(
        mapSh.getLastRow() + 1,
        1,
        rowsToAppend.length,
        rowsToAppend[0].length
      ).setValues(rowsToAppend);
    } else {
      ETI_logNotice_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT', 'No new mappings');
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');


    /* --- SUMMARY --- */
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Scanned=${scanned} | Inserted=${rowsToAppend.length}`
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
FUNCTION: ITEM-BRAND-PRODUCT MAPPING STATE RECONCILIATION
=========================================================*/
/**
 * Script Name: processMapping_Item_Brand_Product_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Reconcile mapping rows with Item, Brand, and Product states
 * - Derive snapshot statuses for all three entities
 * - Maintain mapping activation consistency
 *
 * PRECONDITIONS:
 * - Sheets exist:
 *   - Mapping_Item_Brand_Product
 *   - Lookup_Items
 *   - Lookup_Brands
 *   - Lookup_Products
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load mapping and lookup tables
 * 2. Build state maps for Item, Brand, Product
 * 3. Iterate mapping rows:
 *    - Resolve entity states
 *    - Derive snapshot statuses
 *    - Compute mapping activation
 *    - Detect drift
 * 4. Write updates back
 * 5. Emit summary
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Status derived from:
 *     Is_Approved, Is_Active, Is_Archived
 * - Mapping active if:
 *     none of Item/Brand/Product are Archived
 * - Always enforce:
 *     Is_Analytics_Enabled = TRUE
 * - Drift when:
 *     previous activation ≠ derived activation
 *
 * FAILURE MODES:
 * - Sheet missing
 * - Column missing
 */
function processMapping_Item_Brand_Product_StateMachine() {

  const SCRIPT_NAME = 'Mapping_Item_Brand_Product';
  const FUNCTION_NAME = 'processMapping_Item_Brand_Product_StateMachine';
  const TGT_SHEET = 'Mapping_Item_Brand_Product';

  const MAP_SHEET = 'Mapping_Item_Brand_Product';
  const ITEM_SHEET = 'Lookup_Items';
  const BRAND_SHEET = 'Lookup_Brands';
  const PRODUCT_SHEET = 'Lookup_Products';

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
    const prodSh = ss.getSheetByName(PRODUCT_SHEET);

    if (!mapSh || !itemSh || !brandSh || !prodSh) {
      throw new Error('Required sheet missing');
    }

    const mapData = mapSh.getDataRange().getValues();
    const mapHdr = mapData[0];
    const col = n => mapHdr.indexOf(n);

    const IDX = {
      itemId: col('Item_ID_Machine'),
      brandId: col('Brand_ID_Machine'),
      productId: col('Product_ID_Machine'),
      itemStatus: col('Item_Status_Snapshot'),
      brandStatus: col('Brand_Status_Snapshot'),
      productStatus: col('Product_Status_Snapshot'),
      mapActive: col('Is_Mapping_Active'),
      analytics: col('Is_Analytics_Enabled'),
      notes: col('Notes')
    };

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DATA');


    /* --- BUILD STATE MAPS --- */
    function buildStateMap(sh, idColName) {

      const data = sh.getDataRange().getValues();
      const hdr = data[0];
      const c = n => hdr.indexOf(n);

      const IDX_S = {
        id: c(idColName),
        approved: c('Is_Approved'),
        active: c('Is_Active'),
        archived: c('Is_Archived')
      };

      const map = {};

      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        const id = r[IDX_S.id];

        if (!id) continue;

        map[id] = {
          approved: r[IDX_S.approved],
          active: r[IDX_S.active],
          archived: r[IDX_S.archived]
        };
      }

      return map;
    }

    const itemState = buildStateMap(itemSh, 'Item_ID_Machine');
    const brandState = buildStateMap(brandSh, 'Brand_ID_Machine');
    const productState = buildStateMap(prodSh, 'Product_ID_Machine');


    /*
    ---------------------------------------------------------
    PROCESS LOOP
    ---------------------------------------------------------
    // Reconcile mapping rows with entity states
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'DRIFT_REPAIR');

    let repaired = 0;
    let valid = 0;

    function deriveStatus(s) {
      if (!s) return 'Unknown';
      if (s.archived) return 'Archived';
      if (s.active) return 'Active';
      if (s.approved) return 'Approved (Hidden Dropdown)';
      return 'To be Reviewed';
    }

    for (let i = 1; i < mapData.length; i++) {

      const row = mapData[i];

      const item = itemState[row[IDX.itemId]];
      const brand = brandState[row[IDX.brandId]];
      const product = productState[row[IDX.productId]];

      const prevActive = row[IDX.mapActive];

      const itemStatus = deriveStatus(item);
      const brandStatus = deriveStatus(brand);
      const productStatus = deriveStatus(product);

      const newActive =
        !(itemStatus === 'Archived' ||
          brandStatus === 'Archived' ||
          productStatus === 'Archived');

      row[IDX.itemStatus] = itemStatus;
      row[IDX.brandStatus] = brandStatus;
      row[IDX.productStatus] = productStatus;

      row[IDX.mapActive] = newActive;
      row[IDX.analytics] = true;

      if (prevActive !== newActive) {

        repaired++;

        row[IDX.notes] = 'Mapping state updated due to entity status change';

        ETI_log_({
          scriptName: SCRIPT_NAME,
          functionName: FUNCTION_NAME,
          sheetName: TGT_SHEET,
          level: 'WARN',
          action: 'PROCESS',
          stepName: 'DRIFT_REPAIR',
          details: `Row=${i+1}`
        });

      } else {
        valid++;
      }
    }

    if (repaired === 0) {
      ETI_logNotice_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'DRIFT_REPAIR', 'No drift detected');
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'DRIFT_REPAIR');


    /* --- STEP: WRITE_BACK --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');

    mapSh.getRange(2,1,mapData.length-1,mapHdr.length)
         .setValues(mapData.slice(1));

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');


    /* --- SUMMARY --- */
    ETI_logSummary_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET,
      `Valid=${valid}, Repaired=${repaired}, DurationMs=${new Date()-t0}`
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
FUNCTION: ITEM-BRAND-PRODUCT MAPPING CLEANUP
=========================================================*/
/**
 * Script Name: cleanupMapping_Item_Brand_Product_InvalidRows
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Remove invalid mapping rows
 * - Maintain mapping table integrity
 *
 * PRECONDITIONS:
 * - Sheet exists: Mapping_Item_Brand_Product
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load mapping data
 * 2. Identify invalid rows
 * 3. Delete rows in reverse order
 * 4. Emit summary
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Invalid if:
 *     any of Item_ID, Brand_ID, Product_ID missing
 *     OR canonical values missing
 *
 * FAILURE MODES:
 * - Sheet missing
 * - Column missing
 */

function cleanupMapping_Item_Brand_Product_InvalidRows() {

  const SCRIPT_NAME = 'Mapping_Item_Brand_Product';
  const FUNCTION_NAME = 'cleanupMapping_Item_Brand_Product_InvalidRows';
  const SHEET_NAME = 'Mapping_Item_Brand_Product';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t0 = new Date();

  try {

    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error('Sheet missing');

    const data = sh.getDataRange().getValues();
    const hdr = data[0];
    const col = n => hdr.indexOf(n);

    const IDX = {
      itemId: col('Item_ID_Machine'),
      brandId: col('Brand_ID_Machine'),
      productId: col('Product_ID_Machine'),
      itemCanon: col('Item_Name_Canonical'),
      brandCanon: col('Brand_Name_Canonical'),
      productCanon: col('Product_Name_Canonical')
    };

        /*
    ---------------------------------------------------------
    PROCESS LOOP
    ---------------------------------------------------------*/
    
    // Identify invalid mapping rows
    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) {
        throw new Error(`Missing column: ${k}`);
      }
    }

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'CLEANUP');

    let removed = 0;
    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {

      const r = data[i];

      const isInvalid =
        (!r[IDX.itemId] || !r[IDX.brandId] || !r[IDX.productId]) ||
        (!r[IDX.itemCanon] || !r[IDX.brandCanon] || !r[IDX.productCanon]);

      if (!isInvalid) continue;

      rowsToDelete.push(i + 1);
      removed++;
    }

    if (removed === 0) {
      ETI_logNotice_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'CLEANUP', 'No invalid rows');
    }

    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sh.deleteRow(rowsToDelete[i]);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'CLEANUP');


    ETI_logSummary_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME,
      `Removed=${removed}, DurationMs=${new Date()-t0}`
    );

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

  } catch (err) {
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, err, 'MAIN');
    throw err;
  } finally {
    flushLogs_();
  }
}