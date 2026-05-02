/* 
=========================================================
FUNCTION: PRODUCT STAGING POPULATION
=========================================================*/
/**
 * Script Name: populateStagingLookupProducts_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Populate Staging_Lookup_Products with unresolved product canonicals
 * - Create one staging row per unique canonical requiring governance
 * - Initialize rows with default governance state (Review)
 * - Preserve item + brand lineage context
 *
 * PRECONDITIONS:
 * - Sheet exists: Transaction_Resolution
 * - Sheet exists: Staging_Lookup_Products
 * - Header row present (row 1)
 * - Required columns exist in both sheets
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load staging sheet and build canonical dedupe set
 * 2. Load transaction resolution data
 * 3. Iterate transaction rows:
 *    - Skip invalid / resolved / duplicate canonical rows
 *    - Construct staging row
 *    - Append to buffer
 * 4. Batch append staging rows
 * 5. Emit execution summary
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Process rows where:
 *     - Txn_ID_Machine exists
 *     - Product_ID_Machine is empty
 *     - Product_Name_Canonical exists
 *     - Canonical not already staged
 * - Deduplicate using canonical-level set
 * - Generate UUID for Staging_Product_ID_Machine
 * - Write rows in batch (append-only)
 *
 * FAILURE MODES:
 * - Required sheet missing
 * - Required column missing
 */


function populateStagingLookupProducts_FromTransactionResolution() {

  /* --- FUNCTION-LEVEL CONSTANTS --- */
  const SCRIPT_NAME   = 'Products';
  const FUNCTION_NAME = 'populateStagingLookupProducts_FromTransactionResolution';

  const SRC_SHEET = 'Transaction_Resolution';
  const TGT_SHEET = 'Staging_Lookup_Products';

  const t0 = new Date();
  let shouldExit = false; //
  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tsSh = ss.getSheetByName(SRC_SHEET);
    const stgSh = ss.getSheetByName(TGT_SHEET);

    if (!tsSh || !stgSh) {
      throw new Error('Required sheet not found');
    }


    /* --- STEP: PREPARE STAGING CONTEXT --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');

    const stgData = stgSh.getDataRange().getValues();
    const stgHdr = stgData[0];
    const stgCol = n => stgHdr.indexOf(n);

    const IDX_STG = {
      sourceItemName: stgCol('Source_Item_Name'),
      sourceBrandName: stgCol('Source_Brand_Name'),
      entered: stgCol('Product_Name_Entered'),
      canon: stgCol('Product_Name_Canonical'),
      approvedName: stgCol('Product_Name_Approved'),
      adminAction: stgCol('Admin_Action'),
      isApproved: stgCol('Is_Approved'),
      isActive: stgCol('Is_Active'),
      isArchived: stgCol('Is_Archived'),
      isPromoted: stgCol('Is_Lookup_Promoted'),
      populatedAt: stgCol('Populated_At'),
      notes: stgCol('Notes'),
      stagingId: stgCol('Staging_Product_ID_Machine'),
      mappedId: stgCol('Mapped_Product_ID_Machine'),
      sourceTxn: stgCol('Source_Txn_ID_Machine'),
      sourceItem: stgCol('Source_Item_ID_Machine'),
      sourceProduct: stgCol('Source_Product_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX_STG)) {
      if (v === -1) throw new Error(`Staging_Lookup_Products missing column: ${k}`);
    }

    // Build canonical dedupe set (already staged)
    const stagingCanonSet = new Set();
    for (let i = 1; i < stgData.length; i++) {
      const v = stgData[i][IDX_STG.canon];
      if (v) stagingCanonSet.add(String(v));
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');


    /* --- STEP: LOAD TRANSACTION SOURCE --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');

    const tsData = tsSh.getDataRange().getValues();
    const tsHdr = tsData[0];
    const tsCol = n => tsHdr.indexOf(n);

    const IDX = {
      txnId: tsCol('Txn_ID_Machine'),
      itemId: tsCol('Item_ID_Machine'),
      productId: tsCol('Product_ID_Machine'),
      itemName: tsCol('Item_Name_Entered'),
      brandName: tsCol('Brand_Name_Entered'),
      productEntered: tsCol('Product_Name_Entered'),
      productCanon: tsCol('Product_Name_Canonical')
    };

    for (const [k,v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Transaction_Resolution missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');


    if (tsData.length <= 1) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'SRC_Table contains no data rows. Verify!');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }


    /*
    ---------------------------------------------------------
    PROCESS LOOP
    ---------------------------------------------------------*/
    // Scan transactions → extract unresolved product canonicals
    // Apply validation + deduplication
    // Construct staging rows
    
    let scanned = 0;
    let skipNoTxn = 0;
    let skipHasProduct = 0;
    let skipNoCanon = 0;
    let skipDuplicateCanon = 0;

    const rowsToAppend = [];

    for (let i = 1; i < tsData.length; i++) {

      scanned++;

      const r = tsData[i];

      // Skip rows without transaction identity (invalid source)
      if (!r[IDX.txnId]) { skipNoTxn++; continue; }

      // Skip already resolved products
      if (r[IDX.productId]) { skipHasProduct++; continue; }

      const canon = r[IDX.productCanon];

      // Skip rows without canonical (cannot stage)
      if (!canon) { skipNoCanon++; continue; }

      // Skip already staged canonical (dedupe)
      if (stagingCanonSet.has(canon)) { skipDuplicateCanon++; continue; }

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {

        if (rowsToAppend.length > 0) {
          stgSh.getRange(
            stgSh.getLastRow() + 1,
            1,
            rowsToAppend.length,
            stgHdr.length
          ).setValues(rowsToAppend);

          rowsToAppend.length = 0;
          flushLogs_();
        }

        shouldExit = true;
        break;
      }

      // Construct staging row
      const row = new Array(stgHdr.length).fill('');

      row[IDX_STG.sourceItemName] = r[IDX.itemName];
      row[IDX_STG.sourceBrandName] = r[IDX.brandName];

      row[IDX_STG.entered] = r[IDX.productEntered];
      row[IDX_STG.canon] = canon;
      row[IDX_STG.approvedName] = '';

      row[IDX_STG.adminAction] = 'Review';

      row[IDX_STG.isApproved] = false;
      row[IDX_STG.isActive] = false;
      row[IDX_STG.isArchived] = false;
      row[IDX_STG.isPromoted] = false;

      row[IDX_STG.sourceTxn] = r[IDX.txnId];
      row[IDX_STG.sourceItem] = r[IDX.itemId];
      row[IDX_STG.sourceProduct] = r[IDX.productId];

      row[IDX_STG.stagingId] = Utilities.getUuid();
      row[IDX_STG.mappedId] = '';

      row[IDX_STG.populatedAt] = new Date();
      row[IDX_STG.notes] = 'Staged from Transaction_Resolution';

      rowsToAppend.push(row);

      // Log mutation (row-level trace)
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'INFO',
        rowNumber: i + 1,
        action: 'PROCESS',
        stepName: 'WRITE_OUTPUT',
        details: `Txn_ID=${r[IDX.txnId]}, Canonical=${canon}`
      });

      stagingCanonSet.add(canon);
    }


    /* --- STEP: WRITE OUTPUT --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if (rowsToAppend.length > 0) {

      stgSh.getRange(
        stgSh.getLastRow() + 1,
        1,
        rowsToAppend.length,
        stgHdr.length
      ).setValues(rowsToAppend);
    } else {

      /* --- NOTICE — NO INSERT --- */
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'WRITE_OUTPUT',
        'No rows to write (0 inserts)'
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
      `Skipped: NoTxn=${skipNoTxn}, HasProduct=${skipHasProduct}, NoCanon=${skipNoCanon}, Duplicate=${skipDuplicateCanon} | ` +
      `DurationMs=${durationMs}`
    );


    /* --- SCHEDULER EXIT --- */
    if (shouldExit) {
      return exitAndScheduleContinuation_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        { pipelineName: getExecutionContext_()?.pipeline_name }
      );
    }

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
FUNCTION: PRODUCT STATE MACHINE PROCESSOR
=========================================================*/
/**
 * Script Name: processStagingProducts_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Enforce governance state machine on Staging_Lookup_Products
 * - Repair drift between Admin_Action and state flags
 * - Derive all dependent governance columns deterministically
 * - Maintain integrity audit trail
 *
 * PRECONDITIONS:
 * - Sheet exists: Staging_Lookup_Products
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load staging data
 * 2. Iterate rows:
 *    - Resolve Admin_Action → expected state
 *    - Repair drift in flags
 *    - Validate state constraints
 *    - Derive governance outputs
 * 3. Write updated rows back
 * 4. Emit summary + integrity logs
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Map Admin_Action → binary state (Approved / Active / Archived)
 * - Repair mismatches (drift)
 * - Validate:
 *     - Not (Active AND Archived)
 *     - Not (Promoted AND Not Approved)
 * - Derive:
 *     - Is_Pipeline_ready
 *     - Action_Review_Status
 *     - Product_Status
 *     - Entity_Owner
 * - Mark row as VALID / REPAIRED / INVALID
 *
 * FAILURE MODES:
 * - Required sheet missing
 * - Required column missing
 */


function processStagingProducts_StateMachine() {

  /* --- FUNCTION-LEVEL CONSTANTS --- */
  const SCRIPT_NAME   = 'Products';
  const FUNCTION_NAME = 'processStagingProducts_StateMachine';
  const SRC_SHEET     = 'Staging_Lookup_Products';

  const t0 = new Date();
  let shouldExit = false; // 
  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stgSh = ss.getSheetByName(SRC_SHEET);

    if (!stgSh) throw new Error('Staging_Lookup_Products sheet missing');


    /* --- STEP: LOAD DATA --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'LOAD_DATA');

    const data = stgSh.getDataRange().getValues();
    const hdr = data[0];
    const col = n => hdr.indexOf(n);

    const IDX = {
      adminAction: col('Admin_Action'),
      isApproved: col('Is_Approved'),
      isActive: col('Is_Active'),
      isArchived: col('Is_Archived'),
      isPromoted: col('Is_Lookup_Promoted'),
      pipelineReady: col('Is_Pipeline_ready'),
      validState: col('Valid_State'),
      actionStatus: col('Action_Review_Status'),
      productStatus: col('Product_Status'),
      entityOwner: col('Entity_Owner'),
      integrity: col('Integrity_Status'),
      notes: col('Notes'),
      stagingId: col('Staging_Product_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'LOAD_DATA');


    /* --- STEP: STATE MACHINE PROCESSING --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'STATE_MACHINE');

    let repaired = 0;
    let valid = 0;
    let invalid = 0;

    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "EEEE, MMMM d, yyyy 'at' HH:mm:ss"
    );


    /*
    ---------------------------------------------------------
    PROCESS LOOP
    ---------------------------------------------------------*/
    // Apply Admin_Action → governance state mapping
    // Repair drift between expected and actual state
    // Validate state integrity
    // Derive dependent governance fields
    
    for (let i = 1; i < data.length; i++) {

      const row = data[i];
      const admin = row[IDX.adminAction];
      const stagingId = row[IDX.stagingId];

      // Skip rows without Admin_Action (no processing required)
      if (!admin) continue;

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        flushLogs_();
        shouldExit = true;
        break;
      }

      // Map Admin_Action → expected governance state
      let expected = { approved: false, active: false, archived: false };

      switch (admin) {

        case 'Review': 
        break;

        case 'Activate':
          expected.active = true;
          break;

        case 'Approve (UI Hidden)':
          expected.approved = true;
          break;

        case 'Approve & Activate':
          expected.approved = true;
          expected.active = true;
          break;

        case 'Approve but Deprecate':
          expected.approved = true;
          expected.archived = true;
          break;

        case 'Reject':
          expected.archived = true;
          break;

        default:
          // Invalid admin action → mark integrity failure
          invalid++;
          row[IDX.integrity] = 'INVALID_ADMIN_ACTION';
          continue;
      }

      let drift = [];

      // Repair drift in governance flags (expected vs actual)
      function repair(idx, expectedVal, name) {
        if (row[idx] !== expectedVal) {
          drift.push(`${name} expected=${expectedVal} found=${row[idx]}`);
          row[idx] = expectedVal;
        }
      }

      repair(IDX.isApproved, expected.approved, 'Is_Approved');
      repair(IDX.isActive, expected.active, 'Is_Active');
      repair(IDX.isArchived, expected.archived, 'Is_Archived');


      // Validate state constraints (invalid combinations)
      const promoted = row[IDX.isPromoted];

      const validState =
        !(row[IDX.isActive] && row[IDX.isArchived]) &&
        !(promoted && !row[IDX.isApproved]);

      row[IDX.validState] = validState;

      if (!validState) {
        row[IDX.integrity] = 'INVALID_STATE';
        invalid++;
        continue;
      }


      /* --- DERIVE GOVERNANCE FIELDS --- */

      // Pipeline readiness → eligible for promotion
      row[IDX.pipelineReady] =
        row[IDX.isApproved] && !promoted && validState;

      // Review status lifecycle
      row[IDX.actionStatus] =
        promoted ? 'Promoted' :
        row[IDX.isApproved] ? 'Pending (Promotion)' :
        row[IDX.isArchived] ? 'Rejected' :
        'Pending (Approval)';

      // Product status derivation (UI representation)
      let productStatus = 'To be Reviewed';

      if (promoted) {
        if (row[IDX.isActive]) productStatus = 'Promoted (Live)';
        else if (row[IDX.isArchived]) productStatus = 'Promoted (Archived)';
        else productStatus = 'Promoted (Hidden Dropdown)';
      } else {
        if (row[IDX.isArchived] && !row[IDX.isApproved]) productStatus = 'Rejected';
        else if (row[IDX.isActive] && !row[IDX.isApproved]) productStatus = 'Active (Temporary)';
        else if (row[IDX.isApproved] && !row[IDX.isActive]) productStatus = 'Approved (Hidden Dropdown)';
        else if (row[IDX.isApproved] && row[IDX.isActive]) productStatus = 'Approved & Activated (Temporary)';
        else if (row[IDX.isApproved] && row[IDX.isArchived]) productStatus = 'Approved (Archived)';
      }

      row[IDX.productStatus] = productStatus;

      // Ownership transition (Staging → Lookup)
      row[IDX.entityOwner] = promoted ? 'Lookup' : 'Staging';


      // Logging + integrity tagging
      if (drift.length > 0) {

        repaired++;

        const msg =
          `Integrity drift repaired: ${drift.join(' | ')} — ${timestamp}`;

        row[IDX.notes] = msg;
        row[IDX.integrity] = 'REPAIRED';

        // Log drift repair event
        ETI_log_({
          scriptName: SCRIPT_NAME,
          functionName: FUNCTION_NAME,
          sheetName: SRC_SHEET,
          level: 'WARN',
          action: 'PROCESS',
          stepName: 'STATE_MACHINE',
          details: `Row=${i+1}, Staging_ID=${stagingId}, ${drift.join(' | ')}`
        });

      } else {

        valid++;

        row[IDX.integrity] = 'VALID';
        row[IDX.notes] = `Integrity check passed — ${timestamp}`;
      }
    }


    // No-op notice (no drift or invalid state)
    if (repaired === 0 && invalid === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SRC_SHEET,
        'STATE_MACHINE',
        'No drift detected'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'STATE_MACHINE');


    /* --- STEP: WRITE BACK --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'WRITE_BACK');

    stgSh.getRange(2, 1, data.length - 1, hdr.length)
      .setValues(data.slice(1));

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'WRITE_BACK');


    /* --- SUMMARY --- */
    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SRC_SHEET,
      `Valid=${valid}, Repaired=${repaired}, Invalid=${invalid}, DurationMs=${durationMs}`
    );


    /* --- SCHEDULER EXIT --- */
    if (shouldExit) {
      return exitAndScheduleContinuation_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        { pipelineName: getExecutionContext_()?.pipeline_name }
      );
    }

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET);

  } catch (err) {

    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, err, 'MAIN');
    throw err;

  } finally {

    flushLogs_();
  }
}


/* 
=========================================================
FUNCTION: PRODUCT PROMOTION TO LOOKUP
=========================================================*/
/**
 * Script Name: promoteApprovedProducts_FromStaging_ToLookup
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Promote eligible staging products into Lookup_Products
 * - Maintain one-time promotion guarantee
 * - Preserve lineage via Staging_Product_ID_Machine
 * - Update staging row with promotion metadata
 *
 * PRECONDITIONS:
 * - Sheets exist:
 *   - Staging_Lookup_Products
 *   - Lookup_Products
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load lookup and staging data
 * 2. Iterate staging rows:
 *    - Validate promotion eligibility
 *    - Generate Product_ID_Machine
 *    - Create lookup row
 *    - Prepare staging update
 * 3. Batch append lookup rows
 * 4. Apply staging updates
 * 5. Emit summary
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Promote only when:
 *     - Action_Review_Status = Pending (Promotion)
 *     - Is_Pipeline_ready = TRUE
 *     - Is_Lookup_Promoted = FALSE
 * - Resolve name:
 *     - Approved → fallback Entered
 * - Generate UUID for Product_ID_Machine
 * - Write lookup + staging updates
 *
 * FAILURE MODES:
 * - Required sheet missing
 * - Required column missing
 */

function promoteApprovedProducts_FromStaging_ToLookup() {

  /* --- FUNCTION-LEVEL CONSTANTS --- */
  const SCRIPT_NAME   = 'Products';
  const FUNCTION_NAME = 'promoteApprovedProducts_FromStaging_ToLookup';
  const SRC_SHEET     = 'Staging_Lookup_Products';
  const TGT_SHEET     = 'Lookup_Products';

  const t0 = new Date();
  let shouldExit = false;

  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stgSh = ss.getSheetByName(SRC_SHEET);
    const lkSh  = ss.getSheetByName(TGT_SHEET);

    if (!stgSh || !lkSh) throw new Error('Required sheet not found');


    /* --- STEP: LOAD LOOKUP --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_LOOKUP');

    const lkData = lkSh.getDataRange().getValues();
    const lkHdr  = lkData[0];
    const lkCol  = n => lkHdr.indexOf(n);

    const IDX_LK = {
      productName: lkCol('Product_Name'),
      productCanon: lkCol('Product_Name_Canonical'),
      isApproved: lkCol('Is_Approved'),
      isActive: lkCol('Is_Active'),
      isArchived: lkCol('Is_Archived'),
      productStatus: lkCol('Product_Status'),
      isStgPromoted: lkCol('Is_Staging_Promoted'),
      sourceType: lkCol('Source_Type'),
      createdAt: lkCol('Created_At'),
      notes: lkCol('Notes'),
      productIdMachine: lkCol('Product_ID_Machine'),
      stagingId: lkCol('Staging_Product_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX_LK)) {
      if (v === -1) throw new Error(`Lookup_Products missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_LOOKUP');


    /* --- STEP: LOAD STAGING --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');

    const stgData = stgSh.getDataRange().getValues();
    const stgHdr  = stgData[0];
    const stgCol  = n => stgHdr.indexOf(n);

    const IDX_STG = {
      entered: stgCol('Product_Name_Entered'),
      approved: stgCol('Product_Name_Approved'),
      canon: stgCol('Product_Name_Canonical'),
      reviewStatus: stgCol('Action_Review_Status'),
      pipelineReady: stgCol('Is_Pipeline_ready'),
      isPromoted: stgCol('Is_Lookup_Promoted'),
      stagingId: stgCol('Staging_Product_ID_Machine'),
      mappedId: stgCol('Mapped_Product_ID_Machine'),
      notes: stgCol('Notes'),
      isApproved: stgCol('Is_Approved'),
      isActive: stgCol('Is_Active'),
      isArchived: stgCol('Is_Archived'),
      entityOwner: stgCol('Entity_Owner'),
      promotionLabel: stgCol('Promotion_Label'),
      promotedAt: stgCol('Promoted_At'),
      productStatus: stgCol('Product_Status')
    };

    for (const [k,v] of Object.entries(IDX_STG)) {
      if (v === -1) throw new Error(`Staging_Lookup_Products missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');


    /* --- VALIDATION: SOURCE DATA --- */
    if (stgData.length <= 1) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'No data rows found');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }



    /*
    ---------------------------------------------------------
    PROCESS LOOP
    --------------------------------------------------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROMOTION');

    const lookupAppendRows = [];
    const stagingUpdates   = [];

    let scanned = 0;
    let promoted = 0;
    let skipped = 0;

    for (let i = 1; i < stgData.length; i++) {

      scanned++;

      const rowNum = i + 1;
      const r = stgData[i];

      // Eligibility: only rows ready for promotion
      if (r[IDX_STG.reviewStatus] !== 'Pending (Promotion)') { skipped++; continue; }
      if (!r[IDX_STG.pipelineReady]) { skipped++; continue; }
      if (r[IDX_STG.isPromoted] === true) { skipped++; continue; }

      // Resolve final Product name (Approved → fallback Entered)
      const finalName = r[IDX_STG.approved] || r[IDX_STG.entered];
      if (!finalName) { skipped++; continue; }

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        flushLogs_();
        shouldExit = true;
        break;
      }

      const canon = r[IDX_STG.canon] || '';
      const stagingId = r[IDX_STG.stagingId];

      // Generate unique Product_ID_Machine (UUID)
      const productIdMachine = Utilities.getUuid();


      // Construct lookup row (target table)
      const newLookupRow = new Array(lkHdr.length).fill('');

      newLookupRow[IDX_LK.productName] = finalName;
      newLookupRow[IDX_LK.productCanon] = canon;

      newLookupRow[IDX_LK.isApproved] = r[IDX_STG.isApproved];
      newLookupRow[IDX_LK.isActive]   = r[IDX_STG.isActive];
      newLookupRow[IDX_LK.isArchived] = r[IDX_STG.isArchived];
      newLookupRow[IDX_LK.productStatus] = r[IDX_STG.productStatus];
      newLookupRow[IDX_LK.isStgPromoted] = true;
      newLookupRow[IDX_LK.sourceType] = 'STAGING_PROMOTION';

      newLookupRow[IDX_LK.createdAt] = new Date();

      newLookupRow[IDX_LK.productIdMachine] = productIdMachine;
      newLookupRow[IDX_LK.stagingId] = stagingId;

      newLookupRow[IDX_LK.sourceItem] = r[IDX_STG.sourceItem];
      newLookupRow[IDX_LK.sourceBrand] = r[IDX_STG.sourceBrand];

      newLookupRow[IDX_LK.notes] =
        `Promoted from staging → Staging_ID=${stagingId}`;

      lookupAppendRows.push(newLookupRow);


      // Prepare staging update (post-promotion state)
      let promotedStatus = '';

      if (r[IDX_STG.isApproved] && r[IDX_STG.isArchived])
        promotedStatus = 'Promoted (Archived)';
      else if (r[IDX_STG.isApproved] && r[IDX_STG.isActive])
        promotedStatus = 'Promoted (Live)';
      else if (r[IDX_STG.isApproved])
        promotedStatus = 'Promoted (Hidden Dropdown)';

      const existingNote = r[IDX_STG.notes] || '';

      const newNote =
        (existingNote ? existingNote + ' | ' : '') +
        `Promoted to Lookup_Products → Product_ID_Machine=${productIdMachine}`;

      stagingUpdates.push({
        row: rowNum,
        mappedId: productIdMachine,
        note: newNote,
        status: promotedStatus
      });


      // Log promotion (row-level trace)
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'INFO',
        action: 'PROCESS',
        stepName: 'PROMOTION',
        details:
          `Row=${rowNum}, Staging_ID=${stagingId}, Product_ID=${productIdMachine}, Product_Name=${finalName}`
      });

      promoted++;
    }


    // No-op notice (no eligible promotions)
    if (promoted === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'PROMOTION',
        'No products eligible for promotion'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROMOTION');


    /* --- STEP: WRITE LOOKUP --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_LOOKUP');

    if (lookupAppendRows.length > 0) {
      lkSh.getRange(
        lkSh.getLastRow() + 1,
        1,
        lookupAppendRows.length,
        lookupAppendRows[0].length
      ).setValues(lookupAppendRows);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_LOOKUP');


    /* --- STEP: WRITE BACK STAGING --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK_STAGING');

    for (const u of stagingUpdates) {
      stgSh.getRange(u.row, IDX_STG.mappedId + 1).setValue(u.mappedId);
      stgSh.getRange(u.row, IDX_STG.isPromoted + 1).setValue(true);
      stgSh.getRange(u.row, IDX_STG.reviewStatus + 1).setValue('Promoted');
      stgSh.getRange(u.row, IDX_STG.entityOwner + 1).setValue('Lookup');
      stgSh.getRange(u.row, IDX_STG.promotionLabel + 1).setValue('Promoted');
      stgSh.getRange(u.row, IDX_STG.promotedAt + 1).setValue(new Date());
      stgSh.getRange(u.row, IDX_STG.productStatus + 1).setValue(u.status);
      stgSh.getRange(u.row, IDX_STG.notes + 1).setValue(u.note);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK_STAGING');


    /* --- SUMMARY --- */
    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Scanned=${scanned}, Promoted=${promoted}, Skipped=${skipped}, DurationMs=${durationMs}`
    );


    /* --- SCHEDULER EXIT --- */
    if (shouldExit) {
      return exitAndScheduleContinuation_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        { pipelineName: getExecutionContext_()?.pipeline_name }
      );
    }

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
FUNCTION: PRODUCT ID BACKFILL
=========================================================*/
/**
 * Script Name: backfill_ProductIDs_Machine_LookupProducts
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Assign Product_ID_Machine where missing
 *
 * PRECONDITIONS:
 * - Sheet exists: Lookup_Products
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load lookup data
 * 2. Iterate rows:
 *    - If Product_Name exists and ID missing → generate UUID
 * 3. Batch write updates
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Condition:
 *     Product_Name present AND Product_ID_Machine empty
 * - Generate UUID
 * - Write in-memory → batch update
 *
 * FAILURE MODES:
 * - Sheet missing
 * - Column missing
 */


function backfill_ProductIDs_Machine_LookupProducts() {

  /* --- FUNCTION-LEVEL CONSTANTS --- */
  const SCRIPT_NAME  = 'Products';
  const FUNCTION_NAME = 'backfill_ProductIDs_Machine_LookupProducts';
  const SHEET_NAME   = 'Lookup_Products';

  const t0 = new Date();
  let shouldExit = false;


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
    const data  = range.getValues();


    /* --- VALIDATION: DATA PRESENCE --- */
    if (data.length < 2) {

      ETI_logSkip_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SHEET_NAME,
        'No data rows found'
      );

      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);
      return;
    }


    /* --- HEADER RESOLUTION --- */
    const header = data[0];
    const col = n => header.indexOf(n);

    const IDX = {
      productName: col('Product_Name'),
      productIdM:  col('Product_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
    }


    /*
    ---------------------------------------------------------
    PROCESS LOOP
    ---------------------------------------------------------*/
    // Scan lookup rows → assign missing Product_ID_Machine
    // Condition: Product_Name exists AND ID missing
    
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'GENERATE_ID');

    let generatedCount = 0;
    const output = data.map(r => r.slice());

    for (let i = 1; i < output.length; i++) {

      const rowNum   = i + 1;
      const name     = output[i][IDX.productName];
      const productId = output[i][IDX.productIdM];

      // Skip rows without Product_Name or already assigned ID
      if (!(name && !productId)) continue;

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        shouldExit = true;
        break;
      }

      // Generate unique Product_ID_Machine (UUID)
      const newId = Utilities.getUuid();
      output[i][IDX.productIdM] = newId;
      generatedCount++;

      // Log mutation (row-level trace)
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: SHEET_NAME,
        level: 'INFO',
        rowNumber: rowNum,
        action: 'PROCESS',
        stepName: 'GENERATE_ID',
        details: `Generated Product_ID_Machine: ${newId}`
      });
    }


    /* --- NOTICE: NO GENERATION --- */
    if (generatedCount === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SHEET_NAME,
        'GENERATE_ID',
        'No Product_ID generated (all rows already populated or empty)'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'GENERATE_ID');


    /* --- STEP: WRITE_BACK --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'WRITE_BACK');

    // Batch write entire dataset
    range.setValues(output);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'WRITE_BACK');


    /* --- SUMMARY --- */
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SHEET_NAME,
      `Generated=${generatedCount}`
    );


    /* --- SCHEDULER EXIT --- */
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
  ERROR BLOCK
  ============================================*/
  catch (err) {

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SHEET_NAME,
      err,
      'MAIN'
    );

    throw err;
  }


  /*
  ============================================
  FINALIZATION BLOCK
  ============================================*/
  finally {

    flushLogs_();
  }
}


/* 
=========================================================
FUNCTION: PRODUCT ID CLEANUP
=========================================================*/
/**
 * Script Name: cleanupOrphan_ProductIDs_Machine_LookupProducts
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Remove orphan Product_ID_Machine where Product_Name is missing
 *
 * PRECONDITIONS:
 * - Sheet exists: Lookup_Products
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load lookup data
 * 2. Iterate rows:
 *    - If Product_Name missing AND ID exists → clear ID
 * 3. Batch write updates
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Condition:
 *     Product_Name empty AND Product_ID_Machine exists
 * - Clear ID
 *
 * FAILURE MODES:
 * - Sheet missing
 * - Column missing
 */


function cleanupOrphan_ProductIDs_Machine_LookupProducts() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME  = 'Products';
  const FUNCTION_NAME = 'cleanupOrphan_ProductIDs_Machine_LookupProducts';
  const SHEET_NAME   = 'Lookup_Products';

  const t0 = new Date();
  let shouldExit = false;


  /*
  ============================================
  CORE EXECUTION BLOCK
  ============================================
  */
  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error(`Sheet ${SHEET_NAME} not found`);

    const range = sh.getDataRange();
    const data  = range.getValues();


    /* --- VALIDATION: DATA PRESENCE --- */
    if (data.length < 2) {

      ETI_logSkip_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SHEET_NAME,
        'No data rows found'
      );

      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME);
      return;
    }


    /* --- HEADER RESOLUTION --- */
    const header = data[0];
    const col = n => header.indexOf(n);

    const IDX = {
      productName: col('Product_Name'),
      productIdM:  col('Product_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
    }


    /*
    ---------------------------------------------------------
    PROCESS LOOP
    ---------------------------------------------------------
    // Scan lookup rows → clear orphan Product_ID_Machine
    // Condition: Product_Name missing AND ID exists
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'CLEANUP_ORPHAN_ID');

    let clearedCount = 0;
    const output = data.map(r => r.slice());

    for (let i = 1; i < output.length; i++) {

      const rowNum   = i + 1;
      const name     = output[i][IDX.productName];
      const productId = output[i][IDX.productIdM];

      // Skip rows that are not orphan (valid name or no ID)
      if (!( !name && productId )) continue;

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        shouldExit = true;
        break;
      }

      // Orphan detected → clear Product_ID_Machine
      output[i][IDX.productIdM] = '';
      clearedCount++;

      // Log mutation (row-level trace)
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: SHEET_NAME,
        level: 'WARN',
        rowNumber: rowNum,
        action: 'PROCESS',
        stepName: 'CLEANUP_ORPHAN_ID',
        details: `Product_Name missing; Cleared Product_ID_Machine: ${productId}`
      });
    }


    /* --- NOTICE: NO CLEANUP --- */
    if (clearedCount === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SHEET_NAME,
        'CLEANUP_ORPHAN_ID',
        'No orphan Product_ID found'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'CLEANUP_ORPHAN_ID');


    /* --- STEP: WRITE_BACK --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'WRITE_BACK');

    // Batch write entire dataset
    range.setValues(output);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'WRITE_BACK');


    /* --- SUMMARY --- */
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SHEET_NAME,
      `Cleared=${clearedCount}`
    );


    /* --- SCHEDULER EXIT --- */
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
  ERROR BLOCK
  ============================================
  */
  catch (err) {

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SHEET_NAME,
      err,
      'MAIN'
    );

    throw err;
  }


  /*
  ============================================
  FINALIZATION BLOCK
  ============================================
  */
  finally {

    flushLogs_();
  }
}
