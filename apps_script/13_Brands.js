/* 
=========================================================
FUNCTION: BRAND STAGING POPULATION
=========================================================*/
/**
 * Script Name: populateStagingLookupBrands_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Populate Staging_Lookup_Brands with unresolved Brand canonicals
 * - Create one staging row per unique canonical requiring governance
 * - Initialize rows with default governance state (Review)
 * - Preserve Item + Product lineage context
 *
 * PRECONDITIONS:
 * - Sheet exists: Transaction_Resolution
 * - Sheet exists: Staging_Lookup_Brands
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
 *     - Brand_ID_Machine is empty
 *     - Brand_Name_Canonical exists
 *     - Canonical not already staged
 * - Use Set() for canonical deduplication
 * - Generate UUID for Staging_Brand_ID_Machine
 * - Write rows in batch (append-only)
 *
 * FAILURE MODES:
 * - Required sheet missing
 * - Required column missing
 */


function populateStagingLookupBrands_FromTransactionResolution() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME = 'Brands';
  const FUNCTION_NAME = 'populateStagingLookupBrands_FromTransactionResolution';

  const SRC_SHEET = 'Transaction_Resolution';
  const TGT_SHEET = 'Staging_Lookup_Brands';

  const t0 = new Date();
  let shouldExit = false;

  /*
  ============================================
  CORE EXECUTION BLOCK
  ============================================*/
  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tsSh = ss.getSheetByName(SRC_SHEET);
    const stgSh = ss.getSheetByName(TGT_SHEET);

    if (!tsSh || !stgSh) {
      throw new Error('Required sheet not found');
    }


    /* --- STEP: LOAD_STAGING --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');

    const stgData = stgSh.getDataRange().getValues();
    const stgHdr = stgData[0];
    const stgCol = n => stgHdr.indexOf(n);

    const IDX_STG = {
      sourceItemName: stgCol('Source_Item_Name'),
      sourceProductName: stgCol('Source_Product_Name'),

      entered: stgCol('Brand_Name_Entered'),
      canon: stgCol('Brand_Name_Canonical'),
      approvedName: stgCol('Brand_Name_Approved'),

      adminAction: stgCol('Admin_Action'),

      isApproved: stgCol('Is_Approved'),
      isActive: stgCol('Is_Active'),
      isArchived: stgCol('Is_Archived'),
      isPromoted: stgCol('Is_Lookup_Promoted'),

      populatedAt: stgCol('Populated_At'),
      notes: stgCol('Notes'),

      stagingId: stgCol('Staging_Brand_ID_Machine'),
      mappedId: stgCol('Mapped_Brand_ID_Machine'),

      sourceTxn: stgCol('Source_Txn_ID_Machine'),
      sourceItem: stgCol('Source_Item_ID_Machine'),
      sourceProduct: stgCol('Source_Product_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX_STG)) {
      if (v === -1) throw new Error(`Staging_Lookup_Brands missing column: ${k}`);
    }

    // Build canonical set for deduplication
    const stagingCanonSet = new Set();

    for (let i = 1; i < stgData.length; i++) {
      const v = stgData[i][IDX_STG.canon];
      if (v) stagingCanonSet.add(String(v));
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');


    /* --- STEP: LOAD_TXN --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');

    const tsData = tsSh.getDataRange().getValues();
    const tsHdr = tsData[0];
    const tsCol = n => tsHdr.indexOf(n);

    const IDX = {
      txnId: tsCol('Txn_ID_Machine'),
      itemId: tsCol('Item_ID_Machine'),
      productId: tsCol('Product_ID_Machine'),
      brandId: tsCol('Brand_ID_Machine'),

      itemName: tsCol('Item_Name_Entered'),
      productName: tsCol('Product_Name_Entered'),

      brandEntered: tsCol('Brand_Name_Entered'),
      brandCanon: tsCol('Brand_Name_Canonical')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) {
        throw new Error(`Transaction_Resolution missing column: ${k}`);
      }
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');


    /* --- VALIDATION: SOURCE DATA --- */
    if (tsData.length <= 1) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'SRC_Table contains no data rows. Verify!');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }


    /* 
    ---------------------------------------------------------
    PROCESS LOOP [STAGE NEW BRANDS]
    --------------------------------------------------------- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'STAGE NEW BRANDS');

    let scanned = 0;
    let skipNoTxn = 0;
    let skipHasBrand = 0;
    let skipNoCanon = 0;
    let skipDuplicateCanon = 0;

    const rowsToAppend = [];

    for (let i = 1; i < tsData.length; i++) {

      scanned++;

      const r = tsData[i];

      // Skip rows without Txn_ID_Machine
      if (!r[IDX.txnId]) {
        skipNoTxn++;
        continue;
      }

      // Skip rows already resolved at Brand level
      if (r[IDX.brandId]) {
        skipHasBrand++;
        continue;
      }

      const canon = r[IDX.brandCanon];

      // Skip rows without Brand canonical
      if (!canon) {
        skipNoCanon++;
        continue;
      }

      // Skip rows already staged (canonical-level dedupe)
      if (stagingCanonSet.has(canon)) {
        skipDuplicateCanon++;
        continue;
      }

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        shouldExit = true;
        break;
      }

      // Construct staging row
      const row = new Array(stgHdr.length).fill('');

      row[IDX_STG.sourceItemName] = r[IDX.itemName];
      row[IDX_STG.sourceProductName] = r[IDX.productName];

      row[IDX_STG.entered] = r[IDX.brandEntered];
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
      stagingCanonSet.add(canon);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'STAGE NEW BRANDS');


    /* --- STEP: WRITE_OUTPUT --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if (rowsToAppend.length > 0) {
      stgSh.getRange(
        stgSh.getLastRow() + 1,
        1,
        rowsToAppend.length,
        stgHdr.length
      ).setValues(rowsToAppend);
    } else {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'WRITE_OUTPUT',
        'No new brands to stage'
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
      `Skipped: NoTxn=${skipNoTxn}, HasBrand=${skipHasBrand}, NoCanon=${skipNoCanon}, Duplicate=${skipDuplicateCanon} | ` +
      `DurationMs=${durationMs}`
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

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

  }


  /*
  ============================================
  ERROR BLOCK
  ============================================*/
  catch (err) {

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
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
FUNCTION: BRAND STATE MACHINE PROCESSOR
=========================================================
*/
/**
 * Script Name: processStagingBrands_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Enforce governance state machine on Staging_Lookup_Brands
 * - Repair drift between Admin_Action and state flags
 * - Derive all dependent governance columns deterministically
 * - Maintain integrity audit trail
 *
 * PRECONDITIONS:
 * - Sheet exists: Staging_Lookup_Brands
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load staging data
 * 2. Iterate rows:
 *    - Resolve expected state from Admin_Action
 *    - Repair drift in flags
 *    - Validate state constraints
 *    - Derive governance outputs
 * 3. Write updated rows back
 * 4. Emit summary + integrity logs
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Map Admin_Action → expected binary state
 * - Repair mismatches (drift)
 * - Validate:
 *     - Not (Active AND Archived)
 *     - Not (Promoted AND Not Approved)
 * - Derive:
 *     - Is_Pipeline_ready
 *     - Action_Review_Status
 *     - Brand_Status
 *     - Entity_Owner
 * - Mark row as VALID / REPAIRED / INVALID
 *
 * FAILURE MODES:
 * - Required sheet missing
 * - Required column missing
 */


function processStagingBrands_StateMachine() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME = 'Brands';
  const FUNCTION_NAME = 'processStagingBrands_StateMachine';
  const SRC_SHEET = 'Staging_Lookup_Brands';

  const t0 = new Date();
  let shouldExit = false;


  /*
  ============================================
  CORE EXECUTION BLOCK
  ============================================*/
  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stgSh = ss.getSheetByName(SRC_SHEET);

    if (!stgSh) throw new Error('Staging_Lookup_Brands sheet missing');


    /* --- STEP: LOAD_DATA --- */
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
      brandStatus: col('Brand_Status'),

      entityOwner: col('Entity_Owner'),
      integrity: col('Integrity_Status'),

      notes: col('Notes'),
      stagingId: col('Staging_Brand_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'LOAD_DATA');


    /* --- STEP: DRIFT_REPAIR --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'DRIFT_REPAIR');

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
    PROCESS LOOP [STATE TRANSITION + DRIFT REPAIR]
    --------------------------------------------------------- */
    for (let i = 1; i < data.length; i++) {

      const row = data[i];
      const admin = row[IDX.adminAction];
      const stagingId = row[IDX.stagingId];

      // Skip rows without Admin_Action
      if (!admin) continue;

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        shouldExit = true;
        break;
      }

      // Map Admin_Action → expected governance state
      let expected = { approved: false, active: false, archived: false };

      switch (admin) {

        case 'Review': break;

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
          invalid++;
          row[IDX.integrity] = 'INVALID_ADMIN_ACTION';
          continue;
      }

      let drift = [];

      // Repair drift in governance flags
      function repair(idx, expectedVal, name) {
        if (row[idx] !== expectedVal) {
          drift.push(`${name} expected=${expectedVal} found=${row[idx]}`);
          row[idx] = expectedVal;
        }
      }

      repair(IDX.isApproved, expected.approved, 'Is_Approved');
      repair(IDX.isActive, expected.active, 'Is_Active');
      repair(IDX.isArchived, expected.archived, 'Is_Archived');


      // Validate state constraints
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

      // Pipeline readiness
      row[IDX.pipelineReady] =
        row[IDX.isApproved] && !promoted && validState;

      // Review status
      row[IDX.actionStatus] =
        promoted ? 'Promoted' :
        row[IDX.isApproved] ? 'Pending (Promotion)' :
        row[IDX.isArchived] ? 'Rejected' :
        'Pending (Approval)';

      // Brand status derivation
      let brandStatus = 'To be Reviewed';

      if (promoted) {
        if (row[IDX.isActive]) brandStatus = 'Promoted (Live)';
        else if (row[IDX.isArchived]) brandStatus = 'Promoted (Archived)';
        else brandStatus = 'Promoted (Hidden Dropdown)';
      } else {
        if (row[IDX.isArchived] && !row[IDX.isApproved]) brandStatus = 'Rejected';
        else if (row[IDX.isActive] && !row[IDX.isApproved]) brandStatus = 'Active (Temporary)';
        else if (row[IDX.isApproved] && !row[IDX.isActive]) brandStatus = 'Approved (Hidden Dropdown)';
        else if (row[IDX.isApproved] && row[IDX.isActive]) brandStatus = 'Approved & Activated (Temporary)';
        else if (row[IDX.isApproved] && row[IDX.isArchived]) brandStatus = 'Approved (Archived)';
      }

      row[IDX.brandStatus] = brandStatus;

      row[IDX.entityOwner] = promoted ? 'Lookup' : 'Staging';


      // Logging + integrity tagging
      if (drift.length > 0) {

        repaired++;

        const msg =
          `Integrity drift repaired: ${drift.join(' | ')} — ${timestamp}`;

        row[IDX.notes] = msg;
        row[IDX.integrity] = 'REPAIRED';

        ETI_log_({
          scriptName: SCRIPT_NAME,
          functionName: FUNCTION_NAME,
          sheetName: SRC_SHEET,
          level: 'WARN',
          action: 'PROCESS',
          stepName: 'DRIFT_REPAIR',
          details: `Row=${i+1}, Staging_ID=${stagingId}, ${drift.join(' | ')}`
        });

      } else {

        valid++;

        row[IDX.integrity] = 'VALID';
        row[IDX.notes] = `Integrity check passed — ${timestamp}`;
      }
    }


    // No-op notice
    if (repaired === 0 && invalid === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SRC_SHEET,
        'DRIFT_REPAIR',
        'No drift detected'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'DRIFT_REPAIR');


    /* --- STEP: WRITE_BACK --- */
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
        {
          pipelineName: getExecutionContext_()?.pipeline_name
        }
      );
    }

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET);

  }


  /*
  ============================================
  ERROR BLOCK
  ============================================*/
  catch (err) {

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SRC_SHEET,
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
FUNCTION: BRAND PROMOTION TO LOOKUP
=========================================================
*/
/**
 * Script Name: promoteApprovedBrands_FromStaging_ToLookup
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Promote approved staging brands into Lookup_Brands
 * - Maintain one-time promotion guarantee
 * - Preserve lineage via Staging_Brand_ID_Machine
 * - Update staging row with promotion metadata
 *
 * PRECONDITIONS:
 * - Sheets exist:
 *   - Staging_Lookup_Brands
 *   - Lookup_Brands
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load lookup and staging data
 * 2. Iterate staging rows:
 *    - Validate promotion eligibility
 *    - Generate Brand_ID_Machine
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
 * - Generate UUID for Brand_ID_Machine
 * - Write lookup + staging updates
 *
 * FAILURE MODES:
 * - Required sheet missing
 * - Required column missing
 */
function promoteApprovedBrands_FromStaging_ToLookup() {

  const SCRIPT_NAME  = 'Brands';
  const FUNCTION_NAME = 'promoteApprovedBrands_FromStaging_ToLookup';
  const SRC_SHEET    = 'Staging_Lookup_Brands';
  const TGT_SHEET    = 'Lookup_Brands';

  const t0 = new Date();
  let shouldExit = false;


  /*
  ============================================
  CORE EXECUTION BLOCK
  ============================================*/
  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const stgSh = ss.getSheetByName(SRC_SHEET);
    const lkSh  = ss.getSheetByName(TGT_SHEET);

    if (!stgSh || !lkSh) throw new Error('Required sheet not found');


    /* --- STEP: LOAD_LOOKUP --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_LOOKUP');

    const lkData = lkSh.getDataRange().getValues();
    const lkHdr  = lkData[0];
    const lkCol  = n => lkHdr.indexOf(n);

    const IDX_LK = {
      brandName: lkCol('Brand_Name'),
      brandCanon: lkCol('Brand_Name_Canonical'),

      isApproved: lkCol('Is_Approved'),
      isActive: lkCol('Is_Active'),
      isArchived: lkCol('Is_Archived'),

      isStgPromoted: lkCol('Is_Staging_Promoted'),
      sourceType: lkCol('Source_Type'),

      createdAt: lkCol('Created_At'),
      notes: lkCol('Notes'),

      brandIdMachine: lkCol('Brand_ID_Machine'),
      stagingId: lkCol('Staging_Brand_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX_LK)) {
      if (v === -1) throw new Error(`Lookup_Brands missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_LOOKUP');


    /* --- STEP: LOAD_STAGING --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');

    const stgData = stgSh.getDataRange().getValues();
    const stgHdr  = stgData[0];
    const stgCol  = n => stgHdr.indexOf(n);

    const IDX_STG = {
      entered: stgCol('Brand_Name_Entered'),
      approved: stgCol('Brand_Name_Approved'),
      canon: stgCol('Brand_Name_Canonical'),

      reviewStatus: stgCol('Action_Review_Status'),
      pipelineReady: stgCol('Is_Pipeline_ready'),
      isPromoted: stgCol('Is_Lookup_Promoted'),

      stagingId: stgCol('Staging_Brand_ID_Machine'),
      mappedId: stgCol('Mapped_Brand_ID_Machine'),

      notes: stgCol('Notes'),

      isApproved: stgCol('Is_Approved'),
      isActive: stgCol('Is_Active'),
      isArchived: stgCol('Is_Archived'),

      entityOwner: stgCol('Entity_Owner'),
      promotionLabel: stgCol('Promotion_Label'),
      promotedAt: stgCol('Promoted_At'),
      brandStatus: stgCol('Brand_Status')
    };

    for (const [k,v] of Object.entries(IDX_STG)) {
      if (v === -1) throw new Error(`Staging_Lookup_Brands missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');


    if (stgData.length <= 1) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'No data rows found');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }


    /*
    ---------------------------------------------------------
    PROCESS LOOP [PROMOTE APPROVED BRANDS]
    --------------------------------------------------------- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROMOTION');

    let scanned = 0;
    let promoted = 0;
    let skipped = 0;

    const lookupAppendRows = [];
    const stagingUpdates   = [];

    for (let i = 1; i < stgData.length; i++) {

      scanned++;

      const rowNum = i + 1;
      const r = stgData[i];

      // Eligibility: only rows ready for promotion
      if (r[IDX_STG.reviewStatus] !== 'Pending (Promotion)') { skipped++; continue; }
      if (!r[IDX_STG.pipelineReady]) { skipped++; continue; }
      if (r[IDX_STG.isPromoted] === true) { skipped++; continue; }

      // Resolve final Brand name
      const finalName = r[IDX_STG.approved] || r[IDX_STG.entered];
      if (!finalName) { skipped++; continue; }

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        shouldExit = true;
        break;
      }

      const canon = r[IDX_STG.canon] || '';
      const stagingId = r[IDX_STG.stagingId];

      // Generate unique Brand_ID_Machine (UUID)
      const brandIdMachine = Utilities.getUuid();

      // Construct lookup row
      const newLookupRow = new Array(lkHdr.length).fill('');

      newLookupRow[IDX_LK.brandName] = finalName;
      newLookupRow[IDX_LK.brandCanon] = canon;

      newLookupRow[IDX_LK.isApproved] = r[IDX_STG.isApproved];
      newLookupRow[IDX_LK.isActive]   = r[IDX_STG.isActive];
      newLookupRow[IDX_LK.isArchived] = r[IDX_STG.isArchived];

      newLookupRow[IDX_LK.isStgPromoted] = true;
      newLookupRow[IDX_LK.sourceType] = 'STAGING_PROMOTION';

      newLookupRow[IDX_LK.createdAt] = new Date();

      newLookupRow[IDX_LK.brandIdMachine] = brandIdMachine;
      newLookupRow[IDX_LK.stagingId] = stagingId;

      newLookupRow[IDX_LK.notes] =
        `Promoted from staging → Staging_ID=${stagingId}`;

      lookupAppendRows.push(newLookupRow);


      // Prepare staging update
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
        `Promoted to Lookup_Brands → Brand_ID_Machine=${brandIdMachine}`;

      stagingUpdates.push({
        row: rowNum,
        mappedId: brandIdMachine,
        note: newNote,
        status: promotedStatus
      });


      // Log promotion event
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'INFO',
        action: 'PROCESS',
        stepName: 'PROMOTION',
        details:
          `Row=${rowNum}, Staging_ID=${stagingId}, Brand_ID=${brandIdMachine}, Brand_Name=${finalName}`
      });

      promoted++;
    }

    // No-op notice
    if (promoted === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'PROMOTION',
        'No brands eligible for promotion'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROMOTION');


    /* --- STEP: WRITE_LOOKUP --- */
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


    /* --- STEP: WRITE_BACK_STAGING --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK_STAGING');

    for (const u of stagingUpdates) {
      stgSh.getRange(u.row, IDX_STG.mappedId + 1).setValue(u.mappedId);
      stgSh.getRange(u.row, IDX_STG.isPromoted + 1).setValue(true);
      stgSh.getRange(u.row, IDX_STG.reviewStatus + 1).setValue('Promoted');
      stgSh.getRange(u.row, IDX_STG.entityOwner + 1).setValue('Lookup');
      stgSh.getRange(u.row, IDX_STG.promotionLabel + 1).setValue('Promoted');
      stgSh.getRange(u.row, IDX_STG.promotedAt + 1).setValue(new Date());
      stgSh.getRange(u.row, IDX_STG.brandStatus + 1).setValue(u.status);
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
        {
          pipelineName: getExecutionContext_()?.pipeline_name
        }
      );
    }

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
  }


  /*
  ============================================
  ERROR BLOCK
  ============================================*/
  catch (err) {
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, err, 'MAIN');
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
FUNCTION: BRAND ID BACKFILL
=========================================================
*/
/**
 * Script Name: backfill_BrandIDs_Machine_LookupBrands
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Assign Brand_ID_Machine where missing
 *
 * PRECONDITIONS:
 * - Sheet exists: Lookup_Brands
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load lookup data
 * 2. Iterate rows:
 *    - If Brand_Name exists and ID missing → generate UUID
 * 3. Batch write updates
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Condition:
 *     Brand_Name present AND Brand_ID_Machine empty
 * - Generate UUID
 * - Write in-memory → batch update
 *
 * FAILURE MODES:
 * - Sheet missing
 * - Column missing
 */
function backfill_BrandIDs_Machine_LookupBrands() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME  = 'Brands';
  const FUNCTION_NAME = 'backfill_BrandIDs_Machine_LookupBrands';
  const SHEET_NAME   = 'Lookup_Brands';

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
      brandName: col('Brand_Name'),
      brandIdM:  col('Brand_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
    }


    /*
    ---------------------------------------------------------
    PROCESS LOOP [GENERATE BRAND IDs]
    --------------------------------------------------------- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'GENERATE_ID');

    let generatedCount = 0;
    const output = data.map(r => r.slice());

    for (let i = 1; i < output.length; i++) {

      const rowNum   = i + 1;
      const name     = output[i][IDX.brandName];
      const brandId  = output[i][IDX.brandIdM];

      // Skip rows without Brand_Name or already assigned Brand_ID_Machine
      if (!(name && !brandId)) continue;

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        shouldExit = true;
        break;
      }

      // Generate unique Brand_ID_Machine (UUID)
      const newId = Utilities.getUuid();
      output[i][IDX.brandIdM] = newId;
      generatedCount++;

      // Log mutation
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: SHEET_NAME,
        level: 'INFO',
        rowNumber: rowNum,
        action: 'PROCESS',
        stepName: 'GENERATE_ID',
        details: `Generated Brand_ID_Machine: ${newId}`
      });
    }


    /* --- NOTICE: NO GENERATION --- */
    if (generatedCount === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SHEET_NAME,
        'GENERATE_ID',
        'No Brand_ID generated (all rows already populated or empty)'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'GENERATE_ID');


    /* --- STEP: WRITE_BACK --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'WRITE_BACK');

    // Write full dataset back (single batch write)
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
FUNCTION: BRAND ID CLEANUP
=========================================================
*/
/**
 * Script Name: cleanupOrphan_BrandIDs_Machine_LookupBrands
 * Script Language: Google Apps Script (JavaScript)
 * App Version: v1.3
 *
 * PURPOSE:
 * - Remove orphan Brand_ID_Machine values
 *
 * PRECONDITIONS:
 * - Sheet exists: Lookup_Brands
 * - Required columns exist
 *
 * =========================================================
 * EXECUTION FLOW
 * =========================================================
 * 1. Load lookup data
 * 2. Iterate rows:
 *    - If Brand_Name missing AND ID exists → clear ID
 * 3. Batch write updates
 *
 * =========================================================
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * =========================================================
 * - Condition:
 *     Brand_Name empty AND Brand_ID_Machine exists
 * - Clear ID
 *
 * FAILURE MODES:
 * - Sheet missing
 * - Column missing
 */


function cleanupOrphan_BrandIDs_Machine_LookupBrands() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME  = 'Brands';
  const FUNCTION_NAME = 'cleanupOrphan_BrandIDs_Machine_LookupBrands';
  const SHEET_NAME   = 'Lookup_Brands';

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
      brandName: col('Brand_Name'),
      brandIdM:  col('Brand_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
    }


    /*
    ---------------------------------------------------------
    PROCESS LOOP [CLEAR ORPHAN BRAND IDs]
    --------------------------------------------------------- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'CLEANUP_ORPHAN_ID');

    let clearedCount = 0;
    const output = data.map(r => r.slice());

    for (let i = 1; i < output.length; i++) {

      const rowNum = i + 1;
      const name   = output[i][IDX.brandName];
      const brandId = output[i][IDX.brandIdM];

      // Skip rows that are not orphan (valid name or no ID)
      if (!( !name && brandId )) continue;

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        shouldExit = true;
        break;
      }

      // Orphan = Brand_ID exists but Brand_Name missing → clear ID
      output[i][IDX.brandIdM] = '';
      clearedCount++;

      // Log mutation
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: SHEET_NAME,
        level: 'WARN',
        rowNumber: rowNum,
        action: 'PROCESS',
        stepName: 'CLEANUP_ORPHAN_ID',
        details: `Brand_Name missing; Cleared Brand_ID_Machine: ${brandId}`
      });
    }


    /* --- NOTICE: NO CLEANUP --- */
    if (clearedCount === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SHEET_NAME,
        'CLEANUP_ORPHAN_ID',
        'No orphan Brand_ID found'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'CLEANUP_ORPHAN_ID');


    /* --- STEP: WRITE_BACK --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SHEET_NAME, 'WRITE_BACK');

    // Write full dataset back (single batch write)
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