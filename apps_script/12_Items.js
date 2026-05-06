/* 
=========================================================
  FUNCTION: POPULATE STAGING ITEMS
========================================================= */
/**
 * Script Name: populateStagingLookupItems_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 *
 * ---------------------------------------------------------
 * PURPOSE
 * ---------------------------------------------------------
 * - Insert staging rows for unresolved item canonicals from Transaction_Resolution
 * - Ensure one staging row per unique canonical value
 *
 * ---------------------------------------------------------
 * PRECONDITIONS
 * ---------------------------------------------------------
 * - Sheet must exist: Transaction_Resolution
 * - Sheet must exist: Staging_Lookup_Items
 * - Required columns in Transaction_Resolution:
 *   - Txn_ID_Machine
 *   - Item_ID_Machine
 *   - Item_Name_Entered
 *   - Item_Name_Canonical
 * - Required columns in Staging_Lookup_Items:
 *   - Source_Txn_ID_Machine
 *   - Staging_Item_ID_Machine
 *   - Item_Name_Entered
 *   - Item_Name_Canonical
 *
 * ---------------------------------------------------------
 * EXECUTION FLOW
 * ---------------------------------------------------------
 * 1. Initialize execution context and logging
 * 2. Load staging sheet and build canonical set
 * 3. Load transaction sheet
 * 4. Validate dataset presence
 * 5. Resolve column indices
 * 6. Start processing loop:
 *    a. Skip rows without Txn_ID_Machine
 *    b. Skip rows with existing Item_ID_Machine
 *    c. Skip rows without Item_Name_Canonical
 *    d. Skip rows already present in staging canonical set
 *    e. Check scheduler timeout
 *    f. Construct staging row
 *    g. Append to in-memory buffer
 * 7. Batch append rows to staging sheet
 * 8. Emit execution summary
 * 9. Handle scheduler continuation (if triggered)
 * 10. End execution
 *
 * ---------------------------------------------------------
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * ---------------------------------------------------------
 * - Iterate Transaction_Resolution rows
 * - Process only rows where:
 *     - Txn_ID_Machine exists
 *     - Item_ID_Machine is empty
 *     - Item_Name_Canonical exists
 *     - Canonical not already staged
 * - For each eligible row:
 *     - Generate Staging_Item_ID_Machine (UUID)
 *     - Populate staging row fields
 * - Maintain in-memory canonical set to avoid duplicates
 * - Perform batch write to staging sheet
 *
 * ---------------------------------------------------------
 * FAILURE MODES
 * ---------------------------------------------------------
 * - Required sheet missing
 * - Required column missing
 */

function populateStagingLookupItems_FromTransactionResolution() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME = 'Items';
  const FUNCTION_NAME = 'populateStagingLookupItems_FromTransactionResolution';
  const SRC_SHEET = 'Transaction_Resolution';
  const TGT_SHEET = 'Staging_Lookup_Items';

  const t0 = new Date(); // Execution timer

  let shouldExit = false; // Scheduler-controlled exit flag (loop-safe, no return inside loop)

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

    if (!tsSh || !stgSh) throw new Error('Required sheet not found');


    /* --- STEP: LOAD_STAGING --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');

    const stgData = stgSh.getDataRange().getValues();
    const stgHdr = stgData[0];
    const stgCol = n => stgHdr.indexOf(n);

    const IDX_STG = {
      sourceTxn: stgCol('Source_Txn_ID_Machine'),
      stagingId: stgCol('Staging_Item_ID_Machine'),
      mappedId: stgCol('Mapped_Item_ID_Machine'),
      entered: stgCol('Item_Name_Entered'),
      canon: stgCol('Item_Name_Canonical'),
      approvedName: stgCol('Item_Name_Approved'),
      adminAction: stgCol('Admin_Action'),
      isApproved: stgCol('Is_Approved'),
      isActive: stgCol('Is_Active'),
      isArchived: stgCol('Is_Archived'),
      isPromoted: stgCol('Is_Lookup_Promoted'),
      populatedAt: stgCol('Populated_At'),
      notes: stgCol('Notes')
    };

    for (const [k,v] of Object.entries(IDX_STG)) {
      if (v === -1) throw new Error(`Missing column: ${k}`);
    }

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
      itemEntered: tsCol('Item_Name_Entered'),
      itemCanon: tsCol('Item_Name_Canonical')
    };

    for (const [k,v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');


    /* --- VALIDATION: SOURCE DATA --- */
    if (tsData.length <= 1) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'No source data');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }


    /* 
    ---------------------------------------------------------
    PROCESS LOOP [STAGE NEW ITEMS]
    --------------------------------------------------------- */
    let scanned = 0;
    let skipNoTxn = 0;
    let skipHasItem = 0;
    let skipNoCanon = 0;
    let skipDuplicateCanon = 0;
    let insertedCount = 0;

    const rowsToAppend = [];

    for (let i = 1; i < tsData.length; i++) {

      scanned++;

      const r = tsData[i];

      // Skip: missing txn
      if (!r[IDX.txnId]) { 
        skipNoTxn++; 
        continue; 
      }

      // Skip: already resolved
      if (r[IDX.itemId]) { 
        skipHasItem++; 
        continue; 
      }

      const canon = r[IDX.itemCanon];

      // Skip: no canonical
      if (!canon) { skipNoCanon++; continue; }

      // Skip: already staged
      if (stagingCanonSet.has(canon)) {
        skipDuplicateCanon++;
        continue;
      }

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

      // Create staging row
      const row = new Array(stgHdr.length).fill('');

      row[IDX_STG.sourceTxn] = r[IDX.txnId];
      row[IDX_STG.stagingId] = Utilities.getUuid();
      row[IDX_STG.entered] = r[IDX.itemEntered];
      row[IDX_STG.canon] = canon;
      row[IDX_STG.adminAction] = 'Review';
      row[IDX_STG.isApproved] = false;
      row[IDX_STG.isActive] = false;
      row[IDX_STG.isArchived] = false;
      row[IDX_STG.isPromoted] = false;
      row[IDX_STG.populatedAt] = new Date();
      row[IDX_STG.notes] = 'Staged from Transaction_Resolution';

      rowsToAppend.push(row);
      insertedCount++;
      stagingCanonSet.add(canon);

      // Log mutation
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'INFO',
        rowNumber: i + 1,
        action: 'PROCESS',
        stepName: 'WRITE_OUTPUT',
        details: `Txn_ID=${r[IDX.txnId]}, Item_Canonical=${canon}`
      });


      /* --- PERIODIC FLUSH --- */
      if (i % 240 === 0 && rowsToAppend.length > 0) {

        stgSh.getRange(
          stgSh.getLastRow() + 1,
          1,
          rowsToAppend.length,
          stgHdr.length
        ).setValues(rowsToAppend);

        rowsToAppend.length = 0;
        flushLogs_();
      }
    }
    

    /* --- STEP: WRITE_OUTPUT --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if (rowsToAppend.length > 0) {
      stgSh.getRange(
        stgSh.getLastRow() + 1,
        1,
        rowsToAppend.length,
        stgHdr.length
      ).setValues(rowsToAppend);

      flushLogs_();
    } 
    else {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'WRITE_OUTPUT',
        'No new items to stage'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');


    /* --- SUMMARY --- */
    const durationMs = new Date().getTime() - t0.getTime();
    const effectiveProcessed = scanned - skipNoTxn - skipHasItem - skipNoCanon - skipDuplicateCanon;

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Scanned=${scanned} | Effective=${effectiveProcessed} | Inserted=${insertedCount} | ` +
      `Skipped: NoTxn=${skipNoTxn}, HasItem=${skipHasItem}, NoCanon=${skipNoCanon}, Duplicate=${skipDuplicateCanon} | ` +
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
  FUNCTION: STAGING STATE MACHINE (ITEM GOVERNANCE)
========================================================= */
/**
 * Script Name: processStagingItems_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * ---------------------------------------------------------
 * PURPOSE
 * ---------------------------------------------------------
 * - Enforce deterministic governance state in Staging_Lookup_Items
 * - Repair drift between Admin_Action and binary flags
 * - Derive dependent governance fields within script
 *
 * ---------------------------------------------------------
 * PRECONDITIONS
 * ---------------------------------------------------------
 * - Sheet must exist: Staging_Lookup_Items
 * - Required columns:
 *   - Admin_Action
 *   - Is_Approved
 *   - Is_Active
 *   - Is_Archived
 *   - Is_Lookup_Promoted
 *   - Is_Pipeline_ready
 *   - Valid_State
 *   - Action_Review_Status
 *   - Item_Status
 *   - Entity_Owner
 *   - Integrity_Status
 *   - Notes
 *   - Staging_Item_ID_Machine
 *
 * ---------------------------------------------------------
 * EXECUTION FLOW
 * ---------------------------------------------------------
 * 1. Initialize execution context and logging
 * 2. Load staging dataset
 * 3. Resolve column indices
 * 4. Start processing loop:
 *    a. Skip rows without Admin_Action
 *    b. Check scheduler timeout
 *    c. Resolve expected binary state from Admin_Action
 *    d. Repair drift in Is_Approved, Is_Active, Is_Archived
 *    e. Validate state constraints (Valid_State)
 *    f. Derive Is_Pipeline_ready
 *    g. Derive Action_Review_Status
 *    h. Derive Item_Status
 *    i. Derive Entity_Owner
 *    j. Log drift repair or mark as valid
 * 5. Write updated dataset back to sheet
 * 6. Emit execution summary
 * 7. Handle scheduler continuation (if triggered)
 * 8. End execution
 *
 * ---------------------------------------------------------
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * ---------------------------------------------------------
 * - Iterate all staging rows
 * - Process only rows where Admin_Action exists
 *
 * - Map Admin_Action to expected binary state:
 *     Review                    → approved=false, active=false, archived=false
 *     Activate                  → approved=false, active=true,  archived=false
 *     Approve (UI Hidden)       → approved=true,  active=false, archived=false
 *     Approve & Activate        → approved=true,  active=true,  archived=false
 *     Approve but Deprecate     → approved=true,  active=false, archived=true
 *     Reject                    → approved=false, active=false, archived=true
 *
 * - For each row:
 *     - Compare actual vs expected flags
 *     - Overwrite mismatched values (drift repair)
 *
 * - Compute Valid_State:
 *     - Invalid if:
 *         (Is_Active AND Is_Archived)
 *         OR
 *         (Is_Lookup_Promoted AND NOT Is_Approved)
 *
 * - Skip further processing if invalid
 *
 * - Compute:
 *     - Is_Pipeline_ready
 *     - Action_Review_Status
 *     - Item_Status
 *     - Entity_Owner
 *
 * - Mark:
 *     - Integrity_Status = REPAIRED (if drift found)
 *     - Integrity_Status = VALID (if no drift)
 *
 * - Write all updates in single batch
 *
 * ---------------------------------------------------------
 * FAILURE MODES
 * ---------------------------------------------------------
 * - Required sheet missing
 * - Required column missing
 */

function processStagingItems_StateMachine() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME = 'Items';
  const FUNCTION_NAME = 'processStagingItems_StateMachine';
  const SRC_SHEET = 'Staging_Lookup_Items';

  const t0 = new Date();
  let shouldExit = false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();


  /*
  ============================================
  CORE EXECUTION BLOCK
  ============================================*/
  try {

    /* --- INITIALIZATION --- */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET);

    const stgSh = ss.getSheetByName(SRC_SHEET);
    if (!stgSh) throw new Error('Staging_Lookup_Items sheet missing');


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
      itemStatus: col('Item_Status'),
      entityOwner: col('Entity_Owner'),
      integrity: col('Integrity_Status'),
      notes: col('Notes'),
      stagingId: col('Staging_Item_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'LOAD_DATA');


    /* --- STEP: DRIFT_REPAIR --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'DRIFT_REPAIR');

    let processed = 0;
    let repaired = 0;
    let valid = 0;
    let invalid = 0;

    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    );


    /*
    ---------------------------------------------------------
    PROCESS LOOP [STATE TRANSITION + DRIFT REPAIR]
    --------------------------------------------------------- */
    for (let i = 1; i < data.length; i++) {

      const row = data[i];
      const admin = row[IDX.adminAction];
      const stagingId = row[IDX.stagingId];

      // Skip rows without admin intent
      if (!admin) continue;

      processed++;

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        flushLogs_();
        shouldExit = true;
        break;
      }

      // Resolve expected state from Admin_Action
      let expected = { approved:false, active:false, archived:false };

      // Map Admin_Action → expected governance state
      switch(admin) {

        case 'Review': break;
        case 'Activate': expected.active = true; break;
        case 'Approve (UI Hidden)': expected.approved = true; break;

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

          ETI_log_({
            scriptName: SCRIPT_NAME,
            functionName: FUNCTION_NAME,
            sheetName: SRC_SHEET,
            level: 'ERROR',
            rowNumber: i + 1,
            action: 'PROCESS',
            stepName: 'DRIFT_REPAIR',
            details: `Staging_ID=${stagingId}, Invalid Admin_Action=${admin}`
          });

          continue;
      }

      let drift = [];

      // Repair drift in binary flags
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

        ETI_log_({
          scriptName: SCRIPT_NAME,
          functionName: FUNCTION_NAME,
          sheetName: SRC_SHEET,
          level: 'ERROR',
          rowNumber: i + 1,
          action: 'PROCESS',
          stepName: 'DRIFT_REPAIR',
          details: `Staging_ID=${stagingId}, Invalid State`
        });

        continue;
      }

      /* --- DERIVE GOVERNANCE FIELDS --- */
      // Pipeline readiness
      row[IDX.pipelineReady] = row[IDX.isApproved] && !promoted && validState;

      // Review status
      row[IDX.actionStatus] = promoted ? 'Promoted' :
      row[IDX.isApproved] ? 'Pending (Promotion)' :
      row[IDX.isArchived] ? 'Rejected' : 'Pending (Approval)';

      // Item status derivation
      let itemStatus = 'To be Reviewed';

      if (promoted) {
        if (row[IDX.isActive]) itemStatus = 'Promoted (Live)';
        else if (row[IDX.isArchived]) itemStatus = 'Promoted (Archived)';
        else itemStatus = 'Promoted (Hidden Dropdown)';
      } 
      else {
        if (row[IDX.isArchived] && !row[IDX.isApproved]) itemStatus = 'Rejected';
        else if (row[IDX.isActive] && !row[IDX.isApproved]) itemStatus = 'Active (Temporary)';
        else if (row[IDX.isApproved] && !row[IDX.isActive]) itemStatus = 'Approved (Hidden Dropdown)';
        else if (row[IDX.isApproved] && row[IDX.isActive]) itemStatus = 'Approved & Activated (Temporary)';
        else if (row[IDX.isApproved] && row[IDX.isArchived]) itemStatus = 'Approved (Archived)';
      }

      row[IDX.itemStatus] = itemStatus;

      row[IDX.entityOwner] = promoted ? 'Lookup' : 'Staging';


      /* --- LOGGING --- */
      if (drift.length > 0) {

        repaired++;

        row[IDX.integrity] = 'REPAIRED';
        row[IDX.notes] = `Drift repaired: ${drift.join(' | ')} - ${timestamp}`;

        ETI_log_({
          scriptName: SCRIPT_NAME,
          functionName: FUNCTION_NAME,
          sheetName: SRC_SHEET,
          level: 'INFO',
          rowNumber: i + 1,
          action: 'PROCESS',
          stepName: 'DRIFT_REPAIR',
          details: `Staging_ID=${stagingId}, ${drift.join(' | ')}`
        });

      } else {
        valid++;
        
        row[IDX.integrity] = 'VALID';
        row[IDX.notes] = `Integrity check passed - ${timestamp}`;
      }
    }


    // No-op notice
    if (repaired === 0 && invalid === 0) {
      ETI_logNotice_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'DRIFT_REPAIR', 'No drift detected');
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'DRIFT_REPAIR');


    /* --- STEP: WRITE_BACK --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'WRITE_BACK');

    if (data.length > 1) {
      stgSh.getRange(2,1,data.length-1,hdr.length).setValues(data.slice(1));
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'WRITE_BACK');


    /* --- SUMMARY --- */
    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SRC_SHEET,
      `Processed=${processed} | Valid=${valid} | Repaired=${repaired} | Invalid=${invalid} | DurationMs=${durationMs}`
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
  }


  /*
  ============================================
  ERROR BLOCK
  ============================================*/
  catch (err) {
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, err, 'MAIN');
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
  FUNCTION: PROMOTE APPROVED ITEMS TO LOOKUP
========================================================= */
/**
 * Script Name: promoteApprovedItems_FromStaging_ToLookup
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * ---------------------------------------------------------
 * PURPOSE
 * ---------------------------------------------------------
 * - Promote approved staging items into Lookup_Items
 * - Assign Item_ID_Machine and persist canonical entity
 * - Update staging rows to reflect promotion completion
 *
 * ---------------------------------------------------------
 * PRECONDITIONS
 * ---------------------------------------------------------
 * - Sheet must exist: Staging_Lookup_Items
 * - Sheet must exist: Lookup_Items
 *
 * - Required columns in Staging_Lookup_Items:
 *   - Item_Name_Entered
 *   - Item_Name_Approved
 *   - Item_Name_Canonical
 *   - Action_Review_Status
 *   - Is_Pipeline_ready
 *   - Is_Lookup_Promoted
 *   - Staging_Item_ID_Machine
 *   - Mapped_Item_ID_Machine
 *   - Notes
 *   - Is_Approved
 *   - Is_Active
 *   - Is_Archived
 *   - Entity_Owner
 *   - Promotion_Label
 *   - Promoted_At
 *   - Item_Status
 *
 * - Required columns in Lookup_Items:
 *   - Item_Name
 *   - Item_Name_Canonical
 *   - Is_Approved
 *   - Is_Active
 *   - Is_Archived
 *   - Is_Staging_Promoted
 *   - Source_Type
 *   - Created_At
 *   - Notes
 *   - Item_ID_Machine
 *   - Staging_ID_Machine
 *
 * ---------------------------------------------------------
 * EXECUTION FLOW
 * ---------------------------------------------------------
 * 1. Initialize execution context and logging
 * 2. Load Lookup_Items dataset
 * 3. Load Staging_Lookup_Items dataset
 * 4. Validate dataset presence
 * 5. Resolve column indices
 * 6. Start processing loop:
 *    a. Skip rows not in 'Pending (Promotion)' state
 *    b. Skip rows where Is_Pipeline_ready is not true
 *    c. Skip rows already promoted
 *    d. Resolve final item name (approved or entered)
 *    e. Skip rows without valid name
 *    f. Check scheduler timeout
 *    g. Generate Item_ID_Machine (UUID)
 *    h. Construct lookup row
 *    i. Queue lookup append
 *    j. Prepare staging update payload
 *    k. Log promotion event
 * 7. Append new rows to Lookup_Items
 * 8. Apply updates to staging rows
 * 9. Emit execution summary
 * 10. Handle scheduler continuation (if triggered)
 * 11. End execution
 *
 * ---------------------------------------------------------
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * ---------------------------------------------------------
 * - Iterate all staging rows
 * - Process only rows where:
 *     - Action_Review_Status = 'Pending (Promotion)'
 *     - Is_Pipeline_ready = true
 *     - Is_Lookup_Promoted = false
 *
 * - For each eligible row:
 *     - Resolve finalName:
 *         Item_Name_Approved if exists else Item_Name_Entered
 *     - Generate Item_ID_Machine (UUID)
 *
 *     - Construct Lookup row:
 *         - Map name, canonical, flags
 *         - Set Is_Staging_Promoted = true
 *         - Set Source_Type = 'STAGING_PROMOTION'
 *         - Set Created_At = current timestamp
 *         - Set Staging_ID_Machine reference
 *
 *     - Prepare staging update:
 *         - Set Mapped_Item_ID_Machine
 *         - Set Is_Lookup_Promoted = true
 *         - Set Action_Review_Status = 'Promoted'
 *         - Set Entity_Owner = 'Lookup'
 *         - Set Promotion_Label = 'Promoted'
 *         - Set Promoted_At timestamp
 *         - Derive Item_Status:
 *             Approved + Active   → Promoted (Live)
 *             Approved + Archived → Promoted (Archived)
 *             Approved only       → Promoted (Hidden Dropdown)
 *
 * - Append all lookup rows in batch
 * - Apply staging updates row-by-row
 *
 * ---------------------------------------------------------
 * FAILURE MODES
 * ---------------------------------------------------------
 * - Required sheet missing
 * - Required column missing
 *
 * */


function promoteApprovedItems_FromStaging_ToLookup() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME  = 'Items';
  const FUNCTION_NAME = 'promoteApprovedItems_FromStaging_ToLookup';
  const SRC_SHEET    = 'Staging_Lookup_Items';
  const TGT_SHEET    = 'Lookup_Items';

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
      itemName: lkCol('Item_Name'),
      itemCanon: lkCol('Item_Name_Canonical'),
      isApproved: lkCol('Is_Approved'),
      isActive: lkCol('Is_Active'),
      isArchived: lkCol('Is_Archived'),
      isStgPromoted: lkCol('Is_Staging_Promoted'),
      sourceType: lkCol('Source_Type'),
      createdAt: lkCol('Created_At'),
      notes: lkCol('Notes'),
      itemIdMachine: lkCol('Item_ID_Machine'),
      stagingId: lkCol('Staging_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX_LK)) {
      if (v === -1) throw new Error(`Lookup_Items missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_LOOKUP');


    /* --- STEP: LOAD_STAGING --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');

    const stgData = stgSh.getDataRange().getValues();
    const stgHdr  = stgData[0];
    const stgCol  = n => stgHdr.indexOf(n);

    const IDX_STG = {
      entered: stgCol('Item_Name_Entered'),
      approved: stgCol('Item_Name_Approved'),
      canon: stgCol('Item_Name_Canonical'),
      reviewStatus: stgCol('Action_Review_Status'),
      pipelineReady: stgCol('Is_Pipeline_ready'),
      isPromoted: stgCol('Is_Lookup_Promoted'),
      stagingId: stgCol('Staging_Item_ID_Machine'),
      mappedId: stgCol('Mapped_Item_ID_Machine'),
      notes: stgCol('Notes'),
      isApproved: stgCol('Is_Approved'),
      isActive: stgCol('Is_Active'),
      isArchived: stgCol('Is_Archived'),
      entityOwner: stgCol('Entity_Owner'),
      promotionLabel: stgCol('Promotion_Label'),
      promotedAt: stgCol('Promoted_At'),
      itemStatus: stgCol('Item_Status')
    };

    for (const [k,v] of Object.entries(IDX_STG)) {
      if (v === -1) throw new Error(`Staging_Lookup_Items missing column: ${k}`);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');


    /* --- VALIDATION: SOURCE DATA --- */
    if (stgData.length <= 1) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'No data rows found');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }


    /* --- STEP: PROMOTION --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROMOTION');


    /* 
    ---------------------------------------------------------
    PROCESS LOOP [PROMOTE APPROVED ITEMS]
    --------------------------------------------------------- */
    let scanned = 0;
    let promoted = 0;
    let skipped = 0;

    const lookupAppendRows = [];
    const stagingUpdates   = [];

    for (let i = 1; i < stgData.length; i++) {

      scanned++;

      const rowNum = i + 1;
      const r = stgData[i];

      // Eligibility checks: rows ready for promotion
      if (r[IDX_STG.reviewStatus] !== 'Pending (Promotion)') { skipped++; continue; }
      if (r[IDX_STG.pipelineReady] !== true) { skipped++; continue; }
      if (r[IDX_STG.isPromoted] === true) { skipped++; continue; }

      const finalName = r[IDX_STG.approved] || r[IDX_STG.entered];
      if (!finalName) { skipped++; continue; }

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {

        if (lookupAppendRows.length > 0) {
          lkSh.getRange(
            lkSh.getLastRow() + 1,
            1,
            lookupAppendRows.length,
            lookupAppendRows[0].length
          ).setValues(lookupAppendRows);
          lookupAppendRows.length = 0;
        }

        if (stagingUpdates.length > 0) {
          for (const u of stagingUpdates) {
            stgSh.getRange(u.row, IDX_STG.mappedId + 1).setValue(u.mappedId);
            stgSh.getRange(u.row, IDX_STG.isPromoted + 1).setValue(true);
            stgSh.getRange(u.row, IDX_STG.reviewStatus + 1).setValue('Promoted');
            stgSh.getRange(u.row, IDX_STG.entityOwner + 1).setValue('Lookup');
            stgSh.getRange(u.row, IDX_STG.promotionLabel + 1).setValue('Promoted');
            stgSh.getRange(u.row, IDX_STG.promotedAt + 1).setValue(new Date());
            stgSh.getRange(u.row, IDX_STG.itemStatus + 1).setValue(u.status);
            stgSh.getRange(u.row, IDX_STG.notes + 1).setValue(u.note);
          }
          stagingUpdates.length = 0;
        }

        flushLogs_();
        shouldExit = true;
        break;
      }

      const canon = r[IDX_STG.canon] || '';
      const stagingId = r[IDX_STG.stagingId];
      
      // Generate unique Item_ID_Machine (UUID)
      const itemIdMachine = Utilities.getUuid();

      // Create lookup row
      const newLookupRow = new Array(lkHdr.length).fill('');

      newLookupRow[IDX_LK.itemName] = finalName;
      newLookupRow[IDX_LK.itemCanon] = canon;
      newLookupRow[IDX_LK.isApproved] = r[IDX_STG.isApproved];
      newLookupRow[IDX_LK.isActive]   = r[IDX_STG.isActive];
      newLookupRow[IDX_LK.isArchived] = r[IDX_STG.isArchived];
      newLookupRow[IDX_LK.isStgPromoted] = true;
      newLookupRow[IDX_LK.sourceType] = 'STAGING_PROMOTION';
      newLookupRow[IDX_LK.createdAt] = new Date();
      newLookupRow[IDX_LK.itemIdMachine] = itemIdMachine;
      newLookupRow[IDX_LK.stagingId] = stagingId;
      newLookupRow[IDX_LK.notes] = `Promoted from staging → Staging_ID=${stagingId}`;

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
        `Promoted to Lookup_Items → Item_ID_Machine=${itemIdMachine}`;

      stagingUpdates.push({
        row: rowNum,
        mappedId: itemIdMachine,
        note: newNote,
        status: promotedStatus
      });

      // Log mutation
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'INFO',
        rowNumber: rowNum,
        action: 'PROCESS',
        stepName: 'PROMOTION',
        details:
          `Staging_ID=${stagingId}, Item_ID=${itemIdMachine}, Item_Name=${finalName}`
      });

      promoted++;

      /* --- PERIODIC FLUSH --- */
      if (i % 240 === 0) {

        if (lookupAppendRows.length > 0) {
          lkSh.getRange(
            lkSh.getLastRow() + 1,
            1,
            lookupAppendRows.length,
            lookupAppendRows[0].length
          ).setValues(lookupAppendRows);
          lookupAppendRows.length = 0;
        }

        if (stagingUpdates.length > 0) {
          for (const u of stagingUpdates) {
            stgSh.getRange(u.row, IDX_STG.mappedId + 1).setValue(u.mappedId);
            stgSh.getRange(u.row, IDX_STG.isPromoted + 1).setValue(true);
            stgSh.getRange(u.row, IDX_STG.reviewStatus + 1).setValue('Promoted');
            stgSh.getRange(u.row, IDX_STG.entityOwner + 1).setValue('Lookup');
            stgSh.getRange(u.row, IDX_STG.promotionLabel + 1).setValue('Promoted');
            stgSh.getRange(u.row, IDX_STG.promotedAt + 1).setValue(new Date());
            stgSh.getRange(u.row, IDX_STG.itemStatus + 1).setValue(u.status);
            stgSh.getRange(u.row, IDX_STG.notes + 1).setValue(u.note);
          }
          stagingUpdates.length = 0;
        }

        flushLogs_();
      }
    }

    if (promoted === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'PROMOTION',
        'No items eligible for promotion'
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

      flushLogs_();
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
      stgSh.getRange(u.row, IDX_STG.itemStatus + 1).setValue(u.status);
      stgSh.getRange(u.row, IDX_STG.notes + 1).setValue(u.note);
    }

    if (stagingUpdates.length > 0) {
      flushLogs_();
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK_STAGING');


    /* --- SUMMARY --- */
    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Scanned=${scanned} | Promoted=${promoted} | Skipped=${skipped} | DurationMs=${durationMs}`
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
  FUNCTION: BACKFILL ITEM IDs (LOOKUP)
========================================================= */
/**
 * Script Name: backfill_ItemIDs_Machine_LookupItems
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * ---------------------------------------------------------
 * PURPOSE
 * ---------------------------------------------------------
 * - Assign Item_ID_Machine to lookup rows where missing
 * - Ensure every valid item has a machine identifier
 *
 * ---------------------------------------------------------
 * PRECONDITIONS
 * ---------------------------------------------------------
 * - Sheet must exist: Lookup_Items
 * - Required columns:
 *   - Item_Name
 *   - Item_ID_Machine
 *
 * ---------------------------------------------------------
 * EXECUTION FLOW
 * ---------------------------------------------------------
 * 1. Initialize execution context and logging
 * 2. Load Lookup_Items dataset
 * 3. Validate dataset presence
 * 4. Resolve column indices
 * 5. Start processing loop:
 *    a. Skip rows where Item_Name is empty
 *    b. Skip rows where Item_ID_Machine already exists
 *    c. Check scheduler timeout
 *    d. Generate Item_ID_Machine (UUID)
 *    e. Assign value in memory
 *    f. Log generation event
 * 6. Write updated dataset back to sheet
 * 7. Emit execution summary
 * 8. Handle scheduler continuation (if triggered)
 * 9. End execution
 *
 * ---------------------------------------------------------
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * ---------------------------------------------------------
 * - Iterate all lookup rows
 * - Process only rows where:
 *     - Item_Name exists
 *     - Item_ID_Machine is empty
 *
 * - For each eligible row:
 *     - Generate UUID using Utilities.getUuid()
 *     - Assign to Item_ID_Machine
 *
 * - Perform all updates in-memory
 * - Write full dataset back in single batch
 *
 * ---------------------------------------------------------
 * FAILURE MODES
 * ---------------------------------------------------------
 * - Required sheet missing
 * - Required column missing
 */

function backfill_ItemIDs_Machine_LookupItems() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME  = 'Items';
  const FUNCTION_NAME = 'backfill_ItemIDs_Machine_LookupItems';
  const TGT_SHEET   = 'Lookup_Items';

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
    const sh = ss.getSheetByName(TGT_SHEET);
    if (!sh) throw new Error(`Sheet ${TGT_SHEET} not found`);

    const range = sh.getDataRange();
    const data  = range.getValues();


    /* --- VALIDATION: DATA PRESENCE --- */
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


    /* --- HEADER RESOLUTION --- */
    const header = data[0];
    const col = n => header.indexOf(n);

    const IDX = {
      itemName: col('Item_Name'),
      itemIdM: col('Item_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
    }


    /* 
    ---------------------------------------------------------
    PROCESS LOOP [GENERATE ITEM IDs]
    --------------------------------------------------------- */
    let generatedCount = 0;
    const output = data.map(r => r.slice());

    for (let i = 1; i < output.length; i++) {

      const rowNum  = i + 1;
      const name    = output[i][IDX.itemName];
      const itemId  = output[i][IDX.itemIdM];

      // Skip rows without Item_Name or already assigned Item_ID_Machine
      if (!(name && !itemId)) continue;

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        shouldExit = true;
        break;
      }

      // Generate ID (business logic)
      const newId = Utilities.getUuid();
      output[i][IDX.itemIdM] = newId;
      generatedCount++;

      // Log mutation
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'INFO',
        rowNumber: rowNum,
        action: 'PROCESS',
        stepName: 'GENERATE_ID',
        details: `Generated Item_ID_Machine: ${newId} | Item=${name}`
      });
    }


    /* --- STEP: WRITE_BACK --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');

    
    // Write full dataset back (single batch write)
    range.setValues(output);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');


    /* --- SUMMARY --- */
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Generated=${generatedCount}`
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
  FUNCTION: CLEANUP ORPHAN ITEM IDs (LOOKUP)
========================================================= */
/**
 * Script Name: cleanupOrphan_ItemIDs_Machine_LookupItems
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * ---------------------------------------------------------
 * PURPOSE
 * ---------------------------------------------------------
 * - Remove Item_ID_Machine values where Item_Name is missing
 * - Enforce integrity: IDs must not exist without a valid item name
 *
 * ---------------------------------------------------------
 * PRECONDITIONS
 * ---------------------------------------------------------
 * - Sheet must exist: Lookup_Items
 * - Required columns:
 *   - Item_Name
 *   - Item_ID_Machine
 *
 * ---------------------------------------------------------
 * EXECUTION FLOW
 * ---------------------------------------------------------
 * 1. Initialize execution context and logging
 * 2. Load Lookup_Items dataset
 * 3. Validate dataset presence
 * 4. Resolve column indices
 * 5. Start processing loop:
 *    a. Skip rows where Item_Name exists
 *    b. Skip rows where Item_ID_Machine is empty
 *    c. Check scheduler timeout
 *    d. Clear Item_ID_Machine in memory
 *    e. Log cleanup event
 * 6. Write updated dataset back to sheet
 * 7. Emit execution summary
 * 8. Handle scheduler continuation (if triggered)
 * 9. End execution
 *
 * ---------------------------------------------------------
 * ALGORITHM (IMPLEMENTATION LOGIC)
 * ---------------------------------------------------------
 * - Iterate all lookup rows
 * - Process only rows where:
 *     - Item_Name is empty
 *     - Item_ID_Machine exists
 *
 * - For each eligible row:
 *     - Set Item_ID_Machine to empty string
 *
 * - Perform all updates in-memory
 * - Write full dataset back in single batch
 *
 * ---------------------------------------------------------
 * FAILURE MODES
 * ---------------------------------------------------------
 * - Required sheet missing
 * - Required column missing
 */


function cleanupOrphan_ItemIDs_Machine_LookupItems() {

  /* --- FUNCTION-LEVEL CONSTANTS & STATE --- */
  const SCRIPT_NAME  = 'Items';
  const FUNCTION_NAME = 'cleanupOrphan_ItemIDs_Machine_LookupItems';
  const TGT_SHEET   = 'Lookup_Items';

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
    const sh = ss.getSheetByName(TGT_SHEET);
    if (!sh) throw new Error(`Sheet ${TGT_SHEET} not found`);

    const range = sh.getDataRange();
    const data  = range.getValues();


    /* --- VALIDATION: DATA PRESENCE --- */
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


    /* --- HEADER RESOLUTION --- */
    const header = data[0];
    const col = n => header.indexOf(n);

    const IDX = {
      itemName: col('Item_Name'),
      itemIdM: col('Item_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
    }


    /* 
    ---------------------------------------------------------
    PROCESS LOOP [CLEAR ORPHAN ITEM IDs]
    --------------------------------------------------------- */
    let clearedCount = 0;
    const output = data.map(r => r.slice());

    for (let i = 1; i < output.length; i++) {

      const rowNum = i + 1;
      const name   = output[i][IDX.itemName];
      const itemId = output[i][IDX.itemIdM];

      // Skip non-orphan rows
      if (!( !name && itemId )) continue;

      /* --- SCHEDULER CHECK --- */
      if (shouldExitForTimeout_(t0)) {
        shouldExit = true;
        break;
      }

      // Orphan = Item_ID exists but Item_Name missing → clear ID
      output[i][IDX.itemIdM] = '';
      clearedCount++;

      // Log mutation
      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'WARN',
        rowNumber: rowNum,
        action: 'PROCESS',
        stepName: 'CLEANUP_ORPHAN_ID',
        details: `Item_Name missing; Cleared Item_ID_Machine: ${itemId}`
      });
    }


    /* --- STEP: WRITE_BACK --- */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');

    range.setValues(output);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');


    /* --- SUMMARY --- */
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Cleared=${clearedCount}`
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

