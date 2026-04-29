// ITEMS: LOAD INTO STAGING
/**
 * Script Name: populateStagingLookupItems_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Populate Staging_Lookup_Items with unresolved item canonicals detected in Transaction_Resolution.
 * - Insert one staging row per unique canonical value requiring governance review.
 * - Initialize staging rows with the default governance state ("Review").
 * - Serve as the intake bridge between transaction logging and the staging governance workflow.
 *
 * Preconditions:
 * - Sheet must exist: Transaction_Resolution
 * - Sheet must exist: Staging_Lookup_Items
 * - Header row present in row 1 for both sheets
 *
 * Required columns in Transaction_Resolution:
 *   - Txn_ID_Machine
 *   - Item_ID_Machine
 *   - Item_Name_Entered
 *   - Item_Name_Canonical
 *
 * Required columns in Staging_Lookup_Items:
 *   - Source_Txn_ID_Machine
 *   - Staging_Item_ID_Machine
 *   - Mapped_Item_ID_Machine
 *   - Item_Name_Entered
 *   - Item_Name_Canonical
 *   - Item_Name_Approved
 *   - Admin_Action
 *   - Is_Approved
 *   - Is_Active
 *   - Is_Archived
 *   - Is_Lookup_Promoted
 *   - Populated_At
 *   - Notes
 *
 * Algorithm (Step-by-Step):
 *
 * 1. Load Staging_Lookup_Items and read the header row.
 *    - Resolve column indexes dynamically using header names.
 *
 * 2. Build an in-memory Set of existing staged canonicals.
 *    - Ensures only one staging row exists per canonical value.
 *
 * 3. Load Transaction_Resolution rows.
 *    - Resolve required column indexes dynamically.
 *
 * 4. Iterate through Transaction_Resolution rows:
 *
 *    a. Skip rows where Txn_ID_Machine is missing.
 *
 *    b. Skip rows where Item_ID_Machine already exists
 *       (item already resolved in lookup).
 *
 *    c. Skip rows where Item_Name_Canonical is missing.
 *
 *    d. Skip rows where the canonical value already exists
 *       in the staging canonical Set.
 *
 *    e. For each remaining row:
 *       - Construct a new staging row array sized to the
 *         Staging_Lookup_Items header length.
 *
 *       Populate the following fields:
 *
 *       Source_Txn_ID_Machine   ← transaction reference
 *       Staging_Item_ID_Machine ← generated UUID
 *       Mapped_Item_ID_Machine  ← blank
 *       Item_Name_Entered       ← original entered value
 *       Item_Name_Canonical     ← canonical value
 *       Item_Name_Approved      ← blank
 *
 *       Admin_Action            ← "Review"
 *       Is_Approved             ← FALSE
 *       Is_Active               ← FALSE
 *       Is_Archived             ← FALSE
 *       Is_Lookup_Promoted      ← FALSE
 *
 *       Populated_At            ← script timestamp
 *       Notes                   ← "Staged from Transaction_Resolution"
 *
 *       All other staging columns remain blank so that
 *       sheet formulas populate derived values.
 *
 *    f. Add the canonical value to the in-memory Set to
 *       prevent duplicate staging rows within the same run.
 *
 * 5. Batch append all new rows to Staging_Lookup_Items.
 *
 * 6. Emit execution summary metrics to console logs
 *    and ETI_log_ for auditability.
 *
 * Failure Modes:
 * - Required sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */
function populateStagingLookupItems_FromTransactionResolution() {

  const SCRIPT_NAME = 'Items';
  const FUNCTION_NAME = 'populateStagingLookupItems_FromTransactionResolution';
  const SRC_SHEET = 'Transaction_Resolution';
  const TGT_SHEET = 'Staging_Lookup_Items';

  const t0 = new Date();

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tsSh = ss.getSheetByName(SRC_SHEET);
    const stgSh = ss.getSheetByName(TGT_SHEET);

    if (!tsSh || !stgSh) {
      throw new Error('Required sheet not found');
    }

    /* =========================
       STEP — LOAD_STAGING
    ========================= */
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
      if (v === -1) throw new Error(`Staging_Lookup_Items missing column: ${k}`);
    }

    const stagingCanonSet = new Set();

    for (let i = 1; i < stgData.length; i++) {
      const v = stgData[i][IDX_STG.canon];
      if (v) stagingCanonSet.add(String(v));
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_STAGING');

    /* =========================
       STEP — LOAD_TXN
    ========================= */

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
      if (v === -1) {
        throw new Error(`Transaction_Resolution missing column: ${k}`);
      }
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_TXN');

    /* =====================================
       STEP:NOTICE & FUNCTION:SKIP — NO SRC DATA
    ========================================= */

    if (tsData.length <= 1) {
      ETI_logSkip_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'SRC_Table contains no data rows. Verify!'
      );

      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }

    /* =========================
       PROCESS LOOP
    ========================= */

    let scanned = 0;
    let skipNoTxn = 0;
    let skipHasItem = 0;
    let skipNoCanon = 0;
    let skipDuplicateCanon = 0;

    const rowsToAppend = [];

    for (let i = 1; i < tsData.length; i++) {

      scanned++;

      const r = tsData[i];

      if (!r[IDX.txnId]) {
        skipNoTxn++;
        continue;
      }

      if (r[IDX.itemId]) {
        skipHasItem++;
        continue;
      }

      const canon = r[IDX.itemCanon];

      if (!canon) {
        skipNoCanon++;
        continue;
      }

      if (stagingCanonSet.has(canon)) {
        skipDuplicateCanon++;
        continue;
      }

      const row = new Array(stgHdr.length).fill('');

      row[IDX_STG.sourceTxn] = r[IDX.txnId];
      row[IDX_STG.stagingId] = Utilities.getUuid();
      row[IDX_STG.mappedId] = '';
      row[IDX_STG.entered] = r[IDX.itemEntered];
      row[IDX_STG.canon] = canon;
      row[IDX_STG.approvedName] = '';

      row[IDX_STG.adminAction] = 'Review';
      row[IDX_STG.isApproved] = false;
      row[IDX_STG.isActive] = false;
      row[IDX_STG.isArchived] = false;
      row[IDX_STG.isPromoted] = false;

      row[IDX_STG.populatedAt] = new Date();
      row[IDX_STG.notes] = 'Staged from Transaction_Resolution';

      rowsToAppend.push(row);

      /* =========================
         LOG PER ROW (MUTATION)
      ========================= */
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

    /* =========================
       STEP — WRITE_OUTPUT
    ========================= */

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
        'No rows to write (0 inserts)'
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
      `Skipped: NoTxn=${skipNoTxn}, HasItem=${skipHasItem}, NoCanon=${skipNoCanon}, Duplicate=${skipDuplicateCanon} | ` +
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

    flushLogs_();  // ---- CRITICAL: ensures logs persist even in standalone execution ----

  }
}





// ITEM: PROCESS STAGED ITEM
/**
 * Script Name: processStagingItems_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Deterministically process governance state for Staging_Lookup_Items.
 * - Repair any drift between Admin_Action and binary state flags.
 * - Derive all dependent columns directly in script (no formula reliance).
 * - Preserve audit trace of drift in Notes and ETI_log_.
 *
 * Governance State Machine (Authoritative)
 *
 * Admin_Action → Binary Flags → Derived State
 *
 * Admin_Action                Is_Approved   Is_Active   Is_Archived
 * ------------------------------------------------------------------
 * Review                      FALSE         FALSE       FALSE
 * Activate                    FALSE         TRUE        FALSE
 * Approve (UI Hidden)         TRUE          FALSE       FALSE
 * Approve & Activate          TRUE          TRUE        FALSE
 * Approve but Deprecate       TRUE          FALSE       TRUE
 * Reject                      FALSE         FALSE       TRUE
 *
 *
 * Post Promotion (handled by promotion script)
 *
 * Is_Approved   Is_Active   Is_Archived   Is_Lookup_Promoted
 * -----------------------------------------------------------
 * TRUE          FALSE       TRUE          TRUE   → Promoted (Archived)
 * TRUE          FALSE       FALSE         TRUE   → Promoted (Hidden Dropdown)
 * TRUE          TRUE        FALSE         TRUE   → Promoted (Live)
 *
 *
 * Derived Columns
 *
 * Valid_State =
 * NOT(
 *   (Is_Active = TRUE AND Is_Archived = TRUE)
 *   OR
 *   (Is_Lookup_Promoted = TRUE AND Is_Approved = FALSE)
 * )
 *
 *
 * Is_Pipeline_ready =
 * Is_Approved
 * AND NOT Is_Lookup_Promoted
 * AND Valid_State
 *
 *
 * Action_Review_Status
 *
 * if Is_Lookup_Promoted → Promoted
 * else if Is_Approved → Pending (Promotion)
 * else if Is_Archived → Rejected
 * else → Pending (Approval)
 *
 *
 * Item_Status
 *
 * If promoted:
 *   Approved + Active → Promoted (Live)
 *   Approved + Archived → Promoted (Archived)
 *   Approved only → Promoted (Hidden Dropdown)
 *
 * If not promoted:
 *   none → To be Reviewed
 *   active only → Active (Temporary)
 *   approved only → Approved (Hidden Dropdown)
 *   approved + active → Approved & Activated (Temporary)
 *   approved + archived → Approved (Archived)
 *   archived only → Rejected
 *
 *
 * Entity_Owner
 *
 * If Is_Lookup_Promoted → Lookup
 * Else → Staging
 *
 *
 * Integrity_Status
 *
 * VALID
 * REPAIRED
 * INVALID_STATE
 *
 *
 * Preconditions
 * - Sheet must exist: Staging_Lookup_Items
 * - Sheet must exist: Automation_Control
 * - Integrity control switch: Automation_Control!K2
 *
 * Failure Modes
 * - Required sheet missing
 * - Required column missing
 */
function processStagingItems_StateMachine() {

  const SCRIPT_NAME = 'Items';
  const FUNCTION_NAME = 'processStagingItems_StateMachine';
  const SRC_SHEET = 'Staging_Lookup_Items';

  const t0 = new Date();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET);

    const stgSh = ss.getSheetByName(SRC_SHEET);
    if (!stgSh) throw new Error('Staging_Lookup_Items sheet missing');

    /* =========================
       STEP — LOAD_DATA
    ========================= */

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

    /* =========================
       STEP — DRIFT_REPAIR
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'DRIFT_REPAIR');

    let repaired = 0;
    let valid = 0;
    let invalid = 0;

    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "EEEE, MMMM d, yyyy 'at' HH:mm:ss"
    );

    /* =========================
       PROCESS LOOP
    ========================= */

    for (let i = 1; i < data.length; i++) {

      const row = data[i];
      const admin = row[IDX.adminAction];
      const stagingId = row[IDX.stagingId];

      if (!admin) continue;

      let expected = {
        approved:false,
        active:false,
        archived:false
      };

      switch(admin) {

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
          invalid++;
          row[IDX.integrity] = 'INVALID_ADMIN_ACTION';
          continue;
      }

      let drift = [];

      function repair(idx, expectedVal, name) {

        const actual = row[idx];

        if (actual !== expectedVal) {
          drift.push(`${name} expected=${expectedVal} found=${actual}`);
          row[idx] = expectedVal;
        }
      }

      repair(IDX.isApproved, expected.approved, 'Is_Approved');
      repair(IDX.isActive, expected.active, 'Is_Active');
      repair(IDX.isArchived, expected.archived, 'Is_Archived');

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

      const pipelineReady =
        row[IDX.isApproved] &&
        !promoted &&
        validState;

      row[IDX.pipelineReady] = pipelineReady;

      let reviewStatus = 'Pending (Approval)';

      if (promoted) reviewStatus = 'Promoted';
      else if (row[IDX.isApproved]) reviewStatus = 'Pending (Promotion)';
      else if (row[IDX.isArchived]) reviewStatus = 'Rejected';

      row[IDX.actionStatus] = reviewStatus;

      let itemStatus = 'To be Reviewed';

      if (promoted) {

        if (row[IDX.isActive])
          itemStatus = 'Promoted (Live)';
        else if (row[IDX.isArchived])
          itemStatus = 'Promoted (Archived)';
        else
          itemStatus = 'Promoted (Hidden Dropdown)';

      } else {

        if (row[IDX.isArchived] && !row[IDX.isApproved])
          itemStatus = 'Rejected';

        else if (row[IDX.isActive] && !row[IDX.isApproved])
          itemStatus = 'Active (Temporary)';

        else if (row[IDX.isApproved] && !row[IDX.isActive])
          itemStatus = 'Approved (Hidden Dropdown)';

        else if (row[IDX.isApproved] && row[IDX.isActive])
          itemStatus = 'Approved & Activated (Temporary)';

        else if (row[IDX.isApproved] && row[IDX.isArchived])
          itemStatus = 'Approved (Archived)';
      }

      row[IDX.itemStatus] = itemStatus;

      row[IDX.entityOwner] =
        promoted ? 'Lookup' : 'Staging';

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
        row[IDX.notes] =
          `Integrity check passed — ${timestamp}`;
      }
    }

    /* =========================
       NOTICE — NO DRIFT / NO INVALID
    ========================= */

    if (repaired === 0 && invalid === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        SRC_SHEET,
        'DRIFT_REPAIR',
        'No drift detected; all rows valid'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'DRIFT_REPAIR');

    /* =========================
       STEP — WRITE_BACK
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'WRITE_BACK');

    /* =========================
       WRITE BACK
    ========================= */

    stgSh
      .getRange(2,1,data.length-1,hdr.length)
      .setValues(data.slice(1));

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET, 'WRITE_BACK');

    /* =========================
       SUMMARY
    ========================= */

    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SRC_SHEET,
      `Valid=${valid}, Repaired=${repaired}, Invalid=${invalid}, DurationMs=${durationMs}`
    );

    /* =========================
       END
    ========================= */

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, SRC_SHEET);

  } catch (err) {

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      SRC_SHEET,
      err,
      'MAIN'
    );

    throw err;

  } finally {

    flushLogs_();

  }
}





// ITEMS: PROMOTE ACTIONED ITEMS TO LOOKUP
/**
 * Script Name: promoteApprovedItems_FromStaging_ToLookup
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Version: v1.3.1
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Promote staging items into Lookup_Items once governance conditions are satisfied
 * - Maintain deterministic promotion authority
 * - Preserve lineage via Staging_Item_ID_Machine
 * - Provide full action-level logging for debugging
 *
 * Promotion Gate (ALL must be true):
 *
 *   Action_Review_Status = "Pending (Promotion)"
 *   Is_Pipeline_ready    = TRUE
 *   Is_Lookup_Promoted   = FALSE
 *
 * Promotion Results:
 *
 * Lookup_Items row inserted
 * Staging row updated with:
 *
 *   Mapped_Item_ID_Machine
 *   Is_Lookup_Promoted
 *   Action_Review_Status
 *   Entity_Owner
 *   Promotion_Label
 *   Promoted_At
 *   Item_Status
 *   Notes
 */
function promoteApprovedItems_FromStaging_ToLookup() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME  = 'Items';
  const FUNCTION_NAME = 'promoteApprovedItems_FromStaging_ToLookup';
  const SRC_SHEET    = 'Staging_Lookup_Items';
  const TGT_SHEET    = 'Lookup_Items';

  const t0 = new Date();

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const stgSh = ss.getSheetByName(SRC_SHEET);
    const lkSh  = ss.getSheetByName(TGT_SHEET);

    if (!stgSh || !lkSh) {
      throw new Error('Required sheet not found');
    }

    /* =========================
       STEP — LOAD_LOOKUP
    ========================= */
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

    /* =========================
       STEP — LOAD_STAGING
    ========================= */
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

    /* =========================
       SKIP — NO DATA
    ========================= */
    if (stgData.length <= 1) {
      ETI_logSkip_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'No data rows found');
      ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);
      return;
    }

    /* =========================
       STEP — PROMOTION
    ========================= */
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

      if (r[IDX_STG.reviewStatus] !== 'Pending (Promotion)') { skipped++; continue; }
      if (r[IDX_STG.pipelineReady] !== true) { skipped++; continue; }
      if (r[IDX_STG.isPromoted] === true) { skipped++; continue; }

      const finalName = r[IDX_STG.approved] || r[IDX_STG.entered];
      if (!finalName) { skipped++; continue; }

      const canon = r[IDX_STG.canon] || '';
      const stagingId = r[IDX_STG.stagingId];

      const itemIdMachine = Utilities.getUuid();

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
      newLookupRow[IDX_LK.notes] =
        `Promoted from staging → Staging_ID=${stagingId}`;

      lookupAppendRows.push(newLookupRow);

      /* =========================
         DERIVE PROMOTED STATUS
      ========================= */

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

      ETI_log_({
        scriptName: SCRIPT_NAME,
        functionName: FUNCTION_NAME,
        sheetName: TGT_SHEET,
        level: 'INFO',
        action: 'PROCESS',
        stepName: 'PROMOTION',
        details:
          `Row=${rowNum}, Staging_ID=${stagingId}, Item_ID=${itemIdMachine}, Item_Name=${finalName}`
      });

      promoted++;
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

    /* =========================
       STEP — WRITE_LOOKUP
    ========================= */
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

    /* =========================
       STEP — WRITE_BACK_STAGING
    ========================= */
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

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK_STAGING');

    /* =========================
       SUMMARY
    ========================= */
    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Scanned=${scanned}, Promoted=${promoted}, Skipped=${skipped}, DurationMs=${durationMs}`
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





// ITEMS: LOOKUP - BACKFILL BACKEND-SEEDED(MANUALLY) IDs FOR ITEM RECORDS
function backfill_ItemIDs_Machine_LookupItems() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME  = 'Items';
  const FUNCTION_NAME = 'backfill_ItemIDs_Machine_LookupItems';
  const TGT_SHEET   = 'Lookup_Items';

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(TGT_SHEET);
    if (!sh) throw new Error(`Sheet ${TGT_SHEET} not found`);

    const range = sh.getDataRange();
    const data  = range.getValues();

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
      itemName: col('Item_Name'),
      itemIdM: col('Item_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
    }

    /* =========================
       PROCESS LOOP
    ========================= */
    let generatedCount = 0;
    const output = data.map(r => r.slice());

    for (let i = 1; i < output.length; i++) {

      const rowNum  = i + 1;
      const name    = output[i][IDX.itemName];
      const itemId  = output[i][IDX.itemIdM];

      if (name && !itemId) {

        const newId = Utilities.getUuid();
        output[i][IDX.itemIdM] = newId;
        generatedCount++;

        /* =========================
           LOG PER ROW
        ========================= */
        ETI_log_({
          scriptName: SCRIPT_NAME,
          functionName: FUNCTION_NAME,
          sheetName: TGT_SHEET,
          level: 'INFO',
          rowNumber: rowNum,
          action: 'PROCESS',
          stepName: 'GENERATE_ID',
          details: `Generated Item_ID_Machine: ${newId}`
        });
      }
    }

    /* =========================
       WRITE BACK
    ========================= */
    range.setValues(output);

    /* =========================
       SUMMARY
    ========================= */
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Generated=${generatedCount}`
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





// ITEMS: LOOKUP - CLEANUP OF INVALID RECORDS
/**
 * Script Name: cleanupOrphan_ItemIDs_Machine_LookupItems
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Clear orphan Item_ID_Machine where Item_Name is missing
 *
 * Preconditions:
 * - Spreadsheet contains a sheet named: Lookup_Items
 * - Header row present in row 1
 * - Required columns:
 *   - Item_Name
 *   - Item_ID_Machine
 *
 * Algorithm:
 * 1. Generate Execution_ID
 * 2. Load Lookup_Items
 * 3. Resolve column indexes
 * 4. For each row:
 *    a. If Item_Name is blank AND Item_ID_Machine exists → clear ID
 * 5. Write changes in one batch
 *
 * Failure Modes:
 * - Sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */
function cleanupOrphan_ItemIDs_Machine_LookupItems() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME  = 'Items';
  const FUNCTION_NAME = 'cleanupOrphan_ItemIDs_Machine_LookupItems';
  const TGT_SHEET   = 'Lookup_Items';

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(TGT_SHEET);
    if (!sh) throw new Error(`Sheet ${TGT_SHEET} not found`);

    const range = sh.getDataRange();
    const data  = range.getValues();

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
      itemName: col('Item_Name'),
      itemIdM: col('Item_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
    }

    /* =========================
       PROCESS LOOP
    ========================= */
    let clearedCount = 0;
    const output = data.map(r => r.slice());

    for (let i = 1; i < output.length; i++) {

      const rowNum = i + 1;
      const name   = output[i][IDX.itemName];
      const itemId = output[i][IDX.itemIdM];

      if (!name && itemId) {

        output[i][IDX.itemIdM] = '';
        clearedCount++;

        /* =========================
           LOG PER ROW
        ========================= */
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
    }

    /* =========================
       WRITE BACK
    ========================= */
    range.setValues(output);

    /* =========================
       SUMMARY
    ========================= */
    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Cleared=${clearedCount}`
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