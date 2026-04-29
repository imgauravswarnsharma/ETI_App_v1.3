// PRODUCTS: LOAD INTO STAGING
/**
 * Script Name: populateStagingLookupProducts_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Populate Staging_Lookup_Products with unresolved product canonicals
 *   detected in Transaction_Resolution.
 * - Insert one staging row per unique canonical product requiring governance review.
 * - Initialize staging rows with the default governance state ("Review").
 * - Include item and brand context for governance visibility.
 *
 * Preconditions:
 * - Sheet must exist: Transaction_Resolution
 * - Sheet must exist: Staging_Lookup_Products
 * - Header row present in row 1
 *
 * Required columns in Transaction_Resolution:
 *   - Txn_ID_Machine
 *   - Item_ID_Machine
 *   - Product_ID_Machine
 *   - Item_Name_Entered
 *   - Brand_Name_Entered
 *   - Product_Name_Entered
 *   - Product_Name_Canonical
 *
 * Required columns in Staging_Lookup_Products:
 *   - Source_Item_Name
 *   - Source_Brand_Name
 *   - Product_Name_Entered
 *   - Product_Name_Canonical
 *   - Product_Name_Approved
 *   - Admin_Action
 *   - Is_Approved
 *   - Is_Active
 *   - Is_Archived
 *   - Is_Lookup_Promoted
 *   - Populated_At
 *   - Notes
 *   - Staging_Product_ID_Machine
 *   - Mapped_Product_ID_Machine
 *   - Source_Txn_ID_Machine
 *   - Source_Item_ID_Machine
 *   - Source_Product_ID_Machine
 *
 * Algorithm (Step-by-Step):
 *
 * 1. Load Staging_Lookup_Products and read header row.
 * 2. Build in-memory Set of existing staged product canonicals.
 * 3. Load Transaction_Resolution rows.
 * 4. Iterate rows:
 *
 *    Skip if:
 *      - Txn_ID_Machine missing
 *      - Product_ID_Machine exists
 *      - Product_Name_Canonical missing
 *      - canonical already staged
 *
 * 5. Create staging row:
 *
 *    Source_Item_Name
 *    Source_Brand_Name
 *
 *    Product_Name_Entered
 *    Product_Name_Canonical
 *    Product_Name_Approved
 *
 *    Admin_Action = Review
 *
 *    Is_Approved = FALSE
 *    Is_Active = FALSE
 *    Is_Archived = FALSE
 *    Is_Lookup_Promoted = FALSE
 *
 *    Source_Txn_ID_Machine
 *    Source_Item_ID_Machine
 *    Source_Product_ID_Machine
 *
 *    Staging_Product_ID_Machine (UUID)
 *
 *    Populated_At
 *    Notes
 *
 * 6. Batch append rows.
 * 7. Emit execution summary.
 *
 * Failure Modes:
 * - Required sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */
function populateStagingLookupProducts_FromTransactionResolution() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME   = 'Products';
  const FUNCTION_NAME = 'populateStagingLookupProducts_FromTransactionResolution';

  const SRC_SHEET = 'Transaction_Resolution';
  const TGT_SHEET = 'Staging_Lookup_Products';

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
      productId: tsCol('Product_ID_Machine'),

      itemName: tsCol('Item_Name_Entered'),
      brandName: tsCol('Brand_Name_Entered'),

      productEntered: tsCol('Product_Name_Entered'),
      productCanon: tsCol('Product_Name_Canonical')
    };

    for (const [k,v] of Object.entries(IDX)) {
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
    let skipHasProduct = 0;
    let skipNoCanon = 0;
    let skipDuplicateCanon = 0;

    const rowsToAppend = [];

    for (let i = 1; i < tsData.length; i++) {

      scanned++;

      const r = tsData[i];

      if (!r[IDX.txnId]) { skipNoTxn++; continue; }
      if (r[IDX.productId]) { skipHasProduct++; continue; }

      const canon = r[IDX.productCanon];
      if (!canon) { skipNoCanon++; continue; }

      if (stagingCanonSet.has(canon)) { skipDuplicateCanon++; continue; }

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

      /* =========================
         STEP: NOTICE — NO INSERT
      ========================= */
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
      `Skipped: NoTxn=${skipNoTxn}, HasProduct=${skipHasProduct}, NoCanon=${skipNoCanon}, Duplicate=${skipDuplicateCanon} | ` +
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





// PRODUCTS: PROCESS STAGED PRODUCTS
/**
 * Script Name: processStagingProducts_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Deterministically process governance state for Staging_Lookup_Products.
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
 * Product_Status
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
 * - Sheet must exist: Staging_Lookup_Products
 * - Sheet must exist: Automation_Control
 * - Integrity control switch: Automation_Control!K2
 *
 * Failure Modes
 * - Required sheet missing
 * - Required column missing
 */
function processStagingProducts_StateMachine() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME   = 'Products';
  const FUNCTION_NAME = 'processStagingProducts_StateMachine';

  const TGT_SHEET = 'Staging_Lookup_Products';

  const t0 = new Date();

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const stgSh = ss.getSheetByName(TGT_SHEET);
    if (!stgSh) throw new Error('Staging_Lookup_Products sheet missing');

    /* =========================
       STEP — LOAD_DATA
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DATA');

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

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DATA');

    /* =========================
       STEP — DRIFT_REPAIR
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'DRIFT_REPAIR');

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

      let productStatus = 'To be Reviewed';

      if (promoted) {

        if (row[IDX.isActive])
          productStatus = 'Promoted (Live)';
        else if (row[IDX.isArchived])
          productStatus = 'Promoted (Archived)';
        else
          productStatus = 'Promoted (Hidden Dropdown)';

      } else {

        if (row[IDX.isArchived] && !row[IDX.isApproved])
          productStatus = 'Rejected';

        else if (row[IDX.isActive] && !row[IDX.isApproved])
          productStatus = 'Active (Temporary)';

        else if (row[IDX.isApproved] && !row[IDX.isActive])
          productStatus = 'Approved (Hidden Dropdown)';

        else if (row[IDX.isApproved] && row[IDX.isActive])
          productStatus = 'Approved & Activated (Temporary)';

        else if (row[IDX.isApproved] && row[IDX.isArchived])
          productStatus = 'Approved (Archived)';
      }

      row[IDX.productStatus] = productStatus;

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
          sheetName: TGT_SHEET,
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
        TGT_SHEET,
        'DRIFT_REPAIR',
        'No drift detected; all rows valid'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'DRIFT_REPAIR');

    /* =========================
       STEP — WRITE_BACK
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');

    stgSh
      .getRange(2,1,data.length-1,hdr.length)
      .setValues(data.slice(1));

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_BACK');

    /* =========================
       SUMMARY
    ========================= */

    const durationMs = new Date().getTime() - t0.getTime();

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Valid=${valid}, Repaired=${repaired}, Invalid=${invalid}, DurationMs=${durationMs}`
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





// PRODUCTS: PROMOTE ACTIONED PRODUCTS TO LOOKUP
/**
 * Script Name: promoteApprovedProducts_FromStaging_ToLookup
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Promote approved staging product rows into Lookup_Products exactly once
 * - Promotion authority = staging row (not canonical, not name)
 * - Record lineage via Is_Staging_Promoted (Lookup_Products)
 * - Record completion via Is_Lookup_Promoted (Staging_Lookup_Products)
 * - Preserve auditability and forward-only identity guarantees (ETI v1.3)
 *
 * Preconditions:
 * - Sheets must exist:
 *   - Staging_Lookup_Products
 *   - Lookup_Products
 * - Header row must exist in row 1
 * - Required columns must exist (header-based, order-independent)
 * - Flags Is_Lookup_Promoted and Is_Staging_Promoted are script-owned
 *
 * Algorithm (Step-by-Step):
 * 1. Generate Execution_ID for traceability.
 * 2. Load Lookup_Products and resolve column indexes via headers.
 * 3. Load Staging_Lookup_Products and resolve column indexes via headers.
 * 4. Iterate each staging row:
 *    a. Skip if Is_Approved ≠ TRUE.
 *    b. Skip if Is_Lookup_Promoted = TRUE (hard stop).
 *    c. Resolve final product name (Approved → fallback Entered).
 *    d. Generate full UUID for Product_ID_Machine.
 *    e. Construct a new Lookup_Products row using header-indexed placement.
 *    f. Mark Is_Staging_Promoted = TRUE in Lookup_Products.
 *    g. Write back Mapped_Product_ID_Machine, Is_Lookup_Promoted = TRUE,
 *       and contextual Notes into Staging_Lookup_Products.
 * 5. Batch-append all new Lookup_Products rows.
 * 6. Write back all staging updates.
 * 7. Emit execution summary and completion logs.
 *
 * Failure Modes:
 * - Missing required sheet
 * - Missing required column
 *
 * Reason for Deprecation (if applicable):
 * - N/A
 * - This script remains ACTIVE for ETI v1.3.
 * - Superseded only in v1.4 if product reconciliation
 *   or merge workflows are introduced.
 */
function promoteApprovedProducts_FromStaging_ToLookup() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME   = 'Products';
  const FUNCTION_NAME = 'promoteApprovedProducts_FromStaging_ToLookup';

  const SRC_SHEET = 'Staging_Lookup_Products';
  const TGT_SHEET = 'Lookup_Products';

  const t0 = new Date();

  try {

    /* =========================
       START
    ========================= */
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
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

    /* =========================
       STEP — LOAD_STAGING
    ========================= */
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

      const finalName =
        r[IDX_STG.approved] ||
        r[IDX_STG.entered];

      if (!finalName) { skipped++; continue; }

      const canon = r[IDX_STG.canon] || '';
      const stagingId = r[IDX_STG.stagingId];

      const productIdMachine = Utilities.getUuid();

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

      newLookupRow[IDX_LK.notes] =
        `Promoted from staging → Staging_ID=${stagingId}`;

      lookupAppendRows.push(newLookupRow);

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

    /* =========================
       NOTICE — NO PROMOTION
    ========================= */
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
      stgSh.getRange(u.row, IDX_STG.productStatus + 1).setValue(u.status);
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





// PRODUCTS:  LOOKUP - BACKFILL BACKEND-SEEDED(MANUALLY) IDs FOR BRAND RECORDS
/**
 * Script Name: backfill_ProductIDs_Machine_LookupProducts
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Backfill Product_ID_Machine where Product_Name exists and ID is missing
 *
 * Preconditions:
 * - Spreadsheet contains a sheet named: Lookup_Products
 * - Header row present in row 1
 * - Required columns:
 *   - Product_Name
 *   - Product_ID_Machine
 *
 * Algorithm:
 * 1. Generate Execution_ID
 * 2. Load Lookup_Products
 * 3. Resolve column indexes
 * 4. For each row:
 *    a. If Product_Name exists AND Product_ID_Machine is blank → generate UUID
 * 5. Write changes in one batch
 *
 * Failure Modes:
 * - Sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */
function backfill_ProductIDs_Machine_LookupProducts() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME   = 'Products';
  const FUNCTION_NAME = 'backfill_ProductIDs_Machine_LookupProducts';
  const TGT_SHEET     = 'Lookup_Products';

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
      productName: col('Product_Name'),
      productIdM:  col('Product_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
    }

    /* =========================
       STEP — GENERATE_ID
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'GENERATE_ID');

    let generatedCount = 0;
    const output = data.map(r => r.slice());

    for (let i = 1; i < output.length; i++) {

      const rowNum = i + 1;
      const name   = output[i][IDX.productName];
      const prodId = output[i][IDX.productIdM];

      if (name && !prodId) {

        const newId = Utilities.getUuid();
        output[i][IDX.productIdM] = newId;
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
          details: `Generated Product_ID_Machine: ${newId}`
        });
      }
    }

    /* =========================
       NOTICE — NO GENERATION
    ========================= */

    if (generatedCount === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'GENERATE_ID',
        'No Product_ID generated (all rows already populated or empty)'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'GENERATE_ID');

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
    ETI_logEnd_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET
    );

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





// PRODUCTS: LOOKUP - CLEANUP OF INVALID RECORDS
/**
 * Script Name: cleanupOrphan_ProductIDs_Machine_LookupProducts
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Clear orphan Product_ID_Machine where Product_Name is missing
 *
 * Preconditions:
 * - Spreadsheet contains a sheet named: Lookup_Products
 * - Header row present in row 1
 * - Required columns:
 *   - Product_Name
 *   - Product_ID_Machine
 *
 * Algorithm:
 * 1. Generate Execution_ID
 * 2. Load Lookup_Products
 * 3. Resolve column indexes
 * 4. For each row:
 *    a. If Product_Name is blank AND Product_ID_Machine exists → clear ID
 * 5. Write changes in one batch
 *
 * Failure Modes:
 * - Sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */
function cleanupOrphan_ProductIDs_Machine_LookupProducts() {

  /* =========================
     CONFIG / CONSTANTS
  ========================= */
  const SCRIPT_NAME   = 'Products';
  const FUNCTION_NAME = 'cleanupOrphan_ProductIDs_Machine_LookupProducts';
  const TGT_SHEET     = 'Lookup_Products';

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
      productName: col('Product_Name'),
      productIdM:  col('Product_ID_Machine')
    };

    for (const [k, v] of Object.entries(IDX)) {
      if (v === -1) throw new Error(`Missing required column: ${k}`);
    }

    /* =========================
       STEP — CLEANUP_ORPHAN_ID
    ========================= */

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'CLEANUP_ORPHAN_ID');

    let clearedCount = 0;
    const output = data.map(r => r.slice());

    for (let i = 1; i < output.length; i++) {

      const rowNum = i + 1;
      const name   = output[i][IDX.productName];
      const prodId = output[i][IDX.productIdM];

      if (!name && prodId) {

        output[i][IDX.productIdM] = '';
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
          details: `Product_Name missing; Cleared Product_ID_Machine: ${prodId}`
        });
      }
    }

    /* =========================
       NOTICE — NO CLEANUP
    ========================= */

    if (clearedCount === 0) {
      ETI_logNotice_(
        SCRIPT_NAME,
        FUNCTION_NAME,
        TGT_SHEET,
        'CLEANUP_ORPHAN_ID',
        'No orphan Product_ID found'
      );
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'CLEANUP_ORPHAN_ID');

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
    ETI_logEnd_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET
    );

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
