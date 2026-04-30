/* 
===============================================
SCHEMA SNAPSHOT
===============================================*/
/**
 * Script Name: exportSchemaSnapshot
 *
 * Purpose:
 * - Capture column order and header values for all sheets
 * - Freeze schema exactly as Google Sheets exposes it
 *
 * Explicit Non-Goals:
 * - Does NOT format headers (owned by sheet-level formatter)
 * - Does NOT mutate source sheets
 * - Does NOT interact with AppSheet or transactional data
 *
 * Execution Flow:
 * 1. Load active spreadsheet and metadata spreadsheet
 * 2. Initialize / clear Schema_Snapshot output sheet
 * 3. Iterate all sheets in source spreadsheet
 * 4. Extract header row and determine table role
 * 5. Build output array with schema details
 * 6. Insert separator rows between sheet blocks
 * 7. Write output in single batch
 * 8. Apply separator formatting
 *
 * Input Dependencies:
 * - Active spreadsheet (all visible sheets)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Clear output sheet
 * 2. Build full output array in memory
 * 3. Write once using setValues()
 * 4. Apply separator formatting in batch
 *
 * Output Contract:
 * - Sheet: Schema_Snapshot (fully regenerated each run)
 * - Layout:
 *   Row 1  → Header (values only)
 *   Row 2–3 → Reserved empty rows
 *   Row 4+ → Data
 *   Two empty separator rows inserted when Sheet_Name changes
 *
 * Cosmetic Rules (Intentional & Limited):
 * - Separator rows are visually marked in columns A–D only
 * - Color: #FFFF00 (bright yellow)
 * - Purpose: improve human scanability between sheet blocks
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function exportSchemaSnapshot() {

  const SCRIPT_NAME   = 'Metadata_Sheets';
  const FUNCTION_NAME = 'exportSchemaSnapshot';
  const TGT_SHEET     = 'Schema_Snapshot';
  const t0 = new Date();

  try {
    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    /*
    -----------------
    STEP — LOAD
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    const dataSS = SpreadsheetApp.getActiveSpreadsheet();
    const metaSS = getMetadataSpreadsheet_();
    const sheets = dataSS.getSheets();

    let out = metaSS.getSheetByName(TGT_SHEET);
    if (!out) out = metaSS.insertSheet(TGT_SHEET);

    out.clear();

    const headers = [
      'Sheet_Name',
      'Column_Index',
      'Column_Letter',
      'Header_Value',
      'Table_Role'
    ];

    out.getRange(1, 1, 1, headers.length).setValues([headers]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    /*
    -----------------
    STEP — PROCESS
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    let output = [];
    let separatorRows = [];
    let lastSheetName = null;

    function getTableRole(sheetName) {
      if (sheetName === 'Transaction_Raw') return 'WRITE';
      if (sheetName === 'Transaction_Resolution') return 'RESOLUTION';
      if (sheetName === 'Transaction_Analytics') return 'ANALYTICS';
      if (sheetName === 'Item_Spine_Extract') return 'EXTRACT';
      if (sheetName === 'Itemwise_Analytics') return 'ANALYTICS';
      if (sheetName === 'Item_Buy_Evaluate') return 'EVALUATION';
      if (sheetName === 'Item_Evaluation_Log') return 'LOG';
      if (sheetName === 'Item_Evaluation_Analytics') return 'ANALYTICS';

      if (/^Staging_/.test(sheetName)) return 'STAGING';
      if (/^Lookup_/.test(sheetName)) return 'LOOKUP';
      if (/^Mapping_/.test(sheetName)) return 'MAPPING';

      if (sheetName === 'Automation_Control') return 'CONTROL';
      if (sheetName === 'Data_Flow_Control') return 'CONTROL';

      return 'OTHER';
    }

    sheets.forEach(sh => {

      const sheetName = sh.getName();
      const tableRole = getTableRole(sheetName);
      const lastCol = sh.getLastColumn();

      if (lastCol === 0) {
        lastSheetName = sheetName;
        return;
      }

      if (lastSheetName !== null && sheetName !== lastSheetName) {
        separatorRows.push(output.length + 4);
        separatorRows.push(output.length + 5);
        output.push(new Array(headers.length).fill(''));
        output.push(new Array(headers.length).fill(''));
      }

      const headerValues = sh.getRange(1, 1, 1, lastCol).getValues()[0];

      for (let idx = 0; idx < headerValues.length; idx++) {
        output.push([
          sheetName,
          idx + 1,
          columnToLetter(idx + 1),
          headerValues[idx],
          tableRole
        ]);
      }

      lastSheetName = sheetName;
    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    /*
    -----------------
    STEP — WRITE_OUTPUT
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if (output.length > 0) {
      out.getRange(4, 1, output.length, headers.length).setValues(output);
    }

    separatorRows.forEach(r => {
      out.getRange(r, 1, 1, headers.length).setBackground('#FFFF00');
    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    /*
    -----------------
    SUMMARY
    -----------------*/
    const durationMs = new Date() - t0;

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${output.length} | DurationMs=${durationMs}`
    );

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

  } catch (err) {

    /*
    -------------------------------------
    ERROR LOGGING
    -------------------------------------*/
    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, err, 'MAIN');
    throw err;

  } finally {

    /*
    -------------------------------------
    FINAL LOG FLUSH (MANDATORY)
    -------------------------------------*/
    flushLogs_();
  }
}


/* 
===============================================
FORMULA INVENTORY
===============================================*/
/**
 * Script Name: exportFormulaInventory
 *
 * Purpose:
 * - Capture column-level formulas from all sheets
 * - Extract both A1 and R1C1 representations
 * - Provide formula-level visibility for metadata analysis
 *
 * Explicit Non-Goals:
 * - Does NOT evaluate formulas
 * - Does NOT mutate source sheets
 * - Does NOT interact with transactional logic
 *
 * Execution Flow:
 * 1. Load active spreadsheet and metadata spreadsheet
 * 2. Initialize / clear Formula_Inventory output sheet
 * 3. Iterate all sheets in source spreadsheet
 * 4. Extract header row and formula rows (A1 + R1C1)
 * 5. Normalize formulas (remove leading "=")
 * 6. Build output array
 * 7. Insert separator rows between sheet blocks
 * 8. Write output in batch
 * 9. Apply separator formatting
 *
 * Input Dependencies:
 * - Active spreadsheet (all visible sheets)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Clear output sheet
 * 2. Read headers and formula rows
 * 3. Normalize formulas (strip "=")
 * 4. Build output array in memory
 * 5. Write once using setValues()
 * 6. Apply separator formatting
 *
 * Output Contract:
 * - Sheet: Formula_Inventory (fully regenerated each run)
 * - Layout:
 *   Row 1  → Header
 *   Row 2–3 → Reserved empty rows
 *   Row 4+ → Data
 *   Separator rows inserted between sheet groups
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function exportFormulaInventory() {

  const SCRIPT_NAME   = 'Metadata_Sheets';
  const FUNCTION_NAME = 'exportFormulaInventory';
  const TGT_SHEET     = 'Formula_Inventory';

  const t0 = new Date();

  try {

    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    /*
    -----------------
    STEP — LOAD
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    const dataSS = SpreadsheetApp.getActiveSpreadsheet();
    const metaSS = getMetadataSpreadsheet_();
    const sheets = dataSS.getSheets();

    let out = metaSS.getSheetByName(TGT_SHEET);
    if (!out) out = metaSS.insertSheet(TGT_SHEET);
    out.clear();

    const headers = [
      'Sheet_Name',
      'Column_Index',
      'Column_Letter',
      'Column_Name',
      'Formula_A1_Text',
      'Formula_R1C1_Text'
    ];

    out.getRange(1, 1, 1, headers.length).setValues([headers]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    /*
    -----------------
    STEP — PROCESS
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    let output = [];
    let separatorRows = [];
    let lastSheetName = null;

    sheets.forEach(sh => {

      const sheetName = sh.getName();
      const lastCol = sh.getLastColumn();

      if (lastCol === 0) return;

      if (lastSheetName && sheetName !== lastSheetName) {
        separatorRows.push(output.length + 4);
        separatorRows.push(output.length + 5);
        output.push(new Array(headers.length).fill(''));
        output.push(new Array(headers.length).fill(''));
      }

      const headersRow = sh.getRange(1,1,1,lastCol).getValues()[0];
      const fA1 = sh.getRange(2,1,1,lastCol).getFormulas()[0];
      const fR1C1 = sh.getRange(2,1,1,lastCol).getFormulasR1C1()[0];

      for (let col = 1; col <= lastCol; col++) {

        let a1 = fA1[col-1] || '';
        let r1 = fR1C1[col-1] || '';

        if (a1.startsWith('=')) a1 = a1.slice(1);
        if (r1.startsWith('=')) r1 = r1.slice(1);

        output.push([
          sheetName,
          col,
          columnToLetter(col),
          headersRow[col-1] || '',
          a1,
          r1
        ]);
      }

      lastSheetName = sheetName;
    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    /*
    -----------------
    STEP — WRITE_OUTPUT
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if (output.length > 0) {
      out.getRange(4,1,output.length,headers.length).setValues(output);
    }

    separatorRows.forEach(r => {
      out.getRange(r,1,1,headers.length).setBackground('#FFFF00');
    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    /*
    -----------------
    SUMMARY
    -----------------*/
    const durationMs = new Date() - t0;

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${output.length} | DurationMs=${durationMs}`
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
===============================================
COLUMN CLASSIFICATION
===============================================*/
/**
 * Script Name: classifyColumns_fromManifest
 *
 * Purpose:
 * - Generate Column_Classification as a READ-ONLY analytical artifact
 * - Classify columns using Schema_Snapshot + Formula_Inventory (text-only)
 * - Assign semantic class and column role
 *
 * Explicit Non-Goals:
 * - Does NOT mutate schemas
 * - Does NOT write formulas
 * - Does NOT manage header formatting
 * - Does NOT interact with AppSheet or transactional data
 *
 * Execution Flow:
 * 1. Load Schema_Snapshot and Formula_Inventory
 * 2. Initialize / clear Column_Classification sheet
 * 3. Build formula lookup map
 * 4. Iterate schema rows
 * 5. Classify column type (EMPTY / DERIVED / PASS_THROUGH)
 * 6. Assign semantic role based on naming + classification
 * 7. Build output array
 * 8. Write output in batch
 * 9. Apply separator formatting
 *
 * Input Dependencies:
 * - Schema_Snapshot
 * - Formula_Inventory
 *
 * Algorithm:
 * 1. Build formula lookup map using sheet + column index key
 * 2. Iterate schema rows
 * 3. Determine formula presence and normalize
 * 4. Classify formula type
 * 5. Derive column role using regex-based rules
 * 6. Append results to output array
 * 7. Write results in single batch
 *
 * Output Contract:
 * - Sheet: Column_Classification
 * - Columns:
 *   Sheet_Name, Column_Index, Column_Letter, Column_Name,
 *   Semantic_Class, Column_Role
 *
 * Idempotency:
 * - Safe to re-run; output is fully cleared and rebuilt
 */

function classifyColumns_fromManifest() {

  const SCRIPT_NAME   = 'Metadata_Sheets';
  const FUNCTION_NAME = 'classifyColumns_fromManifest';
  const TGT_SHEET     = 'Column_Classification';

  const t0 = new Date();

  try {

    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    /*
    -----------------
    STEP — LOAD_DEPENDENCIES
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DEPENDENCIES');

    const dataSS = SpreadsheetApp.getActiveSpreadsheet();
    const metaSS = getMetadataSpreadsheet_();

    const SCHEMA_SHEET  = metaSS.getSheetByName('Schema_Snapshot');
    const FORMULA_SHEET = metaSS.getSheetByName('Formula_Inventory');

    if (!SCHEMA_SHEET || !FORMULA_SHEET) {
      throw new Error('Required dependency sheet missing');
    }

    const SCHEMA   = SCHEMA_SHEET.getDataRange().getValues();
    const FORMULAS = FORMULA_SHEET.getDataRange().getValues();

    let out = metaSS.getSheetByName(TGT_SHEET);
    if (!out) out = metaSS.insertSheet(TGT_SHEET);

    out.clear();

    const headers = [
      'Sheet_Name',
      'Column_Index',
      'Column_Letter',
      'Column_Name',
      'Semantic_Class',
      'Column_Role'
    ];

    out.getRange(1, 1, 1, headers.length).setValues([headers]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DEPENDENCIES');

    /*
    -----------------
    STEP — BUILD_FORMULA_MAP
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'BUILD_FORMULA_MAP');

    const formulaMap = {};
    for (let i = 1; i < FORMULAS.length; i++) {
      const [sheet, colIdx,,,, formulaA1] = FORMULAS[i];
      formulaMap[`${sheet}|${colIdx}`] = (formulaA1 || '');
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'BUILD_FORMULA_MAP');

    /*
    -----------------
    STEP — CLASSIFICATION
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'CLASSIFICATION');

    const PASS_THROUGH_REGEX =
      /^IF\s*\(\s*[^,]+,\s*(?:[A-Z0-9_]+!)?\$?[A-Z]+\$?\d+\s*,\s*""\s*\)$/i;

    let output = [];
    let separatorRows = [];
    let lastSheet = null;

    for (let i = 1; i < SCHEMA.length; i++) {

      const [sheet, colIdx, colLetter, colNameRaw] = SCHEMA[i];
      const colName = (colNameRaw || '').toString();

      if (lastSheet !== null && sheet !== lastSheet) {
        separatorRows.push(output.length + 4);
        separatorRows.push(output.length + 5);
        output.push(new Array(headers.length).fill(''));
        output.push(new Array(headers.length).fill(''));
      }

      const rawFormula = formulaMap[`${sheet}|${colIdx}`];

      let cls = 'EMPTY';

      if (rawFormula) {

        const f = rawFormula.replace(/\s+/g, ' ').trim();

        if (/(XLOOKUP|VLOOKUP|INDEX|MATCH)\s*\(/i.test(f)) {
          cls = 'DERIVED_LOOKUP';
        } else if (PASS_THROUGH_REGEX.test(f)) {
          cls = 'PASS_THROUGH';
        } else {
          cls = 'DERIVED_LOCAL';
        }
      }

      const name = colName.toUpperCase();
      let role = 'OTHER';

      if (/_ID\b/.test(name)) role = 'IDENTIFIER';
      else if (/_KEY\b/.test(name)) role = 'FOREIGN_KEY';
      else if (cls === 'EMPTY') role = 'INPUT';
      else if (/^IS_|_FLAG\b/.test(name)) role = 'FLAG';
      else if (/_UI\b/.test(name)) role = 'UI_FIELD';
      else if (/(DATE|TIME|CREATED|UPDATED)/.test(name)) role = 'SYSTEM_FIELD';
      else if (/(RATE|AMOUNT|QTY|VALUE|COUNT)/.test(name)) role = 'METRIC';
      else if (cls === 'DERIVED_LOOKUP') role = 'LOOKUP_DERIVED';
      else if (cls === 'DERIVED_LOCAL') role = 'COMPUTED';

      output.push([
        sheet,
        colIdx,
        colLetter,
        colName,
        cls,
        role
      ]);

      lastSheet = sheet;
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'CLASSIFICATION');

    /*
    -----------------
    STEP — WRITE_OUTPUT
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if (output.length > 0) {
      out.getRange(4, 1, output.length, headers.length).setValues(output);
    }

    separatorRows.forEach(r => {
      out.getRange(r, 1, 1, headers.length).setBackground('#FFFF00');
    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    /*
    -----------------
    SUMMARY
    -----------------*/
    const durationMs = new Date() - t0;

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${output.length} | DurationMs=${durationMs}`
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
===============================================
DERIVED COLUMN LOGIC
===============================================*/
/**
 * Script Name: generateDerivedColumnLogic
 *
 * Purpose:
 * - Resolve derived column dependencies using inert formulas
 * - Expand same-sheet (R1C1) and cross-sheet (A1) references
 *
 * Explicit Non-Goals:
 * - Does NOT evaluate formulas
 * - Does NOT mutate schemas or formulas
 * - Does NOT format headers
 * - Does NOT interact with AppSheet or transactional data
 *
 * Execution Flow:
 * 1. Load schema, formula, and classification sheets
 * 2. Initialize / clear Derived_Column_Logic sheet
 * 3. Build lookup maps (schema, formula, classification)
 * 4. Extract references (R1C1 + A1) and build dependency graph
 * 5. Resolve lineage recursively
 * 6. Enrich semantic meaning and purpose
 * 7. Write output in batch
 * 8. Apply separator formatting
 *
 * Input Dependencies:
 * - Schema_Snapshot
 * - Formula_Inventory
 * - Column_Classification
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Build lookup maps
 * 2. Parse references using regex
 * 3. Construct dependency graph
 * 4. Resolve lineage via recursive traversal
 * 5. Enrich semantics using naming + role rules
 * 6. Batch write output
 *
 * Output Contract:
 * - Sheet: Derived_Column_Logic
 * - Fully regenerated each run
 *
 * Idempotency:
 * - Safe to re-run; deterministic rebuild
 */
function generateDerivedColumnLogic() {

  const SCRIPT_NAME   = 'Metadata_Sheets';
  const FUNCTION_NAME = 'generateDerivedColumnLogic';
  const TGT_SHEET     = 'Derived_Column_Logic';

  const t0 = new Date();

  try {

    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    /*
    -----------------
    STEP — LOAD_DEPENDENCIES
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DEPENDENCIES');

    const dataSS = SpreadsheetApp.getActiveSpreadsheet();
    const metaSS = getMetadataSpreadsheet_();

    const SCHEMA_SHEET   = metaSS.getSheetByName('Schema_Snapshot');
    const FORMULA_SHEET  = metaSS.getSheetByName('Formula_Inventory');
    const CLASS_SHEET    = metaSS.getSheetByName('Column_Classification');

    if (!SCHEMA_SHEET || !FORMULA_SHEET || !CLASS_SHEET) {
      throw new Error('Required dependency sheet missing');
    }

    const SCHEMA   = SCHEMA_SHEET.getDataRange().getValues();
    const FORMULAS = FORMULA_SHEET.getDataRange().getValues();
    const CLASS    = CLASS_SHEET.getDataRange().getValues();

    let out = metaSS.getSheetByName(TGT_SHEET);
    if (!out) out = metaSS.insertSheet(TGT_SHEET);
    out.clear();

    const headers = [
      'Sheet_Name','Column_Index','Column_Letter','Column_Name',
      'Semantic_Class','Formula_A1_Text','Formula_R1C1_Text',
      'Resolved_References','Upstream_Lineage','Table_Dependencies',
      'Semantic_Meaning','Semantic_Purpose'
    ];

    out.getRange(1,1,1,headers.length).setValues([headers]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD_DEPENDENCIES');

    /*
    -----------------
    STEP — BUILD_LOOKUPS
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'BUILD_LOOKUPS');

    const schemaMap = {};
    const schemaLetterMap = {};
    for (let i = 1; i < SCHEMA.length; i++) {
      const [sheet, idx, letter, name] = SCHEMA[i];
      schemaMap[`${sheet}|${idx}`] = { letter, name };
      schemaLetterMap[`${sheet}|${letter}`] = { idx, name };
    }

    const formulaMap = {};
    for (let i = 1; i < FORMULAS.length; i++) {
      const [sheet, colIdx, , , a1, r1c1] = FORMULAS[i];
      formulaMap[`${sheet}|${colIdx}`] = { a1, r1c1 };
    }

    const classMap = {};
    for (let i = 1; i < CLASS.length; i++) {
      const [sheet, colIdx, , , semantic, role] = CLASS[i];
      classMap[`${sheet}|${colIdx}`] = { semantic, role };
    }

    const graph = {};

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'BUILD_LOOKUPS');

    /*
    -----------------
    STEP — BUILD_GRAPH
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'BUILD_GRAPH');

    let output = [];
    let separatorRows = [];
    let lastSheet = null;

    for (let i = 1; i < CLASS.length; i++) {

      const [sheet, colIdx, colLetter, colName, semantic] = CLASS[i];
      if (!['DERIVED_LOCAL','DERIVED_LOOKUP'].includes(semantic)) continue;

      if (lastSheet !== null && sheet !== lastSheet) {
        separatorRows.push(output.length + 4);
        separatorRows.push(output.length + 5);
        output.push(new Array(headers.length).fill(''));
        output.push(new Array(headers.length).fill(''));
      }

      const f = formulaMap[`${sheet}|${colIdx}`];
      if (!f) continue;

      const refs = new Set();
      const refKeys = [];

      const rcMatches = (f.r1c1 || '').match(/RC\[[+-]?\d+\]/g) || [];
      rcMatches.forEach(token => {
        const offset = parseInt(token.match(/[+-]?\d+/)[0],10);
        const targetIdx = colIdx + offset;
        const meta = schemaMap[`${sheet}|${targetIdx}`];
        if (meta) {
          refs.add(`${sheet}.${meta.letter} (${meta.name})`);
          refKeys.push(`${sheet}|${targetIdx}`);
        }
      });

      const a1Matches = (f.a1 || '').match(/([A-Z0-9_]+)!\$?[A-Z]+/gi) || [];
      a1Matches.forEach(ref => {
        const [refSheet, colPart] = ref.split('!');
        const colLetterRef = colPart.replace(/[^A-Z]/gi,'');
        const meta = schemaLetterMap[`${refSheet}|${colLetterRef}`];
        if (meta) {
          refs.add(`${refSheet}.${colLetterRef} (${meta.name})`);
          refKeys.push(`${refSheet}|${meta.idx}`);
        }
      });

      graph[`${sheet}|${colIdx}`] = refKeys;

      output.push([
        sheet,colIdx,colLetter,colName,semantic,
        f.a1,f.r1c1,Array.from(refs).join('\n'),
        '','','',''
      ]);

      lastSheet = sheet;
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'BUILD_GRAPH');

    /*
    -----------------
    STEP — RESOLVE_LINEAGE
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'RESOLVE_LINEAGE');

    function resolveLineage(startKey, visited = new Set(), depth = 0) {
      if (depth > 10) return [];
      const children = graph[startKey] || [];
      let result = [];

      children.forEach(child => {
        if (visited.has(child)) return;
        visited.add(child);

        const meta = schemaMap[child];
        if (meta) {
          result.push(`${child.split('|')[0]}.${meta.letter} (${meta.name})`);
        }

        result = result.concat(resolveLineage(child, visited, depth + 1));
      });

      return result;
    }

    for (let i = 0; i < output.length; i++) {

      const row = output[i];
      if (!row[0]) continue;

      const key = `${row[0]}|${row[1]}`;
      const lineage = Array.from(new Set(resolveLineage(key)));

      row[8] = lineage.join('\n');

      const tables = new Set();
      lineage.forEach(ref => {
        const table = ref.split('.')[0];
        if (table !== row[0]) tables.add(table);
      });
      row[9] = Array.from(tables).join('\n');

      const colName = row[3].toUpperCase();
      const role = (classMap[key] || {}).role || '';

      let meaning = 'General field';
      let purpose = 'General usage';

      if (/_ID\b/.test(colName)) meaning = 'Unique identifier';
      else if (/_KEY\b/.test(colName)) meaning = 'Reference key';
      else if (/RATE/.test(colName)) meaning = 'Unit price';
      else if (/QTY/.test(colName)) meaning = 'Quantity';
      else if (/AMOUNT/.test(colName)) meaning = 'Total value';
      else if (/DATE/.test(colName)) meaning = 'Timestamp';
      else if (/FLAG|^IS_/.test(colName)) meaning = 'Boolean';
      else if (/_UI\b/.test(colName)) meaning = 'Display field';

      switch (role) {
        case 'INPUT': purpose = 'User input'; break;
        case 'IDENTIFIER': purpose = 'Unique identifier'; break;
        case 'FOREIGN_KEY': purpose = 'Entity reference'; break;
        case 'METRIC': purpose = 'Analytical metric'; break;
        case 'LOOKUP_DERIVED': purpose = 'Lookup derived'; break;
        case 'COMPUTED': purpose = 'Computed field'; break;
        case 'FLAG': purpose = 'Control flag'; break;
        case 'UI_FIELD': purpose = 'UI display'; break;
        case 'SYSTEM_FIELD': purpose = 'System tracking'; break;
      }

      row[10] = meaning;
      row[11] = purpose;
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'RESOLVE_LINEAGE');

    /*
    -----------------
    STEP — WRITE_OUTPUT
    -----------------*/
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if (output.length > 0) {
      out.getRange(4,1,output.length,headers.length).setValues(output);
    }

    separatorRows.forEach(r => {
      out.getRange(r,1,1,headers.length).setBackground('#FFFF00');
    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    /*
    -----------------
    SUMMARY
    -----------------*/
    const durationMs = new Date() - t0;

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${output.length} | DurationMs=${durationMs}`
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
=========================================================================
UTILITY FUNCTIONS: Convert column number to letter (1 → A, 27 → AA)
=========================================================================*/
function columnToLetter(column) {
  let temp = '';
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}