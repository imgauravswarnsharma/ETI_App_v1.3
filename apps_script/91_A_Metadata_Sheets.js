// METADATA (PART A): SCHEMA GENERATION 
/**
 * Script Name: STEP3_exportSchemaSnapshot
 * Status: STABLE — ETI v1.3
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

function STEP3_exportSchemaSnapshot() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();
  const sheets = dataSS.getSheets();

  const OUTPUT_SHEET_NAME = 'Schema_Snapshot';
  let out = metaSS.getSheetByName(OUTPUT_SHEET_NAME);
  if (!out) out = metaSS.insertSheet(OUTPUT_SHEET_NAME);

  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Header_Value',
    'Table_Role' //
  ];

  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  let output = [];
  let separatorRows = [];
  let lastSheetName = null;

  /* =========================================================
     NEW: TABLE ROLE FUNCTION
  ========================================================== */

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

  /* ========================================================= */

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

    const headerValues =
      sh.getRange(1, 1, 1, lastCol).getValues()[0];

    for (let idx = 0; idx < headerValues.length; idx++) {
      output.push([
        sheetName,
        idx + 1,
        columnToLetter(idx + 1),
        headerValues[idx],
        tableRole // ✅ NEW
      ]);
    }

    lastSheetName = sheetName;
  });

  /* ========================================================= */

  if (output.length > 0) {
    out.getRange(4, 1, output.length, headers.length)
       .setValues(output);
  }

  separatorRows.forEach(r => {
    out.getRange(r, 1, 1, headers.length)
       .setBackground('#FFFF00');
  });
}





/**
 * Utility: Convert column number to letter (1 → A, 27 → AA)
 */
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





// METADATA: FORMULA REPOSITORY OF SHEET
/**
 * Script Name: exportFormulaInventory_v2_manifest
 * Status: STABLE — ETI v1.3
 *
 * Purpose:
 * - Capture spreadsheet formulas as inert text (manifest mode)
 * - Provide a deterministic, text-only formula inventory
 *
 * Explicit Non-Goals:
 * - Does NOT evaluate formulas
 * - Does NOT format headers (owned by sheet-level formatter)
 * - Does NOT mutate source sheets
 * - Does NOT interact with AppSheet or transactional logic
 *
 * Rules (Locked):
 * - Read formulas from ROW 2 only
 * - Strip leading "="
 * - Store formulas as inert text
 * - Clear + rewrite output on every run
 *
 * Input Dependencies:
 * - Active spreadsheet (all sheets)
 *
 * Output Contract:
 * - Sheet: Formula_Inventory
 * - Layout:
 *   Row 1  → Header (values only)
 *   Row 2–3 → Reserved empty rows
 *   Row 4+ → Data
 *   Two empty separator rows inserted when Sheet_Name changes
 *
 * Cosmetic Rules (Intentional & Limited):
 * - Separator rows are visually marked in columns A–F only
 * - Color: #FFFF00 (bright yellow)
 * - Purpose: improve human scanability between sheet blocks
 *
 * Idempotency:
 * - Safe to re-run; output is fully cleared and rebuilt
 */

function exportFormulaInventory_v2_manifest() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();

  const sheets = dataSS.getSheets();


  const OUTPUT_SHEET = 'Formula_Inventory';
  let out = metaSS.getSheetByName(OUTPUT_SHEET);
  if (!out) out = metaSS.insertSheet(OUTPUT_SHEET);
  
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

  let output = [];
  let separatorRows = [];
  let lastSheetName = null;

  sheets.forEach(sh => {

    const sheetName = sh.getName();
    const lastCol = sh.getLastColumn();

    if (lastCol === 0) {
      lastSheetName = sheetName;
      return;
    }

    if (lastSheetName !== null && sheetName !== lastSheetName) {
      separatorRows.push(output.length + 4); // +4 because metadata starts row 4
      separatorRows.push(output.length + 5);
      output.push(new Array(headers.length).fill(''));
      output.push(new Array(headers.length).fill(''));
    }

    const headersRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const formulaA1 = sh.getRange(2, 1, 1, lastCol).getFormulas()[0];
    const formulaR1C1 = sh.getRange(2, 1, 1, lastCol).getFormulasR1C1()[0];

    for (let col = 1; col <= lastCol; col++) {

      let fA1 = formulaA1[col - 1] || '';
      let fR1C1 = formulaR1C1[col - 1] || '';

      if (fA1.startsWith('=')) fA1 = fA1.slice(1);
      if (fR1C1.startsWith('=')) fR1C1 = fR1C1.slice(1);

      output.push([
        sheetName,
        col,
        columnToLetter(col),
        headersRow[col - 1] || '',
        fA1,
        fR1C1
      ]);
    }

    lastSheetName = sheetName;
  });

  if (output.length > 0) {
    out.getRange(4, 1, output.length, headers.length)
       .setValues(output);
  }

  // Apply separator background in batch
  separatorRows.forEach(r => {
    out.getRange(r, 1, 1, headers.length)
       .setBackground('#FFFF00');
  });
}







// METADATA: AUTOMATIC COLUMN CLASSIFICATION
/**
 * Script Name: classifyColumns_fromManifest
 * Status: UPDATED — ETI v1.3 (Column Role Classification Added)
 *
 * Purpose:
 * - Generate Column_Classification as a READ-ONLY analytical artifact
 * - Classify columns using Schema_Snapshot + Formula_Inventory (text-only)
 * - ADD: Column_Role semantic classification
 *
 * Explicit Non-Goals:
 * - Does NOT mutate schemas
 * - Does NOT write formulas
 * - Does NOT manage header formatting
 * - Does NOT interact with AppSheet or transactional data
 *
 * Input Dependencies:
 * - Schema_Snapshot
 * - Formula_Inventory
 *
 * Output Contract:
 * - Sheet: Column_Classification
 * - Added Column: Column_Role
 *
 * Idempotency:
 * - Safe to re-run; output is fully cleared and rebuilt
 */

function classifyColumns_fromManifest() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();

  const SCHEMA_SHEET   = metaSS.getSheetByName('Schema_Snapshot');
  const FORMULA_SHEET  = metaSS.getSheetByName('Formula_Inventory');

  if (!SCHEMA_SHEET || !FORMULA_SHEET) {
    throw new Error('Required dependency sheet missing');
  }

  const SCHEMA   = SCHEMA_SHEET.getDataRange().getValues();
  const FORMULAS = FORMULA_SHEET.getDataRange().getValues();

  let out = metaSS.getSheetByName('Column_Classification');
  if (!out) out = metaSS.insertSheet('Column_Classification');

  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Column_Name',
    'Semantic_Class',
    'Column_Role' // ✅ NEW
  ];

  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  /* =========================================================
     BUILD FORMULA LOOKUP MAP
  ========================================================== */

  const formulaMap = {};
  for (let i = 1; i < FORMULAS.length; i++) {
    const [sheet, colIdx,,,, formulaA1] = FORMULAS[i];
    formulaMap[`${sheet}|${colIdx}`] = (formulaA1 || '');
  }

  /* =========================================================
     CLASSIFICATION BUILD (IN MEMORY)
  ========================================================== */

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

    /* =========================================================
       NEW: COLUMN ROLE CLASSIFICATION
    ========================================================== */

    const name = colName.toUpperCase();
    let role = 'OTHER';

    if (/_ID\b/.test(name)) {
      role = 'IDENTIFIER';

    } else if (/_KEY\b/.test(name)) {
      role = 'FOREIGN_KEY';

    } else if (cls === 'EMPTY') {
      role = 'INPUT';

    } else if (/^IS_|_FLAG\b/.test(name)) {
      role = 'FLAG';

    } else if (/_UI\b/.test(name)) {
      role = 'UI_FIELD';

    } else if (/(DATE|TIME|CREATED|UPDATED)/.test(name)) {
      role = 'SYSTEM_FIELD';

    } else if (/(RATE|AMOUNT|QTY|VALUE|COUNT)/.test(name)) {
      role = 'METRIC';

    } else if (cls === 'DERIVED_LOOKUP') {
      role = 'LOOKUP_DERIVED';

    } else if (cls === 'DERIVED_LOCAL') {
      role = 'COMPUTED';
    }

    output.push([
      sheet,
      colIdx,
      colLetter,
      colName,
      cls,
      role // ✅ NEW
    ]);

    lastSheet = sheet;
  }

  /* =========================================================
     SINGLE BATCH WRITE
  ========================================================== */

  if (output.length > 0) {
    out.getRange(4, 1, output.length, headers.length)
       .setValues(output);
  }

  /* =========================================================
     BATCH SEPARATOR FORMATTING
  ========================================================== */

  separatorRows.forEach(r => {
    out.getRange(r, 1, 1, headers.length)
       .setBackground('#FFFF00');
  });
}






// METADATA: FINAL SHEET ARTIFACT - DERIVED COLUMN LOGIC
/**
 * Script Name: generateDerivedColumnLogic
 * Status: STABLE — ETI v1.3
 *
 * Purpose:
 * - Resolve derived column dependencies using inert formulas
 * - Expand same-sheet (R1C1) and cross-sheet (A1) references
 *
 * Explicit Non-Goals:
 * - Does NOT evaluate formulas
 * - Does NOT mutate schemas or formulas
 * - Does NOT format headers (owned by sheet-level formatter)
 * - Does NOT interact with AppSheet or transactional data
 *
 * Input Dependencies (Required):
 * - Schema_Snapshot
 * - Formula_Inventory
 * - Column_Classification
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Build schema and formula lookup maps
 * 2. Resolve references in memory
 * 3. Single batch write
 * 4. Batch separator formatting
 *
 * Output Contract:
 * - Sheet: Derived_Column_Logic (fully regenerated each run)
 * - Layout:
 *   Row 1  → Header (values only)
 *   Row 2–3 → Reserved empty rows
 *   Row 4+ → Data
 *   Two empty separator rows inserted when Sheet_Name changes
 *
 * Cosmetic Rules (Intentional & Limited):
 * - Separator rows are visually marked in columns A–H only
 * - Color: #FFFF00 (bright yellow)
 * - Purpose: improve human scanability between derived blocks
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */

function generateDerivedColumnLogic() {

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

  let out = metaSS.getSheetByName('Derived_Column_Logic');
  if (!out) out = metaSS.insertSheet('Derived_Column_Logic');
  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Column_Name',
    'Semantic_Class',
    'Formula_A1_Text',
    'Formula_R1C1_Text',
    'Resolved_References',
    'Upstream_Lineage',
    'Table_Dependencies',
    'Semantic_Meaning',   // 
    'Semantic_Purpose'    // 
  ];

  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  /* ===================== LOOKUPS ===================== */

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

  let output = [];
  let separatorRows = [];
  let lastSheet = null;

  /* ===================== MAIN BUILD ===================== */

  for (let i = 1; i < CLASS.length; i++) {

    const [sheet, colIdx, colLetter, colName, semantic] = CLASS[i];

    if (!['DERIVED_LOCAL', 'DERIVED_LOOKUP'].includes(semantic)) continue;

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

    const r1c1 = f.r1c1 || '';
    const rcMatches = r1c1.match(/RC\[[+-]?\d+\]/g) || [];

    rcMatches.forEach(token => {
      const offset = parseInt(token.match(/[+-]?\d+/)[0], 10);
      const targetIdx = colIdx + offset;
      const meta = schemaMap[`${sheet}|${targetIdx}`];
      if (meta) {
        refs.add(`${sheet}.${meta.letter} (${meta.name})`);
        refKeys.push(`${sheet}|${targetIdx}`);
      }
    });

    const a1 = f.a1 || '';
    const a1Matches = a1.match(/([A-Z0-9_]+)!\$?[A-Z]+/gi) || [];

    a1Matches.forEach(ref => {
      const [refSheet, colPart] = ref.split('!');
      const colLetterRef = colPart.replace(/[^A-Z]/gi, '');
      const meta = schemaLetterMap[`${refSheet}|${colLetterRef}`];
      if (meta) {
        refs.add(`${refSheet}.${colLetterRef} (${meta.name})`);
        refKeys.push(`${refSheet}|${meta.idx}`);
      }
    });

    const nodeKey = `${sheet}|${colIdx}`;
    graph[nodeKey] = refKeys;

    output.push([
      sheet,
      colIdx,
      colLetter,
      colName,
      semantic,
      f.a1,
      f.r1c1,
      Array.from(refs).join('\n'),
      '',
      '',
      '',
      '' // placeholders
    ]);

    lastSheet = sheet;
  }

  /* ===================== LINEAGE ===================== */

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

  /* ===================== ATTACH ALL ===================== */

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

    /* ===================== SEMANTICS ===================== */

    const colName = row[3].toUpperCase();
    const role = (classMap[key] || {}).role || '';

    let meaning = 'General field';
    let purpose = 'General usage';

    if (/_ID\b/.test(colName)) meaning = 'Unique identifier';
    else if (/_KEY\b/.test(colName)) meaning = 'Reference key to another entity';
    else if (/RATE/.test(colName)) meaning = 'Unit price of item';
    else if (/QTY/.test(colName)) meaning = 'Quantity value';
    else if (/AMOUNT/.test(colName)) meaning = 'Total monetary value';
    else if (/DATE/.test(colName)) meaning = 'Event timestamp';
    else if (/FLAG|^IS_/.test(colName)) meaning = 'Boolean indicator';
    else if (/_UI\b/.test(colName)) meaning = 'Formatted display value';

    switch (role) {
      case 'INPUT': purpose = 'User input field'; break;
      case 'IDENTIFIER': purpose = 'Uniquely identifies row'; break;
      case 'FOREIGN_KEY': purpose = 'Links to another entity'; break;
      case 'METRIC': purpose = 'Used for analytical computation'; break;
      case 'LOOKUP_DERIVED': purpose = 'Derived from lookup tables'; break;
      case 'COMPUTED': purpose = 'Computed from other fields'; break;
      case 'FLAG': purpose = 'Controls filtering or logic'; break;
      case 'UI_FIELD': purpose = 'Used for UI display'; break;
      case 'SYSTEM_FIELD': purpose = 'System tracking or audit'; break;
    }

    row[10] = meaning;
    row[11] = purpose;
  }

  /* ===================== WRITE ===================== */

  if (output.length > 0) {
    out.getRange(4, 1, output.length, headers.length)
       .setValues(output);
  }

  separatorRows.forEach(r => {
    out.getRange(r, 1, 1, headers.length)
       .setBackground('#FFFF00');
  });
}