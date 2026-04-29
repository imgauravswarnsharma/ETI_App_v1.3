// METADATA (PART B): DATA VALIDATION
/**
 * Script Name: extractDataValidationConfig
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1
 * Current Status: ACTIVE (Performance Optimized)
 *
 * Purpose:
 * - Extract data validation configuration as read-only manifest
 *
 * Preconditions:
 * - Schema_Snapshot exists
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Build schema lookup
 * 2. Scan sheets for validation rules
 * 3. Build output in memory
 * 4. Single batch write
 *
 */ 

function extractDataValidationConfig() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();

  const SCHEMA = metaSS.getSheetByName('Schema_Snapshot')
                   .getDataRange()
                   .getValues();

  let out = metaSS.getSheetByName('Config_Data_Validation');
  if (!out) out = metaSS.insertSheet('Config_Data_Validation');
  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Letter',
    'Column_Name',
    'Rule_Type',
    'Criteria',
    'Allowed_Values',
    'Help_Text',
    'Strict_Flag'
  ];

  out.getRange(1, 1, 1, headers.length)
     .setValues([headers])
     .setFontWeight('bold');

  let writeRow = 4;
  let lastSheet = null;

/* ===================== BUILD SCHEMA LOOKUP ===================== */
  const schemaMap = {};
  for (let i = 1; i < SCHEMA.length; i++) {
    const [sheet, idx, letter, name] = SCHEMA[i];
    schemaMap[`${sheet}|${idx}`] = { letter, name };
  }

  dataSS.getSheets().forEach(sh => {

    const sheetName = sh.getName();
    const maxCols = sh.getMaxColumns();

    for (let col = 1; col <= maxCols; col++) {

      const rule = sh.getRange(2, col).getDataValidation();
      if (!rule) continue;

      const meta = schemaMap[`${sheetName}|${col}`];
      if (!meta) continue;

      if (lastSheet !== null && sheetName !== lastSheet) {
        writeRow += 2;
      }

      const type = rule.getCriteriaType();
      const args = rule.getCriteriaValues();

      let criteria = type;
      let allowed = '';

      if (args && args.length) {
        allowed = args.map(v => {
          if (Array.isArray(v)) return v.flat().join(', ');
          return String(v);
        }).join(' | ');
      }

      out.getRange(writeRow, 1, 1, headers.length).setValues([[
        sheetName,
        meta.letter,
        meta.name,
        criteria,
        allowed,
        allowed,
        rule.getHelpText() || '',
        rule.getAllowInvalid() ? 'WARN' : 'STRICT'
      ]]);

      lastSheet = sheetName;
      writeRow++;
    }
  });
}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA : CONDITIONAL FORMATTING
/**
 * Script Name: extractConditionalFormattingConfig
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1
 * Current Status: ACTIVE (Performance Optimized)
 *
 * Purpose:
 * - Extract conditional formatting configuration as read-only manifest
 *
 * Preconditions:
 * - Active spreadsheet
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Read conditional rules per sheet
 * 2. Build output in memory
 * 3. Single batch write
 */

function extractConditionalFormattingConfig() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();

  let out = metaSS.getSheetByName('Config_Conditional_Formatting');
  if (!out) out = metaSS.insertSheet('Config_Conditional_Formatting');
  out.clear();

  const headers = [
    'Sheet_Name',
    'Applies_To_Range',
    'Rule_Type',
    'Formula_or_Criteria',
    'Notes'
  ];

  out.getRange(1, 1, 1, headers.length)
     .setValues([headers])
     .setFontWeight('bold');

  let output = [];
  let lastSheet = null;

  dataSS.getSheets().forEach(sh => {

    const sheetName = sh.getName();
    const rules = sh.getConditionalFormatRules();
    if (!rules || rules.length === 0) return;

    if (lastSheet !== null && sheetName !== lastSheet) {
      output.push(new Array(headers.length).fill(''));
      output.push(new Array(headers.length).fill(''));
    }

    rules.forEach(rule => {

      const ranges = rule.getRanges()
        .map(r => r.getA1Notation())
        .join(', ');

      const booleanCond = rule.getBooleanCondition();
      const gradientCond = rule.getGradientCondition();

      let ruleType = '';
      let criteria = '';
      let notes = '';

      if (booleanCond) {
        ruleType = booleanCond.getCriteriaType();
        const vals = booleanCond.getCriteriaValues();
        criteria = vals && vals.length ? vals.join(' | ') : '';
      }

      if (gradientCond) {
        ruleType = 'GRADIENT';

        const midPoint = gradientCond.getMidpoint();

        notes = [
          `MinColor=${gradientCond.getMinColor()}`,
          midPoint && midPoint.getColor() ? `MidColor=${midPoint.getColor()}` : '',
          `MaxColor=${gradientCond.getMaxColor()}`
        ].filter(Boolean).join(' ; ');
      }

      output.push([
        sheetName,
        ranges,
        ruleType,
        criteria,
        notes
      ]);
    });

    lastSheet = sheetName;
  });

  if (output.length > 0) {
    out.getRange(4, 1, output.length, headers.length)
       .setValues(output);
  }
}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA : NAMED RANGES
/**
 * Config_Named_Ranges — Manifest Extract v1
 * Read-only, safe
 */

function extractNamedRangesConfig() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();

  let out = metaSS.getSheetByName('Config_Named_Ranges');
  if (!out) out = metaSS.insertSheet('Config_Named_Ranges');
  out.clear();

  const headers = [
    'Range_Name',
    'Sheet_Name',
    'Range_A1'
  ];

  out.getRange(1, 1, 1, headers.length)
     .setValues([headers])
     .setFontWeight('bold');

  let writeRow = 4;

  const namedRanges = dataSS.getNamedRanges();

  namedRanges.forEach(nr => {
    const range = nr.getRange();
    const sheet = range.getSheet();

    out.getRange(writeRow, 1, 1, headers.length).setValues([[
      nr.getName(),
      sheet.getName(),
      range.getA1Notation()
    ]]);

    writeRow++;
  });
}

