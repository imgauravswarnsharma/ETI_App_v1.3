/*
=========================================================
UTILITY: SHEET CAPACITY MANAGER (GENERIC)
=========================================================

PURPOSE:
- Ensure sheet has sufficient free rows
- Expand capacity only when required

DESIGN:
✔ Sheet-agnostic
✔ Config-driven
✔ Plug-and-play
✔ No dependency on logger/archiver
✔ Safe for reuse across system

=========================================================
*/

function ensureFreeRowCapacity_(sheet, options = {}){

  if (!sheet) return null;

  const MIN_FREE_ROWS = options.minFreeRows || 1200;
  const EXPANSION_CHUNK = options.expansionChunk || 3600;

  const maxRows = sheet.getMaxRows();
  const lastRow = sheet.getLastRow();

  const freeRowsBefore = maxRows - lastRow;

  if (freeRowsBefore >= MIN_FREE_ROWS) {
    return {
      expanded: false,
      rowsAdded: 0,
      freeRowsBefore,
      freeRowsAfter: freeRowsBefore
    };
  }

  sheet.insertRowsAfter(maxRows, EXPANSION_CHUNK);

  const newMaxRows = sheet.getMaxRows();
  const freeRowsAfter = newMaxRows - lastRow;

  return {
    expanded: true,
    rowsAdded: EXPANSION_CHUNK,
    freeRowsBefore,
    freeRowsAfter
  };
}