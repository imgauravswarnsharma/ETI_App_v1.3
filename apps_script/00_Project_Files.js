/*
-----------------------------------
GENERIC FILE RESOLVER (FAST)
-----------------------------------
*/

function getSpreadsheetByName_(fileName){

  const files = DriveApp.getFilesByName(fileName);

  if (!files.hasNext()) {
    throw new Error(`File "${fileName}" not found`);
  }

  const file = files.next();

  // Optional safety: ensure no duplicates
  if(files.hasNext()){
    throw new Error(`Multiple files found for "${fileName}"`);
  }

  return SpreadsheetApp.open(file);
}


/*
-----------------------------------
METADATA
-----------------------------------
*/

function getMetadataSpreadsheet_(){
  return getSpreadsheetByName_('ETI_App_v1.3_Metadata');
}


/*
-----------------------------------
SUPPORT
-----------------------------------
*/

function getSupportSpreadsheet_(){
  return getSpreadsheetByName_('ETI_App_v1.3_Support');
}


/*
-----------------------------------
LOGS
-----------------------------------
*/

function getLogsSpreadsheet_(){
  return getSpreadsheetByName_('ETI_App_v1.3_Logs');
}
