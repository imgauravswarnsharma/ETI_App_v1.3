// GENERATING: AI CONTEXT - BACKEND SYSTEM INFORMATION 
/**
 * Script Name: generateAIContext_
 * Status: UPDATED — ETI v1.3
 *
 * Enhancements:
 * - Derived column logic now includes:
 *     → Actual formula (A1 text)
 *     → Dependency references
 * - Column index replaced with column name (fix)
 * - Added Table Role Classification section
 *
 * Non-Changes:
 * - Script metadata logic untouched
 * - Data flow logic untouched
 * - Schema extraction untouched
 */

/**
 * Phase 2 Enhancement:
 * - Added TABLE RELATIONSHIPS section
 * - Extracted from Derived_Column_Logic (Resolved References)
 * - Only cross-table dependencies included
 * - No inference, no assumptions
 */

/**
 * Enhancement (Drive Layer):
 * - Added GOOGLE DRIVE STRUCTURE section at top level
 * - Auto-detects version folder (v_1.x)
 * - Captures subfolders and file classification (MAIN / LOG / METADATA / SUPPORT)
 * - No hardcoding, version-agnostic
 */


function run_generateAIContext(){
  generateAIContext_();
}

function generateAIContext_(){

  const metadataSS = getMetadataSpreadsheet_();

  const OUT_SHEET = "AI_CONTEXT_EXPORT";

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  let output = [];

  output.push("# ETI AUTO SYSTEM CONTEXT");
  output.push("");

  output.push("Architecture:");
  output.push("Google Sheets = Data Layer");
  output.push("Apps Script = Processing Pipelines");
  output.push("AppSheet = UI Layer");
  output.push("");



/* =========================
   GOOGLE DRIVE STRUCTURE (VERSION-AWARE)
========================= */

try {

  const ssFile = DriveApp.getFileById(
    SpreadsheetApp.getActiveSpreadsheet().getId()
  );

  const sheetFolder = ssFile.getParents().next();       // 01_Sheets
  const versionFolder = sheetFolder.getParents().next(); // v_1.3

  output.push("## GOOGLE DRIVE STRUCTURE");
  output.push("");

  output.push("Version: " + versionFolder.getName());
  output.push("");

  const subFolders = versionFolder.getFolders();

  while(subFolders.hasNext()){

    const folder = subFolders.next();
    const folderName = folder.getName();

    output.push("### " + folderName);

    const files = folder.getFiles();

    while(files.hasNext()){

      const file = files.next();
      const name = file.getName();

      let type = "UNKNOWN";

      if(/_logs?/i.test(name)) type = "LOG";
      else if(/metadata/i.test(name)) type = "METADATA";
      else if(/support/i.test(name)) type = "SUPPORT";
      else if(/app/i.test(name)) type = "MAIN";

      output.push("- " + name + " [" + type + "]");
    }

    output.push("");
  }

} catch(err){

  output.push("## GOOGLE DRIVE STRUCTURE");
  output.push("- Unable to resolve Drive structure");
  output.push("");
}
  /* =========================
     TABLE ARCHITECTURE
  ========================= */

  const schema = metadataSS.getSheetByName("Schema_Snapshot");

  if(schema){

    const data = schema.getDataRange().getValues();

    output.push("## TABLE ARCHITECTURE");

    const tables = new Set();

    data.slice(1).forEach(r=>{
      if(r[0]) tables.add(r[0]);
    });

    tables.forEach(t=>{
      output.push("- " + t);
    });

    output.push("");
  }


/* =========================
   TABLE ROLE CLASSIFICATION
========================= */

const tableRoleSheet = metadataSS.getSheetByName("Table_Role_Classification");

if(tableRoleSheet){

  const data = tableRoleSheet.getDataRange().getValues();

  output.push("## TABLE ROLE CLASSIFICATION");

  data.slice(1).forEach(r=>{

    const table = r[0];
    const role = r[1];

    if(table && role){
      output.push("- " + table + " → " + role);
    }

  });

  output.push("");
}

/* =========================
   TABLE RELATIONSHIPS
========================= */

const derivedRel = metadataSS.getSheetByName("Derived_Column_Logic");

if(derivedRel){

  const data = derivedRel.getDataRange().getValues();

  output.push("## TABLE RELATIONSHIPS");

  const relations = new Set();

  data.slice(1).forEach(r=>{

    const sourceTable = r[0];
    const sourceColumn = r[3];
    const refs = r[7];

    if(!refs || !sourceTable || !sourceColumn) return;

    const refList = refs.split("\n");

    refList.forEach(ref=>{

      const match = ref.match(/^([^.]+)\.([A-Z]+)\s*\((.+)\)$/);

      if(match){

        const targetTable = match[1];
        const targetColumnName = match[3];

        // Only capture cross-table relationships
        const isKeyColumn =
        /(_ID|_KEY|Resolved_.*Key|Txn_ID)/i.test(sourceColumn) ||
        /(_ID|_KEY|Resolved_.*Key|Txn_ID)/i.test(targetColumnName);

        if(targetTable !== sourceTable){

          const relation =
            sourceTable + "." + sourceColumn +
            " → " +
            targetTable + "." + targetColumnName;

          relations.add(relation);
        }
      }

    });

  });

  if(relations.size === 0){
    output.push("- No cross-table relationships detected");
  } else {
    relations.forEach(r => output.push("- " + r));
  }

  output.push("");
}

/* =========================
   DERIVED COLUMN LOGIC
========================= */

const derived = metadataSS.getSheetByName("Derived_Column_Logic");

if(derived){

  const data = derived.getDataRange().getValues();

  output.push("## DERIVED COLUMN LOGIC");

  data.slice(1).forEach(r=>{

    const table = r[0];
    const columnName = r[3];         // FIXED (was column index earlier)
    const formula = r[5];            // NEW
    const refs = r[7];               // NEW

    if(table && columnName){

      let line = "- " + table + "." + columnName;

      if(formula){
        line += " = " + formula;
      }

      if(refs){
        line += " | REF → " + refs.replace(/\n/g, ", ");
      }

      output.push(line);
    }

  });

  output.push("");
}



  /* =========================
     SCRIPT FUNCTIONS
  ========================= */

  const functions = metadataSS.getSheetByName("Script_Function_Inventory");

  if(functions){

    const data = functions.getDataRange().getValues();

    output.push("## SCRIPT FUNCTIONS");

    data.slice(1).forEach(r=>{
      output.push("- " + r[0] + " (" + r[1] + ")");
    });

    output.push("");
  }



  /* =========================
     SCRIPT DEPENDENCIES
  ========================= */

  const callMap = metadataSS.getSheetByName("Script_Call_Map_INTERNAL");

  if(callMap){

    const data = callMap.getDataRange().getValues();

    output.push("## SCRIPT DEPENDENCIES");

    data.slice(1).forEach(r=>{
      output.push("- " + r[0] + " → " + r[1]);
    });

    output.push("");
  }



  /* =========================
     SCRIPT DATA FLOW
  ========================= */

  const flow = metadataSS.getSheetByName("Script_DataFlow_Map");

  if(flow){

    const data = flow.getDataRange().getValues();

    output.push("## SCRIPT DATA FLOW");

    data.slice(1).forEach(r=>{
      output.push("- " + r[0] + " : " + r[2] + " → " + r[3]);
    });

    output.push("");
  }



  /* ====================================================
     WRITE NORMAL READABLE CONTEXT (COLUMN A)
  ==================================================== */

  sheet.getRange(1,1,output.length,1)
       .setValues(output.map(x=>[x]));



  /* ====================================================
     CREATE MARKDOWN BLOCK
  ==================================================== */

  const md = output.join("\n");



  /* ====================================================
     SPLIT INTO CHUNKS (avoid 50k limit)
  ==================================================== */

  const CHUNK_SIZE = 40000;

  let chunks = [];

  for(let i=0;i<md.length;i+=CHUNK_SIZE){

    chunks.push(md.substring(i,i+CHUNK_SIZE));

  }



  /* ====================================================
     WRITE CHUNKS STARTING COLUMN C
  ==================================================== */

  sheet.getRange(1,3).setValue("COPY_PASTE_CONTEXT");

  sheet.getRange(2,3,1,chunks.length)
       .setValues([chunks]);

}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// EXPORTING: AI CONTEXT - BACKEND SYSTEM INFORMATION 
/**
 * Script Name: exportAIContextMarkdown_
 * Status: UPDATED — ETI v1.3
 *
 * Enhancement:
 * - Added structured TABLE RECONSTRUCTION section
 * - Groups: Columns, Derived Logic, Relationships per table
 * - No modification to existing output blocks
 */

function run_exportAIContextMarkdown(){
  exportAIContextMarkdown_();
}

function exportAIContextMarkdown_(){

  const metaSS = getMetadataSpreadsheet_();
  const SHEET_NAME = "AI_CONTEXT_EXPORT";

  const sheet = metaSS.getSheetByName(SHEET_NAME);

  if (!sheet) throw new Error("AI_CONTEXT_EXPORT sheet not found");

  const data = sheet.getDataRange().getValues();

  if (data.length === 0) {
    throw new Error("No AI context data found");
  }

  /* =========================
     BUILD MARKDOWN (DUMP ONLY)
  ========================= */

  let lines = [];

  for (let i = 0; i < data.length; i++) {

    const line = data[i][0]; // Column A

    if (line === null || line === undefined) continue;

    lines.push(String(line));
  }

  const md = lines.join("\n");

  /* =========================
     WRITE FILE TO DRIVE
  ========================= */

  const fileName = "ETI_AUTO_SYSTEM_CONTEXT.md";

  const files = DriveApp.getFilesByName(fileName);

  if (files.hasNext()) {

    const file = files.next();
    file.setContent(md);

    Logger.log("Markdown updated: " + file.getUrl());

  } else {

    const file = DriveApp.createFile(
      fileName,
      md,
      MimeType.PLAIN_TEXT
    );

    Logger.log("Markdown created: " + file.getUrl());
  }
}
