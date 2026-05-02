/* 
===============================================
GENERATE SYSTEM CONTEXT(AI USE) VIA METADATA
===============================================*/
/**
 * Script Name: generateAIContext_
 *
 * Purpose:
 * - Generate a consolidated AI-readable system context from metadata layer
 * - Aggregate Sheets, Script, and Drive structure into a single structured document
 * - Provide both human-readable and markdown-ready context output
 *
 * Explicit Non-Goals:
 * - Does NOT perform any transformation or inference beyond available metadata
 * - Does NOT modify any source data or metadata tables
 * - Does NOT validate or enforce schema integrity
 * - Does NOT interact with AppSheet or transactional pipelines
 *
 * Execution Flow:
 * 1. Initialize metadata spreadsheet and output sheet (AI_CONTEXT_EXPORT)
 * 2. Clear existing output
 * 3. Initialize in-memory output array
 * 4. Append static system header and architecture description
 * 5. Extract and append Google Drive structure:
 *    - Resolve version folder dynamically
 *    - Iterate subfolders and files
 *    - Classify files (LOG / METADATA / SUPPORT / MAIN)
 * 6. Extract and append Table Architecture:
 *    - Read Schema_Snapshot
 *    - Build unique table list
 * 7. Extract and append Table Role Classification:
 *    - Read Table_Role_Classification
 *    - Map table → role
 * 8. Extract and append Table Relationships:
 *    - Read Derived_Column_Logic
 *    - Parse REF column
 *    - Capture cross-table relationships only
 * 9. Extract and append Derived Column Logic:
 *    - Read Derived_Column_Logic
 *    - Combine table + column + formula + references
 * 10. Extract and append Script Functions:
 *     - Read Script_Function_Inventory
 * 11. Extract and append Script Dependencies:
 *     - Read Script_Call_Map_INTERNAL
 * 12. Extract and append Script Data Flow:
 *     - Read Script_DataFlow_Map
 * 13. Write full readable context to Column A (single batch)
 * 14. Generate markdown string from output array
 * 15. Split markdown into chunks (40k limit)
 * 16. Write chunks starting Column C
 *
 * Input Dependencies:
 * - Schema_Snapshot
 * - Table_Role_Classification
 * - Derived_Column_Logic
 * - Script_Function_Inventory
 * - Script_Call_Map_INTERNAL
 * - Script_DataFlow_Map
 * - Google Drive (for structure extraction)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Build output entirely in-memory using sequential push operations
 * 2. Use Set() for deduplication (tables, relationships)
 * 3. Parse references using regex on REF field
 * 4. Avoid any nested sheet writes during processing
 * 5. Perform single batch write using setValues()
 * 6. Generate markdown via join("\n")
 * 7. Chunk markdown using fixed-size slicing (40k)
 * 8. Write chunk row in single batch operation
 *
 * Output Contract:
 * - Sheet: AI_CONTEXT_EXPORT
 *   Column A → Full readable system context
 *   Column C+ → Chunked markdown blocks (copy-paste ready)
 *
 * Idempotency:
 * - Safe to re-run; output sheet is cleared and rebuilt deterministically
 */
function generateAIContext_(){

  const SCRIPT_NAME   = 'Metadata_AI';
  const FUNCTION_NAME = 'generateAIContext_';
  const TGT_SHEET     = 'AI_CONTEXT_EXPORT';

  const t0 = new Date();

  try {

    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    /*
    -----------------
    STEP — LOAD
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    const metadataSS = getMetadataSpreadsheet_();

    const OUT_SHEET = "AI_CONTEXT_EXPORT";

    let sheet = metadataSS.getSheetByName(OUT_SHEET);

    if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

    sheet.clear();

    let output = [];

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');


    /*
    -----------------
    STEP — BUILD_CONTEXT
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'BUILD_CONTEXT');

    output.push("# ETI AUTO SYSTEM CONTEXT");
    output.push("");

    output.push("Architecture:");
    output.push("Google Sheets = Data Layer");
    output.push("Apps Script = Processing Pipelines");
    output.push("AppSheet = UI Layer");
    output.push("");


    /*
    -----------------
    SUBSTEP — GOOGLE DRIVE STRUCTURE
    -----------------
    */
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


    /*
    -----------------
    SUBSTEP — TABLE ARCHITECTURE
    -----------------
    */
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


    /*
    -----------------
    SUBSTEP — TABLE ROLE CLASSIFICATION
    -----------------
    */
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


    /*
    -----------------
    SUBSTEP — TABLE RELATIONSHIPS
    -----------------
    */
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
        /*const isKeyColumn =
        /(_ID|_KEY|Resolved_.*Key|Txn_ID)/i.test(sourceColumn) ||
        /(_ID|_KEY|Resolved_.*Key|Txn_ID)/i.test(targetColumnName);*/

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


    /*
    -----------------
    SUBSTEP — DERIVED COLUMN LOGIC
    -----------------
    */
    const derived = metadataSS.getSheetByName("Derived_Column_Logic");

    if(derived){

      const data = derived.getDataRange().getValues();

      output.push("## DERIVED COLUMN LOGIC");

      data.slice(1).forEach(r=>{

        const table = r[0];
        const columnName = r[3];
        const formula = r[5];
        const refs = r[7];

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


    /*
    -----------------
    SUBSTEP — SCRIPT FUNCTIONS
    -----------------
    */
    const functions = metadataSS.getSheetByName("Script_Function_Inventory");

    if(functions){

      const data = functions.getDataRange().getValues();

      output.push("## SCRIPT FUNCTIONS");

      data.slice(1).forEach(r=>{
        output.push("- " + r[0] + " (" + r[1] + ")");
      });

      output.push("");
    }


    /*
    -----------------
    SUBSTEP — SCRIPT DEPENDENCIES
    -----------------
    */
    const callMap = metadataSS.getSheetByName("Script_Call_Map_INTERNAL");

    if(callMap){

      const data = callMap.getDataRange().getValues();

      output.push("## SCRIPT DEPENDENCIES");

      data.slice(1).forEach(r=>{
        output.push("- " + r[0] + " → " + r[1]);
      });

      output.push("");
    }


    /*
    -----------------
    SUBSTEP — SCRIPT DATA FLOW
    -----------------
    */
    const flow = metadataSS.getSheetByName("Script_DataFlow_Map");

    if(flow){

      const data = flow.getDataRange().getValues();

      output.push("## SCRIPT DATA FLOW");

      data.slice(1).forEach(r=>{
        output.push("- " + r[0] + " : " + r[2] + " → " + r[3]);
      });

      output.push("");
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'BUILD_CONTEXT');


    /*
    -----------------
    STEP — WRITE_OUTPUT
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    sheet.getRange(1,1,output.length,1)
         .setValues(output.map(x=>[x]));

    // Create Markdown block
    const md = output.join("\n");

    // Split Write into Chunks <50K to avoid GAS limits
    const CHUNK_SIZE = 40000;

    let chunks = [];

    for(let i=0;i<md.length;i+=CHUNK_SIZE){
      chunks.push(md.substring(i,i+CHUNK_SIZE));
    }

    // Write Chunks starting Column C
    sheet.getRange(1,3).setValue("COPY_PASTE_CONTEXT");

    sheet.getRange(2,3,1,chunks.length)
         .setValues([chunks]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');


    const durationMs = new Date() - t0;

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Lines=${output.length} | Chunks=${chunks.length} | DurationMs=${durationMs}`
    );

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

  } catch(err){

    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, err, 'MAIN');
    throw err;

  } finally {
    flushLogs_();
  }
}


/* 
===============================================
EXPORT BACKEND SYSTEM CONTEXT (AI USE)
===============================================*/
/**
 * Script Name: exportAIContextMarkdown_
 *
 * Purpose:
 * - Export generated AI context from metadata sheet to Google Drive
 * - Convert Column A content into a markdown file
 * - Maintain a single up-to-date AI context file in Drive
 *
 * Explicit Non-Goals:
 * - Does NOT generate or modify AI context content
 * - Does NOT transform or enrich metadata
 * - Does NOT validate structure of context data
 * - Does NOT interact with metadata generation pipeline logic
 *
 * Execution Flow:
 * 1. Initialize metadata spreadsheet reference
 * 2. Load AI_CONTEXT_EXPORT sheet
 * 3. Validate sheet existence
 * 4. Read full sheet data
 * 5. Validate data availability
 * 6. Iterate Column A:
 *    - Skip null / undefined rows
 *    - Convert each row to string
 *    - Append to lines array
 * 7. Join lines into markdown string using newline separator
 * 8. Check if markdown file exists in Drive
 * 9. If exists:
 *    - Overwrite file content
 * 10. If not exists:
 *    - Create new file with markdown content
 * 11. Log resulting file URL
 *
 * Input Dependencies:
 * - AI_CONTEXT_EXPORT (Sheet)
 * - Google Drive (file read/write access)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Read entire sheet once using getDataRange().getValues()
 * 2. Iterate rows sequentially (single pass)
 * 3. Build lines array in memory
 * 4. Join array into markdown string (join("\n"))
 * 5. Use DriveApp.getFilesByName() for existence check
 * 6. Update or create file accordingly
 *
 * Output Contract:
 * - File: ETI_AUTO_SYSTEM_CONTEXT.md
 * - Location: Google Drive (default root or existing file location)
 * - Content: Full markdown representation of Column A
 *
 * Idempotency:
 * - Safe to re-run; file is overwritten deterministically
 */
function exportAIContextMarkdown_(){

  const SCRIPT_NAME   = 'Metadata_AI';
  const FUNCTION_NAME = 'exportAIContextMarkdown_';
  const TGT_SHEET     = 'AI_CONTEXT_EXPORT';

  const t0 = new Date();

  try {

    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    /*
    -----------------
    STEP — LOAD
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    const metaSS = getMetadataSpreadsheet_();
    const SHEET_NAME = "AI_CONTEXT_EXPORT";

    const sheet = metaSS.getSheetByName(SHEET_NAME);

    if (!sheet) throw new Error("AI_CONTEXT_EXPORT sheet not found");

    const data = sheet.getDataRange().getValues();

    if (data.length === 0) {
      throw new Error("No AI context data found");
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');


    /*
    -----------------
    STEP — BUILD_MARKDOWN
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'BUILD_MARKDOWN');

    let lines = [];

    for (let i = 0; i < data.length; i++) {

    const line = data[i][0]; // Column A

      if (line === null || line === undefined) continue;

      lines.push(String(line));
    }

    const md = lines.join("\n");

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'BUILD_MARKDOWN');


    /*
    -----------------
    STEP — WRITE_FILE
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_FILE');

    const fileName = "ETI_AUTO_SYSTEM_CONTEXT.md";

    const files = DriveApp.getFilesByName(fileName);

    let fileUrl = "";

    if (files.hasNext()) {

      const file = files.next();
      file.setContent(md);

      fileUrl = file.getUrl();

      Logger.log("Markdown updated: " + fileUrl);

    } else {

      const file = DriveApp.createFile(
        fileName,
        md,
        MimeType.PLAIN_TEXT
      );

      fileUrl = file.getUrl();

      Logger.log("Markdown created: " + fileUrl);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_FILE');


    const durationMs = new Date() - t0;

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Lines=${lines.length} | DurationMs=${durationMs}`
    );

    ETI_logEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

  } catch(err){

    ETI_logError_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, err, 'MAIN');
    throw err;

  } finally {
    flushLogs_();
  }
}