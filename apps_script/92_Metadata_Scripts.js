/* 
===============================================
SCRIPT FUNCTION INVENTORY
===============================================*/
/**
 * Script Name: extractScriptFunctionInventory_
 *
 * Purpose:
 * - Extract all script files from Apps Script project
 * - Identify and catalog all functions across files
 * - Build two inventories:
 *   1. File-level metadata (file size, function count)
 *   2. Function-level metadata (name, location, type)
 *
 * Explicit Non-Goals:
 * - Does NOT analyze function logic or dependencies
 * - Does NOT modify script files
 * - Does NOT validate function correctness
 * - Does NOT interact with Sheets data layer (only metadata layer)
 *
 * Execution Flow:
 * 1. Initialize execution context and logging
 * 2. Initialize / clear output sheets (File + Function inventory)
 * 3. Fetch full script project content via Apps Script API
 * 4. Iterate through all script files (SERVER_JS only)
 * 5. For each file:
 *    a. Split source into lines
 *    b. Detect functions using regex
 *    c. Capture function name and line number
 *    d. Classify function type
 *    e. Count functions per file
 * 6. Build file-level and function-level output arrays
 * 7. Write both outputs in batch to respective sheets
 * 8. Log summary and complete execution
 *
 * Input Dependencies:
 * - Apps Script Project (via ScriptApp + API)
 * - Utility: fetchScriptProject_()
 * - Utility: classifyFunctionType_()
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Fetch script project once (cached per execution)
 * 2. Process all files in memory
 * 3. Use regex scan for function detection
 * 4. Build output arrays incrementally
 * 5. Write results using single batch setValues()
 *
 * Output Contract:
 * - Sheet: Script_File_Inventory
 *   Columns:
 *     File_Name | File_Type | Line_Count | Function_Count | Detected_At
 *
 * - Sheet: Script_Function_Inventory
 *   Columns:
 *     Function_Name | File_Name | Line_Number | Function_Type | Detected_At
 *
 * Idempotency:
 * - Safe to re-run; both sheets are cleared and rebuilt deterministically
 */

function extractScriptFunctionInventory_(){

  const SCRIPT_NAME   = 'Metadata_Scripts';
  const FUNCTION_NAME = 'extractScriptFunctionInventory_';
  const TGT_SHEET     = 'Script_Function_Inventory';

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

    const FILE_TABLE = "Script_File_Inventory";
    const FUNC_TABLE = "Script_Function_Inventory";

    let fileSheet = metadataSS.getSheetByName(FILE_TABLE);
    if(!fileSheet) fileSheet = metadataSS.insertSheet(FILE_TABLE);

    let funcSheet = metadataSS.getSheetByName(FUNC_TABLE);
    if(!funcSheet) funcSheet = metadataSS.insertSheet(FUNC_TABLE);

    fileSheet.clear();
    funcSheet.clear();

    const fileHeaders = [
      "File_Name","File_Type","Line_Count","Function_Count","Detected_At"
    ];

    const funcHeaders = [
      "Function_Name","File_Name","Line_Number","Function_Type","Detected_At"
    ];

    fileSheet.getRange(1,1,1,fileHeaders.length).setValues([fileHeaders]);
    funcSheet.getRange(1,1,1,funcHeaders.length).setValues([funcHeaders]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    /*
    -----------------
    STEP — PROCESS
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    const project = fetchScriptProject_();
    const files = project.files || [];

    const fileRows = [];
    const funcRows = [];
    const now = new Date();

    files.forEach(file => {

      if(file.type !== "SERVER_JS") return;

      const fileName = file.name + ".gs";
      const source = file.source || "";
      const lines = source.split("\n");

      const functionRegex = /function\s+([A-Za-z0-9_]+)\s*\(/g;

      let match;
      let functionCount = 0;

      while((match = functionRegex.exec(source)) !== null){

        const functionName = match[1];

        const lineNumber = source.substring(0,match.index).split("\n").length;

        const type = classifyFunctionType_(functionName);

        funcRows.push([
          functionName,
          fileName,
          lineNumber,
          type,
          now
        ]);

        functionCount++;
      }

      fileRows.push([
        fileName,
        "SERVER_JS",
        lines.length,
        functionCount,
        now
      ]);

    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    /*
    -----------------
    STEP — WRITE_OUTPUT
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if(fileRows.length){
      fileSheet.getRange(2,1,fileRows.length,fileRows[0].length).setValues(fileRows);
    }

    if(funcRows.length){
      funcSheet.getRange(2,1,funcRows.length,funcRows[0].length).setValues(funcRows);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Files=${fileRows.length} | Functions=${funcRows.length} | DurationMs=${new Date()-t0}`
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
SCRIPT CALL MAP RAW
===============================================*/
/**
 * Script Name: generateScriptCallMap_RAW_
 *
 * Purpose:
 * - Identify raw function-to-function call relationships across all script files
 * - Capture direct call occurrences without validating against known function inventory
 * - Provide foundational call graph for further internal filtering and analysis
 *
 * Explicit Non-Goals:
 * - Does NOT validate called functions against inventory (handled in INTERNAL map)
 * - Does NOT resolve nested or indirect call chains
 * - Does NOT differentiate between internal vs external libraries beyond basic filtering
 * - Does NOT modify or execute any script logic
 *
 * Execution Flow:
 * 1. Initialize execution context and logging
 * 2. Initialize / clear output sheet (Script_Call_Map_RAW)
 * 3. Fetch full script project content via Apps Script API
 * 4. Iterate all script files (SERVER_JS only)
 * 5. For each file:
 *    a. Identify all function definitions using regex
 *    b. Extract function body (naive block via first closing brace)
 *    c. Scan body for function call patterns using regex
 *    d. Filter out language constructs and ignored keywords
 *    e. Exclude self-referential calls
 *    f. Capture caller → called relationship with line number
 * 6. Build output array of call relationships
 * 7. Write output in single batch
 * 8. Log summary and complete execution
 *
 * Input Dependencies:
 * - Apps Script Project (via fetchScriptProject_)
 * - Utility: fetchScriptProject_()
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Fetch script project once (cached per execution)
 * 2. Iterate files in memory
 * 3. Use functionRegex to locate caller functions
 * 4. Extract approximate function body using index slicing
 * 5. Use callRegex to detect all function-like calls
 * 6. Filter using ignore set + self-call exclusion
 * 7. Compute line numbers via substring split
 * 8. Accumulate rows in memory
 * 9. Write once using setValues()
 *
 * Output Contract:
 * - Sheet: Script_Call_Map_RAW
 *   Columns:
 *     Caller_Function | Called_Function | File_Name | Line_Number | Detected_At
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function generateScriptCallMap_RAW_(){

  const SCRIPT_NAME = 'Metadata_Scripts';
  const FUNCTION_NAME = 'generateScriptCallMap_RAW_';
  const TGT_SHEET = 'Script_Call_Map_RAW';

  const t0 = new Date();

  try {

    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    const metadataSS = getMetadataSpreadsheet_();
    let sheet = metadataSS.getSheetByName(TGT_SHEET);
    if(!sheet) sheet = metadataSS.insertSheet(TGT_SHEET);

    sheet.clear();

    const headers = [
      "Caller_Function","Called_Function","File_Name","Line_Number","Detected_At"
    ];

    sheet.getRange(1,1,1,headers.length).setValues([headers]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    const project = fetchScriptProject_();
    const files = project.files || [];

    const rows = [];
    const now = new Date();

    const callRegex = /([A-Za-z0-9_]+)\s*\(/g;

    const ignore = new Set([
      "if","for","while","switch","catch","function","return","Logger","console"
    ]);

    files.forEach(file => {

      if(file.type !== "SERVER_JS") return;

      const source = file.source || "";
      const fileName = file.name + ".gs";

      const funcRegex = /function\s+([A-Za-z0-9_]+)\s*\(/g;

      let funcMatch;

      while((funcMatch = funcRegex.exec(source)) !== null){

        const caller = funcMatch[1];
        const bodyStart = funcMatch.index;
        const bodyEnd = source.indexOf("}", bodyStart);
        const body = source.substring(bodyStart, bodyEnd);

        let callMatch;

        while((callMatch = callRegex.exec(body)) !== null){

          const called = callMatch[1];

          if(ignore.has(called)) continue;
          if(called === caller) continue;

          const lineNumber = source.substring(0, callMatch.index).split("\n").length;

          rows.push([
            caller,
            called,
            fileName,
            lineNumber,
            now
          ]);
        }
      }
    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if(rows.length){
      sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${rows.length} | DurationMs=${new Date()-t0}`
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
SCRIPT CALL MAP INTERNAL
===============================================*/
/**
 * Script Name: generateScriptCallMap_INTERNAL_
 *
 * Purpose:
 * - Generate validated function-to-function call relationships
 * - Filter RAW call map to retain only calls between known script functions
 * - Produce clean internal call graph for downstream pipeline and architecture mapping
 *
 * Explicit Non-Goals:
 * - Does NOT include calls to external libraries or built-in JS functions
 * - Does NOT resolve multi-level or indirect call chains
 * - Does NOT perform dependency ordering or execution sequencing
 * - Does NOT modify script or metadata sources
 *
 * Execution Flow:
 * 1. Initialize execution context and logging
 * 2. Load Script_Function_Inventory and build valid function set
 * 3. Initialize / clear output sheet (Script_Call_Map_INTERNAL)
 * 4. Fetch full script project content via Apps Script API
 * 5. Iterate all script files (SERVER_JS only)
 * 6. For each file:
 *    a. Detect function definitions (callers)
 *    b. Extract function body (naive block)
 *    c. Scan for function call patterns
 *    d. Filter:
 *       - Exclude self calls
 *       - Include only calls present in valid function set
 *    e. Capture caller → called relationships with line numbers
 * 7. Build output array
 * 8. Write output in single batch
 * 9. Log summary and complete execution
 *
 * Input Dependencies:
 * - Sheet: Script_Function_Inventory (required)
 * - Apps Script Project (via fetchScriptProject_)
 * - Utility: fetchScriptProject_()
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Load function inventory into Set for O(1) lookup
 * 2. Fetch script project once (cached per execution)
 * 3. Iterate files and detect functions using regex
 * 4. Extract approximate function body via index slicing
 * 5. Use callRegex to detect calls
 * 6. Filter using:
 *    - self exclusion (called !== caller)
 *    - membership check (validFunctions.has)
 * 7. Compute line numbers via substring split
 * 8. Accumulate rows in memory
 * 9. Write once using setValues()
 *
 * Output Contract:
 * - Sheet: Script_Call_Map_INTERNAL
 *   Columns:
 *     Caller_Function | Called_Function | File_Name | Line_Number | Detected_At
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function generateScriptCallMap_INTERNAL_(){

  const SCRIPT_NAME = 'Metadata_Scripts';
  const FUNCTION_NAME = 'generateScriptCallMap_INTERNAL_';
  const TGT_SHEET = 'Script_Call_Map_INTERNAL';

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

    const FUNC_SHEET = "Script_Function_Inventory";

    const funcSheet = metadataSS.getSheetByName(FUNC_SHEET);

    if(!funcSheet){
      throw new Error("Script_Function_Inventory not found");
    }

    const funcData = funcSheet.getDataRange().getValues();

    const validFunctions = new Set();

    for(let i=1;i<funcData.length;i++){
      const fn = funcData[i][0];
      if(fn) validFunctions.add(fn);
    }

    let sheet = metadataSS.getSheetByName(TGT_SHEET);
    if(!sheet) sheet = metadataSS.insertSheet(TGT_SHEET);

    sheet.clear();

    const headers = [
      "Caller_Function","Called_Function","File_Name","Line_Number","Detected_At"
    ];

    sheet.getRange(1,1,1,headers.length).setValues([headers]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    /*
    -----------------
    STEP — PROCESS
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    const project = fetchScriptProject_();
    const files = project.files || [];

    const rows = [];
    const now = new Date();

    const callRegex = /([A-Za-z0-9_]+)\s*\(/g;

    files.forEach(file => {

      if(file.type !== "SERVER_JS") return;

      const source = file.source || "";
      const fileName = file.name + ".gs";

      const funcRegex = /function\s+([A-Za-z0-9_]+)\s*\(/g;

      let funcMatch;

      while((funcMatch = funcRegex.exec(source)) !== null){

        const caller = funcMatch[1];

        const bodyStart = funcMatch.index;
        const bodyEnd = source.indexOf("}", bodyStart);
        const body = source.substring(bodyStart, bodyEnd);

        let callMatch;

        while((callMatch = callRegex.exec(body)) !== null){

          const called = callMatch[1];

          if(called === caller) continue;
          if(!validFunctions.has(called)) continue;

          const lineNumber =
            source.substring(0, callMatch.index).split("\n").length;

          rows.push([
            caller,
            called,
            fileName,
            lineNumber,
            now
          ]);

        }

      }

    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    /*
    -----------------
    STEP — WRITE_OUTPUT
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if(rows.length){
      sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${rows.length} | DurationMs=${new Date()-t0}`
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
SCRIPT PIPELINE MAP
===============================================*/
/**
 * Script Name: generateScriptPipelineMap_
 *
 * Purpose:
 * - Identify pipeline functions and map their direct function calls
 * - Generate ordered step list for each pipeline based on call map
 * - Provide a flat, pipeline-scoped view of execution steps
 *
 * Explicit Non-Goals:
 * - Does NOT resolve execution dependencies or actual runtime order
 * - Does NOT traverse nested or indirect function calls
 * - Does NOT validate correctness of pipeline structure
 * - Does NOT infer control flow (only direct call relationships)
 *
 * Execution Flow:
 * 1. Initialize execution context and logging
 * 2. Load Script_Function_Inventory and Script_Call_Map_INTERNAL
 * 3. Build function → type lookup map (funcTypeMap)
 * 4. Identify pipeline functions based on naming convention
 * 5. For each pipeline function:
 *    a. Initialize step counter
 *    b. Iterate all call map rows
 *    c. Filter rows where caller === pipeline
 *    d. Assign step order sequentially based on encounter
 *    e. Derive called function type
 *    f. Append row to output
 * 6. Write output in single batch
 * 7. Log summary and complete execution
 *
 * Input Dependencies:
 * - Sheet: Script_Function_Inventory (Function_Name, Function_Type)
 * - Sheet: Script_Call_Map_INTERNAL (Caller → Called relationships)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Build funcTypeMap for O(1) lookup
 * 2. Identify pipelines via string matching (startsWith / includes)
 * 3. Iterate pipelines × full callData scan
 * 4. Filter direct calls (caller === pipeline)
 * 5. Assign step order incrementally (no sorting)
 * 6. Accumulate rows in memory
 * 7. Write once using setValues()
 *
 * Output Contract:
 * - Sheet: Script_Pipeline_Map
 *   Columns:
 *     Pipeline | Step_Order | Caller_Function | Called_Function | Called_Type | File_Name | Detected_At
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function generateScriptPipelineMap_(){

  const SCRIPT_NAME = 'Metadata_Scripts';
  const FUNCTION_NAME = 'generateScriptPipelineMap_';
  const TGT_SHEET = 'Script_Pipeline_Map';

  const t0 = new Date();

  try {

    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    const metadataSS = getMetadataSpreadsheet_();

    const FUNC_SHEET = "Script_Function_Inventory";
    const CALL_SHEET = "Script_Call_Map_INTERNAL";

    const funcSheet = metadataSS.getSheetByName(FUNC_SHEET);
    const callSheet = metadataSS.getSheetByName(CALL_SHEET);

    if(!funcSheet) throw new Error("Script_Function_Inventory missing");
    if(!callSheet) throw new Error("Script_Call_Map_INTERNAL missing");

    const funcData = funcSheet.getDataRange().getValues();
    const callData = callSheet.getDataRange().getValues();

    const funcTypeMap = {};

    for(let i=1;i<funcData.length;i++){
      const fn = funcData[i][0];
      const type = funcData[i][3];
      funcTypeMap[fn] = type;
    }

    let sheet = metadataSS.getSheetByName(TGT_SHEET);
    if(!sheet) sheet = metadataSS.insertSheet(TGT_SHEET);

    sheet.clear();

    const headers = [
      "Pipeline","Step_Order","Caller_Function","Called_Function","Called_Type","File_Name","Detected_At"
    ];

    sheet.getRange(1,1,1,headers.length).setValues([headers]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    const rows = [];
    const now = new Date();

    const pipelineFunctions = Object.keys(funcTypeMap)
      .filter(fn => fn.startsWith("pipeline_") || fn.includes("pipeline"));

    pipelineFunctions.forEach(pipeline => {

      let step = 1;

      callData.forEach((row,i)=>{

        if(i===0) return;

        const caller = row[0];
        const called = row[1];
        const file = row[2];

        if(caller !== pipeline) return;

        const type = funcTypeMap[called] || "UNKNOWN";

        rows.push([
          pipeline,
          step,
          caller,
          called,
          type,
          file,
          now
        ]);

        step++;

      });

    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if(rows.length){
      sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${rows.length} | DurationMs=${new Date()-t0}`
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
SCRIPT ARCHITECTURE LOGIC
===============================================*/
/**
 * Script Name: generateScriptArchitectureLogic_
 *
 * Purpose:
 * - Construct a hierarchical architecture representation of the script system
 * - Organize functions into 3 levels:
 *   1. System → Pipeline functions
 *   2. Pipeline → Direct called functions
 *   3. Caller → Utility / Logger functions
 *
 * Explicit Non-Goals:
 * - Does NOT compute full dependency graph or recursive call chains
 * - Does NOT deduplicate nodes across levels
 * - Does NOT validate correctness of architecture relationships
 * - Does NOT infer execution flow or control flow logic
 *
 * Execution Flow:
 * 1. Initialize execution context and logging
 * 2. Load:
 *    - Script_Function_Inventory
 *    - Script_Call_Map_INTERNAL
 *    - Script_Pipeline_Map
 * 3. Build:
 *    - funcTypeMap (Function → Type)
 *    - funcFileMap (Function → Source File)
 * 4. Identify pipeline functions via naming convention
 * 5. Construct architecture levels:
 *    a. Level 1:
 *       SYSTEM → pipeline functions
 *    b. Level 2:
 *       pipeline → called functions (from Script_Pipeline_Map)
 *    c. Level 3:
 *       caller → called functions where type is UTILITY or LOGGER (from Call Map)
 * 6. Append all rows into output array
 * 7. Write output in single batch
 * 8. Log summary and complete execution
 *
 * Input Dependencies:
 * - Sheet: Script_Function_Inventory (Function_Name, File_Name, Function_Type)
 * - Sheet: Script_Call_Map_INTERNAL (Caller → Called relationships)
 * - Sheet: Script_Pipeline_Map (Pipeline-level mapping)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Build funcTypeMap and funcFileMap for O(1) lookups
 * 2. Identify pipelines using naming rule (startsWith / includes "pipeline")
 * 3. Level 1:
 *    Add SYSTEM → pipeline rows
 * 4. Level 2:
 *    Iterate pipeData and map pipeline → called function
 * 5. Level 3:
 *    Iterate callData and include only UTILITY / LOGGER types
 * 6. Resolve file names via funcFileMap
 * 7. Accumulate rows in memory
 * 8. Write once using setValues()
 *
 * Output Contract:
 * - Sheet: Script_Architecture_Logic
 *   Columns:
 *     Architecture_Level | Parent_Node | Node_Name | Node_Type | Source_File | Detected_At
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function generateScriptArchitectureLogic_(){

  const SCRIPT_NAME = 'Metadata_Scripts';
  const FUNCTION_NAME = 'generateScriptArchitectureLogic_';
  const TGT_SHEET = 'Script_Architecture_Logic';

  const t0 = new Date();

  try {

    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    const metadataSS = getMetadataSpreadsheet_();

    const funcSheet = metadataSS.getSheetByName("Script_Function_Inventory");
    const callSheet = metadataSS.getSheetByName("Script_Call_Map_INTERNAL");
    const pipeSheet = metadataSS.getSheetByName("Script_Pipeline_Map");

    if(!funcSheet) throw new Error("Script_Function_Inventory missing");
    if(!callSheet) throw new Error("Script_Call_Map_INTERNAL missing");
    if(!pipeSheet) throw new Error("Script_Pipeline_Map missing");

    const funcData = funcSheet.getDataRange().getValues();
    const callData = callSheet.getDataRange().getValues();
    const pipeData = pipeSheet.getDataRange().getValues();

    let sheet = metadataSS.getSheetByName(TGT_SHEET);
    if(!sheet) sheet = metadataSS.insertSheet(TGT_SHEET);

    sheet.clear();

    const headers = [
      "Architecture_Level","Parent_Node","Node_Name","Node_Type","Source_File","Detected_At"
    ];

    sheet.getRange(1,1,1,headers.length).setValues([headers]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    const rows = [];
    const now = new Date();

    const funcTypeMap = {};
    const funcFileMap = {};

    for(let i=1;i<funcData.length;i++){
      const fn   = funcData[i][0];
      const file = funcData[i][1];
      const type = funcData[i][3];
      funcTypeMap[fn] = type;
      funcFileMap[fn] = file;
    }

    const pipelines = Object.keys(funcTypeMap)
      .filter(fn => fn.startsWith("pipeline_") || fn.includes("pipeline"));

    pipelines.forEach(pipeline => {
      rows.push([1,"SYSTEM",pipeline,"PIPELINE",funcFileMap[pipeline] || "",now]);
    });

    pipeData.forEach((row,i)=>{
      if(i===0) return;
      rows.push([2,row[0],row[3],funcTypeMap[row[3]] || "UNKNOWN",funcFileMap[row[3]] || "",now]);
    });

    callData.forEach((row,i)=>{
      if(i===0) return;
      const called = row[1];
      const type = funcTypeMap[called];
      if(type === "UTILITY" || type === "LOGGER"){
        rows.push([3,row[0],called,type,funcFileMap[called] || "",now]);
      }
    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if(rows.length){
      sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${rows.length} | DurationMs=${new Date()-t0}`
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
SCRIPT ARCHITECTURE DIAGRAM
===============================================*/
/**
 * Script Name: generateScriptArchitectureDiagram_
 *
 * Purpose:
 * - Convert Script_Architecture_Logic into a Mermaid diagram representation
 * - Generate a visual graph of function relationships using parent → child mapping
 * - Provide a single consolidated diagram for system architecture visualization
 *
 * Explicit Non-Goals:
 * - Does NOT compute or modify architecture relationships
 * - Does NOT deduplicate nodes or edges
 * - Does NOT preserve architecture levels in diagram (levels ignored)
 * - Does NOT generate multiple diagram formats (only Mermaid)
 *
 * Execution Flow:
 * 1. Initialize execution context and logging
 * 2. Load Script_Architecture_Logic sheet
 * 3. Initialize / clear output sheet (Script_Architecture_Diagram)
 * 4. Iterate all rows in architecture data:
 *    a. Extract Parent_Node and Node_Name
 *    b. Skip invalid rows (missing values)
 *    c. Build edge string:
 *       - SYSTEM → node (for top level)
 *       - parent → node (for others)
 * 5. Combine edges into Mermaid graph definition:
 *    - Prefix with "graph TD"
 *    - Append all edges line-by-line
 * 6. Write single output row with diagram definition
 * 7. Log summary and complete execution
 *
 * Input Dependencies:
 * - Sheet: Script_Architecture_Logic (Parent_Node, Node_Name)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Read full architecture dataset into memory
 * 2. Iterate rows sequentially (skip header)
 * 3. Extract parent and node values
 * 4. Skip rows with missing values
 * 5. Construct edge strings (parent --> node)
 * 6. Accumulate edges in array (no deduplication)
 * 7. Join into Mermaid string using newline separator
 * 8. Write single row output using setValues()
 *
 * Output Contract:
 * - Sheet: Script_Architecture_Diagram
 *   Columns:
 *     Diagram_Type | Diagram_Definition | Generated_At
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function generateScriptArchitectureDiagram_(){

  const SCRIPT_NAME = 'Metadata_Scripts';
  const FUNCTION_NAME = 'generateScriptArchitectureDiagram_';
  const TGT_SHEET = 'Script_Architecture_Diagram';

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

    const ARCH_SHEET = "Script_Architecture_Logic";
    const OUT_SHEET  = "Script_Architecture_Diagram";

    const archSheet = metadataSS.getSheetByName(ARCH_SHEET);

    if(!archSheet){
      throw new Error("Script_Architecture_Logic sheet missing");
    }

    const data = archSheet.getDataRange().getValues();

    let sheet = metadataSS.getSheetByName(OUT_SHEET);
    if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

    sheet.clear();

    sheet.appendRow([
      "Diagram_Type",
      "Diagram_Definition",
      "Generated_At"
    ]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    /*
    -----------------
    STEP — PROCESS
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    const edges = [];

    for(let i=1;i<data.length;i++){

      const parent = data[i][1];
      const node   = data[i][2];

      if(!parent || !node) continue;

      if(parent === "SYSTEM"){
        edges.push(`SYSTEM --> ${node}`);
      } else {
        edges.push(`${parent} --> ${node}`);
      }
    }

    const mermaid = [
      "graph TD",
      ...edges
    ].join("\n");

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    /*
    -----------------
    STEP — WRITE_OUTPUT
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    const rows = [[
      "MERMAID",
      mermaid,
      new Date()
    ]];

    sheet.getRange(2,1,1,3).setValues(rows);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Edges=${edges.length} | DurationMs=${new Date()-t0}`
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
SCRIPT FUNCTION CODE SUMMARY
===============================================*/
/**
 * Script Name: extractScriptFunctionCodeSummary_
 *
 * Purpose:
 * - Extract structural and descriptive metadata for each function in the script
 * - Provide lightweight code profiling including parameters, size, return usage, and comments
 * - Enable basic function-level insights without deep parsing
 *
 * Explicit Non-Goals:
 * - Does NOT perform full syntax parsing or AST analysis
 * - Does NOT correctly handle nested function blocks or complex scopes
 * - Does NOT validate correctness of function structure
 * - Does NOT extract inline or single-line comments (//)
 *
 * Execution Flow:
 * 1. Initialize execution context and logging
 * 2. Initialize / clear output sheet (Script_Function_Code_Summary)
 * 3. Fetch full script project content via Apps Script API
 * 4. Iterate all script files (SERVER_JS only)
 * 5. For each file:
 *    a. Detect function definitions using regex
 *    b. Extract function name and parameter string
 *    c. Compute parameter count
 *    d. Extract function body (naive substring until first closing brace)
 *    e. Determine:
 *       - presence of "return"
 *       - approximate line count
 *    f. Calculate line number from source
 *    g. Extract first  block comment from body
 *    h. Clean and truncate comment text
 *    i. Append row to output
 * 6. Write output in single batch
 * 7. Log summary and complete execution
 *
 * Input Dependencies:
 * - Apps Script Project (via fetchScriptProject_)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Fetch script project once (cached per execution)
 * 2. Use funcRegex to detect functions
 * 3. Extract parameters via regex capture group
 * 4. Count parameters using string split
 * 5. Extract body using first closing brace index (naive)
 * 6. Compute:
 *    - hasReturn via string inclusion
 *    - approxLines via newline split
 * 7. Extract first block comment using regex
 * 8. Normalize comment (remove newlines, asterisks, trim, truncate)
 * 9. Accumulate rows in memory
 * 10. Write once using setValues()
 *
 * Output Contract:
 * - Sheet: Script_Function_Code_Summary
 *   Columns:
 *     Function_Name | File_Name | Line_Number | Parameter_Count | Has_Return |
 *     Approx_Line_Count | First_Comment | Detected_At
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function extractScriptFunctionCodeSummary_(){

  const SCRIPT_NAME = 'Metadata_Scripts';
  const FUNCTION_NAME = 'extractScriptFunctionCodeSummary_';
  const TGT_SHEET = 'Script_Function_Code_Summary';

  const t0 = new Date();

  try {

    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    const metadataSS = getMetadataSpreadsheet_();

    let sheet = metadataSS.getSheetByName(TGT_SHEET);
    if(!sheet) sheet = metadataSS.insertSheet(TGT_SHEET);

    sheet.clear();

    sheet.appendRow([
      "Function_Name",
      "File_Name",
      "Line_Number",
      "Parameter_Count",
      "Has_Return",
      "Approx_Line_Count",
      "First_Comment",
      "Detected_At"
    ]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    const project = fetchScriptProject_();
    const files = project.files || [];

    const rows = [];
    const now = new Date();

    const funcRegex = /function\s+([A-Za-z0-9_]+)\s*\(([^)]*)\)/g;

    files.forEach(file => {

      if(file.type !== "SERVER_JS") return;

      const source = file.source || "";
      const fileName = file.name + ".gs";

      let match;

      while((match = funcRegex.exec(source)) !== null){

        const fnName = match[1];
        const params = match[2];

        const paramCount = params.trim() === "" ? 0 : params.split(",").length;

        const startIndex = match.index;
        const bodyEnd = source.indexOf("}", startIndex);
        const body = source.substring(startIndex, bodyEnd);

        const hasReturn = body.includes("return");
        const approxLines = body.split("\n").length;

        const lineNumber =
          source.substring(0, startIndex).split("\n").length;

        let comment = "";

        const commentRegex = /\/\*\*([\s\S]*?)\*\//;
        const commentMatch = commentRegex.exec(body);

        if(commentMatch){
          comment = commentMatch[1]
            .replace(/\n/g," ")
            .replace(/\*/g,"")
            .trim()
            .substring(0,200);
        }

        rows.push([
          fnName,
          fileName,
          lineNumber,
          paramCount,
          hasReturn,
          approxLines,
          comment,
          now
        ]);
      }
    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if(rows.length){
      sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${rows.length} | DurationMs=${new Date()-t0}`
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
SCRIPT DATA FLOW MAP
===============================================*/
/**
 * Script Name: extractScriptDataFlowMap_
 *
 * Purpose:
 * - Detect and log Google Sheets interaction patterns within script functions
 * - Identify read/write operations and sheet access at line-level granularity
 * - Provide a basic data interaction map for audit and analysis
 *
 * Explicit Non-Goals:
 * - Does NOT perform syntax parsing or variable resolution
 * - Does NOT extract actual sheet names (stores raw line instead)
 * - Does NOT detect multiple operations per line (single operation overwrite logic)
 * - Does NOT handle nested blocks or complex control structures
 *
 * Execution Flow:
 * 1. Initialize execution context and logging
 * 2. Initialize / clear output sheet (Script_DataFlow_Map)
 * 3. Fetch full script project content via Apps Script API
 * 4. Iterate all script files (SERVER_JS only)
 * 5. For each file:
 *    a. Detect function definitions using regex
 *    b. Extract function body (naive substring until first closing brace)
 *    c. Split body into lines
 *    d. For each line:
 *       - Identify operation using string matching
 *       - Determine operation type (single match, overwrite behavior)
 *       - Compute line number
 *       - Capture raw line as reference
 *       - Append row if operation detected
 * 6. Write output in single batch
 * 7. Log summary and complete execution
 *
 * Input Dependencies:
 * - Apps Script Project (via fetchScriptProject_)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Fetch script project once (cached per execution)
 * 2. Use funcRegex to detect functions
 * 3. Extract function body via naive substring
 * 4. Split body into lines
 * 5. Apply sequential string matching for operation detection
 *    (last matching condition overwrites previous)
 * 6. Compute line number using offset + index
 * 7. Store trimmed line as reference
 * 8. Accumulate rows in memory
 * 9. Write once using setValues()
 *
 * Output Contract:
 * - Sheet: Script_DataFlow_Map
 *   Columns:
 *     Function_Name | File_Name | Operation_Type | Sheet_Reference | Detected_Line | Detected_At
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function extractScriptDataFlowMap_(){

  const SCRIPT_NAME = 'Metadata_Scripts';
  const FUNCTION_NAME = 'extractScriptDataFlowMap_';
  const TGT_SHEET = 'Script_DataFlow_Map';

  const t0 = new Date();

  try {

    getOrInitExecutionContext_({ function_name: FUNCTION_NAME });
    ETI_logStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET);

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    const metadataSS = getMetadataSpreadsheet_();

    let sheet = metadataSS.getSheetByName(TGT_SHEET);
    if(!sheet) sheet = metadataSS.insertSheet(TGT_SHEET);

    sheet.clear();

    sheet.appendRow([
      "Function_Name",
      "File_Name",
      "Operation_Type",
      "Sheet_Reference",
      "Detected_Line",
      "Detected_At"
    ]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    const project = fetchScriptProject_();
    const files = project.files || [];

    const rows = [];
    const now = new Date();

    const funcRegex = /function\s+([A-Za-z0-9_]+)\s*\(/g;

    files.forEach(file => {

      if(file.type !== "SERVER_JS") return;

      const source = file.source || "";
      const fileName = file.name + ".gs";

      let funcMatch;

      while((funcMatch = funcRegex.exec(source)) !== null){

        const fnName = funcMatch[1];

        const bodyStart = funcMatch.index;
        const bodyEnd = source.indexOf("}", bodyStart);

        const body = source.substring(bodyStart, bodyEnd);
        const lines = body.split("\n");

        lines.forEach((line,i)=>{

          let operation = "";

          if(line.includes("getSheetByName")) operation = "SHEET_ACCESS";
          if(line.includes("getRange")) operation = "RANGE_ACCESS";
          if(line.includes("getValues")) operation = "READ_VALUES";
          if(line.includes("getValue")) operation = "READ_VALUE";
          if(line.includes("setValues")) operation = "WRITE_VALUES";
          if(line.includes("setValue")) operation = "WRITE_VALUE";
          if(line.includes("appendRow")) operation = "APPEND_ROW";
          if(line.includes("clear")) operation = "CLEAR_RANGE";

          if(operation){

            const lineNumber =
              source.substring(0, funcMatch.index)
              .split("\n").length + i;

            rows.push([
              fnName,
              fileName,
              operation,
              line.trim(),
              lineNumber,
              now
            ]);
          }
        });
      }
    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if(rows.length){
      sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${rows.length} | DurationMs=${new Date()-t0}`
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
SCRIPT PERFORMANCE MAP
===============================================*/
/**
 * Script Name: extractScriptPerformanceMap_
 *
 * Purpose:
 * - Identify potential performance anti-patterns in script functions
 * - Detect expensive operations inside loops and inefficient write patterns
 * - Provide a heuristic-based performance audit at line-level granularity
 *
 * Explicit Non-Goals:
 * - Does NOT perform accurate loop scope tracking or nesting analysis
 * - Does NOT detect loop termination or exit conditions
 * - Does NOT measure actual runtime performance or execution cost
 * - Does NOT detect batching opportunities or optimize code
 *
 * Execution Flow:
 * 1. Initialize execution context and logging
 * 2. Initialize / clear output sheet (Script_Performance_Map)
 * 3. Fetch full script project content via Apps Script API
 * 4. Iterate all script files (SERVER_JS only)
 * 5. For each file:
 *    a. Detect function definitions using regex
 *    b. Extract function body (naive substring until first closing brace)
 *    c. Split body into lines
 *    d. Track loop presence using insideLoop flag (state-based)
 *    e. For each line:
 *       - Detect loop start (for / while)
 *       - Identify performance patterns:
 *         • getValue inside loop → READ_IN_LOOP
 *         • setValue inside loop → WRITE_IN_LOOP
 *         • getRange inside loop → RANGE_ACCESS_IN_LOOP
 *         • appendRow anywhere → APPEND_ROW_USAGE
 *       - Assign single flag (last match wins)
 *       - Compute line number
 *       - Append row if flag detected
 * 6. Write output in single batch
 * 7. Log summary and complete execution
 *
 * Input Dependencies:
 * - Apps Script Project (via fetchScriptProject_)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Fetch script project once (cached per execution)
 * 2. Use funcRegex to detect functions
 * 3. Extract function body via naive substring
 * 4. Split body into lines
 * 5. Maintain insideLoop flag (no reset logic)
 * 6. Apply sequential string matching for flags
 *    (last matching condition overwrites previous)
 * 7. Compute line numbers using offset + index
 * 8. Store trimmed line as reference
 * 9. Accumulate rows in memory
 * 10. Write once using setValues()
 *
 * Output Contract:
 * - Sheet: Script_Performance_Map
 *   Columns:
 *     Function_Name | File_Name | Performance_Flag | Detected_Line | Line_Number | Detected_At
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function extractScriptPerformanceMap_(){

  const SCRIPT_NAME = 'Metadata_Scripts';
  const FUNCTION_NAME = 'extractScriptPerformanceMap_';
  const TGT_SHEET = 'Script_Performance_Map';

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

    let sheet = metadataSS.getSheetByName(TGT_SHEET);
    if(!sheet) sheet = metadataSS.insertSheet(TGT_SHEET);

    sheet.clear();

    sheet.appendRow([
      "Function_Name",
      "File_Name",
      "Performance_Flag",
      "Detected_Line",
      "Line_Number",
      "Detected_At"
    ]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    /*
    -----------------
    STEP — PROCESS
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    const project = fetchScriptProject_();
    const files = project.files || [];

    const rows = [];
    const now = new Date();

    const funcRegex = /function\s+([A-Za-z0-9_]+)\s*\(/g;

    files.forEach(file => {

      if(file.type !== "SERVER_JS") return;

      const source = file.source || "";
      const fileName = file.name + ".gs";

      let funcMatch;

      while((funcMatch = funcRegex.exec(source)) !== null){

        const fnName = funcMatch[1];

        const bodyStart = funcMatch.index;
        const bodyEnd = source.indexOf("}", bodyStart);

        const body = source.substring(bodyStart, bodyEnd);
        const lines = body.split("\n");

        let insideLoop = false;

        lines.forEach((line,i)=>{

          const trimmed = line.trim();

          if(
            trimmed.startsWith("for(") ||
            trimmed.startsWith("for (") ||
            trimmed.startsWith("while(") ||
            trimmed.startsWith("while (")
          ){
            insideLoop = true;
          }

          let flag = "";

          if(trimmed.includes("getValue(") && insideLoop){
            flag = "READ_IN_LOOP";
          }

          if(trimmed.includes("setValue(") && insideLoop){
            flag = "WRITE_IN_LOOP";
          }

          if(trimmed.includes("appendRow(")){
            flag = "APPEND_ROW_USAGE";
          }

          if(trimmed.includes("getRange(") && insideLoop){
            flag = "RANGE_ACCESS_IN_LOOP";
          }

          if(flag){

            const lineNumber =
              source.substring(0, funcMatch.index)
              .split("\n").length + i;

            rows.push([
              fnName,
              fileName,
              flag,
              trimmed,
              lineNumber,
              now
            ]);
          }

        });

      }

    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    /*
    -----------------
    STEP — WRITE_OUTPUT
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if(rows.length){
      sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${rows.length} | DurationMs=${new Date()-t0}`
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
SCRIPT SHEET INTERACTION MAP
===============================================*/
/**
 * Script Name: extractScriptSheetInteractionMap_
 *
 * Purpose:
 * - Detect and map sheet-level interactions within script functions
 * - Identify read/write operations and associate them with specific sheets
 * - Provide a structured view of how script functions interact with spreadsheet data
 *
 * Explicit Non-Goals:
 * - Does NOT perform full variable resolution or scope tracking
 * - Does NOT detect indirect or chained sheet references
 * - Does NOT support complex assignments or reassignment of sheet variables
 * - Does NOT capture multiple operations per line (single overwrite logic)
 *
 * Execution Flow:
 * 1. Initialize execution context and logging
 * 2. Initialize / clear output sheet (Script_Sheet_Interaction_Map)
 * 3. Fetch full script project content via Apps Script API
 * 4. Iterate all script files (SERVER_JS only)
 * 5. For each file:
 *    a. Detect function definitions using regex
 *    b. Extract function body (naive substring until first closing brace)
 *    c. Split body into lines
 *    d. Track sheet variables:
 *       - Identify assignments using getSheetByName("SheetName")
 *       - Map variable → sheet name
 *    e. For each line:
 *       - Detect operation using string matching
 *       - Determine operation type (single match, overwrite behavior)
 *       - Resolve sheet:
 *         • Match variable usage (varName.)
 *         • Else assign "UNKNOWN"
 *       - Compute line number
 *       - Append row if operation detected
 * 6. Write output in single batch
 * 7. Log summary and complete execution
 *
 * Input Dependencies:
 * - Apps Script Project (via fetchScriptProject_)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Fetch script project once (cached per execution)
 * 2. Use funcRegex to detect functions
 * 3. Extract function body via naive substring
 * 4. Split body into lines
 * 5. Build sheetVars map using regex match on getSheetByName
 * 6. Apply sequential string matching for operation detection
 *    (last matching condition overwrites previous)
 * 7. Resolve sheet using variable prefix matching (v.)
 * 8. Compute line numbers using offset + index
 * 9. Accumulate rows in memory
 * 10. Write once using setValues()
 *
 * Output Contract:
 * - Sheet: Script_Sheet_Interaction_Map
 *   Columns:
 *     Function_Name | Sheet_Name | Operation | File_Name | Line_Number | Detected_At
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */
function extractScriptSheetInteractionMap_(){

  const SCRIPT_NAME = 'Metadata_Scripts';
  const FUNCTION_NAME = 'extractScriptSheetInteractionMap_';
  const TGT_SHEET = 'Script_Sheet_Interaction_Map';

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

    let sheet = metadataSS.getSheetByName(TGT_SHEET);
    if(!sheet) sheet = metadataSS.insertSheet(TGT_SHEET);

    sheet.clear();

    sheet.appendRow([
      "Function_Name",
      "Sheet_Name",
      "Operation",
      "File_Name",
      "Line_Number",
      "Detected_At"
    ]);

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'LOAD');

    /*
    -----------------
    STEP — PROCESS
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    const project = fetchScriptProject_();
    const files = project.files || [];

    const rows = [];
    const now = new Date();

    const funcRegex = /function\s+([A-Za-z0-9_]+)\s*\(/g;

    files.forEach(file => {

      if(file.type !== "SERVER_JS") return;

      const source = file.source || "";
      const fileName = file.name + ".gs";

      let funcMatch;

      while((funcMatch = funcRegex.exec(source)) !== null){

        const fnName = funcMatch[1];

        const bodyStart = funcMatch.index;
        const bodyEnd = source.indexOf("}", bodyStart);

        const body = source.substring(bodyStart, bodyEnd);
        const lines = body.split("\n");

        const sheetVars = {};

        lines.forEach((line,i)=>{

          const trimmed = line.trim();

          const assignMatch = trimmed.match(
            /(const|let|var)\s+([A-Za-z0-9_]+)\s*=\s*.*getSheetByName\(["'](.+?)["']\)/
          );

          if(assignMatch){
            const variable = assignMatch[2];
            const sheetName = assignMatch[3];
            sheetVars[variable] = sheetName;
          }

          let operation = "";

          if(trimmed.includes("setValue(")) operation = "WRITE_CELL";
          if(trimmed.includes("setValues(")) operation = "WRITE_RANGE";
          if(trimmed.includes("appendRow(")) operation = "APPEND_ROW";
          if(trimmed.includes("clear(")) operation = "CLEAR";
          if(trimmed.includes("clearContents(")) operation = "CLEAR_CONTENT";
          if(trimmed.includes("getValues(")) operation = "READ_RANGE";
          if(trimmed.includes("getValue(")) operation = "READ_CELL";

          if(operation){

            let detectedSheet = "UNKNOWN";

            Object.keys(sheetVars).forEach(v=>{
              if(trimmed.includes(v + ".")){
                detectedSheet = sheetVars[v];
              }
            });

            const lineNumber =
              source.substring(0, funcMatch.index)
              .split("\n").length + i;

            rows.push([
              fnName,
              detectedSheet,
              operation,
              fileName,
              lineNumber,
              now
            ]);
          }

        });

      }

    });

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'PROCESS');

    /*
    -----------------
    STEP — WRITE_OUTPUT
    -----------------
    */
    ETI_logStepStart_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    if(rows.length){
      sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
    }

    ETI_logStepEnd_(SCRIPT_NAME, FUNCTION_NAME, TGT_SHEET, 'WRITE_OUTPUT');

    ETI_logSummary_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      TGT_SHEET,
      `Rows=${rows.length} | DurationMs=${new Date()-t0}`
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
UTILITY FUNCTIONS — SCRIPT METADATA
===============================================*/
/**
 * Fetch Apps Script project content (cached per execution)
 * NO LOGIC DRIFT — centralizes repeated API call
 */
var SCRIPT_PROJECT_CACHE = null;

function fetchScriptProject_(){

  if (SCRIPT_PROJECT_CACHE) return SCRIPT_PROJECT_CACHE;

  const scriptId = ScriptApp.getScriptId();

  const response = UrlFetchApp.fetch(
    "https://script.googleapis.com/v1/projects/" + scriptId + "/content",
    {
      headers:{
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    }
  );

  SCRIPT_PROJECT_CACHE = JSON.parse(response.getContentText());

  return SCRIPT_PROJECT_CACHE;
}

function classifyFunctionType_(name){

  if(name.startsWith("pipeline_")) return "PIPELINE";

  if(name.includes("Controller")) return "CONTROLLER";

  if(name.includes("process")) return "PROCESSOR";

  if(name.includes("populate")) return "PROCESSOR";

  if(name.includes("promote")) return "PROCESSOR";

  if(name.includes("cleanup")) return "PROCESSOR";

  if(name.includes("metadata")) return "METADATA";

  if(name.includes("access")) return "ACCESS";

  if(name.includes("log")) return "LOGGER";

  return "UTILITY";

}