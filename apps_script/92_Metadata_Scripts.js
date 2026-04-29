// METADATA: SCRIPT - FUNCTION INVENTORY
function run_extractScriptFunctionInventory(){
  extractScriptFunctionInventory_();
}


function extractScriptFunctionInventory_(){

  const metadataSS = getMetadataSpreadsheet_();

  const FILE_TABLE = "Script_File_Inventory";
  const FUNC_TABLE = "Script_Function_Inventory";

  let fileSheet = metadataSS.getSheetByName(FILE_TABLE);
  if(!fileSheet) fileSheet = metadataSS.insertSheet(FILE_TABLE);

  let funcSheet = metadataSS.getSheetByName(FUNC_TABLE);
  if(!funcSheet) funcSheet = metadataSS.insertSheet(FUNC_TABLE);

  fileSheet.clear();
  funcSheet.clear();

  fileSheet.appendRow([
    "File_Name",
    "File_Type",
    "Line_Count",
    "Function_Count",
    "Detected_At"
  ]);

  funcSheet.appendRow([
    "Function_Name",
    "File_Name",
    "Line_Number",
    "Function_Type",
    "Detected_At"
  ]);


  const scriptId = ScriptApp.getScriptId();

  const response = UrlFetchApp.fetch(
    "https://script.googleapis.com/v1/projects/" + scriptId + "/content",
    {
      headers:{
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    }
  );

  const project = JSON.parse(response.getContentText());

  const files = project.files || [];

  const fileRows = [];
  const funcRows = [];

  const now = new Date();

  files.forEach(file => {

    if(file.type !== "SERVER_JS") return;

    const fileName = file.name + ".gs";
    const source = file.source || "";

    const lines = source.split("\n");

    const lineCount = lines.length;

    const functionRegex = /function\s+([A-Za-z0-9_]+)\s*\(/g;

    let match;
    let functionCount = 0;

    while((match = functionRegex.exec(source)) !== null){

      const functionName = match[1];

      const lineNumber = source
        .substring(0,match.index)
        .split("\n").length;

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
      lineCount,
      functionCount,
      now
    ]);

  });


  if(fileRows.length){

    fileSheet.getRange(
      2,
      1,
      fileRows.length,
      fileRows[0].length
    ).setValues(fileRows);

  }


  if(funcRows.length){

    funcSheet.getRange(
      2,
      1,
      funcRows.length,
      funcRows[0].length
    ).setValues(funcRows);

  }

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

/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA: SCRIPT - CALL MAP RAW
function run_generateScriptCallMap_RAW(){
  generateScriptCallMap_RAW_();
}


function generateScriptCallMap_RAW_(){

  const metadataSS = getMetadataSpreadsheet_();
  const SHEET_NAME = "Script_Call_Map_RAW";

  let sheet = metadataSS.getSheetByName(SHEET_NAME);
  if(!sheet) sheet = metadataSS.insertSheet(SHEET_NAME);

  sheet.clear();

  sheet.appendRow([
    "Caller_Function",
    "Called_Function",
    "File_Name",
    "Line_Number",
    "Detected_At"
  ]);

  const scriptId = ScriptApp.getScriptId();

  const response = UrlFetchApp.fetch(
    "https://script.googleapis.com/v1/projects/" + scriptId + "/content",
    {
      headers:{
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    }
  );

  const project = JSON.parse(response.getContentText());

  const files = project.files || [];

  const rows = [];

  const now = new Date();

  const callRegex = /([A-Za-z0-9_]+)\s*\(/g;

  const ignore = new Set([
    "if",
    "for",
    "while",
    "switch",
    "catch",
    "function",
    "return",
    "Logger",
    "console"
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

        const lineNumber = source
          .substring(0, callMatch.index)
          .split("\n").length;

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

  if(rows.length){

    sheet.getRange(
      2,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

  }

}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA: SCRIPT - CALL MAP INTERNAL
function run_generateScriptCallMap_INTERNAL(){
  generateScriptCallMap_INTERNAL_();
}


function generateScriptCallMap_INTERNAL_(){

  const metadataSS = getMetadataSpreadsheet_();

  const FUNC_SHEET = "Script_Function_Inventory";
  const OUT_SHEET = "Script_Call_Map_INTERNAL";

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

  let sheet = metadataSS.getSheetByName(OUT_SHEET);
  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Caller_Function",
    "Called_Function",
    "File_Name",
    "Line_Number",
    "Detected_At"
  ]);

  const scriptId = ScriptApp.getScriptId();

  const response = UrlFetchApp.fetch(
    "https://script.googleapis.com/v1/projects/" + scriptId + "/content",
    {
      headers:{
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    }
  );

  const project = JSON.parse(response.getContentText());

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

        const lineNumber = source
          .substring(0, callMatch.index)
          .split("\n").length;

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

  if(rows.length){

    sheet.getRange(
      2,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

  }

}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA: SCRIPT - PIPELINE MAPPING
function run_generateScriptPipelineMap(){
  generateScriptPipelineMap_();
}


function generateScriptPipelineMap_(){

  const metadataSS = getMetadataSpreadsheet_();

  const FUNC_SHEET = "Script_Function_Inventory";
  const CALL_SHEET = "Script_Call_Map_INTERNAL";
  const OUT_SHEET  = "Script_Pipeline_Map";

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

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Pipeline",
    "Step_Order",
    "Caller_Function",
    "Called_Function",
    "Called_Type",
    "File_Name",
    "Detected_At"
  ]);

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

  if(rows.length){

    sheet.getRange(
      2,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

  }

}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA: SCRIPT - ARCHITECTURE LOGIC
function run_generateScriptArchitectureLogic(){
  generateScriptArchitectureLogic_();
}


function generateScriptArchitectureLogic_(){

  const metadataSS = getMetadataSpreadsheet_();

  const FUNC_SHEET = "Script_Function_Inventory";
  const CALL_SHEET = "Script_Call_Map_INTERNAL";
  const PIPE_SHEET = "Script_Pipeline_Map";
  const OUT_SHEET  = "Script_Architecture_Logic";

  const funcSheet = metadataSS.getSheetByName(FUNC_SHEET);
  const callSheet = metadataSS.getSheetByName(CALL_SHEET);
  const pipeSheet = metadataSS.getSheetByName(PIPE_SHEET);

  if(!funcSheet) throw new Error("Script_Function_Inventory missing");
  if(!callSheet) throw new Error("Script_Call_Map_INTERNAL missing");
  if(!pipeSheet) throw new Error("Script_Pipeline_Map missing");

  const funcData = funcSheet.getDataRange().getValues();
  const callData = callSheet.getDataRange().getValues();
  const pipeData = pipeSheet.getDataRange().getValues();

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Architecture_Level",
    "Parent_Node",
    "Node_Name",
    "Node_Type",
    "Source_File",
    "Detected_At"
  ]);

  const rows = [];

  const now = new Date();

  /* --------------------------------------------------
     Build Function → Type Map
  -------------------------------------------------- */

  const funcTypeMap = {};
  const funcFileMap = {};

  for(let i=1;i<funcData.length;i++){

    const fn   = funcData[i][0];
    const file = funcData[i][1];
    const type = funcData[i][3];

    funcTypeMap[fn] = type;
    funcFileMap[fn] = file;

  }

  /* --------------------------------------------------
     Level 1 — Pipelines
  -------------------------------------------------- */

  const pipelines = Object.keys(funcTypeMap)
    .filter(fn => fn.startsWith("pipeline_") || fn.includes("pipeline"));

  pipelines.forEach(pipeline => {

    rows.push([
      1,
      "SYSTEM",
      pipeline,
      "PIPELINE",
      funcFileMap[pipeline] || "",
      now
    ]);

  });

  /* --------------------------------------------------
     Level 2 — Pipeline Steps
  -------------------------------------------------- */

  pipeData.forEach((row,i)=>{

    if(i===0) return;

    const pipeline = row[0];
    const called   = row[3];

    rows.push([
      2,
      pipeline,
      called,
      funcTypeMap[called] || "UNKNOWN",
      funcFileMap[called] || "",
      now
    ]);

  });

  /* --------------------------------------------------
     Level 3 — Utility Calls
  -------------------------------------------------- */

  callData.forEach((row,i)=>{

    if(i===0) return;

    const caller = row[0];
    const called = row[1];

    const type = funcTypeMap[called];

    if(!type) return;

    if(type === "UTILITY" || type === "LOGGER"){

      rows.push([
        3,
        caller,
        called,
        type,
        funcFileMap[called] || "",
        now
      ]);

    }

  });

  /* --------------------------------------------------
     Write Results
  -------------------------------------------------- */

  if(rows.length){

    sheet.getRange(
      2,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

  }

}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA: SCRIPT - ARCHITECTURE DIAGRAM
function run_generateScriptArchitectureDiagram(){
  generateScriptArchitectureDiagram_();
}


function generateScriptArchitectureDiagram_(){

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

  const rows = [[
    "MERMAID",
    mermaid,
    new Date()
  ]];

  sheet.getRange(2,1,1,3).setValues(rows);

}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA: SCRIPT - FUNCTION CODE SUMMARY
function run_extractScriptFunctionCodeSummary(){
  extractScriptFunctionCodeSummary_();
}


function extractScriptFunctionCodeSummary_(){

  const metadataSS = getMetadataSpreadsheet_();

  const OUT_SHEET = "Script_Function_Code_Summary";

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

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

  const scriptId = ScriptApp.getScriptId();

  const response = UrlFetchApp.fetch(
    "https://script.googleapis.com/v1/projects/" + scriptId + "/content",
    {
      headers:{
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    }
  );

  const project = JSON.parse(response.getContentText());

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

      const lineNumber = source
        .substring(0, startIndex)
        .split("\n").length;

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

  if(rows.length){

    sheet.getRange(
      2,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

  }

}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA: SCRIPT - DATA FLOW MAP
function run_extractScriptDataFlowMap(){
  extractScriptDataFlowMap_();
}


function extractScriptDataFlowMap_(){

  const metadataSS = getMetadataSpreadsheet_();

  const OUT_SHEET = "Script_DataFlow_Map";

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Function_Name",
    "File_Name",
    "Operation_Type",
    "Sheet_Reference",
    "Detected_Line",
    "Detected_At"
  ]);


  const scriptId = ScriptApp.getScriptId();

  const response = UrlFetchApp.fetch(
    "https://script.googleapis.com/v1/projects/" + scriptId + "/content",
    {
      headers:{
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    }
  );

  const project = JSON.parse(response.getContentText());

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

  if(rows.length){

    sheet.getRange(
      2,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

  }

}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA: SCRIPT - PERFORMANCE MAP
function run_extractScriptPerformanceMap(){
  extractScriptPerformanceMap_();
}


function extractScriptPerformanceMap_(){

  const metadataSS = getMetadataSpreadsheet_();
  const OUT_SHEET = "Script_Performance_Map";

  let sheet = metadataSS.getSheetByName(OUT_SHEET);
  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Function_Name",
    "File_Name",
    "Performance_Flag",
    "Detected_Line",
    "Line_Number",
    "Detected_At"
  ]);


  const scriptId = ScriptApp.getScriptId();

  const response = UrlFetchApp.fetch(
    "https://script.googleapis.com/v1/projects/" + scriptId + "/content",
    {
      headers:{
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    }
  );

  const project = JSON.parse(response.getContentText());

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

  if(rows.length){

    sheet.getRange(
      2,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

  }

}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// METADATA: SCRIPT - INTERACTION MAP
function run_extractScriptSheetInteractionMap(){
  extractScriptSheetInteractionMap_();
}


function extractScriptSheetInteractionMap_(){

  const metadataSS = getMetadataSpreadsheet_();
  const SHEET_NAME = "Script_Sheet_Interaction_Map";

  let sheet = metadataSS.getSheetByName(SHEET_NAME);
  if(!sheet) sheet = metadataSS.insertSheet(SHEET_NAME);

  sheet.clear();

  sheet.appendRow([
    "Function_Name",
    "Sheet_Name",
    "Operation",
    "File_Name",
    "Line_Number",
    "Detected_At"
  ]);

  const scriptId = ScriptApp.getScriptId();

  const response = UrlFetchApp.fetch(
    "https://script.googleapis.com/v1/projects/" + scriptId + "/content",
    {
      headers:{
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    }
  );

  const project = JSON.parse(response.getContentText());
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

        const assignMatch = trimmed.match(/(const|let|var)\s+([A-Za-z0-9_]+)\s*=\s*.*getSheetByName\(["'](.+?)["']\)/);

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

  if(rows.length){

    sheet.getRange(
      2,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

  }

}



