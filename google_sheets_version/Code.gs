/***************
 * CONFIG
 ***************/
const MONTHLY_INPUTS_FOLDER_ID = "PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE"; // Monthly_Inputs folder in Drive

const SHEET_GL = "GL";
const SHEET_FINAL = "Final";
const SHEET_QA = "Missing_GL_Mapping";

// Output schema (7 columns)
const FINAL_HEADERS = ["GL Code", "Description", "Category", "Year", "Month", "Department", "Amount"];

// Department sheet name pattern: "DEPARTMENT XXX-F"
const DEPT_SHEET_REGEX = /^DEPARTMENT\s+(\d+)-F$/i;

// File name pattern: "mm.yyyy ..."
const FILE_MMYYYY_REGEX = /(\d{2})\.(\d{4})/;


/***************
 * UI: menu
 ***************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("EXAMPLE_COMPANY Warehouse")
    .addItem("Run Monthly Update", "runMonthlyUpdate")
    .addItem("Initialize/Repair Sheets", "initializeWarehouseSheets")
    .addToUi();
}


/***************
 * MAIN
 ***************/
function runMonthlyUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  initializeWarehouseSheets();

  // Find newest .xlsx in Monthly_Inputs
  const folder = DriveApp.getFolderById(MONTHLY_INPUTS_FOLDER_ID);
  const latestXlsx = getLatestXlsx_(folder);
  if (!latestXlsx) {
    throw new Error("No .xlsx files found in Monthly_Inputs.");
  }

  const { monthNum, year } = parseMonthYearFromFilename_(latestXlsx.getName());
  const monthName = monthNumToName_(monthNum);

  // Convert XLSX -> temp Google Sheet
  const tempSheetFileId = convertXlsxToGoogleSheet_(latestXlsx);
  try {
    const tempSs = SpreadsheetApp.openById(tempSheetFileId);

    // Load GL map from warehouse
    const glMap = loadGlMap_(ss);

    // Parse all dept sheets from temp workbook
    const newRows = parseAllDepartments_(tempSs, glMap, year, monthNum, monthName);

    // Append + dedupe into Final
    const finalSheet = ss.getSheetByName(SHEET_FINAL);
    const existingRows = readFinalRows_(finalSheet);
    const updatedRows = appendAndDedupe_(existingRows, newRows);

    writeFinalRows_(finalSheet, updatedRows);

    // QA: write missing GL codes (only for this run)
    const qaSheet = ss.getSheetByName(SHEET_QA);
    writeQaMissing_(qaSheet, newRows);

    SpreadsheetApp.getUi().alert(
      `Done!\n\nLoaded file: ${latestXlsx.getName()}\nAdded/updated rows: ${newRows.length}\nFinal rows now: ${updatedRows.length}`
    );

  } finally {
    // Trash the temporary converted file
    try {
      DriveApp.getFileById(tempSheetFileId).setTrashed(true); 
    } catch (e) {
      // If trashing fails, don't block the main workflow
      console.log("Warning: could not trash temp file: " + e);
    }
  }
}


/***************
 * SETUP / REPAIR
 ***************/
function initializeWarehouseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure GL exists
  if (!ss.getSheetByName(SHEET_GL)) {
    ss.insertSheet(SHEET_GL);
  }

  // Ensure Final exists + headers
  let finalSheet = ss.getSheetByName(SHEET_FINAL);
  if (!finalSheet) finalSheet = ss.insertSheet(SHEET_FINAL);
  ensureHeaders_(finalSheet, FINAL_HEADERS);

  // Ensure QA exists + headers
  let qaSheet = ss.getSheetByName(SHEET_QA);
  if (!qaSheet) qaSheet = ss.insertSheet(SHEET_QA);
  ensureHeaders_(qaSheet, [...FINAL_HEADERS, "Missing GL in Reference?"]);
}


/***************
 * DRIVE: get newest XLSX + convert
 ***************/
function getLatestXlsx_(folder) {
  const files = folder.getFiles();
  let latest = null;
  let latestTime = 0;

  while (files.hasNext()) {
    const f = files.next();
    const name = (f.getName() || "").toLowerCase();

    // only xlsx, exclude anything that looks like warehouse
    if (!name.endsWith(".xlsx")) continue;
    if (name.includes("data warehouse")) continue;

    const updated = f.getLastUpdated().getTime();
    if (updated > latestTime) {
      latestTime = updated;
      latest = f;
    }
  }
  return latest;
}

function convertXlsxToGoogleSheet_(xlsxFile) {
  const blob = xlsxFile.getBlob();

  const resource = {
    name: "[TEMP] " + xlsxFile.getName(),
    mimeType: MimeType.GOOGLE_SHEETS
  };

  const created = Drive.Files.create(resource, blob);
  return created.id;
}



/***************
 * PARSING: filename month/year, dept sheets, amounts, category
 ***************/
function parseMonthYearFromFilename_(filename) {
  const m = filename.match(FILE_MMYYYY_REGEX);
  if (!m) throw new Error(`Could not find mm.yyyy in filename: ${filename}`);
  const monthNum = parseInt(m[1], 10);
  const year = parseInt(m[2], 10);
  if (!(monthNum >= 1 && monthNum <= 12)) throw new Error(`Month out of range: ${monthNum}`);
  return { monthNum, year };
}

function parseAllDepartments_(tempSs, glMap, year, monthNum, monthName) {
  const rows = [];
  const sheets = tempSs.getSheets();

  for (const sh of sheets) {
    const dept = extractDepartment_(sh.getName());
    if (!dept) continue;

    // Read values. We only care about first 3 columns.
    // Your Excel has a “row 1 header/title” to ignore, and row 2 is the true headers.
    // So start reading from row 2 onward; simplest: read all, then ignore first row.
    const data = sh.getDataRange().getValues(); // includes headers/title
    if (!data || data.length < 3) continue;

    // Drop first row (title), treat next row as column headers, and parse from row index 2
    // We do NOT rely on column header text; we assume col A,B,C are NUMBER, DESCRIPTION, ACTUAL.
    const body = data.slice(2); // skip row 1 + row 2 headers

    let currentCategory = null;

    for (const r of body) {
      const numberCell = (r[0] ?? "").toString().trim();
      const descCell = (r[1] ?? "").toString().trim();
      const amtCell = r[2];

      // Category marker rows
      if (numberCell.toUpperCase() === "REVENUES") {
        currentCategory = "Revenue";
        continue;
      }
      if (numberCell.toUpperCase() === "EXPENSES") {
        currentCategory = "Expenses";
        continue;
      }

      // Only keep 4-digit GL codes
      if (!/^\d{4}$/.test(numberCell)) continue;

      const glCode = numberCell;
      const amount = parseAmount_(amtCell);
      if (amount === null) continue;

      const canonicalDesc = glMap[glCode] || "";
      rows.push([
        glCode,
        canonicalDesc,
        currentCategory || "",
        year,
        monthName, // store month word
        dept,
        amount
      ]);
    }
  }

  return rows;
}

function extractDepartment_(sheetName) {
  const m = sheetName.trim().match(DEPT_SHEET_REGEX);
  return m ? m[1] : null;
}

function parseAmount_(value) {
  if (value === null || value === "" || typeof value === "undefined") return null;

  // If already numeric
  if (typeof value === "number") return value;

  let s = value.toString().trim();
  if (!s) return null;

  // Remove $ and commas
  s = s.replace(/\$/g, "").replace(/,/g, "");

  // Parentheses indicate negative
  let neg = false;
  if (/^\(.*\)$/.test(s)) {
    neg = true;
    s = s.replace(/[()]/g, "").trim();
  }

  // Handle stray spaces
  const num = Number(s);
  if (Number.isNaN(num)) return null;
  return neg ? -num : num;
}

function monthNumToName_(monthNum) {
  return [
    "", "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ][monthNum];
}

function monthNameToNum_(monthName) {
  const names = {
    january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
    july: 7, august: 8, september: 9, october: 10, november: 11, december: 12
  };
  const key = (monthName || "").toString().trim().toLowerCase();
  return names[key] || null;
}


/***************
 * GL MAP: read reference table from GL tab
 ***************/
function loadGlMap_(warehouseSs) {
  const glSheet = warehouseSs.getSheetByName(SHEET_GL);
  const values = glSheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    throw new Error("GL sheet is empty. Add GL codes + descriptions to the GL tab.");
  }

  // Try to identify GL code and description columns
  const headers = values[0].map(h => (h ?? "").toString().trim().toLowerCase());
  let glIdx = headers.findIndex(h => ["gl", "gl code", "glcode", "number", "account", "account number"].includes(h));
  let descIdx = headers.findIndex(h => ["description", "gl description", "account description", "name"].includes(h));

  // Fallback: assume first 2 columns are (GL, Description)
  if (glIdx === -1) glIdx = 0;
  if (descIdx === -1) descIdx = 1;

  const map = {};
  for (let i = 1; i < values.length; i++) {
    const gl = (values[i][glIdx] ?? "").toString().trim();
    const desc = (values[i][descIdx] ?? "").toString().trim();
    if (/^\d{4}$/.test(gl)) {
      map[gl] = desc;
    }
  }
  return map;
}


/***************
 * FINAL: read, append+dedupe, write
 ***************/
function readFinalRows_(finalSheet) {
  const data = finalSheet.getDataRange().getValues();
  if (!data || data.length < 2) return [];
  // assume headers in row 1
  return data.slice(1);
}

function appendAndDedupe_(existingRows, newRows) {
  // Build a map by natural key so reruns don’t duplicate:
  // GL Code + Year + Month + Department + Category
  const outMap = {};

  const upsert = (row) => {
    const gl = (row[0] ?? "").toString().trim();
    const desc = (row[1] ?? "").toString();
    const cat = (row[2] ?? "").toString().trim();
    const year = (row[3] ?? "").toString().trim();
    const month = (row[4] ?? "").toString().trim(); // stored as month name
    const dept = (row[5] ?? "").toString().trim();
    const amt = row[6];

    // Normalize month to a number for stable dedupe if you ever change text/case
    const monthNum = monthNameToNum_(month) || month;

    const key = [gl, year, monthNum, dept, cat].join("|");
    outMap[key] = [gl, desc, cat, Number(year), month, dept, amt];
  };

  existingRows.forEach(upsert);
  newRows.forEach(upsert);

  // Return as an array
  return Object.values(outMap).sort((a, b) => {
    if (a[3] !== b[3]) return a[3] - b[3]; // Year
    const am = monthNameToNum_(a[4]) || 0;
    const bm = monthNameToNum_(b[4]) || 0;
    if (am !== bm) return am - bm;
    if (a[5] !== b[5]) return a[5].localeCompare(b[5]); // Department
    if (a[2] !== b[2]) return a[2].localeCompare(b[2]); // Category
    return a[0].localeCompare(b[0]); // GL Code
  });
}

function writeFinalRows_(finalSheet, rows) {
  finalSheet.clearContents();
  ensureHeaders_(finalSheet, FINAL_HEADERS);

  if (!rows || rows.length === 0) return;

  finalSheet.getRange(2, 1, rows.length, FINAL_HEADERS.length).setValues(rows);
  finalSheet.autoResizeColumns(1, FINAL_HEADERS.length);
}


/***************
 * QA: Missing GL mappings
 ***************/
function writeQaMissing_(qaSheet, newRows) {
  qaSheet.clearContents();
  ensureHeaders_(qaSheet, [...FINAL_HEADERS, "Missing GL in Reference?"]);

  const missing = newRows
    .filter(r => !r[1] || r[1].toString().trim() === "") // Description blank => missing mapping
    .map(r => [...r, "YES"]);

  if (missing.length > 0) {
    qaSheet.getRange(2, 1, missing.length, FINAL_HEADERS.length + 1).setValues(missing);
    qaSheet.autoResizeColumns(1, FINAL_HEADERS.length + 1);
  }
}


/***************
 * UTIL
 ***************/
function ensureHeaders_(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const normalized = current.map(c => (c ?? "").toString().trim());
  const expected = headers.map(h => h.trim());

  let same = true;
  for (let i = 0; i < expected.length; i++) {
    if ((normalized[i] || "") !== expected[i]) {
      same = false;
      break;
    }
  }
  if (!same) {
    sheet.getRange(1, 1, 1, headers.length).setValues([expected]);
  }
}
