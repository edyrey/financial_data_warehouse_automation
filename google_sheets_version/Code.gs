/***************
 * CONFIG
 ***************/
const INCOME_INPUTS_FOLDER_ID  = "PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE";   // Income Statement files folder
const BALANCE_INPUTS_FOLDER_ID = "PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE";   // Balance Sheet files folder (can be same or different)

const SHEET_GL    = "GL";
const SHEET_FINAL = "Final";
const SHEET_QA    = "Missing_GL_Mapping";

// Final schema (includes Group)
const FINAL_HEADERS = ["GL Code", "Description", "Category", "Group", "Year", "Month", "Department", "Amount"];

// QA schema (accumulative tracking)
const QA_HEADERS = [...FINAL_HEADERS, "Missing GL in Reference?", "Status", "Last Seen"];

// Dept sheet pattern for income statements
const DEPT_SHEET_REGEX = /^DEPARTMENT\s+(\d+)\s*[-–—]\s*F/i;

// File name pattern: "mm.yyyy ..."
const FILE_MMYYYY_REGEX = /(\d{2})\.(\d{4})/;


/***************
 * UI: menu
 ***************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("EXAMPLE_COMPANY Warehouse")
    .addItem("Run Update (Income + Balance)", "runWarehouseUpdate")
    .addItem("Initialize/Repair Sheets", "initializeWarehouseSheets")
    .addToUi();
}


/***************
 * MAIN
 ***************/
function runWarehouseUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initializeWarehouseSheets();

  const glMap = loadGlMap_(ss); // GL -> {description, group}

  const finalSheet = ss.getSheetByName(SHEET_FINAL);
  let finalRows = readFinalRows_(finalSheet);

  const qaSheet = ss.getSheetByName(SHEET_QA);

  // ---------
  // 1) Income Statements
  // ---------
  const incomeFolder = DriveApp.getFolderById(INCOME_INPUTS_FOLDER_ID);
  const incomeFiles = getAllMonthlyXlsx_(incomeFolder);
  let incomeExtracted = 0;

  for (const file of incomeFiles) {
    const { monthNum, year } = parseMonthYearFromFilename_(file.getName());
    const monthName = monthNumToName_(monthNum);

    const tempId = convertXlsxToGoogleSheet_(file);
    try {
      const tempSs = SpreadsheetApp.openById(tempId);
      const newRows = parseIncomeStatementWorkbook_(tempSs, glMap, year, monthName);
      incomeExtracted += newRows.length;

      finalRows = appendAndDedupe_(finalRows, newRows);
      writeQaMissingAccumulative_(qaSheet, newRows, glMap);
    } finally {
      safeTrash_(tempId);
    }
  }

  // ---------
  // 2) Balance Sheets
  // ---------
  const balanceFolder = DriveApp.getFolderById(BALANCE_INPUTS_FOLDER_ID);
  const balanceFiles = getAllMonthlyXlsx_(balanceFolder);
  let balanceExtracted = 0;

  for (const file of balanceFiles) {
    const { monthNum, year } = parseMonthYearFromFilename_(file.getName());
    const monthName = monthNumToName_(monthNum);

    const tempId = convertXlsxToGoogleSheet_(file);
    try {
      const tempSs = SpreadsheetApp.openById(tempId);

      // Balance parsing uses:
      // GL col B, Description col C, Amount col E, Department blank
      const newRows = parseBalanceSheetWorkbook_(tempSs, glMap, year, monthName);
      balanceExtracted += newRows.length;

      finalRows = appendAndDedupe_(finalRows, newRows);
      writeQaMissingAccumulative_(qaSheet, newRows, glMap);
    } finally {
      safeTrash_(tempId);
    }
  }

  // Write final output once
  writeFinalRows_(finalSheet, finalRows);

  SpreadsheetApp.getUi().alert(
    `Done!\n\nIncome files: ${incomeFiles.length} | rows extracted: ${incomeExtracted}\n` +
    `Balance files: ${balanceFiles.length} | rows extracted: ${balanceExtracted}\n\n` +
    `Final rows now: ${finalRows.length}\n\n` +
    `Missing GLs accumulate in '${SHEET_QA}'.`
  );
}


/***************
 * SETUP / REPAIR
 ***************/
function initializeWarehouseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_GL)) ss.insertSheet(SHEET_GL);

  let finalSheet = ss.getSheetByName(SHEET_FINAL);
  if (!finalSheet) finalSheet = ss.insertSheet(SHEET_FINAL);
  ensureHeaders_(finalSheet, FINAL_HEADERS);

  let qaSheet = ss.getSheetByName(SHEET_QA);
  if (!qaSheet) qaSheet = ss.insertSheet(SHEET_QA);
  ensureHeaders_(qaSheet, QA_HEADERS);
}


/***************
 * DRIVE HELPERS
 ***************/
function getAllMonthlyXlsx_(folder) {
  const files = [];
  const iter = folder.getFiles();

  while (iter.hasNext()) {
    const f = iter.next();
    const name = (f.getName() || "").toLowerCase();

    if (!name.endsWith(".xlsx")) continue;
    if (!FILE_MMYYYY_REGEX.test(f.getName())) continue;

    files.push(f);
  }

  // Sort by (year, month)
  files.sort((a, b) => {
    const ma = a.getName().match(FILE_MMYYYY_REGEX);
    const mb = b.getName().match(FILE_MMYYYY_REGEX);
    const mA = parseInt(ma[1], 10), yA = parseInt(ma[2], 10);
    const mB = parseInt(mb[1], 10), yB = parseInt(mb[2], 10);
    if (yA !== yB) return yA - yB;
    return mA - mB;
  });

  return files;
}

// Drive API v3 conversion (Advanced Drive Service must be enabled v3)
function convertXlsxToGoogleSheet_(xlsxFile) {
  const blob = xlsxFile.getBlob();
  const resource = { name: "[TEMP] " + xlsxFile.getName(), mimeType: MimeType.GOOGLE_SHEETS };
  const created = Drive.Files.create(resource, blob);
  return created.id;
}

function safeTrash_(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {
    console.log("Warning: could not trash temp file: " + e);
  }
}


/***************
 * PARSING: filename month/year
 ***************/
function parseMonthYearFromFilename_(filename) {
  const m = filename.match(FILE_MMYYYY_REGEX);
  if (!m) throw new Error(`Could not find mm.yyyy in filename: ${filename}`);
  const monthNum = parseInt(m[1], 10);
  const year = parseInt(m[2], 10);
  if (!(monthNum >= 1 && monthNum <= 12)) throw new Error(`Month out of range: ${monthNum}`);
  return { monthNum, year };
}

function monthNumToName_(monthNum) {
  return ["", "January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"][monthNum];
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
 * GL MAP: GL -> {description, group}
 * IMPORTANT: normalize GL codes so "1" becomes "0001"
 ***************/
function loadGlMap_(warehouseSs) {
  const glSheet = warehouseSs.getSheetByName(SHEET_GL);
  const values = glSheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    throw new Error("GL sheet is empty. Add GL Code + Description + Group to the GL tab.");
  }

  const headers = values[0].map(h => (h ?? "").toString().trim().toLowerCase());

  const glIdx = headers.findIndex(h => ["gl", "gl#", "gl code", "glcode", "number", "account", "account number"].includes(h));
  const descIdx = headers.findIndex(h => ["description", "gl description", "account description", "name"].includes(h));
  const groupIdx = headers.findIndex(h => h === "group");

  if (glIdx === -1) throw new Error("Could not find GL code column in GL tab.");
  if (descIdx === -1) throw new Error("Could not find Description column in GL tab.");
  if (groupIdx === -1) throw new Error("Could not find Group column in GL tab.");

  const map = {};
  for (let i = 1; i < values.length; i++) {
    const gl = normalizeGlCode_(values[i][glIdx]); // ✅ normalize here
    if (!gl) continue;

    const desc = (values[i][descIdx] ?? "").toString().trim();
    const grp  = (values[i][groupIdx] ?? "").toString().trim();

    map[gl] = { description: desc, group: grp };
  }
  return map;
}


/***************
 * INCOME STATEMENT PARSING (dept tabs)
 * IMPORTANT: normalize GL codes too (safe even if they never have leading zeros)
 ***************/
function parseIncomeStatementWorkbook_(tempSs, glMap, year, monthName) {
  const out = [];
  const sheets = tempSs.getSheets();

  for (const sh of sheets) {
    const dept = extractDepartment_(sh.getName());
    if (!dept) continue;

    const data = sh.getDataRange().getValues();
    if (!data || data.length < 2) continue;

    // Find header row containing "NUMBER" and "DESCRIPTION"
    let startRow = 0;
    for (let i = 0; i < data.length; i++) {
      const a = (data[i][0] ?? "").toString().trim().toUpperCase();
      const b = (data[i][1] ?? "").toString().trim().toUpperCase();
      if (a === "NUMBER" && b === "DESCRIPTION") {
        startRow = i + 1;
        break;
      }
    }

    const body = data.slice(startRow);
    let currentCategory = null;

    for (const r of body) {
      const numberCellRaw = r[0];
      const amtCell = r[2];

      if ((numberCellRaw ?? "").toString().trim().toUpperCase() === "REVENUES") { currentCategory = "Revenue"; continue; }
      if ((numberCellRaw ?? "").toString().trim().toUpperCase() === "EXPENSES") { currentCategory = "Expenses"; continue; }

      const gl = normalizeGlCode_(numberCellRaw); // ✅ normalize
      if (!gl) continue;

      const amount = parseAmount_(amtCell);
      if (amount === null) continue;

      const info = glMap[gl];
      const desc = info ? info.description : "";
      const group = info ? info.group : "";

      out.push([gl, desc, currentCategory || "", group, year, monthName, dept, amount]);
    }
  }

  return out;
}

function extractDepartment_(sheetName) {
  const m = sheetName.trim().match(DEPT_SHEET_REGEX);
  return m ? m[1] : null;
}


/***************
 * BALANCE SHEET PARSING
 * Preferences:
 * - GL code in column B (index 1)
 * - Description in column C (index 2)
 * - Amount in column E (index 4) ONLY
 * - Department blank
 * - Category switches at TOTAL ASSETS / TOTAL LIABILITIES boundaries
 ***************/
function parseBalanceSheetWorkbook_(tempSs, glMap, year, monthName) {
  const out = [];
  const sheets = tempSs.getSheets();
  for (const sh of sheets) {
    out.push(...parseBalanceSheetSheet_(sh, glMap, year, monthName));
  }
  return out;
}

function parseBalanceSheetSheet_(sh, glMap, year, monthName) {
  const rows = [];
  const data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return rows;

  const glCol = 1;   // B
  const descCol = 2; // C
  const amtCol = 4;  // E
  const dept = "";   // blank department

  let currentCategory = "Assets";

  for (let r = 0; r < data.length; r++) {
    const gl = normalizeGlCode_(data[r][glCol]); // ✅ normalize (handles 0001)
    const descCell = (data[r][descCol] ?? "").toString().trim();
    const descUpper = descCell.toUpperCase();

    // Category boundaries (skip these rows)
    if (descUpper.startsWith("TOTAL ASSETS")) {
      currentCategory = "Liability";
      continue;
    }
    if (descUpper.startsWith("TOTAL LIABILITIES")) {
      currentCategory = "Equity";
      continue;
    }

    // Skip blanks + any TOTAL rows (subtotals)
    if (!descCell) continue;
    if (descUpper.startsWith("TOTAL ")) continue;

    // Require GL
    if (!gl) continue;

    // Amount from column E only
    const amount = parseAmount_(data[r][amtCol]);
    if (amount === null) continue;

    const info = glMap[gl];
    const canonicalDesc = info ? info.description : "";
    const group = info ? info.group : "";

    rows.push([gl, canonicalDesc, currentCategory, group, year, monthName, dept, amount]);
  }

  return rows;
}


/***************
 * AMOUNT PARSER (handles $, commas, parentheses negatives)
 ***************/
function parseAmount_(value) {
  if (value === null || value === "" || typeof value === "undefined") return null;
  if (typeof value === "number") return value;

  let s = value.toString().trim();
  if (!s) return null;

  s = s.replace(/\$/g, "").replace(/,/g, "");

  let neg = false;
  if (/^\(.*\)$/.test(s)) {
    neg = true;
    s = s.replace(/[()]/g, "").trim();
  }

  const num = Number(s);
  if (Number.isNaN(num)) return null;
  return neg ? -num : num;
}

/***************
 * GL NORMALIZER (handles leading zeros)
 ***************/
function normalizeGlCode_(value) {
  if (value === null || value === "" || typeof value === "undefined") return null;

  if (typeof value === "number") {
    const n = Math.trunc(value);
    if (n >= 0 && n <= 9999) return String(n).padStart(4, "0");
    return null;
  }

  const s = value.toString().trim();
  if (/^\d{1,4}$/.test(s)) return s.padStart(4, "0");

  return null;
}


/***************
 * FINAL: read, append+dedupe, write
 ***************/
function readFinalRows_(finalSheet) {
  const data = finalSheet.getDataRange().getValues();
  if (!data || data.length < 2) return [];
  return data.slice(1);
}

function appendAndDedupe_(existingRows, newRows) {
  // Natural key: GL + Year + Month + Department + Category
  const outMap = {};

  const upsert = (row) => {
    const gl    = (row[0] ?? "").toString().trim();
    const desc  = (row[1] ?? "").toString();
    const cat   = (row[2] ?? "").toString().trim();
    const group = (row[3] ?? "").toString().trim();
    const year  = (row[4] ?? "").toString().trim();
    const month = (row[5] ?? "").toString().trim();
    const dept  = (row[6] ?? "").toString().trim();
    const amt   = row[7];

    const monthNum = monthNameToNum_(month) || month;
    const key = [gl, year, monthNum, dept, cat].join("|");

    outMap[key] = [gl, desc, cat, group, Number(year), month, dept, amt];
  };

  existingRows.forEach(upsert);
  newRows.forEach(upsert);

  return Object.values(outMap).sort((a, b) => {
    if (a[4] !== b[4]) return a[4] - b[4]; // Year
    const am = monthNameToNum_(a[5]) || 0;
    const bm = monthNameToNum_(b[5]) || 0;
    if (am !== bm) return am - bm;
    if (a[6] !== b[6]) return a[6].localeCompare(b[6]); // Dept
    if (a[2] !== b[2]) return a[2].localeCompare(b[2]); // Category
    return a[0].localeCompare(b[0]); // GL
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
 * QA: Missing GL mappings (ACCUMULATIVE + auto-resolve)
 ***************/
function writeQaMissingAccumulative_(qaSheet, newRows, glMap) {
  ensureHeaders_(qaSheet, QA_HEADERS);

  const data = qaSheet.getDataRange().getValues();
  const existing = (data.length >= 2) ? data.slice(1) : [];

  // Key: GL|Year|Month|Dept|Category
  const issueMap = {};
  for (const r of existing) {
    const gl = (r[0] ?? "").toString().trim();
    if (!gl) continue;
    const year  = (r[4] ?? "").toString().trim();
    const month = (r[5] ?? "").toString().trim();
    const dept  = (r[6] ?? "").toString().trim();
    const cat   = (r[2] ?? "").toString().trim();

    const key = [gl, year, month, dept, cat].join("|");
    issueMap[key] = r;
  }

  const now = new Date();

  // Add/refresh issues from this run (Description blank => missing mapping)
  for (const row of newRows) {
    const gl    = (row[0] ?? "").toString().trim();
    const desc  = (row[1] ?? "").toString().trim();
    const cat   = (row[2] ?? "").toString().trim();
    const group = (row[3] ?? "").toString().trim();
    const year  = (row[4] ?? "").toString().trim();
    const month = (row[5] ?? "").toString().trim();
    const dept  = (row[6] ?? "").toString().trim();
    const amt   = row[7];

    if (!gl) continue;
    if (desc !== "") continue;

    const key = [gl, year, month, dept, cat].join("|");
    const qaRow = [gl, "", cat, group, Number(year), month, dept, amt, "YES", "Open", now];

    if (issueMap[key]) {
      issueMap[key][7]  = amt;
      issueMap[key][8]  = "YES";
      issueMap[key][9]  = "Open";
      issueMap[key][10] = now;
    } else {
      issueMap[key] = qaRow;
    }
  }

  // Auto-resolve if GL now exists in GL map
  for (const key in issueMap) {
    const r = issueMap[key];
    const gl = (r[0] ?? "").toString().trim();
    if (gl && glMap[gl]) {
      r[1] = glMap[gl].description;
      r[3] = glMap[gl].group;
      r[8] = "";
      r[9] = "Resolved";
    }
  }

  const outRows = Object.values(issueMap);

  // Sort: Open first, then Year/Month/Dept/GL
  outRows.sort((a, b) => {
    const sA = (a[9] ?? "").toString();
    const sB = (b[9] ?? "").toString();
    if (sA !== sB) return sA === "Open" ? -1 : 1;

    if (a[4] !== b[4]) return a[4] - b[4];

    const am = monthNameToNum_(a[5]) || 0;
    const bm = monthNameToNum_(b[5]) || 0;
    if (am !== bm) return am - bm;

    if (a[6] !== b[6]) return (a[6] ?? "").toString().localeCompare((b[6] ?? "").toString());
    return (a[0] ?? "").toString().localeCompare((b[0] ?? "").toString());
  });

  // Clear body only
  const lastRow = qaSheet.getLastRow();
  const lastCol = qaSheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) {
    qaSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }

  if (outRows.length > 0) {
    qaSheet.getRange(2, 1, outRows.length, QA_HEADERS.length).setValues(outRows);
    qaSheet.autoResizeColumns(1, QA_HEADERS.length);
  }
}


/***************
 * UTIL: ensure headers
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
