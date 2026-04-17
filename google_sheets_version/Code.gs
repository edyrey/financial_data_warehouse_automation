/***************
 * CONFIG
 ***************/
const INCOME_INPUTS_FOLDER_ID  = "PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE";
const BALANCE_INPUTS_FOLDER_ID = "PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE";

const SHEET_GL = "GL";
const SHEET_FINAL = "Final";
const SHEET_QA = "Missing_GL_Mapping";

const FINAL_HEADERS = ["GL Code", "Description", "Category", "Group", "Year", "Month", "Department", "Amount"];
const QA_HEADERS = [...FINAL_HEADERS, "Missing GL in Reference?", "Status", "Last Seen"];

const DEPT_SHEET_REGEX = /^DEPARTMENT\s+(\d+)\s*[-–—]\s*F/i;
const FILE_MMYYYY_REGEX = /(\d{2})\.(\d{4})/;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("EXAMPLE_COMPANY Warehouse")
    .addItem("Run Update (Income + Balance)", "runWarehouseUpdate")
    .addItem("Initialize/Repair Sheets", "initializeWarehouseSheets")
    .addToUi();
}

function runWarehouseUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initializeWarehouseSheets();

  const glMap = loadGlMap_(ss);
  const finalSheet = ss.getSheetByName(SHEET_FINAL);
  let finalRows = readFinalRows_(finalSheet);
  const qaSheet = ss.getSheetByName(SHEET_QA);

  const incomeFolder = DriveApp.getFolderById(INCOME_INPUTS_FOLDER_ID);
  const incomeFiles = getAllMonthlyXlsx_(incomeFolder);

  for (const file of incomeFiles) {
    const { monthNum, year } = parseMonthYearFromFilename_(file.getName());
    const monthName = monthNumToName_(monthNum);
    const tempId = convertXlsxToGoogleSheet_(file);

    try {
      const tempSs = SpreadsheetApp.openById(tempId);
      const newRows = parseIncomeStatementWorkbook_(tempSs, glMap, year, monthName);
      finalRows = appendAndDedupe_(finalRows, newRows);
      writeQaMissingAccumulative_(qaSheet, newRows, glMap);
    } finally {
      safeTrash_(tempId);
    }
  }

  const balanceFolder = DriveApp.getFolderById(BALANCE_INPUTS_FOLDER_ID);
  const balanceFiles = getAllMonthlyXlsx_(balanceFolder);

  for (const file of balanceFiles) {
    const { monthNum, year } = parseMonthYearFromFilename_(file.getName());
    const monthName = monthNumToName_(monthNum);
    const tempId = convertXlsxToGoogleSheet_(file);

    try {
      const tempSs = SpreadsheetApp.openById(tempId);
      const newRows = parseBalanceSheetWorkbook_(tempSs, glMap, year, monthName);
      finalRows = appendAndDedupe_(finalRows, newRows);
      writeQaMissingAccumulative_(qaSheet, newRows, glMap);
    } finally {
      safeTrash_(tempId);
    }
  }

  writeFinalRows_(finalSheet, finalRows);
  SpreadsheetApp.getUi().alert(`Done. Final rows now: ${finalRows.length}`);
}

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

function parseMonthYearFromFilename_(filename) {
  const m = filename.match(FILE_MMYYYY_REGEX);
  if (!m) throw new Error(`Could not find mm.yyyy in filename: ${filename}`);
  const monthNum = parseInt(m[1], 10);
  const year = parseInt(m[2], 10);
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
  return names[(monthName || "").toString().trim().toLowerCase()] || null;
}

function loadGlMap_(warehouseSs) {
  const glSheet = warehouseSs.getSheetByName(SHEET_GL);
  const values = glSheet.getDataRange().getValues();
  const headers = values[0].map(h => (h ?? "").toString().trim().toLowerCase());

  const glIdx = headers.findIndex(h => ["gl", "gl#", "gl code", "glcode", "number", "account", "account number"].includes(h));
  const descIdx = headers.findIndex(h => ["description", "gl description", "account description", "name"].includes(h));
  const groupIdx = headers.findIndex(h => h === "group");

  const map = {};
  for (let i = 1; i < values.length; i++) {
    const gl = normalizeGlCode_(values[i][glIdx]);
    if (!gl) continue;
    map[gl] = {
      description: (values[i][descIdx] ?? "").toString().trim(),
      group: (values[i][groupIdx] ?? "").toString().trim()
    };
  }
  return map;
}

function parseIncomeStatementWorkbook_(tempSs, glMap, year, monthName) {
  const out = [];
  for (const sh of tempSs.getSheets()) {
    const dept = extractDepartment_(sh.getName());
    if (!dept) continue;

    const data = sh.getDataRange().getValues();
    let startRow = 0;
    for (let i = 0; i < data.length; i++) {
      const a = (data[i][0] ?? "").toString().trim().toUpperCase();
      const b = (data[i][1] ?? "").toString().trim().toUpperCase();
      if (a === "NUMBER" && b === "DESCRIPTION") {
        startRow = i + 1;
        break;
      }
    }

    let currentCategory = null;
    for (const r of data.slice(startRow)) {
      const numberCellRaw = r[0];
      const amtCell = r[2];

      if ((numberCellRaw ?? "").toString().trim().toUpperCase() === "REVENUES") { currentCategory = "Revenue"; continue; }
      if ((numberCellRaw ?? "").toString().trim().toUpperCase() === "EXPENSES") { currentCategory = "Expenses"; continue; }

      const gl = normalizeGlCode_(numberCellRaw);
      if (!gl) continue;

      const amount = parseAmount_(amtCell);
      if (amount === null) continue;

      const info = glMap[gl];
      out.push([gl, info ? info.description : "", currentCategory || "", info ? info.group : "", year, monthName, dept, amount]);
    }
  }
  return out;
}

function extractDepartment_(sheetName) {
  const m = sheetName.trim().match(DEPT_SHEET_REGEX);
  return m ? m[1] : null;
}

function parseBalanceSheetWorkbook_(tempSs, glMap, year, monthName) {
  const out = [];
  for (const sh of tempSs.getSheets()) {
    const data = sh.getDataRange().getValues();
    let currentCategory = "Assets";

    for (const row of data) {
      const gl = normalizeGlCode_(row[1]);
      const descCell = (row[2] ?? "").toString().trim();
      const descUpper = descCell.toUpperCase();

      if (descUpper.startsWith("TOTAL ASSETS")) { currentCategory = "Liability"; continue; }
      if (descUpper.startsWith("TOTAL LIABILITIES")) { currentCategory = "Equity"; continue; }
      if (!descCell || descUpper.startsWith("TOTAL ")) continue;
      if (!gl) continue;

      const amount = parseAmount_(row[4]);
      if (amount === null) continue;

      const info = glMap[gl];
      out.push([gl, info ? info.description : "", currentCategory, info ? info.group : "", year, monthName, "", amount]);
    }
  }
  return out;
}

function parseAmount_(value) {
  if (value === null || value === "" || typeof value === "undefined") return null;
  if (typeof value === "number") return value;
  let s = value.toString().trim().replace(/\$/g, "").replace(/,/g, "");
  let neg = false;
  if (/^\(.*\)$/.test(s)) {
    neg = true;
    s = s.replace(/[()]/g, "").trim();
  }
  const num = Number(s);
  if (Number.isNaN(num)) return null;
  return neg ? -num : num;
}

function normalizeGlCode_(value) {
  if (value === null || value === "" || typeof value === "undefined") return null;
  if (typeof value === "number") return String(Math.trunc(value)).padStart(4, "0");
  const s = value.toString().trim();
  return /^\d{1,4}$/.test(s) ? s.padStart(4, "0") : null;
}

function readFinalRows_(finalSheet) {
  const data = finalSheet.getDataRange().getValues();
  if (!data || data.length < 2) return [];
  return data.slice(1);
}

function appendAndDedupe_(existingRows, newRows) {
  const outMap = {};
  const upsert = (row) => {
    const month = (row[5] ?? "").toString().trim();
    const key = [(row[0] ?? "").toString().trim(), (row[4] ?? "").toString().trim(), monthNameToNum_(month) || month, (row[6] ?? "").toString().trim(), (row[2] ?? "").toString().trim()].join("|");
    outMap[key] = row;
  };
  existingRows.forEach(upsert);
  newRows.forEach(upsert);
  return Object.values(outMap);
}

function writeFinalRows_(finalSheet, rows) {
  finalSheet.clearContents();
  ensureHeaders_(finalSheet, FINAL_HEADERS);
  if (!rows.length) return;
  finalSheet.getRange(2, 1, rows.length, FINAL_HEADERS.length).setValues(rows);
}

function writeQaMissingAccumulative_(qaSheet, newRows, glMap) {
  ensureHeaders_(qaSheet, QA_HEADERS);
  const data = qaSheet.getDataRange().getValues();
  const existing = (data.length >= 2) ? data.slice(1) : [];
  const issueMap = {};

  for (const r of existing) {
    const key = [(r[0] ?? "").toString().trim(), (r[4] ?? "").toString().trim(), (r[5] ?? "").toString().trim(), (r[6] ?? "").toString().trim(), (r[2] ?? "").toString().trim()].join("|");
    issueMap[key] = r;
  }

  const now = new Date();
  for (const row of newRows) {
    if ((row[1] ?? "").toString().trim() !== "") continue;
    const key = [(row[0] ?? "").toString().trim(), (row[4] ?? "").toString().trim(), (row[5] ?? "").toString().trim(), (row[6] ?? "").toString().trim(), (row[2] ?? "").toString().trim()].join("|");
    issueMap[key] = [row[0], "", row[2], row[3], row[4], row[5], row[6], row[7], "YES", "Open", now];
  }

  for (const key in issueMap) {
    const row = issueMap[key];
    const gl = (row[0] ?? "").toString().trim();
    if (gl && glMap[gl]) {
      row[1] = glMap[gl].description;
      row[3] = glMap[gl].group;
      row[8] = "";
      row[9] = "Resolved";
    }
  }

  const outRows = Object.values(issueMap);
  const lastRow = qaSheet.getLastRow();
  const lastCol = qaSheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) qaSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (outRows.length) qaSheet.getRange(2, 1, outRows.length, QA_HEADERS.length).setValues(outRows);
}

function ensureHeaders_(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}
