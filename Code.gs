// ============================================================
// SATIJA PAPER - SALARY SYSTEM BACKEND
// Google Apps Script (Code.gs)
// ============================================================

// ── Sheet names in THIS spreadsheet (where Code.gs runs) ────
const SHEET_EMPLOYEES = "Employees"; // Name | Mode | MonthlySalary
const SHEET_PAYROLL   = "Payroll";   // Name | Bypass | Status | LastUpdated

// ── External Leaves spreadsheet (Attendance System) ──────────
const LEAVES_SS_ID  = "179U4qy_lPVOV4HtmNHoGfnNy_TgWHhzhxg0Ip-2sR2k";
const LEAVES_TAB    = "Leaves"; // Tab name visible in screenshot

// ─── HTTP ENTRY POINTS ──────────────────────────────────────

function doGet(e) {
  const p      = (e && e.parameter) || {};
  const action = p.action || "";
  const month  = parseInt(p.month) || (new Date().getMonth() + 1);
  const year   = parseInt(p.year)  || new Date().getFullYear();
  let result;

  try {
    if      (action === "getEmployees")  result = getEmployees(month, year);
    else if (action === "getLeaveData")  result = getLeaveData(month, year);
    else { result = { error: "Unknown action: " + action }; }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action || "savePayroll";

    if      (action === "savePayroll")    { savePayroll(payload.data);                     result = { success: true }; }
    else if (action === "addEmployee")    { addEmployee(payload.data);                     result = { success: true }; }
    else if (action === "updateEmployee") { updateEmployee(payload.oldName, payload.data); result = { success: true }; }
    else if (action === "deleteEmployee") { deleteEmployee(payload.name);                  result = { success: true }; }
    else { result = { error: "Unknown action: " + action }; }

  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── EMPLOYEES + LEAVE MERGE ─────────────────────────────────

// Returns combined list: employee info + approved leaves for month/year + payroll state
function getEmployees(month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet(ss, SHEET_EMPLOYEES, ["Name", "Mode", "MonthlySalary"]);
  ensureSheet(ss, SHEET_PAYROLL,   ["Name", "Bypass", "Status", "LastUpdated"]);

  const empRows     = ss.getSheetByName(SHEET_EMPLOYEES).getDataRange().getValues();
  const payrollRows = ss.getSheetByName(SHEET_PAYROLL).getDataRange().getValues();

  if (empRows.length <= 1) return [];

  // Payroll state map: name_lc → {bypass, status}
  const payrollMap = {};
  for (let i = 1; i < payrollRows.length; i++) {
    const r = payrollRows[i];
    if (r[0]) payrollMap[norm(r[0])] = { bypass: parseFloat(r[1]) || 0, status: r[2] || "" };
  }

  // Leave data for selected month from external sheet
  const leaveMap = getLeaveMap(month, year); // name_lc → {total, cl, sl, half, details}

  const employees = [];
  for (let i = 1; i < empRows.length; i++) {
    const r = empRows[i];
    if (!r[0]) continue;

    const name   = String(r[0]).trim();
    const mode   = String(r[1]).trim() || "Cash";
    const salary = parseFloat(r[2]) || 0;
    const key    = norm(name);
    const pr     = payrollMap[key] || {};
    const lv     = leaveMap[key]   || { total: 0, cl: 0, sl: 0, half: 0, details: [] };

    employees.push({
      name,
      mode,
      salary,
      leaves       : lv.total,
      leaveCL      : lv.cl,
      leaveSL      : lv.sl,
      leaveHalf    : lv.half,
      leaveDetails : lv.details,
      bypass       : pr.bypass !== undefined ? pr.bypass : 0,
      status       : pr.status || (mode === "Cash" ? "Pending Manager Review" : "Pending Manager Review")
    });
  }

  return employees;
}

// ─── LEAVE DATA FROM EXTERNAL SHEET ─────────────────────────

// Returns per-employee leave summary for a given month/year
// Only rows where Status column = "Approved" are counted
function getLeaveData(month, year) {
  return Object.values(getLeaveMap(month, year));
}

function getLeaveMap(month, year) {
  let leaveSS, leavesSheet;
  try {
    leaveSS     = SpreadsheetApp.openById(LEAVES_SS_ID);
    leavesSheet = leaveSS.getSheetByName(LEAVES_TAB);
  } catch(e) {
    return {}; // External sheet not accessible
  }
  if (!leavesSheet) return {};

  const rows = leavesSheet.getDataRange().getValues();
  // Row 0 = headers: Timestamp | Name | Type | Start | End | Reason | Status
  // Index:              0           1      2      3      4     5        6

  const map = {}; // name_lc → {name, total, cl, sl, half, details}

  for (let i = 1; i < rows.length; i++) {
    const row    = rows[i];
    const name   = String(row[1] || "").trim();
    const type   = String(row[2] || "").trim(); // CL | SL | 1/2 days
    const start  = parseDate(row[3]);
    const end    = parseDate(row[4]) || start;
    const reason = String(row[5] || "").trim();
    const status = String(row[6] || "").trim().toLowerCase();

    if (!name || !start || status !== "approved") continue;

    const isHalf = type.includes("1/2") || type.toLowerCase().includes("half");
    const key    = norm(name);

    // Count only days that fall in the target month/year
    let days = 0;
    if (isHalf) {
      if (start.getMonth() + 1 === month && start.getFullYear() === year) days = 0.5;
    } else {
      const cur = new Date(start);
      const fin = new Date(end);
      while (cur <= fin) {
        if (cur.getMonth() + 1 === month && cur.getFullYear() === year) days += 1;
        cur.setDate(cur.getDate() + 1);
      }
    }

    if (days === 0) continue;

    if (!map[key]) map[key] = { name, total: 0, cl: 0, sl: 0, half: 0, details: [] };

    const typeLc = type.toUpperCase();
    if      (isHalf)           map[key].half += days;
    else if (typeLc === "CL")  map[key].cl   += days;
    else if (typeLc === "SL")  map[key].sl   += days;
    else                       map[key].cl   += days; // unknown type → treat as CL

    map[key].total += days;
    map[key].details.push({
      type   : isHalf ? "1/2 Day" : typeLc,
      start  : fmtDate(start),
      end    : fmtDate(end),
      days,
      reason
    });
  }

  return map;
}

// ─── PAYROLL SAVE ────────────────────────────────────────────

function savePayroll(employees) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet(ss, SHEET_PAYROLL, ["Name", "Bypass", "Status", "LastUpdated"]);
  const now   = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy HH:mm");

  const existingData = sheet.getDataRange().getValues();
  const existingMap  = {};
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) existingMap[norm(existingData[i][0])] = i + 1;
  }

  employees.forEach(emp => {
    const key    = norm(emp.name);
    const bypass = parseFloat(emp.bypass) || 0;
    const status = emp.status || "";
    if (existingMap[key]) {
      sheet.getRange(existingMap[key], 2, 1, 3).setValues([[bypass, status, now]]);
    } else {
      sheet.appendRow([emp.name, bypass, status, now]);
    }
  });
}

// ─── EMPLOYEE CRUD ────────────────────────────────────────────

function addEmployee(emp) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ensureSheet(ss, SHEET_EMPLOYEES, ["Name", "Mode", "MonthlySalary"]);
  const existing = empSheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (norm(existing[i][0]) === norm(emp.name)) return; // duplicate
  }
  empSheet.appendRow([emp.name, emp.mode, emp.salary]);

  const payrollSheet = ensureSheet(ss, SHEET_PAYROLL, ["Name", "Bypass", "Status", "LastUpdated"]);
  const now = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy HH:mm");
  payrollSheet.appendRow([emp.name, 0, "Pending Manager Review", now]);
}

function updateEmployee(oldName, updatedData) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!empSheet) return;
  const rows = empSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (norm(rows[i][0]) === norm(oldName)) {
      empSheet.getRange(i + 1, 1, 1, 3).setValues([[updatedData.name, updatedData.mode, updatedData.salary]]);
      if (norm(oldName) !== norm(updatedData.name)) {
        const ps = ss.getSheetByName(SHEET_PAYROLL);
        if (ps) {
          const pr = ps.getDataRange().getValues();
          for (let j = 1; j < pr.length; j++) {
            if (norm(pr[j][0]) === norm(oldName)) ps.getRange(j + 1, 1).setValue(updatedData.name);
          }
        }
      }
      break;
    }
  }
}

function deleteEmployee(empName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_EMPLOYEES, SHEET_PAYROLL].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (norm(rows[i][0]) === norm(empName)) sheet.deleteRow(i + 1);
    }
  });
}

// ─── HELPERS ─────────────────────────────────────────────────

function norm(s) { return String(s || "").trim().toLowerCase(); }

function fmtDate(d) {
  return Utilities.formatDate(d, "Asia/Kolkata", "dd-MM-yyyy");
}

// Parses Date objects and common string formats (YYYY-MM-DD, DD-MM-YYYY)
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  const s = String(val).trim();
  if (!s) return null;

  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const d = new Date(s.substring(0, 10));
    return isNaN(d.getTime()) ? null : d;
  }
  // DD-MM-YYYY
  if (/^\d{2}-\d{2}-\d{4}/.test(s)) {
    const parts = s.split("-");
    const d = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
    return isNaN(d.getTime()) ? null : d;
  }
  // Fallback
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1a73e8").setFontColor("white");
  }
  return sheet;
}
