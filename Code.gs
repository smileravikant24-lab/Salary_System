// ============================================================
// SATIJA PAPER - SALARY SYSTEM BACKEND
// Google Apps Script (Code.gs)
// ============================================================

// ── All data lives in this one spreadsheet ───────────────────
const DATA_SS_ID      = "179U4qy_lPVOV4HtmNHoGfnNy_TgWHhzhxg0Ip-2sR2k";

const SHEET_EMPLOYEES = "Employees"; // Name | Mode | MonthlySalary
const SHEET_PAYROLL   = "Payroll";   // Name | Bypass | Status | LastUpdated
const LEAVES_TAB      = "Leaves";    // Timestamp | Name | Type | Start | End | Reason | Status

// ─── Helper: always use the central spreadsheet ──────────────
function getDataSS() {
  return SpreadsheetApp.openById(DATA_SS_ID);
}

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
// Employee names come from the Leaves sheet (unique names).
// Salary & Mode come from the Employees tab (matched by name).
// If an employee has leaves but no Employees tab entry, they
// still appear with salary=0 until set via Add Employee form.

function getEmployees(month, year) {
  const ss = getDataSS();

  ensureSheet(ss, SHEET_EMPLOYEES, ["Name", "Mode", "MonthlySalary"]);
  ensureSheet(ss, SHEET_PAYROLL,   ["Name", "Bypass", "Status", "LastUpdated"]);

  // Salary/Mode lookup from Employees tab
  const empRows  = ss.getSheetByName(SHEET_EMPLOYEES).getDataRange().getValues();
  const empMap   = {}; // name_lc → {mode, salary}
  for (let i = 1; i < empRows.length; i++) {
    const r = empRows[i];
    if (!r[0]) continue;
    empMap[norm(r[0])] = { name: String(r[0]).trim(), mode: String(r[1]).trim() || "Cash", salary: parseFloat(r[2]) || 0 };
  }

  // Payroll state map
  const payrollRows = ss.getSheetByName(SHEET_PAYROLL).getDataRange().getValues();
  const payrollMap  = {};
  for (let i = 1; i < payrollRows.length; i++) {
    const r = payrollRows[i];
    if (r[0]) payrollMap[norm(r[0])] = { bypass: parseFloat(r[1]) || 0, status: r[2] || "" };
  }

  // Full leave map for selected month (ALL months' unique names for roster)
  const leaveMap     = getLeaveMap(month, year);
  const allNamesMap  = getAllEmployeeNamesFromLeaves(); // unique names ever in Leaves

  // Merge: union of names from Leaves sheet + Employees tab
  const seen     = {};
  const employees = [];

  const addEmployee_ = (key, displayName) => {
    if (seen[key]) return;
    seen[key] = true;
    const emp  = empMap[key]     || { name: displayName, mode: "Cash", salary: 0 };
    const pr   = payrollMap[key] || {};
    const lv   = leaveMap[key]   || { total: 0, cl: 0, sl: 0, half: 0, details: [] };
    employees.push({
      name         : emp.name,
      mode         : emp.mode,
      salary       : emp.salary,
      leaves       : lv.total,
      leaveCL      : lv.cl,
      leaveSL      : lv.sl,
      leaveHalf    : lv.half,
      leaveDetails : lv.details,
      bypass       : pr.bypass !== undefined ? pr.bypass : 0,
      status       : pr.status || "Pending Manager Review"
    });
  };

  // First: all names that appear in Leaves sheet
  Object.keys(allNamesMap).forEach(key => addEmployee_(key, allNamesMap[key]));

  // Then: any extras in Employees tab not in Leaves
  Object.keys(empMap).forEach(key => addEmployee_(key, empMap[key].name));

  return employees;
}

// Returns all unique employee names ever seen in Leaves sheet
function getAllEmployeeNamesFromLeaves() {
  const ss    = getDataSS();
  const sheet = ss.getSheetByName(LEAVES_TAB);
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][1] || "").trim();
    if (name) map[norm(name)] = name;
  }
  return map;
}

// ─── LEAVE DATA FROM CENTRAL SHEET ──────────────────────────

function getLeaveData(month, year) {
  return Object.values(getLeaveMap(month, year));
}

function getLeaveMap(month, year) {
  let ss, leavesSheet;
  try {
    ss          = getDataSS();
    leavesSheet = ss.getSheetByName(LEAVES_TAB);
  } catch(e) {
    return {};
  }
  if (!leavesSheet) return {};

  const rows = leavesSheet.getDataRange().getValues();
  // Row 0 = headers: Timestamp | Name | Type | Start | End | Reason | Status

  const map = {};

  for (let i = 1; i < rows.length; i++) {
    const row    = rows[i];
    const name   = String(row[1] || "").trim();
    const type   = String(row[2] || "").trim();
    const start  = parseDate(row[3]);
    const end    = parseDate(row[4]) || start;
    const reason = String(row[5] || "").trim();
    const status = String(row[6] || "").trim().toLowerCase();

    if (!name || !start || status !== "approved") continue;

    const isHalf = type.includes("1/2") || type.toLowerCase().includes("half");
    const key    = norm(name);

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
    else                       map[key].cl   += days;

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
  const ss    = getDataSS();
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
  const ss       = getDataSS();
  const empSheet = ensureSheet(ss, SHEET_EMPLOYEES, ["Name", "Mode", "MonthlySalary"]);
  const existing = empSheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (norm(existing[i][0]) === norm(emp.name)) return;
  }
  empSheet.appendRow([emp.name, emp.mode, emp.salary]);

  const payrollSheet = ensureSheet(ss, SHEET_PAYROLL, ["Name", "Bypass", "Status", "LastUpdated"]);
  const now = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy HH:mm");
  payrollSheet.appendRow([emp.name, 0, "Pending Manager Review", now]);
}

function updateEmployee(oldName, updatedData) {
  const ss       = getDataSS();
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
  const ss = getDataSS();
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

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  const s = String(val).trim();
  if (!s) return null;

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const d = new Date(s.substring(0, 10));
    return isNaN(d.getTime()) ? null : d;
  }
  if (/^\d{2}-\d{2}-\d{4}/.test(s)) {
    const parts = s.split("-");
    const d = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
    return isNaN(d.getTime()) ? null : d;
  }
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
