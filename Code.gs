// ============================================================
// SATIJA PAPER - SALARY SYSTEM BACKEND
// Google Apps Script (Code.gs)
// ============================================================

const SHEET_EMPLOYEES  = "Employees";   // Name | Mode | MonthlySalary
const SHEET_ATTENDANCE = "Attendance";  // Name | Day1 | Day2 | ... (Admin fills this)
const SHEET_PAYROLL    = "Payroll";     // Name | Bypass | Status | LastUpdated

// ─── HTTP ENTRY POINTS ──────────────────────────────────────

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "";
  let result;
  try {
    if (action === "getEmployees") {
      result = getEmployees();
    } else {
      result = { error: "Unknown action: " + action };
    }
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

// ─── MAIN DATA FUNCTIONS ─────────────────────────────────────

// Returns combined employee list: basic info + leaves from Attendance + bypass/status from Payroll
function getEmployees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet(ss, SHEET_EMPLOYEES,  ["Name", "Mode", "MonthlySalary"]);
  ensureSheet(ss, SHEET_PAYROLL,    ["Name", "Bypass", "Status", "LastUpdated"]);

  const empSheet     = ss.getSheetByName(SHEET_EMPLOYEES);
  const payrollSheet = ss.getSheetByName(SHEET_PAYROLL);

  const empRows     = empSheet.getDataRange().getValues();
  const payrollRows = payrollSheet.getDataRange().getValues();

  if (empRows.length <= 1) return [];

  // Build payroll map: lowercase name → {bypass, status}
  const payrollMap = {};
  for (let i = 1; i < payrollRows.length; i++) {
    const r = payrollRows[i];
    if (r[0]) payrollMap[String(r[0]).trim().toLowerCase()] = {
      bypass : parseFloat(r[1]) || 0,
      status : r[2] || ""
    };
  }

  // Calculate leaves per employee from Attendance sheet
  const leavesMap = calculateAllLeaves(ss);

  const employees = [];
  for (let i = 1; i < empRows.length; i++) {
    const r = empRows[i];
    if (!r[0]) continue;

    const name    = String(r[0]).trim();
    const mode    = String(r[1]).trim() || "Cash";
    const salary  = parseFloat(r[2]) || 0;
    const key     = name.toLowerCase();
    const pr      = payrollMap[key] || {};
    const bypass  = pr.bypass !== undefined ? pr.bypass : 0;
    const status  = pr.status || (mode === "Cash" ? "Pending Cash Approval (MD)" : "Pending Bank Verification");
    const leaves  = leavesMap[key] || 0;

    employees.push({ name, mode, salary, leaves, bypass, status });
  }

  return employees;
}

// Read Attendance sheet and return {name_lowercase: totalLeaveDays}
function calculateAllLeaves(ss) {
  const sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) return {};

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return {};

  const map = {};

  // Row 0 = headers (Name, Day1, Day2 …). Rows 1+ = employee data.
  for (let i = 1; i < rows.length; i++) {
    const row  = rows[i];
    if (!row[0]) continue;

    const name = String(row[0]).trim().toLowerCase();
    let leaves = 0;

    for (let j = 1; j < row.length; j++) {
      const val = String(row[j] || "").trim().toLowerCase();
      if (!val || val === "p" || val === "present" || val === "h" || val === "holiday") continue;

      if (val === "cl" || val === "sl" || val === "el" || val.includes("leave") || val === "a" || val === "absent") {
        leaves += 1;
      } else if (val === "1/2" || val === "half" || val === "p+1/2" || val === "1/2+p" || val.includes("half")) {
        leaves += 0.5;
      }
    }

    map[name] = leaves;
  }
  return map;
}

// Save bypass + status changes to Payroll sheet (called on any status/bypass change)
function savePayroll(employees) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet(ss, SHEET_PAYROLL, ["Name", "Bypass", "Status", "LastUpdated"]);
  const now   = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy HH:mm");

  // Build existing map to preserve rows for employees not in payload
  const existingData  = sheet.getDataRange().getValues();
  const existingMap   = {};
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) existingMap[String(existingData[i][0]).trim().toLowerCase()] = i + 1; // 1-indexed row number
  }

  employees.forEach(emp => {
    const key    = emp.name.toLowerCase();
    const bypass = emp.bypass !== undefined ? parseFloat(emp.bypass) || 0 : 0;
    const status = emp.status || "";

    if (existingMap[key]) {
      // Update existing row
      const row = existingMap[key];
      sheet.getRange(row, 2, 1, 3).setValues([[bypass, status, now]]);
    } else {
      // Append new row
      sheet.appendRow([emp.name, bypass, status, now]);
    }
  });
}

// Add employee to Employees sheet + initialize Payroll row
function addEmployee(emp) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ensureSheet(ss, SHEET_EMPLOYEES, ["Name", "Mode", "MonthlySalary"]);

  // Check duplicate
  const existing = empSheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (String(existing[i][0]).trim().toLowerCase() === emp.name.toLowerCase()) return; // already exists
  }

  empSheet.appendRow([emp.name, emp.mode, emp.salary]);

  const payrollSheet = ensureSheet(ss, SHEET_PAYROLL, ["Name", "Bypass", "Status", "LastUpdated"]);
  const now = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy HH:mm");
  const defaultStatus = emp.mode === "Cash" ? "Pending Cash Approval (MD)" : "Pending Bank Verification";
  payrollSheet.appendRow([emp.name, 0, defaultStatus, now]);
}

// Update employee name/mode/salary in Employees sheet
function updateEmployee(oldName, updatedData) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!empSheet) return;

  const rows = empSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === oldName.toLowerCase()) {
      empSheet.getRange(i + 1, 1, 1, 3).setValues([[updatedData.name, updatedData.mode, updatedData.salary]]);

      // Also rename in Payroll sheet if name changed
      if (oldName.toLowerCase() !== updatedData.name.toLowerCase()) {
        const paySheet = ss.getSheetByName(SHEET_PAYROLL);
        if (paySheet) {
          const payRows = paySheet.getDataRange().getValues();
          for (let j = 1; j < payRows.length; j++) {
            if (String(payRows[j][0]).trim().toLowerCase() === oldName.toLowerCase()) {
              paySheet.getRange(j + 1, 1).setValue(updatedData.name);
            }
          }
        }
      }
      break;
    }
  }
}

// Delete employee from Employees and Payroll sheets
function deleteEmployee(empName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_EMPLOYEES, SHEET_PAYROLL].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]).trim().toLowerCase() === empName.toLowerCase()) {
        sheet.deleteRow(i + 1);
      }
    }
  });
}

// ─── HELPER ──────────────────────────────────────────────────

function ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1a73e8").setFontColor("white");
  }
  return sheet;
}
