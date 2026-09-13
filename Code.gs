// Satija Paper payroll web-app backend.
// Deploy this file as a Google Apps Script web app for the target spreadsheet.

const DATA_SS_ID = "1C3_dzbSSQWaldvHh8vSYd-ZaYfdZUfPry3vlU1GEtAQ";
const LEAVES_SS_ID = "179U4qy_lPVOV4HtmNHoGfnNy_TgWHhzhxg0Ip-2sR2k";
const EMPLOYEES_TAB = "Employees";
const PAYROLL_TAB = "Payroll";
const LEAVES_TAB = "Leaves";
const USERS = {
  "mis@satijapaper.com": { role: "Admin", pass: "admin123" },
  "mukesh.shukla@satijapaper.com": { role: "Manager", pass: "mukesh123" },
  "pranavsatija@satijapaper.com": { role: "MD", pass: "pranav123" },
  "satijapaper@gmail.com": { role: "Accountant", pass: "acc123" }
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === "getEmployees") {
      return json_(getEmployees_(Number(p.month) || new Date().getMonth() + 1, Number(p.year) || new Date().getFullYear()));
    }
    return json_({ error: "Unknown action: " + (p.action || "") });
  } catch (err) {
    return json_({ error: err.message });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === "login") return login_(payload);
    if (payload.action === "savePayroll") savePayroll_(payload.data || []);
    else if (payload.action === "addEmployee") addEmployee_(payload.data);
    else if (payload.action === "updateEmployee") updateEmployee_(payload.oldName, payload.data);
    else if (payload.action === "deleteEmployee") deleteEmployee_(payload.name);
    else throw new Error("Unknown action: " + (payload.action || ""));
    return json_({ success: true });
  } catch (err) {
    return json_({ error: err.message });
  }
}

function login_(payload) {
  const email = String(payload.email || "").trim().toLowerCase();
  const user = USERS[email];
  if (!user || user.pass !== String(payload.password || "")) {
    return json_({ success: false, error: "Wrong password! Please try again." });
  }
  return json_({ success: true, email: email, role: user.role });
}

function getEmployees_(month, year) {
  const ss = SpreadsheetApp.openById(DATA_SS_ID);
  const employees = ensureSheet_(ss, EMPLOYEES_TAB, ["Name", "Mode", "MonthlySalary"]);
  const payroll = payrollSheet_(ss);
  const employeeRows = employees.getDataRange().getValues();
  const payrollRows = payroll.getDataRange().getValues();
  const payrollMap = {};
  for (let i = 1; i < payrollRows.length; i++) {
    if (payrollRows[i][0]) {
      payrollMap[norm_(payrollRows[i][0])] = {
        leaveOverride: payrollRows[i][1],
        deductionOverride: payrollRows[i][2],
        status: payrollRows[i][3] || ""
      };
    }
  }
  const leaves = getLeaveMap_(month, year);
  const employeeMap = {};
  employeeRows.slice(1).forEach(function (r) {
    if (!r[0]) return;
    const name = String(r[0]).trim();
    employeeMap[norm_(name)] = {
      name: name,
      mode: String(r[1] || "Cash"),
      salary: Number(r[2]) || 0
    };
  });

  // The attendance spreadsheet is also the roster source. Include names
  // that have leave records even when salary details are not entered yet.
  const roster = getLeaveRoster_();
  Object.keys(roster).forEach(function (key) {
    if (!employeeMap[key]) {
      employeeMap[key] = { name: roster[key], mode: "Cash", salary: 0 };
    }
  });

  return Object.keys(employeeMap).map(function (key) {
    const employee = employeeMap[key];
    const leave = leaves[key] || { total: 0, cl: 0, sl: 0, half: 0, details: [] };
    const saved = payrollMap[key] || {};
    return {
      name: employee.name,
      mode: employee.mode,
      salary: employee.salary,
      leaves: leave.total,
      leaveCL: leave.cl,
      leaveSL: leave.sl,
      leaveHalf: leave.half,
      leaveDetails: leave.details,
      leaveOverride: saved.leaveOverride > 0 ? Number(saved.leaveOverride) : "",
      deductionOverride: saved.deductionOverride > 0 ? Number(saved.deductionOverride) : "",
      status: saved.status || "Pending Manager Review"
    };
  });
}

function getLeaveRoster_() {
  const sheet = SpreadsheetApp.openById(LEAVES_SS_ID).getSheetByName(LEAVES_TAB);
  if (!sheet) return {};
  const roster = {};
  sheet.getDataRange().getValues().slice(1).forEach(function (row) {
    const name = String(row[1] || "").trim();
    if (name) roster[norm_(name)] = name;
  });
  return roster;
}

function savePayroll_(employees) {
  const ss = SpreadsheetApp.openById(DATA_SS_ID);
  const employeeSheet = ensureSheet_(ss, EMPLOYEES_TAB, ["Name", "Mode", "MonthlySalary"]);
  const sheet = payrollSheet_(ss);
  const employeeRows = employeeSheet.getDataRange().getValues();
  const rows = sheet.getDataRange().getValues();
  const employeeRowByName = {};
  const rowByName = {};
  for (let i = 1; i < employeeRows.length; i++) {
    if (employeeRows[i][0]) employeeRowByName[norm_(employeeRows[i][0])] = i + 1;
  }
  for (let i = 1; i < rows.length; i++) if (rows[i][0]) rowByName[norm_(rows[i][0])] = i + 1;
  const now = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy HH:mm");
  employees.forEach(function (employee) {
    const employeeKey = norm_(employee.name);
    const employeeValues = [employee.name, employee.mode || "Cash", Number(employee.salary) || 0];
    if (employeeRowByName[employeeKey]) {
      employeeSheet.getRange(employeeRowByName[employeeKey], 1, 1, 3).setValues([employeeValues]);
    } else {
      employeeSheet.appendRow(employeeValues);
    }
    const values = [
      employee.name,
      employee.leaveOverride === "" || employee.leaveOverride == null ? "" : Number(employee.leaveOverride),
      employee.deductionOverride === "" || employee.deductionOverride == null ? "" : Number(employee.deductionOverride),
      employee.status || "",
      now
    ];
    const row = rowByName[employeeKey];
    if (row) sheet.getRange(row, 1, 1, values.length).setValues([values]);
    else sheet.appendRow(values);
  });
}

function addEmployee_(employee) {
  const sheet = ensureSheet_(SpreadsheetApp.openById(DATA_SS_ID), EMPLOYEES_TAB,
    ["Name", "Mode", "MonthlySalary"]);
  const name = String(employee && employee.name || "").trim();
  if (!name) throw new Error("Employee name is required.");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (norm_(rows[i][0]) === norm_(name)) {
      sheet.getRange(i + 1, 1, 1, 3).setValues([[name, employee.mode || "Cash", Number(employee.salary) || 0]]);
      return;
    }
  }
  sheet.appendRow([name, employee.mode || "Cash", Number(employee.salary) || 0]);
}

function updateEmployee_(oldName, employee) {
  const sheet = ensureSheet_(SpreadsheetApp.openById(DATA_SS_ID), EMPLOYEES_TAB,
    ["Name", "Mode", "MonthlySalary"]);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (norm_(values[i][0]) === norm_(oldName)) {
      sheet.getRange(i + 1, 1, 1, 3).setValues([[employee.name, employee.mode || "Cash", Number(employee.salary) || 0]]);
      return;
    }
  }
  throw new Error("Employee not found: " + oldName);
}

function deleteEmployee_(name) {
  const ss = SpreadsheetApp.openById(DATA_SS_ID);
  [EMPLOYEES_TAB, PAYROLL_TAB].forEach(function (tab) {
    const sheet = ss.getSheetByName(tab);
    if (!sheet) return;
    const values = sheet.getDataRange().getValues();
    for (let i = values.length - 1; i > 0; i--) {
      if (norm_(values[i][0]) === norm_(name)) sheet.deleteRow(i + 1);
    }
  });
}

function getLeaveMap_(month, year) {
  const sheet = SpreadsheetApp.openById(LEAVES_SS_ID).getSheetByName(LEAVES_TAB);
  if (!sheet) return {};
  const map = {};
  sheet.getDataRange().getValues().slice(1).forEach(function (row) {
    const name = String(row[1] || "").trim();
    const type = String(row[2] || "").trim();
    const start = date_(row[3]);
    const end = date_(row[4]) || start;
    if (!name || !start || String(row[6] || "").toLowerCase() !== "approved") return;
    const half = type.indexOf("1/2") >= 0 || type.toLowerCase().indexOf("half") >= 0;
    let days = 0;
    if (half) {
      if (start.getMonth() + 1 === month && start.getFullYear() === year) days = 0.5;
    } else {
      for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
        if (d.getMonth() + 1 === month && d.getFullYear() === year) days++;
      }
    }
    if (!days) return;
    const key = norm_(name);
    if (!map[key]) map[key] = { total: 0, cl: 0, sl: 0, half: 0, details: [] };
    if (half) map[key].half += days;
    else if (type.toUpperCase() === "SL") map[key].sl += days;
    else map[key].cl += days;
    map[key].total += days;
    map[key].details.push({ type: half ? "1/2 Day" : type.toUpperCase(), start: formatDate_(start), end: formatDate_(end), days: days, reason: String(row[5] || "") });
  });
  return map;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}

function payrollSheet_(ss) {
  const sheet = ensureSheet_(ss, PAYROLL_TAB, ["Name", "Bypass", "Days", "Status", "LastUpdated"]);
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  if (headers[1] === "Bypass" || headers[2] === "Days") {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const oldValues = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
      const migrated = oldValues.map(function (row) { return ["", row[0]]; });
      sheet.getRange(2, 2, migrated.length, 2).setValues(migrated);
    }
  }
  if (headers[2] === "AmountOverride") {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 3, lastRow - 1, 1).clearContent();
  }
  if (headers[1] !== "LeaveOverride") sheet.getRange(1, 2).setValue("LeaveOverride");
  if (headers[2] !== "DeductionOverride") sheet.getRange(1, 3).setValue("DeductionOverride");
  return sheet;
}

function norm_(value) { return String(value || "").trim().toLowerCase(); }
function date_(value) { return value instanceof Date ? new Date(value) : (value ? new Date(value) : null); }
function formatDate_(value) { return Utilities.formatDate(value, "Asia/Kolkata", "dd-MM-yyyy"); }
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
