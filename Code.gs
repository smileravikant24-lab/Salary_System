// Satija Paper payroll & expenses web-app backend.
// Deploy this file as a Google Apps Script web app for the target spreadsheet.

const DATA_SS_ID = "1C3_dzbSSQWaldvHh8vSYd-ZaYfdZUfPry3vlU1GEtAQ";
const LEAVES_SS_ID = "179U4qy_lPVOV4HtmNHoGfnNy_TgWHhzhxg0Ip-2sR2k";
const EMPLOYEES_TAB = "Employees";
const PAYROLL_TAB = "Payroll";
const LEAVES_TAB = "Leaves";
const FIXED_EXPENSES_TAB = "FixedExpenses";

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
      return json_(getEmployees_(Number(p.month) || (new Date().getMonth() + 1), Number(p.year) || new Date().getFullYear()));
    }
    if (p.action === "getExpenses") {
      return json_(getExpenses_(Number(p.month) || (new Date().getMonth() + 1), Number(p.year) || new Date().getFullYear()));
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
    if (payload.action === "savePayroll") savePayroll_(payload);
    else if (payload.action === "saveExpenses") saveExpenses_(payload);
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

  const roster = getLeaveRoster_();
  Object.keys(roster).forEach(function (key) {
    if (!employeeMap[key]) {
      employeeMap[key] = { name: roster[key], mode: "Cash", salary: 0 };
    }
  });

  const SEPTEMBER_2026_SALARIES = {
    "ms. khushi painter": 15000,
    "mr. indresh pratap singh": 29500,
    "ms. neha nidhi xaxa": 15000,
    "mr. sonu chauhan": 40500,
    "mr. mukesh shukla": 40000,
    "mr. pawan negi": 40000,
    "ms. pooja": 10000,
    "mr. sandeep dhawan": 40000,
    "mr. rishabh jain": 22000,
    "ms. richa": 8000,
    "mrs. rashmi di": 6500,
    "mr. ravi kant": 11500,
    "mr. pawan bhaiya (driver)": 22000,
    "mrs.sushma shukla": 35000
  };

  return Object.keys(employeeMap).map(function (key) {
    const employee = employeeMap[key];
    const leave = leaves[key] || { total: 0, cl: 0, sl: 0, half: 0, details: [] };
    const saved = payrollMap[key] || {};
    let sal = employee.salary;
    if (month === 9 && year === 2026 && SEPTEMBER_2026_SALARIES[key] !== undefined) {
      sal = SEPTEMBER_2026_SALARIES[key];
    }
    return {
      name: employee.name,
      mode: employee.mode,
      salary: sal,
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

function savePayroll_(payload) {
  const ss = SpreadsheetApp.openById(DATA_SS_ID);
  const employees = payload.data || payload || [];
  const month = Number(payload.month) || (new Date().getMonth() + 1);
  const year = Number(payload.year) || new Date().getFullYear();
  const now = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy HH:mm");
  const MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const mName = MONTH_NAMES[month] || ("Month_" + month);

  // 1. Update Master Employees Sheet (only for current/active months, do not overwrite with past September salaries)
  if (!(month === 9 && year === 2026)) {
    const employeeSheet = ensureSheet_(ss, EMPLOYEES_TAB, ["Name", "Mode", "MonthlySalary"]);
    const empRows = employeeSheet.getDataRange().getValues();
    const empRowByName = {};
    for (let i = 1; i < empRows.length; i++) {
      if (empRows[i][0]) empRowByName[norm_(empRows[i][0])] = i + 1;
    }
    employees.forEach(function (emp) {
      const key = norm_(emp.name);
      const empVals = [emp.name, emp.mode || "Cash", Number(emp.salary) || 0];
      if (empRowByName[key]) {
        employeeSheet.getRange(empRowByName[key], 1, 1, 3).setValues([empVals]);
      } else {
        employeeSheet.appendRow(empVals);
      }
    });
  }

  // 2. Update Master Payroll Sheet (Overrides & Status)
  const payrollSheet = payrollSheet_(ss);
  const pRows = payrollSheet.getDataRange().getValues();
  const pRowByName = {};
  for (let i = 1; i < pRows.length; i++) {
    if (pRows[i][0]) pRowByName[norm_(pRows[i][0])] = i + 1;
  }
  employees.forEach(function (emp) {
    const key = norm_(emp.name);
    const pVals = [
      emp.name,
      emp.leaveOverride === "" || emp.leaveOverride == null ? "" : Number(emp.leaveOverride),
      emp.deductionOverride === "" || emp.deductionOverride == null ? "" : Number(emp.deductionOverride),
      emp.status || "",
      now
    ];
    const r = pRowByName[key];
    if (r) payrollSheet.getRange(r, 1, 1, pVals.length).setValues([pVals]);
    else payrollSheet.appendRow(pVals);
  });

  // 3. Save / Update Month-Specific Sheet (e.g. "September")
  let monthTabName = mName;
  if (!ss.getSheetByName(monthTabName) && ss.getSheetByName("Payroll_" + mName + "_" + year)) {
    monthTabName = "Payroll_" + mName + "_" + year;
  } else if (!ss.getSheetByName(monthTabName) && ss.getSheetByName(mName + " " + year)) {
    monthTabName = mName + " " + year;
  }
  let monthSheet = ss.getSheetByName(monthTabName);
  if (!monthSheet) monthSheet = ss.insertSheet(monthTabName);

  monthSheet.clear();

  const headers = [
    "S.No",
    "Employee Name",
    "Payment Mode",
    "Leave Days",
    "Monthly Salary (₹)",
    "Per Day Rate (₹)",
    "Leave Deduction (₹)",
    "Day-Based Net (₹)",
    "Final Payable (₹)",
    "Approval Status",
    "Last Updated"
  ];

  const monthRows = [headers];
  let totSal = 0, totDed = 0, totPay = 0, totLeaves = 0;

  employees.forEach(function (emp, idx) {
    const sal = Number(emp.salary) || 0;
    const lv = Number(emp.leaveDays != null ? emp.leaveDays : (emp.leaves || 0));
    const perDay = Number(emp.perDay != null ? emp.perDay : (sal ? (sal / 26).toFixed(2) : 0));
    const ded = Number(emp.leaveDeduction != null ? emp.leaveDeduction : (emp.deductionOverride !== "" ? emp.deductionOverride : Math.round(perDay * lv * 100) / 100));
    const net = Number(emp.dayBasedAmount != null ? emp.dayBasedAmount : Math.max(0, sal - ded));
    const fin = Number(emp.finalAmount != null ? emp.finalAmount : net);

    totSal += sal;
    totDed += ded;
    totPay += fin;
    totLeaves += lv;

    monthRows.push([
      idx + 1,
      emp.name,
      emp.mode || "Cash",
      lv,
      sal,
      perDay,
      ded,
      net,
      fin,
      emp.status || (month === 9 && year === 2026 ? "Approved by MD ✓" : "Pending"),
      now
    ]);
  });

  // TOTAL Row
  monthRows.push([
    "TOTAL",
    "TOTAL STAFF: " + employees.length,
    "",
    totLeaves,
    totSal,
    "",
    Math.round(totDed * 100) / 100,
    Math.round(totPay * 100) / 100,
    Math.round(totPay * 100) / 100,
    "Official Record",
    now
  ]);

  monthSheet.getRange(1, 1, monthRows.length, headers.length).setValues(monthRows);

  // Formatting Header
  monthSheet.getRange(1, 1, 1, headers.length)
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  // Formatting Total Row
  monthSheet.getRange(monthRows.length, 1, 1, headers.length)
    .setBackground("#f0f4f8")
    .setFontWeight("bold");

  monthSheet.setFrozenRows(1);
}

function getExpenses_(month, year) {
  const ss = SpreadsheetApp.openById(DATA_SS_ID);
  const sheet = ss.getSheetByName(FIXED_EXPENSES_TAB);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (Number(r[0]) === Number(month) && Number(r[1]) === Number(year)) {
      list.push({
        sno: r[2],
        head: r[3],
        bank: r[4],
        to: r[5],
        particular: r[6],
        amount: Number(r[7]) || 0,
        status: r[8] || ""
      });
    }
  }
  return list;
}

function saveExpenses_(payload) {
  const ss = SpreadsheetApp.openById(DATA_SS_ID);
  const sheet = ensureSheet_(ss, FIXED_EXPENSES_TAB, [
    "Month", "Year", "S.No", "Head", "Bank", "TO", "Particular", "Amount", "Status", "LastUpdated"
  ]);
  const month = Number(payload.month);
  const year = Number(payload.year);
  const data = payload.data || [];
  const now = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy HH:mm");

  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (Number(rows[i][0]) === month && Number(rows[i][1]) === year) {
      sheet.deleteRow(i + 1);
    }
  }

  const newRows = data.map(function (item, idx) {
    return [
      month,
      year,
      idx + 1,
      item.head || "",
      item.bank || "",
      item.to || "TO",
      item.particular || "",
      Number(item.amount) || 0,
      item.status || "",
      now
    ];
  });
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
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
