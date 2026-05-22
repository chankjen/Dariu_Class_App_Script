// ============================================================
//  ATTENDANCE AUTOMATION — Google Apps Script
//  Based on: Lesson.xlsx (Student Roster sheet)
//  Sheets required: "Student Roster", "Attendance Records"
// ============================================================

// ---------- CONFIGURATION ----------
const CONFIG = {
  ROSTER_SHEET:     "Student Roster",
  ATTENDANCE_SHEET: "Attendance Records",
  // Column indices (1-based) in Attendance Records
  COL_DATE:         1,
  COL_SESSION:      2,
  COL_STUDENT_ID:   3,
  COL_NAME:         4,
  COL_STATUS:       5,
  COL_TIME_IN:      6,
  COL_NOTES:        7,
  COL_REVIEW:       8,
  COL_REVIEWED_BY:  9,
  COL_TIMESTAMP:    10,
  // Valid attendance statuses
  STATUSES:         ["Present", "Late", "Absent", "Excused"],
};

// ============================================================
// 1. SETUP — Run once to create sheets and seed the roster
// ============================================================
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _ensureRosterSheet(ss);
  _ensureAttendanceSheet(ss);

  SpreadsheetApp.getUi().alert("✅ Setup complete! Both sheets are ready.");
}

function _ensureRosterSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.ROSTER_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.ROSTER_SHEET);

  if (sheet.getLastRow() === 0) {
    const headers = ["Student ID", "Name", "Email", "Specialty", "Enrollment Date", "Active"];
    sheet.appendRow(headers);
    _styleHeader(sheet, 1);

    // Seed with the 8 learners from Lesson.xlsx
    const roster = [
      ["FD001", "Maya Johnson",    "maya.j@university.edu",   "Sustainable Fashion",  "2026-01-01", "Yes"],
      ["FD002", "James Liu",       "james.l@university.edu",  "Textile Innovation",   "2026-01-01", "Yes"],
      ["FD003", "Sofia Rodriguez", "sofia.r@university.edu",  "Pattern Making",       "2026-01-01", "Yes"],
      ["FD004", "Aisha Okonkwo",   "aisha.o@university.edu",  "African Print Design", "2026-01-01", "Yes"],
      ["FD005", "Lucas Martin",    "lucas.m@university.edu",  "Digital Fashion",      "2026-01-01", "Yes"],
      ["FD006", "Priya Sharma",    "priya.s@university.edu",  "Accessory Design",     "2026-01-01", "Yes"],
      ["FD007", "Ethan Kim",       "ethan.k@university.edu",  "Fashion Tech",         "2026-01-01", "Yes"],
      ["FD008", "Zara Ahmed",      "zara.a@university.edu",   "Couture Design",       "2026-01-01", "Yes"],
    ];
    roster.forEach(row => sheet.appendRow(row));
    sheet.autoResizeColumns(1, headers.length);
  }
}

function _ensureAttendanceSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.ATTENDANCE_SHEET);

  if (sheet.getLastRow() === 0) {
    const headers = [
      "Date", "Class Session", "Student ID", "Student Name",
      "Status", "Time In", "Notes", "Review Status", "Reviewed By", "Review Timestamp"
    ];
    sheet.appendRow(headers);
    _styleHeader(sheet, 1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

// ============================================================
// 2. MARK ATTENDANCE — Pre-populates today's rows as "Pending"
//    then lets the instructor update each status.
// ============================================================
function markAttendanceForToday() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const roster    = _getActiveStudents(ss);
  const attSheet  = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  const today     = new Date();
  const session   = _getSessionLabel(today);

  if (roster.length === 0) {
    SpreadsheetApp.getUi().alert("No active students found in the roster.");
    return;
  }

  // Check for duplicate entries for today
  const existing = _getExistingEntriesForDate(attSheet, today);
  if (existing.length > 0) {
    const ui  = SpreadsheetApp.getUi();
    const res = ui.alert(
      "Attendance already recorded for today.",
      "Do you want to add new rows anyway?",
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return;
  }

  const timeIn = Utilities.formatDate(today, Session.getScriptTimeZone(), "HH:mm");

  roster.forEach(student => {
    attSheet.appendRow([
      Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      session,
      student.id,
      student.name,
      "Present",   // Default — instructor updates as needed
      timeIn,
      "",          // Notes
      "Pending",   // Review Status
      "",          // Reviewed By
      "",          // Review Timestamp
    ]);
  });

  // Add data validation to the Status column for new rows
  _applyStatusValidation(attSheet);

  SpreadsheetApp.getUi().alert(
    `✅ Attendance rows created for ${roster.length} students (${session} — ${Utilities.formatDate(today, Session.getScriptTimeZone(), "dd MMM yyyy")}).\n\nDefault status is "Present". Please update any absences or late arrivals.`
  );
}

// ============================================================
// 3. MARK INDIVIDUAL STUDENT
//    Use from the custom menu to mark a single student quickly.
// ============================================================
function markIndividualStudent() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const idRes = ui.prompt("Mark Individual Student", "Enter Student ID (e.g. FD001):", ui.ButtonSet.OK_CANCEL);
  if (idRes.getSelectedButton() !== ui.Button.OK) return;
  const studentId = idRes.getResponseText().trim().toUpperCase();

  const student = _getStudentById(ss, studentId);
  if (!student) {
    ui.alert(`Student "${studentId}" not found in the roster.`);
    return;
  }

  const statusRes = ui.prompt(
    `Mark ${student.name}`,
    `Enter status (${CONFIG.STATUSES.join(" / ")}):`,
    ui.ButtonSet.OK_CANCEL
  );
  if (statusRes.getSelectedButton() !== ui.Button.OK) return;

  const status = _capitalise(statusRes.getResponseText().trim());
  if (!CONFIG.STATUSES.includes(status)) {
    ui.alert(`Invalid status. Please use one of: ${CONFIG.STATUSES.join(", ")}`);
    return;
  }

  const noteRes = ui.prompt("Notes (optional)", "Enter any notes or leave blank:", ui.ButtonSet.OK_CANCEL);
  const notes   = noteRes.getSelectedButton() === ui.Button.OK ? noteRes.getResponseText().trim() : "";

  const today    = new Date();
  const attSheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  const session  = _getSessionLabel(today);
  const timeIn   = Utilities.formatDate(today, Session.getScriptTimeZone(), "HH:mm");

  attSheet.appendRow([
    Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd"),
    session,
    student.id,
    student.name,
    status,
    timeIn,
    notes,
    "Pending",
    "",
    "",
  ]);

  _applyStatusValidation(attSheet);
  ui.alert(`✅ Marked ${student.name} as "${status}" for ${session}.`);
}

// ============================================================
// 4. GENERATE SUMMARY REPORT
//    Creates (or overwrites) a "Attendance Summary" sheet.
// ============================================================
function generateAttendanceSummary() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet  = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  const roster    = _getActiveStudents(ss);

  if (!attSheet || attSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("No attendance records found.");
    return;
  }

  const data = attSheet.getDataRange().getValues();
  // Build a map: studentId -> { Present, Late, Absent, Excused }
  const summary = {};
  roster.forEach(s => {
    summary[s.id] = { name: s.name, Present: 0, Late: 0, Absent: 0, Excused: 0, Total: 0 };
  });

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const id     = row[CONFIG.COL_STUDENT_ID - 1];
    const status = row[CONFIG.COL_STATUS - 1];
    if (summary[id] && CONFIG.STATUSES.includes(status)) {
      summary[id][status]++;
      summary[id].Total++;
    }
  }

  // Write to summary sheet
  let sumSheet = ss.getSheetByName("Attendance Summary");
  if (sumSheet) ss.deleteSheet(sumSheet);
  sumSheet = ss.insertSheet("Attendance Summary");

  const headers = ["Student ID", "Name", "Present", "Late", "Absent", "Excused", "Total Sessions", "Attendance %"];
  sumSheet.appendRow(headers);
  _styleHeader(sumSheet, 1);

  Object.entries(summary).forEach(([id, s]) => {
    const pct = s.Total > 0 ? ((s.Present + s.Late) / s.Total * 100).toFixed(1) + "%" : "N/A";
    sumSheet.appendRow([id, s.name, s.Present, s.Late, s.Absent, s.Excused, s.Total, pct]);
  });

  // Conditional formatting: highlight low attendance (<75%)
  const lastRow = sumSheet.getLastRow();
  if (lastRow > 1) {
    const pctRange = sumSheet.getRange(2, 8, lastRow - 1, 1);
    const rules    = sumSheet.getConditionalFormatRules();
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("%")
        .setBackground("#FFD2D2")
        .setRanges([pctRange])
        .build()
    );
    sumSheet.setConditionalFormatRules(rules);
  }

  sumSheet.autoResizeColumns(1, headers.length);
  SpreadsheetApp.getUi().alert("✅ Attendance Summary sheet generated!");
}

// ============================================================
// 5. SEND ABSENCE ALERTS (Gmail)
//    Emails students marked Absent today.
// ============================================================
function sendAbsenceAlerts() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if (!attSheet) return;

  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const data     = attSheet.getDataRange().getValues();
  const roster   = _buildRosterMap(ss);
  let   count    = 0;

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const date   = row[CONFIG.COL_DATE - 1];
    const status = row[CONFIG.COL_STATUS - 1];
    const id     = row[CONFIG.COL_STUDENT_ID - 1];

    if (String(date) === today && status === "Absent" && roster[id]) {
      const student = roster[id];
      GmailApp.sendEmail(
        student.email,
        "Attendance Notice — You Were Marked Absent Today",
        `Hi ${student.name},\n\nOur records show you were marked absent for today's class session.\n\nIf this is an error, or if you have submitted an excuse, please contact your instructor.\n\nBest regards,\nCourse Administration`
      );
      count++;
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Sent ${count} absence alert email(s).`);
}

// ============================================================
// 6. AUTO-TRIGGER — Run daily at class time via installable trigger
// ============================================================
function createDailyTrigger() {
  // Delete any existing triggers for markAttendanceForToday
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "markAttendanceForToday")
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Create a new daily trigger at 10:00 AM
  ScriptApp.newTrigger("markAttendanceForToday")
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();

  SpreadsheetApp.getUi().alert("✅ Daily trigger set: attendance rows will auto-generate at 10:00 AM every day.");
}

// ============================================================
// CUSTOM MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 Attendance")
    .addItem("▶ Setup Spreadsheet",         "setupSpreadsheet")
    .addSeparator()
    .addItem("✅ Mark Attendance — Today",   "markAttendanceForToday")
    .addItem("👤 Mark Individual Student",   "markIndividualStudent")
    .addSeparator()
    .addItem("📊 Generate Summary Report",  "generateAttendanceSummary")
    .addItem("📧 Send Absence Alerts",       "sendAbsenceAlerts")
    .addSeparator()
    .addItem("⏰ Set Daily Auto-Trigger",    "createDailyTrigger")
    .addToUi();
}

// ============================================================
// HELPERS
// ============================================================

function _getActiveStudents(ss) {
  const sheet = ss.getSheetByName(CONFIG.ROSTER_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][5]).toLowerCase() === "yes") {
      results.push({ id: data[i][0], name: data[i][1], email: data[i][2] });
    }
  }
  return results;
}

function _getStudentById(ss, id) {
  return _getActiveStudents(ss).find(s => s.id === id) || null;
}

function _buildRosterMap(ss) {
  const map = {};
  _getActiveStudents(ss).forEach(s => { map[s.id] = s; });
  return map;
}

function _getExistingEntriesForDate(sheet, date) {
  if (sheet.getLastRow() < 2) return [];
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const data    = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => String(r[CONFIG.COL_DATE - 1]) === dateStr);
}

function _getSessionLabel(date) {
  // Returns "Week N" based on how many distinct dates are already logged
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return "Week 1";
  const data    = sheet.getDataRange().getValues().slice(1);
  const dates   = [...new Set(data.map(r => String(r[CONFIG.COL_DATE - 1])))];
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (!dates.includes(dateStr)) dates.push(dateStr);
  return "Week " + dates.indexOf(dateStr);
}

function _applyStatusValidation(sheet) {
  if (sheet.getLastRow() < 2) return;
  const range = sheet.getRange(2, CONFIG.COL_STATUS, sheet.getLastRow() - 1, 1);
  const rule  = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STATUSES, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function _styleHeader(sheet, row) {
  const range = sheet.getRange(row, 1, 1, sheet.getLastColumn() || 10);
  range.setBackground("#4A90D9")
       .setFontColor("#FFFFFF")
       .setFontWeight("bold");
}

function _capitalise(str) {
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}
