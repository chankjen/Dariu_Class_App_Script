/**
 * 🎨 FASHION DESIGN CLASS AUTOMATION SUITE
 * =========================================
 */

// ================= GLOBAL CONFIGURATION =================
const CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '',
  LESSON_PLAN_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('LESSON_PLAN_FOLDER_ID') || '',
  ATTENDANCE_SHEET_NAME: 'Attendance Records',
  LESSON_PLAN_SHEET_NAME: 'Lesson Plans',
  ROSTER_SHEET_NAME: 'Student Roster',
  TRIGGER_TIME_HOUR: 14,
  NOTIFICATION_EMAIL: Session.getActiveUser().getEmail()
};

// ================= SYNTHETIC DATA =================
const FASHION_DESIGN_DATA = {
  classInfo: {
    className: "Fashion Design 101",
    courseCode: "FD-101",
    instructor: "Prof. Alexandra Chen",
    semester: "Spring 2026",
    schedule: { days: ["Monday", "Wednesday"], time: "10:00 AM - 12:00 PM", room: "Design Studio B" }
  },
  students: [
    { id: "FD001", name: "Maya Johnson", email: "maya.j@university.edu", specialty: "Sustainable Fashion" },
    { id: "FD002", name: "James Liu", email: "james.l@university.edu", specialty: "Textile Innovation" },
    { id: "FD003", name: "Sofia Rodriguez", email: "sofia.r@university.edu", specialty: "Pattern Making" },
    { id: "FD004", name: "Aisha Okonkwo", email: "aisha.o@university.edu", specialty: "African Print Design" },
    { id: "FD005", name: "Lucas Martin", email: "lucas.m@university.edu", specialty: "Digital Fashion" },
    { id: "FD006", name: "Priya Sharma", email: "priya.s@university.edu", specialty: "Accessory Design" },
    { id: "FD007", name: "Ethan Kim", email: "ethan.k@university.edu", specialty: "Fashion Tech" },
    { id: "FD008", name: "Zara Ahmed", email: "zara.a@university.edu", specialty: "Couture Design" }
  ],
  lessonPlans: [
    {
      week: 1, date: "2026-01-13", topic: "Introduction to Fashion Design Principles",
      objectives: ["Understand design elements: line, shape, color, texture", "Explore fashion history timelines", "Complete initial sketch assignment"],
      materials: ["Sketch pads", "Color pencils", "Fashion magazines", "Projector"],
      activities: [
        { time: "10:00-10:30", activity: "Icebreaker: Fashion Inspiration Share" },
        { time: "10:30-11:15", activity: "Lecture: Design Elements & Principles" },
        { time: "11:15-11:45", activity: "Hands-on: Quick Sketch Challenge" },
        { time: "11:45-12:00", activity: "Q&A + Assignment Brief" }
      ],
      assessment: "Submit 3 fashion sketches by next class",
      resources: ["https://fashionhistory.edu/timeline", "https://colortheory.guide"]
    },
    {
      week: 2, date: "2026-01-20", topic: "Color Theory & Fabric Selection",
      objectives: ["Master color wheel applications", "Identify fabric properties", "Create a mood board"],
      materials: ["Fabric swatches", "Color wheels", "Glue sticks", "Poster boards"],
      activities: [
        { time: "10:00-10:45", activity: "Workshop: Color Harmony Exercises" },
        { time: "10:45-11:30", activity: "Fabric Lab: Touch & Analyze Session" },
        { time: "11:30-12:00", activity: "Mood Board Creation Start" }
      ],
      assessment: "Complete mood board with rationale",
      resources: ["https://pantone.com/fashion", "https://fabriclibrary.org"]
    }
  ],
  attendanceTemplate: [
    { date: "2026-01-13", records: [
      { studentId: "FD001", status: "Present", notes: "" }, { studentId: "FD002", status: "Present", notes: "" },
      { studentId: "FD003", status: "Late", notes: "10:15 AM arrival" }, { studentId: "FD004", status: "Present", notes: "" },
      { studentId: "FD005", status: "Absent", notes: "Medical excuse submitted" }, { studentId: "FD006", status: "Present", notes: "" },
      { studentId: "FD007", status: "Present", notes: "" }, { studentId: "FD008", status: "Present", notes: "" }
    ]},
    { date: "2026-01-20", records: [
      { studentId: "FD001", status: "Present", notes: "" }, { studentId: "FD002", status: "Absent", notes: "Travel notification" },
      { studentId: "FD003", status: "Present", notes: "" }, { studentId: "FD004", status: "Present", notes: "" },
      { studentId: "FD005", status: "Present", notes: "" }, { studentId: "FD006", status: "Present", notes: "" },
      { studentId: "FD007", status: "Late", notes: "10:20 AM arrival" }, { studentId: "FD008", status: "Present", notes: "" }
    ]}
  ]
};

// ================= 🔧 CONNECT TO EXISTING SPREADSHEET =================
function connectToExistingSpreadsheet() {
  try {
    Logger.log("🔍 Searching for Google Sheet named 'Lesson'...");
    
    const files = DriveApp.getFilesByName('Lesson');
    let spreadsheetFile = null;
    
    // Find the Google Sheet (not Apps Script)
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        spreadsheetFile = file;
        break;
      }
    }
    
    if (!spreadsheetFile) {
      throw new Error('❌ No Google Sheet named "Lesson" found in your Drive');
    }
    
    const ssId = spreadsheetFile.getId();
    const ssUrl = spreadsheetFile.getUrl();
    
    // Save to Properties
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ssId);
    CONFIG.SPREADSHEET_ID = ssId;
    
    Logger.log(`✅ Connected to: ${ssUrl}`);
    Logger.log(`📋 Spreadsheet ID: ${ssId}`);
    
    // Initialize
    const ss = SpreadsheetApp.openById(ssId);
    initializeSheets(ss);
    
    Logger.log('✅ Setup complete!');
    return ssId;
  } catch (error) {
    Logger.log('❌ Connection failed: ' + error.message);
    throw error;
  }
}

function initializeSheets(ss) {
  Logger.log('📄 Creating required sheets...');
  createAttendanceSheet(ss);
  createLessonPlanSheet(ss);
  createStudentRosterSheet(ss);
  Logger.log('✅ All sheets created/verified');
  
  Logger.log('📊 Populating with sample data...');
  populateSyntheticData(ss);
  
  Logger.log('⏰ Setting up automated triggers...');
  createPostClassReviewTrigger();
  
  Logger.log('✅ Initialization complete!');
}

// ================= ✅ FIXED: SMART SPREADSHEET ACCESS =================
function getSpreadsheet() {
  let ssId = CONFIG.SPREADSHEET_ID;
  
  if (!ssId || ssId === '') {
    Logger.log('🔍 No spreadsheet ID configured. Auto-discovering...');
    
    const files = DriveApp.getFilesByName('Lesson');
    let foundFile = null;
    
    // Search for Google Sheet specifically
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        foundFile = file;
        break;
      }
    }
    
    if (foundFile) {
      ssId = foundFile.getId();
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ssId);
      CONFIG.SPREADSHEET_ID = ssId;
      Logger.log(`✅ Found spreadsheet: ${foundFile.getUrl()}`);
    } else {
      throw new Error('❌ No Google Sheet named "Lesson" found. Run connectToExistingSpreadsheet() first.');
    }
  }
  
  try {
    return SpreadsheetApp.openById(ssId);
  } catch (e) {
    throw new Error(`❌ Cannot access spreadsheet. Error: ${e.message}`);
  }
}

// ================= SHEET CREATION =================
function createAttendanceSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ATTENDANCE_SHEET_NAME);
  }
  const headers = ['Date', 'Class Session', 'Student ID', 'Student Name', 'Status', 'Time In', 'Notes', 'Review Status', 'Reviewed By', 'Review Timestamp'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f3f4f6');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['Present', 'Absent', 'Late', 'Excused'], true).build();
  sheet.getRange('E2:E').setDataValidation(statusRule);
  return sheet;
}

function createLessonPlanSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.LESSON_PLAN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LESSON_PLAN_SHEET_NAME);
  }
  const headers = ['Week', 'Date', 'Topic', 'Learning Objectives', 'Materials Needed', 'Activities Timeline', 'Assessment', 'Resources', 'Status', 'Last Updated'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f3f4f6');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function createStudentRosterSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.ROSTER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ROSTER_SHEET_NAME);
  }
  const headers = ['Student ID', 'Name', 'Email', 'Specialty', 'Enrollment Date', 'Active'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f3f4f6');
  sheet.setFrozenRows(1);
  return sheet;
}

// ================= DATA POPULATION =================
function populateSyntheticData(ss) {
  const rosterSheet = ss.getSheetByName(CONFIG.ROSTER_SHEET_NAME);
  if (rosterSheet.getLastRow() <= 1) {
    rosterSheet.getRange(2, 1, FASHION_DESIGN_DATA.students.length, 6).setValues(
      FASHION_DESIGN_DATA.students.map(s => [s.id, s.name, s.email, s.specialty, '2026-01-06', 'Yes'])
    );
  }

  const lessonSheet = ss.getSheetByName(CONFIG.LESSON_PLAN_SHEET_NAME);
  if (lessonSheet.getLastRow() <= 1) {
    lessonSheet.getRange(2, 1, FASHION_DESIGN_DATA.lessonPlans.length, 10).setValues(
      FASHION_DESIGN_DATA.lessonPlans.map(lp => [
        lp.week, lp.date, lp.topic, lp.objectives.join(' | '), lp.materials.join(', '),
        lp.activities.map(a => `${a.time}: ${a.activity}`).join('\n'), lp.assessment,
        lp.resources.join(', '), 'Draft', new Date()
      ])
    );
  }

  const attSheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);
  if (attSheet.getLastRow() <= 1) {
    let attendanceData = [];
    FASHION_DESIGN_DATA.attendanceTemplate.forEach(session => {
      const weekObj = FASHION_DESIGN_DATA.lessonPlans.find(lp => lp.date === session.date);
      session.records.forEach(rec => {
        const student = FASHION_DESIGN_DATA.students.find(s => s.id === rec.studentId);
        attendanceData.push([
          session.date, `Week ${weekObj?.week || 'N/A'}`, rec.studentId, student?.name || 'Unknown',
          rec.status, rec.status === 'Late' ? '10:15 AM' : '10:00 AM', rec.notes, 'Pending', '', ''
        ]);
      });
    });
    if (attendanceData.length > 0) {
      attSheet.getRange(2, 1, attendanceData.length, attendanceData[0].length).setValues(attendanceData);
    }
  }
}

// ================= LESSON PLAN MANAGEMENT =================
function createLessonPlan(lessonPlanObj) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.LESSON_PLAN_SHEET_NAME);
  const newRow = [
    lessonPlanObj.week, lessonPlanObj.date, lessonPlanObj.topic,
    Array.isArray(lessonPlanObj.objectives) ? lessonPlanObj.objectives.join(' | ') : lessonPlanObj.objectives,
    Array.isArray(lessonPlanObj.materials) ? lessonPlanObj.materials.join(', ') : lessonPlanObj.materials,
    Array.isArray(lessonPlanObj.activities) ? lessonPlanObj.activities.map(a => `${a.time}: ${a.activity}`).join('\n') : lessonPlanObj.activities,
    lessonPlanObj.assessment, Array.isArray(lessonPlanObj.resources) ? lessonPlanObj.resources.join(', ') : lessonPlanObj.resources,
    lessonPlanObj.status || 'Draft', new Date()
  ];
  sheet.appendRow(newRow);
  Logger.log(`✅ Lesson plan added: Week ${lessonPlanObj.week}`);
}

function generateLessonPlanDoc(weekNumber) {
  const ss = getSpreadsheet();
  const data = ss.getSheetByName(CONFIG.LESSON_PLAN_SHEET_NAME).getDataRange().getValues();
  const lessonRow = data.find(row => row[0] == weekNumber);
  if (!lessonRow) throw new Error(`Week ${weekNumber} not found`);

  const doc = DocumentApp.create(`FD101_Week${weekNumber}_${lessonRow[2]}`);
  const body = doc.getBody();
  body.appendParagraph(`Fashion Design 101 - Week ${weekNumber}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Topic: ${lessonRow[2]}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(`Date: ${Utilities.formatDate(new Date(lessonRow[1]), Session.getScriptTimeZone(), 'MMM dd, yyyy')}\n`);
  body.appendParagraph('🎯 Learning Objectives:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  lessonRow[3].split(' | ').forEach(obj => body.appendParagraph(`• ${obj.trim()}`));
  body.appendParagraph('\n🧵 Materials Needed:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(lessonRow[4]);
  body.appendParagraph('\n📅 Class Activities:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  lessonRow[5].split('\n').forEach(act => body.appendParagraph(`🕐 ${act}`));
  body.appendParagraph('\n📝 Assessment:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(lessonRow[6]);
  body.appendParagraph('\n🔗 Resources:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(lessonRow[7]);
  
  Logger.log(`📄 Doc created: ${doc.getUrl()}`);
  return doc.getUrl();
}

// ================= ATTENDANCE TRACKING =================
function logAttendance(studentId, status, notes = '', timeIn = null) {
  const ss = getSpreadsheet();
  const rosterData = ss.getSheetByName(CONFIG.ROSTER_SHEET_NAME).getDataRange().getValues();
  const student = rosterData.find(row => row[0] === studentId);
  if (!student) throw new Error(`Student ${studentId} not found`);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const lessons = ss.getSheetByName(CONFIG.LESSON_PLAN_SHEET_NAME).getDataRange().getValues();
  const currentLesson = lessons.find(row => row[1] === today);
  const weekNum = currentLesson ? currentLesson[0] : 'N/A';

  ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME).appendRow([
    today, `Week ${weekNum}`, studentId, student[1], status, timeIn || new Date(), notes, 'Pending', '', ''
  ]);
  Logger.log(`✅ Attendance logged: ${student[1]} - ${status}`);
}

function bulkLogAttendance(attendanceRecords) {
  const ss = getSpreadsheet();
  const rows = attendanceRecords.map(rec => {
    const student = FASHION_DESIGN_DATA.students.find(s => s.id === rec.studentId);
    return [rec.date, rec.classSession || '', rec.studentId, student?.name || 'Unknown', rec.status, rec.timeIn || '', rec.notes || '', 'Pending', '', ''];
  });
  if (rows.length > 0) {
    ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME).getRange(ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME).getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return rows.length;
}

// ================= POST-CLASS REVIEW & TRIGGERS =================
function createPostClassReviewTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'postClassAttendanceReview') ScriptApp.deleteTrigger(t);
  });

  [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.WEDNESDAY].forEach(day => {
    ScriptApp.newTrigger('postClassAttendanceReview').timeBased().onWeekDay(day).atHour(CONFIG.TRIGGER_TIME_HOUR).nearMinute(0).create();
  });
  Logger.log('✅ Triggers created for Mon & Wed at 2:00 PM');
}

function postClassAttendanceReview() {
  Logger.log('🔍 Starting post-class review...');
  try {
    const ss = getSpreadsheet();
    const attSheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);
    const data = attSheet.getDataRange().getValues();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const todayRecords = data.slice(1).filter(row => row[0] === today);
    
    if (todayRecords.length === 0) {
      Logger.log(`ℹ️ No records for ${today}`);
      return null;
    }

    const summary = generateAttendanceSummary(todayRecords, today);
    updateReviewStatus(todayRecords, attSheet);
    sendAttendanceReportEmail(summary, todayRecords);
    logReviewActivity(summary, today);
    Logger.log('✅ Review completed');
    return summary;
  } catch (error) {
    Logger.log("❌ Review failed: " + error.message);
    throw error;
  }
}

function generateAttendanceSummary(records, date) {
  const total = records.length;
  const present = records.filter(r => r[4] === 'Present').length;
  const absent = records.filter(r => r[4] === 'Absent').length;
  const late = records.filter(r => r[4] === 'Late').length;
  const excused = records.filter(r => r[4] === 'Excused').length;
  const rate = total > 0 ? ((present / total) * 100).toFixed(1) + '%' : '0%';
  
  return {
    date, className: FASHION_DESIGN_DATA.classInfo.className, totalStudents: total,
    present, absent, late, excused, attendanceRate: rate,
    followUpNeeded: records.filter(r => (r[4] === 'Absent' || r[4] === 'Late') && r[6] && !r[6].toLowerCase().includes('excuse'))
      .map(r => ({ name: r[3], id: r[2], status: r[4], notes: r[6] })),
    timestamp: new Date()
  };
}

function updateReviewStatus(records, sheet) {
  const values = sheet.getDataRange().getValues();
  records.forEach(rec => {
    const idx = values.findIndex(row => row[0] === rec[0] && row[2] === rec[2]);
    if (idx > 0) {
      sheet.getRange(idx + 1, 8).setValue('Reviewed');
      sheet.getRange(idx + 1, 9).setValue('Automated Script');
      sheet.getRange(idx + 1, 10).setValue(new Date());
    }
  });
}

function sendAttendanceReportEmail(summary, records) {
  const classInfo = FASHION_DESIGN_DATA.classInfo;
  const subject = `📊 Attendance: ${summary.className} - ${summary.date}`;
  
  let html = `<h2>🎨 Attendance Review</h2>
    <p><strong>Class:</strong> ${classInfo.className}<br><strong>Date:</strong> ${summary.date}</p>
    <h3>📈 Summary</h3>
    <table border="1" cellpadding="8" style="border-collapse: collapse;">
      <tr style="background:#f0f0f0;"><th>Total</th><th>Present</th><th>Absent</th><th>Late</th><th>Rate</th></tr>
      <tr style="text-align:center;"><td>${summary.totalStudents}</td>
      <td style="color:green;">${summary.present}</td><td style="color:red;">${summary.absent}</td>
      <td style="color:orange;">${summary.late}</td><td><b>${summary.attendanceRate}</b></td></tr>
    </table>`;
    
  if (summary.followUpNeeded.length > 0) {
    html += `<h3>⚠️ Follow-Up Needed</h3><ul>${summary.followUpNeeded.map(s => `<li><b>${s.name}</b>: ${s.status}</li>`).join('')}</ul>`;
  }
  
  html += `<p><a href="https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}">View Sheet</a></p>`;
  
  MailApp.sendEmail({ to: CONFIG.NOTIFICATION_EMAIL, subject, htmlBody: html });
  Logger.log(`📧 Email sent to ${CONFIG.NOTIFICATION_EMAIL}`);
}

function logReviewActivity(summary, date) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('Review Logs');
  if (!sheet) {
    sheet = ss.insertSheet('Review Logs');
    sheet.getRange('A1:E1').setValues([['Review Date', 'Class Date', 'Rate', 'Follow-ups', 'Timestamp']]).setFontWeight('bold');
  }
  sheet.appendRow([new Date(), date, summary.attendanceRate, summary.followUpNeeded.length, summary.timestamp]);
}

// ================= UTILITIES =================
function getAttendanceReport(startDate, endDate) {
  const ss = getSpreadsheet();
  const data = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME).getDataRange().getValues().slice(1);
  const s = new Date(startDate), e = new Date(endDate);
  return data.filter(row => { const d = new Date(row[0]); return d >= s && d <= e; })
    .map(r => ({ date: r[0], studentId: r[2], studentName: r[3], status: r[4] }));
}

function exportAttendanceToCSV(date) {
  const ss = getSpreadsheet();
  const data = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME).getDataRange().getValues().filter(r => r[0] === date);
  if (data.length <= 1) return null;
  const csv = [['Date', 'Student ID', 'Name', 'Status', 'Notes'].join(','),
    ...data.slice(1).map(r => [r[0], r[2], r[3], r[4], r[6]].map(f => `"${String(f).replace(/"/g, '""')}"`).join(','))
  ].join('\n');
  const file = DriveApp.createFile(Utilities.newBlob(csv, 'text/csv', `Attendance_${date}.csv`));
  return file.getUrl();
}

// ================= UI & HELPERS =================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎨 Fashion Design Tools')
    .addItem('🔍 Run Review Now', 'postClassAttendanceReview')
    .addItem('📤 Export CSV (Today)', 'exportTodayCSV')
    .addItem('⚙️ Re-Connect', 'connectToExistingSpreadsheet')
    .addToUi();
}

function exportTodayCSV() {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const url = exportAttendanceToCSV(today);
  if (url) SpreadsheetApp.getUi().alert(`✅ CSV: ${url}`);
  else SpreadsheetApp.getUi().alert('ℹ️ No records today');
}

function testPostClassReview() {
  postClassAttendanceReview();
}

function verifySetup() {
  try {
    const ss = getSpreadsheet();
    Logger.log(`✅ Spreadsheet: ${ss.getName()} (${ss.getUrl()})`);
    Logger.log(`📄 Sheets: ${ss.getSheets().map(s => s.getName()).join(', ')}`);
    const roster = ss.getSheetByName(CONFIG.ROSTER_SHEET_NAME);
    if (roster && roster.getLastRow() > 1) Logger.log(`👥 Students: ${roster.getLastRow() - 1}`);
    Logger.log('✅ All systems operational!');
    return true;
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    return false;
  }
}
