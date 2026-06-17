// ============================================================
// CLASS MANAGEMENT SYSTEM - Code.gs
// Google Apps Script Backend
// ============================================================

// ---- SHEET NAMES ----
const SHEETS = {
  STUDENTS: 'Students',
  ADMINS: 'Admins',
  ATTENDANCE: 'Attendance',
  DOCUMENTS: 'Documents',
  MATERIALS: 'Materials',
  RESULTS: 'Results'
};

// ---- COLUMN DEFINITIONS ----
const COLS = {
  STUDENTS:    ['Enrollment Number', 'Name'],
  ADMINS:      ['Enrollment Number', 'Name'],
  ATTENDANCE:  ['Enrollment Number', 'Name', 'Total Classes', 'उपस्थित (Present)', 'Percentage'],
  DOCUMENTS:   ['Timestamp', 'Student Name', 'Enrollment Number', 'Selected Admin', 'Description', 'File URL'],
  MATERIALS:   ['Timestamp', 'Admin Name', 'Title', 'Description', 'File URL'],
  RESULTS:     ['Enrollment Number', 'Name', 'Semester', 'CGPA']
};

// ============================================================
// ENTRY POINT
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Class Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================================
// SHEET SETUP
// ============================================================
function getOrCreateSheet(ss, name, cols) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(cols);
    sheet.getRange(1, 1, 1, cols.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
  } else {
    // Ensure columns exist
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    cols.forEach(function(col) {
      if (existingHeaders.indexOf(col) === -1) {
        var nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol).setValue(col).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
      }
    });
  }
  return sheet;
}

function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function(key) {
    getOrCreateSheet(ss, SHEETS[key], COLS[key]);
  });
  return { success: true, message: 'Sheets initialized.' };
}

// ============================================================
// AUTHENTICATION
// ============================================================
function login(enrollmentNumber, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var expectedPwd = String(enrollmentNumber).slice(-4);
    if (password !== expectedPwd) {
      return { success: false, message: 'Invalid Login' };
    }

    // Check student
    var studentSheet = getOrCreateSheet(ss, SHEETS.STUDENTS, COLS.STUDENTS);
    var studentData = studentSheet.getDataRange().getValues();
    for (var i = 1; i < studentData.length; i++) {
      if (String(studentData[i][0]).trim() === String(enrollmentNumber).trim()) {
        return {
          success: true,
          message: 'Valid Login',
          role: 'student',
          name: studentData[i][1],
          enrollmentNumber: String(enrollmentNumber)
        };
      }
    }

    // Check admin
    var adminSheet = getOrCreateSheet(ss, SHEETS.ADMINS, COLS.ADMINS);
    var adminData = adminSheet.getDataRange().getValues();
    for (var j = 1; j < adminData.length; j++) {
      if (String(adminData[j][0]).trim() === String(enrollmentNumber).trim()) {
        return {
          success: true,
          message: 'Valid Login',
          role: 'admin',
          name: adminData[j][1],
          enrollmentNumber: String(enrollmentNumber)
        };
      }
    }

    return { success: false, message: 'Invalid Login' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// STUDENT FUNCTIONS
// ============================================================
function getStudentAttendance(enrollmentNumber) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEETS.ATTENDANCE, COLS.ATTENDANCE);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var enIdx = headers.indexOf('Enrollment Number');
    var nameIdx = headers.indexOf('Name');
    var totalIdx = headers.indexOf('Total Classes');
    var presentIdx = headers.indexOf('उपस्थित (Present)');
    var pctIdx = headers.indexOf('Percentage');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][enIdx]).trim() === String(enrollmentNumber).trim()) {
        return {
          success: true,
          enrollmentNumber: data[i][enIdx],
          name: data[i][nameIdx],
          totalClasses: data[i][totalIdx] || 0,
          present: data[i][presentIdx] || 0,
          percentage: data[i][pctIdx] || 0
        };
      }
    }
    return { success: false, message: 'No attendance record found.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAdminList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEETS.ADMINS, COLS.ADMINS);
    var data = sheet.getDataRange().getValues();
    var admins = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) admins.push({ enrollmentNumber: data[i][0], name: data[i][1] });
    }
    return { success: true, admins: admins };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function uploadDocument(studentName, enrollmentNumber, selectedAdmin, description, fileUrl) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEETS.DOCUMENTS, COLS.DOCUMENTS);
    sheet.appendRow([
      new Date().toLocaleString(),
      studentName,
      enrollmentNumber,
      selectedAdmin,
      description,
      fileUrl
    ]);
    return { success: true, message: 'Document uploaded successfully.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function uploadCGPA(enrollmentNumber, name, semester, cgpa) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEETS.RESULTS, COLS.RESULTS);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var enIdx = headers.indexOf('Enrollment Number');
    var nameIdx = headers.indexOf('Name');
    var semIdx = headers.indexOf('Semester');
    var cgpaIdx = headers.indexOf('CGPA');

    // Check if row exists for this student + semester
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][enIdx]).trim() === String(enrollmentNumber).trim() &&
          String(data[i][semIdx]).trim() === String(semester).trim()) {
        sheet.getRange(i + 1, cgpaIdx + 1).setValue(cgpa);
        return { success: true, message: 'CGPA updated for Semester ' + semester };
      }
    }
    // New row
    var row = [];
    row[enIdx] = enrollmentNumber;
    row[nameIdx] = name;
    row[semIdx] = semester;
    row[cgpaIdx] = cgpa;
    // Fill gaps
    for (var k = 0; k < headers.length; k++) {
      if (row[k] === undefined) row[k] = '';
    }
    sheet.appendRow(row);
    return { success: true, message: 'CGPA saved for Semester ' + semester };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getStudentCGPA(enrollmentNumber) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEETS.RESULTS, COLS.RESULTS);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var enIdx = headers.indexOf('Enrollment Number');
    var semIdx = headers.indexOf('Semester');
    var cgpaIdx = headers.indexOf('CGPA');
    var results = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][enIdx]).trim() === String(enrollmentNumber).trim()) {
        results.push({ semester: data[i][semIdx], cgpa: data[i][cgpaIdx] });
      }
    }
    return { success: true, results: results };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getMaterials() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEETS.MATERIALS, COLS.MATERIALS);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var materials = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var row = {};
        headers.forEach(function(h, idx) { row[h] = data[i][idx]; });
        materials.push(row);
      }
    }
    return { success: true, materials: materials.reverse() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// ADMIN FUNCTIONS
// ============================================================
function getAllStudents() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var studSheet = getOrCreateSheet(ss, SHEETS.STUDENTS, COLS.STUDENTS);
    var attSheet = getOrCreateSheet(ss, SHEETS.ATTENDANCE, COLS.ATTENDANCE);
    var studData = studSheet.getDataRange().getValues();
    var attData = attSheet.getDataRange().getValues();
    var attHeaders = attData[0];
    var attEnIdx = attHeaders.indexOf('Enrollment Number');
    var attPresentIdx = attHeaders.indexOf('उपस्थित (Present)');
    var attTotalIdx = attHeaders.indexOf('Total Classes');
    var attPctIdx = attHeaders.indexOf('Percentage');

    var attMap = {};
    for (var a = 1; a < attData.length; a++) {
      attMap[String(attData[a][attEnIdx]).trim()] = {
        present: attData[a][attPresentIdx] || 0,
        total: attData[a][attTotalIdx] || 0,
        percentage: attData[a][attPctIdx] || 0
      };
    }

    var students = [];
    for (var i = 1; i < studData.length; i++) {
      if (!studData[i][0]) continue;
      var en = String(studData[i][0]).trim();
      var att = attMap[en] || { present: 0, total: 0, percentage: 0 };
      students.push({
        enrollmentNumber: en,
        name: studData[i][1],
        present: att.present,
        total: att.total,
        percentage: att.percentage
      });
    }
    return { success: true, students: students };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAllDocuments() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEETS.DOCUMENTS, COLS.DOCUMENTS);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var docs = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var row = {};
        headers.forEach(function(h, idx) { row[h] = data[i][idx]; });
        docs.push(row);
      }
    }
    return { success: true, documents: docs.reverse() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function uploadMaterial(adminName, title, description, fileUrl) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEETS.MATERIALS, COLS.MATERIALS);
    sheet.appendRow([
      new Date().toLocaleString(),
      adminName,
      title,
      description,
      fileUrl
    ]);
    return { success: true, message: 'Material uploaded successfully.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateAttendance(enrollmentNumber, name, totalClasses, present) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEETS.ATTENDANCE, COLS.ATTENDANCE);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var enIdx = headers.indexOf('Enrollment Number');
    var nameIdx = headers.indexOf('Name');
    var totalIdx = headers.indexOf('Total Classes');
    var presentIdx = headers.indexOf('उपस्थित (Present)');
    var pctIdx = headers.indexOf('Percentage');

    var percentage = totalClasses > 0 ? ((present / totalClasses) * 100).toFixed(2) : 0;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][enIdx]).trim() === String(enrollmentNumber).trim()) {
        sheet.getRange(i + 1, nameIdx + 1).setValue(name);
        sheet.getRange(i + 1, totalIdx + 1).setValue(totalClasses);
        sheet.getRange(i + 1, presentIdx + 1).setValue(present);
        sheet.getRange(i + 1, pctIdx + 1).setValue(percentage);
        return { success: true, message: 'Attendance updated.' };
      }
    }
    // New row
    var newRow = [];
    newRow[enIdx] = enrollmentNumber;
    newRow[nameIdx] = name;
    newRow[totalIdx] = totalClasses;
    newRow[presentIdx] = present;
    newRow[pctIdx] = percentage;
    for (var k = 0; k < headers.length; k++) {
      if (newRow[k] === undefined) newRow[k] = '';
    }
    sheet.appendRow(newRow);
    return { success: true, message: 'Attendance added.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAllCGPA() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEETS.RESULTS, COLS.RESULTS);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var results = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var row = {};
        headers.forEach(function(h, idx) { row[h] = data[i][idx]; });
        results.push(row);
      }
    }
    return { success: true, results: results };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function bulkUpdateAttendance(updates) {
  try {
    updates.forEach(function(u) {
      updateAttendance(u.enrollmentNumber, u.name, u.totalClasses, u.present);
    });
    return { success: true, message: 'Bulk attendance updated.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// TIMETABLE (static - admin can customize later)
// ============================================================
function getTimetable() {
  return {
    success: true,
    timetable: [
      { day: 'Monday',    periods: ['Math 9:00', 'Physics 10:00', 'Break', 'CS 12:00', 'English 1:00'] },
      { day: 'Tuesday',   periods: ['Chemistry 9:00', 'Math 10:00', 'Break', 'CS Lab 12:00', 'Sports 2:00'] },
      { day: 'Wednesday', periods: ['English 9:00', 'Physics Lab 10:00', 'Break', 'Math 12:00', 'Library 1:00'] },
      { day: 'Thursday',  periods: ['CS 9:00', 'Chemistry Lab 10:00', 'Break', 'Physics 12:00', 'English 1:00'] },
      { day: 'Friday',    periods: ['Math 9:00', 'CS 10:00', 'Break', 'Project Work 12:00', 'Seminar 2:00'] },
      { day: 'Saturday',  periods: ['Revision 9:00', 'Doubt Session 10:00', 'Break', 'Practical 12:00'] }
    ]
  };
}
