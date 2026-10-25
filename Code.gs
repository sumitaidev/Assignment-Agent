const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const DRIVE_FOLDER_NAME = "Assignment_Submissions";

// Serves the HTML Web App
function doGet(e) {
  // Safely check if 'e' and 'e.parameter' exist
  let page = (e && e.parameter && e.parameter.page === 'Student') ? 'Student' : 'Index';
  
  return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .setTitle(page === 'Index' ? 'Teacher Portal' : 'Student Submission')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------- STUDENT MANAGEMENT -----------------
function getStudents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sheet.getDataRange().getValues();
  let students = [];
  for (let i = 1; i < data.length; i++) {
    students.push({ id: data[i][0], name: data[i][1], email: data[i][2] });
  }
  return students;
}

function addStudent(name, email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const id = 'STU' + new Date().getTime(); // Generate unique ID
  sheet.appendRow([id, name, email]);
  return getStudents();
}

function removeStudent(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getStudents();
}

// ----------------- ASSIGNMENT MANAGEMENT -----------------
function getAssignments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Assignments");
  const data = sheet.getDataRange().getValues();
  let assignments = [];
  for (let i = 1; i < data.length; i++) {
    assignments.push({
      id: data[i][0],
      studentId: data[i][1],
      studentName: data[i][2],
      title: data[i][3],
      fileUrl: data[i][4],
      status: data[i][5] || 'Pending',
      mark: data[i][6] || '',
      review: data[i][7] || ''
    });
  }
  return assignments;
}

function gradeAssignment(id, mark, review, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Assignments");
  const data = sheet.getDataRange().getValues();
  let studentId = "";
  let title = "";
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 6).setValue(status);
      sheet.getRange(i + 1, 7).setValue(mark);
      sheet.getRange(i + 1, 8).setValue(review);
      
      studentId = data[i][1];
      title = data[i][3];
      break;
    }
  }
  
  // Gmail API Integration: Notify the student automatically
  notifyStudent(studentId, title, status, mark, review);
  return getAssignments();
}

// ----------------- GMAIL API INTEGRATION -----------------
function notifyStudent(studentId, title, status, mark, review) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sheet.getDataRange().getValues();
  let email = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      email = data[i][2];
      break;
    }
  }
  
  if (email) {
    const subject = `Assignment Update: ${title}`;
    const body = `Hello,\n\nYour assignment "${title}" has been reviewed by your teacher.\n\nStatus: ${status}\nMark: ${mark}\nTeacher's Feedback: ${review}\n\nPlease check your portal for more details.\n\nBest Regards,\nTeacher Portal Automation`;
    GmailApp.sendEmail(email, subject, body);
  }
}

// ----------------- GOOGLE DRIVE API INTEGRATION -----------------
// (Use this function if you build a student-facing submission form)
function uploadAssignmentFile(studentId, studentName, title, base64Data, fileName) {
  // 1. Get or create folder
  let folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  
  // 2. Decode and save file to Drive
  let blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.PDF, fileName);
  let file = folder.createFile(blob);
  let fileUrl = file.getUrl();
  
  // 3. Record in database
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Assignments");
  const id = 'ASN' + new Date().getTime();
  sheet.appendRow([id, studentId, studentName, title, fileUrl, "Pending", "", ""]);
  
  return "Successfully Uploaded & Saved to Drive";
}
