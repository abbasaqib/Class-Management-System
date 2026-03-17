// Configuration - USE YOUR ACTUAL IDS
const SPREADSHEET_ID = '1-Pmiszn-nUqQyt0nabKYHFRfa2VmzdzMEOYktNNrPCY';
const ROOT_FOLDER_ID = '18YAdJmIdWqlSZn_fnLWfMDd3bV3bD1aZ'; // Root folder for all class materials

const SHEET_NAMES = {
  USERS: 'Users',
  CLASSES: 'Classes',
  LECTURES: 'Lectures',
  MEETINGS: 'Meetings',
  SUBMISSIONS: 'Submissions',
  SESSIONS: 'Sessions'
};

// Web App entry point
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Class Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== AUTHENTICATION FUNCTIONS ====================

function authenticate(userId, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    if (!usersSheet) {
      return { success: false, message: 'System not initialized. Please contact admin.' };
    }
    
    const data = usersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const storedUserId = data[i][0];
      const storedPassword = data[i][2];
      const status = data[i][7] || 'Active';
      
      if (storedUserId === userId) {
        if (status !== 'Active') {
          return { success: false, message: 'Account is inactive. Please contact admin.' };
        }
        
        if (storedPassword === password) {
          const sessionId = generateSessionId();
          const sessionsSheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
          
          if (sessionsSheet) {
            sessionsSheet.appendRow([
              sessionId,
              userId,
              new Date(),
              new Date(Date.now() + 8 * 60 * 60 * 1000)
            ]);
          }
          
          return {
            success: true,
            sessionId: sessionId,
            user: {
              userId: data[i][0],
              name: data[i][1],
              role: data[i][3],
              classId: data[i][4] || '',
              folderId: data[i][5] || '' // Student's personal folder ID
            }
          };
        } else {
          return { success: false, message: 'Invalid password' };
        }
      }
    }
    
    return { success: false, message: 'User ID not found' };
  } catch (error) {
    return { success: false, message: 'Login error: ' + error.toString() };
  }
}

function checkSession(sessionId) {
  try {
    if (!sessionId) return null;
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sessionsSheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
    
    if (!sessionsSheet) return null;
    
    const data = sessionsSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sessionId) {
        const expiryTime = new Date(data[i][3]);
        if (expiryTime > new Date()) {
          return getUserById(data[i][1]);
        }
      }
    }
    return null;
  } catch (error) {
    return null;
  }
}

function logout(sessionId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sessionsSheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
    
    if (!sessionsSheet) return;
    
    const data = sessionsSheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] === sessionId) {
        sessionsSheet.deleteRow(i + 1);
        break;
      }
    }
  } catch (error) {
    console.error('Logout error:', error);
  }
}

function getUserById(userId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    if (!usersSheet) return null;
    
    const data = usersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        return {
          userId: data[i][0],
          name: data[i][1],
          role: data[i][3],
          classId: data[i][4] || '',
          folderId: data[i][5] || ''
        };
      }
    }
    return null;
  } catch (error) {
    return null;
  }
}

// ==================== USER CRUD FUNCTIONS ====================

function getAllUsers(sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user) return [];
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    if (!usersSheet) return [];
    
    const data = usersSheet.getDataRange().getDisplayValues();
    const users = [];
    
    for (let i = 1; i < data.length; i++) {
      users.push({
        userId: data[i][0] || '',
        name: data[i][1] || '',
        role: data[i][3] || '',
        classId: data[i][4] || '',
        folderId: data[i][5] || '',
        status: data[i][7] || 'Active'
      });
    }
    
    return users;
  } catch (error) {
    return [];
  }
}

function getProfessors(sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user) return [];
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    if (!usersSheet) return [];
    
    const data = usersSheet.getDataRange().getDisplayValues();
    const professors = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === 'professor') {
        professors.push({
          userId: data[i][0] || '',
          name: data[i][1] || ''
        });
      }
    }
    
    return professors;
  } catch (error) {
    return [];
  }
}

function getStudentsByClass(classId, sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user) return [];
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    if (!usersSheet) return [];
    
    const data = usersSheet.getDataRange().getDisplayValues();
    const students = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === 'student' && data[i][4] === classId && data[i][7] === 'Active') {
        students.push({
          userId: data[i][0] || '',
          name: data[i][1] || '',
          folderId: data[i][5] || ''
        });
      }
    }
    
    return students;
  } catch (error) {
    return [];
  }
}

function createUser(userData, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Unauthorized' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    // Check if user exists
    const existingData = usersSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] === userData.userId) {
        return { success: false, message: 'User ID already exists' };
      }
    }
    
    // Create personal folder for student
    let personalFolderId = '';
    if (userData.role === 'student' && userData.classId) {
      try {
        const classFolder = getClassFolder(userData.classId);
        if (classFolder) {
          const studentFolder = classFolder.createFolder(userData.userId + ' - ' + userData.name);
          personalFolderId = studentFolder.getId();
        }
      } catch (folderError) {
        console.error('Error creating student folder:', folderError);
      }
    }
    
    usersSheet.appendRow([
      userData.userId,
      userData.name,
      userData.password,
      userData.role,
      userData.classId || '',
      personalFolderId || '',
      new Date(),
      'Active'
    ]);
    
    return { success: true, message: 'User created successfully' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function updateUser(userId, updateData, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Unauthorized' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    const data = usersSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        const row = i + 1;
        if (updateData.name) usersSheet.getRange(row, 2).setValue(updateData.name);
        if (updateData.role) usersSheet.getRange(row, 4).setValue(updateData.role);
        if (updateData.classId !== undefined) usersSheet.getRange(row, 5).setValue(updateData.classId);
        if (updateData.status) usersSheet.getRange(row, 8).setValue(updateData.status);
        if (updateData.password) usersSheet.getRange(row, 3).setValue(updateData.password);
        
        // Update folder if class changed for student
        if (updateData.role === 'student' && updateData.classId && updateData.classId !== data[i][4]) {
          try {
            const classFolder = getClassFolder(updateData.classId);
            if (classFolder) {
              const studentFolder = classFolder.createFolder(userId + ' - ' + (updateData.name || data[i][1]));
              usersSheet.getRange(row, 6).setValue(studentFolder.getId());
            }
          } catch (folderError) {
            console.error('Error updating student folder:', folderError);
          }
        }
        
        return { success: true, message: 'User updated successfully' };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function deleteUser(userId, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Unauthorized' };
    }
    
    if (userId === admin.userId) {
      return { success: false, message: 'Cannot delete your own account' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    const data = usersSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        usersSheet.deleteRow(i + 1);
        return { success: true, message: 'User deleted successfully' };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ==================== CLASS CRUD FUNCTIONS ====================

function getClassFolder(classId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const classesSheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
    
    if (!classesSheet) return null;
    
    const data = classesSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === classId) {
        const folderId = data[i][3];
        if (folderId) {
          return DriveApp.getFolderById(folderId);
        }
        break;
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting class folder:', error);
    return null;
  }
}

function getClassesList(sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user) return [];
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const classesSheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
    
    if (!classesSheet) return [];
    
    const data = classesSheet.getDataRange().getDisplayValues();
    const classes = [];
    
    for (let i = 1; i < data.length; i++) {
      classes.push({
        classId: data[i][0] || '',
        className: data[i][1] || '',
        professorId: data[i][2] || '',
        folderId: data[i][3] || '',
        status: data[i][5] || 'Active'
      });
    }
    
    return classes;
  } catch (error) {
    return [];
  }
}

function getProfessorClasses(professorId, sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user) return [];
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const classesSheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
    
    if (!classesSheet) return [];
    
    const data = classesSheet.getDataRange().getDisplayValues();
    const classes = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === professorId) {
        classes.push({
          classId: data[i][0] || '',
          className: data[i][1] || '',
          folderId: data[i][3] || ''
        });
      }
    }
    
    return classes;
  } catch (error) {
    return [];
  }
}

function createClass(className, professorId, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Unauthorized' };
    }
    
    // Create folder in Google Drive
    let classFolderId = '';
    try {
      const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
      const classFolder = rootFolder.createFolder(className);
      classFolderId = classFolder.getId();
    } catch (folderError) {
      console.error('Folder creation error:', folderError);
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const classesSheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
    
    const classId = 'CLASS' + String(Math.floor(Math.random() * 10000)).padStart(3, '0');
    
    classesSheet.appendRow([
      classId,
      className,
      professorId,
      classFolderId,
      new Date(),
      'Active'
    ]);
    
    return { success: true, message: 'Class created successfully', classId: classId };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function updateClass(classId, updateData, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Unauthorized' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const classesSheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
    
    const data = classesSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === classId) {
        const row = i + 1;
        if (updateData.className) classesSheet.getRange(row, 2).setValue(updateData.className);
        if (updateData.professorId) classesSheet.getRange(row, 3).setValue(updateData.professorId);
        if (updateData.status) classesSheet.getRange(row, 6).setValue(updateData.status);
        return { success: true, message: 'Class updated successfully' };
      }
    }
    return { success: false, message: 'Class not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function deleteClass(classId, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Unauthorized' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const classesSheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
    
    const data = classesSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === classId) {
        classesSheet.deleteRow(i + 1);
        return { success: true, message: 'Class deleted successfully' };
      }
    }
    return { success: false, message: 'Class not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ==================== LECTURE CRUD FUNCTIONS ====================

function getClassLectures(classId, sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user) return [];
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const lecturesSheet = ss.getSheetByName(SHEET_NAMES.LECTURES);
    
    if (!lecturesSheet) return [];
    
    const data = lecturesSheet.getDataRange().getDisplayValues();
    const lectures = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === classId) {
        lectures.push({
          lectureId: data[i][0] || '',
          classId: data[i][1] || '',
          title: data[i][2] || '',
          description: data[i][3] || '',
          videoUrl: data[i][4] || '',
          createdDate: data[i][5] || new Date(),
          createdBy: data[i][6] || ''
        });
      }
    }
    
    return lectures;
  } catch (error) {
    return [];
  }
}

function addLecture(lectureData, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Only admins can add lectures' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const lecturesSheet = ss.getSheetByName(SHEET_NAMES.LECTURES);
    
    const lectureId = 'LEC' + String(Math.floor(Math.random() * 10000)).padStart(3, '0');
    
    lecturesSheet.appendRow([
      lectureId,
      lectureData.classId,
      lectureData.title,
      lectureData.description || '',
      lectureData.videoUrl,
      new Date(),
      admin.userId
    ]);
    
    return { success: true, message: 'Lecture added successfully', lectureId: lectureId };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function updateLecture(lectureId, updateData, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Only admins can update lectures' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const lecturesSheet = ss.getSheetByName(SHEET_NAMES.LECTURES);
    
    const data = lecturesSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === lectureId) {
        const row = i + 1;
        if (updateData.title) lecturesSheet.getRange(row, 3).setValue(updateData.title);
        if (updateData.description !== undefined) lecturesSheet.getRange(row, 4).setValue(updateData.description);
        if (updateData.videoUrl) lecturesSheet.getRange(row, 5).setValue(updateData.videoUrl);
        return { success: true, message: 'Lecture updated successfully' };
      }
    }
    return { success: false, message: 'Lecture not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function deleteLecture(lectureId, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Only admins can delete lectures' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const lecturesSheet = ss.getSheetByName(SHEET_NAMES.LECTURES);
    
    const data = lecturesSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === lectureId) {
        lecturesSheet.deleteRow(i + 1);
        return { success: true, message: 'Lecture deleted successfully' };
      }
    }
    return { success: false, message: 'Lecture not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ==================== MEETING CRUD FUNCTIONS ====================

function getClassMeetings(classId, sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user) return [];
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const meetingsSheet = ss.getSheetByName(SHEET_NAMES.MEETINGS);
    
    if (!meetingsSheet) return [];
    
    const data = meetingsSheet.getDataRange().getDisplayValues();
    const meetings = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === classId) {
        meetings.push({
          meetingId: data[i][0] || '',
          classId: data[i][1] || '',
          title: data[i][2] || '',
          description: data[i][3] || '',
          meetingLink: data[i][4] || '',
          meetingDate: data[i][5] || '',
          startTime: data[i][6] || '',
          endTime: data[i][7] || '',
          createdDate: data[i][8] || new Date(),
          createdBy: data[i][9] || ''
        });
      }
    }
    
    // Sort by date
    meetings.sort((a, b) => {
      if (a.meetingDate && b.meetingDate) {
        return new Date(a.meetingDate + ' ' + a.startTime) - new Date(b.meetingDate + ' ' + b.startTime);
      }
      return 0;
    });
    
    return meetings;
  } catch (error) {
    return [];
  }
}

function addMeeting(meetingData, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Only admins can add meetings' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const meetingsSheet = ss.getSheetByName(SHEET_NAMES.MEETINGS);
    
    const meetingId = 'MTG' + String(Math.floor(Math.random() * 10000)).padStart(3, '0');
    
    meetingsSheet.appendRow([
      meetingId,
      meetingData.classId,
      meetingData.title,
      meetingData.description || '',
      meetingData.meetingLink,
      meetingData.meetingDate,
      meetingData.startTime,
      meetingData.endTime,
      new Date(),
      admin.userId
    ]);
    
    return { success: true, message: 'Meeting added successfully', meetingId: meetingId };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function updateMeeting(meetingId, updateData, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Only admins can update meetings' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const meetingsSheet = ss.getSheetByName(SHEET_NAMES.MEETINGS);
    
    const data = meetingsSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === meetingId) {
        const row = i + 1;
        if (updateData.title) meetingsSheet.getRange(row, 3).setValue(updateData.title);
        if (updateData.description !== undefined) meetingsSheet.getRange(row, 4).setValue(updateData.description);
        if (updateData.meetingLink) meetingsSheet.getRange(row, 5).setValue(updateData.meetingLink);
        if (updateData.meetingDate) meetingsSheet.getRange(row, 6).setValue(updateData.meetingDate);
        if (updateData.startTime) meetingsSheet.getRange(row, 7).setValue(updateData.startTime);
        if (updateData.endTime) meetingsSheet.getRange(row, 8).setValue(updateData.endTime);
        return { success: true, message: 'Meeting updated successfully' };
      }
    }
    return { success: false, message: 'Meeting not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function deleteMeeting(meetingId, sessionId) {
  try {
    const admin = checkSession(sessionId);
    if (!admin || admin.role !== 'admin') {
      return { success: false, message: 'Only admins can delete meetings' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const meetingsSheet = ss.getSheetByName(SHEET_NAMES.MEETINGS);
    
    const data = meetingsSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === meetingId) {
        meetingsSheet.deleteRow(i + 1);
        return { success: true, message: 'Meeting deleted successfully' };
      }
    }
    return { success: false, message: 'Meeting not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ==================== SUBMISSION FUNCTIONS WITH FILE UPLOAD ====================

function uploadAudioToDrive(base64Data, fileName, classId, studentId, sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user || user.role !== 'student') {
      return { success: false, message: 'Unauthorized' };
    }
    
    // Get student's folder
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getDisplayValues();
    
    let studentFolderId = '';
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][0] === studentId) {
        studentFolderId = usersData[i][5];
        break;
      }
    }
    
    if (!studentFolderId) {
      return { success: false, message: 'Student folder not found' };
    }
    
    // Decode base64 data
    const decodedData = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(decodedData, 'audio/webm', fileName);
    
    // Upload to student's folder
    const studentFolder = DriveApp.getFolderById(studentFolderId);
    const file = studentFolder.createFile(blob);
    
    return { 
      success: true, 
      fileId: file.getId(), 
      fileUrl: file.getUrl(),
      fileName: file.getName()
    };
    
  } catch (error) {
    return { success: false, message: 'Upload error: ' + error.toString() };
  }
}

function submitToLecture(submissionData, sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user || user.role !== 'student') {
      return { success: false, message: 'Unauthorized' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const submissionsSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
    
    const submissionId = 'SUB' + String(Math.floor(Math.random() * 10000)).padStart(3, '0');
    
    submissionsSheet.appendRow([
      submissionId,
      submissionData.lectureId,
      user.userId,
      submissionData.content, // File URL or text
      submissionData.type || 'audio',
      new Date(),
      submissionData.parentSubmissionId || '',
      submissionData.grade || '',
      submissionData.fileName || ''
    ]);
    
    return { success: true, message: 'Submitted successfully', submissionId: submissionId };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function getLectureThread(lectureId, sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user) return [];
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const submissionsSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
    
    if (!submissionsSheet) return [];
    
    const data = submissionsSheet.getDataRange().getDisplayValues();
    const threads = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === lectureId) {
        threads.push({
          submissionId: data[i][0] || '',
          lectureId: data[i][1] || '',
          userId: data[i][2] || '',
          content: data[i][3] || '',
          type: data[i][4] || 'audio',
          submittedDate: data[i][5] || new Date(),
          parentSubmissionId: data[i][6] || '',
          grade: data[i][7] || '',
          fileName: data[i][8] || ''
        });
      }
    }
    
    // Organize into threads
    const threadMap = {};
    const rootThreads = [];
    
    threads.forEach(item => {
      threadMap[item.submissionId] = { ...item, replies: [] };
    });
    
    threads.forEach(item => {
      if (item.parentSubmissionId && threadMap[item.parentSubmissionId]) {
        threadMap[item.parentSubmissionId].replies.push(threadMap[item.submissionId]);
      } else if (!item.parentSubmissionId) {
        rootThreads.push(threadMap[item.submissionId]);
      }
    });
    
    // Sort by date
    rootThreads.sort((a, b) => new Date(a.submittedDate) - new Date(b.submittedDate));
    
    return rootThreads;
  } catch (error) {
    return [];
  }
}

function getStudentSubmissions(studentId, sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user) return [];
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const submissionsSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
    
    if (!submissionsSheet) return [];
    
    const data = submissionsSheet.getDataRange().getDisplayValues();
    const submissions = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === studentId && data[i][4] === 'audio') {
        submissions.push({
          submissionId: data[i][0] || '',
          lectureId: data[i][1] || '',
          content: data[i][3] || '',
          submittedDate: data[i][5] || new Date(),
          grade: data[i][7] || '',
          fileName: data[i][8] || ''
        });
      }
    }
    
    return submissions;
  } catch (error) {
    return [];
  }
}

function updateSubmissionGrade(submissionId, grade, sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user || user.role !== 'professor') {
      return { success: false, message: 'Unauthorized' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const submissionsSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
    
    const data = submissionsSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === submissionId) {
        const row = i + 1;
        submissionsSheet.getRange(row, 8).setValue(grade);
        return { success: true, message: 'Grade updated successfully' };
      }
    }
    return { success: false, message: 'Submission not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ==================== DASHBOARD STATS ====================

function getDashboardStats(sessionId) {
  try {
    const user = checkSession(sessionId);
    if (!user || user.role !== 'admin') return null;
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const classesSheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
    const lecturesSheet = ss.getSheetByName(SHEET_NAMES.LECTURES);
    const submissionsSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
    
    let totalUsers = 0, totalProfessors = 0, totalStudents = 0;
    let totalClasses = 0, totalLectures = 0, totalSubmissions = 0, totalGraded = 0;
    
    if (usersSheet) {
      const users = usersSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < users.length; i++) {
        if (users[i][7] !== 'Inactive') {
          totalUsers++;
          if (users[i][3] === 'professor') totalProfessors++;
          else if (users[i][3] === 'student') totalStudents++;
        }
      }
    }
    
    if (classesSheet) {
      totalClasses = classesSheet.getLastRow() - 1;
      if (totalClasses < 0) totalClasses = 0;
    }
    
    if (lecturesSheet) {
      totalLectures = lecturesSheet.getLastRow() - 1;
      if (totalLectures < 0) totalLectures = 0;
    }
    
    if (submissionsSheet) {
      const submissions = submissionsSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < submissions.length; i++) {
        if (submissions[i][4] === 'audio') {
          totalSubmissions++;
          if (submissions[i][7]) totalGraded++;
        }
      }
    }
    
    return {
      totalUsers,
      totalProfessors,
      totalStudents,
      totalClasses,
      totalLectures,
      totalSubmissions,
      totalGraded
    };
  } catch (error) {
    return null;
  }
}

// ==================== HELPER FUNCTIONS ====================

function generateSessionId() {
  return Utilities.getUuid();
}
