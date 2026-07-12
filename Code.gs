function doGet() {
  return HtmlService.createHtmlOutputFromFile('Student Attendance System')
    .setTitle('Student Attendance System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function initializeUserCache() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) throw new Error('Users sheet not found');
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  // Read all columns
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  const userDataCache = {};
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0]) {
      userDataCache[data[i][0]] = {
        password: data[i][1],
        type: data[i][2] || 'teacher',
        teacherName: data[i][3] || data[i][0],
        allowedDays: data[i][4] || '',
        startTime: data[i][5] || '',
        endTime: data[i][6] || '',
        status: data[i][12] || 'Active', // Status column
        creationDate: data[i][11], // Creation_Date column
        inactiveDate: data[i][13] // Inactive_Date column
      };
    }
  }
  
  CacheService.getScriptCache().put('userDataCache', JSON.stringify(userDataCache), 300);
}

// Us date pe jin class-sections ki NORMAL attendance mark ho chuki hai un ki list
function getMarkedClassSections(date) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Attendance');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const headers = data[0];
    const dateCol = headers.indexOf('Date');
    const classCol = headers.indexOf('Class_Section');
    const typeCol = headers.indexOf('Attendance_Type');
    if (dateCol === -1 || classCol === -1) return [];
    const targetDate = new Date(date).toLocaleDateString('en-CA');
    const marked = new Set();
    for (let i = 1; i < data.length; i++) {
      let rowDate;
      try { rowDate = new Date(data[i][dateCol]).toLocaleDateString('en-CA'); } catch (e) { continue; }
      if (rowDate !== targetDate) continue;
      if (typeCol !== -1 && data[i][typeCol] === 'Episodic') continue; // sirf normal attendance
      if (data[i][classCol]) marked.add(data[i][classCol]);
    }
    return Array.from(marked);
  } catch (e) {
    Logger.log('Error in getMarkedClassSections: ' + e.message);
    return [];
  }
}

/**
 * Forces immediate cache refresh for all users
 * Call this after any user status change
 */
function forceCacheRefresh() {
  const cache = CacheService.getScriptCache();
  cache.remove('userDataCache');
  initializeUserCache();
}

/**
 * Updates user status (Active/Inactive) and sets Inactive_Date when status changes to Inactive
 * Also invalidates cache immediately
 * @param {string} username - The username to update
 * @param {string} newStatus - New status ('Active' or 'Inactive')
 * @returns {boolean} - Success status
 */
function updateUserStatus(username, newStatus) {
  try {
    const cache = CacheService.getScriptCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) throw new Error('Users sheet not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const usernameCol = headers.indexOf('Username');
    const statusCol = headers.indexOf('Status');
    const inactiveDateCol = headers.indexOf('Inactive_Date');
    const creationDateCol = headers.indexOf('Creation_Date');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameCol] === username) {
        const currentStatus = data[i][statusCol] || 'Active';
        const row = i + 1;
        
        // Update status
        sheet.getRange(row, statusCol + 1).setValue(newStatus);
        
        // If status is changing to Inactive, set Inactive_Date
        if (newStatus === 'Inactive' && currentStatus !== 'Inactive') {
          const inactiveDate = new Date();
          sheet.getRange(row, inactiveDateCol + 1).setValue(inactiveDate);
          
          // CRITICAL: Force immediate cache invalidation
          cache.remove(`user_${username}`); // Remove individual user cache
          cache.remove('userDataCache'); // Remove entire user cache
          
          // Also invalidate any active sessions
          const userCache = cache.get(`session_${username}`);
          if (userCache) {
            cache.remove(`session_${username}`);
          }
        }
        
        // If status is changing to Active, clear Inactive_Date
        if (newStatus === 'Active' && currentStatus === 'Inactive') {
          sheet.getRange(row, inactiveDateCol + 1).clearContent();
          
          // Update cache
          cache.remove(`user_${username}`);
          cache.remove('userDataCache');
        }
        
        // Set Creation_Date if it's empty (for new users)
        if (!data[i][creationDateCol] && creationDateCol !== -1) {
          sheet.getRange(row, creationDateCol + 1).setValue(new Date());
        }
        
        // Spreadsheet flush to ensure changes are saved
        SpreadsheetApp.flush();
        
        return true;
      }
    }
    
    return false; // User not found
    
  } catch (e) {
    console.error('Error updating user status:', e);
    throw e;
  }
}

/**
 * Real-time status check - bypasses cache completely
 * Use this for critical status checks
 * @param {string} username - The username to check
 * @returns {Object} - User status information
 */
function checkUserStatusRealTime(username) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) throw new Error('Users sheet not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const usernameCol = headers.indexOf('Username');
    const statusCol = headers.indexOf('Status');
    const inactiveDateCol = headers.indexOf('Inactive_Date');
    const creationDateCol = headers.indexOf('Creation_Date');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameCol] === username) {
        return {
          username: data[i][usernameCol],
          status: data[i][statusCol] || 'Active',
          inactiveDate: data[i][inactiveDateCol],
          creationDate: data[i][creationDateCol]
        };
      }
    }
    
    return null; // User not found
  } catch (e) {
    console.error('Error in real-time status check:', e);
    throw e;
  }
}

/**
 * Gets user creation date
 * @param {string} username - The username
 * @returns {Date|null} - Creation date or null if not found
 */
function getUserCreationDate(username) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) throw new Error('Users sheet not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const usernameCol = headers.indexOf('Username');
    const creationDateCol = headers.indexOf('Creation_Date');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameCol] === username) {
        return data[i][creationDateCol] || null;
      }
    }
    
    return null;
  } catch (e) {
    console.error('Error getting user creation date:', e);
    throw e;
  }
}

/**
 * Gets list of inactive users
 * @returns {Array} - Array of inactive users with their details
 */
function getInactiveUsers() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) throw new Error('Users sheet not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const usernameCol = headers.indexOf('Username');
    const statusCol = headers.indexOf('Status');
    const inactiveDateCol = headers.indexOf('Inactive_Date');
    const creationDateCol = headers.indexOf('Creation_Date');
    const teacherNameCol = headers.indexOf('Teacher_Name');
    
    const inactiveUsers = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][statusCol] === 'Inactive') {
        inactiveUsers.push({
          username: data[i][usernameCol],
          teacherName: data[i][teacherNameCol],
          status: data[i][statusCol],
          creationDate: data[i][creationDateCol],
          inactiveDate: data[i][inactiveDateCol]
        });
      }
    }
    
    return inactiveUsers;
  } catch (e) {
    console.error('Error getting inactive users:', e);
    throw e;
  }
}


function login(username, password) {
  try {
    const cache = CacheService.getScriptCache();
    let userDataCache = cache.get('userDataCache');
    
    // Always check the sheet directly for critical status information
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) throw new Error('Users sheet not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const usernameCol = headers.indexOf('Username');
    const passwordCol = headers.indexOf('Password');
    const typeCol = headers.indexOf('UserType');
    const teacherNameCol = headers.indexOf('Teacher_Name');
    const statusCol = headers.indexOf('Status');
    
    // Find the user in the sheet (DIRECT READ - NO CACHE)
    let userFound = false;
    let userRow = null;
    let userStatus = null;
    let userData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameCol] === username) {
        userFound = true;
        userRow = i;
        userStatus = data[i][statusCol] || 'Active';
        
        // FIRST: Check if user is active before checking password
if (userStatus === 'Inactive') {
  // Force cache update
  cache.remove('userDataCache');
  initializeUserCache();
  
  // Throw a clear error that will be caught and displayed properly
  throw new Error('User account is inactive. Please contact administrator.');
}
        
        // SECOND: Check password
        if (data[i][passwordCol] === password) {
          userData = {
            password: data[i][passwordCol],
            type: data[i][typeCol] || 'teacher',
            teacherName: data[i][teacherNameCol] || data[i][usernameCol],
            status: userStatus
          };
        }
        break;
      }
    }
    
    if (!userFound) {
      return null; // User not found
    }
    
    if (!userData) {
      return null; // Wrong password
    }
    
    // Update cache with latest data
    if (userDataCache) {
      const users = JSON.parse(userDataCache);
      users[username] = userData;
      cache.put('userDataCache', JSON.stringify(users), 300);
    } else {
      // If cache is empty, initialize it
      initializeUserCache();
    }
    
    // Also cache individual user data for quick access
    cache.put(`user_${username}`, JSON.stringify(userData), 300);
    
  const recId = String(password).split('-')[0].trim();
const recName = getRecName(recId);

return {
  username: username,
  type: userData.type,
  teacherName: userData.teacherName,
  recId: recId,
  recName: recName
};
    
  } catch (e) {
    console.error('Login error:', e);
    throw e;
  }
}

function onOpen() {
  // Ensure Users sheet has the required columns
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (sheet) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Add missing columns if needed
    const requiredColumns = ['Creation_Date', 'Status', 'Inactive_Date'];
    requiredColumns.forEach((column, index) => {
      if (!headers.includes(column)) {
        const columnIndex = sheet.getLastColumn() + 1;
        sheet.getRange(1, columnIndex).setValue(column);
        
        // If adding Status column, set default value to 'Active' for existing users
        if (column === 'Status') {
          const lastRow = sheet.getLastRow();
          if (lastRow > 1) {
            const range = sheet.getRange(2, columnIndex, lastRow - 1, 1);
            range.setValue('Active');
          }
        }
        
        // If adding Creation_Date column, set current date for existing users
        if (column === 'Creation_Date') {
          const lastRow = sheet.getLastRow();
          if (lastRow > 1) {
            const range = sheet.getRange(2, columnIndex, lastRow - 1, 1);
            range.setValue(new Date());
          }
        }
      }
    });
  }
  
  initializeUserCache();
}

/**
 * Validates if a user can login based on their status
 * @param {string} username - The username to validate
 * @returns {boolean} - True if user can login (active), false otherwise
 */
function validateUserLogin(username) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) throw new Error('Users sheet not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const usernameCol = headers.indexOf('Username');
    const statusCol = headers.indexOf('Status');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameCol] === username) {
        const status = data[i][statusCol] || 'Active';
        return status === 'Active';
      }
    }
    
    return false; // User not found
  } catch (e) {
    console.error('Error validating user login:', e);
    throw e;
  }
}

function getUserSchedule(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) throw new Error('Users sheet not found');

  const data = userSheet.getDataRange().getValues();
  const headers = data[0];
  const usernameCol = headers.indexOf('Username');
  const daysCol = headers.indexOf('AllowedDays');
  const startTimeCol = headers.indexOf('StartTime');
  const endTimeCol = headers.indexOf('EndTime');

  if (usernameCol === -1 || daysCol === -1 || startTimeCol === -1 || endTimeCol === -1) {
    throw new Error('Required columns not found in Users sheet');
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][usernameCol] === username) {
      const allowedDays = data[i][daysCol] ? data[i][daysCol].split(',').map(day => day.trim()) : [];
      const startTime = data[i][startTimeCol] || '';
      const endTime = data[i][endTimeCol] || '';
      return { allowedDays, startTime, endTime };
    }
  }
  throw new Error('No schedule assigned for this user');
}

function isWithinSchedule(username) {
  const schedule = getUserSchedule(username);
  const now = new Date();
  const currentDay = now.toLocaleString('en-US', { weekday: 'long', timeZone: 'Asia/Karachi' });
  const currentTime = now.toLocaleString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false, timeZone: 'Asia/Karachi' });

  if (!schedule.allowedDays.includes(currentDay)) {
    throw new Error(`Attendance marking not allowed on ${currentDay}. Allowed days: ${schedule.allowedDays.join(', ')}`);
  }

  if (!schedule.startTime || !schedule.endTime) {
    throw new Error('Start or end time not defined for this user');
  }

  const [startHour, startMinute] = schedule.startTime.split(':').map(Number);
  const [endHour, endMinute] = schedule.endTime.split(':').map(Number);
  const [currentHour, currentMinute] = currentTime.split(':').map(Number);

  const startTimeInMinutes = startHour * 60 + startMinute;
  const endTimeInMinutes = endHour * 60 + endMinute;
  const currentTimeInMinutes = currentHour * 60 + currentMinute;

  if (currentTimeInMinutes < startTimeInMinutes || currentTimeInMinutes > endTimeInMinutes) {
    throw new Error(`Attendance marking only allowed between ${schedule.startTime} and ${schedule.endTime}`);
  }

  return true;
}



function getSummaryData(username) {
  try {
    const assignedSections = getAssignedClassSections(username);
    Logger.log('Assigned Sections for ' + username + ': ' + assignedSections);
    if (!assignedSections || assignedSections.length === 0) {
      Logger.log('No assigned sections for user: ' + username);
      return [];
    }

    const summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Summary');
    if (!summarySheet) throw new Error('Summary sheet not found');
    
    const data = summarySheet.getDataRange().getValues();
    Logger.log('Summary sheet data: ' + JSON.stringify(data));
    if (data.length <= 1) {
      Logger.log('Summary sheet is empty or contains only headers');
      return [];
    }

    const headers = data[0];
    const sectionCol = headers.indexOf('Class_Section');
    const presentCol = headers.indexOf('Present');
    const absentCol = headers.indexOf('Absent');
    const leavesCol = headers.indexOf('Leaves');
    const hoursCol = headers.indexOf('Hours');
    const categoryCol = headers.indexOf('Category');
    
    if (sectionCol === -1 || presentCol === -1 || absentCol === -1 || leavesCol === -1) {
      throw new Error('Required columns (Class_Section, Present, Absent, Leaves) not found');
    }

    const filteredData = data.filter(row => assignedSections.includes(row[sectionCol]));
    Logger.log('Filtered summary data: ' + JSON.stringify(filteredData));
    if (filteredData.length === 0) {
      Logger.log('No data found for assigned sections: ' + assignedSections);
    }

    return filteredData.map(row => ({
      classSection: row[sectionCol] || '',
      hours: hoursCol !== -1 ? (row[hoursCol] || '') : '',
      category: categoryCol !== -1 ? (row[categoryCol] || '') : '',
      present: row[presentCol] !== undefined && row[presentCol] !== '' && !isNaN(row[presentCol]) ? Number(row[presentCol]) : 0,
      absent: row[absentCol] !== undefined && row[absentCol] !== '' && !isNaN(row[absentCol]) ? Number(row[absentCol]) : 0,
      leaves: row[leavesCol] !== undefined && row[leavesCol] !== '' && !isNaN(row[leavesCol]) ? Number(row[leavesCol]) : 0,
      total: (row[presentCol] !== undefined && row[presentCol] !== '' && !isNaN(row[presentCol]) ? Number(row[presentCol]) : 0) +
             (row[absentCol] !== undefined && row[absentCol] !== '' && !isNaN(row[absentCol]) ? Number(row[absentCol]) : 0) +
             (row[leavesCol] !== undefined && row[leavesCol] !== '' && !isNaN(row[leavesCol]) ? Number(row[leavesCol]) : 0)
    }));
  } catch (e) {
    Logger.log('Error in getSummaryData: ' + e.message);
    throw e;
  }
}

function getAssignedClassSections(username) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teacher_Assignments');
    if (!sheet) throw new Error('Teacher_Assignments sheet not found');

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        const sections = data[i][1] ? data[i][1].split(',').filter(s => s.trim()) : [];
        const validSections = getClassSections();
        return sections.filter(section => validSections.includes(section));
      }
    }
    return [];
  } catch (e) {
    Logger.log('Error in getAssignedClassSections: ' + e.message);
    throw e;
  }
}

function verifyPreviousPassword(username, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Previous_Date_Permissions');
    if (!sheet) throw new Error('Previous_Date_Permissions sheet not found');

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const sheetUsername = String(data[i][0]).trim();
      const sheetPassword = String(data[i][2]).trim();
      if (sheetUsername === username && sheetPassword === password) {
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log(`Error in verifyPreviousPassword: ${e.message}`);
    throw e;
  }
}

function checkPreviousDatePermission(username, date) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Previous_Date_Permissions');
    if (!sheet) throw new Error('Previous_Date_Permissions sheet not found');

    const data = sheet.getDataRange().getValues();
    const targetDate = new Date(date);
    const dateString = targetDate.toLocaleDateString('en-CA'); // YYYY-MM-DD

    for (let i = 1; i < data.length; i++) {
      let sheetDate;
      try {
        sheetDate = new Date(data[i][1]);
        if (isNaN(sheetDate.getTime())) {
          Logger.log(`Invalid date in Previous_Date_Permissions row ${i + 2}: ${data[i][1]}`);
          continue;
        }
        const sheetDateString = sheetDate.toLocaleDateString('en-CA');
        if (data[i][0] === username && sheetDateString === dateString) {
          return true;
        }
      } catch (e) {
        Logger.log(`Error parsing date in row ${i + 2}: ${e.message}`);
      }
    }
    return false;
  } catch (e) {
    Logger.log(`Error in checkPreviousDatePermission: ${e.message}`);
    throw e;
  }
}

let classSectionsCache = null;

function getClassSections() {
  if (classSectionsCache) return classSectionsCache;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_Sections');
  if (!sheet) throw new Error('Class_Sections sheet not found');
  const data = sheet.getDataRange().getValues();
  const classSections = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      classSections.push(data[i][0]);
    }
  }
  classSectionsCache = classSections;
  return classSections;
}

let allowedDatesCache = null;

function getAllowedPreviousDates(username) {
  if (allowedDatesCache && allowedDatesCache.username === username) return allowedDatesCache.dates;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Previous_Date_Permissions');
  if (!sheet) throw new Error('Previous_Date_Permissions sheet not found');
  const data = sheet.getDataRange().getValues();
  const allowedDates = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1]) {
      try {
        const date = new Date(data[i][1]);
        if (!isNaN(date.getTime())) {
          allowedDates.push(date.toLocaleDateString('en-CA')); // YYYY-MM-DD
        } else {
          Logger.log(`Invalid date in Previous_Date_Permissions row ${i + 2}: ${data[i][1]}`);
        }
      } catch (e) {
        Logger.log(`Error parsing date in row ${i + 2}: ${e.message}`);
      }
    }
  }
  allowedDatesCache = { username, dates: allowedDates };
  return allowedDates;
}

function parseSheetDate(dateValue) {
  if (!dateValue) return null;
  
  // If it's already a Date object and valid
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
    return dateValue;
  }
  
  // If it's a string, try to parse it
  if (typeof dateValue === 'string') {
    // Try different date formats
    const parsed = new Date(dateValue);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
    
    // Try common spreadsheet date formats
    const formats = [
      'YYYY-MM-DD',
      'MM/DD/YYYY', 
      'DD/MM/YYYY',
      'MMM DD, YYYY'
    ];
    
    for (const format of formats) {
      try {
        const formatted = Utilities.formatDate(new Date(dateValue), 'Asia/Karachi', 'yyyy-MM-dd');
        const finalDate = new Date(formatted);
        if (!isNaN(finalDate.getTime())) {
          return finalDate;
        }
      } catch (e) {
        continue;
      }
    }
  }
  
  Logger.log(`Warning: Could not parse date value: ${dateValue}`);
  return null;
}

function getStudentList(classSection, date) {
  try {
    Logger.log(`getStudentList called with classSection: ${classSection}, date: ${date}`);
    if (!classSection || !classSection.includes('-')) {
      throw new Error(`Invalid classSection format: ${classSection}. Expected format: Class-Section`);
    }
    const [className, section] = classSection.split('-');
    if (!className || !section) {
      throw new Error('Class or section cannot be empty');
    }

    const studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Data');
    if (!studentSheet) throw new Error('Student_Data sheet not found');
    const studentHeaders = studentSheet.getRange(1, 1, 1, studentSheet.getLastColumn()).getValues()[0];
    
    // Check for required columns including Date_of_Joining
    const requiredColumns = ['Std_ID', 'Student_Name', 'Status', 'Date_of_Joining'];
    requiredColumns.forEach(col => {
      if (!studentHeaders.includes(col)) {
        throw new Error(`Required column '${col}' missing in Student_Data sheet`);
      }
    });

    const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Attendance');
    if (!attendanceSheet) throw new Error('Student_Attendance sheet not found');

    const studentData = studentSheet.getDataRange().getValues();
    const studentsMap = new Map();
    
    // Get column indices
    const stdIdCol = studentHeaders.indexOf('Std_ID');
    const barcodeIdCol = studentHeaders.indexOf('Barcode_ID');
    const nameCol = studentHeaders.indexOf('Student_Name');
    const classCol = studentHeaders.indexOf('Student_Class');
    const sectionCol = studentHeaders.indexOf('Student_Section');
    const statusCol = studentHeaders.indexOf('Status');
    const dojCol = studentHeaders.indexOf('Date_of_Joining');

    for (let i = 1; i < studentData.length; i++) {
      const status = studentData[i][statusCol] ? studentData[i][statusCol].toString().toLowerCase() : '';
      const dateOfJoining = parseSheetDate(studentData[i][dojCol]);
      const targetDate = new Date(date);
      
      // Check if student should be included based on joining date
      const shouldIncludeByDate = !dateOfJoining || dateOfJoining <= targetDate;
      
      if (studentData[i][classCol] === className && 
          studentData[i][sectionCol] === section && 
          status === 'active' &&
          shouldIncludeByDate) {
        
        studentsMap.set(String(studentData[i][stdIdCol]), {
          Std_ID: String(studentData[i][stdIdCol]),
          Barcode_ID: studentData[i][barcodeIdCol],
          Student_Name: studentData[i][nameCol],
          Student_Class: studentData[i][classCol],
          Student_Section: studentData[i][sectionCol],
          Date_of_Joining: dateOfJoining ? dateOfJoining.toLocaleDateString('en-CA') : null,
          attendanceStatus: 'Not marked'
        });
      }
    }

    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!dateRegex.test(date)) {
      throw new Error(`Invalid date format: ${date}. Please use YYYY-MM-DD (e.g., 2025-10-03).`);
    }
    
    const targetDate = new Date(date);
    const dateString = targetDate.toLocaleDateString('en-CA');

    const attendanceData = attendanceSheet.getDataRange().getValues();
    const headers = attendanceData[0];
    const attendanceHeaderMap = new Map(headers.map((h, i) => [h, i]));
    if (!attendanceHeaderMap.has('Std_ID') || !attendanceHeaderMap.has('Date') || !attendanceHeaderMap.has('Status')) {
      throw new Error('Required headers missing in Student_Attendance sheet');
    }

        for (let j = 1; j < attendanceData.length; j++) {
      let rowDate = parseSheetDate(attendanceData[j][attendanceHeaderMap.get('Date')]);
      if (!rowDate) continue;
      const rowDateString = rowDate.toLocaleDateString('en-CA');
      if (rowDateString === dateString) {
        const stdId = String(attendanceData[j][attendanceHeaderMap.get('Std_ID')]);
        if (studentsMap.has(stdId)) {
          // ✅ CRITICAL: Only get Normal attendance, NOT Episodic
          const attendanceType = attendanceData[j][attendanceHeaderMap.get('Attendance_Type')];
          // Skip if this is an episodic record
          if (attendanceType !== 'Episodic') {
            studentsMap.get(stdId).attendanceStatus = attendanceData[j][attendanceHeaderMap.get('Status')] || 'Not marked';
          }
        }
      }
    }

    const students = Array.from(studentsMap.values());
    Logger.log(`Returning ${students.length} students with attendance statuses after Date_of_Joining filter`);
    
    // Log students who were filtered out due to future joining dates
    const totalStudents = studentData.length - 1; // Exclude header
    const filteredStudents = totalStudents - students.length;
    if (filteredStudents > 0) {
      Logger.log(`${filteredStudents} students filtered out due to future Date_of_Joining or other conditions`);
    }
    
    return students;
  } catch (e) {
    Logger.log(`Error in getStudentList: ${e.message}`);
    throw e;
  }
}



/**
 * Get student list with episodic attendance status for a specific event
 * @param {string} classSection - Class section (e.g., "10-A")
 * @param {string} date - Date in YYYY-MM-DD format
 * @param {string} eventName - Name of the episodic event
 * @returns {Array} - Students with attendanceStatus set to 'Present' or 'Not Applicable'
 */
function getEpisodicStudentList(classSection, date, eventName) {
  try {
    Logger.log(`getEpisodicStudentList called with classSection: ${classSection}, date: ${date}, eventName: ${eventName}`);
    
    // First, get all students for this class/date (without attendance status)
    const studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Data');
    if (!studentSheet) throw new Error('Student_Data sheet not found');
    
    const [className, section] = classSection.split('-');
    if (!className || !section) {
      throw new Error(`Invalid classSection format: ${classSection}. Expected format: Class-Section`);
    }
    
    const studentHeaders = studentSheet.getRange(1, 1, 1, studentSheet.getLastColumn()).getValues()[0];
    const stdIdCol = studentHeaders.indexOf('Std_ID');
    const barcodeIdCol = studentHeaders.indexOf('Barcode_ID');
    const nameCol = studentHeaders.indexOf('Student_Name');
    const classCol = studentHeaders.indexOf('Student_Class');
    const sectionCol = studentHeaders.indexOf('Student_Section');
    const statusCol = studentHeaders.indexOf('Status');
    const dojCol = studentHeaders.indexOf('Date_of_Joining');
    
    const studentData = studentSheet.getDataRange().getValues();
    const studentsMap = new Map();
    const targetDate = new Date(date);
    
    // Get all ACTIVE students for this class
    for (let i = 1; i < studentData.length; i++) {
      const studentStatus = studentData[i][statusCol] ? studentData[i][statusCol].toString().toLowerCase() : '';
      const dateOfJoining = parseSheetDate(studentData[i][dojCol]);
      const shouldIncludeByDate = !dateOfJoining || dateOfJoining <= targetDate;
      
      if (studentData[i][classCol] === className && 
          studentData[i][sectionCol] === section && 
          studentStatus === 'active' &&
          shouldIncludeByDate) {
        
        studentsMap.set(String(studentData[i][stdIdCol]), {
          Std_ID: String(studentData[i][stdIdCol]),
          Barcode_ID: studentData[i][barcodeIdCol],
          Student_Name: studentData[i][nameCol],
          Student_Class: studentData[i][classCol],
          Student_Section: studentData[i][sectionCol],
          Date_of_Joining: dateOfJoining ? dateOfJoining.toLocaleDateString('en-CA') : null,
          attendanceStatus: 'Not Applicable'  // Default for episodic
        });
      }
    }
    
    // If no event name, return all students with default status
    if (!eventName || eventName.trim() === '') {
      Logger.log('No event name provided, returning students with default status');
      return Array.from(studentsMap.values());
    }
    
    // Get EPISODIC attendance ONLY for this specific event
    const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Attendance');
    if (!attendanceSheet) throw new Error('Student_Attendance sheet not found');
    
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const headers = attendanceData[0];
    
    // Find column indices
    const stdIdColAtt = headers.indexOf('Std_ID');
    const dateColAtt = headers.indexOf('Date');
    const classColAtt = headers.indexOf('Class_Section');
    const typeColAtt = headers.indexOf('Attendance_Type');
    const eventColAtt = headers.indexOf('Event_Name');
    const statusColAtt = headers.indexOf('Status');
    
    const targetDateStr = new Date(date).toLocaleDateString('en-CA');
    
    // Create a map of studentId -> status ONLY for EPISODIC records
    const episodicStatusMap = new Map();
    
    for (let i = 1; i < attendanceData.length; i++) {
      const row = attendanceData[i];
      const rowDate = row[dateColAtt] ? new Date(row[dateColAtt]).toLocaleDateString('en-CA') : null;
      const rowType = row[typeColAtt];
      const rowEvent = row[eventColAtt];
      const rowClass = row[classColAtt];
      
      // ✅ CRITICAL: Only match EPISODIC type!
      if (rowDate === targetDateStr &&
          rowType === 'Episodic' &&           // ← ONLY Episodic records
          rowEvent === eventName &&
          rowClass === classSection) {
        const studentId = String(row[stdIdColAtt]);
        episodicStatusMap.set(studentId, row[statusColAtt]);
        Logger.log(`Found EPISODIC record: Student ${studentId} → ${row[statusColAtt]}`);
      }
    }
    
    // Update student status based ONLY on episodic records
    for (const student of studentsMap.values()) {
      const studentId = String(student.Std_ID);
      if (episodicStatusMap.has(studentId)) {
        student.attendanceStatus = episodicStatusMap.get(studentId);
        Logger.log(`Student ${studentId} has episodic status: ${student.attendanceStatus}`);
      } else {
        // No episodic record found - default to 'Not Applicable'
        student.attendanceStatus = 'Not Applicable';
      }
    }
    
    Logger.log(`Returning ${studentsMap.size} students with episodic-only statuses`);
    return Array.from(studentsMap.values());
    
  } catch (e) {
    Logger.log(`Error in getEpisodicStudentList: ${e.message}`);
    throw e;
  }
}

function generateUUID() {
  return Utilities.getUuid();
}

function markAttendance(attendanceData, class_section, date, username, isBarcode = false, recId = '') {
   const startTime = Date.now();

// ✅ ADD THIS LINE - SERVER-SIDE SCHEDULE CHECK
  isWithinSchedule(username); // Will throw error if outside allowed time

  Logger.log(`Starting markAttendance with date: ${date}, classSection: ${class_section}, isBarcode: ${isBarcode}, records: ${attendanceData.length}`);

  try {
    // ========================================================
    // TERM VALIDATION - Check if class has active term
    // Checks both Class_Terms.Is_Active AND Terms.Is_Active
    // ========================================================
    const term = getTermForClassSection(class_section, date);

    if (!term) {
      // First check Class_Terms
      const classTermsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_Terms');
      if (classTermsSheet) {
        const data = classTermsSheet.getDataRange().getValues();
        const headers = data[0];
        const classCol = headers.indexOf('Class_Section');
        const activeCol = headers.indexOf('Is_Active');
        
        for (let i = 1; i < data.length; i++) {
          if (data[i][classCol] === class_section) {
            const isActiveInClassTerms = data[i][activeCol] === true || data[i][activeCol] === 'TRUE';
            
            if (!isActiveInClassTerms) {
              throw new Error(`Class ${class_section} is INACTIVE in Class_Terms sheet.`);
            }
            
            // If class is active in Class_Terms but term not found, check Terms sheet
            const termId = data[i][headers.indexOf('Term_ID')];
            const termDetails = getTermDetails(termId);
            
            if (termDetails && !termDetails.isActive) {
              throw new Error(`Term "${termDetails.termName}" is INACTIVE in Terms sheet. Contact administrator.`);
            }
            
            break;
          }
        }
      }
      
      throw new Error(`No active term found for ${class_section} on ${date}.`);
    }

    // Double-check term is active (should already be checked, but just in case)
    if (!term.isActive) {
      throw new Error(`Cannot mark attendance for ${class_section}. Term "${term.termName}" is INACTIVE.`);
    }

    Logger.log(`Using ACTIVE term: ${term.termName} (${term.termId}) from ${term.startDate} to ${term.endDate}`);
    // ========================================================
    
    // -------------------------------------------------
    // NEW: Check that hours & category are defined for the day
    // -------------------------------------------------
const hoursCheck = checkDailyHours(date, class_section, 'normal');
console.log('Hours check result:', hoursCheck);  // Add this for debugging

    if (!hoursCheck.exists) {
      throw new Error('Hours and category must be set before marking attendance.');
    }

    // -------------------------------------------------
    // Permissions check - remains outside the lock (read-only, fast)
    // -------------------------------------------------
    const permissions = getUserPermissions(username);
    if (isBarcode && !permissions.canMarkBarcode) {
      throw new Error('User does not have permission to mark barcode attendance');
    }
    if (!isBarcode && !permissions.canMarkManual) {
      throw new Error('User does not have permission to mark manual attendance');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Attendance');
    if (!sheet) throw new Error('Student_Attendance sheet not found');

    // -------------------------------------------------
    // Validate date format
    // -------------------------------------------------
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!dateRegex.test(date)) {
      throw new Error(`Invalid date format: ${date}. Please use YYYY-MM-DD (e.g., 2025-10-03).`);
    }

    // -------------------------------------------------
    // --- Acquire Lock ---
    // -------------------------------------------------
    // -------------------------------------------------
    // Fetch headers ONCE before acquiring lock
    // -------------------------------------------------
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = new Map(headers.map((h, i) => [h, i]));

    // -------------------------------------------------
    // --- Acquire Lock ---
    // -------------------------------------------------
    const lock = LockService.getScriptLock();
    const maxAttempts = 5;
    let attempt = 0;
    const baseDelay = 1000;
    const maxDelay = 8000;

    while (attempt < maxAttempts) {
      try {
        if (lock.tryLock(10000)) { // 10-second timeout per attempt
          Logger.log(`Lock acquired on attempt ${attempt + 1} after ${Date.now() - startTime}ms`);

          // Required original columns (using headers fetched above)
          if (!headerMap.has('Attendance_ID') || !headerMap.has('Std_ID') ||
              !headerMap.has('Date') || !headerMap.has('Status') || !headerMap.has('Hours') || !headerMap.has('Category')) {
            throw new Error('Required headers missing in Student_Attendance sheet');
          }

          // ========================================================
          // NEW: Get Term column indices
          // ========================================================
          const termIdCol = headerMap.get('Term_ID');
          const termNameCol = headerMap.get('Term_Name');
          
          // Log warning if term columns missing but continue
          if (termIdCol === undefined || termNameCol === undefined) {
            Logger.log(`WARNING: Term_ID or Term_Name columns not found in Student_Attendance. Attendance will be marked without term info.`);
          } else {
            Logger.log(`Term_ID column: ${termIdCol}, Term_Name column: ${termNameCol}`);
          }
          // ========================================================

          // NEW columns – safe-guard if they are missing
       const hoursCol = headerMap.get('Hours');
        const categoryCol = headerMap.get('Category');
        const recIdCol = headerMap.get('REC_ID');

// -------------------------------------------------
// Build map of existing rows for the target date + class
// ✅ FIXED: Only consider Normal attendance records, ignore Episodic
// -------------------------------------------------
const existingRecords = sheet.getDataRange().getValues();
const existingRecordsByStudentDate = new Map();
const targetDateString = new Date(date).toLocaleDateString('en-CA');

for (let i = 1; i < existingRecords.length; i++) { // skip header row
  const rowData = existingRecords[i];
  const stdId = String(rowData[headerMap.get('Std_ID')]);
  
  // ✅ CRITICAL FIX: Get Attendance_Type and skip Episodic records
  const attendanceTypeCol = headerMap.get('Attendance_Type');
  const rowAttendanceType = attendanceTypeCol !== undefined ? rowData[attendanceTypeCol] : '';
  
  // Skip Episodic records - they belong to a different attendance type
  if (rowAttendanceType === 'Episodic') {
    continue;
  }

  let rowDateStr;
  try {
    rowDateStr = new Date(rowData[headerMap.get('Date')]).toLocaleDateString('en-CA');
  } catch (e) {
    Logger.log(`Warning: Invalid date in row ${i + 1}: ${rowData[headerMap.get('Date')]}. Skipping.`);
    continue;
  }

  const rowClassSection = rowData[headerMap.get('Class_Section')];

  if (rowDateStr === targetDateString && rowClassSection === class_section) {
    existingRecordsByStudentDate.set(stdId, {
      rowIndex: i + 1,
      data: rowData,
      currentStatus: rowData[headerMap.get('Status')]
    });
  }
}
Logger.log(`Found ${existingRecordsByStudentDate.size} existing NORMAL records for ${class_section} on ${date}`);

          // -------------------------------------------------
          // -------------------------------------------------
          // Prepare bulk updates / new rows
          // -------------------------------------------------
          const updatesToApply = [];   // { rowIndex, rowValues }
          const newRowsToAppend = [];

          for (const record of attendanceData) {
            const studentKey = String(record.Std_ID);
            const existing = existingRecordsByStudentDate.get(studentKey);

            if (existing) {
              // ---- UPDATE existing row ----
              if (existing.currentStatus !== record.Status) {
                const updatedRow = [...existing.data]; // shallow copy
                updatedRow[headerMap.get('Status')]      = record.Status;
                updatedRow[headerMap.get('Timestamp')]   = new Date().toLocaleString();
                updatedRow[headerMap.get('Teacher_ID')]  = username;
                updatedRow[headerMap.get('Class_Section')] = class_section;


    // Add this line in the update section
    const attendanceTypeCol = headerMap.get('Attendance_Type');
    if (attendanceTypeCol !== undefined) {
      updatedRow[attendanceTypeCol] = 'Normal';
    }

                                // NEW: write Hours & Category
                updatedRow[hoursCol] = hoursCheck.hours;
                updatedRow[categoryCol] = hoursCheck.category;
                if (recIdCol !== undefined) updatedRow[recIdCol] = recId;
                
                // ========================================================
                // NEW: Add Term information if columns exist
                // ========================================================
                if (termIdCol !== undefined && termNameCol !== undefined) {
                  updatedRow[termIdCol] = term.termId;
                  updatedRow[termNameCol] = term.termName;
                  Logger.log(`Added term info to UPDATE: ${term.termId} - ${term.termName}`);
                }
                // ========================================================
                


                updatesToApply.push({
                  rowIndex: existing.rowIndex,
                  rowValues: updatedRow
                });
                Logger.log(`Prepared UPDATE Std_ID ${record.Std_ID} → ${record.Status} (row ${existing.rowIndex})`);
              } else {
                Logger.log(`Skipping Std_ID ${record.Std_ID} – status already ${record.Status}`);
              }
            } else {
              // ---- CREATE new row ----
              const newRow = new Array(headers.length).fill('');
              newRow[headerMap.get('Attendance_ID')]   = generateUUID();
              newRow[headerMap.get('Date')]           = record.Date;
              newRow[headerMap.get('Std_ID')]         = String(record.Std_ID);
              newRow[headerMap.get('Barcode_ID')]     = record.Barcode_ID || '';
              newRow[headerMap.get('Student_Name')]   = record.Student_Name;
              newRow[headerMap.get('Student_Class')]  = record.Student_Class;
              newRow[headerMap.get('Student_Section')]= record.Student_Section;
              newRow[headerMap.get('Status')]         = record.Status;
              newRow[headerMap.get('Timestamp')]      = new Date().toLocaleString();
              newRow[headerMap.get('Teacher_ID')]     = username;
              newRow[headerMap.get('Class_Section')]  = class_section;

              // Add this line - set Attendance_Type to 'Normal'
const attendanceTypeCol = headerMap.get('Attendance_Type');
if (attendanceTypeCol !== undefined) {
  newRow[attendanceTypeCol] = 'Normal';
}

              // NEW: write Hours & Category
              newRow[hoursCol] = hoursCheck.hours;
              newRow[categoryCol] = hoursCheck.category;
              if (recIdCol !== undefined) newRow[recIdCol] = recId;

              // ========================================================
              // NEW: Add Term information if columns exist
              // ========================================================
              if (termIdCol !== undefined && termNameCol !== undefined) {
                newRow[termIdCol] = term.termId;
                newRow[termNameCol] = term.termName;
                Logger.log(`Added term info to NEW row: ${term.termId} - ${term.termName}`);
              }
              // ========================================================


              newRowsToAppend.push(newRow);
              Logger.log(`Prepared NEW row for Std_ID ${record.Std_ID} → ${record.Status}`);
            }
          }

          // -------------------------------------------------
          // Apply updates (individual setValues – still fast for <100 rows)
          // -------------------------------------------------
          if (updatesToApply.length > 0) {
            for (const upd of updatesToApply) {
              sheet.getRange(upd.rowIndex, 1, 1, upd.rowValues.length)
                   .setValues([upd.rowValues]);
            }
            Logger.log(`Applied ${updatesToApply.length} row update(s).`);
          }

          // -------------------------------------------------
          // Append new rows in one batch
          // -------------------------------------------------
          if (newRowsToAppend.length > 0) {
            const startRow = sheet.getLastRow() + 1;
            sheet.getRange(startRow, 1, newRowsToAppend.length, headers.length)
                 .setValues(newRowsToAppend);
            Logger.log(`Appended ${newRowsToAppend.length} new row(s).`);
          }

          // One final flush to guarantee all writes are sent
          SpreadsheetApp.flush();

          Logger.log(`Attendance processing complete in ${Date.now() - startTime}ms`);
          lock.releaseLock();
          Logger.log(`Lock released at ${Date.now() - startTime}ms`);
          return true; // SUCCESS
        } else {
          // ---- lock not acquired → retry with exponential back-off ----
          attempt++;
          const delay = Math.min(baseDelay * Math.pow(2, attempt - 1), maxDelay);
          Logger.log(`Lock failed (attempt ${attempt}) after ${Date.now() - startTime}ms – retry in ${delay}ms`);
          Utilities.sleep(delay);
        }
      } catch (e) {
        // ---- error inside a lock attempt ----
        Logger.log(`Error on attempt ${attempt + 1} at ${Date.now() - startTime}ms: ${e.message}`);
        if (lock.hasLock()) {
          lock.releaseLock();
          Logger.log(`Lock released due to error at ${Date.now() - startTime}ms`);
        }
        throw e; // re-throw so the outer catch can log it once
      }
    }

    // -------------------------------------------------
    // All retry attempts exhausted
    // -------------------------------------------------
    const errMsg = `Failed to acquire lock after ${maxAttempts} attempts in ${Date.now() - startTime}ms. System busy – try again later.`;
    Logger.log(errMsg);
    throw new Error(errMsg);

  } catch (e) {
    Logger.log(`Final error in markAttendance at ${Date.now() - startTime}ms: ${e.message}`);
    throw e;
  }
}

// =============== EPISODIC ATTENDANCE FUNCTIONS ===============
// ADD ALL OF THIS CODE AT THE END OF YOUR code.gs FILE

/**
 * Get categories filtered by attendance type from Hour_Categories sheet
 */
function getCategoriesByType(type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hour_Categories');
  if (!sheet) throw new Error('Hour_Categories sheet not found');
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const nameCol = headers.indexOf('Category_Name');
  const activeCol = headers.indexOf('Is_Active');
  const typeCol = headers.indexOf('Attendance_Type');
  
  const categories = [];
  
  for (let i = 1; i < data.length; i++) {
    const isActive = data[i][activeCol] === true || data[i][activeCol] === 'TRUE';
    if (!isActive) continue;
    
    let allowedType = 'both';
    if (typeCol !== -1 && data[i][typeCol]) {
      allowedType = String(data[i][typeCol]).toLowerCase();
    }
    
    if (!type || allowedType === 'both' || allowedType === type) {
      categories.push(String(data[i][nameCol]));
    }
  }
  
  return categories;
}

/**
 * Get episodic categories
 */
function getEpisodicCategories() {
  return getCategoriesByType('episodic');
}

/**
 * Get hours configuration filtered by attendance type from Hours_Setup sheet
 */
function getHoursByType(type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hours_Setup');
  if (!sheet) throw new Error('Hours_Setup sheet not found');
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const hoursCol = headers.indexOf('Hours');
  const typeCol = headers.indexOf('Attendance_Type');
  const maxCustomCol = headers.indexOf('Max_Custom');
  
  const presetHours = [];
  let allowCustom = false;
  let maxCustom = 8;
  
  for (let i = 1; i < data.length; i++) {
    const hoursValue = String(data[i][hoursCol]).toLowerCase();
    let allowedType = 'both';
    
    if (typeCol !== -1 && data[i][typeCol]) {
      allowedType = String(data[i][typeCol]).toLowerCase();
    }
    
    const matchesType = !type || allowedType === 'both' || allowedType === type;
    
    if (hoursValue === 'custom') {
      if (matchesType) {
        allowCustom = true;
        if (maxCustomCol !== -1 && data[i][maxCustomCol]) {
          maxCustom = Number(data[i][maxCustomCol]);
        }
      }
    } else if (matchesType) {
      presetHours.push(hoursValue);
    }
  }
  
  return {
    presetHours: presetHours,
    allowCustom: allowCustom,
    maxCustom: maxCustom
  };
}

/**
 * Get episodic hours config
 */
function getEpisodicHoursConfig() {
  return getHoursByType('episodic');
}

/**
 * Mark episodic attendance (Field Trips, Camps, Visits)
 * Only marks Present for selected students
 */
/**
 * Mark episodic attendance for ALL students (Present or Not Applicable)
 * Every student gets a record
 */
/**
 * Mark episodic attendance for ALL students (Present or Not Applicable)
 * Every student gets a record - updates existing or creates new
 */
function markEpisodicAttendance(studentStatuses, classSection, date, username, eventName, category, hours, isCustomHours, recId = '') {
  const startTime = Date.now();
  
  // studentStatuses = [{ studentId: "123", status: "Present" }, { studentId: "456", status: "Not Applicable" }]
  Logger.log(`Episodic: ${eventName} | ${classSection} | ${date} | Hours: ${hours}`);
  Logger.log(`Processing ${studentStatuses.length} students`);
  
  try {
    // ========================================================
    // 1. DATE FORMAT VALIDATION
    // ========================================================
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!dateRegex.test(date)) {
      throw new Error(`Invalid date format: ${date}. Please use YYYY-MM-DD (e.g., 2025-10-03).`);
    }
    
    // ========================================================
    // 2. TERM VALIDATION - Check if class has active term
    // Checks both Class_Terms.Is_Active AND Terms.Is_Active
    // ========================================================
    const term = getTermForClassSection(classSection, date);
    if (!term) {
      // First check Class_Terms for more detailed error
      const classTermsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_Terms');
      if (classTermsSheet) {
        const data = classTermsSheet.getDataRange().getValues();
        const headers = data[0];
        const classCol = headers.indexOf('Class_Section');
        const activeCol = headers.indexOf('Is_Active');
        
        for (let i = 1; i < data.length; i++) {
          if (data[i][classCol] === classSection) {
            const isActiveInClassTerms = data[i][activeCol] === true || data[i][activeCol] === 'TRUE';
            
            if (!isActiveInClassTerms) {
              throw new Error(`Class ${classSection} is INACTIVE in Class_Terms sheet. Cannot mark episodic attendance.`);
            }
            
            // If class is active in Class_Terms but term not found, check Terms sheet
            const termId = data[i][headers.indexOf('Term_ID')];
            const termDetails = getTermDetails(termId);
            
            if (termDetails && !termDetails.isActive) {
              throw new Error(`Term "${termDetails.termName}" is INACTIVE in Terms sheet. Contact administrator.`);
            }
            
            break;
          }
        }
      }
      
      throw new Error(`No active term found for ${classSection} on ${date}. Cannot mark episodic attendance.`);
    }

    // Double-check term is active
    if (!term.isActive) {
      throw new Error(`Cannot mark episodic attendance for ${classSection}. Term "${term.termName}" is INACTIVE.`);
    }

    Logger.log(`Using ACTIVE term: ${term.termName} (${term.termId}) from ${term.startDate} to ${term.endDate}`);
    
    // ========================================================
    // 3. SCHEDULE CHECK - Server-side schedule validation
    // ========================================================
    isWithinSchedule(username);
    Logger.log(`Schedule check passed for user: ${username}`);
    
    // ========================================================
    // 4. HOURS VALIDATION - Check that hours & category are defined for the day
    // CRITICAL: Episodic attendance also needs hours validation
    // ========================================================
const hoursCheck = checkDailyHours(date, classSection, 'episodic');

let finalHours = hours;
let finalCategory = category;

if (hoursCheck.exists) {
  finalHours = hoursCheck.hours;
  finalCategory = hoursCheck.category;
  Logger.log(`Using pre-configured episodic hours: ${finalHours} (${finalCategory})`);
} else {
  Logger.log(`No pre-configured hours. Using user selection: ${finalHours} (${finalCategory})`);
}
    
    // ========================================================
    // 5. PERMISSION CHECK - Server-side permission validation
    // CRITICAL: Verify user has episodic permission
    // ========================================================
    const permissions = getUserPermissions(username);
    if (!permissions.canMarkEpisodic) {
      throw new Error('User does not have permission to mark episodic attendance');
    }
    Logger.log(`Permission check passed for user: ${username}`);
    
    // ========================================================
    // 6. Get the attendance sheet and headers
    // ========================================================
    const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Attendance');
    if (!attendanceSheet) throw new Error('Student_Attendance sheet not found');
    
    const headers = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0];
    const headerMap = new Map(headers.map((h, i) => [h, i]));
    
    // Get column indices
    const typeCol = headerMap.get('Attendance_Type');
    const eventCol = headerMap.get('Event_Name');
    const termIdCol = headerMap.get('Term_ID');
    const termNameCol = headerMap.get('Term_Name');
    const statusCol = headerMap.get('Status');
    const timestampCol = headerMap.get('Timestamp');
    const teacherCol = headerMap.get('Teacher_ID');
    const stdIdCol = headerMap.get('Std_ID');
    const dateCol = headerMap.get('Date');
    const classCol = headerMap.get('Class_Section');
    const hoursCol = headerMap.get('Hours');
    const categoryCol = headerMap.get('Category');
    const recIdCol = headerMap.get('REC_ID');
    const studentNameCol = headerMap.get('Student_Name');
    const studentClassCol = headerMap.get('Student_Class');
    const studentSectionCol = headerMap.get('Student_Section');
    const barcodeCol = headerMap.get('Barcode_ID');
    const attendanceIdCol = headerMap.get('Attendance_ID');
    
    // Validate required columns exist
    if (stdIdCol === undefined || dateCol === undefined || classCol === undefined || 
        statusCol === undefined || hoursCol === undefined || categoryCol === undefined) {
      throw new Error('Required columns missing in Student_Attendance sheet for episodic attendance');
    }
    
    const targetDate = new Date(date).toLocaleDateString('en-CA');
    
    // ========================================================
    // 7. Get all students for this class
    // ========================================================
    const allStudents = getStudentList(classSection, date);
    const studentMap = new Map();
    allStudents.forEach(s => {
      studentMap.set(String(s.Std_ID), s);
    });
    Logger.log(`Found ${studentMap.size} students for class ${classSection}`);
    
    // ========================================================
    // 8. Create a map of existing records for this date/class/event
    // ========================================================
    const data = attendanceSheet.getDataRange().getValues();
    const existingRecords = new Map(); // studentId -> { rowIndex, currentStatus, rowData }
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowClass = row[classCol];
      let rowDate;
      try {
        rowDate = new Date(row[dateCol]).toLocaleDateString('en-CA');
      } catch (e) {
        continue;
      }
      const rowType = row[typeCol];
      const rowEvent = row[eventCol];
      
      if (rowClass === classSection &&
          rowDate === targetDate &&
          rowType === 'Episodic' &&
          rowEvent === eventName) {
        const studentId = String(row[stdIdCol]);
        existingRecords.set(studentId, {
          rowIndex: i + 1,
          currentStatus: row[statusCol],
          rowData: row
        });
        Logger.log(`Found existing episodic record for student ${studentId}: ${row[statusCol]}`);
      }
    }
    Logger.log(`Found ${existingRecords.size} existing episodic records for ${eventName}`);
    
    // ========================================================
    // 9. Prepare updates and inserts
    // ========================================================
    const updatesToApply = []; // { rowIndex, newStatus, rowData }
    const newRowsToAppend = [];
    
    // Process each student's status
    for (const item of studentStatuses) {
      const studentId = String(item.studentId);
      const newStatus = item.status;
      const student = studentMap.get(studentId);
      
      if (!student) {
        Logger.log(`Warning: Student ${studentId} not found in class ${classSection}`);
        continue;
      }
      
      const existing = existingRecords.get(studentId);
      const now = new Date();
      
      if (existing) {
        // Update existing record if status changed
        if (existing.currentStatus !== newStatus) {
          // Create updated row data based on existing row
          const updatedRow = [...existing.rowData];
          updatedRow[statusCol] = newStatus;
          updatedRow[timestampCol] = now.toLocaleString();
          updatedRow[teacherCol] = username;
          
          // Ensure hours and category are from Daily_Hours_Setup (not custom if provided)
 updatedRow[hoursCol] = finalHours;
updatedRow[categoryCol] = finalCategory;
          if (recIdCol !== undefined) updatedRow[recIdCol] = recId;

          // Ensure term info is correct
          if (termIdCol !== undefined && termNameCol !== undefined) {
            updatedRow[termIdCol] = term.termId;
            updatedRow[termNameCol] = term.termName;
          }
          
          updatesToApply.push({
            rowIndex: existing.rowIndex,
            rowValues: updatedRow,
            studentId: studentId,
            oldStatus: existing.currentStatus,
            newStatus: newStatus
          });
          Logger.log(`Will UPDATE student ${studentId}: ${existing.currentStatus} → ${newStatus}`);
        } else {
          Logger.log(`Skipping student ${studentId} - status unchanged (${newStatus})`);
        }
      } else {
        // Create new record for this student
        const newRow = new Array(headers.length).fill('');
        
        newRow[attendanceIdCol] = generateUUID();
        newRow[dateCol] = date;
        newRow[stdIdCol] = studentId;
        newRow[barcodeCol] = student.Barcode_ID || '';
        newRow[studentNameCol] = student.Student_Name;
        newRow[studentClassCol] = student.Student_Class;
        newRow[studentSectionCol] = student.Student_Section;
        newRow[statusCol] = newStatus;
        newRow[timestampCol] = now.toLocaleString();
        newRow[teacherCol] = username;
        newRow[classCol] = classSection;
        
        // Use hours from Daily_Hours_Setup validation (not the passed hours parameter)
        newRow[hoursCol] = finalHours;
newRow[categoryCol] = finalCategory;
        if (recIdCol !== undefined) newRow[recIdCol] = recId;
        if (typeCol !== undefined) newRow[typeCol] = 'Episodic';
        if (eventCol !== undefined) newRow[eventCol] = eventName;
        if (termIdCol !== undefined && termNameCol !== undefined) {
          newRow[termIdCol] = term.termId;
          newRow[termNameCol] = term.termName;
        }
        
        newRowsToAppend.push(newRow);
        Logger.log(`Will INSERT new record for student ${studentId}: ${newStatus}`);
      }
    }
    
    // ========================================================
    // 10. Apply updates with lock protection
    // ========================================================
    if (updatesToApply.length > 0) {
      const lock = LockService.getScriptLock();
      let lockAcquired = false;
      
      for (let attempt = 0; attempt < 5; attempt++) {
        if (lock.tryLock(10000)) {
          lockAcquired = true;
          break;
        }
        Utilities.sleep(1000 * (attempt + 1));
      }
      
      if (!lockAcquired) {
        throw new Error('System busy. Unable to acquire lock for updates. Please try again.');
      }
      
      try {
        for (const update of updatesToApply) {
          attendanceSheet.getRange(update.rowIndex, 1, 1, update.rowValues.length)
            .setValues([update.rowValues]);
          Logger.log(`Applied UPDATE for student ${update.studentId}: ${update.oldStatus} → ${update.newStatus}`);
        }
        Logger.log(`Updated ${updatesToApply.length} records`);
      } finally {
        lock.releaseLock();
      }
    }
    
    // ========================================================
    // 11. Insert new records with lock protection
    // ========================================================
    if (newRowsToAppend.length > 0) {
      const lock = LockService.getScriptLock();
      let lockAcquired = false;
      
      for (let attempt = 0; attempt < 5; attempt++) {
        if (lock.tryLock(10000)) {
          lockAcquired = true;
          break;
        }
        Utilities.sleep(1000 * (attempt + 1));
      }
      
      if (!lockAcquired) {
        throw new Error('System busy. Unable to acquire lock for inserts. Please try again.');
      }
      
      try {
        const lastRow = attendanceSheet.getLastRow();
        attendanceSheet.getRange(lastRow + 1, 1, newRowsToAppend.length, headers.length)
          .setValues(newRowsToAppend);
        Logger.log(`Inserted ${newRowsToAppend.length} new records`);
      } finally {
        lock.releaseLock();
      }
    }
    
   // ========================================================
// ========================================================
// 12. Save episodic hours to Daily_Hours_Setup
// ONLY check for existing EPISODIC hours - normal hours are independent
// ========================================================
SpreadsheetApp.flush();

const episodicHoursExist = checkDailyHours(date, classSection, 'episodic');

if (!episodicHoursExist.exists) {
  // No episodic hours for this date/class - safe to save
  try {
        setDailyHours([{
      date: date,
      classSection: classSection,
      category: finalCategory,
      hours: finalHours,
      teacherId: username
    }], 'episodic', recId);
Logger.log(`Added episodic hours to Daily_Hours_Setup: ${finalHours} hours (${finalCategory})`);

  } catch (e) {
    Logger.log(`Warning: Could not save episodic hours: ${e.message}`);
  }
} else {
  Logger.log('Episodic hours already exist for this date/class - using existing configuration');
}

    const presentCount = studentStatuses.filter(s => s.status === 'Present').length;
    const notApplicableCount = studentStatuses.filter(s => s.status === 'Not Applicable').length;
    
    const elapsedTime = Date.now() - startTime;
    Logger.log(`Episodic attendance completed in ${elapsedTime}ms: ${presentCount} Present, ${notApplicableCount} Not Applicable`);
    
    return {
      success: true,
      presentCount: presentCount,
      notApplicableCount: notApplicableCount,
      updatesApplied: updatesToApply.length,
      insertsApplied: newRowsToAppend.length,
      message: `${presentCount} students Present, ${notApplicableCount} Not Applicable`
    };
    
  } catch (e) {
    const elapsedTime = Date.now() - startTime;
    Logger.log(`Episodic error after ${elapsedTime}ms: ${e.message}`);
    throw e;
  }
}

// This function is still here but the appendRow logic is now inside markAttendance.
// It's not directly used in the provided markAttendance anymore.
function appendRow(...rows) { 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Attendance');
  if (!sheet) throw new Error('Student_Attendance sheet not found');
  sheet.appendRow(...rows); 
}

function getActiveCategories(type = 'normal') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hour_Categories');
  if (!sheet) throw new Error('Hour_Categories sheet not found');
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const typeCol = headers.indexOf('Attendance_Type');
  
  const categories = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === true || data[i][2] === 'TRUE') {
      let allowedType = 'both';
      if (typeCol !== -1 && data[i][typeCol]) {
        allowedType = String(data[i][typeCol]).toLowerCase();
      }
      // Include if: no type specified OR type is 'both' OR matches requested type
      if (!type || allowedType === 'both' || allowedType === type) {
        categories.push(data[i][1]);
      }
    }
  }
  return categories;
}

function getHoursOptions(type = 'normal') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hours_Setup');
  if (!sheet) throw new Error('Hours_Setup sheet not found');
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const typeCol = headers.indexOf('Attendance_Type');
  
  const presetHours = [];
  let allowCustom = false;
  let maxCustom = 8;
  
  for (let i = 1; i < data.length; i++) {
    const hoursValue = String(data[i][0]).toLowerCase();
    let allowedType = 'both';
    
    if (typeCol !== -1 && data[i][typeCol]) {
      allowedType = String(data[i][typeCol]).toLowerCase();
    }
    
    const matchesType = !type || allowedType === 'both' || allowedType === type;
    
    if (hoursValue === 'custom') {
      if (matchesType) {
        allowCustom = true;
        // You can add Max_Custom column logic here if needed
      }
    } else if (matchesType) {
      presetHours.push(hoursValue);
    }
  }
  
  // For backward compatibility, just return presetHours array
  return presetHours;
}

// NEW: Check if hours/category are set for a date and class_section
function checkDailyHours(date, classSection, attendanceType = null) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Hours_Setup');
  if (!sheet) throw new Error('Daily_Hours_Setup sheet not found');
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const typeCol = headers.indexOf('Attendance_Type');
  const targetDate = new Date(date).toLocaleDateString('en-CA');
  
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]).toLocaleDateString('en-CA');
    if (rowDate === targetDate && data[i][1] === classSection) {
      const rowType = (typeCol !== -1 && data[i][typeCol]) ? String(data[i][typeCol]) : 'normal';
      
      // If type specified (like 'normal' or 'episodic'), only match that type
      if (attendanceType && rowType === attendanceType) {
        return { exists: true, category: data[i][2], hours: data[i][3] };
      }
      // If no type specified, return any (for backward compatibility)
      if (!attendanceType) {
        return { exists: true, category: data[i][2], hours: data[i][3] };
      }
    }
  }
  return { exists: false };
}

function setDailyHours(entries, attendanceType = 'normal', recId = '') {
  const startTime = Date.now();
  Logger.log(`setDailyHours started with ${entries.length} entries, type: ${attendanceType}`);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Hours_Setup');
  if (!sheet) throw new Error('Daily_Hours_Setup sheet not found');
  
  // Ensure Attendance_Type column exists
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let typeColIndex = headers.indexOf('Attendance_Type');
  
  if (typeColIndex === -1) {
    // Add Attendance_Type column at the end
    typeColIndex = headers.length;
    sheet.getRange(1, typeColIndex + 1).setValue('Attendance_Type');
  }
  const recIdColIndex = headers.indexOf('REC_ID');
  
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('Unable to acquire lock. Please try again.');
  }
  
  try {
    const existingData = sheet.getDataRange().getValues();
    const now = new Date();
    const rowWidth = Math.max(7, headers.length, typeColIndex + 1, recIdColIndex + 1);
    
    const updates = [];
    const newRows = [];
    
    entries.forEach(entry => {
      const targetDate = new Date(entry.date).toLocaleDateString('en-CA');
      let found = false;
      
      for (let i = 1; i < existingData.length; i++) {
        const rowDate = new Date(existingData[i][0]).toLocaleDateString('en-CA');
        if (rowDate === targetDate && existingData[i][1] === entry.classSection) {
          
          // ✅ CRITICAL FIX: Get the existing Attendance_Type
          const existingType = existingData[i][typeColIndex] || 'normal';
          
          // ✅ ONLY update if the existing row has the SAME attendance type
          if (existingType === attendanceType) {
            // Update existing row of same type
            const values = new Array(Math.max(7, existingData[i].length, recIdColIndex + 1)).fill('');
            values[0] = entry.date;
            values[1] = entry.classSection;
            values[2] = entry.category;
            values[3] = entry.hours;
            values[4] = entry.teacherId;
            values[5] = now;
            values[typeColIndex] = attendanceType;
            if (recIdColIndex !== -1) values[recIdColIndex] = recId;
            
            updates.push({
              rowIndex: i + 1,
              values: values
            });
            found = true;
            break;  // Stop searching once found and updated
          }
          // If different type, continue searching (don't update, don't break)
          // This allows both normal and episodic to coexist
        }
      }
      
      if (!found) {
        // Create new row only if no existing row of this type exists
        const newRow = new Array(rowWidth).fill('');
        newRow[0] = entry.date;
        newRow[1] = entry.classSection;
        newRow[2] = entry.category;
        newRow[3] = entry.hours;
        newRow[4] = entry.teacherId;
        newRow[5] = now;
        newRow[typeColIndex] = attendanceType;
        if (recIdColIndex !== -1) newRow[recIdColIndex] = recId;
        newRows.push(newRow);
      }
    });
    
    // Apply updates
    if (updates.length > 0) {
      updates.forEach(update => {
        const numCols = update.values.length;
        sheet.getRange(update.rowIndex, 1, 1, numCols).setValues([update.values]);
      });
    }
    
    // Append new rows
    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, rowWidth).setValues(newRows);
    }
    
    SpreadsheetApp.flush();
    
    Logger.log(`setDailyHours completed, updated: ${updates.length}, new: ${newRows.length}, type: ${attendanceType}`);
    return true;
    
  } catch (error) {
    Logger.log(`Error in setDailyHours: ${error.message}`);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

// Add these functions to your existing GS code

// DEBUG: Get hours conflicts for a specific date and username
function getHoursConflicts(date, username) {
  try {
    console.log('getHoursConflicts called with:', date, username);
    
    const sheet = getHoursSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    console.log('Sheet headers:', headers);
    
    const dateCol = headers.indexOf('Date');
    const teacherCol = headers.indexOf('Teacher_ID');
    const classSectionCol = headers.indexOf('Class_Section');
    const hoursCol = headers.indexOf('Hours');
    const categoryCol = headers.indexOf('Category');
    const teacherNameCol = headers.indexOf('Teacher_Name');
    
    console.log('Column indices - date:', dateCol, 'teacher:', teacherCol, 'classSection:', classSectionCol);
    
    if (dateCol === -1 || teacherCol === -1) {
      console.log('Required columns not found, returning empty array');
      return []; // No conflicts if columns not found
    }
    
    const targetDate = new Date(date).toLocaleDateString('en-CA');
    console.log('Target date:', targetDate);
    
    const conflicts = data.slice(1)
      .filter((row, index) => {
        if (!row[dateCol]) {
          console.log(`Row ${index + 2}: No date found`);
          return false;
        }
        
        let rowDate;
        try {
          rowDate = new Date(row[dateCol]).toLocaleDateString('en-CA');
        } catch (e) {
          console.log(`Row ${index + 2}: Invalid date format: ${row[dateCol]}`);
          return false;
        }
        
        const isSameDate = rowDate === targetDate;
        const isDifferentTeacher = row[teacherCol] && row[teacherCol] !== username;
        
        console.log(`Row ${index + 2}: date=${rowDate}, teacher=${row[teacherCol]}, isSameDate=${isSameDate}, isDifferentTeacher=${isDifferentTeacher}`);
        
        return isSameDate && isDifferentTeacher;
      })
      .map(row => {
        const conflict = {
          classSection: row[classSectionCol] || 'Unknown',
          hours: row[hoursCol] || 'Unknown',
          category: row[categoryCol] || 'Unknown',
          teacherName: row[teacherNameCol] || 'Unknown Teacher'
        };
        console.log('Found conflict:', conflict);
        return conflict;
      });
    
    console.log('Total conflicts found:', conflicts.length);
    return conflicts;
    
  } catch (error) {
    console.error('Error in getHoursConflicts:', error);
    return [];
  }
}

// FIXED: Get class sections with hours status - properly filtered by user
function getClassSectionsWithHoursStatus(date, username) {
  try {
    const sheet = getHoursSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const dateCol = headers.indexOf('Date');
    const classSectionCol = headers.indexOf('Class_Section');
    const hoursCol = headers.indexOf('Hours');
    const categoryCol = headers.indexOf('Category');
    
    const existingHours = {};
    const targetDate = new Date(date).toLocaleDateString('en-CA');
    
    if (dateCol !== -1 && classSectionCol !== -1) {
      data.slice(1)
        .filter(row => {
          if (!row[dateCol]) return false;
          const rowDate = new Date(row[dateCol]).toLocaleDateString('en-CA');
          return rowDate === targetDate;
        })
        .forEach(row => {
          existingHours[row[classSectionCol]] = {
            hours: row[hoursCol],
            category: row[categoryCol]
          };
        });
    }
    
    // Get assigned class sections for the specific user
    let userClassSections = [];
    if (username) {
      // Get user type first
      const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
      if (userSheet) {
        const userData = userSheet.getDataRange().getValues();
        const headers = userData[0];
        const usernameCol = headers.indexOf('Username');
        const typeCol = headers.indexOf('Type');
        
        let userType = 'teacher';
        for (let i = 1; i < userData.length; i++) {
          if (userData[i][usernameCol] === username) {
            userType = userData[i][typeCol] || 'teacher';
            break;
          }
        }
        
        if (userType === "teacher") {
          userClassSections = getAssignedClassSections(username);
        } else {
          userClassSections = getClassSections();
        }
      }
    }
    
    return userClassSections.map(cs => ({
      classSection: cs,
      hasHours: !!existingHours[cs],
      hours: existingHours[cs]?.hours || '',
      category: existingHours[cs]?.category || ''
    }));
    
  } catch (error) {
    console.error('Error in getClassSectionsWithHoursStatus:', error);
    return [];
  }
}

// NEW: Handle batch hours form submission
function handleBatchHoursSubmission(formData) {
  const lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(10000)) {
      throw new Error('Unable to acquire lock. Please try again.');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Hours_Setup');
    if (!sheet) throw new Error('Daily_Hours_Setup sheet not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const now = new Date().toLocaleString();
    
    let successCount = 0;
    
    formData.entries.forEach(entry => {
      const targetDate = new Date(entry.date).toLocaleDateString('en-CA');
      let found = false;
      
      // Check for existing entry
      for (let i = 1; i < data.length; i++) {
        const rowDate = new Date(data[i][0]).toLocaleDateString('en-CA');
        if (rowDate === targetDate && data[i][1] === entry.classSection) {
          // Update existing
          sheet.getRange(i + 1, 3).setValue(entry.category); // Category
          sheet.getRange(i + 1, 4).setValue(entry.hours);   // Hours
          sheet.getRange(i + 1, 5).setValue(entry.teacherId); // Teacher ID
          sheet.getRange(i + 1, 6).setValue(now);           // Timestamp
          found = true;
          successCount++;
          break;
        }
      }
      
      if (!found) {
        // Append new
        sheet.appendRow([
          entry.date,
          entry.classSection,
          entry.category,
          entry.hours,
          entry.teacherId,
          now
        ]);
        successCount++;
      }
    });
    
    return { success: true, count: successCount };
    
  } catch (error) {
    console.error('Batch hours submission error:', error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

// Helper function to get the hours sheet
function getHoursSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Daily_Hours_Setup');
  if (!sheet) {
    throw new Error('Daily_Hours_Setup sheet not found');
  }
  return sheet;
}

// Get hours status for dashboard visualization
function getHoursStatus(date) {
  try {
    const sheet = getHoursSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const dateCol = headers.indexOf('Date');
    const classSectionCol = headers.indexOf('Class_Section');
    const hoursCol = headers.indexOf('Hours');
    const categoryCol = headers.indexOf('Category');
    const teacherCol = headers.indexOf('Teacher_ID');
    
    const targetDate = new Date(date).toLocaleDateString('en-CA');
    const status = [];
    
    data.slice(1)
      .filter(row => {
        if (!row[dateCol]) return false;
        const rowDate = new Date(row[dateCol]).toLocaleDateString('en-CA');
        return rowDate === targetDate;
      })
      .forEach(row => {
        status.push({
          classSection: row[classSectionCol],
          hours: row[hoursCol],
          category: row[categoryCol],
          teacherId: row[teacherCol]
        });
      });
    
    return status;
  } catch (error) {
    console.error('Error in getHoursStatus:', error);
    return [];
  }
}

// NEW: Update batch hours for existing attendance records
function updateBatchHours(date, classSection, category, hours, username) {
  const lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(10000)) {
      return { success: false, message: 'Unable to acquire lock. Please try again.' };
    }
    
    // First update Daily_Hours_Setup
    const hoursSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Hours_Setup');
    if (!hoursSheet) throw new Error('Daily_Hours_Setup sheet not found');
    
    const hoursData = hoursSheet.getDataRange().getValues();
    const targetDate = new Date(date).toLocaleDateString('en-CA');
    let updatedHoursSetup = false;
    const now = new Date().toLocaleString();
    
    // Update or create entry in Daily_Hours_Setup
    for (let i = 1; i < hoursData.length; i++) {
      const rowDate = new Date(hoursData[i][0]).toLocaleDateString('en-CA');
      if (rowDate === targetDate && hoursData[i][1] === classSection) {
        hoursSheet.getRange(i + 1, 3).setValue(category); // Category
        hoursSheet.getRange(i + 1, 4).setValue(hours);   // Hours
        hoursSheet.getRange(i + 1, 5).setValue(username); // Teacher ID
        hoursSheet.getRange(i + 1, 6).setValue(now);     // Timestamp
        updatedHoursSetup = true;
        break;
      }
    }
    
    if (!updatedHoursSetup) {
      hoursSheet.appendRow([date, classSection, category, hours, username, now]);
    }
    
    // Now update existing attendance records
    const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Attendance');
    if (!attendanceSheet) throw new Error('Student_Attendance sheet not found');
    
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const headers = attendanceData[0];
    
    const dateCol = headers.indexOf('Date');
    const classSectionCol = headers.indexOf('Class_Section');
    const hoursCol = headers.indexOf('Hours');
    const categoryCol = headers.indexOf('Category');
    
    if (dateCol === -1 || classSectionCol === -1 || hoursCol === -1 || categoryCol === -1) {
      return { success: false, message: 'Required columns not found in attendance sheet.' };
    }
    
    let updatedCount = 0;
    
    for (let i = 1; i < attendanceData.length; i++) {
      const rowDate = new Date(attendanceData[i][dateCol]).toLocaleDateString('en-CA');
      if (rowDate === targetDate && attendanceData[i][classSectionCol] === classSection) {
        attendanceSheet.getRange(i + 1, hoursCol + 1).setValue(hours);
        attendanceSheet.getRange(i + 1, categoryCol + 1).setValue(category);
        updatedCount++;
      }
    }
    
    SpreadsheetApp.flush();
    
    return {
      success: true,
      updated: updatedCount,
      message: updatedCount > 0 
        ? `Updated ${updatedCount} attendance records.` 
        : 'No attendance records found to update.'
    };
    
  } catch (error) {
    console.error('Error in updateBatchHours:', error);
    return { success: false, message: error.message };
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}


// NEW: Extend getUserPermissions to include canSetupHours
// UPDATE THIS EXISTING FUNCTION (replace it completely):
function getUserPermissions(username) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User_Permissions');
    if (!sheet) throw new Error('User_Permissions sheet not found');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        return {
          canMarkManual: data[i][1] === 'TRUE' || data[i][1] === true,
          canMarkBarcode: data[i][2] === 'TRUE' || data[i][2] === true,
          canSeeSummary: data[i][3] === 'TRUE' || data[i][3] === true,
          canSetupHours: data[i][4] === 'TRUE' || data[i][4] === true,
          canUpdateBatchHours: data[i][5] === 'TRUE' || data[i][5] === true,
          canDownloadReports: data[i][6] === 'TRUE' || data[i][6] === true,
          canDownloadAllReports: data[i][7] === 'TRUE' || data[i][7] === true,
          canMarkEpisodic: data[i][8] === 'TRUE' || data[i][8] === true  // ADD THIS LINE
        };
      }
    }
    return { 
      canMarkManual: false, 
      canMarkBarcode: false, 
      canSeeSummary: false, 
      canSetupHours: false,
      canUpdateBatchHours: false,
      canDownloadReports: false,
      canDownloadAllReports: false,
      canMarkEpisodic: false  // ADD THIS LINE
    };
  } catch (e) {
    Logger.log('Error in getUserPermissions: ' + e.message);
    throw e;
  }
}


// =============== REPORT FUNCTIONS ===============

// NEW: Generate student attendance report for PDF/CSV download
function generateStudentAttendanceReport(date, classSection) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get student data
    const studentSheet = ss.getSheetByName('Student_Data');
    if (!studentSheet) throw new Error('Student_Data sheet not found');
    
    const attendanceSheet = ss.getSheetByName('Student_Attendance');
    if (!attendanceSheet) throw new Error('Student_Attendance sheet not found');
    
    // Get all attendance for the given date and class
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const headers = attendanceData[0];
    const headerMap = new Map(headers.map((h, i) => [h, i]));
    
    const requiredHeaders = ['Std_ID', 'Date', 'Status', 'Student_Name', 'Student_Class', 'Student_Section'];
    requiredHeaders.forEach(h => {
      if (!headerMap.has(h)) throw new Error(`Required column ${h} not found in Student_Attendance`);
    });
    
    const [className, section] = classSection.split('-');
    const targetDate = new Date(date).toLocaleDateString('en-CA');
    
    // Filter attendance records for the specific date and class
    const filteredAttendance = [];
    for (let i = 1; i < attendanceData.length; i++) {
      const row = attendanceData[i];
      const rowDate = new Date(row[headerMap.get('Date')]).toLocaleDateString('en-CA');
      const rowClass = row[headerMap.get('Student_Class')];
      const rowSection = row[headerMap.get('Student_Section')];
      
      if (rowDate === targetDate && rowClass === className && rowSection === section) {
        filteredAttendance.push({
          studentId: row[headerMap.get('Std_ID')],
          studentName: row[headerMap.get('Student_Name')],
          className: rowClass,
          section: rowSection,
          date: date,
          status: row[headerMap.get('Status')]
        });
      }
    }
    
    // If no attendance records found, get all students from Student_Data
    if (filteredAttendance.length === 0) {
      const studentData = studentSheet.getDataRange().getValues();
      const studentHeaders = studentData[0];
      const studentHeaderMap = new Map(studentHeaders.map((h, i) => [h, i]));
      
      const stdIdCol = studentHeaderMap.get('Std_ID');
      const nameCol = studentHeaderMap.get('Student_Name');
      const classCol = studentHeaderMap.get('Student_Class');
      const sectionCol = studentHeaderMap.get('Student_Section');
      
      for (let i = 1; i < studentData.length; i++) {
        const row = studentData[i];
        if (row[classCol] === className && row[sectionCol] === section) {
          filteredAttendance.push({
            studentId: row[stdIdCol],
            studentName: row[nameCol],
            className: className,
            section: section,
            date: date,
            status: 'Not Marked'
          });
        }
      }
    }
    
    // Sort by student name
    filteredAttendance.sort((a, b) => a.studentName.localeCompare(b.studentName));
    
    return filteredAttendance;
    
  } catch (error) {
    console.error('Error generating student attendance report:', error);
    throw error;
  }
}

// NEW: Generate student monthly summary report
function generateStudentMonthlySummary(studentId, month, year) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get student info
    const studentSheet = ss.getSheetByName('Student_Data');
    if (!studentSheet) throw new Error('Student_Data sheet not found');
    
    const studentData = studentSheet.getDataRange().getValues();
    const studentHeaders = studentData[0];
    const studentHeaderMap = new Map(studentHeaders.map((h, i) => [h, i]));
    
    // Find student
    let studentInfo = null;
    const stdIdCol = studentHeaderMap.get('Std_ID');
    const nameCol = studentHeaderMap.get('Student_Name');
    const classCol = studentHeaderMap.get('Student_Class');
    const sectionCol = studentHeaderMap.get('Student_Section');
    
    for (let i = 1; i < studentData.length; i++) {
      if (String(studentData[i][stdIdCol]) === String(studentId)) {
        studentInfo = {
          studentId: studentData[i][stdIdCol],
          studentName: studentData[i][nameCol],
          className: studentData[i][classCol],
          section: studentData[i][sectionCol]
        };
        break;
      }
    }
    
    if (!studentInfo) {
      throw new Error(`Student with ID ${studentId} not found`);
    }
    
    // Get attendance data for the month
    const attendanceSheet = ss.getSheetByName('Student_Attendance');
    if (!attendanceSheet) throw new Error('Student_Attendance sheet not found');
    
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const headers = attendanceData[0];
    const headerMap = new Map(headers.map((h, i) => [h, i]));
    
    const stdIdColAttendance = headerMap.get('Std_ID');
    const dateCol = headerMap.get('Date');
    const statusCol = headerMap.get('Status');
    const hoursCol = headerMap.has('Hours') ? headerMap.get('Hours') : -1;
    const categoryCol = headerMap.has('Category') ? headerMap.get('Category') : -1;
    
    // Calculate month start and end dates
    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 0);
    
    // Initialize counters
    let presents = 0;
    let absents = 0;
    let leaves = 0;
    let totalDays = 0;
    
    // Process attendance records
    for (let i = 1; i < attendanceData.length; i++) {
      const row = attendanceData[i];
      
      // Check if this is the right student
      if (String(row[stdIdColAttendance]) !== String(studentId)) {
        continue;
      }
      
      // Check date
      const rowDate = new Date(row[dateCol]);
      if (rowDate >= startDate && rowDate <= endDate) {
        const status = String(row[statusCol]).toLowerCase();
        
        switch (status) {
          case 'present':
            presents++;
            break;
          case 'absent':
            absents++;
            break;
          case 'leave':
            leaves++;
            break;
        }
        totalDays++;
      }
    }
    
    // Calculate percentage
    const presentPercentage = totalDays > 0 ? (presents / totalDays) * 100 : 0;
    
    return {
      ...studentInfo,
      month: month,
      year: year,
      presents: presents,
      absents: absents,
      leaves: leaves,
      totalDays: totalDays,
      presentPercentage: Math.round(presentPercentage * 100) / 100 // Round to 2 decimal places
    };
    
  } catch (error) {
    console.error('Error generating monthly summary:', error);
    throw error;
  }
}

// NEW: Get available months for a student
function getAvailableMonthsForStudent(studentId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = ss.getSheetByName('Student_Attendance');
    if (!attendanceSheet) throw new Error('Student_Attendance sheet not found');
    
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const headers = attendanceData[0];
    const headerMap = new Map(headers.map((h, i) => [h, i]));
    
    const stdIdCol = headerMap.get('Std_ID');
    const dateCol = headerMap.get('Date');
    
    const monthsSet = new Set();
    
    for (let i = 1; i < attendanceData.length; i++) {
      const row = attendanceData[i];
      if (String(row[stdIdCol]) === String(studentId)) {
        const date = new Date(row[dateCol]);
        const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
        monthsSet.add(monthKey);
      }
    }
    
    // Convert to array and sort descending (newest first)
    const months = Array.from(monthsSet)
      .sort((a, b) => b.localeCompare(a))
      .map(monthStr => {
        const [year, month] = monthStr.split('-');
        return {
          year: parseInt(year),
          month: parseInt(month),
          display: `${new Date(year, month - 1).toLocaleString('default', { month: 'long' })} ${year}`
        };
      });
    
    return months;
    
  } catch (error) {
    console.error('Error getting available months:', error);
    throw error;
  }
}

// NEW: Get all students for dropdown
function getAllStudents(classSection = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = ss.getSheetByName('Student_Data');
    if (!studentSheet) throw new Error('Student_Data sheet not found');
    
    const studentData = studentSheet.getDataRange().getValues();
    const headers = studentData[0];
    const headerMap = new Map(headers.map((h, i) => [h, i]));
    
    const requiredCols = ['Std_ID', 'Student_Name', 'Student_Class', 'Student_Section'];
    requiredCols.forEach(col => {
      if (!headerMap.has(col)) throw new Error(`Required column ${col} not found`);
    });
    
    const stdIdCol = headerMap.get('Std_ID');
    const nameCol = headerMap.get('Student_Name');
    const classCol = headerMap.get('Student_Class');
    const sectionCol = headerMap.get('Student_Section');
    
    const students = [];
    
    for (let i = 1; i < studentData.length; i++) {
      const row = studentData[i];
      const studentClass = row[classCol];
      const studentSection = row[sectionCol];
      
      if (classSection) {
        const [className, section] = classSection.split('-');
        if (studentClass !== className || studentSection !== section) {
          continue;
        }
      }
      
      students.push({
        studentId: row[stdIdCol],
        studentName: row[nameCol],
        className: studentClass,
        section: studentSection,
        classSection: `${studentClass}-${studentSection}`
      });
    }
    
    // Sort by name
    students.sort((a, b) => a.studentName.localeCompare(b.studentName));
    
    return students;
    
  } catch (error) {
    console.error('Error getting students:', error);
    throw error;
  }
}

// NEW: Generate CSV content
function generateCSVContent(data) {
  if (!data || data.length === 0) return '';
  
  // Get headers from first object
  const headers = Object.keys(data[0]);
  
  // Create CSV content
  let csv = headers.join(',') + '\n';
  
  data.forEach(row => {
    const values = headers.map(header => {
      const value = row[header];
      // Escape quotes and wrap in quotes if contains comma or quotes
      if (typeof value === 'string' && (value.includes(',') || value.includes('"'))) {
        return '"' + value.replace(/"/g, '""') + '"';
      }
      return value;
    });
    csv += values.join(',') + '\n';
  });
  
  return csv;
}


// Add this function
function healthCheck() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const requiredSheets = ['Users', 'Student_Data', 'Student_Attendance'];
    
    const missingSheets = requiredSheets.filter(sheetName => 
      !sheets.some(sheet => sheet.getName() === sheetName)
    );
    
    return {
      status: missingSheets.length === 0 ? 'healthy' : 'degraded',
      sheets: sheets.map(s => s.getName()),
      missingSheets: missingSheets,
      timestamp: new Date().toISOString()
    };
  } catch (e) {
    return {
      status: 'unhealthy',
      error: e.message,
      timestamp: new Date().toISOString()
    };
  }
}



// =============== TERM VALIDATION FUNCTIONS ===============
// Add these functions to your code.gs (anywhere, preferably after markAttendance)

/**
 * Get ACTIVE term for a class section on specific date
 * Checks date range, Class_Terms.Is_Active, AND Terms.Is_Active
 */
function getTermForClassSection(classSection, date = null) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_Terms');
    if (!sheet) {
      console.error('Class_Terms sheet not found');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const classCol = headers.indexOf('Class_Section');
    const termIdCol = headers.indexOf('Term_ID');
    const activeCol = headers.indexOf('Is_Active');
    const startCol = headers.indexOf('Start_Date');
    const endCol = headers.indexOf('End_Date');
    
    if (classCol === -1 || termIdCol === -1 || activeCol === -1 || startCol === -1 || endCol === -1) {
      console.error('Required columns not found in Class_Terms');
      return null;
    }
    
    const targetDate = date ? new Date(date) : new Date();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][classCol] === classSection) {
        // 1. Check if term assignment is ACTIVE in Class_Terms
        const isAssignmentActive = data[i][activeCol] === true || data[i][activeCol] === 'TRUE';
        if (!isAssignmentActive) {
          console.log(`Class ${classSection}: Assignment is INACTIVE in Class_Terms`);
          continue; // Skip inactive assignments
        }
        
        // 2. Check date range using dates from Class_Terms sheet
        const startDate = new Date(data[i][startCol]);
        const endDate = new Date(data[i][endCol]);
        
        if (!(targetDate >= startDate && targetDate <= endDate)) {
          console.log(`Class ${classSection}: Outside date range (${startDate} to ${endDate})`);
          continue; // Skip if outside date range
        }
        
        // 3. Get term details AND check if term is ACTIVE in Terms sheet
        const termDetails = getTermDetails(data[i][termIdCol]);
        
        if (!termDetails) {
          console.log(`Class ${classSection}: Term ${data[i][termIdCol]} not found in Terms sheet`);
          continue;
        }
        
        // 4. CRITICAL: Check if term itself is ACTIVE in Terms sheet
        if (!termDetails.isActive) { // CHANGED FROM isCurrent to isActive
          console.log(`Class ${classSection}: Term ${data[i][termIdCol]} is INACTIVE in Terms sheet`);
          return null; // Term is inactive
        }
        
        return {
          termId: data[i][termIdCol],
          termName: termDetails.termName,
          startDate: data[i][startCol],
          endDate: data[i][endCol],
          isActive: termDetails.isActive, // From Terms.Is_Active (formerly Is_Current)
          description: termDetails.description
        };
      }
    }
    
    console.log(`No active term found for ${classSection} on ${date || 'today'}`);
    return null;
    
  } catch (error) {
    console.error('Error getting term for class section:', error);
    return null;
  }
}

/**
 * Get term details from Terms sheet
 * Checks Is_Active status (formerly Is_Current)
 */
function getTermDetails(termId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Terms');
    if (!sheet) {
      console.error('Terms sheet not found');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const termIdCol = headers.indexOf('Term_ID');
    const nameCol = headers.indexOf('Term_Name');
    const activeCol = headers.indexOf('Is_Active'); // CHANGED FROM Is_Current
    
    if (termIdCol === -1 || nameCol === -1) {
      console.error('Required columns not found in Terms sheet');
      return null;
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][termIdCol] === termId) {
        // Check Is_Active (formerly Is_Current)
        const isActive = activeCol !== -1 ? 
          (data[i][activeCol] === true || data[i][activeCol] === 'TRUE') : 
          false; // If column doesn't exist, assume inactive for safety
        
        return {
          termId: data[i][termIdCol],
          termName: data[i][nameCol],
          startDate: data[i][headers.indexOf('Start_Date')],
          endDate: data[i][headers.indexOf('End_Date')],
          isActive: isActive, // CHANGED FROM isCurrent
          description: data[i][headers.indexOf('Description')]
        };
      }
    }
    
    console.log(`Term ${termId} not found in Terms sheet`);
    return null;
    
  } catch (error) {
    console.error('Error getting term details:', error);
    return null;
  }
}

/**
 * Get all ACTIVE class-term assignments
 * FIXED: Uses dates from Class_Terms, not Terms
 */
function getActiveClassTerms() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_Terms');
    if (!sheet) {
      console.error('Class_Terms sheet not found');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const classCol = headers.indexOf('Class_Section');
    const termIdCol = headers.indexOf('Term_ID');
    const activeCol = headers.indexOf('Is_Active');
    const startCol = headers.indexOf('Start_Date');
    const endCol = headers.indexOf('End_Date');
    
    const activeAssignments = [];
    
    for (let i = 1; i < data.length; i++) {
      const isActive = data[i][activeCol] === true || data[i][activeCol] === 'TRUE';
      
      if (isActive && data[i][classCol] && data[i][termIdCol]) {
        const termDetails = getTermDetails(data[i][termIdCol]);
        
        activeAssignments.push({
          classSection: data[i][classCol],
          termId: data[i][termIdCol],
          termName: termDetails ? termDetails.termName : `Term ${data[i][termIdCol]}`,
          startDate: data[i][startCol],  // From Class_Terms
          endDate: data[i][endCol],      // From Class_Terms
          isActive: true,
          termDetails: termDetails
        });
      }
    }
    
    return activeAssignments;
    
  } catch (error) {
    console.error('Error getting active class terms:', error);
    return [];
  }
}

/**
 * Validate term system setup
 */
function validateTermSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const results = [];
    
    // 1. Check if sheets exist
    const termsSheet = ss.getSheetByName('Terms');
    const classTermsSheet = ss.getSheetByName('Class_Terms');
    
    if (!termsSheet) {
      results.push('❌ Terms sheet not found');
    } else {
      results.push('✅ Terms sheet exists');
      
      // Check columns
      const termsHeaders = termsSheet.getRange(1, 1, 1, termsSheet.getLastColumn()).getValues()[0];
      const requiredTermsCols = ['Term_ID', 'Term_Name', 'Is_Active']; // CHANGED FROM Is_Current
      requiredTermsCols.forEach(col => {
        if (termsHeaders.includes(col)) {
          results.push(`  ✅ ${col} column exists`);
        } else {
          results.push(`  ❌ ${col} column missing`);
        }
      });
    }
    
    if (!classTermsSheet) {
      results.push('❌ Class_Terms sheet not found');
    } else {
      results.push('✅ Class_Terms sheet exists');
      
      // Check columns
      const classTermsHeaders = classTermsSheet.getRange(1, 1, 1, classTermsSheet.getLastColumn()).getValues()[0];
      const requiredClassTermsCols = ['Class_Section', 'Term_ID', 'Is_Active'];
      requiredClassTermsCols.forEach(col => {
        if (classTermsHeaders.includes(col)) {
          results.push(`  ✅ ${col} column exists`);
        } else {
          results.push(`  ❌ ${col} column missing`);
        }
      });
      
      // Count active assignments
      const data = classTermsSheet.getDataRange().getValues();
      const headers = data[0];
      const activeCol = headers.indexOf('Is_Active');
      const classCol = headers.indexOf('Class_Section');
      
      let activeCount = 0;
      let totalCount = 0;
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][classCol]) {
          totalCount++;
          const isActive = data[i][activeCol] === true || data[i][activeCol] === 'TRUE';
          if (isActive) activeCount++;
        }
      }
      
      results.push(`  📊 ${activeCount} of ${totalCount} class terms are ACTIVE`);
    }
    
    // 2. Check Student_Attendance columns
    const attendanceSheet = ss.getSheetByName('Student_Attendance');
    if (attendanceSheet) {
      const attendanceHeaders = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0];
      
      if (attendanceHeaders.includes('Term_ID')) {
        results.push('✅ Term_ID column exists in Student_Attendance');
      } else {
        results.push('⚠️ Term_ID column missing in Student_Attendance (attendance will work but no term info)');
      }
      
      if (attendanceHeaders.includes('Term_Name')) {
        results.push('✅ Term_Name column exists in Student_Attendance');
      } else {
        results.push('⚠️ Term_Name column missing in Student_Attendance');
      }
    }
    
    // 3. Test a sample class
    const testClasses = ['10-A', '9-C'];
    testClasses.forEach(classSection => {
      const term = getTermForClassSection(classSection, '2025-10-15');
      if (term) {
        results.push(`✅ ${classSection}: ${term.termName} (${term.termId}) - ACTIVE`);
      } else {
        results.push(`❌ ${classSection}: No active term found`);
      }
    });
    
    return results.join('\n');
    
  } catch (error) {
    return `Validation error: ${error.message}`;
  }
}

/**
 * Quick test: Check term for a class
 */
function testTermForClass() {
  const classSection = '10-A'; // Change to your class
  const date = '2025-10-15';   // Change to test date
  
  const term = getTermForClassSection(classSection, date);
  
  if (term) {
    return `✅ ${classSection} has ACTIVE term:\n` +
           `Term: ${term.termName}\n` +
           `ID: ${term.termId}\n` +
           `Dates: ${term.startDate} to ${term.endDate}\n` +
           `Is Active: ${term.isActive ? 'Yes' : 'No'}`; // CHANGED FROM Is_Current
  } else {
    return `❌ No ACTIVE term found for ${classSection} on ${date}`;
  }
}

/**
 * List all active class-term assignments
 */
function listActiveAssignments() {
  const assignments = getActiveClassTerms();
  
  if (assignments.length === 0) {
    return 'No active class-term assignments found';
  }
  
  let result = `Active Class-Term Assignments (${assignments.length}):\n\n`;
  
  assignments.forEach((assignment, index) => {
    result += `${index + 1}. ${assignment.classSection}\n`;
    result += `   Term: ${assignment.termName} (${assignment.termId})\n`;
    result += `   Dates: ${assignment.startDate} to ${assignment.endDate}\n`;
    result += `   Term Status: ${assignment.termDetails?.isActive ? 'Active' : 'Inactive'}\n`; // CHANGED FROM isCurrent
    result += `\n`;
  });
  
  return result;
}

/**
 * Get term status with detailed error information
 * Returns {hasTerm: boolean, term: object, error: string}
 */
function getTermStatusWithDetails(classSection, date) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_Terms');
    if (!sheet) {
      return {
        hasTerm: false,
        term: null,
        error: 'Class_Terms sheet not found. Please create it.'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const classCol = headers.indexOf('Class_Section');
    const termIdCol = headers.indexOf('Term_ID');
    const activeCol = headers.indexOf('Is_Active');
    const startCol = headers.indexOf('Start_Date');
    const endCol = headers.indexOf('End_Date');
    
    const targetDate = new Date(date);
    let classExists = false;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][classCol] === classSection) {
        classExists = true;
        
        const isActive = data[i][activeCol] === true || data[i][activeCol] === 'TRUE';
        const startDate = new Date(data[i][startCol]);
        const endDate = new Date(data[i][endCol]);
        
        if (isActive) {
          if (targetDate >= startDate && targetDate <= endDate) {
            const term = getTermDetails(data[i][termIdCol]);
            return {
              hasTerm: true,
              term: term,
              error: null
            };
          } else {
            return {
              hasTerm: false,
              term: null,
              error: `Term exists but date ${date} is outside term range (${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()})`
            };
          }
        } else {
          return {
            hasTerm: false,
            term: null,
            error: `Term exists but is NOT ACTIVE for ${classSection}`
          };
        }
      }
    }
    
    if (!classExists) {
      return {
        hasTerm: false,
        term: null,
        error: `Class ${classSection} not found in Class_Terms sheet`
      };
    }
    
    return {
      hasTerm: false,
      term: null,
      error: 'Unknown error'
    };
    
  } catch (error) {
    return {
      hasTerm: false,
      term: null,
      error: error.message
    };
  }
}


function debugTermSystem() {
  console.log("=== DEBUG TERM SYSTEM ===");
  console.log("Today's date:", new Date());
  console.log("Today in YYYY-MM-DD:", new Date().toISOString().split('T')[0]);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classTermsSheet = ss.getSheetByName('Class_Terms');
  const termsSheet = ss.getSheetByName('Terms');
  
  // Check Class_Terms data
  console.log("\n=== Class_Terms Sheet Data ===");
  const classTermsData = classTermsSheet.getDataRange().getValues();
  const classTermsHeaders = classTermsData[0];
  console.log("Headers:", classTermsHeaders);
  
  const classCol = classTermsHeaders.indexOf('Class_Section');
  const termIdCol = classTermsHeaders.indexOf('Term_ID');
  const activeCol = classTermsHeaders.indexOf('Is_Active');
  const startCol = classTermsHeaders.indexOf('Start_Date');
  const endCol = classTermsHeaders.indexOf('End_Date');
  
  console.log("\nClass_Terms entries for 10-A:");
  for (let i = 1; i < classTermsData.length; i++) {
    if (classTermsData[i][classCol] === '10-A') {
      console.log(`Row ${i + 1}:`);
      console.log(`  Term_ID: ${classTermsData[i][termIdCol]} (type: ${typeof classTermsData[i][termIdCol]})`);
      console.log(`  Is_Active: ${classTermsData[i][activeCol]} (type: ${typeof classTermsData[i][activeCol]})`);
      console.log(`  Start_Date: ${classTermsData[i][startCol]} (type: ${typeof classTermsData[i][startCol]})`);
      console.log(`  End_Date: ${classTermsData[i][endCol]} (type: ${typeof classTermsData[i][endCol]})`);
      
      // Try to parse dates
      try {
        const startDate = new Date(classTermsData[i][startCol]);
        const endDate = new Date(classTermsData[i][endCol]);
        const today = new Date();
        
        console.log(`  Parsed Start: ${startDate} (Valid: ${!isNaN(startDate.getTime())})`);
        console.log(`  Parsed End: ${endDate} (Valid: ${!isNaN(endDate.getTime())})`);
        console.log(`  Today within range? ${today >= startDate && today <= endDate}`);
      } catch (e) {
        console.log(`  Date parsing error: ${e.message}`);
      }
    }
  }
  
  // Test getTermForClassSection function
  console.log("\n=== Testing getTermForClassSection ===");
  const testDate = '2026-01-04';
  const result = getTermForClassSection('10-A', testDate);
  console.log(`Result for 10-A on ${testDate}:`, result);
  
  if (!result) {
    console.log("Why no term found? Let's trace...");
    
    // Manually trace through the logic
    const targetDate = new Date(testDate);
    console.log(`Target date: ${targetDate}`);
    
    for (let i = 1; i < classTermsData.length; i++) {
      if (classTermsData[i][classCol] === '10-A') {
        const isActive = classTermsData[i][activeCol];
        console.log(`\nRow ${i + 1} - Is_Active: ${isActive} (type: ${typeof isActive})`);
        
        const startDate = new Date(classTermsData[i][startCol]);
        const endDate = new Date(classTermsData[i][endCol]);
        
        console.log(`Start date: ${startDate} (${classTermsData[i][startCol]})`);
        console.log(`End date: ${endDate} (${classTermsData[i][endCol]})`);
        console.log(`Target >= Start? ${targetDate >= startDate}`);
        console.log(`Target <= End? ${targetDate <= endDate}`);
        console.log(`Within range? ${targetDate >= startDate && targetDate <= endDate}`);
        
        if (isActive === true || isActive === 'TRUE') {
          console.log("✅ Is_Active check PASSED");
          if (targetDate >= startDate && targetDate <= endDate) {
            console.log("✅ Date range check PASSED");
          } else {
            console.log("❌ Date range check FAILED");
          }
        } else {
          console.log("❌ Is_Active check FAILED");
        }
      }
    }
  }
}


function testTermSystemUpdated() {
  console.log("Testing Updated Term System...");
  
  const testClass = '10-A'; // Change to your actual class
  const testDate = new Date().toISOString().split('T')[0]; // Today
  
  console.log(`Testing term for ${testClass} on ${testDate}:`);
  const term = getTermForClassSection(testClass, testDate);
  
  if (term) {
    console.log("✅ Term found:");
    console.log(`   Term ID: ${term.termId}`);
    console.log(`   Term Name: ${term.termName}`);
    console.log(`   Class-Specific Dates: ${term.startDate} to ${term.endDate}`);
    console.log(`   Is Active (from Terms): ${term.isActive ? 'Yes' : 'No'}`); // CHANGED
    console.log(`   Is Active (from Class_Terms): ${term.isActive ? 'Yes' : 'No'}`);
    
    // Also check what's in Terms sheet for comparison
    const termDetails = getTermDetails(term.termId);
    if (termDetails) {
      console.log(`   Master Dates in Terms sheet: ${termDetails.startDate} to ${termDetails.endDate}`);
      console.log(`   Are dates different? ${term.startDate !== termDetails.startDate || term.endDate !== termDetails.endDate ? 'Yes' : 'No'}`);
    }
  } else {
    console.log("❌ No active term found for this class on today's date");
  }
  
  // Test multiple classes
  console.log("\nTesting all active class-term assignments:");
  const assignments = getActiveClassTerms();
  if (assignments.length > 0) {
    assignments.forEach((assignment, index) => {
      console.log(`${index + 1}. ${assignment.classSection}`);
      console.log(`   Term: ${assignment.termName} (${assignment.termId})`);
      console.log(`   Dates: ${assignment.startDate} to ${assignment.endDate}`);
      console.log(`   Is Active: ${assignment.termDetails?.isActive ? 'Yes' : 'No'}`); // CHANGED
    });
  } else {
    console.log("No active class-term assignments found");
  }
}


/**
 * Get detailed term status including why a term might be inaccessible
 */
function getTermAccessStatus(classSection, date) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_Terms');
    if (!sheet) {
      return {
        accessible: false,
        error: 'Class_Terms sheet not found'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const classCol = headers.indexOf('Class_Section');
    const termIdCol = headers.indexOf('Term_ID');
    const activeCol = headers.indexOf('Is_Active');
    const startCol = headers.indexOf('Start_Date');
    const endCol = headers.indexOf('End_Date');
    
    const targetDate = new Date(date);
    
    // Check if class exists in Class_Terms
    for (let i = 1; i < data.length; i++) {
      if (data[i][classCol] === classSection) {
        const assignmentActive = data[i][activeCol] === true || data[i][activeCol] === 'TRUE';
        const startDate = new Date(data[i][startCol]);
        const endDate = new Date(data[i][endCol]);
        
        if (!assignmentActive) {
          return {
            accessible: false,
            error: `Class-term assignment is inactive for ${classSection}.`
          };
        }
        
        if (targetDate < startDate) {
          return {
            accessible: false,
            error: `Term starts on ${startDate.toLocaleDateString()}. Cannot mark attendance before term start.`
          };
        }
        
        if (targetDate > endDate) {
          return {
            accessible: false,
            error: `Term ended on ${endDate.toLocaleDateString()}. Cannot mark attendance after term end.`
          };
        }
        
        // Now check Terms sheet
        const termDetails = getTermDetails(data[i][termIdCol]);
        if (!termDetails) {
          return {
            accessible: false,
            error: `Associated term not found in Terms sheet.`
          };
        }
        
        if (!termDetails.isActive) {
          return {
            accessible: false,
            error: `The term "${termDetails.termName}" is marked as inactive. Contact administrator to activate it.`
          };
        }
        
        // All checks passed
        return {
          accessible: true,
          term: termDetails,
          assignment: {
            startDate: startDate,
            endDate: endDate
          }
        };
      }
    }
    
    return {
      accessible: false,
      error: `No term assignment found for ${classSection} in Class_Terms sheet.`
    };
    
  } catch (error) {
    return {
      accessible: false,
      error: `Error checking term access: ${error.message}`
    };
  }
}

/**
 * Find and list all references to isCurrent/Is_Current in the code
 */
function findIsCurrentReferences() {
  console.log("Searching for isCurrent/Is_Current references...");
  
  // These are the functions that had isCurrent references:
  const functionsWithIsCurrent = [
    'getTermDetails',
    'getTermForClassSection', 
    'markAttendance',
    'getActiveClassTerms',
    'validateTermSetup',
    'testTermForClass',
    'listActiveAssignments',
    'testTermSystemUpdated',
    'getTermAccessStatus'
  ];
  
  console.log("Functions that need updating:");
  functionsWithIsCurrent.forEach(func => {
    console.log(`- ${func}`);
  });
  
  return functionsWithIsCurrent;
}


function testGetTermDetails() {
  const termId = "2025-2026";
  const result = getTermDetails(termId);
  console.log('getTermDetails("2025-2026") returns:');
  console.log('- termId:', result.termId);
  console.log('- termName:', result.termName);
  console.log('- startDate:', result.startDate, 'Type:', typeof result.startDate);
  console.log('- endDate:', result.endDate, 'Type:', typeof result.endDate);
  console.log('- isActive:', result.isActive);
  
  // Check what Date() creates from it
  const startAsDate = new Date(result.startDate);
  const endAsDate = new Date(result.endDate);
  console.log('\nAs Date objects:');
  console.log('- startAsDate:', startAsDate, 'Valid?', !isNaN(startAsDate.getTime()));
  console.log('- endAsDate:', endAsDate, 'Valid?', !isNaN(endAsDate.getTime()));
}

function testTermFinding() {
  console.log('=== Comprehensive Term Test ===');
  
  const testDate = new Date(); // Today's date
  console.log('Test Date:', testDate);
  console.log('Test Date ISO:', testDate.toISOString());
  
  // Test with a class you know should work
  const classSection = "10-A";
  
  console.log('\n--- Testing getTermForClassSection ---');
  const termResult = getTermForClassSection(classSection, testDate);
  console.log('Result:', termResult);
  
  console.log('\n--- Testing getTermAccessStatus ---');
  const accessResult = getTermAccessStatus(classSection, testDate);
  console.log('Access Result:', accessResult);
  
  console.log('\n--- Manual Date Check ---');
  
  // Manually check what's in Class_Terms
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_Terms');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const classCol = headers.indexOf('Class_Section');
  const termIdCol = headers.indexOf('Term_ID');
  const activeCol = headers.indexOf('Is_Active');
  const classStartCol = headers.indexOf('Start_Date');
  const classEndCol = headers.indexOf('End_Date');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][classCol] === classSection) {
      console.log(`Found ${classSection} at row ${i}`);
      console.log('Raw data:', data[i]);
      
      const classStartDate = new Date(data[i][classStartCol]);
      const classEndDate = new Date(data[i][classEndCol]);
      
      console.log('Class Start:', classStartDate);
      console.log('Class End:', classEndDate);
      console.log('Today >= Start?', testDate >= classStartDate);
      console.log('Today <= End?', testDate <= classEndDate);
      
      // Check Terms sheet dates
      const termId = data[i][termIdCol];
      const termDetails = getTermDetails(termId);
      if (termDetails) {
        const termStart = new Date(termDetails.startDate);
        const termEnd = new Date(termDetails.endDate);
        console.log('Term Start:', termStart);
        console.log('Term End:', termEnd);
        console.log('Today >= Term Start?', testDate >= termStart);
        console.log('Today <= Term End?', testDate <= termEnd);
      }
    }
  }
}


// Get months with attendance data for a class
function getMonthsForClass(classSection) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = ss.getSheetByName('Student_Attendance');
    if (!attendanceSheet) throw new Error('Student_Attendance sheet not found');
    
    const data = attendanceSheet.getDataRange().getValues();
    const headers = data[0];
    
    const classSectionCol = headers.indexOf('Class_Section');
    const dateCol = headers.indexOf('Date');
    
    if (classSectionCol === -1 || dateCol === -1) {
      return []; // Return empty array if columns not found
    }
    
    const monthsSet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][classSectionCol] === classSection) {
        const date = new Date(data[i][dateCol]);
        if (!isNaN(date.getTime())) {
          const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
          monthsSet.add(monthKey);
        }
      }
    }
    
    // Convert to array and sort descending
    const months = Array.from(monthsSet)
      .sort((a, b) => b.localeCompare(a))
      .map(monthStr => {
        const [year, month] = monthStr.split('-');
        return {
          year: parseInt(year),
          month: parseInt(month),
          display: `${new Date(year, month - 1).toLocaleString('default', { month: 'long' })} ${year}`
        };
      });
    
    return months;
    
  } catch (error) {
    console.error('Error in getMonthsForClass:', error);
    return []; // Return empty array on error
  }
}


function getBarcodeScannerHTML() {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
  <title>Barcode Scanner</title>
  <script src="https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.6.0/dist/confetti.browser.min.js"></script>
  <style>
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    
    body { 
      margin: 0; 
      background: #000; 
      font-family: 'Segoe UI', sans-serif; 
      overflow: hidden;
      position: fixed;
      width: 100%;
      height: 100%;
    }
    
    .scanner-container {
      position: relative;
      width: 100%;
      height: 100%;
      display: flex;
      justify-content: center;
      align-items: center;
    }
    
    #reader {
      width: 100%;
      height: 100%;
      position: absolute;
      top: 0;
      left: 0;
    }
    
    #reader video {
      width: 100%;
      height: 100%;
      object-fit: cover;
    }
    
    .scanner-overlay {
      position: absolute;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      pointer-events: none;
      z-index: 10;
    }
    
    .scan-frame {
      position: absolute;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      width: 280px;
      height: 280px;
      border: 2px solid #4CAF50;
      border-radius: 20px;
      box-shadow: 0 0 0 9999px rgba(0,0,0,0.6);
    }
    
    .scan-line {
      position: absolute;
      top: 0;
      left: 0;
      width: 100%;
      height: 2px;
      background: #4CAF50;
      animation: scan 2s linear infinite;
      border-radius: 2px;
    }
    
    @keyframes scan {
      0% { top: 0; }
      100% { top: 100%; }
    }
    
    .corner {
      position: absolute;
      width: 30px;
      height: 30px;
      border: 3px solid #4CAF50;
    }
    
    .corner-tl {
      top: -2px;
      left: -2px;
      border-right: none;
      border-bottom: none;
      border-radius: 20px 0 0 0;
    }
    
    .corner-tr {
      top: -2px;
      right: -2px;
      border-left: none;
      border-bottom: none;
      border-radius: 0 20px 0 0;
    }
    
    .corner-bl {
      bottom: -2px;
      left: -2px;
      border-right: none;
      border-top: none;
      border-radius: 0 0 0 20px;
    }
    
    .corner-br {
      bottom: -2px;
      right: -2px;
      border-left: none;
      border-top: none;
      border-radius: 0 0 20px 0;
    }
    
    .info-panel {
      position: absolute;
      bottom: 20px;
      left: 20px;
      right: 20px;
      background: rgba(0,0,0,0.85);
      color: white;
      padding: 15px;
      text-align: center;
      font-size: 14px;
      border-radius: 15px;
      z-index: 20;
    }
    
    .close-btn {
      position: absolute;
      top: 20px;
      right: 20px;
      background: #ff4444;
      color: white;
      border: none;
      width: 50px;
      height: 50px;
      border-radius: 50%;
      font-size: 24px;
      cursor: pointer;
      z-index: 30;
      box-shadow: 0 2px 10px rgba(0,0,0,0.3);
      display: flex;
      align-items: center;
      justify-content: center;
    }

    .camera-switch-btn {
  
}

.camera-switch-btn:active {
  transform: scale(0.95);
}
    
    .status {
      position: absolute;
      top: 20px;
      left: 20px;
      background: rgba(0,0,0,0.7);
      color: #4CAF50;
      padding: 8px 16px;
      border-radius: 20px;
      font-size: 14px;
      z-index: 30;
      font-weight: bold;
    }
    
    .counter-panel {
      position: absolute;
      top: 20px;
      right: 90px;
      background: rgba(0,0,0,0.7);
      color: white;
      padding: 8px 16px;
      border-radius: 20px;
      font-size: 14px;
      z-index: 30;
      font-weight: bold;
    }
    
    .counter-panel span {
      color: #4CAF50;
      font-size: 18px;
    }
    
    .last-scanned {
      position: absolute;
      bottom: 90px;
      left: 20px;
      right: 20px;
      background: rgba(0,0,0,0.85);
      color: #4CAF50;
      padding: 10px;
      text-align: center;
      font-size: 12px;
      font-family: monospace;
      word-break: break-all;
      border-radius: 10px;
      z-index: 20;
    }
    
    .last-scanned .label {
      color: #aaa;
      font-size: 10px;
      display: block;
      margin-bottom: 4px;
    }
    
    .status.valid {
      color: #4CAF50;
    }
    
    .status.invalid {
      color: #ff4444;
    }
  </style>
</head>
<body>
  <button class="close-btn" onclick="window.close()">✕</button>
<button class="camera-switch-btn" id="switchCameraBtn" style="position: absolute; bottom: 90px; right: 20px; background: #2196F3; color: white; border: none; width: 50px; height: 50px; border-radius: 50%; font-size: 24px; cursor: pointer; z-index: 30; box-shadow: 0 2px 10px rgba(0,0,0,0.3); display: flex; align-items: center; justify-content: center;">
  🔄
</button>
  <div class="scanner-container">
    <div id="reader"></div>
    <div class="scanner-overlay">
      <div class="scan-frame">
        <div class="corner corner-tl"></div>
        <div class="corner corner-tr"></div>
        <div class="corner corner-bl"></div>
        <div class="corner corner-br"></div>
        <div class="scan-line"></div>
      </div>
    </div>
  </div>
  
  <div class="status" id="scanStatus">Ready to scan</div>
  <div class="counter-panel" id="counterPanel">
    Scanned: <span id="scanCounter">0</span>
  </div>
  <div class="info-panel">
    <i class="fas fa-qrcode"></i> Position QR code within the frame
  </div>
<div class="last-scanned" id="lastScannedPanel">
  <span class="label">Last Scanned:</span>
  <span id="lastScannedValue">-</span>
</div>

  <script>
    let html5QrCode = null;
let isProcessing = false;
let lastScanTime = 0;
let lastScannedCode = null;
let validScanCount = 0;
let currentCameraId = null;
let allCameras = [];
    
    // TRACK SCANNED STUDENTS TO PREVENT DUPLICATES
    let scannedStudentIds = new Set();
    
    // Student data passed from parent
    let studentData = [];
    let dataReceived = false;

    function updateCounter() {
      document.getElementById('scanCounter').textContent = validScanCount;
    }

function updateLastScanned(barcode, isValid, studentName, studentId, isDuplicate) {
  const element = document.getElementById('lastScannedValue');
  if (isValid && studentName) {
    if (isDuplicate) {
      element.textContent = studentName + ' (SIRN: ' + studentId + ', Barcode: ' + barcode + ') - ALREADY SCANNED';
      element.style.color = '#ff9800';
    } else {
      element.textContent = studentName + ' (SIRN: ' + studentId + ', Barcode: ' + barcode + ')';
      element.style.color = '#4CAF50';
    }
  } else {
    element.textContent = 'INVALID: ' + barcode;
    element.style.color = '#ff4444';
  }
}

    function updateStatus(message, isValid = null) {
      const statusEl = document.getElementById('scanStatus');
      statusEl.textContent = message;
      if (isValid === true) {
        statusEl.style.color = '#4CAF50';
        statusEl.classList.add('valid');
        statusEl.classList.remove('invalid');
      } else if (isValid === false) {
        statusEl.style.color = '#ff4444';
        statusEl.classList.add('invalid');
        statusEl.classList.remove('valid');
      } else {
        statusEl.style.color = '#4CAF50';
      }
    }

    // Request student data from parent
    function requestStudentData() {
      if (window.opener && !dataReceived) {
        console.log('Requesting student data from parent...');
        window.opener.postMessage({
          type: 'REQUEST_STUDENT_DATA'
        }, '*');
        
        // Set timeout to show error if no response
        setTimeout(() => {
          if (!dataReceived) {
            console.warn('No student data received from parent');
            updateStatus('Waiting for data...', false);
          }
        }, 3000);
      }
    }



  async function startScanner(cameraId) {
  try {
    // Stop existing scanner
    if (html5QrCode && html5QrCode.isScanning) {
      await html5QrCode.stop();
      await html5QrCode.clear();
    }
    
    const readerElement = document.getElementById("reader");
    readerElement.innerHTML = "";
    
    html5QrCode = new Html5Qrcode("reader");
    
    await html5QrCode.start(
      { deviceId: { exact: cameraId } },
      {
        fps: 15,
        qrbox: { width: 280, height: 280 },
        aspectRatio: 1.0,
        disableFlip: false
      },
      onScanSuccess,
      onScanError
    );
    
    currentCameraId = cameraId;
    
    // Update button text to show camera type
    const currentCamera = allCameras.find(c => c.id === cameraId);
    const cameraLabel = currentCamera?.label.toLowerCase() || '';
    const cameraType = cameraLabel.includes('back') || cameraLabel.includes('rear') ? '📷' : 
                      (cameraLabel.includes('front') ? '🤳' : '🔄');
    const switchBtn = document.getElementById('switchCameraBtn');
    if (switchBtn && allCameras.length > 1) {
      switchBtn.innerHTML = cameraType;
      switchBtn.style.opacity = '1';
    } else if (switchBtn) {
      switchBtn.style.opacity = '0.5';
    }
    
    updateStatus('Scanner active - Ready', true);
    
  } catch (err) {
    console.error('Scanner start error:', err);
    updateStatus('Camera error: ' + err.message, false);
  }
}  

 async function initScanner() {
  try {
    const cameras = await Html5Qrcode.getCameras();
    allCameras = cameras;
    
    if (!cameras || cameras.length === 0) {
      updateStatus('No camera found', false);
      return;
    }
    
    // Find back camera
    let backCameraId = null;
    for (const camera of cameras) {
      const label = camera.label.toLowerCase();
      if (label.includes('back') || label.includes('rear') || label.includes('environment')) {
        backCameraId = camera.id;
        break;
      }
    }
    
    // Try to get saved preference
    const savedCameraId = localStorage.getItem('preferredCameraId');
    const savedCameraExists = cameras.some(c => c.id === savedCameraId);
    
    let cameraToUse = null;
    if (savedCameraExists) {
      cameraToUse = savedCameraId;
    } else if (backCameraId) {
      cameraToUse = backCameraId;
    } else {
      cameraToUse = cameras[0].id;
    }
    
    await startScanner(cameraToUse);
    
  } catch (err) {
    console.error('Camera init error:', err);
    updateStatus('Camera error: ' + err.message, false);
  }
}

async function switchCamera() {
  if (allCameras.length <= 1) {
    updateStatus('Only one camera available', false);
    setTimeout(() => updateStatus('Scanner active - Ready', true), 1500);
    return;
  }
  
  updateStatus('Switching camera...', false);
  
  // Find next camera
  const currentIndex = allCameras.findIndex(c => c.id === currentCameraId);
  const nextIndex = (currentIndex + 1) % allCameras.length;
  const nextCameraId = allCameras[nextIndex].id;
  
  // Save preference
  localStorage.setItem('preferredCameraId', nextCameraId);
  
  await startScanner(nextCameraId);
}

 function onScanSuccess(decodedText, decodedResult) {
  const now = Date.now();
  
  // Prevent rapid successive scans
  if (isProcessing) return;
  if (now - lastScanTime < 1500) return;
  if (lastScannedCode === decodedText && now - lastScanTime < 3000) return;
  
  lastScanTime = now;
  lastScannedCode = decodedText;
  isProcessing = true;
  
  // Wait for student data if not received yet
  if (!dataReceived || studentData.length === 0) {
    updateStatus('Loading student data...', false);
    setTimeout(function() {
      isProcessing = false;
    }, 1000);
    return;
  }
  
  // VALIDATE BARCODE LOCALLY
  var student = null;
  for (var i = 0; i < studentData.length; i++) {
    if (studentData[i].Barcode_ID === decodedText) {
      student = studentData[i];
      break;
    }
  }
  var isValid = (student !== null);
  
  if (isValid) {
    // CHECK ALL CONDITIONS:
    var isInDatabase = (student.attendanceStatus === 'Present');
    var isInUI = (student.isUIPresent === true);
    var isInSession = scannedStudentIds.has(student.Std_ID);
    
    // If already marked ANYWHERE, reject
    if (isInDatabase || isInUI || isInSession) {
      var message = '';
      if (isInDatabase) {
        message = 'ALREADY SAVED: ' + student.Student_Name;
      } else if (isInUI) {
        message = 'ALREADY MARKED: ' + student.Student_Name + ' in list!';
      } else {
        message = 'DUPLICATE: ' + student.Student_Name + ' already scanned!';
      }
      updateLastScanned(decodedText, true, student.Student_Name, student.Std_ID, true);
      updateStatus(message, false);
      
      // Send duplicate notification to parent
      if (window.opener) {
        window.opener.postMessage({
          type: 'BARCODE_SCAN',
          barcode: decodedText,
          studentName: student.Student_Name,
          studentId: student.Std_ID,
          isValid: true,
          isDuplicate: true,
          timestamp: now
        }, '*');
      }
    } else {
      // NEW VALID STUDENT - INCREMENT COUNTER
      validScanCount++;
      scannedStudentIds.add(student.Std_ID);
      updateCounter();
      updateLastScanned(decodedText, true, student.Student_Name, student.Std_ID, false);
      updateStatus('NEW: ' + student.Student_Name, true);
      
      // TRIGGER CONFETTI IN THE POPUP ITSELF
      triggerPopupConfetti();
      
      // Send scan result to parent
      if (window.opener) {
        window.opener.postMessage({
          type: 'BARCODE_SCAN',
          barcode: decodedText,
          studentName: student.Student_Name,
          studentId: student.Std_ID,
          isValid: true,
          isDuplicate: false,
          timestamp: now
        }, '*');
      }
    }
  } else {
    // Invalid barcode - DO NOT increment counter
    updateLastScanned(decodedText, false, null, null, false);
    updateStatus('INVALID BARCODE: ' + decodedText.substring(0, 20) + '...', false);
    
    // Send invalid notification to parent
    if (window.opener) {
      window.opener.postMessage({
        type: 'BARCODE_SCAN',
        barcode: decodedText,
        isValid: false,
        error: 'Barcode not found in system',
        timestamp: now
      }, '*');
    }
  }
  
  setTimeout(function() {
    isProcessing = false;
    updateStatus('Ready to scan', true);
  }, 1000);
}

// Add this function inside the popup script
function triggerPopupConfetti() {
  // Simple canvas confetti for the popup
  const canvas = document.createElement('canvas');
  canvas.style.position = 'fixed';
  canvas.style.top = '0';
  canvas.style.left = '0';
  canvas.style.width = '100%';
  canvas.style.height = '100%';
  canvas.style.pointerEvents = 'none';
  canvas.style.zIndex = '9999';
  document.body.appendChild(canvas);
  
  const myConfetti = confetti.create(canvas, {
    resize: true,
    useWorker: true
  });
  
  myConfetti({
    particleCount: 100,
    spread: 70,
    origin: { y: 0.6 },
    startVelocity: 20,
    colors: ['#4CAF50', '#2196F3', '#FF9800', '#E91E63', '#9C27B0']
  });
  
  // Remove canvas after animation
  setTimeout(() => {
    canvas.remove();
  }, 2000);
}

    function onScanError(errorMessage) {
      // Silent fail - scanner continues
      console.warn('Scan error:', errorMessage);
    }

// Listen for student data from parent
window.addEventListener('message', function(event) {
  if (event.data && event.data.type === 'STUDENT_DATA_RESPONSE') {
    studentData = event.data.students;
    dataReceived = true;
    
    // Pre-fill scannedStudentIds with students already marked (Database OR UI)
    let preMarkedCount = 0;
    studentData.forEach(student => {
      // Check BOTH: Database status AND UI button state
      if (student.attendanceStatus === 'Present' || student.isUIPresent === true) {
        scannedStudentIds.add(student.Std_ID);
        preMarkedCount++;
        console.log('Pre-marked: ' + student.Student_Name + 
                    (student.attendanceStatus === 'Present' ? ' (Database)' : ' (UI)'));
      }
    });
    
    // Set counter to number of already marked students
    validScanCount = preMarkedCount;
    updateCounter();
    
    console.log('Loaded ' + studentData.length + ' students, ' + preMarkedCount + ' already marked');
    updateStatus('Ready - ' + studentData.length + ' students', true);
  }
});

// Request data and start scanner
requestStudentData();
initScanner();

  // Add camera switch button event listener
    document.getElementById('switchCameraBtn').addEventListener('click', function(e) {
      e.preventDefault();
      if (allCameras.length > 1) {
        switchCamera();
      }
    });

    // Handle cleanup on page unload
    window.addEventListener('beforeunload', async () => {
      if (html5QrCode && html5QrCode.isScanning) {
        try {
          await html5QrCode.stop();
        } catch(e) {}
      }
    });
  </script>
</body>
</html>
  `;
}


function getRecName(recId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RECNames');
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const recCodeCol = headers.indexOf('REC Code');
    const recNameCol = headers.indexOf('REC Name');
    
    if (recCodeCol === -1 || recNameCol === -1) return null;
    
    const searchId = String(recId).trim();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][recCodeCol]).trim() === searchId) {
        return data[i][recNameCol];
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}