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
      if (typeCol !== -1 && String(data[i][typeCol] || '').trim().toLowerCase() === 'episodic') continue; // sirf normal attendance
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
 * @param {string} callerUsername - The admin making the change
 * @returns {boolean} - Success status (false when the user does not exist)
 */
function updateUserStatus(username, newStatus, callerUsername) {
  // v5.6: only an administrator may change a user's status, and the new value
  // must be Active or Inactive. (No HTML caller — direct calls only.)
  // v5.8 (F5, corrected 19 Nov 2026): school scoping — NO cross-school access
  // for anyone, superadmin included. The caller may change status only for
  // users at a school where the caller themself holds an admin/superadmin
  // Users row (getAdminCallerScope_); the row updated is the target's row at
  // one of those schools, not blindly the first row with the username. A
  // target with a blank REC_ID matches (v5.6 grandfathering). To work at
  // another school, the caller must be assigned there and log in with that
  // school's credentials.
  requireAdminCaller_(callerUsername);
  if (newStatus !== 'Active' && newStatus !== 'Inactive') throw new Error('Status must be Active or Inactive.');
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
    const recIdCol = headers.indexOf('REC_ID');
    
    // v5.8 (F5, corrected): pick the target row the caller is allowed to
    // touch — the first row whose REC_ID is one of the caller's schools, or
    // whose REC_ID is blank (v5.6 grandfathering). A superadmin gets no
    // exception: without a row at the target's school they are refused like
    // any other admin. A target with no row at the caller's schools is refused.
    const scope = getAdminCallerScope_(callerUsername);
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameCol] !== username) continue;
      const targetRec = recIdCol !== -1 ? String(data[i][recIdCol] || '').trim() : '';
      if (!targetRec || scope.recIds.indexOf(targetRec) !== -1) { targetRow = i; break; }
    }
    if (targetRow === -1) {
      const found = data.slice(1).some(function (row) { return row[usernameCol] === username; });
      if (!found) return false; // User not found (unchanged)
      throw new Error('You can only change the status of users at your own school.');
    }
    const i = targetRow;
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
    
    // v5.7: the row is picked by username AND password together, so the same
    // username can live at two schools — the password prefix picks the school.
    // (Before, the first row with the username always won and the other
    // school's row was never reached.)
    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameCol] !== username) continue;
      userFound = true;
      if (data[i][passwordCol] !== password) continue; // another row may match
      userRow = i;
      userStatus = data[i][statusCol] || 'Active';
      if (userStatus === 'Inactive') {
        // Force cache update
        cache.remove('userDataCache');
        initializeUserCache();
        throw new Error('User account is inactive. Please contact administrator.');
      }
      // Sirf UserType 'teacher' wale users hi login kar sakte hain
      const userType = String(data[i][typeCol] || '').trim();
      if (userType.toLowerCase() !== 'teacher') {
        return { blocked: true, message: 'This account is not authorized to use the attendance system.' };
      }
      userData = {
            password: data[i][passwordCol],
        type: data[i][typeCol] || 'teacher',
        teacherName: data[i][teacherNameCol] || data[i][usernameCol],
        status: userStatus
      };
      break;
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

  // v5.9: sheet-structure self-heal, once per ~6h, best-effort only.
  try { maybeAutoStructureCheck_('Attendance', MARKING_STRUCTURE_SCHEMA_); } catch (e) {}

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

function getUserSchedule(username, schoolRecId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) throw new Error('Users sheet not found');

  const data = userSheet.getDataRange().getValues();
  const headers = data[0];
  const usernameCol = headers.indexOf('Username');
  const daysCol = headers.indexOf('AllowedDays');
  const startTimeCol = headers.indexOf('StartTime');
  const endTimeCol = headers.indexOf('EndTime');
  const dayScheduleCol = headers.indexOf('DaySchedule');

  if (usernameCol === -1 || daysCol === -1 || startTimeCol === -1 || endTimeCol === -1) {
    throw new Error('Required columns not found in Users sheet');
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][usernameCol] === username) {
      // v5.7: same teacher at two schools — use the row for the school they
      // logged into (a blank REC_ID on the row matches any school).
      if (schoolRecId) {
        const recCol = headers.indexOf('REC_ID');
        if (recCol !== -1) {
          const rowRec = String(data[i][recCol] || '').trim();
          if (rowRec && rowRec !== String(schoolRecId).trim()) continue;
        }
      }
      const allowedDays = data[i][daysCol] ? data[i][daysCol].split(',').map(day => day.trim()) : [];
      const startTime = data[i][startTimeCol] || '';
      const endTime = data[i][endTimeCol] || '';
      let daySchedule = {};
      if (dayScheduleCol !== -1 && data[i][dayScheduleCol]) {
        try {
          daySchedule = JSON.parse(data[i][dayScheduleCol]) || {};
        } catch (e) {
          daySchedule = {};
        }
      }
      return { allowedDays, startTime, endTime, daySchedule };
    }
  }
  throw new Error('No schedule assigned for this user');
}

function isWithinSchedule(username, schoolRecId) {
  const schedule = getUserSchedule(username, schoolRecId);
  const now = new Date();
  const currentDay = now.toLocaleString('en-US', { weekday: 'long', timeZone: 'Asia/Karachi' });
  const currentTime = now.toLocaleString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false, timeZone: 'Asia/Karachi' });

  if (!schedule.allowedDays.includes(currentDay)) {
    throw new Error(`Attendance marking not allowed on ${currentDay}. Allowed days: ${schedule.allowedDays.join(', ')}`);
  }

  const dayOverride = schedule.daySchedule && schedule.daySchedule[currentDay];
  const effectiveStart = (dayOverride && dayOverride.start && dayOverride.end) ? dayOverride.start : schedule.startTime;
  const effectiveEnd = (dayOverride && dayOverride.start && dayOverride.end) ? dayOverride.end : schedule.endTime;

  if (!effectiveStart || !effectiveEnd) {
    throw new Error('Start or end time not defined for this user');
  }

  const [startHour, startMinute] = effectiveStart.split(':').map(Number);
  const [endHour, endMinute] = effectiveEnd.split(':').map(Number);
  const [currentHour, currentMinute] = currentTime.split(':').map(Number);

  const startTimeInMinutes = startHour * 60 + startMinute;
  const endTimeInMinutes = endHour * 60 + endMinute;
  const currentTimeInMinutes = currentHour * 60 + currentMinute;

  if (currentTimeInMinutes < startTimeInMinutes || currentTimeInMinutes > endTimeInMinutes) {
    throw new Error(`Attendance marking only allowed between ${effectiveStart} and ${effectiveEnd}`);
  }

  return true;
}



// Summary report (v5.1): computed in the app from Student_Attendance +
// Daily_Hours_Setup for a picked date (default TODAY, Asia/Karachi). The old
// version read the formula-backed Summary sheet, whose J1 lookup errored with
// #N/A on any day before the first marking; this version has no Summary-sheet
// dependency.
function getSummaryData(username, classSection, dateStr, schoolRecId) {
  try {
    // v5.8 (O5): the logged-in school scopes the assignment list.
    const assignedSections = getAssignedClassSections(username, schoolRecId);
    Logger.log('Assigned Sections for ' + username + ': ' + assignedSections);
    if (!assignedSections || assignedSections.length === 0) {
      Logger.log('No assigned sections for user: ' + username);
      return [];
    }

    // A direct call may ask for one class only: refuse it unless assigned
    // (same rule as the Daily/Monthly/Day Wise reports).
    let sections = assignedSections.slice();
    if (classSection) {
      if (!isClassSectionAssigned_(assignedSections, classSection)) {
        throw new Error('You are not assigned to class ' + classSection + '.');
      }
      sections = [classSection];
    }

    // Rows in Class_Sections order, as the old sheet listed them.
    const order = getClassSections();
    sections.sort(function (a, b) {
      const ia = order.indexOf(a), ib = order.indexOf(b);
      return (ia === -1 ? 9999 : ia) - (ib === -1 ? 9999 : ib);
    });

    // v5.1: optional summary date (yyyy-MM-dd). No date = today (Asia/Karachi).
    // A future date has no attendance to summarise, so it is refused.
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Karachi', 'yyyy-MM-dd');
    let summaryDate = todayStr;
    if (dateStr !== undefined && dateStr !== null && String(dateStr).trim() !== '') {
      const picked = String(dateStr).trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(picked)) {
        throw new Error('Invalid summary date: ' + picked + '. Use yyyy-MM-dd.');
      }
      const parts = picked.split('-').map(Number);
      const d = new Date(parts[0], parts[1] - 1, parts[2]);
      if (d.getFullYear() !== parts[0] || d.getMonth() !== parts[1] - 1 || d.getDate() !== parts[2]) {
        throw new Error('Invalid summary date: ' + picked + '.');
      }
      if (picked > todayStr) {
        throw new Error('Future dates are not allowed for the summary.');
      }
      summaryDate = picked;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const attSheet = ss.getSheetByName('Student_Attendance');
    if (!attSheet) throw new Error('Student_Attendance sheet not found');
    const attData = attSheet.getDataRange().getValues();
    const attHeaders = attData[0].map(function (h) { return String(h).trim(); });
    const dateCol = attHeaders.indexOf('Date');
    const statusCol = attHeaders.indexOf('Status');
    const clsSecCol = attHeaders.indexOf('Class_Section');
    const clsCol = attHeaders.indexOf('Student_Class');
    const secCol = attHeaders.indexOf('Student_Section');
    if (dateCol === -1 || statusCol === -1 || (clsSecCol === -1 && (clsCol === -1 || secCol === -1))) {
      throw new Error('Required columns (Date, Status, Class_Section) not found in Student_Attendance');
    }

    // Count the picked day's rows per class-section, by Status. Like the old sheet
    // COUNTIFS, every row counts (no Attendance_Type filter — episodic rows
    // count too) and duplicate rows count twice.
    const counts = {};
    sections.forEach(function (s) { counts[s] = { present: 0, absent: 0, leave: 0 }; });
    for (let i = 1; i < attData.length; i++) {
      const row = attData[i];
      if (normalizeYMD_(row[dateCol]) !== summaryDate) continue;
      let cs = clsSecCol !== -1 ? String(row[clsSecCol] || '').trim() : '';
      if (!cs && clsCol !== -1 && secCol !== -1) {
        const c = String(row[clsCol] || '').trim();
        const s = String(row[secCol] || '').trim();
        if (c && s) cs = c + '-' + s;
      }
      const bucket = cs ? counts[cs] : null;
      if (!bucket) continue;
      const st = String(row[statusCol] || '').toLowerCase().trim();
      if (st === 'present') bucket.present++;
      else if (st === 'absent') bucket.absent++;
      else if (st === 'leave') bucket.leave++;
      // unknown statuses (e.g. 'Not Applicable') are ignored in the total,
      // same as the old sheet version
    }

    // Category + Hours are the day's SETUP values from Daily_Hours_Setup
    // (same lookup the sheet did): normal-type row first, fallback any type.
    return sections.map(function (cs) {
      const c = counts[cs];
      let category = '';
      let hours = '';
      try {
        // v5.8 (O3): the day's setup hours from this school's rows only.
        let wh = checkDailyHours(summaryDate, cs, 'normal', schoolRecId);
        if (!wh || !wh.exists) wh = checkDailyHours(summaryDate, cs, null, schoolRecId);
        if (wh && wh.exists) {
          category = wh.category == null ? '' : wh.category;
          hours = wh.hours == null ? '' : wh.hours;
        }
      } catch (e) {
        category = '';
        hours = ''; // no setup sheet/row: summary still works, blank setup
      }
      return {
        classSection: cs,
        hours: hours,
        category: category,
        present: c.present,
        absent: c.absent,
        leaves: c.leave,
        total: c.present + c.absent + c.leave
      };
    });
  } catch (e) {
    Logger.log('Error in getSummaryData: ' + e.message);
    throw e;
  }
}

function getAssignedClassSections(username, schoolRecId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teacher_Assignments');
    if (!sheet) throw new Error('Teacher_Assignments sheet not found');

    const data = sheet.getDataRange().getValues();
    const recCol = data[0].indexOf('REC_ID');
    const wantRec = String(schoolRecId || '').trim();
    // v5.8 (O5): no school, no classes. The old first-row-wins fallback
    // silently offered another school's classes (F4's server side). Callers
    // must pass the logged-in school. An empty list is what every caller
    // already handles: requireAssignedClass_ treats it as "no constraint"
    // (same as a teacher with no assignment rows), getSummaryData returns
    // [], and the report guards refuse with "not assigned".
    if (!wantRec) return [];
    const validSections = getClassSections();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        // v5.6: only this school's rows count (blank REC_ID = any school).
        if (recCol !== -1) {
          const rowRec = String(data[i][recCol] || '').trim();
          if (rowRec && rowRec !== wantRec) continue;
        }
        const sections = data[i][1] ? data[i][1].split(',').filter(s => s.trim()) : [];
        const filtered = sections.filter(section => validSections.includes(section));
        filtered.forEach(function (s) { if (out.indexOf(s) === -1) out.push(s); });
      }
    }
    return out;
  } catch (e) {
    Logger.log('Error in getAssignedClassSections: ' + e.message);
    throw e;
  }
}
// ============ PREVIOUS DATE PERMISSION HELPERS ============

/** Previous_Date_Permissions ki rows — naye columns header se padhta hai, position se farq nahi */
function getPrevPermRows_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Previous_Date_Permissions');
  if (!sheet) throw new Error('Previous_Date_Permissions sheet not found');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim());
  const expCol  = headers.indexOf('Expires_At');
  const grantedCol = headers.indexOf('Granted_At');
  const typeCol = headers.indexOf('Attendance_Type');
  const clsCol  = headers.indexOf('Class_Section');
  const recCol  = headers.indexOf('REC_ID');

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      username:  String(data[i][0]).trim(),
      dateRaw:   data[i][1],
      password:  String(data[i][2]).trim(),
      expiresAt: expCol  !== -1 ? data[i][expCol]  : '',
      grantedAt: grantedCol !== -1 ? data[i][grantedCol] : '',
      type:      typeCol !== -1 ? String(data[i][typeCol] || '').trim() : '',
      classSection: clsCol !== -1 ? String(data[i][clsCol] || '').trim() : '',
      recId:   recCol !== -1 ? String(data[i][recCol] || '').trim() : ''
    });
  }
  return rows;
}

/** Effective expiry nikalo: Expires_At ho to wohi; warna Granted_At + 24 ghante; dono na hon to null. */
function prevPermEffectiveExpiry_(row) {
  const exp = row.expiresAt instanceof Date ? row.expiresAt : new Date(row.expiresAt);
  if (row.expiresAt && !isNaN(exp.getTime())) return exp;
  const g = row.grantedAt instanceof Date ? row.grantedAt : new Date(row.grantedAt);
  if (row.grantedAt && !isNaN(g.getTime())) return new Date(g.getTime() + 24 * 60 * 60 * 1000);
  return null;
}

/** 24-hour window valid hai? Effective expiry (Expires_At, warna Granted_At + 24h) abhi future mein ho. Dono khaali/kharab hon to row live nahi. */
function prevPermNotExpired_(row) {
  const exp = prevPermEffectiveExpiry_(row);
  if (!exp) return false;
  return new Date().getTime() <= exp.getTime();
}

/** Attendance_Type match karta hai? ('' ya 'Both' = dono allowed) */
function prevPermTypeMatches_(row, wantedType) {
  const want = String(wantedType || '').trim().toLowerCase();
  if (!want) return true;
  const have = String(row.type || '').trim().toLowerCase();
  if (!have || have === 'both') return true;
  return have === want;
}

/** Class match karta hai? (row ki Class_Section khaali = sab classes allowed) */
function prevPermClassMatches_(row, classSection) {
  const want = String(classSection || '').trim();
  if (!want) return true;
  const have = String(row.classSection || '').trim();
  if (!have) return true;
  return have === want;
}

/** '1102-A78941' jaisi value se REC prefix hatao — sirf asal password compare karo */
function stripPrevPermPrefix_(value) {
  return String(value || '').trim().replace(/^\d+-/, '');
}

/** Effective expiry ko readable string banao ('16 Nov 2026, 01:13'); khaali/kharab = '' */
function formatPrevPermExpiry_(expiresAt) {
  if (!expiresAt) return '';
  const d = expiresAt instanceof Date ? expiresAt : new Date(expiresAt);
  if (isNaN(d.getTime())) return '';
  try {
    return Utilities.formatDate(d, 'Asia/Karachi', 'dd MMM yyyy, HH:mm');
  } catch (e) {
    return '';
  }
}

const PREV_PERM_EXPIRED_MSG =
  'Your previous-date access has expired (24-hour limit crossed). Please request a new permission from the administrator.';

// v5.8 (O4): a permission row that was never live (no Granted_At and no
// Expires_At) is NOT expired — it gets the no-permission message instead.
const PREV_PERM_NO_PERMISSION_MSG =
  'You do not have permission to mark attendance for this previous date. Please request a new permission from the administrator.';

function verifyPreviousPassword(username, password, attendanceType, date, classSection) {
  try {
    const rows = getPrevPermRows_();
    let expired = false;
    let neverLive = false;
    const typed = stripPrevPermPrefix_(password);

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (r.username !== username) continue;
      if (stripPrevPermPrefix_(r.password) !== typed) continue;
      if (!prevPermTypeMatches_(r, attendanceType)) continue;
      if (date) {
        const sheetDate = new Date(r.dateRaw);
        if (isNaN(sheetDate.getTime())) continue;
        if (sheetDate.toLocaleDateString('en-CA') !== new Date(date).toLocaleDateString('en-CA')) continue;
      }
      if (!prevPermClassMatches_(r, classSection)) continue;
      // v5.8 (O4): only a row that HAD an expiry which has passed counts as
      // expired. A row that was never live (no dates at all) is refused with
      // the no-permission message, not the expired message.
      if (!prevPermNotExpired_(r)) {
        if (prevPermEffectiveExpiry_(r)) expired = true;
        else neverLive = true;
        continue;
      }
      return true;
    }

    if (expired) throw new Error(PREV_PERM_EXPIRED_MSG);
    if (neverLive) throw new Error(PREV_PERM_NO_PERMISSION_MSG);
    return false;
  } catch (e) {
    Logger.log(`Error in verifyPreviousPassword: ${e.message}`);
    throw e;
  }
}

function checkPreviousDatePermission(username, date, attendanceType, classSection, schoolRecId) {
  try {
    const rows = getPrevPermRows_();
    const dateString = new Date(date).toLocaleDateString('en-CA'); // YYYY-MM-DD
    let expired = false;

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (r.username !== username) continue;
      // v5.6: a permission granted at another school does not unlock this one
      // (blank REC_ID on the row = any school, for older rows).
      if (!rowRecMatches_(r.recId, schoolRecId)) continue;

      const sheetDate = new Date(r.dateRaw);
      if (isNaN(sheetDate.getTime())) {
        Logger.log(`Invalid date in Previous_Date_Permissions for ${username}: ${r.dateRaw}`);
        continue;
      }
      if (sheetDate.toLocaleDateString('en-CA') !== dateString) continue;
      if (!prevPermTypeMatches_(r, attendanceType)) continue;
      if (!prevPermClassMatches_(r, classSection)) continue;
      // v5.8 (O4): only a row that HAD an expiry which has passed counts as
      // expired. A row that was never live (no dates at all) does not set
      // expired, so this returns false and the caller throws its own
      // no-permission message ("You do not have permission to ... for this
      // previous date."). Refusal and zero writes are unchanged.
      if (!prevPermNotExpired_(r)) {
        if (prevPermEffectiveExpiry_(r)) expired = true;
        continue;
      }

      return true;
    }

    if (expired) throw new Error(PREV_PERM_EXPIRED_MSG);
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

function getAllowedPreviousDates(username, attendanceType, classSection) {
  const cacheKey = username + '|' + String(attendanceType || '') + '|' + String(classSection || '');
  if (allowedDatesCache && allowedDatesCache.key === cacheKey) return allowedDatesCache.dates;

  const rows = getPrevPermRows_();
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Karachi', 'yyyy-MM-dd');
  const allowedDates = [];
  const seen = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r.username !== username || !r.dateRaw) continue;
    if (!prevPermTypeMatches_(r, attendanceType)) continue;   // Normal / Episodic / Both
    if (!prevPermNotExpired_(r)) continue;                    // 24 ghante guzar gaye
    if (!prevPermClassMatches_(r, classSection)) continue;    // class scope (khaali row = sab classes)

    const d = new Date(r.dateRaw);
    if (isNaN(d.getTime())) {
      Logger.log(`Invalid date in Previous_Date_Permissions row for ${username}: ${r.dateRaw}`);
      continue;
    }
    const dateStr = d.toLocaleDateString('en-CA');
    if (dateStr > todayStr) continue;                          // future dates kabhi offer nahi
    if (seen[dateStr]) continue;
    seen[dateStr] = true;
    allowedDates.push({ date: dateStr, expiresAt: formatPrevPermExpiry_(prevPermEffectiveExpiry_(r)) });
  }

  allowedDatesCache = { key: cacheKey, dates: allowedDates };
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

function getStudentList(classSection, date, schoolRecId) {
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
    const recIdCol = studentHeaders.indexOf('REC_ID');

    for (let i = 1; i < studentData.length; i++) {
      const status = studentData[i][statusCol] ? studentData[i][statusCol].toString().toLowerCase() : '';
      const dateOfJoining = parseSheetDate(studentData[i][dojCol]);
      const targetDate = new Date(date);
      
      // Check if student should be included based on joining date
      const shouldIncludeByDate = !dateOfJoining || dateOfJoining <= targetDate;
      // v5.6: school-scoped roster (blank REC_ID on the student row = any school).
      if (!rowRecMatches_(recIdCol !== -1 ? studentData[i][recIdCol] : '', schoolRecId)) continue;
      
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
          attendanceStatus: 'Not marked',
          Hours: null,
          Category: null
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
          if (String(attendanceType || '').trim().toLowerCase() !== 'episodic') {
            studentsMap.get(stdId).attendanceStatus = attendanceData[j][attendanceHeaderMap.get('Status')] || 'Not marked';
            const hrsIdx = attendanceHeaderMap.get('Hours');
            if (hrsIdx !== undefined) {
              const hrsVal = attendanceData[j][hrsIdx];
              studentsMap.get(stdId).Hours = (hrsVal === '' || hrsVal === null || hrsVal === undefined) ? null : hrsVal;
                          const catIdx = attendanceHeaderMap.get('Category');
            if (catIdx !== undefined) studentsMap.get(stdId).Category = attendanceData[j][catIdx] || null;
            }
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
          attendanceStatus: 'Not Applicable',  // Default for episodic
          Hours: null
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
    const hoursColAtt = headers.indexOf('Hours');
    
    const targetDateStr = new Date(date).toLocaleDateString('en-CA');
    
    // Create a map of studentId -> {status, hours} ONLY for EPISODIC records
    const episodicStatusMap = new Map();
    
    for (let i = 1; i < attendanceData.length; i++) {
      const row = attendanceData[i];
      const rowDate = row[dateColAtt] ? new Date(row[dateColAtt]).toLocaleDateString('en-CA') : null;
      const rowType = row[typeColAtt];
      const rowEvent = row[eventColAtt];
      const rowClass = row[classColAtt];
      
      // ✅ CRITICAL: Only match EPISODIC type!
      if (rowDate === targetDateStr &&
          String(rowType || '').trim().toLowerCase() === 'episodic' &&           // ← ONLY Episodic records
          rowEvent === eventName &&
          rowClass === classSection) {
        const studentId = String(row[stdIdColAtt]);
        episodicStatusMap.set(studentId, { status: row[statusColAtt], hours: hoursColAtt !== -1 ? row[hoursColAtt] : null });
        Logger.log(`Found EPISODIC record: Student ${studentId} → ${row[statusColAtt]}`);
      }
    }
    
    // Update student status based ONLY on episodic records
    for (const student of studentsMap.values()) {
      const studentId = String(student.Std_ID);
      if (episodicStatusMap.has(studentId)) {
        const rec = episodicStatusMap.get(studentId);
        student.attendanceStatus = rec.status;
        student.Hours = (rec.hours === '' || rec.hours === null || rec.hours === undefined) ? null : rec.hours;
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

function markAttendance(attendanceData, class_section, date, username, isBarcode = false, recId = '', category = '', hours = '') {
   const startTime = Date.now();

// ✅ SERVER-SIDE SCHEDULE CHECK for Current Date Only
  const todayStr_ = Utilities.formatDate(new Date(), 'Asia/Karachi', 'yyyy-MM-dd');

  // v5.6: caller verification (unknown / non-teacher / Inactive refused). The
  // school comes from the caller's Users row; the sent recId only disambiguates a username that exists at two schools.
  const caller = requireTeacherCaller_(username, recId);
  const schoolRecId = caller.recId || String(recId || '').trim();
  requireAssignedClass_(username, class_section, schoolRecId);
  if (!isRealYMD_(date)) throw new Error('Invalid date: ' + date + '. Use a real yyyy-MM-dd date.');
  if (date > todayStr_) throw new Error('Future dates cannot be marked.');
  // v5.6: past-date permission is checked BEFORE any write, so a refused
  // marking leaves no setup row behind.
  if (date < todayStr_) {
    const allowed = checkPreviousDatePermission(username, date, 'Normal', class_section, schoolRecId);
    if (!allowed) {
      throw new Error('You do not have permission to mark attendance for this previous date. Please request a new permission from the administrator.');
    }
  }


  if (new Date(date).toLocaleDateString('en-CA') === todayStr_) {
    isWithinSchedule(username, schoolRecId); // Will throw error if outside allowed time
  }

  Logger.log(`Starting markAttendance with date: ${date}, classSection: ${class_section}, isBarcode: ${isBarcode}, records: ${attendanceData.length}`);

  try {
    // ========================================================
    // TERM CHECK REMOVED (per request) - term info still fetched (non-blocking) for record-keeping only
    // ========================================================
    const term = getTermForClassSection(class_section, date);
    if (term) {
      Logger.log(`Using term: ${term.termName} (${term.termId}) from ${term.startDate} to ${term.endDate}`);
    } else {
      Logger.log(`No active term found for ${class_section} on ${date} - proceeding anyway (term check disabled).`);
    }
    // ========================================================
    
    // -------------------------------------------------
    // NEW: Check that hours & category are defined for the day
    // -------------------------------------------------
let hoursCheck = checkDailyHours(date, class_section, 'normal', schoolRecId);
console.log('Hours check result:', hoursCheck);  // Add this for debugging

    if (!hoursCheck.exists) {
      // Daily_Hours_Setup mein abhi tak entry nahi thi - client se chuni gayi category/hours se
      // abhi (attendance save hote waqt) entry banao, popup-continue ke waqt nahi.
      if (!category || !hours) {
        throw new Error('Hours and category must be set before marking attendance.');
      }
      // v5.6: the client-sent category/hours must be valid configured values.
      validateHoursValue_(hours, 'normal');
      validateCategoryForClass_(category, 'normal', class_section);
      setDailyHours([{ date, classSection: class_section, category, hours, teacherId: username }], 'normal', recId, username);
      hoursCheck = { exists: true, category, hours };
    }
    // v5.6: the day's setup hours must be a usable number.
    const dayHours = Number(hoursCheck.hours);
    if (isNaN(dayHours) || dayHours <= 0) throw new Error('The hours set for ' + class_section + ' on ' + date + ' are not valid. Fix the hours setup first.');

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
  if (String(rowAttendanceType || '').trim().toLowerCase() === 'episodic') {
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

  // v5.6: only this school's rows can be overwritten (blank REC_ID = any school).
  if (!rowRecMatches_(recIdCol !== undefined ? rowData[recIdCol] : '', schoolRecId)) continue;

  if (rowDateStr === targetDateString && rowClassSection === class_section) {
    existingRecordsByStudentDate.set(stdId, {
      rowIndex: i + 1,
      data: rowData,
      currentStatus: rowData[headerMap.get('Status')]
    });
  }
}
Logger.log(`Found ${existingRecordsByStudentDate.size} existing NORMAL records for ${class_section} on ${date}`);

          // v5.6: roster check — only students enrolled in this class on this day
          // can be marked (same rule as the episodic path). Client-sent names and
          // classes are never trusted; new rows take them from the roster.
          const roster = new Map();
          getStudentList(class_section, date, schoolRecId).forEach(function (s) { roster.set(String(s.Std_ID), s); });

          // -------------------------------------------------
          // -------------------------------------------------
          // Prepare bulk updates / new rows
          // -------------------------------------------------
          const updatesToApply = [];   // { rowIndex, rowValues }
          const newRowsToAppend = [];

          for (const record of attendanceData) {
            const studentKey = String(record.Std_ID);
            const rosterStudent = roster.get(studentKey);
            if (!rosterStudent) {
              Logger.log(`Skipping Std_ID ${studentKey} – not on the roster for ${class_section} on ${date}`);
              continue;
            }
            // v5.6: only Present / Absent / Leave / Other may be stored.
            const recordStatus = String(record.Status || '').trim().toLowerCase();
            if (['present', 'absent', 'leave', 'other'].indexOf(recordStatus) === -1) {
              throw new Error('Invalid attendance status: ' + record.Status + '. Use Present, Absent, Leave or Other.');
            }
            const existing = existingRecordsByStudentDate.get(studentKey);

                        // Hours to store: use sent per-student value if provided, else class hours
            const desiredHours = (record.Hours !== undefined && record.Hours !== '' && record.Hours !== null)
              ? record.Hours
              : hoursCheck.hours;
            // v5.6: per-student hours must be a number within the day's total.
            // A value different from the day total is an hours edit, which needs
            // the canEditStudentHours flag (same gate as the screen).
            const desiredHoursNum = Number(desiredHours);
            if (isNaN(desiredHoursNum) || desiredHoursNum <= 0 || desiredHoursNum > dayHours) {
              throw new Error('Invalid hours for Std_ID ' + studentKey + ': ' + desiredHours + '. Hours must be more than 0 and no more than the day total (' + dayHours + ').');
            }
            if (desiredHoursNum !== dayHours) requirePermissionFlag_(username, 'canEditStudentHours');
              const dayItems = parseCategoryBreakdown_(hoursCheck.category);
            const desiredCategory = dayItems
              ? resolveStudentCategory_(dayItems, hoursCheck.category, hoursCheck.hours, desiredHours, record.Category)
              : hoursCheck.category;

            if (existing) {
              // ---- UPDATE existing row ----
              const existingHours = existing.data[hoursCol];
              const hoursDiff = String(existingHours) !== String(desiredHours);
              const catDiff = !!dayItems && String(existing.data[categoryCol]) !== String(desiredCategory);
              if (existing.currentStatus !== record.Status || hoursDiff || catDiff) {
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
                                updatedRow[hoursCol] = desiredHours;
                updatedRow[categoryCol] = desiredCategory;
                if (recIdCol !== undefined) updatedRow[recIdCol] = schoolRecId;
                
                // ========================================================
                // NEW: Add Term information if columns exist
                // ========================================================
                if (termIdCol !== undefined && termNameCol !== undefined && term) {
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
              newRow[headerMap.get('Date')]           = date;
              newRow[headerMap.get('Std_ID')]         = String(record.Std_ID);
              newRow[headerMap.get('Barcode_ID')]     = rosterStudent.Barcode_ID || '';
              newRow[headerMap.get('Student_Name')]   = rosterStudent.Student_Name;
              newRow[headerMap.get('Student_Class')]  = rosterStudent.Student_Class;
              newRow[headerMap.get('Student_Section')]= rosterStudent.Student_Section;
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
                            newRow[hoursCol] = desiredHours;
              newRow[categoryCol] = desiredCategory;
              if (recIdCol !== undefined) newRow[recIdCol] = schoolRecId;

              // ========================================================
              // NEW: Add Term information if columns exist
              // ========================================================
              if (termIdCol !== undefined && termNameCol !== undefined && term) {
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

// ---------- Applicable_Classes (v5.4) ----------
// Parse an Applicable_Classes cell into an array of class tokens. Accepts '{VIII,IX,X,XI}',
// a bare 'VII', '' (blank = all classes) or an array. Same brace rules as the MIS portal. Never throws.
function parseApplicableClasses_(cell) {
  if (Array.isArray(cell)) return cell.map(function (t) { return String(t == null ? '' : t).trim(); }).filter(Boolean);
  const str = String(cell == null ? '' : cell).trim();
  if (!str) return [];
  let s = str;
  if (s.charAt(0) === '{' && s.charAt(s.length - 1) === '}') s = s.slice(1, -1);
  return s.split(',').map(function (t) { return t.trim(); }).filter(Boolean);
}

// Does this category apply to the given class-section(s)? Blank/missing Applicable_Classes = all
// classes. classSection may be a single 'VII-A', an array of them (category must apply to EVERY
// one — the batch/event rule), or null/[] (no class in scope → no filter, fail-open).
function categoryAppliesToClass_(cell, classSection) {
  if (classSection == null) return true;
  const wanted = (Array.isArray(classSection) ? classSection : [classSection])
    .map(function (s) { return String(s || '').trim(); }).filter(Boolean);
  if (wanted.length === 0) return true;
  const allowed = parseApplicableClasses_(cell);
  if (allowed.length === 0) return true;
  const allowedLower = allowed.map(function (t) { return String(t).trim().toLowerCase(); });
  return wanted.every(function (cs) {
    const base = String(cs).split('-')[0].trim().toLowerCase();
    return allowedLower.indexOf(base) !== -1;
  });
}

/**
 * Get categories filtered by attendance type from Hour_Categories sheet
 * (v5.4: optional classSection filters by Applicable_Classes; blank = all classes)
 */
function getCategoriesByType(type, classSection) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hour_Categories');
  if (!sheet) throw new Error('Hour_Categories sheet not found');
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const nameCol = headers.indexOf('Category_Name');
  const activeCol = headers.indexOf('Is_Active');
  const typeCol = headers.indexOf('Attendance_Type');
  const applicableCol = headers.indexOf('Applicable_Classes');
  
  const categories = [];
  
  for (let i = 1; i < data.length; i++) {
    const isActive = data[i][activeCol] === true || data[i][activeCol] === 'TRUE';
    if (!isActive) continue;
    
    let allowedType = 'both';
    if (typeCol !== -1 && data[i][typeCol]) {
      allowedType = String(data[i][typeCol]).toLowerCase();
    }
    
    if (!type || allowedType === 'both' || allowedType === type) {
      // v5.4: only categories bound to this class (blank Applicable_Classes = all classes).
      if (!categoryAppliesToClass_(applicableCol === -1 ? '' : data[i][applicableCol], classSection)) continue;
      categories.push(String(data[i][nameCol]));
    }
  }
  
  return categories;
}

/**
 * Get episodic categories (v5.4: optional classSection, passed through)
 */
function getEpisodicCategories(classSection) {
  return getCategoriesByType('episodic', classSection);
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
  const activeCol = headers.indexOf('Is_Active');
  
  const presetHours = [];
  let allowCustom = false;
  let maxCustom = 8;
  
  for (let i = 1; i < data.length; i++) {
    // v5.3: skip rows deactivated in the MIS portal (blank/missing Is_Active = active).
    if (activeCol !== -1) {
      const av = data[i][activeCol];
      if (av === false || String(av).trim().toLowerCase() === 'false') continue;
    }
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

    // v5.6: caller verification (unknown / non-teacher / Inactive refused). The
    // school comes from the caller's Users row; the sent recId only disambiguates a username that exists at two schools.
    const caller = requireTeacherCaller_(username, recId);
    const schoolRecId = caller.recId || String(recId || '').trim();
    requireAssignedClass_(username, classSection, schoolRecId);
    if (!isRealYMD_(date)) throw new Error('Invalid date: ' + date + '. Use a real yyyy-MM-dd date.');

    // ========================================================
    // 1b. PREVIOUS-DATE PERMISSION CHECK (server-side)
    // Pichhli date ke liye live Previous_Date_Permissions row
    // (user + date + class + Episodic type) zaroori hai.
    // ========================================================
    const todayStr_ = Utilities.formatDate(new Date(), 'Asia/Karachi', 'yyyy-MM-dd');
    if (date > todayStr_) throw new Error('Future dates cannot be marked.');
    if (date < todayStr_) {
      const allowed = checkPreviousDatePermission(username, date, 'Episodic', classSection, schoolRecId);
      if (!allowed) {
        throw new Error('You do not have permission to mark episodic attendance for this previous date. Please request a new permission from the administrator.');
      }
    }
    
    // v5.6: the event must exist for this date + class (marking under a
    // never-created event is refused).
    const eventRow = findEpisodicEvent_(eventName, date, classSection, schoolRecId);
    if (!eventRow) throw new Error('Event "' + eventName + '" does not exist for ' + classSection + ' on ' + date + '. Create the event first.');
    
    // ========================================================
    // 2. TERM CHECK REMOVED (per request) - term info still fetched (non-blocking) for record-keeping only
    // ========================================================
    const term = getTermForClassSection(classSection, date);
    if (term) {
      Logger.log(`Using term: ${term.termName} (${term.termId}) from ${term.startDate} to ${term.endDate}`);
    } else {
      Logger.log(`No active term found for ${classSection} on ${date} - proceeding anyway (term check disabled).`);
    }
    
    // ========================================================
    // ========================================================
    Logger.log(`Schedule check skipped for episodic attendance: ${username}`);
    
    // ========================================================
    // 4. HOURS VALIDATION - Check that hours & category are defined for the day
    // CRITICAL: Episodic attendance also needs hours validation
    // ========================================================
const hoursCheck = checkDailyHours(date, classSection, 'episodic', schoolRecId);

let finalHours = hours;
let finalCategory = category;

if (!hoursCheck.exists && (!category || !hours)) {
  throw new Error('Hours and category must be set before marking episodic attendance.');
}
if (!hoursCheck.exists) {
  // v5.6: client-sent hours/category must be valid configured values.
  validateHoursValue_(hours, 'episodic');
  validateCategoryForClass_(category, 'episodic', classSection);
}
// v5.6: the event's hours must be a usable number.
const eventHoursNum = Number(hoursCheck.exists ? hoursCheck.hours : finalHours);
if (isNaN(eventHoursNum) || eventHoursNum <= 0) throw new Error('The hours for this event are not valid. Fix the event or the hours setup first.');

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
    const allStudents = getStudentList(classSection, date, schoolRecId);
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
          String(rowType || '').trim().toLowerCase() === 'episodic' &&
          rowEvent === eventName &&
          rowRecMatches_(recIdCol !== undefined ? row[recIdCol] : '', schoolRecId)) {
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
      // v5.6: only Present / Not Applicable may be stored for episodic attendance.
      const itemStatus = String(newStatus || '').trim().toLowerCase();
      if (itemStatus !== 'present' && itemStatus !== 'not applicable') {
        throw new Error('Invalid episodic status for Std_ID ' + studentId + ': ' + newStatus + '. Use Present or Not Applicable.');
      }
      
            const existing = existingRecords.get(studentId);
      const now = new Date();
      const desiredHours = (newStatus === 'Present' && item.hours !== undefined && item.hours !== '' && item.hours !== null) ? item.hours : finalHours;
      // v5.6: per-student hours must be a number within the event's hours.
      // A value different from the event total is an hours edit, which needs
      // the canEditStudentHours flag (same gate as the screen).
      const desiredHoursNum = Number(desiredHours);
      if (isNaN(desiredHoursNum) || desiredHoursNum <= 0 || desiredHoursNum > eventHoursNum) {
        throw new Error('Invalid hours for Std_ID ' + studentId + ': ' + desiredHours + '. Hours must be more than 0 and no more than the event total (' + eventHoursNum + ').');
      }
      if (newStatus === 'Present' && desiredHoursNum !== eventHoursNum) requirePermissionFlag_(username, 'canEditStudentHours');
      
      if (existing) {
        // Update existing record if status OR hours changed
        const hoursDiff = String(existing.rowData[hoursCol]) !== String(desiredHours);
        if (existing.currentStatus !== newStatus || hoursDiff) {
          // Create updated row data based on existing row
          const updatedRow = [...existing.rowData];
          updatedRow[statusCol] = newStatus;
          updatedRow[timestampCol] = now.toLocaleString();
          updatedRow[teacherCol] = username;
          
          // Hours: per-student value when Present (if provided), else event's default hours
 updatedRow[hoursCol] = desiredHours;
updatedRow[categoryCol] = finalCategory;
          if (recIdCol !== undefined) updatedRow[recIdCol] = schoolRecId;

          // Ensure term info is correct
          if (termIdCol !== undefined && termNameCol !== undefined && term) {
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
        
                // Hours: per-student value when Present (if provided), else event's default hours
        newRow[hoursCol] = desiredHours;
newRow[categoryCol] = finalCategory;
        if (recIdCol !== undefined) newRow[recIdCol] = schoolRecId;
        if (typeCol !== undefined) newRow[typeCol] = 'Episodic';
        if (eventCol !== undefined) newRow[eventCol] = eventName;
        if (termIdCol !== undefined && termNameCol !== undefined && term) {
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

const episodicHoursExist = checkDailyHours(date, classSection, 'episodic', schoolRecId);

if (!episodicHoursExist.exists) {
  // No episodic hours for this date/class - safe to save
  try {
        setDailyHours([{
      date: date,
      classSection: classSection,
      category: finalCategory,
      hours: finalHours,
      teacherId: username
    }], 'episodic', recId, username);
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

// v5.6: appendRow deleted — it was an unguarded direct write to Student_Attendance with no callers.
function getActiveCategories(type = 'normal', classSection) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hour_Categories');
  if (!sheet) throw new Error('Hour_Categories sheet not found');
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const typeCol = headers.indexOf('Attendance_Type');
  const applicableCol = headers.indexOf('Applicable_Classes');
  
  const categories = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === true || data[i][2] === 'TRUE') {
      let allowedType = 'both';
      if (typeCol !== -1 && data[i][typeCol]) {
        allowedType = String(data[i][typeCol]).toLowerCase();
      }
      // Include if: no type specified OR type is 'both' OR matches requested type
      if (!type || allowedType === 'both' || allowedType === type) {
        // v5.4: only categories bound to this class (blank Applicable_Classes = all classes).
        if (!categoryAppliesToClass_(applicableCol === -1 ? '' : data[i][applicableCol], classSection)) continue;
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
  const activeCol = headers.indexOf('Is_Active');
  
  const presetHours = [];
  let allowCustom = false;
  let maxCustom = 8;
  
  for (let i = 1; i < data.length; i++) {
    // v5.3: skip rows deactivated in the MIS portal (blank/missing Is_Active = active).
    if (activeCol !== -1) {
      const av = data[i][activeCol];
      if (av === false || String(av).trim().toLowerCase() === 'false') continue;
    }
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
// NEW: Check if hours/category are set for a date and class_section
function checkDailyHours(date, classSection, attendanceType = null, schoolRecId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Hours_Setup');
  if (!sheet) throw new Error('Daily_Hours_Setup sheet not found');
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const typeCol = headers.indexOf('Attendance_Type');
  const recCol = headers.indexOf('REC_ID');
  const targetDate = new Date(date).toLocaleDateString('en-CA');
  
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]).toLocaleDateString('en-CA');
    if (rowDate === targetDate && data[i][1] === classSection) {
      // v5.6: only this school's setup rows (blank REC_ID = any school).
      if (!rowRecMatches_(recCol !== -1 ? data[i][recCol] : '', schoolRecId)) continue;
      const rowType = (typeCol !== -1 && data[i][typeCol]) ? String(data[i][typeCol]) : 'normal';
      
      // If type specified (like 'normal' or 'episodic'), only match that type
      // (compared case-insensitively, so 'Normal' and 'normal' are the same row).
      if (attendanceType && rowType.toLowerCase() === String(attendanceType).toLowerCase()) {
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
function setDailyHours(entries, attendanceType = 'normal', recId = '', username) {
  // v5.6: caller verification, assignment, school, canSetupHours flag and
  // per-entry validation. (Future dates are allowed here — the batch and
  // hours-setup screens let teachers set hours ahead; no past-date
  // permission is demanded for setup rows.)
  const caller = requireTeacherCaller_(username, recId);
  const schoolRecId = caller.recId || String(recId || '').trim();
  requirePermissionFlag_(username, 'canSetupHours');
  const typeNorm = String(attendanceType || 'normal').trim().toLowerCase();
  if (typeNorm !== 'normal' && typeNorm !== 'episodic') throw new Error('Invalid attendance type: ' + attendanceType + '. Use normal or episodic.');
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
      // v5.6: every entry is validated — real date, assigned class, valid hours/category.
      if (!isRealYMD_(entry.date)) throw new Error('Invalid date: ' + entry.date + '. Use a real yyyy-MM-dd date.');
      requireAssignedClass_(username, entry.classSection, schoolRecId);
      validateHoursValue_(entry.hours, typeNorm);
      validateCategoryForClass_(entry.category, typeNorm, entry.classSection);
            const bd_ = parseCategoryBreakdown_(entry.category);
      if (bd_ && Math.abs(sumBreakdown_(bd_) - Number(entry.hours)) > 0.001) {
        throw new Error('Category hours breakup is not equal to the total Hours defined.');
      }
      const targetDate = new Date(entry.date).toLocaleDateString('en-CA');
      let found = false;
      
      for (let i = 1; i < existingData.length; i++) {
        const rowDate = new Date(existingData[i][0]).toLocaleDateString('en-CA');
        if (rowDate === targetDate && existingData[i][1] === entry.classSection) {
          // v5.6: only this school's rows can be overwritten (blank REC_ID = any school).
          if (!rowRecMatches_(recIdColIndex !== -1 ? existingData[i][recIdColIndex] : '', schoolRecId)) continue;
          
          // ✅ CRITICAL FIX: Get the existing Attendance_Type
          const existingType = existingData[i][typeColIndex] || 'normal';
          
          // ✅ ONLY update if the existing row has the SAME attendance type
          // (compared case-insensitively, so 'Normal' and 'normal' are the same row).
          if (String(existingType).toLowerCase() === typeNorm) {
            // Update existing row of same type
            const values = new Array(Math.max(7, existingData[i].length, recIdColIndex + 1)).fill('');
            values[0] = entry.date;
            values[1] = entry.classSection;
            values[2] = entry.category;
            values[3] = entry.hours;
            values[4] = username;
            values[5] = now;
            values[typeColIndex] = attendanceType;
            if (recIdColIndex !== -1) values[recIdColIndex] = schoolRecId;
            
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
        newRow[4] = username;
        newRow[5] = now;
        newRow[typeColIndex] = attendanceType;
        if (recIdColIndex !== -1) newRow[recIdColIndex] = schoolRecId;
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
function getClassSectionsWithHoursStatus(date, username, schoolRecId) {
  try {
    const sheet = getHoursSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const dateCol = headers.indexOf('Date');
    const classSectionCol = headers.indexOf('Class_Section');
    const hoursCol = headers.indexOf('Hours');
    const categoryCol = headers.indexOf('Category');
    const recCol = headers.indexOf('REC_ID');
    
    const existingHours = {};
    const targetDate = new Date(date).toLocaleDateString('en-CA');
    
    if (dateCol !== -1 && classSectionCol !== -1) {
      data.slice(1)
        .filter(row => {
          if (!row[dateCol]) return false;
          // v5.8 (O5): only this school's setup rows (blank REC_ID = any
          // school; no school passed = old behaviour, every row counts).
          if (!rowRecMatches_(recCol !== -1 ? row[recCol] : '', schoolRecId)) return false;
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
          // v5.8 (O5): the logged-in school scopes the assignment list.
          userClassSections = getAssignedClassSections(username, schoolRecId);
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

// v5.6: handleBatchHoursSubmission was deleted — it had no caller (the batch
// form calls setDailyHours) and no validation at all.
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
// ================= BATCH HOURS UPDATE LOG =================
const BATCH_HOURS_LOG_SHEET = 'Batch_Hours_Log';
const BATCH_HOURS_LOG_HEADERS = ['Date','Class_Section','Old_Hours','Old_Category','New_Hours','New_Category','UserID','REC_ID','Timestamp'];

function getOrCreateBatchHoursLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BATCH_HOURS_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BATCH_HOURS_LOG_SHEET);
    sheet.getRange(1, 1, 1, BATCH_HOURS_LOG_HEADERS.length).setValues([BATCH_HOURS_LOG_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// blank ya "normal" (chhota/bara koi bhi) = Normal attendance
function isNormalType_(v) {
  const t = String(v || '').trim().toLowerCase();
  return t === '' || t === 'normal';
}

function updateBatchHours(date, classSection, category, hours, username, recId = '') {
  const lock = LockService.getScriptLock();
  
  try {
    // v5.6: caller verification, assignment, school, date rules, permission flag,
    // and hours/category validation. (Throws here become {success:false} below.)
    const caller = requireTeacherCaller_(username, recId);
    const schoolRecId = caller.recId || String(recId || '').trim();
    requireAssignedClass_(username, classSection, schoolRecId);
    if (!isRealYMD_(date)) throw new Error('Invalid date: ' + date + '. Use a real yyyy-MM-dd date.');
    const todayStr_ = Utilities.formatDate(new Date(), 'Asia/Karachi', 'yyyy-MM-dd');
    if (date > todayStr_) throw new Error('Future dates cannot be updated.');
    if (date < todayStr_) {
      const allowed = checkPreviousDatePermission(username, date, 'Normal', classSection, schoolRecId);
      if (!allowed) throw new Error('You do not have permission to change hours for this previous date. Please request a new permission from the administrator.');
    }
    requirePermissionFlag_(username, 'canUpdateBatchHours');
    validateHoursValue_(hours, 'normal');
    validateCategoryForClass_(category, 'normal', classSection);
    const bd_ = parseCategoryBreakdown_(category);
    if (bd_ && Math.abs(sumBreakdown_(bd_) - Number(hours)) > 0.001) {
      return { success: false, message: 'Category hours breakup is not equal to the total Hours defined.' };
    }
    if (!lock.tryLock(10000)) {
      return { success: false, message: 'Unable to acquire lock. Please try again.' };
    }
    
    // First update Daily_Hours_Setup
    const hoursSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Hours_Setup');
    if (!hoursSheet) throw new Error('Daily_Hours_Setup sheet not found');
    
    const hoursData = hoursSheet.getDataRange().getValues();
    const hoursTypeCol = hoursData[0].indexOf('Attendance_Type');
    const hoursRecCol = hoursData[0].indexOf('REC_ID');
    const targetDate = new Date(date).toLocaleDateString('en-CA');
    let updatedHoursSetup = false;
    const now = new Date().toLocaleString();
    let oldHours = '';
    let oldCategory = '';
    
    // Update or create entry in Daily_Hours_Setup
    for (let i = 1; i < hoursData.length; i++) {
      const rowDate = new Date(hoursData[i][0]).toLocaleDateString('en-CA');
      if (rowDate === targetDate && hoursData[i][1] === classSection &&
    (hoursTypeCol === -1 || isNormalType_(hoursData[i][hoursTypeCol])) &&
    rowRecMatches_(hoursRecCol !== -1 ? hoursData[i][hoursRecCol] : '', schoolRecId)) {
        oldCategory = hoursData[i][2];
        oldHours = hoursData[i][3];
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
      // v5.6: stamp the new setup row with its type + school (header-indexed).
      const newSetupRow = hoursSheet.getLastRow();
      if (hoursTypeCol !== -1) hoursSheet.getRange(newSetupRow, hoursTypeCol + 1).setValue('Normal');
      if (hoursRecCol !== -1) hoursSheet.getRange(newSetupRow, hoursRecCol + 1).setValue(schoolRecId);
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
    const attTypeCol = headers.indexOf('Attendance_Type');
    const attRecCol = headers.indexOf('REC_ID');

    if (dateCol === -1 || classSectionCol === -1 || hoursCol === -1 || categoryCol === -1) {
      return { success: false, message: 'Required columns not found in attendance sheet.' };
    }
    
    let updatedCount = 0;
    
    for (let i = 1; i < attendanceData.length; i++) {
      const rowDate = new Date(attendanceData[i][dateCol]).toLocaleDateString('en-CA');
      if (rowDate === targetDate && attendanceData[i][classSectionCol] === classSection &&
    (attTypeCol === -1 || isNormalType_(attendanceData[i][attTypeCol])) &&
    rowRecMatches_(attRecCol !== -1 ? attendanceData[i][attRecCol] : '', schoolRecId)) {
        attendanceSheet.getRange(i + 1, hoursCol + 1).setValue(hours);
        attendanceSheet.getRange(i + 1, categoryCol + 1).setValue(category);
        updatedCount++;
      }
    }
    
    SpreadsheetApp.flush();
    
    // Log this batch hours update — sirf tab jab kam-az-kam 1 record actually update hua ho
    if (updatedCount > 0) {
      getOrCreateBatchHoursLogSheet_().appendRow([date, classSection, oldHours, oldCategory, hours, category, username, schoolRecId, new Date().toLocaleString()]);
    }
    
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

function parseCategoryBreakdown_(str) {          
  const s = String(str || '');
  if (s.indexOf(' + ') === -1) return null;
  const items = [];
  for (const part of s.split(' + ')) {
    const m = part.trim().match(/^(.+) \((\d+(?:\.\d+)?)\)$/);
    if (!m) return null;
    items.push({ name: m[1], hours: Number(m[2]) });
  }
  return items;
}
function formatCategoryBreakdown_(items) {          // 1 item = sadi category, 2+ = "A (1) + B (2)"
  if (items.length === 1) return items[0].name;
  return items.map(i => `${i.name} (${Number(i.hours)})`).join(' + ');
}
function sumBreakdown_(items) { return items.reduce((a, i) => a + i.hours, 0); }
function splitHoursByRatio_(items, studentHours) {  // din ke ratio se 0.25 ke step mein baanto
  // largest-remainder split in quarter-hour units, so the parts always add up exactly
  const q = 0.25, dayTotal = sumBreakdown_(items);
  const exact = items.map(i => (studentHours * i.hours / dayTotal) / q);
  const out = exact.map(Math.floor);
  let left = Math.round(studentHours / q) - out.reduce((a, b) => a + b, 0);
  const order = exact.map((_, k) => k).sort((a, b) => (exact[b] - out[b]) - (exact[a] - out[a]));
  for (let k = 0; left > 0; k = (k + 1) % order.length) { out[order[k]]++; left--; }
  return items.map((i, k) => ({ name: i.name, hours: out[k] * q })).filter(i => i.hours > 0);
}
function resolveStudentCategory_(dayItems, dayCategory, dayHours, studentHours, sentCategory) {
  const sh = Number(studentHours);
  if (!(sh > 0) || Math.abs(sh - Number(dayHours)) < 0.001) return dayCategory;
  const sent = parseCategoryBreakdown_(sentCategory);
  if (sent && sent.every(s => dayItems.some(d => d.name === s.name && s.hours <= d.hours)) &&
      Math.abs(sumBreakdown_(sent) - sh) < 0.001) return formatCategoryBreakdown_(sent);
  if (sentCategory && !sent) {                       // sirf ek category
    const d = dayItems.find(i => i.name === sentCategory);
    if (d && sh <= d.hours) return sentCategory;
  }
  return formatCategoryBreakdown_(splitHoursByRatio_(dayItems, sh));
}


// NEW: Extend getUserPermissions to include canSetupHours
// UPDATE THIS EXISTING FUNCTION (replace it completely):
function getUserPermissions(username) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User_Permissions');
    if (!sheet) throw new Error('User_Permissions sheet not found');
    const data = sheet.getDataRange().getValues();
    const mcCol = data[0].indexOf('canMultiCategory');
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
          canMarkEpisodic: data[i][8] === 'TRUE' || data[i][8] === true,
          canManageEpisodicEvents: data[i][10] === 'TRUE' || data[i][10] === true,
          canEditStudentHours: data[i][11] === 'TRUE' || data[i][11] === true,
          canMultiCategory: mcCol !== -1 && (data[i][mcCol] === 'TRUE' || data[i][mcCol] === true)
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
      canMarkEpisodic: false,
      canManageEpisodicEvents: false,
      canEditStudentHours: false,
      canMultiCategory: false
    };
  } catch (e) {
    Logger.log('Error in getUserPermissions: ' + e.message);
    throw e;
  }
}


// =============== v5.6: SERVER-SIDE CALLER VALIDATION ===============
// Every write below verifies its caller through these helpers. The caller's
// school always comes from the Users row (never from a client-sent recId).
// A teacher with no Teacher_Assignments rows on file has no assignment
// constraint (fail-open, same as the class pickers' fallback); a teacher WITH
// assignments may only write to their own classes. A missing User_Permissions
// row is fail-open for the flags below; a present row with the flag off refuses.

function getCallerUser_(username, recId) {
  // The caller's Users row. One row with this username: that row. Several
  // rows (same teacher at two schools): the one whose REC_ID matches the
  // school the caller logged into (recId); a row with a blank REC_ID counts
  // as a match for any school. No row matches -> null (ambiguous caller).
  // Returns { username, userType, status, recId } or null.
  if (!username || !String(username).trim()) return null;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) throw new Error('Users sheet not found');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const h = data[0];
  const uCol = h.indexOf('Username');
  const tCol = h.indexOf('UserType');
  const sCol = h.indexOf('Status');
  const rCol = h.indexOf('REC_ID');
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][uCol] || '').trim() === String(username).trim()) {
      rows.push({
        username: String(data[i][uCol] || '').trim(),
        userType: tCol !== -1 ? String(data[i][tCol] || '').trim() : '',
        status: sCol !== -1 ? String(data[i][sCol] || '').trim() : 'Active',
        recId: rCol !== -1 ? String(data[i][rCol] || '').trim() : ''
      });
    }
  }
  if (rows.length === 0) return null;
  if (rows.length === 1) return rows[0];
  const wantRec = String(recId || '').trim();
  if (wantRec) {
    for (let k = 0; k < rows.length; k++) {
      if (rows[k].recId === wantRec) return rows[k];
    }
    for (let k = 0; k < rows.length; k++) {
      if (!rows[k].recId) return rows[k];
    }
    return null;
  }
  return rows[0];
}

function requireTeacherCaller_(username, recId) {
  // Refuses unknown usernames, accounts that are not teacher/admin, and
  // Inactive accounts. Returns { username, userType, status, recId }.
  const caller = getCallerUser_(username, recId);
  if (!caller) throw new Error('Unknown user. Please log in again.');
  const type = caller.userType.toLowerCase();
  if (type !== 'teacher' && type !== 'admin') throw new Error('This account is not allowed to change attendance data.');
  if (caller.status.toLowerCase() !== 'active') throw new Error('Your account is inactive. Please contact the administrator.');
  return caller;
}

function requireAdminCaller_(username, recId) {
  // v5.8: an Active admin OR superadmin may proceed (used by updateUserStatus).
  const caller = getCallerUser_(username, recId);
  if (!caller) throw new Error('Unknown user. Please log in again.');
  const type = caller.userType.toLowerCase();
  if (type !== 'admin' && type !== 'superadmin') throw new Error('Only an administrator can do this.');
  if (caller.status.toLowerCase() !== 'active') throw new Error('Your account is inactive. Please contact the administrator.');
  return caller;
}

// v5.8 (F5, corrected 19 Nov 2026): the schools an admin caller may change
// users at, derived from ALL of the caller's Users rows — not just the first
// row getCallerUser_ picks. Only Active admin/superadmin rows count, so an
// Inactive admin row at a second school grants nothing there. isSuperadmin
// is reported but grants NO cross-school access: a superadmin may act only
// at schools where they hold a row, like any admin. A blank REC_ID grants
// no school (fail-closed): a row with no school on file does not silently
// become every school. Returns { isSuperadmin: boolean, recIds: string[] }.
function getAdminCallerScope_(callerUsername) {
  const scope = { isSuperadmin: false, recIds: [] };
  if (!callerUsername || !String(callerUsername).trim()) return scope;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) throw new Error('Users sheet not found');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return scope;
  const h = data[0];
  const uCol = h.indexOf('Username');
  const tCol = h.indexOf('UserType');
  const sCol = h.indexOf('Status');
  const rCol = h.indexOf('REC_ID');
  const want = String(callerUsername).trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][uCol] || '').trim() !== want) continue;
    const type = tCol !== -1 ? String(data[i][tCol] || '').trim().toLowerCase() : '';
    if (type !== 'admin' && type !== 'superadmin') continue;
    const status = sCol !== -1 ? String(data[i][sCol] || '').trim().toLowerCase() : 'active';
    if (status !== 'active') continue;
    if (type === 'superadmin') scope.isSuperadmin = true;
    const rec = rCol !== -1 ? String(data[i][rCol] || '').trim() : '';
    if (rec && scope.recIds.indexOf(rec) === -1) scope.recIds.push(rec);
  }
  return scope;
}

function requireAssignedClass_(username, classSection, schoolRecId) {
  // The class must be in the teacher's assigned list. No Teacher_Assignments
  // rows on file for this teacher/school = no constraint (fail-open).
  const assigned = getAssignedClassSections(username, schoolRecId);
  if (!assigned || assigned.length === 0) return;
  if (!isClassSectionAssigned_(assigned, classSection)) {
    throw new Error('You are not assigned to class ' + classSection + '.');
  }
}

function requireAssignedClasses_(username, classSections, schoolRecId) {
  const list = classSections || [];
  for (let i = 0; i < list.length; i++) requireAssignedClass_(username, list[i], schoolRecId);
}

function hasUserPermissionsRow_(username) {
  // True when the User_Permissions sheet has a row for this username.
  // (getUserPermissions returns all-false both for a missing row and for a
  // row with every flag off, so the existence check has to be separate.)
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User_Permissions');
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

function requirePermissionFlag_(username, flagName) {
  // Missing permissions row = fail-open (teachers with no flags row keep
  // today's behaviour); a present row with the flag off = refuse.
  if (!hasUserPermissionsRow_(username)) return;
  const perms = getUserPermissions(username);
  if (!perms[flagName]) throw new Error('User does not have permission for this action.');
}

function isRealYMD_(s) {
  // Strict real-date check: shape yyyy-MM-dd AND a real calendar date
  // (2026-02-31 is refused, not just badly shaped values).
  const v = String(s || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) return false;
  const parts = v.split('-').map(Number);
  const d = new Date(parts[0], parts[1] - 1, parts[2]);
  return d.getFullYear() === parts[0] && d.getMonth() === parts[1] - 1 && d.getDate() === parts[2];
}

function rowRecMatches_(rowRec, schoolRecId) {
  // Grandfathering: a stored row with a blank/missing REC_ID matches any
  // school, so older rows keep working instead of duplicating or vanishing.
  const r = String(rowRec == null ? '' : rowRec).trim();
  if (!r) return true;
  const want = String(schoolRecId || '').trim();
  if (!want) return true;
  return r === want;
}

function validateHoursValue_(hours, type) {
  // Hours must be a number > 0 and one of the configured preset hours for
  // this type, or a custom value up to Max_Custom when custom is allowed.
  // A type with no configured hours at all accepts any positive number.
  const n = Number(hours);
  if (hours === '' || hours === null || hours === undefined || isNaN(n) || n <= 0) {
    throw new Error('Hours must be a number greater than 0.');
  }
  const t = String(type || 'normal').trim().toLowerCase();
  let cfg = null;
  try { cfg = getHoursByType(t); } catch (e) { cfg = null; }
  if (!cfg) return n;
  const presets = (cfg.presetHours || []).map(function (h) { return Number(h); }).filter(function (h) { return !isNaN(h); });
  if (presets.indexOf(n) !== -1) return n;
  if (cfg.allowCustom && n <= Number(cfg.maxCustom || 0)) return n;
  if (presets.length === 0 && !cfg.allowCustom) return n;
  throw new Error('Hours ' + hours + ' is not an allowed value for ' + t + ' hours. Use one of the configured hours' + (cfg.allowCustom ? ' or a custom value up to ' + cfg.maxCustom + '.' : '.'));
}

function validateCategoryForClass_(category, type, classSections) {
  // Every part of a (possibly multi-part) category must be an active category
  // of this type that applies to every listed class.
  const t = String(type || 'normal').trim().toLowerCase();
  const allowed = getCategoriesByType(t, classSections);
  const items = parseCategoryBreakdown_(category);
  const names = items ? items.map(function (i) { return i.name; }) : [String(category || '').trim()];
  for (let i = 0; i < names.length; i++) {
    if (!names[i] || allowed.indexOf(names[i]) === -1) {
      throw new Error('Category "' + names[i] + '" is not an active ' + t + ' category for ' + (Array.isArray(classSections) ? classSections.join(', ') : classSections) + '.');
    }
  }
}

function findEpisodicEventById_(eventId) {
  // Episodic_Events row by Event_ID, or null. Returns the row's fields plus rowIndex.
  const sheet = getOrCreateEpisodicEventsSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const h = data[0];
  const c = {
    id: h.indexOf('Event_ID'), name: h.indexOf('Event_Name'), date: h.indexOf('Event_Date'),
    cat: h.indexOf('Category'), hrs: h.indexOf('Hours'), cls: h.indexOf('Class_Sections'),
    user: h.indexOf('Username'), rec: h.indexOf('REC_ID')
  };
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][c.id]) === String(eventId)) {
      return {
        rowIndex: i + 1,
        eventId: data[i][c.id],
        eventName: data[i][c.name],
        eventDate: normalizeYMD_(data[i][c.date]),
        category: data[i][c.cat],
        hours: data[i][c.hrs],
        classSections: String(data[i][c.cls] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean),
        username: String(data[i][c.user] || ''),
        recId: c.rec !== -1 ? String(data[i][c.rec] || '').trim() : ''
      };
    }
  }
  return null;
}

function findEpisodicEvent_(eventName, date, classSection, schoolRecId) {
  // The event with this name + date whose class list contains the class and
  // whose school matches (blank REC_ID = any school), or null.
  const sheet = getOrCreateEpisodicEventsSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const h = data[0];
  const c = {
    id: h.indexOf('Event_ID'), name: h.indexOf('Event_Name'), date: h.indexOf('Event_Date'),
    cat: h.indexOf('Category'), hrs: h.indexOf('Hours'), cls: h.indexOf('Class_Sections'),
    user: h.indexOf('Username'), rec: h.indexOf('REC_ID')
  };
  const wantDate = normalizeYMD_(date);
  const wantName = String(eventName || '').trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][c.name] || '').trim().toLowerCase() !== wantName) continue;
    if (normalizeYMD_(data[i][c.date]) !== wantDate) continue;
    const sections = String(data[i][c.cls] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    if (sections.indexOf(classSection) === -1) continue;
    if (!rowRecMatches_(c.rec !== -1 ? data[i][c.rec] : '', schoolRecId)) continue;
    return {
      rowIndex: i + 1,
      eventId: data[i][c.id],
      eventName: data[i][c.name],
      eventDate: normalizeYMD_(data[i][c.date]),
      category: data[i][c.cat],
      hours: data[i][c.hrs],
      classSections: sections,
      username: String(data[i][c.user] || ''),
      recId: c.rec !== -1 ? String(data[i][c.rec] || '').trim() : ''
    };
  }
  return null;
}

function markedSectionsForEvent_(eventName, eventDate, schoolRecId) {
  // Class-sections with Student_Attendance rows under this event (name + date,
  // this school only). Used to block deleting an event — or removing a class
  // from it — after attendance was marked.
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Attendance');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const h = data[0];
  const dateCol = h.indexOf('Date');
  const classCol = h.indexOf('Class_Section');
  const eventCol = h.indexOf('Event_Name');
  const recCol = h.indexOf('REC_ID');
  if (dateCol === -1 || classCol === -1 || eventCol === -1) return [];
  const wantDate = normalizeYMD_(eventDate);
  const wantName = String(eventName || '').trim();
  const marked = new Set();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][eventCol] || '').trim() !== wantName) continue;
    let rowDate;
    try { rowDate = new Date(data[i][dateCol]).toLocaleDateString('en-CA'); } catch (e) { continue; }
    if (rowDate !== wantDate) continue;
    if (!rowRecMatches_(recCol !== -1 ? data[i][recCol] : '', schoolRecId)) continue;
    if (data[i][classCol]) marked.add(data[i][classCol]);
  }
  return Array.from(marked);
}

// =============== REPORT FUNCTIONS ===============

// NEW: Generate student attendance report for PDF/CSV download
// ===== REPORT MATH — MIS portal Daily/Monthly rules (v3 port) =====
// Pure helpers (no sheet access) so the maths can be unit-checked in Node.
// Portal rules implemented here:
// - Only Present earns hours (Absent/Leave rows never contribute hours).
// - Statuses other than present/absent/leave are ignored (never counted).
// - Multiple rows for one student+day count as multiple attendances
//   (portal rule: daily shows Present (2); monthly counts the rows), so
//   daily and monthly always agree. Marking prevents new duplicates.
// - Episodic rows (Attendance_Type = 'Episodic') are excluded: the daily
//   and monthly reports here are the NORMAL reports (same as getStudentList).
// - Inactive_Date is the first day the student is NOT active (portal rule).

// First ACTIVE Class_Terms row wins (sheet order), REC-matched when recId given.
// rows: [{Class_Section, Term_ID, Is_Active, Start_Date, End_Date, REC_ID}]
// Returns {termId, fromDate, toDate} (yyyy-MM-dd) or null.
function pickTermRange_(rows, classSection, termId, recId) {
  const wantTerm = String(termId || '').trim().toLowerCase();
  const wantRec = String(recId || '').trim();
  const list = rows || [];
  for (let i = 0; i < list.length; i++) {
    const r = list[i];
    if (!r) continue;
    if (String(r.Term_ID || '').trim().toLowerCase() !== wantTerm) continue;
    if (String(r.Is_Active == null ? '' : r.Is_Active).trim().toLowerCase() !== 'true') continue;
    if (!classCellMatches_(String(r.Class_Section || ''), classSection)) continue;
    if (wantRec && r.REC_ID !== undefined && r.REC_ID !== null && String(r.REC_ID).trim() !== '' && String(r.REC_ID).trim() !== wantRec) continue;
    const s = normalizeYMD_(r.Start_Date);
    const e = normalizeYMD_(r.End_Date);
    if (!s || !e) continue;
    return { termId: String(r.Term_ID || '').trim(), fromDate: s, toDate: e };
  }
  return null;
}

// Is the student enrolled on targetYMD (yyyy-MM-dd)? Portal rule:
// joined (no DoJ, or DoJ <= day) AND not left (no Inactive_Date, or Inactive_Date > day).
function isEnrolledOnYMD_(targetYMD, dojYMD, inactiveYMD) {
  const d = String(dojYMD || '').trim();
  const ia = String(inactiveYMD || '').trim();
  const joinedOK = !d || d <= targetYMD;
  const notLeftOK = !ia || ia > targetYMD;
  return joinedOK && notLeftOK;
}

// Is the student enrolled at ANY POINT during [startYMD, endYMD]? Portal rule
// (same formula as the portal's isEnrolledInMonth, on yyyy-MM-dd strings so no
// timezone shift is possible).
function isEnrolledInMonthYMD_(startYMD, endYMD, dojYMD, inactiveYMD) {
  const d = String(dojYMD || '').trim();
  const ia = String(inactiveYMD || '').trim();
  const joinedByMonthEnd = !d || d <= endYMD;
  const notLeftBeforeStart = !ia || ia >= startYMD;
  return joinedByMonthEnd && notLeftBeforeStart;
}

// 'Active' / 'Inactive' in a month (portal monthStatusFor, on yyyy-MM-dd strings).
function monthStatusForYMD_(startYMD, endYMD, dojYMD, inactiveYMD) {
  return isEnrolledInMonthYMD_(startYMD, endYMD, dojYMD, inactiveYMD) ? 'Active' : 'Inactive';
}

// Daily aggregation (pure). rows: [{stdId, status, hours, type}] already filtered
// to one class-section + one date. Returns per-student {present, absent, leave, other, hours}.
// Hours count only on Present rows; Episodic rows are skipped.
function aggregateDailyRows_(rows) {
  const out = {};
  (rows || []).forEach(function (r) {
    if (String(r.type || '').toLowerCase().trim() === 'episodic') return;
    const sid = String(r.stdId || '').trim();
    if (!sid) return;
    if (!out[sid]) out[sid] = { present: 0, absent: 0, leave: 0, other: 0, hours: 0 };
    const st = String(r.status || '').toLowerCase().trim();
    const hrs = Number(r.hours || 0);
    if (st === 'present') {
      out[sid].present++;
      if (!isNaN(hrs) && hrs > 0) out[sid].hours += hrs;
    } else if (st === 'absent') out[sid].absent++;
    else if (st === 'leave') out[sid].leave++;
    else out[sid].other++;
  });
  return out;
}

// Portal daily status label for one student (counts shown when > 1).
function dailyStatusLabel_(c) {
  if (!c) return 'Not Marked';
  if (c.present > 0) return 'Present' + (c.present > 1 ? ' (' + c.present + ')' : '');
  if (c.absent > 0) return 'Absent' + (c.absent > 1 ? ' (' + c.absent + ')' : '');
  if (c.leave > 0) return 'Leave' + (c.leave > 1 ? ' (' + c.leave + ')' : '');
  if (c.other > 0) return 'Other' + (c.other > 1 ? ' (' + c.other + ')' : '');
  return 'Not Marked';
}

// Monthly aggregation (pure). rows: [{stdId, dateYMD, status, hours, type}] for one
// student. Only rows with startYMD <= dateYMD <= endYMD count. Unknown statuses
// are ignored (never counted). Hours count only on Present rows.
function aggregateMonthlyRows_(rows, startYMD, endYMD) {
  const agg = { presents: 0, absents: 0, leaves: 0, totalDays: 0, hours: 0 };
  (rows || []).forEach(function (r) {
    const k = String(r.dateYMD || '').trim();
    if (!k || k < startYMD || k > endYMD) return;
    if (String(r.type || '').toLowerCase().trim() === 'episodic') return;
    const st = String(r.status || '').toLowerCase().trim();
    const hrs = Number(r.hours || 0);
    if (st === 'present') {
      agg.presents++; agg.totalDays++;
      if (!isNaN(hrs) && hrs > 0) agg.hours += hrs;
    } else if (st === 'absent') { agg.absents++; agg.totalDays++; }
    else if (st === 'leave') { agg.leaves++; agg.totalDays++; }
  });
  return agg;
}

// Working days in a range (pure). setupRows: [{dateYMD, classSection}] from
// Daily_Hours_Setup. Counts DISTINCT dates in range for the class-section.
function countWorkingDays_(setupRows, classSection, startYMD, endYMD) {
  const seen = {};
  let n = 0;
  (setupRows || []).forEach(function (r) {
    if (String(r.classSection || '').trim() !== String(classSection || '').trim()) return;
    const k = String(r.dateYMD || '').trim();
    if (!k || k < startYMD || k > endYMD) return;
    if (!seen[k]) { seen[k] = true; n++; }
  });
  return n;
}

// Assignment check (pure). The report pickers list assigned classes only; the
// server re-checks so a direct call can never show another class.
function isClassSectionAssigned_(assignedList, classSection) {
  const list = assignedList || [];
  for (let i = 0; i < list.length; i++) {
    if (String(list[i] || '').trim() === String(classSection || '').trim()) return true;
  }
  return false;
}

// Distinct Daily_Hours_Setup dates in [startYMD, endYMD] for one class-section.
// Returns 0 when the sheet is missing (reports must not die on missing setup).
function countWorkingDaysInRange_(classSection, startYMD, endYMD) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Hours_Setup');
    if (!sheet) return 0;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0;
    const h = data[0].map(function (x) { return String(x).trim(); });
    const dateCol = h.indexOf('Date');
    const clsCol = h.indexOf('Class_Section');
    const dCol = dateCol !== -1 ? dateCol : 0;
    const cCol = clsCol !== -1 ? clsCol : 1;
    const setupRows = [];
    for (let i = 1; i < data.length; i++) {
      setupRows.push({ dateYMD: normalizeYMD_(data[i][dCol]), classSection: data[i][cCol] });
    }
    return countWorkingDays_(setupRows, classSection, startYMD, endYMD);
  } catch (e) {
    Logger.log('Error in countWorkingDaysInRange_: ' + e.message);
    return 0;
  }
}
// Working-day set in a range (pure). setupRows: [{dateYMD, classSection}]
// from Daily_Hours_Setup. Returns { 'yyyy-MM-dd': true } for distinct dates.
function workingDateSet_(setupRows, classSection, startYMD, endYMD) {
  const set = {};
  (setupRows || []).forEach(function (r) {
    if (String(r.classSection || '').trim() !== String(classSection || '').trim()) return;
    const k = String(r.dateYMD || '').trim();
    if (!k || k < startYMD || k > endYMD) return;
    set[k] = true;
  });
  return set;
}

// Daily_Hours_Setup dates in [startYMD, endYMD] for one class-section, as a set.
// Returns {} when the sheet is missing (reports must not die on missing setup).
function getWorkingDatesInRange_(classSection, startYMD, endYMD, schoolRecId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Hours_Setup');
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return {};
    const h = data[0].map(function (x) { return String(x).trim(); });
    const dateCol = h.indexOf('Date');
    const clsCol = h.indexOf('Class_Section');
    const recCol = h.indexOf('REC_ID');
    const dCol = dateCol !== -1 ? dateCol : 0;
    const cCol = clsCol !== -1 ? clsCol : 1;
    const setupRows = [];
    for (let i = 1; i < data.length; i++) {
      // v5.8 (O3): only this school's setup rows (blank REC_ID = any school;
      // no school passed = old behaviour, every row counts).
      if (!rowRecMatches_(recCol !== -1 ? data[i][recCol] : '', schoolRecId)) continue;
      setupRows.push({ dateYMD: normalizeYMD_(data[i][dCol]), classSection: data[i][cCol] });
    }
    return workingDateSet_(setupRows, classSection, startYMD, endYMD);
  } catch (e) {
    Logger.log('Error in getWorkingDatesInRange_: ' + e.message);
    return {};
  }
}

// Per-day counts for one student (pure). rows: [{stdId, dateYMD, status, hours, type}].
// Only rows in [startYMD, endYMD] count; Episodic rows are skipped. Unknown
// statuses are tracked as 'other' (shown as '·' in the grid) but never counted
// in totals — the same counting as aggregateMonthlyRows_.
function aggregateDayWiseRows_(rows, startYMD, endYMD) {
  const out = {};
  (rows || []).forEach(function (r) {
    const k = String(r.dateYMD || '').trim();
    if (!k || k < startYMD || k > endYMD) return;
    if (String(r.type || '').toLowerCase().trim() === 'episodic') return;
    const day = Number(k.slice(8, 10));
    if (!day) return;
    if (!out[day]) out[day] = { present: 0, absent: 0, leave: 0, other: 0 };
    const st = String(r.status || '').toLowerCase().trim();
    if (st === 'present') out[day].present++;
    else if (st === 'absent') out[day].absent++;
    else if (st === 'leave') out[day].leave++;
    else out[day].other++;
  });
  return out;
}

// Portal MonthlyBarReport cell symbol: P/P2, A/A2, L/L3, ·/·2 (priority
// Present > Absent > Leave > Other). '' when the day has no records.
function dayWiseCellSymbol_(c) {
  if (!c) return '';
  if (c.present > 0) return 'P' + (c.present > 1 ? c.present : '');
  if (c.absent > 0) return 'A' + (c.absent > 1 ? c.absent : '');
  if (c.leave > 0) return 'L' + (c.leave > 1 ? c.leave : '');
  if (c.other > 0) return '·' + (c.other > 1 ? c.other : '');
  return '';
}

// Build one student's Day Wise grid row (pure). agg comes from
// aggregateMonthlyRows_ (the same totals as the v3 Monthly report, so the two
// reports always agree); perDay comes from aggregateDayWiseRows_.
// Day symbols: 'P'/'P2'/'A'/'A2'/'L'/'L3'/'·'/'·2' when the day has records
// (a marked day is never dashed, even if it falls outside enrolment — the row
// still counts in the totals, portal rule); '—' when the day is outside the
// term-intersected range or the student was not enrolled that day;
// '×' when enrolled but not a working day (Daily_Hours_Setup); '-' when it is
// a working day with no record.
function dayWiseRow_(student, agg, perDay, workingSet, effStartYMD, effEndYMD, daysInMonth, monthStartYMD, monthEndYMD) {
  const pad2 = function (n) { return String(n).padStart(2, '0'); };
  const days = {};
  for (let d = 1; d <= daysInMonth; d++) {
    const ymd = String(monthStartYMD).slice(0, 8) + pad2(d);
    const c = perDay[d];
    if (c && (c.present > 0 || c.absent > 0 || c.leave > 0 || c.other > 0)) {
      days[d] = dayWiseCellSymbol_(c);
    } else if (ymd < effStartYMD || ymd > effEndYMD) {
      days[d] = '—';
    } else if (!isEnrolledOnYMD_(ymd, student.dojYMD, student.inactiveYMD)) {
      days[d] = '—';
    } else if (!workingSet[ymd]) {
      days[d] = '×';
    } else {
      days[d] = '-';
    }
  }
  const pct = agg.totalDays > 0 ? Math.round((agg.presents / agg.totalDays) * 100 * 100) / 100 : 0;
  return {
    studentId: student.studentId,
    studentName: student.studentName,
    className: student.className,
    section: student.section,
    status: monthStatusForYMD_(monthStartYMD, monthEndYMD, student.dojYMD, student.inactiveYMD),
    days: days,
    presents: agg.presents,
    absents: agg.absents,
    leaves: agg.leaves,
    totalDays: agg.totalDays,
    percentage: pct,
    hours: Math.round((agg.hours || 0) * 100) / 100
  };
}

function generateStudentAttendanceReport(date, classSection, termId, recId, username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get student data
    const studentSheet = ss.getSheetByName('Student_Data');
    if (!studentSheet) throw new Error('Student_Data sheet not found');

    const attendanceSheet = ss.getSheetByName('Student_Attendance');
    if (!attendanceSheet) throw new Error('Student_Attendance sheet not found');

    if (!classSection || String(classSection).indexOf('-') === -1) {
      throw new Error('Invalid class section: ' + classSection);
    }
    const parts = String(classSection).split('-');
    const className = String(parts[0] || '').trim();
    const section = String(parts[1] || '').trim();
    const targetDate = new Date(date).toLocaleDateString('en-CA');

    // Reports are scoped to the teacher's assignments: a direct call can never
    // show a class the user is not assigned to (the pickers already list
    // assigned classes only; getSummaryData enforces the same rule).
    // v5.6: a username is required — a call with no username is refused, not
    // silently allowed.
    if (!username) throw new Error('Please log in again to view this report.');
    const assigned = getAssignedClassSections(username, recId);
    if (!isClassSectionAssigned_(assigned, classSection)) {
      throw new Error('You are not assigned to class ' + classSection + '.');
    }

    // Term guard (portal rule): when a term is picked, the date must lie inside
    // this class-section's term range (first ACTIVE Class_Terms row wins).
    if (termId) {
      const range = getTermDateRangeForClass(classSection, termId, recId || null);
      if (!range) {
        throw new Error('No term dates set up for ' + classSection + ' in Term ' + termId + '. Add a Class_Terms row first.');
      }
      if (targetDate < range.fromDate || targetDate > range.toDate) {
        throw new Error('Selected date is outside Term ' + termId + ' for ' + classSection + ' (' + range.fromDate + ' to ' + range.toDate + ').');
      }
    }

    // ---- Roster: Student_Data, class/section + enrolled on that day (portal rule) ----
    const studentData = studentSheet.getDataRange().getValues();
    const studentHeaders = studentData[0];
    const shm = new Map(studentHeaders.map((h, i) => [h, i]));
    const stdIdCol = shm.get('Std_ID');
    const nameCol = shm.get('Student_Name');
    const classCol = shm.get('Student_Class');
    const sectionCol = shm.get('Student_Section');
    const dojCol = shm.has('Date_of_Joining') ? shm.get('Date_of_Joining') : -1;
    const inactCol = shm.has('Inactive_Date') ? shm.get('Inactive_Date') : -1;
    const recCol = shm.has('REC_ID') ? shm.get('REC_ID') : -1;
    if (stdIdCol === undefined || nameCol === undefined || classCol === undefined || sectionCol === undefined) {
      throw new Error('Required columns missing in Student_Data sheet');
    }

    const roster = [];
    for (let i = 1; i < studentData.length; i++) {
      const row = studentData[i];
      if (String(row[classCol] || '').trim() !== className) continue;
      if (String(row[sectionCol] || '').trim() !== section) continue;
      // v5.8 (O3): only this school's students (blank REC_ID = any school,
      // so rows from before REC_ID was recorded still show).
      if (recCol !== -1 && !rowRecMatches_(row[recCol], recId)) continue;
      const dojYMD = dojCol !== -1 ? normalizeYMD_(row[dojCol]) : '';
      const inactYMD = inactCol !== -1 ? normalizeYMD_(row[inactCol]) : '';
      if (!isEnrolledOnYMD_(targetDate, dojYMD, inactYMD)) continue;
      roster.push({ studentId: row[stdIdCol], studentName: row[nameCol] });
    }

    if (roster.length === 0) {
      throw new Error('No students enrolled in ' + classSection + ' on ' + date + '.');
    }

    // ---- Day rows: aggregate per student (portal rule) ----
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const headers = attendanceData[0];
    const headerMap = new Map(headers.map((h, i) => [h, i]));
    const aStdCol = headerMap.get('Std_ID');
    const aDateCol = headerMap.get('Date');
    const aStatusCol = headerMap.get('Status');
    const aClassCol = headerMap.get('Student_Class');
    const aSectionCol = headerMap.get('Student_Section');
    const aHoursCol = headerMap.has('Hours') ? headerMap.get('Hours') : -1;
    const aTypeCol = headerMap.has('Attendance_Type') ? headerMap.get('Attendance_Type') : -1;
    const aRecCol = headerMap.has('REC_ID') ? headerMap.get('REC_ID') : -1;
    if (aStdCol === undefined || aDateCol === undefined || aStatusCol === undefined || aClassCol === undefined || aSectionCol === undefined) {
      throw new Error('Required columns missing in Student_Attendance sheet');
    }

    const dayRows = [];
    for (let i = 1; i < attendanceData.length; i++) {
      const row = attendanceData[i];
      if (String(row[aClassCol] || '').trim() !== className) continue;
      if (String(row[aSectionCol] || '').trim() !== section) continue;
      // v5.8 (O3): only this school's attendance rows count — without this,
      // a same-Std_ID namesake at another school would feed this student's
      // status (the roster filter alone cannot fix the counts).
      if (aRecCol !== -1 && !rowRecMatches_(row[aRecCol], recId)) continue;
      if (normalizeYMD_(row[aDateCol]) !== targetDate) continue;
      dayRows.push({
        stdId: row[aStdCol],
        status: row[aStatusCol],
        hours: aHoursCol !== -1 ? row[aHoursCol] : 0,
        type: aTypeCol !== -1 ? row[aTypeCol] : ''
      });
    }
    const agg = aggregateDailyRows_(dayRows);

    // ---- Working hours for the day (Daily_Hours_Setup, normal type) ----
    let workingHours = 0;
    try {
      // v5.8 (O3): working hours from this school's setup rows only.
      let wh = checkDailyHours(date, classSection, 'normal', recId);
      if (!wh || !wh.exists) wh = checkDailyHours(date, classSection, null, recId);
      if (wh && wh.exists && !isNaN(Number(wh.hours))) workingHours = Number(wh.hours);
    } catch (e) {
      workingHours = 0; // no setup sheet/row: report still works, working hours 0
    }
    workingHours = Math.round(workingHours * 100) / 100;

    // ---- Build rows: full roster, one row per student ----
    const cleanName = function (s) { return String(s || '').replace(/[\u00A0\u200B-\u200D\uFEFF]/g, '').trim(); };
    const rows = roster.map(function (s) {
      const c = agg[String(s.studentId).trim()] || { present: 0, absent: 0, leave: 0, other: 0, hours: 0 };
      return {
        studentId: s.studentId,
        studentName: s.studentName,
        className: className,
        section: section,
        date: date,
        status: dailyStatusLabel_(c),
        hours: Math.round((c.hours || 0) * 100) / 100,
        workingHours: workingHours
      };
    });
    rows.sort(function (a, b) { return cleanName(a.studentName).localeCompare(cleanName(b.studentName), undefined, { sensitivity: 'base' }); });

    return rows;

  } catch (error) {
    console.error('Error generating student attendance report:', error);
    throw error;
  }
}

// NEW: Generate student monthly summary report
// Generate student monthly summary report (portal Monthly rules, v3 port)
function generateStudentMonthlySummary(studentId, month, year, termId, recId, username) {
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
    const dojCol = studentHeaderMap.has('Date_of_Joining') ? studentHeaderMap.get('Date_of_Joining') : -1;
    const inactCol = studentHeaderMap.has('Inactive_Date') ? studentHeaderMap.get('Inactive_Date') : -1;

    for (let i = 1; i < studentData.length; i++) {
      if (String(studentData[i][stdIdCol]) === String(studentId)) {
        studentInfo = {
          studentId: studentData[i][stdIdCol],
          studentName: studentData[i][nameCol],
          className: studentData[i][classCol],
          section: studentData[i][sectionCol],
          dojYMD: dojCol !== -1 ? normalizeYMD_(studentData[i][dojCol]) : '',
          inactiveYMD: inactCol !== -1 ? normalizeYMD_(studentData[i][inactCol]) : ''
        };
        break;
      }
    }

    if (!studentInfo) {
      throw new Error(`Student with ID ${studentId} not found`);
    }

    const classSection = studentInfo.className + '-' + studentInfo.section;

    // Reports are scoped to the teacher's assignments (same rule as the daily
    // report and getSummaryData).
    // v5.6: a username is required — a call with no username is refused, not
    // silently allowed.
    if (!username) throw new Error('Please log in again to view this report.');
    const assigned = getAssignedClassSections(username, recId);
    if (!isClassSectionAssigned_(assigned, classSection)) {
      throw new Error('You are not assigned to class ' + classSection + '.');
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
    const typeCol = headerMap.has('Attendance_Type') ? headerMap.get('Attendance_Type') : -1;

    // Calculate month start and end dates (end = last moment of the last day,
    // so rows on the last day are never cut off).
    let startDate = new Date(year, month - 1, 1, 0, 0, 0, 0);
    let endDate = new Date(year, month, 0, 23, 59, 59, 999);
    const monthStartYMD = startDate.getFullYear() + '-' + String(startDate.getMonth() + 1).padStart(2, '0') + '-' + String(startDate.getDate()).padStart(2, '0');
    const monthEndYMD = endDate.getFullYear() + '-' + String(endDate.getMonth() + 1).padStart(2, '0') + '-' + String(endDate.getDate()).padStart(2, '0');
    const monthName = new Date(year, month - 1).toLocaleString('default', { month: 'long' });

    // Not enrolled at any point in this month (portal roster rule) → zeros + warning.
    if (!isEnrolledInMonthYMD_(monthStartYMD, monthEndYMD, studentInfo.dojYMD, studentInfo.inactiveYMD)) {
      return {
        ...studentInfo, month: month, year: year,
        presents: 0, absents: 0, leaves: 0, totalDays: 0, presentPercentage: 0,
        hours: 0, workingDays: 0, status: 'Inactive',
        termId: termId ? String(termId).trim() : null, termName: null,
        warning: 'Student was not enrolled during ' + monthName + ' ' + year + '.'
      };
    }

    // ── Term-scoped filtering ──
    // If termId supplied: intersect month range with the term's date range (from Class_Terms).
    // Never reads Term_ID from attendance rows — filtering is by Date only.
    let appliedTermId = null, appliedTermName = null, appliedWarning = null;
    if (termId) {
      const range = getTermDateRangeForClass(classSection, termId, recId || null);
      if (range && range.fromDate && range.toDate) {
        const tFrom = new Date(range.fromDate + 'T00:00:00');
        const tTo   = new Date(range.toDate + 'T23:59:59');
        if (tFrom > startDate) startDate = tFrom;
        if (tTo   < endDate)   endDate   = tTo;
        const details = getTermDetails(termId) || {};
        appliedTermId = String(termId).trim();
        appliedTermName = details.termName || String(termId).trim();
        if (startDate > endDate) {
          appliedWarning = `Selected month is outside the term range ${range.fromDate} to ${range.toDate}. Report is empty.`;
          return {
            ...studentInfo, month: month, year: year,
            presents: 0, absents: 0, leaves: 0, totalDays: 0, presentPercentage: 0,
            hours: 0, workingDays: 0, status: monthStatusForYMD_(monthStartYMD, monthEndYMD, studentInfo.dojYMD, studentInfo.inactiveYMD),
            termId: appliedTermId, termName: appliedTermName, warning: appliedWarning
          };
        }
      }
    }
    const effStartYMD = normalizeYMD_(startDate);
    const effEndYMD = normalizeYMD_(endDate);

    // ---- Aggregate this student's rows (portal rule, single pass) ----
    // Only present/absent/leave count toward Total Days; unknown statuses are
    // ignored; hours count only on Present rows; Episodic rows are excluded.
    const rows = [];
    for (let i = 1; i < attendanceData.length; i++) {
      const row = attendanceData[i];
      if (String(row[stdIdColAttendance]) !== String(studentId)) continue;
      rows.push({
        stdId: row[stdIdColAttendance],
        dateYMD: normalizeYMD_(row[dateCol]),
        status: row[statusCol],
        hours: hoursCol !== -1 ? row[hoursCol] : 0,
        type: typeCol !== -1 ? row[typeCol] : ''
      });
    }
    const agg = aggregateMonthlyRows_(rows, effStartYMD, effEndYMD);

    // ---- Working days in the (term-intersected) month ----
    const workingDays = countWorkingDaysInRange_(classSection, effStartYMD, effEndYMD);

    // Calculate percentage
    const presentPercentage = agg.totalDays > 0 ? (agg.presents / agg.totalDays) * 100 : 0;

    return {
      ...studentInfo,
      month: month,
      year: year,
      presents: agg.presents,
      absents: agg.absents,
      leaves: agg.leaves,
      totalDays: agg.totalDays,
      presentPercentage: Math.round(presentPercentage * 100) / 100,
      hours: Math.round((agg.hours || 0) * 100) / 100,
      workingDays: workingDays,
      status: monthStatusForYMD_(monthStartYMD, monthEndYMD, studentInfo.dojYMD, studentInfo.inactiveYMD),
      termId: appliedTermId,
      termName: appliedTermName,
      warning: appliedWarning
    };

  } catch (error) {
    console.error('Error generating monthly summary:', error);
    throw error;
  }
}

// Monthly Day Wise report (v4 port of the MIS portal's MonthlyBarReport).
// One row per student, one column per day of the month, plus per-student totals
// (Present/Absent/Leave, Total Days, Total %, Total Hours) and a class total.
// Same rules as the v3 Monthly: Student_Data roster (enrolled at some point in
// the month), Student_Attendance rows, only Present earns hours, unknown
// statuses ignored in totals, term intersected from Class_Terms, Episodic rows
// excluded, single assigned class only (no multi-class selection).
// Percentage is 2dp like the v3 Monthly (the portal Day Wise rounds to whole %).
function generateStudentDayWiseReport(month, year, classSection, termId, recId, username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const studentSheet = ss.getSheetByName('Student_Data');
    if (!studentSheet) throw new Error('Student_Data sheet not found');

    const attendanceSheet = ss.getSheetByName('Student_Attendance');
    if (!attendanceSheet) throw new Error('Student_Attendance sheet not found');

    if (!classSection || String(classSection).indexOf('-') === -1) {
      throw new Error('Invalid class section: ' + classSection);
    }
    const parts = String(classSection).split('-');
    const className = String(parts[0] || '').trim();
    const section = String(parts[1] || '').trim();

    // Reports are scoped to the teacher's assignments: a direct call can never
    // show a class the user is not assigned to (the pickers already list
    // assigned classes only; same rule as the v3 Daily/Monthly reports).
    // v5.6: a username is required — a call with no username is refused, not
    // silently allowed.
    if (!username) throw new Error('Please log in again to view this report.');
    const assigned = getAssignedClassSections(username, recId);
    if (!isClassSectionAssigned_(assigned, classSection)) {
      throw new Error('You are not assigned to class ' + classSection + '.');
    }

    month = Number(month); year = Number(year);
    if (!month || month < 1 || month > 12 || !year) throw new Error('Invalid month/year: ' + month + '/' + year);
    const daysInMonth = new Date(year, month, 0).getDate();
    const pad2 = function (n) { return String(n).padStart(2, '0'); };
    const monthStartYMD = year + '-' + pad2(month) + '-01';
    const monthEndYMD = year + '-' + pad2(month) + '-' + pad2(daysInMonth);
    const monthName = new Date(year, month - 1).toLocaleString('default', { month: 'long' });

    // ---- Term: intersect the month with the class-section's term range ----
    // (first ACTIVE Class_Terms row wins, same as the v3 Monthly).
    let effStartYMD = monthStartYMD, effEndYMD = monthEndYMD;
    let appliedTermId = null, appliedTermName = null;
    if (termId) {
      const range = getTermDateRangeForClass(classSection, termId, recId || null);
      if (!range) {
        throw new Error('No term dates set up for ' + classSection + ' in Term ' + termId + '. Add a Class_Terms row first.');
      }
      if (range.fromDate > effStartYMD) effStartYMD = range.fromDate;
      if (range.toDate < effEndYMD) effEndYMD = range.toDate;
      const details = getTermDetails(termId) || {};
      appliedTermId = String(termId).trim();
      appliedTermName = details.termName || String(termId).trim();
      if (effStartYMD > effEndYMD) {
        return {
          month: month, year: year, monthName: monthName,
          classSection: classSection, className: className, section: section,
          daysInMonth: daysInMonth, workingDays: 0,
          termId: appliedTermId, termName: appliedTermName,
          effStartYMD: effStartYMD, effEndYMD: effEndYMD,
          rows: [],
          totals: { presents: 0, absents: 0, leaves: 0, totalDays: 0, hours: 0, percentage: 0, workingDays: 0 },
          warning: 'Selected month is outside the term range ' + range.fromDate + ' to ' + range.toDate + '. Report is empty.'
        };
      }
    }

    // ---- Roster: Student_Data, class/section + enrolled at some point in the month ----
    const studentData = studentSheet.getDataRange().getValues();
    const studentHeaders = studentData[0];
    const shm = new Map(studentHeaders.map((h, i) => [h, i]));
    const stdIdCol = shm.get('Std_ID');
    const nameCol = shm.get('Student_Name');
    const classCol = shm.get('Student_Class');
    const sectionCol = shm.get('Student_Section');
    const dojCol = shm.has('Date_of_Joining') ? shm.get('Date_of_Joining') : -1;
    const inactCol = shm.has('Inactive_Date') ? shm.get('Inactive_Date') : -1;
    const recCol = shm.has('REC_ID') ? shm.get('REC_ID') : -1;
    if (stdIdCol === undefined || nameCol === undefined || classCol === undefined || sectionCol === undefined) {
      throw new Error('Required columns missing in Student_Data sheet');
    }

    const roster = [];
    for (let i = 1; i < studentData.length; i++) {
      const row = studentData[i];
      if (String(row[classCol] || '').trim() !== className) continue;
      if (String(row[sectionCol] || '').trim() !== section) continue;
      // v5.8 (O3): only this school's students (blank REC_ID = any school).
      if (recCol !== -1 && !rowRecMatches_(row[recCol], recId)) continue;
      const dojYMD = dojCol !== -1 ? normalizeYMD_(row[dojCol]) : '';
      const inactYMD = inactCol !== -1 ? normalizeYMD_(row[inactCol]) : '';
      if (!isEnrolledInMonthYMD_(monthStartYMD, monthEndYMD, dojYMD, inactYMD)) continue;
      roster.push({ studentId: row[stdIdCol], studentName: row[nameCol], className: className, section: section, dojYMD: dojYMD, inactiveYMD: inactYMD });
    }

    if (roster.length === 0) {
      throw new Error('No students enrolled in ' + classSection + ' during ' + monthName + ' ' + year + '.');
    }

    // ---- Month rows: attendance for the class-section in the month ----
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const headers = attendanceData[0];
    const headerMap = new Map(headers.map((h, i) => [h, i]));
    const aStdCol = headerMap.get('Std_ID');
    const aDateCol = headerMap.get('Date');
    const aStatusCol = headerMap.get('Status');
    const aClassCol = headerMap.get('Student_Class');
    const aSectionCol = headerMap.get('Student_Section');
    const aHoursCol = headerMap.has('Hours') ? headerMap.get('Hours') : -1;
    const aTypeCol = headerMap.has('Attendance_Type') ? headerMap.get('Attendance_Type') : -1;
    const aRecCol = headerMap.has('REC_ID') ? headerMap.get('REC_ID') : -1;
    if (aStdCol === undefined || aDateCol === undefined || aStatusCol === undefined || aClassCol === undefined || aSectionCol === undefined) {
      throw new Error('Required columns missing in Student_Attendance sheet');
    }

    const rowsByStudent = {};
    for (let i = 1; i < attendanceData.length; i++) {
      const row = attendanceData[i];
      if (String(row[aClassCol] || '').trim() !== className) continue;
      if (String(row[aSectionCol] || '').trim() !== section) continue;
      // v5.8 (O3): only this school's attendance rows count.
      if (aRecCol !== -1 && !rowRecMatches_(row[aRecCol], recId)) continue;
      const dateYMD = normalizeYMD_(row[aDateCol]);
      if (!dateYMD || dateYMD < monthStartYMD || dateYMD > monthEndYMD) continue;
      const sid = String(row[aStdCol] || '').trim();
      if (!sid) continue;
      if (!rowsByStudent[sid]) rowsByStudent[sid] = [];
      rowsByStudent[sid].push({
        stdId: row[aStdCol],
        dateYMD: dateYMD,
        status: row[aStatusCol],
        hours: aHoursCol !== -1 ? row[aHoursCol] : 0,
        type: aTypeCol !== -1 ? row[aTypeCol] : ''
      });
    }

    // ---- Working days in the (term-intersected) month ----
    // v5.8 (O3): working days from this school's setup rows only.
    const workingSet = getWorkingDatesInRange_(classSection, effStartYMD, effEndYMD, recId);
    const workingDays = Object.keys(workingSet).length;

    // ---- Build the grid: one row per student, one cell per day ----
    const cleanName = function (s) { return String(s || '').replace(/[\u00A0\u200B-\u200D\uFEFF]/g, '').trim(); };
    const rows = roster.map(function (s) {
      const list = rowsByStudent[String(s.studentId).trim()] || [];
      const agg = aggregateMonthlyRows_(list, effStartYMD, effEndYMD);
      const perDay = aggregateDayWiseRows_(list, effStartYMD, effEndYMD);
      return dayWiseRow_(s, agg, perDay, workingSet, effStartYMD, effEndYMD, daysInMonth, monthStartYMD, monthEndYMD);
    });
    rows.sort(function (a, b) { return cleanName(a.studentName).localeCompare(cleanName(b.studentName), undefined, { sensitivity: 'base' }); });

    // ---- Class totals (sums; % = total presents / total days) ----
    const tPresents = rows.reduce(function (sum, r) { return sum + r.presents; }, 0);
    const tAbsents = rows.reduce(function (sum, r) { return sum + r.absents; }, 0);
    const tLeaves = rows.reduce(function (sum, r) { return sum + r.leaves; }, 0);
    const tDays = rows.reduce(function (sum, r) { return sum + r.totalDays; }, 0);
    const tHours = rows.reduce(function (sum, r) { return sum + (r.hours || 0); }, 0);
    const totals = {
      presents: tPresents,
      absents: tAbsents,
      leaves: tLeaves,
      totalDays: tDays,
      hours: Math.round(tHours * 100) / 100,
      percentage: tDays > 0 ? Math.round((tPresents / tDays) * 100 * 100) / 100 : 0,
      workingDays: workingDays
    };

    return {
      month: month, year: year, monthName: monthName,
      classSection: classSection, className: className, section: section,
      daysInMonth: daysInMonth, workingDays: workingDays,
      termId: appliedTermId, termName: appliedTermName,
      effStartYMD: effStartYMD, effEndYMD: effEndYMD,
      rows: rows, totals: totals, warning: null
    };

  } catch (error) {
    console.error('Error generating monthly day wise report:', error);
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
    const dojCol = headerMap.has('Date_of_Joining') ? headerMap.get('Date_of_Joining') : -1;
    const inactCol = headerMap.has('Inactive_Date') ? headerMap.get('Inactive_Date') : -1;

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
        classSection: `${studentClass}-${studentSection}`,
        dateOfJoining: dojCol !== -1 ? (normalizeYMD_(row[dojCol]) || null) : null,
        inactiveDate: inactCol !== -1 ? (normalizeYMD_(row[inactCol]) || null) : null
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

// Report Term picker data (v3 port): all Terms rows for the picker.
// Terms holds names only; dates live in Class_Terms. Returns [] when the
// sheet is missing or empty — never throws, so reports keep working
// (no term picked = no term filtering, portal rule).
function getReportTerms() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Terms');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const h = data[0].map(function (x) { return String(x).trim(); });
    const idCol = h.indexOf('Term_ID');
    if (idCol === -1) return [];
    const nameCol = h.indexOf('Term_Name');
    const actCol = h.indexOf('Is_Active');
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][idCol] || '').trim();
      if (!id) continue;
      out.push({
        Term_ID: id,
        Term_Name: nameCol !== -1 ? String(data[i][nameCol] || '').trim() : id,
        Is_Active: actCol !== -1 ? data[i][actCol] : false
      });
    }
    return out;
  } catch (e) {
    Logger.log('Error in getReportTerms: ' + e.message);
    return [];
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
      bottom: 95px;
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
      bottom: 200px;
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
  
<button onclick="window.close()" style="position:absolute; bottom:20px; left:20px; right:20px; height:60px; border:none; border-radius:14px; background:#2e7d32; color:#fff; font-size:18px; font-weight:700; z-index:40; cursor:pointer; box-shadow:0 4px 14px rgba(0,0,0,0.4);">✓ DONE — CLOSE SCANNER</button>
<button onclick="window.close()" style="position:absolute; top:16px; right:16px; width:44px; height:44px; border-radius:50%; background:#ea4335; color:#fff; font-size:22px; border:none; z-index:40; cursor:pointer;">✕</button>


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

// If the parent page that launched us is gone, close this scanner
setInterval(function() {
  if (!window.opener || window.opener.closed) {
    try { window.close(); } catch (e) {}
  }
}, 2000);

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
  try {
    if (window.opener && !window.opener.closed) {
      window.opener.postMessage({ type: "SCANNER_CLOSED" }, "*");
    }
  } catch (e) {}
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

// ================= EPISODIC EVENTS (CRUD) =================
const EPISODIC_EVENTS_SHEET = 'Episodic_Events';
const EPISODIC_EVENTS_HEADERS = ['Event_ID','Event_Name','Event_Date','Category','Hours','Class_Sections','Username','REC_ID','Timestamp'];

// Sheet get-or-create (headers ke saath auto-banati hai agar mojood nahi)
function getOrCreateEpisodicEventsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(EPISODIC_EVENTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(EPISODIC_EVENTS_SHEET);
    sheet.getRange(1, 1, 1, EPISODIC_EVENTS_HEADERS.length).setValues([EPISODIC_EVENTS_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normalizeYMD_(v) {
  if (!v) return '';
  try { return new Date(v).toLocaleDateString('en-CA'); } catch (e) { return String(v); }
}

// CREATE
function createEpisodicEvent(payload) {
  // payload = { eventName, date, category, hours, classSections:[...], username, recId }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('System busy. Please try again.');
  try {
    if (!payload || !payload.eventName || !payload.date || !payload.category ||
        !payload.hours || !payload.classSections || payload.classSections.length === 0) {
      throw new Error('All fields are required.');
    }
    // v5.6: caller verification, assignment for EVERY listed class, school from
    // the Users row, and full input validation (future dates ARE allowed here —
    // events are planned ahead).
    const caller = requireTeacherCaller_(payload.username, payload.recId);
    const schoolRecId = caller.recId || String(payload.recId || '').trim();
    requireAssignedClasses_(payload.username, payload.classSections, schoolRecId);
    requirePermissionFlag_(payload.username, 'canManageEpisodicEvents');
    const evDate = String(payload.date).trim();
    if (!isRealYMD_(evDate)) throw new Error('Invalid event date: ' + payload.date + '. Use a real yyyy-MM-dd date.');
    validateHoursValue_(payload.hours, 'episodic');
    validateCategoryForClass_(payload.category, 'episodic', payload.classSections);
    const sheet = getOrCreateEpisodicEventsSheet_();
    // Duplicate: same name (case-insensitive) + same date + an overlapping class.
    const evData = sheet.getDataRange().getValues();
    if (evData.length > 1) {
      const evH = evData[0];
      const evNameCol = evH.indexOf('Event_Name'), evDateCol = evH.indexOf('Event_Date'),
            evClsCol = evH.indexOf('Class_Sections'), evRecCol = evH.indexOf('REC_ID');
      const wantName = String(payload.eventName).trim().toLowerCase();
      for (let i = 1; i < evData.length; i++) {
        if (String(evData[i][evNameCol] || '').trim().toLowerCase() !== wantName) continue;
        if (normalizeYMD_(evData[i][evDateCol]) !== evDate) continue;
        const existingCls = String(evData[i][evClsCol] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
        const overlap = payload.classSections.filter(function (cs) { return existingCls.indexOf(cs) !== -1; });
        if (overlap.length === 0) continue;
        if (!rowRecMatches_(evRecCol !== -1 ? evData[i][evRecCol] : '', schoolRecId)) continue;
        throw new Error('An event named "' + payload.eventName + '" already exists on ' + evDate + ' for ' + overlap.join(', ') + '.');
      }
    }
    sheet.appendRow([
      Utilities.getUuid(),
      payload.eventName,
      evDate,
      payload.category,
      payload.hours,
      payload.classSections.join(','),
      payload.username,
      schoolRecId,
      new Date().toLocaleString()
    ]);
    SpreadsheetApp.flush();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// LIST (sirf isi user ke banaye events)
function getEpisodicEvents(username) {
  const sheet = getOrCreateEpisodicEventsSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const h = data[0];
  const c = {
    id: h.indexOf('Event_ID'), name: h.indexOf('Event_Name'), date: h.indexOf('Event_Date'),
    cat: h.indexOf('Category'), hrs: h.indexOf('Hours'), cls: h.indexOf('Class_Sections'),
    user: h.indexOf('Username')
  };
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][c.user]) !== String(username)) continue;
    out.push({
      eventId: data[i][c.id],
      eventName: data[i][c.name],
      date: normalizeYMD_(data[i][c.date]),
      category: data[i][c.cat],
      hours: data[i][c.hrs],
      classSections: data[i][c.cls]
    });
  }
  return out;
}

// DELETE (Event_ID se)
// DELETE (Event_ID se) — v5.6: caller verified, creator-or-admin only, and
// refused when attendance is already marked under the event.
function deleteEpisodicEvent(eventId, username) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('System busy. Please try again.');
  try {
    const ev = findEpisodicEventById_(eventId);
    if (!ev) return { success: false, message: 'Event not found' };
    const caller = requireTeacherCaller_(username, ev.recId);
    const schoolRecId = caller.recId || ev.recId || '';
    requirePermissionFlag_(username, 'canManageEpisodicEvents');
    if (!rowRecMatches_(ev.recId, schoolRecId)) return { success: false, message: 'Event not found' };
    if (ev.username !== username && caller.userType.toLowerCase() !== 'admin') {
      return { success: false, message: 'Only the teacher who created this event (or an administrator) can delete it.' };
    }
    if (markedSectionsForEvent_(ev.eventName, ev.eventDate, schoolRecId).length > 0) {
      return { success: false, message: 'Attendance has already been marked for this event, so it cannot be deleted.' };
    }
    const sheet = getOrCreateEpisodicEventsSheet_();
    sheet.deleteRow(ev.rowIndex);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}
// Ek date + class-section ke liye events (Episodic Hours page ke dropdown ke liye)
function getEpisodicEventsForDate(classSection, date) {
  const sheet = getOrCreateEpisodicEventsSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const h = data[0];
  const c = {
    id: h.indexOf('Event_ID'), name: h.indexOf('Event_Name'), date: h.indexOf('Event_Date'),
    cat: h.indexOf('Category'), hrs: h.indexOf('Hours'), cls: h.indexOf('Class_Sections')
  };
  const target = normalizeYMD_(date);
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (normalizeYMD_(data[i][c.date]) !== target) continue;
    const sections = String(data[i][c.cls] || '').split(',').map(s => s.trim());
    if (sections.indexOf(classSection) === -1) continue;
    out.push({
      eventId: data[i][c.id],
      eventName: data[i][c.name],
      category: data[i][c.cat],
      hours: data[i][c.hrs]
    });
  }
  return out;
}

// Update: sirf Class_Sections change hote hain (baaki fields lock hain)
// Update: sirf Class_Sections change hote hain (baaki fields lock hain)
// v5.6: caller verified, creator-or-admin only, every new class assigned,
// and a class with marked attendance cannot be removed.
function updateEpisodicEvent(eventId, classSections, username) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('System busy. Please try again.');
  try {
    if (!classSections || classSections.length === 0) throw new Error('At least one class section is required.');
    const ev = findEpisodicEventById_(eventId);
    if (!ev) return { success: false, message: 'Event not found' };
    const caller = requireTeacherCaller_(username, ev.recId);
    const schoolRecId = caller.recId || ev.recId || '';
    requirePermissionFlag_(username, 'canManageEpisodicEvents');
    requireAssignedClasses_(username, classSections, schoolRecId);
    if (!rowRecMatches_(ev.recId, schoolRecId)) return { success: false, message: 'Event not found' };
    if (ev.username !== username && caller.userType.toLowerCase() !== 'admin') {
      return { success: false, message: 'Only the teacher who created this event (or an administrator) can change it.' };
    }
    const marked = markedSectionsForEvent_(ev.eventName, ev.eventDate, schoolRecId);
    const stranded = marked.filter(function (s) { return classSections.indexOf(s) === -1; });
    if (stranded.length > 0) {
      return { success: false, message: 'Attendance is already marked for ' + stranded.join(', ') + ' under this event. Those classes cannot be removed.' };
    }
    const sheet = getOrCreateEpisodicEventsSheet_();
    const h = sheet.getDataRange().getValues()[0];
    const clsCol = h.indexOf('Class_Sections');
    const tsCol = h.indexOf('Timestamp');
    sheet.getRange(ev.rowIndex, clsCol + 1).setValue(classSections.join(','));
    if (tsCol !== -1) sheet.getRange(ev.rowIndex, tsCol + 1).setValue(new Date().toLocaleString());
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}
// Is event ke liye jin class-sections ki attendance already mark ho chuki hai (lock karne ke liye)
function getMarkedClassSectionsForEvent(eventName, date) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Student_Attendance');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const h = data[0];
  const dateCol = h.indexOf('Date');
  const classCol = h.indexOf('Class_Section');
  const typeCol = h.indexOf('Attendance_Type');
  const eventCol = h.indexOf('Event_Name');
  if (dateCol === -1 || classCol === -1 || eventCol === -1) return [];
  const target = normalizeYMD_(date);
  const marked = new Set();
  for (let i = 1; i < data.length; i++) {
    if (typeCol !== -1 && String(data[i][typeCol] || '').trim().toLowerCase() !== 'episodic') continue;
    if (String(data[i][eventCol]) !== String(eventName)) continue;
    let rowDate;
    try { rowDate = new Date(data[i][dateCol]).toLocaleDateString('en-CA'); } catch (e) { continue; }
    if (rowDate !== target) continue;
    if (data[i][classCol]) marked.add(data[i][classCol]);
  }
  return Array.from(marked);
}


// ==================== TERM REPORTING — PHASE 1 (ADDITIONS ONLY) ====================
// These are pure additions. They do not modify any existing function.
// They are not called by any existing code. Nothing in the app changes behavior.
// Phase 2 (later) will wire the frontend to these.

/**
 * Parse a Class_Terms.Class_Section cell into tokens.
 * Accepts: "10-A" | "10-A,10-B" | "{I,II,IV}" | "{I,II,III-B},HRE-A"
 * Returns: ["10-A"] | ["10-A","10-B"] | ["I","II","IV"] | ["I","II","III-B","HRE-A"]
 */
function parseClassCell_(cell) {
  let s = String(cell || '').trim();
  if (!s) return [];
  s = s.replace(/[{}]/g, '');
  return s.split(',').map(t => t.trim()).filter(t => t.length > 0);
}

/**
 * Does one token match a concrete class-section?
 *   "I-A" vs "I-A"    → true   (exact section)
 *   "I"   vs "I-A"    → true   (bare class, any section)
 *   "I"   vs "IBT-A"  → false  (different class)
 *   "I-B" vs "I-A"    → false
 */
function tokenMatches_(token, classSection) {
  if (!token || !classSection) return false;
  if (token.indexOf('-') !== -1) return token === classSection;
  const dash = classSection.indexOf('-');
  const cls  = dash === -1 ? classSection : classSection.slice(0, dash);
  return token === cls;
}

/**
 * Does a whole Class_Terms.Class_Section cell match a concrete class-section?
 */
function classCellMatches_(cell, classSection) {
  const tokens = parseClassCell_(cell);
  for (let i = 0; i < tokens.length; i++) {
    if (tokenMatches_(tokens[i], classSection)) return true;
  }
  return false;
}

/**
 * Get all terms ever assigned to a given class-section.
 * Returns newest-start-first. Reads Class_Terms, expands brace/comma tokens.
 */
function getTermsForClassSection(classSection, recId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_Terms');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const h = data[0].map(x => String(x).trim());
    const clsCol   = h.indexOf('Class_Section');
    const termCol  = h.indexOf('Term_ID');
    const startCol = h.indexOf('Start_Date');
    const endCol   = h.indexOf('End_Date');
    const activeCol = h.indexOf('Is_Active');
    const recCol   = h.indexOf('REC_ID');
    if (clsCol === -1 || termCol === -1 || startCol === -1 || endCol === -1) return [];

    const out = [];
    const seen = new Set();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[termCol]) continue;
      if (!classCellMatches_(row[clsCol], classSection)) continue;
      if (recId && recCol !== -1 && String(row[recCol]).trim() !== String(recId).trim()) continue;

      const termId = String(row[termCol]).trim();
      const startYMD = normalizeYMD_(row[startCol]);
      const endYMD   = normalizeYMD_(row[endCol]);
      const key = termId + '|' + startYMD + '|' + endYMD;

      if (seen.has(key)) continue;
      seen.add(key);

      const termDetails = getTermDetails(termId) || {};
      out.push({
        termId: termId,
        termName: termDetails.termName || termId,
        startDate: startYMD,
        endDate: endYMD,
        isActive: activeCol !== -1 ? (row[activeCol] === true || row[activeCol] === 'TRUE') : false
      });
    }

    out.sort((a, b) => String(b.startDate).localeCompare(String(a.startDate)));
    return out;
  } catch (e) {
    Logger.log('Error in getTermsForClassSection: ' + e.message);
    return [];
  }
}

/**
 * Get the date range for (classSection, termId).
 * Portal rule (v4): FIRST matching ACTIVE Class_Terms row wins (sheet order),
 * REC-matched when recId is given. No merging across rows or RECs.
 * Returns {termId, fromDate, toDate, segments} or null.
 */
function getTermDateRangeForClass(classSection, termId, recId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_Terms');
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return null;

    const h = data[0].map(x => String(x).trim());
    const clsCol   = h.indexOf('Class_Section');
    const termCol  = h.indexOf('Term_ID');
    const startCol = h.indexOf('Start_Date');
    const endCol   = h.indexOf('End_Date');
    const activeCol = h.indexOf('Is_Active');
    const recCol   = h.indexOf('REC_ID');
    if (clsCol === -1 || termCol === -1 || startCol === -1 || endCol === -1) return null;

    const rows = [];
    for (let i = 1; i < data.length; i++) {
      rows.push({
        Class_Section: data[i][clsCol],
        Term_ID: data[i][termCol],
        Is_Active: activeCol !== -1 ? data[i][activeCol] : false,
        Start_Date: data[i][startCol],
        End_Date: data[i][endCol],
        REC_ID: recCol !== -1 ? data[i][recCol] : ''
      });
    }
    const picked = pickTermRange_(rows, classSection, termId, recId);
    if (!picked) return null;

    return {
      termId: picked.termId,
      fromDate: picked.fromDate,
      toDate: picked.toDate,
      segments: [{ fromDate: picked.fromDate, toDate: picked.toDate, isActive: true }]
    };
  } catch (e) {
    Logger.log('Error in getTermDateRangeForClass: ' + e.message);
    return null;
  }
}

// ==================== END TERM REPORTING PHASE 1 ====================



// ==================== TERM REPORTING — PHASE 1 TESTS ====================

function testParseClassCell() {
  const cases = [
    '10-A',
    '10-A,10-B,10-C',
    '{I,II,IV}',
    '{I,II,III-B},HRE-A,V-C',
    '{I,II,III-B,IV,V-C,VI,HRE,XI,VIII-A}',
    '  { I , II } ',
    ''
  ];
  cases.forEach(c => {
    Logger.log(JSON.stringify(c) + '  →  ' + JSON.stringify(parseClassCell_(c)));
  });
}

function testClassCellMatches() {
  const cell = '{I,II,III-B,IV,V-C,VI,HRE,XI,VIII-A}';
  const targets = [
    'I-A', 'I-B', 'II-A', 'III-A', 'III-B', 'III-C',
    'IV-A', 'IV-D', 'V-A', 'V-C', 'VI-B', 'HRE-A',
    'XI-D', 'VIII-A', 'VIII-B', 'VII-A', 'IBT-A'
  ];
  targets.forEach(t => {
    Logger.log(t + '  →  ' + classCellMatches_(cell, t));
  });
}

function testGetTermsForClassSection() {
  Logger.log(JSON.stringify(getTermsForClassSection('10-A', null), null, 2));
  Logger.log(JSON.stringify(getTermsForClassSection('I-A', null), null, 2));
  Logger.log(JSON.stringify(getTermsForClassSection('VII-A', null), null, 2));
  Logger.log(JSON.stringify(getTermsForClassSection('HRE-A', null), null, 2));
}

function testGetTermDateRange() {
  Logger.log(JSON.stringify(getTermDateRangeForClass('10-A', '2025-2026', null), null, 2));
  Logger.log(JSON.stringify(getTermDateRangeForClass('I-A', '2025-2026', null), null, 2));
  Logger.log(JSON.stringify(getTermDateRangeForClass('VII-A', '2025-2026', null), null, 2));
  Logger.log(JSON.stringify(getTermDateRangeForClass('9-C', '2025-2026-A', null), null, 2));
}


// ==================== TERM REPORTING — PHASE 2 ====================

/**
 * Check whether a specific date falls inside a term's date range for a class-section.
 * Read-only. Returns { termId, termName, fromDate, toDate, dateInsideRange, warning }
 * or null if the term range can't be resolved.
 * Used by the Daily report to decide whether to show an "outside term" warning.
 */
function checkDateInTerm(classSection, termId, date) {
  try {
    if (!classSection || !termId || !date) return null;
    const range = getTermDateRangeForClass(classSection, termId, null);
    if (!range) return null;
    const dateYMD = normalizeYMD_(date);
    const inside = dateYMD >= range.fromDate && dateYMD <= range.toDate;
    const details = getTermDetails(termId) || {};
    return {
      termId: String(termId).trim(),
      termName: details.termName || String(termId).trim(),
      fromDate: range.fromDate,
      toDate: range.toDate,
      dateInsideRange: inside,
      warning: inside
        ? null
        : `Selected date ${dateYMD} is outside the term range ${range.fromDate} to ${range.toDate}. Report may be empty.`
    };
  } catch (e) {
    Logger.log('Error in checkDateInTerm: ' + e.message);
    return null;
  }
}



function TEST_STEP1() {
  // v5.6: reports refuse calls with no username — use a real teacher username
  // assigned to student 3543's class.
  var TEST_USERNAME = 'REPLACE_WITH_A_TEACHER_USERNAME';
  Logger.log('=== A. checkDateInTerm — I-A, 2025-2026, 2025-11-15 (expect inside:true) ===');
  Logger.log(JSON.stringify(checkDateInTerm('I-A', '2025-2026', '2025-11-15'), null, 2));

  Logger.log('=== B. checkDateInTerm — I-A, 2025-2026, 2024-03-01 (expect inside:false) ===');
  Logger.log(JSON.stringify(checkDateInTerm('I-A', '2025-2026', '2024-03-01'), null, 2));

  Logger.log('=== C. checkDateInTerm — VIII-A, 2026-2027, 2026-08-15 (expect inside:true) ===');
  Logger.log(JSON.stringify(checkDateInTerm('VIII-A', '2026-2027', '2026-08-15'), null, 2));

  Logger.log('=== D. checkDateInTerm — I-A, 2026-2027, 2026-08-15 (expect inside:false) ===');
  Logger.log(JSON.stringify(checkDateInTerm('I-A', '2026-2027', '2026-08-15'), null, 2));

  // Replace 3543 with a real Std_ID that has attendance in Oct 2025
  Logger.log('=== E. Monthly — no term (should behave like today) ===');
  Logger.log(JSON.stringify(generateStudentMonthlySummary('3543', 10, 2025, null, null, TEST_USERNAME), null, 2));

  Logger.log('=== F. Monthly — WITH term 2025-2026, Oct 2025 (expect same counts) ===');
  Logger.log(JSON.stringify(generateStudentMonthlySummary('3543', 10, 2025, '2025-2026', null, TEST_USERNAME), null, 2));

  Logger.log('=== G. Monthly — WITH term 2025-2026, Mar 2024 (expect 0s + warning) ===');
  Logger.log(JSON.stringify(generateStudentMonthlySummary('3543', 3, 2024, '2025-2026', null, TEST_USERNAME), null, 2));
}


// ============ SHEET STRUCTURE SELF-HEAL (v5.9) ============
// The 65 schools each have their own spreadsheet, made at different
// times, so older sheets are missing newer columns. This block checks
// every sheet the app uses against the expected columns below and
// APPENDS any missing column at the END of the header row. It never
// inserts in the middle, never renames, never deletes and never touches
// a data row -- a new column simply starts blank, and the code already
// reads blanks safely (a blank right is No, a blank Is_Active is
// active). A column that exists under a slightly different name is NOT
// guessed at: it is reported as a possible mismatch for a human. For
// the few sheets that positional readers depend on (fixed: true), the
// present columns must also be in the expected relative order, or
// nothing is appended to that sheet and a warning is logged instead.
// Every addition and warning is written to the Structure_Log sheet,
// which is created lazily on the FIRST entry only -- a school that
// never needs one never gets the sheet. Sheets that do not exist are
// never created here; they appear when their feature is first used,
// exactly as before. setupSpreadsheet() does not call any of this.
// The check runs quietly once per ~6 hours at login (CacheService
// flag), and in the portal also on demand from the Sheet Structure
// tab (runStructureCheck). Failures are swallowed -- structure work
// must never break a login.
const STRUCTURE_LOG_SHEET = 'Structure_Log';
const STRUCTURE_LOG_HEADERS = ['Timestamp', 'App', 'Sheet_Name', 'Detail', 'Via'];
const STRUCTURE_CACHE_KEY_ = 'structureChecked_v1';
const STRUCTURE_CACHE_SECONDS_ = 6 * 3600;

// The expected columns per sheet, in order. Derived from what the
// portal + attendance code reads or writes by header name, plus the
// fixed layouts. fixed: true marks the sheets whose columns positional
// readers depend on:
//   Users -- the portal login (authenticateUser) reads indexes
//     0,1,2,3,7,8,9,10,15,16.
//   User_Permissions -- the attendance app's getUserPermissions reads
//     fixed indexes 0-8, 10, 11 (canMultiCategory is read by name).
//     Index 9 is canCreateEpisodic (the one portal-written flag not
//     otherwise placed; nothing reads it by position).
//   Previous_Date_Permissions -- the attendance app reads Username,
//     Allowed_Date and Password positionally (columns 0, 1, 2).
//   Daily_Hours_Setup -- the attendance app reads Date, Class_Section,
//     Category, Hours positionally and appends new rows positionally.
//   RECNames -- the portal's getRecInfo_ reads LB / REC Code /
//     REC Name at indexes 2, 3, 4.
const MARKING_STRUCTURE_SCHEMA_ = [
  { sheet: 'Users', fixed: true, columns: ['Username', 'Password', 'UserType', 'Teacher_Name', 'AllowedDays', 'StartTime', 'EndTime', 'allowedTabs', 'student_delete', 'canQueryAttendance', 'canDownload', 'Status', 'Creation_Date', 'Inactive_Date', 'Email', 'REC_ID', 'canAmendSIRN', 'canCorrectAttendance', 'DaySchedule', 'canDeleteUsers'] },
  { sheet: 'User_Permissions', fixed: true, columns: ['Username', 'canMarkManual', 'canMarkBarcode', 'canSeeSummary', 'canSetupHours', 'Can_Update_Batch_Hours', 'canDownloadReports', 'canDownloadAllReports', 'canMarkEpisodic', 'canCreateEpisodic', 'canManageEpisodicEvents', 'canEditStudentHours', 'canMultiCategory', 'REC_ID'] },
  { sheet: 'Teacher_Assignments', fixed: false, columns: ['Username', 'Class_Section', 'REC_ID'] },
  { sheet: 'Class_Sections', fixed: false, columns: ['Class_Section', 'REC_ID'] },
  { sheet: 'Student_Data', fixed: false, columns: ['Std_ID', 'Barcode_ID', 'Student_Name', 'Student_Class', 'Student_Section', 'Status', 'Date_of_Joining', 'REC_ID'] },
  { sheet: 'Student_Attendance', fixed: false, columns: ['Attendance_ID', 'Date', 'Std_ID', 'Barcode_ID', 'Student_Name', 'Student_Class', 'Student_Section', 'Class_Section', 'Status', 'Attendance_Type', 'Hours', 'Category', 'Event_Name', 'Teacher_ID', 'Timestamp', 'REC_ID', 'Term_ID', 'Term_Name'] },
  { sheet: 'Previous_Date_Permissions', fixed: true, columns: ['Username', 'Allowed_Date', 'Day', 'Password', 'Attendance_Type', 'Class_Section', 'Granted_At', 'Expires_At', 'REC_ID'] },
  { sheet: 'Daily_Hours_Setup', fixed: true, columns: ['Date', 'Class_Section', 'Category', 'Hours', 'Teacher_ID', 'Timestamp', 'Attendance_Type', 'REC_ID'] },
  { sheet: 'Hours_Setup', fixed: false, columns: ['Hours', 'Attendance_Type', 'Max_Custom', 'REC_ID', 'Is_Active'] },
  { sheet: 'Hour_Categories', fixed: false, columns: ['Category_ID', 'Category_Name', 'Is_Active', 'Attendance_Type', 'REC_ID', 'Applicable_Classes'] },
  { sheet: 'Terms', fixed: false, columns: ['Term_ID', 'Term_Name', 'Is_Active'] },
  { sheet: 'Class_Terms', fixed: false, columns: ['Class_Section', 'Term_ID', 'Is_Active', 'Start_Date', 'End_Date', 'Notes', 'REC_ID'] },
  { sheet: 'Batch_Hours_Log', fixed: false, columns: ['Date', 'Class_Section', 'Old_Hours', 'Old_Category', 'New_Hours', 'New_Category', 'UserID', 'REC_ID', 'Timestamp'] },
  { sheet: 'Episodic_Events', fixed: false, columns: ['Event_ID', 'Event_Name', 'Event_Date', 'Category', 'Hours', 'Class_Sections', 'Username', 'REC_ID', 'Timestamp'] },
  { sheet: 'RECNames', fixed: true, columns: ['S.#', 'LB Code', 'LB', 'REC Code', 'REC Name', 'LB_REC', 'Is_Active'] }
];

function structureNorm_(h) {
  // 'Student Class', 'Student_Class' and 'StudentClass' all normalise
  // to the same key, so a renamed column is spotted instead of
  // duplicated.
  return String(h == null ? '' : h).toLowerCase().replace(/[_\s]+/g, '');
}

function getOrCreateStructureLogSheet_(ss) {
  let sheet = ss.getSheetByName(STRUCTURE_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(STRUCTURE_LOG_SHEET);
    sheet.appendRow(STRUCTURE_LOG_HEADERS);
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  }
  return sheet;
}

// Append one Structure_Log row per entry. Called only when there is at
// least one entry, so the log sheet appears lazily on first use.
function logStructureEntries_(ss, appLabel, entries, via) {
  if (!entries || entries.length === 0) return;
  const sheet = getOrCreateStructureLogSheet_(ss);
  const now = new Date();
  const rows = entries.map(function (e) {
    return [now, appLabel, e.sheet, e.detail, via];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, STRUCTURE_LOG_HEADERS.length).setValues(rows);
  SpreadsheetApp.flush();
}

// Check every sheet in the schema and append its missing columns.
// Returns the per-sheet results, or null when the script lock is busy
// (the caller then skips quietly). via is 'Auto' or 'Check' and is
// stamped on every Structure_Log entry this run writes.
function ensureSheetStructure_(appLabel, via, schema) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return null;
  try {
    const results = [];
    const logEntries = [];
    schema.forEach(function (entry) {
      const res = { sheet: entry.sheet, status: 'OK', added: [], warnings: [] };
      const sheet = ss.getSheetByName(entry.sheet);
      if (!sheet) {
        // Never created here -- the sheet appears when its feature is
        // first used. Auto mode skips it silently; the manual report
        // shows it as 'Not present'.
        res.status = 'Not present';
        results.push(res);
        return;
      }
      const lastCol = sheet.getLastColumn();
      if (lastCol < 1) {
        res.status = 'Warning';
        res.warnings.push('Sheet exists but has no header row — not touched.');
        logEntries.push({ sheet: entry.sheet, detail: 'Warning: sheet exists but has no header row — not touched' });
        results.push(res);
        return;
      }
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
        return String(h == null ? '' : h).trim();
      });
      const present = {};
      headers.forEach(function (h) { if (h) present[h] = true; });
      const missing = entry.columns.filter(function (c) { return !present[c]; });
      let toAppend = [];
      missing.forEach(function (c) {
        const n = structureNorm_(c);
        const found = headers.filter(function (h) { return h && structureNorm_(h) === n; })[0];
        if (found) {
          res.warnings.push('Possible mismatch: expected "' + c + '", found "' + found + '" — not touched.');
          logEntries.push({ sheet: entry.sheet, detail: 'Warning: possible mismatch — expected "' + c + '", found "' + found + '" — not touched' });
        } else {
          toAppend.push(c);
        }
      });
      if (entry.fixed && toAppend.length > 0) {
        const inSheetOrder = headers.filter(function (h) { return entry.columns.indexOf(h) !== -1; });
        const inSchemaOrder = entry.columns.filter(function (c) { return !!present[c]; });
        if (JSON.stringify(inSheetOrder) !== JSON.stringify(inSchemaOrder)) {
          res.warnings.push('Column order differs from the expected layout — not touched.');
          logEntries.push({ sheet: entry.sheet, detail: 'Warning: column order differs from the expected layout — not touched' });
          toAppend = [];
        }
      }
      if (toAppend.length > 0) {
        sheet.getRange(1, lastCol + 1, 1, toAppend.length).setValues([toAppend]);
        SpreadsheetApp.flush();
        res.added = toAppend.slice();
        toAppend.forEach(function (c) {
          logEntries.push({ sheet: entry.sheet, detail: 'Added column: ' + c });
        });
      }
      if (res.added.length > 0) res.status = 'Added';
      else if (res.warnings.length > 0) res.status = 'Warning';
      results.push(res);
    });
    logStructureEntries_(ss, appLabel, logEntries, via);
    return results;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// The login-time auto-run: at most once per ~6 hours (script cache),
// and any failure inside is swallowed so a login can never break
// because of structure work.
function maybeAutoStructureCheck_(appLabel, schema) {
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get(STRUCTURE_CACHE_KEY_)) return;
    const results = ensureSheetStructure_(appLabel, 'Auto', schema);
    if (results) cache.put(STRUCTURE_CACHE_KEY_, '1', STRUCTURE_CACHE_SECONDS_);
  } catch (e) {
    try { Logger.log('Structure auto-check skipped: ' + e.message); } catch (e2) {}
  }
}
