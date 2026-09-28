/**
 * EduObserve Scheduler - Google Apps Script WebApp
 * Helps educational leaders find and observe teachers
 */

// Configuration - You need to set your spreadsheet ID here
//const SPREADSHEET_ID = '1yLj4St9lrqzrBkWWGMbRYlX_E2JCdNFQSiCS74c0MQdFTgjrvVqxtH6K'; // Replace with your actual spreadsheet ID

// Helper function to get the spreadsheet
function getSpreadsheet() {
  try {
    // Try to get the spreadsheet by ID first
    //if (SPREADSHEET_ID) {
    //  return SpreadsheetApp.openById(SPREADSHEET_ID);
    //}
    // Fallback to active spreadsheet if ID is not set
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (error) {
    Logger.log('Error getting spreadsheet: ' + error.toString());
    throw new Error('Cannot access spreadsheet. Please check the SPREADSHEET_ID configuration.');
  }
}

// Entry point for the web app
function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();
  
  // Check if user is authorized
  if (!isAuthorizedUser(userEmail)) {
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <title>Access Denied</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; margin-top: 100px; }
            .error { color: #d32f2f; font-size: 18px; }
          </style>
        </head>
        <body>
          <h1 class="error">Access Denied</h1>
          <p>You do not have permission to access this application.</p>
          <p>Please contact your administrator if you believe this is an error.</p>
        </body>
      </html>
    `);
  }
  
  // Return the main application HTML
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('EduObserve Scheduler')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include HTML files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Check if user is authorized
function isAuthorizedUser(email) {
  try {
    const ss = getSpreadsheet();
    const usersSheet = ss.getSheetByName('USERS');
    const users = usersSheet.getRange('A:A').getValues().flat();
    return users.includes(email);
  } catch (error) {
    Logger.log('Error checking authorization: ' + error.toString());
    return false;
  }
}

// Get dropdown options for Day, Period, and Teachers
function getDropdownOptions() {
  try {
    const ss = getSpreadsheet();
    const timetableSheet = ss.getSheetByName('TIMETABLE_MASTER');
    const data = timetableSheet.getDataRange().getValues();
    
    if (data.length <= 1) return { days: [], periods: [], teachers: [] };
    
    // Remove header row
    const rows = data.slice(1);
    
    // Natural sort helper for days and periods
    const naturalSort = (a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });

    // Extract unique days and periods (columns: day=4, periodName=2)
    const days = [...new Set(rows.map(row => row[4]).filter(day => day))].sort(naturalSort);
    const periods = [...new Set(rows.map(row => row[2]).filter(period => period))].sort(naturalSort);

    // Extract unique teachers (columns: email=0, teacherName=1)
    const teacherMap = new Map();
    rows.forEach(row => {
      const email = row[0] ? row[0].toString().trim() : '';
      const name = row[1] ? row[1].toString().trim() : email;
      if (email && !teacherMap.has(email)) {
        teacherMap.set(email, name);
      }
    });

    const teachers = Array.from(teacherMap.entries()).map(([email, name]) => ({
      email: email,
      name: name
    })).sort((a, b) => a.name.localeCompare(b.name));
    
    return { days, periods, teachers };
  } catch (error) {
    Logger.log('Error getting dropdown options: ' + error.toString());
    return { days: [], periods: [], teachers: [] };
  }
}

// Get teachers for specific day and period
function getTeachersForDayPeriod(day, period) {
  try {
    const ss = getSpreadsheet();
    const timetableSheet = ss.getSheetByName('TIMETABLE_MASTER');
    const observedSheet = ss.getSheetByName('OBSERVED');
    
    // Get timetable data
    const timetableData = timetableSheet.getDataRange().getValues();
    if (timetableData.length <= 1) return { notObserved: [], observed: [] };
    
    // Get observed teachers
    const observedData = observedSheet.getDataRange().getValues();
    const observedEmails = new Set();
    if (observedData.length > 1) {
      observedData.slice(1).forEach(row => observedEmails.add(row[0])); // Teacher Email in column A
    }
    
    // Filter teachers for the specific day and period
    const filteredTeachers = timetableData.slice(1).filter(row => 
      row[4] === day && row[2] === period // day=4, periodName=2
    ).map(row => ({
      email: row[0],
      teacherName: row[1],
      periodName: row[2],
      className: row[3],
      day: row[4],
      time: row[5],
      room: row[6]
    }));
    
    // Split into observed and not observed
    const notObserved = filteredTeachers.filter(teacher => !observedEmails.has(teacher.email));
    const observed = filteredTeachers.filter(teacher => observedEmails.has(teacher.email));
    
    return { notObserved, observed };
  } catch (error) {
    Logger.log('Error getting teachers: ' + error.toString());
    return { notObserved: [], observed: [] };
  }
}

// Return timetable rows relevant to the current common-free search.
function getCommonFreeData(selectedTeacherEmails, selectedDays, selectedPeriods) {
  try {
    if (!selectedTeacherEmails || !selectedTeacherEmails.length ||
        !selectedDays || !selectedDays.length ||
        !selectedPeriods || !selectedPeriods.length) {
      return [];
    }

    const ss = getSpreadsheet();
    const timetableSheet = ss.getSheetByName('TIMETABLE_MASTER');
    const timetableData = timetableSheet.getDataRange().getValues();

    if (timetableData.length <= 1) return { teachers: [], scheduleRows: [] };

    const selectedEmails = new Set(selectedTeacherEmails.map(email => email.toString().trim()));
    const selectedDaySet = new Set(selectedDays.map(day => day.toString().trim()));
    const selectedPeriodSet = new Set(selectedPeriods.map(period => period.toString().trim()));

    const normalizedRows = timetableData.slice(1).map(row => ({
        email: row[0] ? row[0].toString().trim() : '',
        name: row[1] ? row[1].toString().trim() : '',
        day: row[4] ? row[4].toString().trim() : '',
        period: row[2] ? row[2].toString().trim() : ''
      }));

    const teachers = [];
    const teacherNames = {};
    normalizedRows.forEach(row => {
      if (selectedEmails.has(row.email) && row.name && !teacherNames[row.email]) {
        teacherNames[row.email] = row.name;
        teachers.push({ email: row.email, name: row.name });
      }
    });

    const scheduleRows = normalizedRows.filter(row => selectedEmails.has(row.email) &&
        selectedDaySet.has(row.day) && selectedPeriodSet.has(row.period));

    return { teachers: teachers, scheduleRows: scheduleRows };
  } catch (error) {
    Logger.log('Error getting common-free data: ' + error.toString());
    return { teachers: [], scheduleRows: [] };
  }
}

function exportCommonFrees(destination, existingSheetReference, values) {
  if (!values || !values.length || !values[0].length) {
    throw new Error('No Common Frees data is available to export.');
  }

  const spreadsheet = destination === 'existing'
    ? SpreadsheetApp.openById(extractSpreadsheetId(existingSheetReference))
    : SpreadsheetApp.create('Common Frees Export');
  const sheetName = getUniqueSheetName(spreadsheet, 'Common Frees');
  const sheet = destination === 'existing'
    ? spreadsheet.insertSheet(sheetName)
    : spreadsheet.getSheets()[0].setName(sheetName);
  const rowCount = values.length;
  const columnCount = values[0].length;

  if (sheet.getMaxRows() < rowCount) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rowCount - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < columnCount) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), columnCount - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, rowCount, columnCount).setValues(values);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, columnCount).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.autoResizeColumns(1, columnCount);

  return {
    sheetName: sheetName,
    url: spreadsheet.getUrl() + '#gid=' + sheet.getSheetId()
  };
}

function extractSpreadsheetId(reference) {
  const match = (reference || '').toString().match(/[a-zA-Z0-9-_]{25,}/);
  if (!match) throw new Error('Enter a valid Google Sheet URL or ID.');
  return match[0];
}

function getUniqueSheetName(spreadsheet, baseName) {
  const existingNames = new Set(spreadsheet.getSheets().map(sheet => sheet.getName()));
  if (!existingNames.has(baseName)) return baseName;

  let index = 1;
  while (existingNames.has(baseName + ' (' + index + ')')) {
    index++;
  }
  return baseName + ' (' + index + ')';
}

// Add observation record
function addObservation(teacherEmail, teacherName, date, className) {
  try {
    const observerEmail = Session.getActiveUser().getEmail();
    const ss = getSpreadsheet();
    const observedSheet = ss.getSheetByName('OBSERVED');
    
    // Add new row
    observedSheet.appendRow([teacherEmail, teacherName, date, className, observerEmail]);
    
    return { success: true, message: 'Observation recorded successfully' };
  } catch (error) {
    Logger.log('Error adding observation: ' + error.toString());
    return { success: false, message: 'Error recording observation: ' + error.toString() };
  }
}

// Get all observed records
function getObservedList() {
  try {
    const ss = getSpreadsheet();
    const observedSheet = ss.getSheetByName('OBSERVED');
    const data = observedSheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      teacherEmail: row[0],
      teacherName: row[1],
      date: row[2] ? new Date(row[2]).toISOString().split('T')[0] : '', // Fix date serialization
      className: row[3],
      observerEmail: row[4]
    }));
  } catch (error) {
    Logger.log('Error getting observed list: ' + error.toString());
    return [];
  }
}

// Get all classes for a specific teacher
function getTeacherSchedule(teacherEmail) {
  try {
    const ss = getSpreadsheet();
    const timetableSheet = ss.getSheetByName('TIMETABLE_MASTER');
    const data = timetableSheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    // Filter and return all classes for this teacher
    return data.slice(1)
      .filter(row => row[0] === teacherEmail) // Teacher EMAIL in column 0
      .map(row => ({
        email: row[0],
        teacherName: row[1],
        periodName: row[2],
        className: row[3],
        day: row[4],
        time: row[5],
        room: row[6]
      }));
  } catch (error) {
    Logger.log('Error getting teacher schedule: ' + error.toString());
    return [];
  }
}

// Setup function to create required sheets with headers (run this once to initialize)
function setupSheets() {
  try {
    const ss = getSpreadsheet();
    
    // Setup TIMETABLE_MASTER sheet
    let timetableSheet = ss.getSheetByName('TIMETABLE_MASTER');
    if (!timetableSheet) {
      timetableSheet = ss.insertSheet('TIMETABLE_MASTER');
    }
    
    // Check if headers exist, if not, add them
    const timetableHeaders = timetableSheet.getRange(1, 1, 1, 7).getValues()[0];
    if (!timetableHeaders[0] || timetableHeaders[0] !== 'Teacher EMAIL') {
      timetableSheet.getRange(1, 1, 1, 7).setValues([
        ['Teacher EMAIL', 'teacherName', 'periodName', 'class', 'day', 'time', 'room']
      ]);
      timetableSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }
    
    // Setup OBSERVED sheet
    let observedSheet = ss.getSheetByName('OBSERVED');
    if (!observedSheet) {
      observedSheet = ss.insertSheet('OBSERVED');
    }
    
    // Check if headers exist, if not, add them
    const observedHeaders = observedSheet.getRange(1, 1, 1, 5).getValues()[0];
    if (!observedHeaders[0] || observedHeaders[0] !== 'Teacher Email') {
      observedSheet.getRange(1, 1, 1, 5).setValues([
        ['Teacher Email', 'teacherName', 'date', 'className', 'observer Email']
      ]);
      observedSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    
    // Setup USERS sheet
    let usersSheet = ss.getSheetByName('USERS');
    if (!usersSheet) {
      usersSheet = ss.insertSheet('USERS');
    }
    
    // Check if headers exist, if not, add them
    const userHeaders = usersSheet.getRange(1, 1, 1, 1).getValues()[0];
    if (!userHeaders[0] || userHeaders[0] !== 'Authorized User Emails') {
      usersSheet.getRange(1, 1, 1, 1).setValues([
        ['Authorized User Emails']
      ]);
      usersSheet.getRange(1, 1, 1, 1).setFontWeight('bold');
      
      // Add the current user as the first authorized user
      const currentUser = Session.getActiveUser().getEmail();
      if (currentUser) {
        usersSheet.getRange(2, 1).setValue(currentUser);
      }
    }
    
    Logger.log('Sheets setup completed successfully');
    return { success: true, message: 'Sheets setup completed successfully' };
    
  } catch (error) {
    Logger.log('Error setting up sheets: ' + error.toString());
    return { success: false, message: 'Error setting up sheets: ' + error.toString() };
  }
}
