function generateDayPeriods() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Days");
  
  // Get all data in the sheet
  const data = sheet.getDataRange().getValues();
  
  const days = [];
  const periods = [];
  
  // Start from index 1 to skip the header row (DAY, PERIODS)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) days.push(data[i][0]);
    if (data[i][1]) periods.push(data[i][1]);
  }
  
  const results = [["Day-Period"]]; // Header for Column C
  
  // Generate the combinations
  for (let i = 0; i < days.length; i++) {
    for (let j = 0; j < periods.length; j++) {
      results.push([`${days[i]}-${periods[j]}`]);
    }
  }
  
  // Clear Column C before writing (optional, but good practice)
  sheet.getRange("C:C").clearContent();
  
  // Write the results to Column C
  sheet.getRange(1, 3, results.length, 1).setValues(results);
}


function listFreeTeachers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get the Day-Periods from the "Days" sheet (Column C)
  const daysSheet = ss.getSheetByName("Days");
  // Assuming row 1 is the header, get data from row 2 down
  const daysData = daysSheet.getRange("C2:C" + daysSheet.getLastRow()).getValues();
  const dayPeriods = daysData.map(row => row[0]).filter(val => val !== "");
  
  // 2. Get the Timetable data
  const ttSheet = ss.getSheetByName("TIMETABLE_MASTER");
  const ttData = ttSheet.getDataRange().getValues();
  
  // Indices based on TIMETABLE_MASTER structure:
  // [0] EMAIL, [1] teacherName, [2] periodName, [3] class, [4] day, [5] time, [6] room, [7] Day-Period
  const TEACHER_NAME_COL = 1; 
  const DAY_PERIOD_COL = 7;   
  
  const allTeachers = new Set();
  const busyMap = {}; // Will store: teacherName -> Set of busy Day-Periods
  
  // Start from 1 to skip the header row in TIMETABLE_MASTER
  for (let i = 1; i < ttData.length; i++) {
    // Fallback to EMAIL if teacherName is blank
    const teacherName = ttData[i][TEACHER_NAME_COL] ? ttData[i][TEACHER_NAME_COL].toString().trim() : ttData[i][0].toString().trim();
    const dayPeriod = ttData[i][DAY_PERIOD_COL] ? ttData[i][DAY_PERIOD_COL].toString().trim() : "";
    
    if (!teacherName) continue; // Skip empty rows
    
    allTeachers.add(teacherName);
    
    if (!busyMap[teacherName]) {
      busyMap[teacherName] = new Set();
    }
    if (dayPeriod) {
      busyMap[teacherName].add(dayPeriod);
    }
  }
  
  // Sort the list of all unique teachers alphabetically
  const uniqueTeachers = Array.from(allTeachers).sort();
  
  // 3. Prepare the data for the "Free Teachers" sheet
  const outputData = [];
  outputData.push(dayPeriods); // Row 1: The Day-Period headers
  
  const freeTeachersPerPeriod = [];
  let maxFreeCount = 0;
  
  // For each Day-Period, find which teachers are NOT in the busyMap for that period
  for (let i = 0; i < dayPeriods.length; i++) {
    const dp = dayPeriods[i];
    const freeTeachers = uniqueTeachers.filter(t => !busyMap[t] || !busyMap[t].has(dp));
    
    freeTeachersPerPeriod.push(freeTeachers);
    
    if (freeTeachers.length > maxFreeCount) {
      maxFreeCount = freeTeachers.length;
    }
  }
  
  // Populate the rows underneath the headers
  for (let rowIdx = 0; rowIdx < maxFreeCount; rowIdx++) {
    const row = [];
    for (let colIdx = 0; colIdx < dayPeriods.length; colIdx++) {
      // Add the teacher's name, or a blank string if we've run out of free teachers for this column
      row.push(freeTeachersPerPeriod[colIdx][rowIdx] || "");
    }
    outputData.push(row);
  }
  
  // 4. Create or clear the "Free Teachers" sheet
  let freeSheet = ss.getSheetByName("Free Teachers");
  if (!freeSheet) {
    freeSheet = ss.insertSheet("Free Teachers");
  } else {
    freeSheet.clear();
  }
  
  // 5. Write the data to the sheet
  if (outputData.length > 0 && outputData[0].length > 0) {
    freeSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    // Format the headers to make them stand out
    freeSheet.setFrozenRows(1);
    freeSheet.getRange(1, 1, 1, outputData[0].length).setFontWeight("bold").setBackground("#f3f3f3");
    freeSheet.autoResizeColumns(1, outputData[0].length);
  }
}

function listFreeTeachersByEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get the Day-Periods from the "Days" sheet (Column C)
  const daysSheet = ss.getSheetByName("Days");
  // Assuming row 1 is the header, get data from row 2 down
  const daysData = daysSheet.getRange("C2:C" + daysSheet.getLastRow()).getValues();
  const dayPeriods = daysData.map(row => row[0]).filter(val => val !== "");
  
  // 2. Get the Timetable data
  const ttSheet = ss.getSheetByName("TIMETABLE_MASTER");
  const ttData = ttSheet.getDataRange().getValues();
  
  // Indices based on TIMETABLE_MASTER structure:
  // [0] EMAIL, [1] teacherName, [2] periodName, [3] class, [4] day, [5] time, [6] room, [7] Day-Period
  const EMAIL_COL = 0; 
  const DAY_PERIOD_COL = 7;   
  
  const allEmails = new Set();
  const busyMap = {}; // Will store: email -> Set of busy Day-Periods
  
  // Start from 1 to skip the header row in TIMETABLE_MASTER
  for (let i = 1; i < ttData.length; i++) {
    const email = ttData[i][EMAIL_COL] ? ttData[i][EMAIL_COL].toString().trim() : "";
    const dayPeriod = ttData[i][DAY_PERIOD_COL] ? ttData[i][DAY_PERIOD_COL].toString().trim() : "";
    
    if (!email) continue; // Skip rows that do not have an email listed
    
    allEmails.add(email);
    
    if (!busyMap[email]) {
      busyMap[email] = new Set();
    }
    if (dayPeriod) {
      busyMap[email].add(dayPeriod);
    }
  }
  
  // Sort the list of all unique emails alphabetically
  const uniqueEmails = Array.from(allEmails).sort();
  
  // 3. Prepare the data for the "Free Teachers" sheet
  const outputData = [];
  outputData.push(dayPeriods); // Row 1: The Day-Period headers
  
  const freeEmailsPerPeriod = [];
  let maxFreeCount = 0;
  
  // For each Day-Period, find which emails are NOT in the busyMap for that period
  for (let i = 0; i < dayPeriods.length; i++) {
    const dp = dayPeriods[i];
    const freeEmails = uniqueEmails.filter(e => !busyMap[e] || !busyMap[e].has(dp));
    
    freeEmailsPerPeriod.push(freeEmails);
    
    if (freeEmails.length > maxFreeCount) {
      maxFreeCount = freeEmails.length;
    }
  }
  
  // Populate the rows underneath the headers
  for (let rowIdx = 0; rowIdx < maxFreeCount; rowIdx++) {
    const row = [];
    for (let colIdx = 0; colIdx < dayPeriods.length; colIdx++) {
      // Add the email, or a blank string if we've run out of free emails for this column
      row.push(freeEmailsPerPeriod[colIdx][rowIdx] || "");
    }
    outputData.push(row);
  }
  
  // 4. Create or clear the "Free Teachers" sheet
  let freeSheet = ss.getSheetByName("Free Teachers");
  if (!freeSheet) {
    freeSheet = ss.insertSheet("Free Teachers");
  } else {
    freeSheet.clear();
  }
  
  // 5. Write the data to the sheet
  if (outputData.length > 0 && outputData[0].length > 0) {
    freeSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    // Format the headers to make them stand out
    freeSheet.setFrozenRows(1);
    freeSheet.getRange(1, 1, 1, outputData[0].length).setFontWeight("bold").setBackground("#f3f3f3");
    freeSheet.autoResizeColumns(1, outputData[0].length);
  }
}