/**
 * @OnlyCurrentDoc
 */

// =================================================================
// MAIN SETUP FUNCTIONS
// =================================================================

/**
 * Main function to initialize all necessary sheets.
 */
function initializeAllSheets() {
  initializeOwnersSheet(); // This was missing, now re-added
  recreateProjectTrackingSheet();
  createRecurringTasksSheet();
  SpreadsheetApp.getUi().alert('All sheets have been successfully initialized!');
  console.log('All required sheets initialized.');
}

/**
 * Creates and populates the 'Owners' sheet with sample data.
 */
function initializeOwnersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Owners';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  const headers = ['Owner', 'Email', 'First Name', 'Last Name'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
             .setFontWeight('bold')
             .setBackground('#e6f3ff');

  const data = [
    ['Justin', 'justin@example.com', 'Justin', ''],
    ['PWA', 'pwa@example.com', 'PWA', ''],
    ['Naokimi', 'naokimi@example.com', 'Naokimi', '']
  ];

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  console.log('Owners sheet initialized.');
}

/**
 * Recreates the "Project Tracking" sheet with a predefined structure and data.
 */
function recreateProjectTrackingSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Project Tracking';
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }
  sheet = spreadsheet.insertSheet(sheetName);

  const headers = [
    'Project Name', 'Priority', 'Due Date', 'Description',
    'Deliverables', 'Owner', 'Status', 'Notes'
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
             .setFontWeight('bold')
             .setBackground('#e6f3ff');

  const data = [
    ['PSWM - Resource Promotion', 'Medium', new Date('2025-06-04'), 'Highlight PSWM resources on website and engage non-attendees', '', 'Justin', 'Done', ''],
    ['Discourse Series - Marketing', 'High', new Date('2025-06-04'), 'Slides, newsletter, announcements, registration form with Zoom reminders', 'Slides\nnewsletter\nannouncements\nregistration form with Zoom reminders', 'PWA', 'Done', ''],
    ['Discourse Series - Online Communuty Tutorial (During First Session)', 'Medium', new Date('2025-06-09'), 'We will do the community involvement', '', 'PWA', 'In Progress', ''],
    ['New Job Description', 'High', new Date('2025-06-04'), 'Draft job description aligned with taking on SG responsibilities', '', 'Justin', 'Done', ''],
    ['Contract with PWA', 'High', new Date('2025-06-04'), 'Define needs from PWA and review proposal package', '', 'Justin', 'Done', ''],
    ['Weds NE Pastor\'s Agenda', 'High', new Date('2025-06-04'), 'Finalize agenda for NE Pastor\'s Weds call', '', 'Naokimi', 'Done', ''],
    ['Community Goal Setting Form', 'High', new Date('2025-06-04'), 'Create form for communities to set goals by size, week-to-week results, baseline data', '', 'PWA', 'Done', ''],
    ['NE Financial Review', 'High', new Date('2025-06-05'), 'Set meeting with Shizuko to review NE finances', '', 'Naokimi', 'Done', ''],
    ['Discourse Series - Night Vigil', 'High', new Date('2025-06-09'), 'Run nightly vigil 10pm-10:40pm with rotating pastors, includes tech setup and cue sheet', '', 'PWA', 'Done', ''],
    ['Transition FamChu to FamGrHost', 'Medium', new Date('2025-06-10'), 'Transition all family churches to family group hosts. Family group hosts means they would no longer be their own community, they would be a part of the Northeast Community, tithe directly to the Northeast and report to a leader in the northeast.', '- All Family Church size pastors are aware and on board\n- Central Figure for FGH is secured\n- FGH are trained in new setup\n- Content creation structure in place', 'PWA', 'Not Started', ''],
    ['New Membership - Digital Forms', 'High', new Date('2025-06-11'), 'Integrated digital membership forms', 'See below', 'PWA', 'Done', ''],
    ['Meeting with Kaeleigh', 'Medium', new Date('2025-06-11'), 'Discuss how providential orgs can support core metrics', '', 'Naokimi', 'Done', ''],
    ['Scaling the 3 metrics', 'High', new Date('2025-06-20'), 'Build system and strategy for reporting on PSWM Intro, New Members, Discourse', '- New Member Form Training\n- Event Management Forms (apply to PSWM) \n- EM Training for Pastors and Ministry\n- Connecting current discourse registration to Attendance Sheets\n- Connect our dashboards to HQ data', 'Justin', 'In Progress', 'Setup a weekly meeting with Rev. Rendel and Taka?'],
    ['Jake and Kikuchi Plan', 'High', new Date('2025-06-14'), 'Transition Jake out by July 1st?\nDicision will be made by end of the week June 14', 'Clarified roles and destination for Kikuchi and Jake from July', 'Naokimi', 'In Progress', ''],
    ['Hosting NE True Family Tour', 'High', new Date('2025-06-21'), 'True Grand Children are coming 16-22nd to the NE\nThere will be a youth event on the 21st 3pm at the New Yorker\nEvent on the 22nd at Belvedere at 4pm', '', '', '', ''],
    ['Heavenly Fortune Review', 'Medium', '', 'Review course content and impact', '', 'Naokimi', 'Done', ''],
    ['NE Leadership Summit Prep', 'High', new Date('2025-06-20'), 'Plan agenda and content for Leadership Summit\nFriday day Belvedere TC\nNight - Tudor\nSat - NYC\nAfternoon youth event', '- Event Sheet filled out\n- Budget\n- Catering is planned\n- Location is secured', 'PWA', 'Not Started', '']
  ];

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }

  formatProjectTrackingSheet(sheet);
  applyDataValidations(sheet, 'Project Tracking');
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
  console.log('Project Tracking sheet recreated in: ' + spreadsheet.getUrl());
}

/**
 * Creates or refreshes the "Recurring Tasks" sheet.
 */
function createRecurringTasksSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Recurring Tasks';
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  const headers = [
    'Task Name', 'Frequency', 'Day/Pattern', 'Next Due Date',
    'Owner', 'Status', 'Last Completed Date', 'Notes'
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
             .setFontWeight('bold')
             .setBackground('#e6f3ff');

  const data = [
    ['Send Weekly Newsletter', 'Weekly', 'Monday', new Date('2025-06-10'), 'Justin', 'Not Started', new Date('2025-06-03'), ''],
    ['Monthly Planning Meeting', 'Monthly', '3rd Thursday', new Date('2025-06-20'), 'PWA', 'Not Started', new Date('2025-05-23'), '']
  ];

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }

  applyDataValidations(sheet, 'Recurring Tasks');
  formatRecurringTasksSheet(sheet); // Format function now handles all aesthetics
  console.log('Recurring Tasks sheet created in current spreadsheet');
}

// =================================================================
// FORMATTING AND VALIDATION
// =================================================================

/**
 * Applies conditional formatting to the "Project Tracking" sheet.
 */
function formatProjectTrackingSheet(sheet) {
  const numRows = sheet.getLastRow() - 1;
  if (numRows <= 0) return;

  const rules = [
    { range: sheet.getRange(2, 2, numRows, 1), colors: { 'High': '#ffebee', 'Medium': '#fff3e0', 'Low': '#e8f5e8' } },
    { range: sheet.getRange(2, 7, numRows, 1), colors: { 'Done': '#e8f5e8', 'In Progress': '#fff3e0', 'Not Started': '#ffebee' } }
  ];

  rules.forEach(rule => {
    const conditionalFormatRules = [];
    for (const [value, color] of Object.entries(rule.colors)) {
      conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(value)
        .setBackground(color)
        .setRanges([rule.range])
        .build());
    }
    sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat(conditionalFormatRules));
  });

  // Apply other formatting - REMOVED setNumberFormat
  sheet.getRange(2, 4, numRows, 2).setWrap(true); // Description and Deliverables
  sheet.getRange(2, 8, numRows, 1).setWrap(true); // Notes
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setBorder(true, true, true, true, true, true);
  sheet.setRowHeights(2, numRows, 60);
}

/**
 * Applies conditional formatting and other styles to the "Recurring Tasks" sheet.
 */
function formatRecurringTasksSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const numRows = lastRow - 1;

  if (numRows <= 0) return; // Exit if no data rows

  // --- 1. Conditional Formatting Rules ---
  const statusRange = sheet.getRange(2, 6, numRows, 1);
  const dueDateRange = sheet.getRange(2, 4, numRows, 1);
  const rules = [
    // Status Rules
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Done').setBackground('#e8f5e8').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('In Progress').setBackground('#fff3e0').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Not Started').setBackground('#ffebee').setRanges([statusRange]).build(),
    // Overdue Rule for 'Next Due Date'
    SpreadsheetApp.newConditionalFormatRule().whenDateBefore(SpreadsheetApp.RelativeDate.TODAY).setBackground('#ffcccc').setRanges([dueDateRange]).build()
  ];
  sheet.setConditionalFormatRules(rules);

  // --- 2. Data Body Formatting ---
  // Apply alternating row colors (banding)
  sheet.getBandings().forEach(b => b.remove());
  sheet.getRange(2, 1, numRows, lastColumn).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  // Set standard row height and wrap text in the 'Notes' column
  sheet.setRowHeights(2, numRows, 40);
  sheet.getRange(2, 8, numRows, 1).setWrap(true);

  // --- 3. Sheet-Wide Formatting ---
  // Add borders to the entire data range
  sheet.getRange(1, 1, lastRow, lastColumn).setBorder(true, true, true, true, true, true);
  
  // Freeze the first column and header row
  sheet.setFrozenColumns(1);
  sheet.setFrozenRows(1);
  
  // Create a filter for the entire data range
  sheet.getRange(1, 1, lastRow, lastColumn).createFilter();

  // --- 4. Column Widths (Resolved Merge Conflict) ---
  // Auto-resize columns first, then set specific widths for readability
  sheet.autoResizeColumns(1, lastColumn);
  const widths = [200, 100, 130, 120, 120, 120, 140, 250];
  for (let i = 0; i < widths.length && i < lastColumn; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}

/**
 * Applies data validation rules to a sheet.
 */
function applyDataValidations(sheet, sheetType) {
  const maxRows = sheet.getMaxRows();
  const ownersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Owners');
  let ownerValues = ['Justin', 'PWA', 'Naokimi', 'Other']; // Fallback

  if (ownersSheet) {
    const owners = ownersSheet.getRange('A2:A').getValues().flat().filter(String);
    if (owners.length > 0) {
      ownerValues = owners;
    }
  }

  const ownerRule = SpreadsheetApp.newDataValidation().requireValueInList(ownerValues, true).build();
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['Not Started', 'In Progress', 'Done'], true).build();

  if (sheetType === 'Project Tracking') {
    const priorityRule = SpreadsheetApp.newDataValidation().requireValueInList(['High', 'Medium', 'Low'], true).build();
    sheet.getRange(2, 2, maxRows - 1, 1).setDataValidation(priorityRule);
    sheet.getRange(2, 6, maxRows - 1, 1).setDataValidation(ownerRule);
    sheet.getRange(2, 7, maxRows - 1, 1).setDataValidation(statusRule);
  } else if (sheetType === 'Recurring Tasks') {
    sheet.getRange(2, 5, maxRows - 1, 1).setDataValidation(ownerRule);
    sheet.getRange(2, 6, maxRows - 1, 1).setDataValidation(statusRule);
  }
}

/**
 * Simple onEdit trigger to update the "Next Due Date" column when the
 * "Last Completed Date" is edited in the Recurring Tasks sheet.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Recurring Tasks' || range.getColumn() !== 7 || range.getRow() < 2) {
    return;
  }

  const lastCompleted = range.getValue();
  const frequency = sheet.getRange(range.getRow(), 2).getValue();
  const pattern = sheet.getRange(range.getRow(), 3).getValue();
  const nextDue = calculateNextDueDate(new Date(lastCompleted), frequency, pattern);

  if (nextDue) {
    sheet.getRange(range.getRow(), 4).setValue(nextDue);
  }
}

/**
 * Calculate the next due date for a recurring task.
 * @param {Date} lastDate The last completed date.
 * @param {string} frequency The recurrence frequency (Weekly, Monthly).
 * @param {string} pattern The day or pattern (e.g. "Monday" or "3rd Thursday").
 * @return {Date|null} The calculated next due date or null if it cannot be determined.
 */
function calculateNextDueDate(lastDate, frequency, pattern) {
  if (!lastDate || isNaN(lastDate) || !frequency) return null;
  frequency = String(frequency).toLowerCase();
  if (frequency === 'weekly') {
    return getNextWeeklyDate(lastDate, pattern);
  }
  if (frequency === 'monthly') {
    return getNextMonthlyDate(lastDate, pattern);
  }
  return null;
}

function getNextWeeklyDate(lastDate, dayStr) {
  if (!dayStr) return null;
  const dayNames = { sunday: 0, monday: 1, tuesday: 2, wednesday: 3, thursday: 4, friday: 5, saturday: 6 };
  const target = dayNames[String(dayStr).toLowerCase()];
  if (target === undefined) return null;
  const start = new Date(lastDate);
  for (let i = 1; i <= 7; i++) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    if (d.getDay() === target) return d;
  }
  return null;
}

function getNextMonthlyDate(lastDate, patternStr) {
  if (!patternStr) return null;
  const pattern = parseMonthlyPattern(patternStr);
  if (!pattern) return null;
  let year = lastDate.getFullYear();
  let month = lastDate.getMonth();
  for (let i = 0; i < 12; i++) {
    const d = computeDateForMonth(year, month, pattern);
    if (d && d > lastDate) return d;
    month++;
    if (month > 11) {
      month = 0;
      year++;
    }
  }
  return null;
}

function parseMonthlyPattern(pattern) {
  pattern = String(pattern).toLowerCase().trim();
  const dayNames = { sunday: 0, monday: 1, tuesday: 2, wednesday: 3, thursday: 4, friday: 5, saturday: 6 };
  const ordMap = { '1st': 1, first: 1, '2nd': 2, second: 2, '3rd': 3, third: 3, '4th': 4, fourth: 4, last: 'last' };

  const m = pattern.match(/^(\d{1,2})(?:st|nd|rd|th)?$/);
  if (m) {
    const day = parseInt(m[1], 10);
    if (day >= 1 && day <= 31) return { type: 'dayOfMonth', day: day };
  }

  const parts = pattern.split(/\s+/);
  if (parts.length === 2 && ordMap[parts[0]] && dayNames[parts[1]]) {
    return { type: 'weekdayOfMonth', ordinal: ordMap[parts[0]], weekday: dayNames[parts[1]] };
  }
  if (parts.length === 2 && parts[0] === 'last' && dayNames[parts[1]]) {
    return { type: 'weekdayOfMonth', ordinal: 'last', weekday: dayNames[parts[1]] };
  }
  return null;
}

function computeDateForMonth(year, month, p) {
  if (p.type === 'dayOfMonth') {
    const d = new Date(year, month, p.day);
    return d.getMonth() === month ? d : null;
  }
  if (p.type === 'weekdayOfMonth') {
    if (p.ordinal === 'last') {
      let d = new Date(year, month + 1, 0);
      while (d.getDay() !== p.weekday) d.setDate(d.getDate() - 1);
      return d;
    } else {
      let d = new Date(year, month, 1);
      while (d.getDay() !== p.weekday) d.setDate(d.getDate() + 1);
      d.setDate(d.getDate() + (p.ordinal - 1) * 7);
      return d.getMonth() === month ? d : null;
    }
  }
  return null;
}
