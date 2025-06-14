/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of execution
 * to the current spreadsheet only. This is a good practice for security
 * and performance.
 */

/**
 * Main function to initialize all necessary sheets.
 * This function acts as the primary entry point for setting up the spreadsheet.
 */
function initializeAllSheets() {
  initializeOwnersSheet();
  recreateProjectTrackingSheet();
  createRecurringTasksSheet();
  SpreadsheetApp.getUi().alert('All sheets have been successfully initialized!');
  console.log('All required sheets initialized.');
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

  // Set up headers
  const headers = [
    'Project Name', 'Priority', 'Due Date', 'Description',
    'Deliverables', 'Owner', 'Status', 'Notes'
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
             .setFontWeight('bold')
             .setBackground('#e6f3ff');

  // Define data for the sheet
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

  // Set up headers
  const headers = [
    'Task Name', 'Frequency', 'Day/Pattern', 'Next Due Date',
    'Owner', 'Status', 'Last Completed Date', 'Notes'
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
             .setFontWeight('bold')
             .setBackground('#e6f3ff');

  // Define data for the sheet
  const data = [
    ['Send Weekly Newsletter', 'Weekly', 'Monday', new Date('2025-06-10'), 'Justin', 'Not Started', new Date('2025-06-03'), ''],
    ['Monthly Planning Meeting', 'Monthly', '3rd Thursday', new Date('2025-06-20'), 'PWA', 'Not Started', new Date('2025-05-23'), '']
  ];

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }

  formatRecurringTasksSheet(sheet);
  applyDataValidations(sheet, 'Recurring Tasks');
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
  console.log('Recurring Tasks sheet created in current spreadsheet');
}

/**
 * Applies conditional formatting to the "Project Tracking" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to format.
 */
function formatProjectTrackingSheet(sheet) {
  const numRows = sheet.getLastRow() - 1;
  if (numRows <= 0) return; // No data to format

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

  // Apply other formatting
  sheet.getRange(2, 3, numRows, 1).setNumberFormat('m/d/yyyy');
  sheet.getRange(2, 4, numRows, 2).setWrap(true); // Description and Deliverables
  sheet.getRange(2, 8, numRows, 1).setWrap(true); // Notes
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setBorder(true, true, true, true, true, true);
  sheet.setRowHeights(2, numRows, 60);
}

/**
 * Applies conditional formatting to the "Recurring Tasks" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to format.
 */
function formatRecurringTasksSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const numRows = lastRow - 1;
  if (numRows <= 0) return;

  const statusRange = sheet.getRange(2, 6, numRows, 1);
  const dueDateRange = sheet.getRange(2, 4, numRows, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Done').setBackground('#e8f5e8').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('In Progress').setBackground('#fff3e0').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Not Started').setBackground('#ffebee').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenDateBefore(SpreadsheetApp.RelativeDate.YESTERDAY).setBackground('#ffcccc').setRanges([dueDateRange]).build()
  ];
  sheet.setConditionalFormatRules(rules);
  
  // Apply other formatting
  sheet.getRange(1, 1, 1, lastColumn)
       .setFontColor('#1a73e8')
       .setFontSize(12)
       .setFontWeight('bold');

  sheet.setFrozenColumns(1);

  sheet.getRange(2, 4, numRows, 1).setNumberFormat('m/d/yyyy'); // Next Due Date
  sheet.getRange(2, 7, numRows, 1).setNumberFormat('m/d/yyyy'); // Last Completed
  sheet.getRange(2, 8, numRows, 1).setWrap(true); // Notes

  sheet.getRange(1, 1, lastRow, lastColumn).setBorder(true, true, true, true, true, true);
  if (numRows > 0) {
    sheet.setRowHeights(2, numRows, 40);
  }

  sheet.getBandings().forEach(function(b) { b.remove(); });
  if (numRows > 0) {
    sheet.getRange(2, 1, numRows, lastColumn).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  }

  sheet.getRange(1, 1, lastRow, lastColumn).createFilter();
}


/**
 * Applies data validation rules to a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to apply validations to.
 * @param {string} sheetType The type of sheet ('Project Tracking' or 'Recurring Tasks').
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
