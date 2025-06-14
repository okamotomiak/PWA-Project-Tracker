function ensureOwnersSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Owners');
  if (!sheet) {
    sheet = ss.insertSheet('Owners');
    const headers = ['Owner', 'Email', 'First Name', 'Last Name'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    formatOwnersSheet(sheet);
  }
}

// Completely (re)create the Owners sheet with some sample data
function initializeOwnersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Owners');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Owners');
  }

  const headers = ['Owner', 'Email', 'First Name', 'Last Name'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const data = [
    ['Justin', '', 'Justin', ''],
    ['PWA', '', 'PWA', ''],
    ['Naokimi', '', 'Naokimi', '']
  ];

  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  formatOwnersSheet(sheet);
}

function sendReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectSheet = ss.getSheetByName('Project Tracking');
  const ownersSheet = ss.getSheetByName('Owners');
  if (!projectSheet || !ownersSheet) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }

  const ownersData = ownersSheet.getRange(2, 1, ownersSheet.getLastRow() - 1, 4).getValues();
  const ownersMap = {};
  ownersData.forEach(function(row) {
    const [owner, email, firstName, lastName] = row;
    if (owner) {
      ownersMap[owner] = { email: email, firstName: firstName, lastName: lastName };
    }
  });

  const projectData = projectSheet.getRange(2, 1, projectSheet.getLastRow() - 1, 8).getValues();
  const projectsByOwner = {};
  projectData.forEach(function(row) {
    const owner = row[5];
    if (!owner) return;
    if (!projectsByOwner[owner]) {
      projectsByOwner[owner] = [];
    }
    projectsByOwner[owner].push({ project: row[0], status: row[6] });
  });

  Object.keys(projectsByOwner).forEach(function(owner) {
    const info = ownersMap[owner];
    if (!info || !info.email) {
      return;
    }
    const projects = projectsByOwner[owner];
    let body = 'Hello ' + (info.firstName || owner) + ',\n\nHere is the current status of your projects:\n';
    projects.forEach(function(p) {
      body += '- ' + p.project + ': ' + p.status + '\n';
    });
    body += '\nRegards,\nProject Tracker';
    const subject = 'Project Status Reminder';
    MailApp.sendEmail(info.email, subject, body);
  });

  SpreadsheetApp.getUi().alert('Reminders sent.');
}

/**
 * Applies consistent formatting to the Owners sheet for better readability.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Owners sheet.
 */
function formatOwnersSheet(sheet) {
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  // Style header row
  sheet.getRange(1, 1, 1, lastColumn)
       .setBackground('#e6f3ff')
       .setFontWeight('bold')
       .setFontColor('#1a73e8')
       .setFontSize(12);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Remove existing banding and apply new light grey banding
  sheet.getBandings().forEach(function(b) { b.remove(); });
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn)
         .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  }

  // Add borders
  sheet.getRange(1, 1, Math.max(lastRow, 1), lastColumn)
       .setBorder(true, true, true, true, true, true);

  // Standard row height and auto size columns
  if (lastRow > 0) {
    sheet.setRowHeights(2, lastRow - 1, 30);
  }
  sheet.autoResizeColumns(1, lastColumn);
  // Improve column widths for readability
  var widths = [140, 220, 140, 140];
  for (var i = 0; i < widths.length && i < lastColumn; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}
