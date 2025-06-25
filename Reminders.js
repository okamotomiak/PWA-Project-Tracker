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

  // Corrected data with placeholder emails to ensure sendReminders can work
  const data = [
    ['Justin', 'justin@example.com', 'Justin', ''],
    ['PWA', 'pwa@example.com', 'PWA', ''],
    ['Naokimi', 'naokimi@example.com', 'Naokimi', '']
  ];

  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  formatOwnersSheet(sheet);
}

function sendReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectSheet = ss.getSheetByName('Project Tracking');
  const ownersSheet = ss.getSheetByName('Owners');
  const tasksSheet = ss.getSheetByName('Recurring Tasks');
  if (!projectSheet || !ownersSheet) {
    SpreadsheetApp.getUi().alert('Required sheets not found. Please run the initialization script.');
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
    const owner = row[5]; // Owner is in column F (index 5)
    if (!owner) return;
    if (!projectsByOwner[owner]) {
      projectsByOwner[owner] = [];
    }
    projectsByOwner[owner].push({
      project: row[0], // Project Name from column A
      dueDate: row[2], // Due Date from column C
      status: row[6]   // Status from column G
    });
  });

  const tasksByOwner = {};
  if (tasksSheet) {
    const tasksData = tasksSheet.getRange(2, 1, tasksSheet.getLastRow() - 1, 8).getValues();
    tasksData.forEach(function(row) {
      const owner = row[4]; // Owner column E
      if (!owner) return;
      if (!tasksByOwner[owner]) {
        tasksByOwner[owner] = [];
      }
      tasksByOwner[owner].push({
        task: row[0], // Task Name
        nextDue: row[3], // Next Due Date
        status: row[5] // Status
      });
    });
  }

  let remindersSentCount = 0;
  const sheetLink = ss.getUrl() + '#gid=' + projectSheet.getSheetId();
  const allOwners = new Set([...Object.keys(projectsByOwner), ...Object.keys(tasksByOwner)]);
  allOwners.forEach(function(owner) {
    const info = ownersMap[owner];
    if (!info || !info.email) {
      console.log('Skipping reminder for ' + owner + ' due to missing email.');
      return; // Skip this owner if no email is listed
    }

    const projects = projectsByOwner[owner] || [];
    // Exclude projects marked as Done
    const activeProjects = projects.filter(function(p) {
      return String(p.status).toLowerCase() !== 'done';
    });
    // Skip sending if owner has no active projects
    if (activeProjects.length === 0) {
      console.log('Skipping reminder for ' + owner + ' due to no active projects.');
      return;
    }
    const tasks = tasksByOwner[owner] || [];

    let body = 'Hello ' + (info.firstName || owner) + ',\n\n';
    if (activeProjects.length > 0) {
      body += 'Here is the current status of your projects:\n';
      activeProjects.forEach(function(p) {
        const due = p.dueDate ? Utilities.formatDate(new Date(p.dueDate), ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy') : 'No due date';
        body += '- ' + p.project + ' (Due: ' + due + '): ' + p.status + '\n';
      });
    }

    if (tasks.length > 0) {
      if (activeProjects.length > 0) body += '\n';
      body += 'Recurring tasks:\n';
      tasks.forEach(function(t) {
        const due = t.nextDue ? Utilities.formatDate(new Date(t.nextDue), ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy') : 'No due date';
        body += '- ' + t.task + ' (Next Due: ' + due + '): ' + t.status + '\n';
      });
    }

    body += '\nCheck project sheet for more details:\n' + sheetLink + '\n\nRegards,\nProject Tracker';

    const subject = 'Project Status Reminder';
    MailApp.sendEmail(info.email, subject, body);
    remindersSentCount++;
  });

  SpreadsheetApp.getUi().alert(remindersSentCount + ' reminder(s) sent.');
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

  // Freeze header row for easy scrolling
  sheet.setFrozenRows(1);

  // Remove any existing banding before applying a new one to prevent conflicts
  sheet.getBandings().forEach(function(b) { b.remove(); });
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn)
         .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  }

  // Add borders to all cells with data
  sheet.getRange(1, 1, Math.max(lastRow, 1), lastColumn)
       .setBorder(true, true, true, true, true, true);

  // Standard row height for data rows
  if (lastRow > 1) {
    sheet.setRowHeights(2, lastRow - 1, 30);
  }

  // Auto-resize columns first to set a baseline
  sheet.autoResizeColumns(1, lastColumn);
  
  // Then, set specific widths for key columns to improve readability
  const widths = [140, 220, 140, 140]; // Widths for Owner, Email, First Name, Last Name
  for (let i = 0; i < widths.length && i < lastColumn; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}
