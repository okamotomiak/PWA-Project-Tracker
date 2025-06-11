function ensureOwnersSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Owners');
  if (!sheet) {
    sheet = ss.insertSheet('Owners');
    const headers = ['Owner', 'Email', 'First Name', 'Last Name'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
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
