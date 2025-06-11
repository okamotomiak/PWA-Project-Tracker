/**
 * AppScript to recreate the Project Tracking Google Sheet
 * This script creates a new spreadsheet with the exact structure and data
 */

function recreateProjectTrackingSheet() {
  // Create a new spreadsheet
  const spreadsheet = SpreadsheetApp.create('Project Tracking - Copy');
  const sheet = spreadsheet.getActiveSheet();
  
  // Rename the sheet to "Project Tracking"
  sheet.setName('Project Tracking');
  
  // Set up the headers
  const headers = [
    'Project Name',
    'Priority', 
    'Due Date',
    'Description',
    'Deliverables',
    'Owner',
    'Status',
    'Notes'
  ];
  
  // Add headers to row 1
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e6f3ff');
  
  // Data rows
  const data = [
    [
      'PSWM - Resource Promotion',
      'Medium',
      new Date('6/4/2025'),
      'Highlight PSWM resources on website and engage non-attendees',
      '',
      'Justin',
      'Done',
      ''
    ],
    [
      'Discourse Series - Marketing',
      'High',
      new Date('6/4/2025'),
      'Slides, newsletter, announcements, registration form with Zoom reminders',
      'Slides\nnewsletter\nannouncements\nregistration form with Zoom reminders',
      'PWA',
      'Done',
      ''
    ],
    [
      'Discourse Series - Online Communuty Tutorial (During First Session)',
      'Medium',
      new Date('6/9/2025'),
      'We will do the community involvement',
      '',
      'PWA',
      'In Progres',
      ''
    ],
    [
      'New Job Description',
      'High',
      new Date('6/4/2025'),
      'Draft job description aligned with taking on SG responsibilities',
      '',
      'Justin',
      'Done',
      ''
    ],
    [
      'Contract with PWA',
      'High',
      new Date('6/4/2025'),
      'Define needs from PWA and review proposal package',
      '',
      'Justin',
      'Done',
      ''
    ],
    [
      'Weds NE Pastor\'s Agenda',
      'High',
      new Date('6/4/2025'),
      'Finalize agenda for NE Pastor\'s Weds call',
      '',
      'Naokimi',
      'Done',
      ''
    ],
    [
      'Community Goal Setting Form',
      'High',
      new Date('6/4/2025'),
      'Create form for communities to set goals by size, week-to-week results, baseline data',
      '',
      'PWA',
      'Done',
      ''
    ],
    [
      'NE Financial Review',
      'High',
      new Date('6/5/2025'),
      'Set meeting with Shizuko to review NE finances',
      '',
      'Naokimi',
      'Done',
      ''
    ],
    [
      'Discourse Series - Night Vigil',
      'High',
      new Date('6/9/2025'),
      'Run nightly vigil 10pm-10:40pm with rotating pastors, includes tech setup and cue sheet',
      '',
      'PWA',
      'Done',
      ''
    ],
    [
      'Transition FamChu to FamGrHost',
      'Medium',
      new Date('6/10/2025'),
      'Transition all family churches to family group hosts. Family group hosts means they would no longer be their own community, they would be a part of the Northeast Community, tithe directly to the Northeast and report to a leader in the northeast.',
      '- All Family Church size pastors are aware and on board\n- Central Figure for FGH is secured\n- FGH are trained in new setup\n- Content creation structure in place',
      'PWA',
      'Not Started',
      ''
    ],
    [
      'New Membership - Digital Forms',
      'High',
      new Date('6/11/2025'),
      'Integrated digital membership forms',
      'See below',
      'PWA',
      'Done',
      ''
    ],
    [
      'Meeting with Kaeleigh',
      'Medium',
      new Date('6/11/2025'),
      'Discuss how providential orgs can support core metrics',
      '',
      'Naokimi',
      'Done',
      ''
    ],
    [
      'Scaling the 3 metrics',
      'High',
      new Date('6/20/2025'),
      'Build system and strategy for reporting on PSWM Intro, New Members, Discourse',
      '- New Member Form Training\n- Event Management Forms (apply to PSWM) \n- EM Training for Pastors and Ministry\n- Connecting current discourse registration to Attendance Sheets\n- Connect our dashboards to HQ data',
      'Justin',
      'In Progres',
      'Setup a weekly meeting with Rev. Rendel and Taka?'
    ],
    [
      'Jake and Kikuchi Plan',
      'High',
      new Date('6/14/2025'),
      'Transition Jake out by July 1st?\nDicision will be made by end of the week June 14',
      'Clarified roles and destination for Kikuchi and Jake from July',
      'Naokimi',
      'In Progres',
      ''
    ],
    [
      'Hosting NE True Family Tour',
      'High',
      new Date('6/21/2025'),
      'True Grand Children are coming 16-22nd to the NE\nThere will be a youth event on the 21st 3pm at the New Yorker\nEvent on the 22nd at Belvedere at 4pm',
      '',
      '',
      '',
      ''
    ],
    [
      'Heavenly Fortune Review',
      'Medium',
      '',
      'Review course content and impact',
      '',
      'Naokimi',
      'Done',
      ''
    ],
    [
      'NE Leadership Summit Prep',
      'High',
      new Date('6/20/2025'),
      'Plan agenda and content for Leadership Summit\nFriday day Belvedere TC\nNight - Tudor\nSat - NYC\nAfternoon youth event',
      '- Event Sheet filled out\n- Budget\n- Catering is planned\n- Location is secured',
      'PWA',
      'Not Started',
      ''
    ]
  ];
  
  // Add data starting from row 2
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  // Format the sheet
  formatProjectTrackingSheet(sheet, data.length + 1);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  
  // Log the URL of the new spreadsheet
  console.log('New Project Tracking sheet created: ' + spreadsheet.getUrl());
  
  return spreadsheet;
}

function formatProjectTrackingSheet(sheet, totalRows) {
  // Format Priority column with color coding
  const priorityRange = sheet.getRange(2, 2, totalRows - 1, 1);
  const priorityValues = priorityRange.getValues();
  
  for (let i = 0; i < priorityValues.length; i++) {
    const priority = priorityValues[i][0];
    const cellRange = sheet.getRange(i + 2, 2);
    
    switch (priority) {
      case 'High':
        cellRange.setBackground('#ffebee').setFontColor('#c62828');
        break;
      case 'Medium':
        cellRange.setBackground('#fff3e0').setFontColor('#ef6c00');
        break;
      case 'Low':
        cellRange.setBackground('#e8f5e8').setFontColor('#2e7d32');
        break;
    }
  }
  
  // Format Status column with color coding
  const statusRange = sheet.getRange(2, 7, totalRows - 1, 1);
  const statusValues = statusRange.getValues();
  
  for (let i = 0; i < statusValues.length; i++) {
    const status = statusValues[i][0];
    const cellRange = sheet.getRange(i + 2, 7);
    
    switch (status) {
      case 'Done':
        cellRange.setBackground('#e8f5e8').setFontColor('#2e7d32');
        break;
      case 'In Progres':
        cellRange.setBackground('#fff3e0').setFontColor('#ef6c00');
        break;
      case 'Not Started':
        cellRange.setBackground('#ffebee').setFontColor('#c62828');
        break;
    }
  }
  
  // Format Due Date column
  const dueDateRange = sheet.getRange(2, 3, totalRows - 1, 1);
  dueDateRange.setNumberFormat('m/d/yyyy');
  
  // Set text wrapping for Description and Deliverables columns
  const descriptionRange = sheet.getRange(2, 4, totalRows - 1, 1);
  descriptionRange.setWrap(true);
  
  const deliverablesRange = sheet.getRange(2, 5, totalRows - 1, 1);
  deliverablesRange.setWrap(true);
  
  const notesRange = sheet.getRange(2, 8, totalRows - 1, 1);
  notesRange.setWrap(true);
  
  // Add borders to all data
  const dataRange = sheet.getRange(1, 1, totalRows, 8);
  dataRange.setBorder(true, true, true, true, true, true);
  
  // Freeze the header row
  sheet.setFrozenRows(1);
  
  // Set row heights for better readability
  sheet.setRowHeights(2, totalRows - 1, 60);
}

// Alternative function to create sheet in existing spreadsheet
function addProjectTrackingSheetToExisting() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.insertSheet('Project Tracking');
  
  // Use the same setup logic as above
  const headers = [
    'Project Name',
    'Priority', 
    'Due Date',
    'Description',
    'Deliverables',
    'Owner',
    'Status',
    'Notes'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ... (rest of the data and formatting code would be the same)
  
  console.log('Project Tracking sheet added to current spreadsheet');
}

// Function to test the script
function testRecreateSheet() {
  recreateProjectTrackingSheet();
}
