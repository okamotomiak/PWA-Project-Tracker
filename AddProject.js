/**
 * Google Apps Script for adding new projects to the Project Tracking sheet
 */

function onOpen() {
  // Add a custom menu to the spreadsheet
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Project Management')
    .addItem('Add New Project', 'openAddProjectDialog')
    .addItem('Send Reminders', 'sendReminders')
    .addItem('Initialize Sheets', 'initializeAllSheets')
    .addToUi();

  // Ensure the Owners sheet exists
  ensureOwnersSheet();
}

function openAddProjectDialog() {
  // Create and show the HTML dialog
  const htmlOutput = HtmlService.createHtmlOutputFromFile('AddProjectDialog')
    .setWidth(600)
    .setHeight(700)
    .setTitle('Add New Project');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Add New Project');
}

function addNewProject(projectData) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('Project Tracking');
    
    // If sheet doesn't exist, try to get the active sheet
    if (!sheet) {
      sheet = spreadsheet.getActiveSheet();
    }
    
    // Find the next empty row
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    // Prepare the data array
    const rowData = [
      projectData.projectName,
      projectData.priority,
      projectData.dueDate ? new Date(projectData.dueDate) : '',
      projectData.description,
      projectData.deliverables,
      projectData.owner,
      projectData.status,
      projectData.notes
    ];
    
    // Add the data to the sheet
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Apply formatting to the new row
    formatNewProjectRow(sheet, newRow, projectData.priority, projectData.status);
    
    // Auto-resize columns if needed
    sheet.autoResizeColumns(1, 8);
    
    return {
      success: true,
      message: 'Project added successfully!',
      rowNumber: newRow
    };
    
  } catch (error) {
    console.error('Error adding project:', error);
    return {
      success: false,
      message: 'Error adding project: ' + error.toString()
    };
  }
}

function formatNewProjectRow(sheet, rowNumber, priority, status) {
  // Format Priority column with color coding
  const priorityCell = sheet.getRange(rowNumber, 2);
  switch (priority) {
    case 'High':
      priorityCell.setBackground('#ffebee').setFontColor('#c62828');
      break;
    case 'Medium':
      priorityCell.setBackground('#fff3e0').setFontColor('#ef6c00');
      break;
    case 'Low':
      priorityCell.setBackground('#e8f5e8').setFontColor('#2e7d32');
      break;
  }
  
  // Format Status column with color coding
  const statusCell = sheet.getRange(rowNumber, 7);
  switch (status) {
    case 'Done':
      statusCell.setBackground('#e8f5e8').setFontColor('#2e7d32');
      break;
    case 'In Progres':
      statusCell.setBackground('#fff3e0').setFontColor('#ef6c00');
      break;
    case 'Not Started':
      statusCell.setBackground('#ffebee').setFontColor('#c62828');
      break;
  }
  
  // Format Due Date column
  const dueDateCell = sheet.getRange(rowNumber, 3);
  try {
    dueDateCell.setNumberFormat('m/d/yyyy');
  } catch (e) {
    // If the Due Date column is a typed column, setting the number format can
    // throw a ScriptError. We simply skip formatting in that case.
    console.warn('Unable to set number format for Due Date column:', e);
  }
  
  // Set text wrapping for Description, Deliverables, and Notes columns
  const descriptionCell = sheet.getRange(rowNumber, 4);
  descriptionCell.setWrap(true);
  
  const deliverablesCell = sheet.getRange(rowNumber, 5);
  deliverablesCell.setWrap(true);
  
  const notesCell = sheet.getRange(rowNumber, 8);
  notesCell.setWrap(true);
  
  // Add borders to the new row
  const rowRange = sheet.getRange(rowNumber, 1, 1, 8);
  rowRange.setBorder(true, true, true, true, true, true);
  
  // Set row height for better readability
  sheet.setRowHeight(rowNumber, 60);
}

function getDropdownOptions() {
  // Return the dropdown options for the form
  return {
    priorities: ['High', 'Medium', 'Low'],
    owners: ['Justin', 'PWA', 'Naokimi', 'Other'],
    statuses: ['Not Started', 'In Progres', 'Done']
  };
}

function validateProjectData(projectData) {
  const errors = [];
  
  if (!projectData.projectName || projectData.projectName.trim() === '') {
    errors.push('Project Name is required');
  }
  
  if (!projectData.priority) {
    errors.push('Priority is required');
  }
  
  if (!projectData.description || projectData.description.trim() === '') {
    errors.push('Description is required');
  }
  
  if (!projectData.owner) {
    errors.push('Owner is required');
  }
  
  if (!projectData.status) {
    errors.push('Status is required');
  }
  
  return errors;
}
