# PWA-Project-Tracker

This repository contains Google Apps Script code for managing project and task tracking within Google Sheets.

## Scripts

- **recreateProjectTrackingSheet** – creates or refreshes a "Project Tracking" sheet in the active spreadsheet with sample data.
- **createRecurringTasksSheet** – creates or refreshes a "Recurring Tasks" sheet for scheduling repeating tasks in the active spreadsheet.
- **initializeAllSheets** – sets up the Project Tracking, Recurring Tasks, and Owners sheets in the active spreadsheet.
- **initializeOwnersSheet** – recreates the Owners sheet with `Owner`, `Email`, `First Name`, and `Last Name` columns.
- **sendReminders** – emails each owner a list of their active projects and recurring tasks. Owners without active projects are skipped and projects marked **Done** are omitted from the emails.

All creation scripts now apply data validation drop-downs. Priority and Status columns use predefined lists while Owner selections reference the **Owners** sheet.

Each script formats the sheet with headers, example rows, and color-coded statuses without generating a separate file.
