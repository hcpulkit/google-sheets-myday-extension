const THEME = {
  COLORS: {
    BG_GRAY: '#f6f8fa',
    BORDER_GRAY: '#d0d7de',
    HEADER_BG: '#24292f',
    TASK_BG: '#fff3cd',
    ASSIGNMENT_BG: '#d1ecf1',
    NOTES_BG: '#eaf7ea',
    STATUS_COMPLETED: '#d4edda',
    STATUS_IN_PROGRESS: '#fff3cd',
    STATUS_NOT_STARTED: '#f8d7da'
  },
  STATUSES: ['⭕ Not Started', '🟡 In Progress', '✅ Completed']
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('myday')
    .addItem('Add New Day\'s Schedule', 'addNewDaySchedule')
    .addToUi();
}

function addNewDaySchedule() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  if (sheet.getLastRow() === 0) {
    setupSheet(sheet);
  }

  const startRow = sheet.getLastRow() + 1;

  // Add date header with styling
  const dateRange = sheet.getRange(startRow, 1, 1, 8); // Changed to span all 8 columns
  dateRange.merge()
    .setValue(`📅 ${today}`)
    .setBackground(THEME.COLORS.BORDER_GRAY)
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Setup for separate "Tasks" and "Assignments" columns
  const headers = [
    ['S. No.', 'Tasks', 'Deadline', 'Status', 'S. No.', 'Assignments', 'Deadline', 'Status']
  ];
  sheet.getRange(startRow + 1, 1, 1, 8).setValues(headers)
    .setBackground(THEME.COLORS.BG_GRAY)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Add rows for tasks and assignments
  const rows = [
    [1, '', '', THEME.STATUSES[0], 1, '', '', THEME.STATUSES[0]],
    [2, '', '', THEME.STATUSES[0], 2, '', '', THEME.STATUSES[0]],
    [3, '', '', THEME.STATUSES[0], 3, '', '', THEME.STATUSES[0]],
    [4, '', '', THEME.STATUSES[0], 4, '', '', THEME.STATUSES[0]],
    [5, '', '', THEME.STATUSES[0], 5, '', '', THEME.STATUSES[0]]
  ];
  sheet.getRange(startRow + 2, 1, rows.length, 8).setValues(rows);

  // Styling borders and cell sizes
  const taskRange = sheet.getRange(startRow + 1, 1, rows.length + 1, 8);
  taskRange.setBorder(true, true, true, true, true, true);
  
  // Center-align all text and dates in the entire table
  sheet.getRange(startRow + 1, 1, rows.length + 1, 8).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Add data validations for status columns (both Tasks and Assignments)
  const taskStatusRange = sheet.getRange(startRow + 2, 4, rows.length, 1);
  const assignmentStatusRange = sheet.getRange(startRow + 2, 8, rows.length, 1);
  const statusValidation = SpreadsheetApp.newDataValidation().requireValueInList(THEME.STATUSES).build();
  
  taskStatusRange.setDataValidation(statusValidation);
  assignmentStatusRange.setDataValidation(statusValidation);

  // Add data validations for deadline columns (both Tasks and Assignments)
  const taskDeadlineRange = sheet.getRange(startRow + 2, 3, rows.length, 1);
  const assignmentDeadlineRange = sheet.getRange(startRow + 2, 7, rows.length, 1);
  const dateValidation = SpreadsheetApp.newDataValidation().requireDate().build();
  
  taskDeadlineRange.setDataValidation(dateValidation);
  assignmentDeadlineRange.setDataValidation(dateValidation);

  // Format deadline columns as date
  const dateFormat = "dd/MM/yyyy";
  taskDeadlineRange.setNumberFormat(dateFormat);
  assignmentDeadlineRange.setNumberFormat(dateFormat);

  // Conditional formatting for both status columns
  addConditionalFormatting(sheet, startRow + 2, rows.length);

  // Add notes section
  const notesRow = startRow + rows.length + 2;
  const notesRange = sheet.getRange(notesRow, 1, 1, 8);
  notesRange.merge()
    .setValue('📝 Notes:')
    .setBackground(THEME.COLORS.NOTES_BG)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

function setupSheet(sheet) {
  // Title with distinctive header style
  const titleRange = sheet.getRange(1, 1, 1, 8);
  titleRange.merge()
    .setValue('🗓️ myday 🗓️')
    .setBackground(THEME.COLORS.HEADER_BG)
    .setFontColor('white')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Set column widths for readability
  sheet.setColumnWidths(1, 8, 150);
  sheet.setColumnWidth(1, 50);  // Serial No.
  sheet.setColumnWidth(2, 200); // Tasks
  sheet.setColumnWidth(3, 100); // Deadline
  sheet.setColumnWidth(4, 120); // Status
  sheet.setColumnWidth(5, 50);  // Serial No.
  sheet.setColumnWidth(6, 200); // Assignments
  sheet.setColumnWidth(7, 100); // Deadline
  sheet.setColumnWidth(8, 120); // Status
}

function addConditionalFormatting(sheet, startRow, numRows) {
  const rules = [];
  
  // Add rules for both status columns (column 4 and 8)
  [4, 8].forEach(col => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('✅ Completed')
        .setBackground(THEME.COLORS.STATUS_COMPLETED)
        .setFontColor('#155724')
        .setRanges([sheet.getRange(startRow, col, numRows)])
        .build(),
        
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('⭕ Not Started')
        .setBackground(THEME.COLORS.STATUS_NOT_STARTED)
        .setFontColor('#721c24')
        .setRanges([sheet.getRange(startRow, col, numRows)])
        .build()
    );
  });
  
  sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat(rules));
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const nextRow = range.getRow() + 1;
  const col = range.getColumn();
  
  // Only proceed for task and assignment description columns (columns 2 and 6)
  if ((col === 2 || col === 6) && range.getValue() !== '') {
    const nextCell = sheet.getRange(nextRow, col);
    if (nextCell.getValue() === '') {
      nextCell.activate();
    }
  }
}
