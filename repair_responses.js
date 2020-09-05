function onOpen() {
  // Add Admin menu to the spreadsheet
  SpreadsheetApp.getUi()
    .createMenu('Admin')
      .addItem('Mark as in progress', 'menuIP')
      .addItem('Mark as waiting on parts', 'menuParts')
      .addItem('Upgrade severity', 'menuUpgrade')
      .addItem('Downgrade severity', 'menuDowngrade')
      .addItem('Mark as completed', 'menuCompleted')
      .addToUi();

  // Sort by Severity column
  SpreadsheetApp.getActiveSheet().sort(5, false);
}

function onEdit(e) {
  // Actions to take place upon editing of the sheet
  if (e.value == '0 - Completed') {
    // Copy the item to the "Completed" sheet and delete the old entry
    var lastRow = e.source.getSheetByName('Completed').getLastRow();
    var newRange = e.source.getActiveSheet().getRange(e.range.getRow(), 1, 1, 12);
    e.source.getSheetByName('Completed').insertRowsAfter(lastRow, 1);
    newRange.copyValuesToRange(e.source.getSheetByName('Completed'), 1, 12, lastRow + 1, lastRow + 1);
    e.source.getActiveSheet().deleteRow(newRange.getRow());
    }

  // Sort by Severity column
  e.source.getActiveSheet().sort(5, false);
  flush();
}

function menuIP() {
  //Mark currently selected item as in progress
  var cell = SpreadsheetApp.getActiveSheet().getActiveRange().getCell(1, 5);
  cell.setValue('7 - Repair in progress');
}

function menuParts() {
  //Mark currently selected item to waiting on parts
  var cell = SpreadsheetApp.getActiveSheet().getActiveRange().getCell(1, 5);
  cell.setValue('5 - Waiting on parts');
}

function menuUpgrade() {
  // Upgrade the severity of the currently selected item
  var cell = SpreadsheetApp.getActiveSheet().getActiveRange().getCell(1, 5);
  switch (cell.getValue()) {
    case '1 - Eventually':
      cell.setValue('2 - Not that soon');
      break;
    case '2 - Not that soon':
      cell.setValue('3 - Soon');
      break;
    case '3 - Soon':
      cell.setValue('4 - Very soon');
      break;
    case '4 - Very soon':
      cell.setValue('5 - ASAP');
      break;
    case '5 - ASAP':
      cell.setValue('6 - Waiting on parts');
      break;
    case '6 - Waiting on parts':
      cell.setValue('7 - Repair in progress');
      break;
  }
}

function menuDowngrade() {
  // Downgrade the severity of the currently selected item
  var cell = SpreadsheetApp.getActiveSheet().getActiveRange().getCell(1, 5);
  switch (cell.getValue()) {
    case '7 - Repair in progress':
      cell.setValue('6 - Waiting on parts');
      break;
    case '6 - Waiting on parts':
      cell.setValue('5 - ASAP');
      break;
    case '5 - ASAP':
      cell.setValue('4 - Very soon');
      break;
    case '4 - Very soon':
      cell.setValue('3 - Soon');
      break;
    case '3 - Soon':
      cell.setValue('2 - Not that soon');
      break;
    case '2 - Not that soon':
      cell.setValue('1 - Eventually');
      break;
    case '1 - Eventually':
      menuCompleted();
      break;
  }
}

function menuCompleted() {
  // Mark the currently selected item as completed
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getActiveRange();
  var cell = range.getCell(1, 5);
  cell.setValue('0 - Completed');

  // Copy the item to the "Completed" sheet and delete the old entry
  var lastRow = ss.getSheetByName('Completed').getLastRow();
  ss.getSheetByName('Completed').insertRowsAfter(lastRow, 1);
  range.copyValuesToRange(ss.getSheetByName('Completed'), 1, 12, lastRow + 1, lastRow + 1);
  ss.getActiveSheet().deleteRow(range.getRow());
}
