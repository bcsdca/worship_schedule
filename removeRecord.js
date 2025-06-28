function removeRecord(name, date, tab_name) {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(tab_name);
  
  // Check if the sheet exists
  if (!dataSheet) {
    logMessageError(getCallStackTrace() + `Sheet with name "${tab_name}" does not exist.`);
    return;
  }

  // Fetch all data from the sheet
  var data = dataSheet.getDataRange().getValues();
  if (data.length <= 1) { // Check if the sheet is empty or only has a header row
    logMessageError(getCallStackTrace() + `: Sheet is empty or contains only headers.`);
    return;
  }

  // Loop through every row (starting from the first data row, index 1)
  // to look for the target name to delete
  for (var row = 1; row < data.length; row++) {
    // Extract the date in a readable format
    var rowDate = Utilities.formatDate(new Date(data[row][0]), Session.getScriptTimeZone(), 'MM/dd/yyyy');
   
    // Track how many cells have been deleted in the row to adjust column indexing
    var shiftOffset = 0;

    // Loop through the columns (excluding the first column which is the date)
    for (var col = 1; col < data[row].length; col++) {
      if (data[row][col] === name) {
        logMessage(getCallStackTrace() + `: Deleting "${name}" from row ${row + 1}, column ${col + 1 - shiftOffset}.`);

        // Delete the cell and shift the row contents left
        // Adds 1 to the row and column index 
        // because spreadsheet row indexing starts at 1, while JavaScript arrays start at 0.
        dataSheet.getRange(row + 1, col + 1 - shiftOffset).deleteCells(SpreadsheetApp.Dimension.COLUMNS);
        shiftOffset++;
      }
    }
  }
}

