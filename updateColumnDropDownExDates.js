// Update the dropDownValue for the taskName column pointed by headerRow[columnIndex]
// from startRow to lastRow, one row at a time. It checks the "Exception Dates" sheet 
// to find out all the names unavailable for each week and removes them from the drop-down list.

/*function updateColumnDropDownExDates(startRow, lastRow, headerRow, columnIndex, dropDownValue) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var exceptionDatesSheet = ss.getSheetByName("Exception Dates");
    var scheduleSheet = ss.getSheets()[0];
    
    // Get the exception data range using a helper function
    const rowColStartEnd = getRowColStartEnd(exceptionDatesSheet);
    const [exRowStart, exRowEnd, exColumnStart, exColumnEnd] = rowColStartEnd[1];
    const headerRow = rowColStartEnd[0];
    logMessage(getCallStackTrace() + `: headerRow: ${JSON.stringify(headerRow)}`);
    // Retrieve all exception rows in one call
    var exceptionRows = exceptionDatesSheet.getRange(exRowStart, exColumnStart, exRowEnd - exRowStart + 1, exColumnEnd - exColumnStart + 1).getValues();
    logMessage(getCallStackTrace() + `: exceptionRows: ${JSON.stringify(exceptionRows)}`);
    return;

    // Generate array of row indices from startRow to lastRow
    var rowIndices = Array.from({ length: lastRow - startRow + 1 }, (_, i) => i + startRow);

    // Iterate over row indices to update dropdowns
    // Iterate through the exceptionRows
    // Create an array of row indices to iterate over (from startRow to lastRow)
    // rowIndices = [2,3,4,5,6,7,8,9,10,11,12,13,14]
    // Use forEach to iterate over each row
    // using rowIndices method below combined with getting all the exceptionRows at one time help the execution time about 6x
    // the original code getting one exception row at a time, api call in a loop waste a lot of time

    rowIndices.forEach(rowIndex => {
      // Retrieve the exception row and filter out empty cells
      var exceptionRow = exceptionRows[rowIndex - startRow].filter(String);
      var dateValue = Utilities.formatDate(exceptionRow.shift(), Session.getScriptTimeZone(), 'MM/dd/yyyy');

      logMessage(getCallStackTrace() + `: Working on the dropdown for task = \"${headerRow[columnIndex]}\", date = \"${dateValue}\", exceptionRow = ${JSON.stringify(exceptionRow)}, old dropdown value = ${JSON.stringify(dropDownValue)}`);
    

      if (!dropDownModifyExcludeList.includes(headerRow[columnIndex])) {

        // Filter dropdown to exclude exceptions date for each person
        var newDropDownValue = dropDownValue.filter(item => !exceptionRow.includes(item));

        // Apply the new dropdown values to the sheet
        //setDropDown(scheduleSheet, rowIndex, columnIndex, newDropDownValue);

      } else {
        logMessage(getCallStackTrace() + `: NOT modifying the dropdown list affected by the \"exception tab\" for task = \"${headerRow[columnIndex]}\", date = \"${dateValue}\"`);
        var newDropDownValue = dropDownValue;
      }
      
      // Apply the new dropdown values to the sheet
      setDropDown(scheduleSheet, rowIndex, columnIndex, newDropDownValue);

      // Log changes only if dropdown was modified
      if (!arraysEqual(dropDownValue, newDropDownValue)) {
        logDropDownModification(rowIndex, columnIndex, headerRow[columnIndex], dropDownValue, newDropDownValue, dateValue);
      }
    });
  } catch (error) {
    //logMessage(getCallStackTrace() + ": Error in updateColumnDropDown1: " + error.message);
    logMessageError(getCallStackTrace() + error); // Log to the console for more detailed stack trace
  }
}
// Helper function to log dropdown changes
function logDropDownModification(rowIndex, columnIndex, taskName, originalValues, modifiedValues, dateValue) {
  logMessage(getCallStackTrace() + `: Row: ${rowIndex}, Column: ${columnIndex + 1}, Task: ${taskName}, Date: ${dateValue}`);
  logMessage(getCallStackTrace() + `: Original Dropdown: ${JSON.stringify(originalValues)}`);
  logMessage(getCallStackTrace() + `: Modified Dropdown: ${JSON.stringify(modifiedValues)}`);
}

// Helper function to apply dropdown values to a specific cell
function setDropDown(sheet, rowIndex, columnIndex, dropDownValues) {
  var dropdownCell = sheet.getRange(rowIndex, columnIndex + 1); // columnIndex + 1 to match 1-based indexing
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(dropDownValues).build();
  dropdownCell.setDataValidation(rule);
}

// Helper function to compare two arrays
function arraysEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) return false; // Different lengths
  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) return false; // Mismatch found
  }
  return true; // All elements match
}
*/
