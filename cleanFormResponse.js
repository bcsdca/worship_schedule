function cleanFormResponse() {
  
  clearLogSheet();
  var sheet = getSheetByWildcardName();
  var lastRow = sheet.getLastRow()
  var lastColumn = sheet.getLastColumn();
  logMessage(getCallStackTrace() + `: lastRow is ${lastRow}, lastColumn is ${lastColumn}`);

  //sheet.getRange(2, 1, lastRow, lastColumn).clearContent();

  if (lastRow > 1) { // Ensure there are rows to delete
    sheet.deleteRows(2, lastRow - 1); // Delete rows from row 2 to last row
  }
  flushLogsToSheet
  SpreadsheetApp.getActive().toast("Done, Clean up all the old data in \"Form Responses For Not Available Date\" sheet 👍 !!!");

}

//find the 1st "Form Responses XX" sheet
function getSheetByWildcardName() {
  var sheets = SpreadsheetApp.getActive().getSheets();
  var sheetNamePrefix = "Form Responses "; // The prefix you're looking for
  var matchingSheet = null;

  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();

    // Check if the sheet name starts with "Form Responses "
    if (sheetName.startsWith(sheetNamePrefix)) {
      matchingSheet = sheets[i];
      break; // Stop the loop if you find a match
    }
  }

  if (matchingSheet) {
    logMessage(getCallStackTrace() + `: Found sheet: \"${matchingSheet.getName()}\", with sheet index of ${i}`);
    return matchingSheet;
  } else {
    logMessage(getCallStackTrace() + `: No sheet matching the wildcard found.`);
    return null;
  }
}

