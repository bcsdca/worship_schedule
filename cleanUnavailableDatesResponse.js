function cleanUnavailableDatesResponse() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Unavailable Dates Response")
  var lastRow = sheet.getLastRow()
  var lastColumn = sheet.getLastColumn();
  logMessage(getCallStackTrace() + `: lastRow is ${lastRow}, lastColumn is ${lastColumn}`);

  sheet.getRange(2, 1, lastRow, lastColumn).clearContent();

  SpreadsheetApp.getActive().toast("Done, Clean up all the old previous \"Unavailable Dates Response\" 👍 !!!");

}

