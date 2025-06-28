function addRecord(name, dates, tab_name) {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(tab_name);

  // Check if the sheet exists
  if (!dataSheet) {
    logMessageError(getCallStackTrace() + `: Sheet with name "${tab_name}" does not exist.`);
    return;
  }

  // Fetch all data from the sheet
  var data = dataSheet.getDataRange().getValues();
  if (data.length <= 1) { // Check if the sheet is empty or only contains headers
    logMessageError(getCallStackTrace() + ": Sheet is empty or contains only headers.");
    return;
  }

  // Loop through each date in the provided dates array
  // Make sure the dates or record[1] is not missing, 
  if (dates != "Yes" && dates != "No") {
    dates.forEach((inputDate) => {
      let dateFound = false; // Flag to check if date was found in the sheet

      // Loop through the rows to find the matching date
      for (var row = 1; row < data.length; row++) {
        // Format the date in the sheet to match the input format
        var rowDate = Utilities.formatDate(new Date(data[row][0]), Session.getScriptTimeZone(), 'MM/dd/yyyy');

        // Check if the date matches
        if (rowDate === inputDate) {
          dateFound = true;
          logMessage(getCallStackTrace() + `: Found matching date "${inputDate}" at row ${row + 1}.`);

          // Find the first empty cell in the row
          var lastColumn = data[row].findIndex(cell => cell === "") + 1; // First empty cell
          if (lastColumn === 0) lastColumn = data[row].length + 1; // If no empty cell, append at the end

          // Add the name to the first empty cell in the matching row
          dataSheet.getRange(row + 1, lastColumn).setValue(name);
          logMessage(getCallStackTrace() + `: Added name "${name}" to row ${row + 1}, column ${lastColumn}.`);
          break;
        }
      }

      if (!dateFound) {
        console.warn(getCallStackTrace() + `: No matching row found for date "${inputDate}".`);
      }
    });
  } else {
    console.warn(getCallStackTrace() + `: No records was added, because no input date was selected, instead it shows as "${dates}".`);
  }


  // Flush pending changes to the sheet
  SpreadsheetApp.flush();
}
