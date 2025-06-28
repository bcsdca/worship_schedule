function cleanUnavailableDates() {
  try {

    clearLogSheet();

    // Get active spreadsheet and the source sheet (assumed to be the first sheet)
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var srcSheet = spreadsheet.getSheets()[0];

    // Find the start and end positions for rows and columns (assuming findRowColStartEnd returns correct values)
    const rowColStartEnd = getRowColStartEnd(srcSheet);

    // Extract the start and end row/column numbers for the data range
    const srcDataRowStart = rowColStartEnd[1][0];
    const srcDataRowEnd = rowColStartEnd[1][1];
    const srcDataColumnStart = rowColStartEnd[1][2];
    const srcDataColumnEnd = rowColStartEnd[1][3];

    // Calculate the number of rows containing data in the source sheet
    const totalSrcRows = srcDataRowEnd - srcDataRowStart + 1;

    logMessage(getCallStackTrace() + `: srcDataRowStart = ${srcDataRowStart}, srcDataRowEnd = ${srcDataRowEnd}, srcDataColumnStart = ${srcDataColumnStart}, srcDataColumnEnd = ${srcDataColumnEnd}, total 'date' rows = ${totalSrcRows}`);

    // Fetch the data from the "date" column in the source sheet
    var srcValues = srcSheet.getRange(srcDataRowStart, srcDataColumnStart, totalSrcRows, 1).getValues();

    // Select the "Unavailable Dates" sheet
    var unavailableDatesSheet = spreadsheet.getSheetByName("Unavailable Dates");
    if (!unavailableDatesSheet) {
      throw new Error('Sheet "Unavailable Dates" not found!');
    }

    var lastRow = unavailableDatesSheet.getLastRow();
    var lastColumn = unavailableDatesSheet.getLastColumn();

    logMessage(getCallStackTrace() + `: "Unavailable Dates" sheet before cleanup - lastRow: ${lastRow}, lastColumn: ${lastColumn}`);

    // Clear content of all rows and columns except the first row (assumed to contain headers)
    if (lastRow > 1) {
      //clean out the 1st row, 2nd column, which is the 1st row with all the names, except column 1 of the 1st row, which is "Date"
      unavailableDatesSheet.getRange(1, 2, 1, lastColumn).clearContent();
      //clean out the rest of data
      unavailableDatesSheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
    }

    // Copy the date column from the source sheet into the "Unavailable Dates" sheet
    unavailableDatesSheet.getRange(2, 1, totalSrcRows, 1).setValues(srcValues);

    logMessage(getCallStackTrace() + `: "Unavailable Dates" sheet after copying date column - copied ${totalSrcRows} rows.`);

    setupUnavailableDatesSheet();

    cleanUnavailableDatesResponse();

    logMessage(getCallStackTrace() + `: "Done! Cleaned up all previous "Unavailable Date" and "Unavailable Date Response" data and updated with new values`);
    flushLogsToSheet();

    // Notify the user that the cleanup and data copy process is complete
    SpreadsheetApp.getActive().toast('Done! Cleaned up all previous "Unavailable Date" and "Unavailable Date Response" data and updated with new values 👍');

  } catch (error) {
    logMessage.Error(getCallStackTrace() + `: Error in cleanUnavailableDates: ${error.message}`);
    flushLogsToSheet();
    SpreadsheetApp.getActive().toast(`Error: ${error.message}`);

  }

}

function setupUnavailableDatesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const unavailableDatesSheet = ss.getSheetByName("Unavailable Dates");
  const contactSheet = ss.getSheetByName("contact");

  if (!unavailableDatesSheet || !contactSheet) {
    throw new Error("Either the Unavailable Dates sheet or 'contact' sheet is missing.");
  }

  const contactData = contactSheet.getRange("A2:E").getValues(); // A2:E covers name + responsibility

  const names = contactData
    .filter(row => {
      const name = row[0];
      if (!name || excludedNames.includes(name)) return false;

      const responsibilities = (row[4] || "").split(",").map(item => item.trim());
      return responsibilities.some(task => !excludedTasks.includes(task));
    })
    .map(row => row[0])
    .sort(); // Sort the names alphabetically

  logMessage(getCallStackTrace() + `: All the names get populated into the "Unavailable Dates" sheet= ${names}`);

  // Clear all old conditional formatting rules
  unavailableDatesSheet.clearConditionalFormatRules();

  // Remove all existing protections from the "Unavailable Dates" sheet
  const protections = unavailableDatesSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(protection => protection.remove());

  const startColumn = 2; // Column B
  const startRow = 2;    // Row 2 (row 1 is headers)

  // Set header row with names
  for (let i = 0; i < names.length; i++) {
    const cell = unavailableDatesSheet.getRange(1, startColumn + i);
    cell.setValue(names[i]);
    cell.setHorizontalAlignment("center"); // Center the header text
  }

  // Get all Sundays from column A (until blank)
  const sundays = unavailableDatesSheet.getRange("A2:A").getValues()
    .map(row => row[0])
    .filter(date => date instanceof Date); // Only valid dates

  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();

  // Add edit permissions to each name's column based on email
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    const emailRow = contactData.find(row => row[0] === name);
    const email = emailRow?.[1]; // Column B has emails

    if (email) {
      const column = startColumn + i;
      const protectionRange = unavailableDatesSheet.getRange(startRow, column, sundays.length);

      const protection = protectionRange.protect().setDescription(`Edit permission for ${name}`);
      protection.setWarningOnly(false); // Block others from editing

      // Remove all editors and add only the correct one
      protection.removeEditors(protection.getEditors());
      protection.addEditor(email);

      // Optionally remove the domain editor group if it's there
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
    }
  }

  const rules = [];

  for (let col = 0; col < names.length; col++) {
    const columnIndex = startColumn + col;
    const range = unavailableDatesSheet.getRange(startRow, columnIndex, sundays.length, 1);
    range.setDataValidation(checkboxRule);
    range.setHorizontalAlignment("center");

    // Dynamically get column letter (works for AA, AB, etc.)
    const colLetter = unavailableDatesSheet.getRange(1, columnIndex).getA1Notation().replace(/\d+/g, '');
    const formula = `=$${colLetter}${startRow}=TRUE`;

    const conditionalRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setFontColor("#FF0000")  // Change only checkmark color
      .setBackground(null)      // Keep background clear
      .setRanges([range])
      .build();

    rules.push(conditionalRule);
  }

  unavailableDatesSheet.setConditionalFormatRules(rules);
}

