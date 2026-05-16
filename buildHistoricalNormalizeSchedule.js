function buildHistoricalNormalizeSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheets()[0]; // Always the first sheet
  const targetSheetName = 'Historical Normalized Data';

  // Get or create the 'Normalized' sheet
  let targetSheet = ss.getSheetByName(targetSheetName);
  var output = [];
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
    output = [["Date", "Role", "Person"]];
  }

  const targetSheetLastRow = targetSheet.getLastRow() + 1;
  Logger.log("The starting row of the 'Historical Normalized' sheet to append data = " + targetSheetLastRow)

  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];
  const startColIndex = headers.indexOf("Worship Chairperson");

  if (startColIndex === -1) {
    throw new Error('"Worship Chairperson" column not found.');
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateCell = row[0];

    // Stop if the first cell is not a valid date
    if (!(dateCell instanceof Date)) {
      break;
    }

    for (let j = startColIndex; j < row.length; j++) {
      const role = headers[j];
      const person = row[j];

      // Skip empty cells or 'Speaker' column
      //if (!person || role === "Speaker") continue;
      if (!person) continue;

      output.push([dateCell, role, person]);
    }
  }

  // Write the normalized data to the 'Historical Normalized' sheet
  targetSheet.getRange(targetSheetLastRow, 1, output.length, 3).setValues(output);
}
