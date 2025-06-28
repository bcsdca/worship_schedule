function getWorkerName(sheet, row, targetTaskName, headerRow) {
  // Find the column index for the target task
  var targetTaskColumnNum = headerRow.indexOf(targetTaskName);

  // Handle the case where the target task name is not found in the header row
  if (targetTaskColumnNum === -1) {
    throw new Error(`${getCallStackTrace()}: Target task name "${targetTaskName}" not found in the header row.`);
  }

  // Get the worker name from the target cell
  // Note: Row and column are 1-based in Google Sheets API, hence adding 1 to column index
  var workerName = sheet.getRange(row, targetTaskColumnNum + 1).getValue();

  // Log the information if a worker's name is found
  if (workerName) {
    logMessage(`${getCallStackTrace()}: Found worker "${workerName}" in row ${row}, column ${targetTaskColumnNum + 1}, task "${targetTaskName}".`);
  }else {
    logMessage(`${getCallStackTrace()}: No worker found in row ${row}, column ${targetTaskColumnNum + 1}, for task "${targetTaskName}"...`);
  }

  return workerName;
}