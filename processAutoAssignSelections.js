function processAutoAssignSelections(runOption, formData) {
  clearLogSheet();
  logMessage(`${getCallStackTrace()}: Form data object: ${JSON.stringify(formData)}, and the selected run option is "${runOption}"`);

  // Get the list of selected options and reorder based on priority
  let selectedOptions = Object.keys(formData);
  selectedOptions = reorderTasksByPriority(selectedOptions, priorityOrder);

  logMessage(`${getCallStackTrace()}: All tasks selected with the correct prioity order based on \"priorityOrder\" : ${JSON.stringify(selectedOptions)}`);

  // Get the spreadsheet and necessary row/column info
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const rowColStartEnd = getRowColStartEnd(sheet);
  const headerRow = rowColStartEnd[0];
  const startRow = rowColStartEnd[1][0];
  const endRow = rowColStartEnd[1][1];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contactSheet = ss.getSheetByName("contact");
  const avSMEColumn = 5; // "Worship SME" is in column 5

  // Build SME object array and assign tasks
  const smeObjectArray = buildSMEObjectArray(contactSheet, avSMEColumn);
  //logMessage(getCallStackTrace() + ": SME object array:", JSON.stringify(smeObjectArray));

  //clean up all the old data 1st before we start
  selectedOptions.forEach(task => {
    //index is 0-base, and column is 1-base
    const taskColumnNum = headerRow.indexOf(task) + 1;
    sheet.getRange(startRow, taskColumnNum, endRow - startRow + 1, 1).clearContent();
    logMessage(`${getCallStackTrace()}: Just delete old data on task Column "${taskColumnNum}" for Task name "${task}"`);
  });

  // Filter names based on smeLabel for all the tasks in smeLabelsToFilter array  and ensure uniqueness using Set
  var workerQueue = [...new Set(
    smeObjectArray
      .filter(item => smeLabelsToFilter.includes(item.smeLabel))
      .map(item => item.name)
  )];

  //logMessage(`${getCallStackTrace()}: Starting worker queue: %s for Task name: %s, and auto assignment is running in %s mode`, JSON.stringify(workerQueue), taskName, runOption);

  var workerCount = new Map();
  workerQueue.forEach(worker => workerCount.set(worker, 0));

  selectedOptions.forEach(task => {
    const taskColumnNum = headerRow.indexOf(task);
    const col_id = String.fromCharCode(64 + (taskColumnNum + 1));
    logMessage(`${getCallStackTrace()}: Starting Working on the task Column id of "${col_id}" for Task name "${task}"`);
    [workerQueue, workerCount] = assignTasks(sheet, task, col_id, workerQueue, workerCount, startRow, endRow, headerRow, runOption);
  });

  logMessage(`${getCallStackTrace()}: Done, finished the following tasks automatic assignment: ${JSON.stringify(selectedOptions)}`);

  buildPivotTable()
  logMessage(`${getCallStackTrace()}: Done, rebuilding the Dash Board !!!`);
  flushLogsToSheet();
  return `Done, you have selected the following tasks "${selectedOptions.join(', ')}" for automatic worker assignment 👍 !!!`;
  
}

function reorderTasksByPriority(selectedOptions, priorityOrder) {
  // Sort the selected options based on their position in the priorityOrder array
  return selectedOptions.sort((a, b) => {
    const priorityA = priorityOrder.indexOf(a);
    const priorityB = priorityOrder.indexOf(b);
    return priorityA - priorityB;
  });
}
