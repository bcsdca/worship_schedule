function updateAllDropDowns(arr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contactSheet = ss.getSheetByName("contact");
  const scheduleSheet = ss.getSheets()[0];
  const avSMEColumn = 5; // "Worship SME" is in column 5

  logMessage(getCallStackTrace() + `: Task names array pass in to build pull down:", ${JSON.stringify(arr)}`);

  //build SME Data with contact names
  const objectArray = buildSMEObjectArray(contactSheet, avSMEColumn);
  //logMessage(getCallStackTrace() + ": SME object array:", JSON.stringify(objectArray));

  // Find row and column range in the schedule sheet
  const rowColStartEnd = getRowColStartEnd(scheduleSheet);

  //dataRowStart is the row after the header row
  const dataRowStart = rowColStartEnd[1][0];
  const dataRowEnd = rowColStartEnd[1][1];
  const dataColumnStart = rowColStartEnd[1][2];
  const dataColumnEnd = rowColStartEnd[1][3];
  const totalRows = dataRowEnd - dataRowStart + 1;
  const headerRow = rowColStartEnd[0];

  logMessage(getCallStackTrace() + `: Schedule sheet's dataRowStart = ${dataRowStart}, dataRowEnd = ${dataRowEnd}, dataColumnStart = ${dataColumnStart}, dataColumnEnd = ${dataColumnEnd}, total # of rows = ${totalRows}`);


  // Iterate over header columns, starting at column 5 (array index 4)
  for (let headerRowIndex = 4; headerRowIndex < headerRow.length; headerRowIndex++) {
    const taskName = headerRow[headerRowIndex];

    if (arr) {
      // If arr is defined, update dropdown for only the tasks in arr
      if (arr.includes(taskName)) {
        logMessage(getCallStackTrace() + `: Task name array pass in was defined, updating dropdown for task name: ${taskName}`);
        //this dropDownValues has all the names that are qualified for this taskname, based on the column5 of the contact sheet
        //this is an ideal list, and might be less when it finally insert into the sschedule sheet, 
        //because of the unavialable dates in exception date 
        const dropDownValues = searchSMEObjectArray(objectArray, taskName);
        updateColumnDropDownUnDates(dataRowStart, dataRowEnd, headerRow, headerRowIndex, dropDownValues);

      }
    } else {
      // If arr is undefined, regenerate dropdowns for all tasks
      logMessage(getCallStackTrace() + `: No task names array pass in, updating dropdown for task name: ${taskName}`);
      const dropDownValues = searchSMEObjectArray(objectArray, taskName);
      updateColumnDropDownUnDates(dataRowStart, dataRowEnd, headerRow, headerRowIndex, dropDownValues);
    }
  }

  SpreadsheetApp.getActive().toast("Done! Successfully updated all worship schedule dropdowns 👍");
  logMessage(getCallStackTrace() + `: updateScheduleDropDown: Done updating all dropdowns.`);
}


