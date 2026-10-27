// Update the dropDownValue for the taskName column pointed by headerRow[columnIndex]
// from startRow to lastRow, one row at a time. It checks the "Unavailable Dates" sheet
// to find out all the names unavailable for each week and removes them from the drop-down list.

function updateColumnDropDownUnDates(startRow, lastRow, headerRow, columnIndex, dropDownValue) {

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const exceptionDatesSheet = ss.getSheetByName("Unavailable Dates");
    const scheduleSheet = ss.getSheets()[0];

    const taskName = headerRow[columnIndex];

    logMessage(
      `${getCallStackTrace()} : Starting update for task="${taskName}", column=${columnIndex + 1}`
    );

    // Get exception data range
    const rowColStartEnd = getRowColStartEnd(exceptionDatesSheet);
    const [exRowStart, exRowEnd, exColumnStart, exColumnEnd] = rowColStartEnd[1];
    const exHeaderRow = rowColStartEnd[0];

    logMessage(
      `${getCallStackTrace()} : Unavailable Dates headerRow = ${JSON.stringify(exHeaderRow)}`
    );

    // Read all exception data in one API call
    // trueFalseRows: 2D array from getValues(), each row is [date, false, false, ..., true, ...]
    // returning [date, name, name...] if true, returning name,if false, return null, but null got filter out
    const trueFalseRows = exceptionDatesSheet
      .getRange(
        exRowStart,
        exColumnStart,
        exRowEnd - exRowStart + 1,
        exColumnEnd - exColumnStart + 1
      )
      .getValues();

    const exceptionRows = trueFalseRows.map(row => {
      // Always include the date (row[0])
      const date = row[0];
      // Start from index 1 because index 0 is the date
      const names = row
        .slice(1)
        .map((val, idx) => val === true ? exHeaderRow[idx + 1] : null)
        .filter(name => name);

      return [date, ...names];
    });

    logMessage(
      `${getCallStackTrace()} : Built ${exceptionRows.length} exception rows`
    );

    // Iterate over row indices to update dropdowns
    // Iterate through the exceptionRows
    // Create an array of row indices to iterate over (from startRow to lastRow)
    // rowIndices = [2,3,4,5,6,7,8,9,10,11,12,13,14]
    // Use forEach to iterate over each row
    // using rowIndices method below combined with getting all the exceptionRows at one time help the execution time about 6x
    // the original code getting one exception row at a time, api call in a loop waste a lot of time

    const rowIndices = Array.from(
      { length: lastRow - startRow + 1 },
      (_, i) => i + startRow
    );

    rowIndices.forEach(rowIndex => {

      try {

        const scheduleCell = scheduleSheet.getRange(
          rowIndex,
          columnIndex + 1
        );

        const a1Notation = scheduleCell.getA1Notation();

        const rowData = exceptionRows[rowIndex - startRow];

        if (!rowData) {
          throw new Error(
            `Missing exception row data for schedule row ${rowIndex}`
          );
        }

        const exceptionRow = rowData.filter(String);

        const dateValue = Utilities.formatDate(
          exceptionRow.shift(),
          Session.getScriptTimeZone(),
          "MM/dd/yyyy"
        );

        logMessage(
          `${getCallStackTrace()} : Processing cell="${a1Notation}", task="${taskName}", date="${dateValue}"`
        );

        logMessage(
          `${getCallStackTrace()} : Unavailable names = ${JSON.stringify(exceptionRow)}`
        );

        let newDropDownValue;

        if (!dropDownModifyExcludeList.includes(taskName)) {
          // Filter dropdown to exclude exceptions date for each person
          newDropDownValue = dropDownValue.filter(
            item => !exceptionRow.includes(item)
          );

        } else {

          logMessage(
            `${getCallStackTrace()} : Task "${taskName}" excluded from unavailable-date filtering`
          );

          newDropDownValue = dropDownValue;
        }

        logMessage(
          `${getCallStackTrace()} : Original dropdown size=${dropDownValue.length}, New dropdown size=${newDropDownValue.length}`
        );
        // Apply the new dropdown values to the sheet
        setDropDown(
          scheduleSheet,
          rowIndex,
          columnIndex,
          newDropDownValue
        );

        if (!arraysEqual(dropDownValue, newDropDownValue)) {

          logDropDownModification(
            rowIndex,
            columnIndex,
            taskName,
            dropDownValue,
            newDropDownValue,
            dateValue
          );
        }

      } catch (rowError) {

        logMessageError(
          `${getCallStackTrace()} : FAILED processing row=${rowIndex}, ` +
          `column=${columnIndex + 1}, task="${taskName}"\n` +
          `Error: ${rowError}\n` +
          `Stack: ${rowError.stack || "No stack trace"}`
        );

        throw rowError;
      }

    });

    logMessage(
      `${getCallStackTrace()} : Finished update for task="${taskName}"`
    );

    return true;

  } catch (error) {

    logMessageError(
      `${getCallStackTrace()} : FAILED updateColumnDropDownUnDates()
          Task="${headerRow[columnIndex]}"
          Column=${columnIndex + 1}
          Error=${error}
          Stack=${error.stack || "No stack trace"}`
    );

    return false;
  }
}


// Helper function to apply dropdown values to a specific cell
function setDropDown(sheet, rowIndex, columnIndex, dropDownValues) {

  const dropdownCell = sheet.getRange(
    rowIndex,
    columnIndex + 1
  );

  const a1Notation = dropdownCell.getA1Notation();

  try {

    logMessage(
      `${getCallStackTrace()} : Setting dropdown for cell="${a1Notation}"`
    );

    const rule = SpreadsheetApp
      .newDataValidation()
      .requireValueInList(dropDownValues)
      .build();

    dropdownCell.setDataValidation(rule);

    //SpreadsheetApp.flush();   // TEMPORARY DEBUGGING

    logMessage(
      `${getCallStackTrace()} : Successfully set dropdown for cell="${a1Notation}"`
    );

  } catch (error) {

    logMessageError(
      `${getCallStackTrace()} : FAILED setting dropdown for cell="${a1Notation}"\n` +
      `Row=${rowIndex}\n` +
      `Column=${columnIndex + 1}\n` +
      `Dropdown Values=${JSON.stringify(dropDownValues)}\n` +
      `Error=${error}\n` +
      `Stack=${error.stack || "No stack trace"}`
    );

    throw error;
  }
}
// Helper function to log dropdown changes
function logDropDownModification(rowIndex, columnIndex, taskName, originalValues, modifiedValues, dateValue) {
  logMessage(`${getCallStackTrace()} : Row: ${rowIndex}, Column: ${columnIndex + 1}, Task: ${taskName}, Date: ${dateValue}`);
  logMessage(`${getCallStackTrace()} : Original Dropdown: ${JSON.stringify(originalValues)}`);
  logMessage(`${getCallStackTrace()} : Modified Dropdown: ${JSON.stringify(modifiedValues)}`);
}


// Helper function to compare two arrays
function arraysEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) return false; // Different lengths
  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) return false; // Mismatch found
  }
  return true; // All elements match
}

