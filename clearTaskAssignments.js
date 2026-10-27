/**********************************************************************
 * Clear all task assignments.
 * Starting Column = "Power Point Preparation"
 * Ending Column = "Mandarin A/V helper"
 * Clears values only for rows whose column A contains a Date.
 **********************************************************************/

/**
 * ***************************************************************
 * clearTaskAssignments()
 *
 * Clear all task assignments.
 *
 * Starting Column = "Power Point Preparation"
 * Ending Column   = "Mandarin A/V helper"
 *
 * Clears values only for rows whose Column A contains a Date.
 *
 * The sheet object is supplied by the caller so the spreadsheet
 * does not need to be opened again.
 * ***************************************************************
 */

function clearTaskAssignments(sheet) {

  if (!sheet) {
    throw new Error(
      "clearTaskAssignments: sheet is required."
    );
  }

  const headerInfo = buildHeaderMap(sheet); 
  const headerMap = headerInfo.headerMap;

  const startCol =
    headerMap["Power Point Preparation"];

  const endCol =
    headerMap["Mandarin A/V helper"];

  if (!startCol || !endCol) {
    throw new Error(
      "Task columns not found."
    );
  }

  const bounds =
    getScheduleRowBounds(sheet);

  if (!bounds) {
    return;
  }

  sheet
    .getRange(
      `${startCol}${bounds.firstRow}:` +
      `${endCol}${bounds.lastRow}`
    )
    .clearContent();

  aiLog(
    `Cleared task assignments ` +
    `(${startCol}${bounds.firstRow}:` +
    `${endCol}${bounds.lastRow})`
  );
}


/**
 * ***************************************************************
 * getScheduleRowBounds()
 *
 * Find the first and last worksheet rows whose Column A
 * contains a valid Date.
 *
 * Returns:
 *
 * {
 *   firstRow: 2,
 *   lastRow: 15
 * }
 *
 * Returns null if no valid schedule rows are found.
 *
 * The sheet object is passed in by the caller so we do not
 * repeatedly access the spreadsheet.
 * ***************************************************************
 */

function getScheduleRowBounds(sheet) {

  if (!sheet) {
    throw new Error(
      "getScheduleRowBounds: sheet is required."
    );
  }

  const lastSheetRow = sheet.getLastRow();

  if (lastSheetRow < 1) {
    return null;
  }

  const dates =
    sheet
      .getRange(1, 1, lastSheetRow, 1)
      .getValues();

  let firstRow = null;
  let lastRow = null;

  dates.forEach((row, i) => {

    if (!(row[0] instanceof Date)) {
      return;
    }

    const rowNum = i + 1;

    if (firstRow === null) {
      firstRow = rowNum;
    }

    lastRow = rowNum;
  });

  if (firstRow === null) {
    aiLog("No schedule rows found.");
    return null;
  }

  return {
    firstRow: firstRow,
    lastRow: lastRow
  };
}

