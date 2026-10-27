function runAssignAll(assignMethod) {
  clearLogSheet();
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheets()[0];

  if (!sheet) {
    throw new Error(`sheet not found`);
  }

  aiLog(`Using sheet: ${sheet.getName()}`);

  /**************************************************************
   * Resolve/write all forced assignments and return the
   * initial historical count.
   **************************************************************/
  let currentHistoricalCount =
    runForceAssign(sheet);

  /**************************************************************
   * Find schedule rows.
   **************************************************************/
  const bounds =
    getScheduleRowBounds(sheet);

  if (!bounds) {
    aiLog(
      `first row and last row to assign task were not found`
    );
    return;
  }

  const firstRow = bounds.firstRow;
  const lastRow = bounds.lastRow;

  aiLog(
    `Using ${assignMethod} to assign rows ` +
    `${firstRow} through ${lastRow}`
  );

  /**************************************************************
   * Assign ONE Sunday at a time.
   **************************************************************/
  for (
    let row = firstRow;
    row <= lastRow;
    row++
  ) {

    currentHistoricalCount =
      runRowAssign(
        row,
        assignMethod,
        currentHistoricalCount,
        sheet
      );
  }

  aiLog(
    `Schedule assignment completed using ${assignMethod}`
  );
  flushLogsToSheet();
}



