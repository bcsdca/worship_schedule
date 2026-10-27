/**********************************************************************
 * writeForcedAssignments.gs
 *
 * Write resolved forced assignments to the worship schedule.
 **********************************************************************/

function writeForcedAssignments(forcedAssignments, sheet) {
  aiLog(`Writing ${forcedAssignments.length} forced assignments`);

  if (!forcedAssignments || !forcedAssignments.length) {
    aiLog("No forced assignments to write");
    return;
  }
  
  if (!sheet) throw new Error(`sheet not found`);

  aiLog(`Using sheet: ${sheet.getName()}`);

  const headerInfo = buildHeaderMap(sheet); 
  const headerMap = headerInfo.headerMap;

  forcedAssignments.forEach(a => {

    const col = headerMap[a.columnName];

    if (!col) {
      aiLog(
        `Column not found: ${a.columnName}`
      );
      return;
    }

    if (!a._rowNum) {
      aiLog(
        `Row number missing: ${a.date} | ${a.columnName}`
      );
      return;
    }

    const cell = sheet.getRange(`${col}${a._rowNum}`);

    // Safety check: never overwrite an existing assignment.
    if (cell.getValue()) {
      aiLog(
        `Skipped existing assignment: ${a.date} | ${a.columnName} | ${cell.getValue()}`
      );
      return;
    }

    cell.setValue(a.worker).setFontStyle("normal");

    aiLog(
      `Forced written: ${a.date} | ${a.columnName} -> ${a.worker} [${a.source}]`
    );

  });

  SpreadsheetApp.flush();

  aiLog(
    `Finished writing ${forcedAssignments.length} forced assignments`
  );
}
