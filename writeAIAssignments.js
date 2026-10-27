
/**
 * ***************************************************************
 * writeAIAssignments.gs
 *
 * Write validated AI assignments to the worship schedule.
 *
 * The AI is now called for ONE Sunday at a time, but this
 * function accepts an array so it remains compatible with
 * the existing AI engine.
 *
 * Expected assignment format:
 *
 * {
 *   date: "7/5/2026",
 *   columnName: "Camera Person",
 *   worker: "Cyrus Tang",
 *   reason: "Lowest historical count"
 * }
 *
 * This function:
 *
 *   1. Finds the worksheet row by date.
 *   2. Finds the task column by column name.
 *   3. Verifies the cell is still empty.
 *   4. Writes the worker.
 *   5. Logs the assignment.
 *
 * It does NOT:
 *
 *   - resolve forced assignments
 *   - validate AI assignments
 *   - calculate historicalCount
 *   - calculate futureCount
 *
 * Those responsibilities belong to other modules.
 * ***************************************************************
 */

function writeAIAssignments(assignments,sheet) {

  if (!Array.isArray(assignments)) {
    throw new Error(
      "writeAIAssignments: assignments must be an array."
    );
  }

  if (assignments.length === 0) {
    aiLog("writeAIAssignments: No assignments to write.");
    return;
  }
 
  if (!sheet) {
    throw new Error(
      `Sheet not found`
    );
  }

  const headerInfo = buildHeaderMap(sheet); 
  const headerMap = headerInfo.headerMap; 
 
  let written = 0;

  assignments.forEach(a => {

    if (!a.date) {
      throw new Error(
        "AI assignment is missing date."
      );
    }

    if (!a.columnName) {
      throw new Error(
        `AI assignment on ${a.date} is missing columnName.`
      );
    }

    if (!a.worker) {
      throw new Error(
        `${a.date} | ${a.columnName}: missing worker.`
      );
    }

    /************************************************************
     * Find task column
     ************************************************************/
    const col = headerMap[a.columnName];

    if (!col) {
      throw new Error(
        `${a.date} | ${a.columnName}: column not found.`
      );
    }

    /************************************************************
     * Find worksheet row by displayed date.
     *
     * A1 notation is intentionally retained.
     ************************************************************/
    let rowNum = null;

    for (
      let r = 2;
      r <= sheet.getLastRow();
      r++
    ) {

      const dateDisplay =
        sheet.getRange(`A${r}`).getDisplayValue();

      if (dateDisplay === a.date) {
        rowNum = r;
        break;
      }
    }

    if (!rowNum) {
      throw new Error(
        `${a.date}: worksheet row not found.`
      );
    }

    /************************************************************
     * Make sure the task is still empty.
     *
     * This protects against accidentally overwriting an
     * assignment made after the AI data was generated.
     ************************************************************/
    const cell =
      sheet.getRange(`${col}${rowNum}`);

    const existingValue =
      cell.getValue();

    if (existingValue) {
      throw new Error(
        `${a.date} | ${a.columnName}: ` +
        `cell already contains "${existingValue}".`
      );
    }

    /************************************************************
     * Write AI assignment
     * Italic
     ************************************************************/
    cell.setValue(a.worker).setFontStyle("italic");;

    written++;

    aiLog(
      `AI assignment written: ` +
      `${a.date} | ${a.columnName} → ${a.worker}` +
      (a.reason
        ? ` (${a.reason})`
        : "")
    );
  });

  aiLog(
    `writeAIAssignments: ${written} assignment(s) written.`
  );

  return written;
}

