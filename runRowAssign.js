/**
 * ***************************************************************
 * runRowAssign.gs
 *
 * Assign ONE Sunday / row.
 *
 * Supports:
 *
 *   - chatgpt
 *   - deepseek
 *   - claude
 *   - gemini
 *   - noai
 *
 * The caller is responsible for:
 *
 *   - forced assignments
 *   - creating initial historicalCount
 *
 * This function:
 *
 *   1. Builds scheduling data for the current row
 *   2. Selects AI or GAS assignment method
 *   3. Validates assignments
 *   4. Writes assignments
 *   5. Updates historicalCount
 *
 * IMPORTANT:
 *
 *   AI path:
 *     historicalCount is updated HERE after successful write.
 *
 *   No-AI path:
 *     historicalCount is updated INSIDE runGASAssign()
 *     after EACH task.
 *
 * ***************************************************************
 */

function runRowAssign(
  rowNum,
  assignMethod,
  historicalCount,
  sheet
) {

  aiLog(
    `Starting row assignment for row ${rowNum} ` +
    `using ${assignMethod}`
  );

  /**************************************************************
   * STEP 1
   *
   * Build scheduling data for THIS Sunday only.
   **************************************************************/
  const buildResult =
    aiScheduleData(
      rowNum,
      historicalCount,
      sheet
    );

  const scheduleRows =
    buildResult.scheduleRows || [];

  if (scheduleRows.length === 0) {

    aiLog(
      `Row ${rowNum}: no schedule data`
    );

    return historicalCount;
  }

  const currentRow =
    scheduleRows[0];

  const dateDisplay =
    currentRow.date;

  /**************************************************************
   * Nothing left to assign.
   **************************************************************/
  if (
    !Array.isArray(currentRow.emptyTasks) ||
    currentRow.emptyTasks.length === 0
  ) {

    aiLog(
      `${dateDisplay}: no empty tasks to assign`
    );

    return historicalCount;
  }

  /**************************************************************
   * Determine assignment method.
   **************************************************************/
  const method =
    (assignMethod || "")
      .toLowerCase()
      .trim();

  const noAI =
    method === "noai" ||
    method === "no ai";

  let assignments = [];

  /**************************************************************
   *
   * NO-AI PATH
   *
   **************************************************************/
  if (noAI) {

    aiLog(
      `${dateDisplay}: using GAS assignment engine`
    );

    /************************************************************
     * runGASAssign() assigns ONE TASK AT A TIME.
     *
     * It updates:
     *
     *   historicalCount
     *   blockingWorkers
     *
     * immediately after each successful task decision.
     *
     * Therefore DO NOT update historicalCount again below.
     ************************************************************/
    assignments =
      runGASAssign(
        currentRow,
        historicalCount
      );

    /************************************************************
     * Validate GAS result.
     ************************************************************/
    aiValidator(
      assignments,
      [currentRow]
    );

    aiLog(
      `${dateDisplay}: GAS assignments passed validation`
    );

    /************************************************************
     * Write GAS assignments.
     ************************************************************/
    if (assignments.length > 0) {

      writeAIAssignments(
        assignments,
        sheet
      );

      aiLog(
        `${dateDisplay}: wrote ` +
        `${assignments.length} GAS assignment(s)`
      );
    }

    /************************************************************
     * IMPORTANT:
     *
     * DO NOT increment historicalCount here.
     *
     * runGASAssign() already did it after EACH task.
     ************************************************************/

    logLargeText(
      `=== HISTORICAL COUNT AFTER ${dateDisplay} ===`,
      JSON.stringify(
        historicalCount,
        null,
        2
      )
    );

    return historicalCount;
  }

  /**************************************************************
   *
   * AI PATH
   *
   * Remove internal row number before sending data to AI.
   *
   **************************************************************/
  const {
    _rowNum,
    ...rowForAI
  } = currentRow;

  const rowsForAI = [
    rowForAI
  ];

  const MAX_ATTEMPTS = 3;

  let unassigned = [];
  let priorErrors = null;

  /**************************************************************
   * STEP 2
   *
   * Call selected AI provider.
   **************************************************************/
  for (
    let attempt = 1;
    attempt <= MAX_ATTEMPTS;
    attempt++
  ) {

    const temperature =
      attempt === 1
        ? 0
        : 0.4;

    const retryContext =
      priorErrors
        ? priorErrors.concat(
            unassigned.length > 0
              ? [
                  "Unassigned slots: " +
                  unassigned.join("; ")
                ]
              : []
          )
        : null;

    aiLog(
      `${dateDisplay}: AI attempt ${attempt}`
    );

    const aiResult =
      aiCallEngine(
        method,
        rowsForAI,
        {
          temperature: temperature,
          priorErrors: retryContext
        }
      );

    assignments =
      Array.isArray(aiResult)
        ? aiResult
        : aiResult.assignments || [];

    /************************************************************
     * Only accept assignments for THIS Sunday.
     ************************************************************/
    assignments =
      assignments.filter(
        a => a.date === dateDisplay
      );

    /************************************************************
     * Find anything AI failed to assign.
     ************************************************************/
    unassigned =
      findUnassignedSlots(
        rowsForAI,
        assignments
      );

    if (unassigned.length === 0) {

      aiLog(
        `${dateDisplay}: assignment complete ` +
        `on attempt ${attempt}`
      );

      break;
    }

    aiLog(
      `${dateDisplay}: attempt ${attempt}: ` +
      `${unassigned.length} slot(s) remain unassigned`
    );

    unassigned.forEach(
      u => aiLog(u)
    );

    priorErrors = [
      `Unassigned slots: ${unassigned.join("; ")}`
    ];
  }

  /**************************************************************
   * STEP 3
   *
   * AI assignments have already been validated inside
   * aiCallEngine().
   *
   * Log final assignments.
   **************************************************************/
  aiLog(
    `${dateDisplay}: AI assignments generated: ` +
    `${assignments.length}`
  );

  assignments.forEach(a => {

    aiLog(
      `${a.date} | ${a.columnName} -> ${a.worker}` +
      (
        a.reason
          ? ` (${a.reason})`
          : ""
      )
    );

  });

  /**************************************************************
   * STEP 4
   *
   * Write successful AI assignments.
   **************************************************************/
  if (assignments.length > 0) {

    writeAIAssignments(
      assignments,
      sheet
    );

    aiLog(
      `${dateDisplay}: wrote ` +
      `${assignments.length} AI assignment(s)`
    );
  }

  /**************************************************************
   * STEP 5
   *
   * AI PATH ONLY:
   *
   * Update historicalCount after assignments were successfully
   * validated and written.
   *
   * Do NOT execute this for the No-AI path because GAS already
   * updates historicalCount task-by-task.
   **************************************************************/
  assignments.forEach(a => {

    historicalCount[a.worker] =
      (historicalCount[a.worker] || 0) + 1;

  });

  logLargeText(
    `=== HISTORICAL COUNT AFTER ${dateDisplay} ===`,
    JSON.stringify(
      historicalCount,
      null,
      2
    )
  );

  /**************************************************************
   * Final status.
   **************************************************************/
  if (unassigned.length > 0) {

    aiLog(
      `${dateDisplay}: ` +
      `${unassigned.length} slot(s) remain unassigned`
    );

    unassigned.forEach(
      u => aiLog(u)
    );

  } else {

    aiLog(
      `${dateDisplay}: all slots assigned`
    );
  }

  /**************************************************************
   * Return historical count for NEXT Sunday.
   **************************************************************/
  return historicalCount;
}
