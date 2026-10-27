
/**
 * ***************************************************************
 * aiValidator.gs
 *
 * Validate AI-generated worship schedule assignments.
 *
 * The AI engine now processes ONE Sunday at a time.
 *
 * Therefore scheduleRows normally contains exactly ONE row.
 *
 * Validation checks:
 *
 *   1. Date exists in the current schedule data
 *   2. Task is actually empty
 *   3. Worker is in the task's valid options
 *   4. Blocking rule is respected for that Sunday
 *   5. No duplicate task assignment
 * ***************************************************************
 */

function aiValidator(
  assignments,
  scheduleRows
) {

  /**************************************************************
   * Normalize inputs
   **************************************************************/
  if (!Array.isArray(assignments)) {
    assignments = [];
  }

  if (!Array.isArray(scheduleRows)) {
    scheduleRows = [];
  }

  aiLog(
    `Validating ${assignments.length} AI assignments`
  );

  const errors = [];

  /**************************************************************
   * Build date → schedule row lookup.
   *
   * Since runAIAssign() now sends one Sunday at a time,
   * this normally contains one entry.
   **************************************************************/
  const rowMap = {};

  scheduleRows.forEach(row => {

    if (!row || !row.date) {
      return;
    }

    rowMap[row.date] = row;

  });

  /**************************************************************
   * Validate each AI assignment.
   **************************************************************/
  assignments.forEach(a => {

    /************************************************************
     * Basic assignment structure check
     ************************************************************/
    if (
      !a ||
      !a.date ||
      !a.columnName ||
      !a.worker
    ) {

      errors.push(
        "Invalid AI assignment structure"
      );

      return;
    }

    /************************************************************
     * Find the corresponding Sunday.
     ************************************************************/
    const row =
      rowMap[a.date];

    if (!row) {

      errors.push(
        `${a.date}: date not found`
      );

      return;
    }

    /************************************************************
     * Find the task in the current Sunday's emptyTasks.
     ************************************************************/
    const task =
      row.emptyTasks.find(
        t =>
          t.columnName ===
          a.columnName
      );

    if (!task) {

      errors.push(
        `${a.date} | ${a.columnName}: task not empty`
      );

      return;
    }

    /************************************************************
     * Priority 1 — Dropdown validity
     *
     * The worker must be one of the workers available for
     * THIS task on THIS Sunday.
     ************************************************************/
    if (
      !task.options.some(
        o => o.worker === a.worker
      )
    ) {

      errors.push(
        `${a.date} | ${a.columnName}: ` +
        `${a.worker} not in dropdown`
      );

    }

    /************************************************************
     * Priority 2 — Blocking rule
     *
     * A blocking task cannot use a worker who is already
     * assigned to another blocking role on THIS Sunday.
     ************************************************************/
    if (
      task.blocking &&
      row.blockingWorkers.includes(
        a.worker
      )
    ) {

      errors.push(
        `${a.date} | ${a.columnName}: ` +
        `${a.worker} violates blocking rule`
      );

    }

  });

  /**************************************************************
   * Duplicate task check
   *
   * The same task on the same Sunday may only be assigned once.
   **************************************************************/
  const seen = new Set();

  assignments.forEach(a => {

    if (
      !a ||
      !a.date ||
      !a.columnName
    ) {
      return;
    }

    const key =
      `${a.date}|${a.columnName}`;

    if (seen.has(key)) {

      errors.push(
        `${key}: duplicate assignment`
      );

    }

    seen.add(key);

  });

  /**************************************************************
   * Final validation result
   **************************************************************/
  if (errors.length === 0) {

    aiLog(
      "✅ AI assignments passed validation"
    );

    return true;

  }

  /**************************************************************
   * Validation failed
   **************************************************************/
  aiLog(
    `❌ Found ${errors.length} validation errors`
  );

  errors.forEach(e => {
    aiLog("  " + e);
  });

  const error =
    new Error(
      "AI assignment validation failed:\n" +
      errors.join("\n")
    );

  error.validationErrors =
    errors;

  throw error;
}
