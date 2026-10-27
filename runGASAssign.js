/**
 * ***************************************************************
 * runGASAssign.gs
 *
 * Assign ONE Sunday using deterministic GAS logic.
 *
 * Selection priority:
 *
 *   1. Hard constraints
 *   2. Lowest historicalCount
 *   3. If tied, lowest futureCount
 *   4. If tied, first valid worker in task.options
 *
 * IMPORTANT:
 *
 *   - historicalCount is updated AFTER EACH successful task.
 *   - blockingWorkers is updated AFTER EACH successful
 *     blocking task.
 *   - Both updated values are used immediately by the next task
 *     on the same Sunday.
 *
 *   - historicalCount must NOT be incremented again outside
 *     this function for the No-AI path.
 *
 * @param {Object} row
 *   One schedule row produced by aiScheduleData().
 *
 * @param {Object} historicalCount
 *   Current cumulative historical counts.
 *
 * @return {Array}
 *   Assignments generated for this Sunday.
 * ***************************************************************
 */

function runGASAssign(row, historicalCount) {

  aiLog(`Running GAS assignment for ${row.date}`);

  const assignments = [];

  if (!Array.isArray(row.emptyTasks) || row.emptyTasks.length === 0) {
    aiLog(`${row.date}: no empty tasks for GAS`);
    return assignments;
  }

  /**************************************************************
   * Live blocking state for THIS Sunday.
   *
   * Start with workers already blocking before GAS assignments.
   **************************************************************/
  const blockingWorkers = new Set(row.blockingWorkers || []);

  /**************************************************************
   * Process tasks ONE AT A TIME, in emptyTasks order.
   **************************************************************/
  row.emptyTasks.forEach(task => {

    const columnName = task.columnName;
    const options = Array.isArray(task.options) ? task.options : [];

    if (!options.length) {
      aiLog(`${row.date} | ${columnName}: no options`);
      return;
    }

    /************************************************************
     * STEP 1 — Hard constraints
     *
     * aiScheduleData() already filtered the original options.
     *
     * However, blockingWorkers may have changed because GAS
     * assigned an earlier task on THIS SAME Sunday.
     ************************************************************/
    const validOptions = options.filter(option => {

      const worker = option.worker;

      if (!worker) return false;

      if (
        task.blocking === true &&
        blockingWorkers.has(worker)
      ) {
        return false;
      }

      return true;
    });

    if (!validOptions.length) {
      aiLog(
        `${row.date} | ${columnName}: no valid workers remain`
      );
      return;
    }

    /************************************************************
     * STEP 2 — Start with FIRST valid option.
     *
     * This automatically becomes the final tie-breaker if
     * historicalCount and futureCount are both equal.
     ************************************************************/
    let selectedWorker = validOptions[0].worker;

    let selectedHistorical =
      historicalCount[selectedWorker] || 0;

    let selectedFuture =
      Number(task.futureCount?.[selectedWorker]) || 0;

    /************************************************************
     * Compare remaining valid workers.
     ************************************************************/
    for (let i = 1; i < validOptions.length; i++) {

      const worker = validOptions[i].worker;

      const historical =
        historicalCount[worker] || 0;

      const future =
        Number(task.futureCount?.[worker]) || 0;

      /**********************************************************
       * PRIORITY 1 — Lowest historicalCount
       **********************************************************/
      if (historical < selectedHistorical) {

        selectedWorker = worker;
        selectedHistorical = historical;
        selectedFuture = future;

        continue;
      }

      /**********************************************************
       * PRIORITY 2 — Lowest futureCount
       *
       * ONLY when historicalCount is tied.
       **********************************************************/
      if (
        historical === selectedHistorical &&
        future < selectedFuture
      ) {

        selectedWorker = worker;
        selectedHistorical = historical;
        selectedFuture = future;

        continue;
      }

      /**********************************************************
       * PRIORITY 3 — First valid option
       *
       * If historical AND future counts are tied,
       * do nothing.
       *
       * The earlier option remains selected.
       **********************************************************/
    }

    /************************************************************
     * STEP 3 — Determine reason BEFORE updating history.
     ************************************************************/
    let reason;

    if (validOptions.length === 1) {

      reason =
        "Only eligible worker";

    } else {

      const lowestHistorical =
        Math.min(
          ...validOptions.map(option =>
            historicalCount[option.worker] || 0
          )
        );

      const historicalCandidates =
        validOptions.filter(option =>
          (historicalCount[option.worker] || 0) ===
          lowestHistorical
        );

      if (historicalCandidates.length === 1) {

        reason =
          "Lowest historical count";

      } else {

        const lowestFuture =
          Math.min(
            ...historicalCandidates.map(option =>
              Number(task.futureCount?.[option.worker]) || 0
            )
          );

        const futureCandidates =
          historicalCandidates.filter(option =>
            (Number(task.futureCount?.[option.worker]) || 0) ===
            lowestFuture
          );

        if (futureCandidates.length === 1) {

          reason =
            "Historical count tied; lowest future count";

        } else {

          reason =
            "Historical and future counts tied; first valid option";
        }
      }
    }

    /************************************************************
     * STEP 4 — Final local validation before changing state.
     ************************************************************/
    if (
      !validOptions.some(
        option => option.worker === selectedWorker
      )
    ) {
      throw new Error(
        `${row.date} | ${columnName}: ` +
        `${selectedWorker} is not a valid option`
      );
    }

    if (
      task.blocking === true &&
      blockingWorkers.has(selectedWorker)
    ) {
      throw new Error(
        `${row.date} | ${columnName}: ` +
        `${selectedWorker} violates blocking rule`
      );
    }

    /************************************************************
     * STEP 5 — Create successful assignment.
     ************************************************************/
    const assignment = {
      date: row.date,
      columnName: columnName,
      worker: selectedWorker,
      reason: reason
    };

    assignments.push(assignment);

    aiLog(
      `${row.date} | ${columnName} -> ${selectedWorker}` +
      ` | historical=${selectedHistorical}` +
      ` | future=${selectedFuture}` +
      ` | ${reason}`
    );

    /************************************************************
     * STEP 6 — Update historicalCount IMMEDIATELY.
     *
     * This is the SAME historicalCount object passed in by
     * runRowAssign().
     *
     * The NEXT task on THIS SAME Sunday sees the updated count.
     ************************************************************/
    historicalCount[selectedWorker] =
      (historicalCount[selectedWorker] || 0) + 1;

    aiLog(
      `${selectedWorker} historicalCount: ` +
      `${selectedHistorical} -> ` +
      `${historicalCount[selectedWorker]}`
    );

    /************************************************************
     * STEP 7 — Update blockingWorkers IMMEDIATELY.
     *
     * Only blocking assignments affect this list.
     *
     * Non-blocking assignments require no separate worker list.
     ************************************************************/
    if (task.blocking === true) {

      blockingWorkers.add(selectedWorker);

      aiLog(
        `${selectedWorker} added to blockingWorkers`
      );
    }

  });

  /**************************************************************
   * Final status.
   **************************************************************/
  aiLog(
    `${row.date}: GAS generated ` +
    `${assignments.length} assignment(s)`
  );

  aiLog(
    `${row.date}: final blockingWorkers = ` +
    JSON.stringify([...blockingWorkers])
  );

  return assignments;
}