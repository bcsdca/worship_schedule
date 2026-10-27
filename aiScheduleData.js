/**
 * aiScheduleData.gs
 *
 * Build structured AI schedule data for ONE Sunday.
 *
 * Includes:
 *   - historicalCount
 *   - futureCount
 *   - occupiedRoles
 *   - blockingWorkers
 *   - emptyTasks
 *
 * rowNum identifies the Sunday being processed.
 * historicalCount is supplied by the calling process.
 *
 * Uses a worksheet snapshot so futureCount calculations do not
 * repeatedly access Spreadsheet services.
 */

function aiScheduleData(rowNum, historicalCount = {}, sheet) {
  if (!sheet) throw new Error(`Sheet not found`);

  aiLog(`Building schedule data for row ${rowNum} from ${sheet.getName()}`);

  if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
    throw new Error(`Invalid row number: ${rowNum}`);
  }

  const headerInfo = buildHeaderMap(sheet);
  const headerMap = headerInfo.headerMap;

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const displayValues = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();

  const snapshot = buildScheduleSnapshot(
    sheet,
    headerMap,
    values,
    lastRow
  );

  const rowIndex = rowNum - 1;
  const date = values[rowIndex][0];

  if (!(date instanceof Date)) {
    throw new Error(`Row ${rowNum} does not contain a valid date`);
  }

  const dateDisplay = displayValues[rowIndex][0];
  const snapshotRow = snapshot[rowNum];

  if (!snapshotRow) {
    throw new Error(`No schedule data found for row ${rowNum}`);
  }

  const row = {
    date: dateDisplay,
    weekOfMonth: getWeekOfMonth(date),
    context: {},
    emptyTasks: [],
    occupiedRoles: [],
    blockingWorkers: [],
    _rowNum: rowNum
  };

  const blocking = new Set();

  /**************************************************************
   * Existing context assignments
   **************************************************************/
  CONTEXT_COLUMN_NAMES.forEach(name => {
    const worker = snapshotRow.context[name];
    if (!worker) return;

    row.context[name] = worker;
    addRole(row, blocking, name, worker);
  });

  // Log blocking workers from existing context assignments
  aiLog(
    `${dateDisplay}: blocking workers after context assignments = ` +
    JSON.stringify([...blocking])
  );

  /**************************************************************
   * Existing task assignments
   **************************************************************/
  TASK_COLUMN_NAMES.forEach(name => {
    const worker = snapshotRow.tasks[name];
    if (!worker) return;

    addRole(row, blocking, name, worker);
  });

  aiLog(
  `${dateDisplay}: blocking workers after task assignments = ` +
  JSON.stringify([...blocking])
);

  /**************************************************************
   * Build empty tasks for THIS Sunday only
   **************************************************************/
  TASK_COLUMN_NAMES.forEach(name => {
    const col = headerMap[name];
    if (!col) return;

    // Already assigned.
    if (snapshotRow.tasks[name]) return;

    let options = snapshotRow.dropdownOptions[name] || [];

    options = options.map(worker => ({ worker }));

    options = applyAvailabilityRules(
      name,
      row.weekOfMonth,
      options
    );

    /************************************************************
     * Blocking task:
     *
     * Only workers not already blocked THIS Sunday.
     ************************************************************/
    if (isBlockingTask(name)) {
      options = options.filter(
        option => !blocking.has(option.worker)
      );
    }

    if (!options.length) {
      aiLog(`No valid options: ${dateDisplay} | ${name}`);
      return;
    }

    /************************************************************
     * Single-option assignments should already have been
     * resolved by resolveForcedAssignments().
     ************************************************************/
    if (options.length === 1) {
      aiLog(
        `single-option already resolved: ${dateDisplay} | ${name} -> ${options[0].worker}`
      );
      return;
    }

    /************************************************************
     * Historical count:
     *
     * Only workers who are valid choices THIS Sunday.
     ************************************************************/
    const taskHistoricalCount = {};

    options.forEach(option => {
      taskHistoricalCount[option.worker] =
        historicalCount[option.worker] || 0;
    });

    /************************************************************
     * Future count:
     *
     * Only count future opportunities for workers who are
     * valid choices THIS Sunday.
     ************************************************************/
    const futureCount = getFutureCount(
      snapshot,
      headerMap,
      name,
      rowNum,
      options
    );

    row.emptyTasks.push({
      columnName: name,
      options: options,
      blocking: isBlockingTask(name),
      historicalCount: taskHistoricalCount,
      futureCount: futureCount
    });
  });

  row.blockingWorkers = [...blocking];

  const result = {
    scheduleRows: [row]
  };

  aiLog(
    `Schedule data built for ${dateDisplay}: ${row.emptyTasks.length} AI tasks`
  );

  logLargeText(
    "=== AI INPUT ===",
    JSON.stringify(result, null, 2)
  );

  return result;
}


/**************************************************************
 * buildScheduleSnapshot()
 *
 * Build all worksheet information needed by aiScheduleData().
 *
 * Spreadsheet values are read once. Dropdown validation is also
 * resolved once for every empty task cell.
 **************************************************************/

function buildScheduleSnapshot(sheet, headerMap, values, lastRow) {
  const snapshot = {};

  for (let r = 2; r <= lastRow; r++) {
    const rowIndex = r - 1;
    const date = values[rowIndex][0];

    if (!(date instanceof Date)) continue;

    const snapshotRow = {
      date: date,
      context: {},
      tasks: {},
      dropdownOptions: {},
      blockingWorkers: new Set()
    };

    /************************************************************
     * Context
     ************************************************************/
    CONTEXT_COLUMN_NAMES.forEach(name => {
      const col = headerMap[name];
      if (!col) return;

      const worker = sheet.getRange(`${col}${r}`).getValue();
      if (!worker) return;

      snapshotRow.context[name] = worker;

      if (isBlockingTask(name)) {
        snapshotRow.blockingWorkers.add(worker);
      }
    });

    /************************************************************
     * Tasks
     ************************************************************/
    TASK_COLUMN_NAMES.forEach(name => {
      const col = headerMap[name];
      if (!col) return;

      const worker = sheet.getRange(`${col}${r}`).getValue();

      if (worker) {
        snapshotRow.tasks[name] = worker;

        if (isBlockingTask(name)) {
          snapshotRow.blockingWorkers.add(worker);
        }

        return;
      }

      const cell = sheet.getRange(`${col}${r}`);

      snapshotRow.dropdownOptions[name] =
        getDropdownOptions(cell);
    });

    snapshot[r] = snapshotRow;
  }

  return snapshot;
}


/**************************************************************
 * getFutureCount()
 *
 * Count future opportunities for one specific task.
 *
 * Only workers who are available for the CURRENT task are
 * included.
 *
 * A future opportunity exists only when:
 *   1. Future row contains a valid date.
 *   2. Future task is completely empty.
 *   3. Worker is in the future task dropdown.
 *   4. Worker is allowed by availability rules.
 *   5. For blocking tasks, worker is not already blocked
 *      on that same future Sunday.
 **************************************************************/

function getFutureCount(snapshot, headerMap, taskName, currentRow, currentOptions) {
  const futureCount = {};

  const currentWorkers = new Set(
    currentOptions.map(option => option.worker)
  );

  const rows = Object.keys(snapshot)
    .map(Number)
    .filter(r => r > currentRow)
    .sort((a, b) => a - b);

  rows.forEach(r => {
    const futureRow = snapshot[r];

    // Future task already assigned -> not an opportunity.
    if (futureRow.tasks[taskName]) return;

    let options = futureRow.dropdownOptions[taskName] || [];

    options = options
      .filter(worker => currentWorkers.has(worker))
      .map(worker => ({ worker }));

    if (!options.length) return;

    const futureWeek = getWeekOfMonth(futureRow.date);

    options = applyAvailabilityRules(
      taskName,
      futureWeek,
      options
    );

    // Blocking check applies only to the SAME future Sunday.
    if (isBlockingTask(taskName)) {
      options = options.filter(
        option => !futureRow.blockingWorkers.has(option.worker)
      );
    }

    options.forEach(option => {
      if (futureCount[option.worker] === undefined) {
        futureCount[option.worker] = 0;
      }

      futureCount[option.worker]++;
    });
  });

  /**************************************************************
   * Make sure every current worker appears, even when their
   * future opportunity count is zero.
   **************************************************************/
  currentWorkers.forEach(worker => {
    if (futureCount[worker] === undefined) {
      futureCount[worker] = 0;
    }
  });

  return futureCount;
}


/**************************************************************
 * getDropdownOptions()
 **************************************************************/

function getDropdownOptions(cell) {
  const rule = cell.getDataValidation();
  if (!rule) return [];

  const criteria = rule.getCriteriaType();

  if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return rule.getCriteriaValues()[0];
  }

  if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return rule.getCriteriaValues()[0]
      .getValues()
      .flat()
      .filter(v => v !== '');
  }

  return [];
}


/**************************************************************
 * isBlockingTask()
 **************************************************************/

function isBlockingTask(name) {
  return !NON_BLOCKING_TASKS.includes(name);
}


/**************************************************************
 * addRole()
 **************************************************************/

function addRole(row, blocking, name, worker) {
  const isBlocking = isBlockingTask(name);

  row.occupiedRoles.push({
    columnName: name,
    worker: worker,
    blocking: isBlocking
  });

  if (isBlocking) {
    blocking.add(worker);
  }
}
