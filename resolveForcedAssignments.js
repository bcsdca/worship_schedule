/**********************************************************************
 * resolveForcedAssignments.gs
 *
 * Resolve ALL forced assignments before building AI schedule data.
 *
 * Pass 1:
 *   - rule
 *   - availability
 *
 * Pass 2:
 *   - single-option
 *
 * Does NOT write assignments to the worksheet.
 **********************************************************************/

function resolveForcedAssignments(sheet) {

  if (!sheet) throw new Error(`Sheet not found`);

  const sheetName = sheet.getName();
  aiLog(`Resolving forced assignments from ${sheetName}`);

  const headerInfo = buildHeaderMap(sheet);
  const headerMap = headerInfo.headerMap;

  const forcedAssignments = [];

  for (let r = 2; r <= sheet.getLastRow(); r++) {
    const date = sheet.getRange(`A${r}`).getValue();
    if (!(date instanceof Date)) continue;

    const dateDisplay = sheet.getRange(`A${r}`).getDisplayValue();
    const weekOfMonth = getWeekOfMonth(date);
    const context = {};
    const blocking = new Set();
    const forcedTaskNames = new Set();

    // Existing assignments
    CONTEXT_COLUMN_NAMES.forEach(name => {
      const col = headerMap[name];
      if (!col) return;

      const worker = sheet.getRange(`${col}${r}`).getValue();
      if (!worker) return;

      context[name] = worker;

      if (isBlockingTask(name)) blocking.add(worker);
    });

    // Log all current blocking workers
    aiLog(
      `${dateDisplay}: blocking workers after existing assignments = ` +
      JSON.stringify([...blocking])
    );

    /**************************************************************
     * Pass 1 - Resolve rule and availability forced assignments
     **************************************************************/
    TASK_COLUMN_NAMES.forEach(name => {
      const col = headerMap[name];
      if (!col) return;

      const cell = sheet.getRange(`${col}${r}`);
      if (cell.getValue()) return;

      let options = getDropdownOptions(cell).map(worker => ({ worker }));
      let worker = getForcedWorker(
        dateDisplay,
        name,
        context,
        options,
        blocking
      );

      let source = "rule";

      if (!worker) {
        options = applyAvailabilityRules(name, weekOfMonth, options);

        worker = getAvailabilityForcedWorker(
          name,
          weekOfMonth,
          options,
          blocking
        );

        source = "availability";
      }

      if (!worker) return;

      forcedAssignments.push({
        date: dateDisplay,
        columnName: name,
        worker: worker,
        source: source,
        _rowNum: r
      });

      forcedTaskNames.add(name);

      if (isBlockingTask(name)) blocking.add(worker);

      aiLog(`${source}: ${dateDisplay} | ${name} -> ${worker}`);
    });

    /**************************************************************
     * Pass 2 - Resolve cascading single-option assignments
     *
     * A single-option assignment may block a worker and cause
     * another task on the same Sunday to become single-option.
     *
     * Repeat until a complete pass produces no new assignments.
     **************************************************************/
    let singleOptionPassCount = 0;
    let singleOptionAssignmentCount = 0;
    let noValidOptions = [];
    let changed;

    do {
      changed = false;
      noValidOptions = [];
      singleOptionPassCount++;

      TASK_COLUMN_NAMES.forEach(name => {
        const col = headerMap[name];
        if (!col) return;

        const cell = sheet.getRange(`${col}${r}`);

        // Already assigned in worksheet
        if (cell.getValue()) return;

        // Already resolved during Pass 1 or an earlier loop
        if (forcedTaskNames.has(name)) return;

        let options = getDropdownOptions(cell)
          .map(worker => ({ worker }));

        // Apply availability restrictions
        options = applyAvailabilityRules(
          name,
          weekOfMonth,
          options
        );

        // Remove workers already holding a blocking role
        if (isBlockingTask(name)) {
          options = options.filter(
            option => !blocking.has(option.worker)
          );
        }

        // Remember tasks with no valid options
        if (!options.length) {
          noValidOptions.push(name);
          return;
        }

        // Multiple choices remain — leave for normal assignment
        if (options.length !== 1) return;

        const worker = options[0].worker;

        forcedAssignments.push({
          date: dateDisplay,
          columnName: name,
          worker: worker,
          source: "single-option",
          _rowNum: r
        });

        forcedTaskNames.add(name);

        // Immediately update blocking state
        if (isBlockingTask(name)) {
          blocking.add(worker);
        }

        singleOptionAssignmentCount++;
        changed = true;

        aiLog(
          `single-option: ${dateDisplay} | ${name} -> ${worker}`
        );
      });

    } while (changed);

    /**************************************************************
     * Log final single-option resolution results
     **************************************************************/
    aiLog(
      `${dateDisplay}: single-option resolution completed after ` +
      `${singleOptionPassCount} pass(es), ` +
      `${singleOptionAssignmentCount} assignment(s)`
    );

    noValidOptions.forEach(name => {
      aiLog(
        `No valid options: ${dateDisplay} | ${name}`
      );
    });
  }

  aiLog(
    `Forced assignments resolved: ${forcedAssignments.length}`
  );

  logLargeText(
    "=== FORCED ASSIGNMENTS ===",
    JSON.stringify(forcedAssignments, null, 2)
  );

  return forcedAssignments;
}

function getForcedWorker(date, name, context, options, blockingWorkers) {
  let worker = null;

  FORCED_RULES.forEach(rule => {
    if (rule.taskName !== name) return;

    if (context[rule.contextColumn] !== rule.contextValue) {
      aiLog(
        `${date} | ${name}: context mismatch ` +
        `(${rule.contextColumn}="${context[rule.contextColumn]}", ` +
        `expected "${rule.contextValue}")`
      );
      return;
    }

    if (!options.some(o => o.worker === rule.forcedWorker)) {
      aiLog(
        `${date} | ${name}: forced worker unavailable ` +
        `(${rule.forcedWorker})`
      );
      return;
    }

    if (
      blockingWorkers.has(rule.forcedWorker) &&
      isBlockingTask(name)
    ) {
      aiLog(
        `${date} | ${name}: forced worker blocked ` +
        `(${rule.forcedWorker})`
      );
      return;
    }

    worker = rule.forcedWorker;
  });

  return worker;
}

function applyAvailabilityRules(name, week, options) {
  AVAILABILITY_RESTRICTIONS.forEach(rule => {
    if (rule.taskName !== name) return;
    if (rule.allowedWeeks.includes(week)) return;

    options = options.filter(
      o => o.worker !== rule.workerName
    );
  });

  return options;
}

function getAvailabilityForcedWorker(name, week, options, occupied) {
  let worker = null;

  AVAILABILITY_RESTRICTIONS.forEach(rule => {
    if (rule.taskName !== name) return;
    if (!rule.forceIfEligible) return;
    if (!rule.allowedWeeks.includes(week)) return;
    if (!options.some(o => o.worker === rule.workerName)) return;

    if (
      occupied.has(rule.workerName) &&
      isBlockingTask(name)
    ) {
      aiLog(
        `${name} | ${rule.workerName}: availability worker blocked`
      );
      return;
    }

    worker = rule.workerName;
  });

  return worker;
}

