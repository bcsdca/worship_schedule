function buildScheduleData(sheet) {
  
  if (!sheet) throw new Error(`sheet not found`);

  aiLog(`Using sheet: ${sheet.getName()}`);
    
  const headerInfo = buildHeaderMap(sheet); 
  const headerMap = headerInfo.headerMap;

  const bounds = getScheduleRowBounds(sheet);

  if (!bounds) {
    aiLog(`first row and last row for all the data rows were not found`);
    return;
  }
  
  const firstDataRow = bounds.firstRow;
  const lastDataRow = bounds.lastRow;

  aiLog(`firstDataRow = ${firstRow}, and lastDataRow = ${lastDataRow}`);
  
  const rows = [];
  const forcedAssignments = [];

  for (let r = firstDataRow; r <= lastDataRow; r++) {
    const rawDate = sheet.getRange(r, 1).getValue();
    if (!(rawDate instanceof Date)) continue;

    const weekOfMonth = getWeekOfMonthObject(rawDate);
    const dateVal = sheet.getRange(r, 1).getDisplayValue();

    // Collect context values for this row
    const contextValues = {};
    CONTEXT_COLUMN_NAMES.forEach(name => {
      const colLetter = headerMap[name];
      if (!colLetter) return;
      const val = sheet.getRange(colLetter + r).getValue();
      if (val) contextValues[name] = val;
    });

    // Collect all names used this row (context + forced so far) for double-booking check
    const usedNamesThisRow = Object.values(contextValues).slice();

    const rowObj = {
      date: dateVal,
      weekOfMonth: weekOfMonth,
      context: contextValues,
      emptyTasks: [],
      forcedTasks: [],
      _rowNum: r  // internal only — used for sheet write-back, never sent to Claude
    };

    TASK_COLUMN_NAMES.forEach(name => {
      const colLetter = headerMap[name];
      if (!colLetter) {
        throw new Error('Column header not found: "' + name + '". Check TASK_COLUMN_NAMES matches row 1 exactly.');
      }
      const cell = sheet.getRange(colLetter + r);
      if (cell.getValue()) return; // already filled

      let options = getDropdownOptions(cell);

      // ── Step 1: Apply FORCED_RULES first (Priority 1) ───────────────────
      let forcedByRule = null;
      FORCED_RULES.forEach(rule => {
        if (rule.taskName !== name) return;
        if (contextValues[rule.contextColumn] !== rule.contextValue) return;
        // Rule matches — check worker is in dropdown and not double-booked
        if (!options.includes(rule.forcedWorker)) {
          Logger.log('WARNING: Forced rule requires "' + rule.forcedWorker + '" for row ' + r +
            ' "' + name + '" but they are not in the dropdown. Skipping force.');
          return;
        }
        if (usedNamesThisRow.includes(rule.forcedWorker) && name !== 'Power Point Preparation') {
          Logger.log('WARNING: Forced rule requires "' + rule.forcedWorker + '" for row ' + r +
            ' "' + name + '" but they are already used this row (double-booking). Skipping force.');
          return;
        }
        forcedByRule = rule.forcedWorker;
      });

      if (forcedByRule) {
        rowObj.forcedTasks.push({ columnName: name, worker: forcedByRule });
        forcedAssignments.push({ date: dateVal, columnName: name, worker: forcedByRule });
        if (name !== 'Power Point Preparation') usedNamesThisRow.push(forcedByRule);
        return; // don't add to emptyTasks
      }

      // ── Step 2: Strip ineligible workers by week restriction ────────────
      AVAILABILITY_RESTRICTIONS.forEach(rule => {
        if (rule.taskName !== name) return;
        if (rule.allowedWeeks.includes(weekOfMonth)) return;
        options = options.filter(o => o !== rule.workerName);
      });

      // ── Step 3: Apply forceIfEligible for availability-restricted workers ─
      let forcedByAvailability = null;
      AVAILABILITY_RESTRICTIONS.forEach(rule => {
        if (rule.taskName !== name) return;
        if (!rule.forceIfEligible) return;
        if (!rule.allowedWeeks.includes(weekOfMonth)) return;
        if (!options.includes(rule.workerName)) return;
        if (usedNamesThisRow.includes(rule.workerName) && name !== 'Power Point Preparation') return;
        forcedByAvailability = rule.workerName;
      });

      if (forcedByAvailability) {
        rowObj.forcedTasks.push({ columnName: name, worker: forcedByAvailability });
        forcedAssignments.push({ date: dateVal, columnName: name, worker: forcedByAvailability });
        if (name !== 'Power Point Preparation') usedNamesThisRow.push(forcedByAvailability);
        return;
      }

      // ------------------------------------------------------------------
      // Build occupiedWorkers and occupiedRoles for AI convenience.
      // This is redundant information derived from context + forcedTasks,
      // but it makes the AI's double-booking check much easier.
      // ------------------------------------------------------------------

      const occupiedWorkerSet = new Set();
      const occupiedRoles = [];

      // Context roles
      Object.entries(rowObj.context).forEach(([role, worker]) => {

        if (!worker) return;

        occupiedWorkerSet.add(worker);

        occupiedRoles.push({
          role: role,
          worker: worker
        });

      });

      // Forced task assignments
      rowObj.forcedTasks.forEach(task => {

        if (!task.worker) return;

        occupiedWorkerSet.add(task.worker);

        occupiedRoles.push({
          role: task.columnName,
          worker: task.worker
        });

      });

      rowObj.occupiedWorkers = [...occupiedWorkerSet];
      rowObj.occupiedRoles = occupiedRoles;

      // ── Step 4: Remaining slot goes to AI ────────────────────────────
      rowObj.emptyTasks.push({ columnName: name, options: options });
    });

    if (rowObj.emptyTasks.length > 0 || rowObj.forcedTasks.length > 0) rows.push(rowObj);
  }

  return { scheduleRows: rows, forcedAssignments: forcedAssignments };
}

function getWeekOfMonthObject(date) {
  return Math.ceil(date.getDate() / 7);
}


/**
 * Build a header map from the worksheet.
 *
 * The header row is NOT assumed to be row 1.
 *
 * The first row where column A contains exactly "Date"
 * is treated as the true header row.
 *
 * Returns:
 * {
 *   headerMap: {
 *     "Worship Chairperson": "B",
 *     "Speaker": "C",
 *     ...
 *   },
 *   headerRow: 5
 * }
 */
function buildHeaderMap(sheet) {

  if (!sheet) throw new Error(`sheet not found`);

  aiLog(`Using sheet: ${sheet.getName()}`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    throw new Error("Sheet is empty.");
  }

  /**************************************************************
   * Find the true header row.
   *
   * Look for the first cell in column A whose value is "Date".
   **************************************************************/
  const columnA =
    sheet
      .getRange(1, 1, lastRow, 1)
      .getValues();

  let headerRow = null;

  for (let r = 0; r < columnA.length; r++) {

    const value = columnA[r][0];

    if (
      value !== null &&
      value !== undefined &&
      value.toString().trim().toLowerCase() === "date"
    ) {
      headerRow = r + 1;
      break;
    }
  }

  if (headerRow === null) {
    throw new Error(
      `Header row not found: column A does not contain "Date".`
    );
  }

  /**************************************************************
   * Read the actual header row.
   **************************************************************/
  const headerValues =
    sheet
      .getRange(headerRow, 1, 1, lastCol)
      .getValues()[0];

  const headerMap = {};

  headerValues.forEach((text, idx) => {

    if (
      text === null ||
      text === undefined ||
      text.toString().trim() === ""
    ) {
      return;
    }

    const headerName =
      text.toString().trim();

    headerMap[headerName] =
      columnIndexToLetter(idx + 1);
  });

  return {
    headerMap: headerMap,
    headerRow: headerRow
  };
}

function columnIndexToLetter(colIndex) {
  let letter = '';
  while (colIndex > 0) {
    const rem = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    colIndex = Math.floor((colIndex - 1) / 26);
  }
  return letter;
}

