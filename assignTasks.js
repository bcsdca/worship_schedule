function assignTasks(sheet, taskName, col_id, workerQueue, workerCount, startRow, endRow, headerRow, runOption) {

  let previousAssignedWorker = null;

  logMessage(`${getCallStackTrace()}: Starting worker queue: ${JSON.stringify(workerQueue)} for Task name: ${taskName}, and auto assignment is running in ${runOption} mode`);

  for (let row = startRow; row <= endRow; row++) {
    const cellAddress = col_id + row;
    let combineAvailableQueue = getDropdownList(cellAddress);

    const dateAddress = "A" + row;
    const dateValue = sheet.getRange(dateAddress).getValue();
    const date = Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), 'MM/dd/yyyy');

    // Adjust available queue based on task dependencies
    combineAvailableQueue = adjustAvailableQueueForTask(taskName, combineAvailableQueue, row, sheet, headerRow);
    logMessage(`${getCallStackTrace()}: combine available worker queue after adjusting for special cases: ${JSON.stringify(combineAvailableQueue)} for Task name: ${taskName} on ${date}`);

    // Filter the available queue based on worker restrictions (availability per week)
    combineAvailableQueue = filterWorkersByAvailability(date, taskName, combineAvailableQueue, workerRestrictionsArray);
    logMessage(`${getCallStackTrace()}: combine available worker queue after filtering for availability: ${JSON.stringify(combineAvailableQueue)} for Task name: ${taskName} on ${date}`);

    // Sort workerQueue if running in balance mode
    if (runOption === "balance") {
      workerQueue.sort((a, b) => (workerCount.get(a) || 0) - (workerCount.get(b) || 0));
      logMessage(`${getCallStackTrace()}: row ${row}: running in ${runOption} mode. Sorted workerQueue: ${JSON.stringify(workerQueue)} based on workerCount: ${mapToString(workerCount)} on ${date}`);
    } else {
      logMessage(`${getCallStackTrace()}: row ${row}: running in ${runOption} mode. WorkerQueue: ${JSON.stringify(workerQueue)} on ${date}`);
    }

    // Assign worker from available queue
    const assignedWorker = assignWorker(taskName, workerQueue, combineAvailableQueue, previousAssignedWorker);

    if (assignedWorker) {
      sheet.getRange(cellAddress).setValue(assignedWorker);
      workerCount.set(assignedWorker, (workerCount.get(assignedWorker) || 0) + 1);
      previousAssignedWorker = assignedWorker;
      workerQueue = rotateQueue(workerQueue, assignedWorker);
      logMessage(`${getCallStackTrace()}: After assigning worker ${assignedWorker}, the updated workerQueue: ${JSON.stringify(workerQueue)}, combineAvailableQueue: ${JSON.stringify(combineAvailableQueue)}, previousAssignedWorker: ${previousAssignedWorker}, workerCount: ${mapToString(workerCount)}`);
    } else {
      logMessage(`No available worker found for row ${row}. Worker queue: ${JSON.stringify(workerQueue)}, combineAvailableQueue: ${JSON.stringify(combineAvailableQueue)}`);
      //sheet.getRange(cellAddress).setValue("NO BODY");
      sheet.getRange(cellAddress).setValue("");
    }
  }

  return [workerQueue, workerCount];
}

function filterWorkersByAvailability(date, taskName, availableQueue, workerRestrictionsArray) {
  // Get the current week of the month and day of the week
  const weekOfMonth = getWeekOfMonth(date); // Helper function to get the week of the month

  // Filter workers based on their restrictions for the current task and week
  const restrictedWorkers = availableQueue.filter(worker => {
    // Find the restriction for the current worker and task
    const workerRestriction = workerRestrictionsArray.find(
      w => w.workerName === worker && w.taskName === taskName
    );

    // Include workers only if they have a restriction and are available this week
    return workerRestriction && workerRestriction.availableWeeks.includes(weekOfMonth);
  });

  // If any restricted workers are available for this week, return only them
  if (restrictedWorkers.length > 0) {
    logMessage(`${getCallStackTrace()}: Found and return restrictedWorkers ${restrictedWorkers}. for taskName: ${taskName}, on date: ${date}, from the availableQueue: ${JSON.stringify(availableQueue)}`);
    return restrictedWorkers;
  }

  logMessage(`${getCallStackTrace()}: NO restrictedWorkers found for taskName: ${taskName}, on date: ${date}, for the availableQueue: ${JSON.stringify(availableQueue)}`);

  // Otherwise, return all workers without any restriction
  return availableQueue.filter(worker => {
    // Check if the worker has no restriction for the current task
    const workerRestriction = workerRestrictionsArray.find(
      w => w.workerName === worker && w.taskName === taskName
    );
    if (workerRestriction) {
      logMessage(`${getCallStackTrace()}: Remove worker: ${worker}, from availableQueue, with taskName: ${taskName} has restriction, but not on date: ${date}`);
    }
    if (!workerRestriction) {
      logMessage(`${getCallStackTrace()}: Include worker: ${worker}, from availableQueue, with taskName: ${taskName} has no restriction on date: ${date}`);
    }
    return !workerRestriction; // Include only if no restriction exists
  });

}

function adjustAvailableQueueForTask(taskName, availableQueue, row, sheet, headerRow) {
  if (taskDependencies[taskName]) {
    availableQueue = chkNameOnOtherColumn(sheet, taskName, taskDependencies[taskName], row, headerRow, availableQueue);
  }

  return availableQueue;
}

function assignWorker(taskName, workerQueue, availableQueue, previousAssignedWorker) {
  let assignedWorker = null;

  for (const worker of workerQueue) {
    if (availableQueue.length === 1 && availableQueue[0] === worker) {
      assignedWorker = worker;
      logMessage(`${getCallStackTrace()}: Only one worker ${availableQueue[0]} for ${taskName}, found in combineAvailableQueue: ${JSON.stringify(availableQueue)}, and is matching the targeted worker ${worker}, Assigning worker without choice.`);
      break;
    } else if (availableQueue.includes(worker) && previousAssignedWorker !== worker) {
      assignedWorker = worker;
      break;
    } else if (!availableQueue.includes(worker)) {
      //logMessage(`${getCallStackTrace()}: Worker ${worker} for ${taskName}, not found in combineAvailableQueue: ${JSON.stringify(availableQueue)}. Trying next worker.`);
    } else if (previousAssignedWorker === worker) {
      logMessage(`${getCallStackTrace()}: Worker ${worker} for ${taskName}, was assigned last week. Trying next worker.`);
    }
  }

  if (assignedWorker) {
    logMessage(`${getCallStackTrace()}: Assigning worker ${assignedWorker} for ${taskName},  with workerQueue: ${JSON.stringify(workerQueue)}, combineAvailableQueue: ${JSON.stringify(availableQueue)}, previousAssignedWorker: ${previousAssignedWorker}`);
  }

  return assignedWorker;
}

function rotateQueue(queue, worker) {
  const newQueue = queue.filter(w => w !== worker);
  newQueue.push(worker);
  return newQueue;
}

function mapToString(map) {
  return Array.from(map.entries()).map(([key, value]) => `${key}=${value}`).join(',');
}

function chkNameOnOtherColumn(sheet, taskName, otherColumnTaskName, row, headerRow, combineAvailableQueue) {
  const allWorkersOnRow = sheet.getRange(row, 5, 1, sheet.getLastColumn()).getDisplayValues().flat();
  logMessage(`${getCallStackTrace()}: All workers found on row ${row} = ${JSON.stringify(allWorkersOnRow)}, with the input combineAvailableQueue = ${JSON.stringify(combineAvailableQueue)}, and the otherColumnTaskName = ${otherColumnTaskName}`);

  for (const otherTask of otherColumnTaskName) {
    const otherTaskWorker = getWorkerName(sheet, row, otherTask, headerRow);
    if (otherTaskWorker != "") {
      logMessage(`${getCallStackTrace()}: Working on this item = otherTaskWorker is "${otherTaskWorker}" and is assigned to otherTask "${otherTask}"`);
    }
    
    // Handle special cases with forced assignments
    if (specialCases[otherTaskWorker] && specialCases[otherTaskWorker].task === otherTask && specialCases[otherTaskWorker].forceTask === taskName) {
      const { forceWorker } = specialCases[otherTaskWorker];
      
      logMessage(`${getCallStackTrace()}: Checking Special case... otherTaskWorker of "${otherTaskWorker}" as "${specialCases[otherTaskWorker].task}" is checking against "${forceWorker}" as "${taskName}"`);
      
      //check if this forceWorker like wellington is already working on this row or not
      //check if the taskName == "Power Point Preparation" or not?
      //if it is, special case is allowed, because the "Power Point Preparation" was done on Saturday
      //if also the forceWorker is available that week
      if (combineAvailableQueue.includes(forceWorker)) {
        if (!allWorkersOnRow.includes(forceWorker) || taskName == "Power Point Preparation") {
          combineAvailableQueue = [forceWorker];
          logMessage(`${getCallStackTrace()}: Special case detected. otherTaskWorker of "${otherTaskWorker}" as "${specialCases[otherTaskWorker].task}" is forcing "${forceWorker}" as "${taskName}"`);
          logMessage(`${getCallStackTrace()}: Special case detected. Forcing combineAvailableQueue to "${JSON.stringify(combineAvailableQueue)}"`);
          continue;
        }
      }

    }

    // Remove the worker from the queue if they are already assigned to another task
    if (otherTaskWorker) {
      const position = combineAvailableQueue.indexOf(otherTaskWorker);
      if (position !== -1) {
        combineAvailableQueue.splice(position, 1);
        logMessage(`${getCallStackTrace()}: The updated combineAvailableQueue is ${JSON.stringify(combineAvailableQueue)}, after removing worker "${otherTaskWorker}" from it, because he/she is already assigned to "${otherTask}"`);
      }
    } else {
      //logMessage(`${getCallStackTrace()}: No action taken for worker "${otherTaskWorker}" assigned to "${otherTask}"`);
    }
  }

  return combineAvailableQueue;
}
