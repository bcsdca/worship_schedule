function scheduleQuarterlyTriggerCleanup() {
  const triggerDate = getLastSundayOfQuarterAt2PM();

  // Prevent duplicate triggers
  deleteTriggerByFunction('quarterlyTriggerCleanup');

  ScriptApp.newTrigger('quarterlyTriggerCleanup')
    .timeBased()
    .at(triggerDate)
    .create();

  logMessage(getCallStackTrace() + `: Trigger scheduled for: ${triggerDate}`);
}

function getLastSundayOfQuarterAt2PM() {
  const now = new Date();

  const year = now.getFullYear();
  const month = now.getMonth(); // 0–11

  // Determine current quarter end month (0-based)
  const quarterEndMonth = Math.floor(month / 3) * 3 + 2;

  // Get last day of the quarter
  let lastDay = new Date(year, quarterEndMonth + 1, 0);

  // Move back to the last Sunday
  const dayOfWeek = lastDay.getDay(); // 0 = Sunday
  lastDay.setDate(lastDay.getDate() - dayOfWeek);

  // Set time to 2:00 PM
  lastDay.setHours(14, 0, 0, 0);

  return lastDay;
}

function quarterlyTriggerCleanup() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  logMessage(getCallStackTrace() + `: All triggers removed at end of quarter.`);
}

function deleteTriggerByFunction(functionName) {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}