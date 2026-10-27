function addEndOfQuarterCleanupTrigger() {
  const triggerDate = getLastSundayOfNextQuarterAt2PM();

  // Prevent duplicate triggers
  deleteTriggerByFunction('quarterlyCleanupTrigger');

  ScriptApp.newTrigger('quarterlyCleanupTrigger')
    .timeBased()
    .at(triggerDate)
    .create();

  logMessage(getCallStackTrace() + `: Trigger scheduled for: ${triggerDate}`);
}

function getLastSundayOfNextQuarterAt2PM() {
  const now = new Date();

  let year = now.getFullYear();
  const month = now.getMonth(); // 0–11

  // Determine next quarter's end month
  let quarterEndMonth = Math.floor(month / 3) * 3 + 5;

  // Handle next quarter being in the following year
  if (quarterEndMonth > 11) {
    quarterEndMonth -= 12;
    year++;
  }

  // Get last day of next quarter
  const lastDay = new Date(year, quarterEndMonth + 1, 0);

  // Move back to the last Sunday
  lastDay.setDate(lastDay.getDate() - lastDay.getDay());

  // Set time to 2:00 PM
  lastDay.setHours(14, 0, 0, 0);

  logMessage(
    getCallStackTrace() +
    `: Last Sunday of Next Quarter: ${lastDay}`
  );

  return lastDay;
}


function quarterlyCleanupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  logMessage(getCallStackTrace() + `: All triggers removed at end of quarter.`);
}