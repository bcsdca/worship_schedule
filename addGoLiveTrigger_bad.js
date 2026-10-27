/*function addGoLiveTrigger_bad() {
  const triggerDate = getFirstMondayOfQuarterAt11AM();

  // Prevent duplicate triggers
  deleteTriggerByFunction('goLive');

  ScriptApp.newTrigger('goLive')
    .timeBased()
    .at(triggerDate)
    .create();

  logMessage(
    getCallStackTrace() +
    `: goLive trigger scheduled for: ${triggerDate}`
  );
}
*/
/**
 * Returns the first Monday of the current or next quarter at 11:00 AM.
 */
/*function getFirstMondayOfQuarterAt11AM() {
  const now = new Date();

  const currentQuarter = Math.floor(now.getMonth() / 3);

  // First month of the next quarter
  let year = now.getFullYear();
  let quarterStartMonth = (currentQuarter + 1) * 3;

  if (quarterStartMonth >= 12) {
    quarterStartMonth = 0;
    year++;
  }

  // First day of next quarter
  const firstDay = new Date(year, quarterStartMonth, 1);

  // Find first Monday
  const dayOfWeek = firstDay.getDay(); // Sunday=0, Monday=1
  const daysToMonday = (8 - dayOfWeek) % 7;

  firstDay.setDate(firstDay.getDate() + daysToMonday);

  // Set 11:00 AM
  firstDay.setHours(11, 0, 0, 0);
  logMessage(getCallStackTrace() +`: First Monday of the coming quarter = ${firstDay}`);

  return firstDay;
}
*/