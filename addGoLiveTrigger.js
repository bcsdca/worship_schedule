function addGoLiveTrigger() {
  const triggerDate = getMondayBeforeFirstSundayOfNextQuarterAt11AM();

  // Prevent duplicate triggers
  deleteTriggerByFunction('goLive');
  
  // goLive function is resided in onOpen
  ScriptApp.newTrigger('goLive')
    .timeBased()
    .at(triggerDate)
    .create();

  logMessage(
    getCallStackTrace() +
    `: goLive trigger scheduled for: ${triggerDate}`
  );
}

/**
 * Returns the Monday before the first Sunday of the next quarter at 11:00 AM.
 */
function getMondayBeforeFirstSundayOfNextQuarterAt11AM() {

  const now = new Date();

  const currentQuarter = Math.floor(now.getMonth() / 3);

  // First month of next quarter
  let year = now.getFullYear();
  let quarterStartMonth = (currentQuarter + 1) * 3;

  if (quarterStartMonth >= 12) {
    quarterStartMonth = 0;
    year++;
  }

  // First day of the next quarter
  let date = new Date(year, quarterStartMonth, 1);

  // Find the first Sunday
  const daysToSunday = (7 - date.getDay()) % 7;
  date.setDate(date.getDate() + daysToSunday);

  // Go back 6 days to Monday
  date.setDate(date.getDate() - 6);

  // Set to 11:00 AM
  date.setHours(15, 0, 0, 0);

  logMessage(getCallStackTrace() +`: Monday Before the 1st Sunday of the coming quarter = ${date}`);

  return date;
}