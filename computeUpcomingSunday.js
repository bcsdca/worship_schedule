// Function to compute the upcoming Sunday and return it in "MM/DD/YYYY" format
function computeUpcomingSunday(today) {
  const daysUntilSunday = (7 - today.getDay()) % 7;
  const upcomingSunday = new Date(today);
  upcomingSunday.setDate(today.getDate() + daysUntilSunday);
  const upcomingSundayStr = Utilities.formatDate(new Date(upcomingSunday), Session.getScriptTimeZone(), 'MM/dd/yyyy');
    
  logMessage(getCallStackTrace() + `: Upcoming sunday = ${upcomingSundayStr}.`);
  return upcomingSundayStr;
  
}

