function getWeekOfMonth(dateStr) {
  // Parse the input date string (format: MM/DD/YYYY)
  var date = new Date(dateStr);

  const dayOfWeek = date.getDay(); // 0 for Sunday

  // Ensure this function is called only on Sundays
  if (dayOfWeek !== 0) {
    throw new Error(getCallStackTrace() + ": This function is designed to work only on Sundays.");
  }

  // Get the day of the month (1st, 2nd, 3rd, etc.)
  var dayOfMonth = date.getDate();

  // Calculate the week of the month (week starts on Sunday)
  var weekOfMonth = Math.ceil(dayOfMonth / 7);

  //console.log(getCallStackTrace() + `: date = ${date}`);
  //console.log(getCallStackTrace() + `: dayOfMonth = ${dayOfMonth}`);
  console.log(getCallStackTrace() + `: Returning weekOfMonth = ${weekOfMonth}`);

  return weekOfMonth;
}