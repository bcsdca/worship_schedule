function cleanWorshipSchedule() {

  var src_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  //var src_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  const rowColStartEnd = findRowColStartEnd(src_sheet);

  const src_dataRowStart = rowColStartEnd[1][0];
  const src_dataRowEnd = rowColStartEnd[1][1];
  const src_dataColumnStart = rowColStartEnd[1][2];
  const src_dataColumnEnd = rowColStartEnd[1][3];

  //figuring out how many rows to clean on the worship schedule
  var total_delete_rows = src_dataRowEnd - (src_dataRowStart + 1) + 1;
  //starting delete from column 5
  var total_delete_columns = src_dataColumnEnd - src_dataColumnStart + 1 - 5;

  logMessage(getCallStackTrace() + `: src_dataRowStart = ${src_dataRowStart}, src_dataRowEnd = ${src_dataRowEnd}, src_dataColumnStart = ${src_dataColumnStart}, src_dataColumnEnd = ${src_dataColumnEnd}, total delete rows = ${total_delete_rows}, and total delete columns = ${total_delete_columns}`);
 
  //clean up the all the data, except 1st row with "Date", and starting column 5 "Speaker" column
  src_sheet.getRange(src_dataRowStart + 1, src_dataColumnStart + 4, total_delete_rows, total_delete_columns).clearContent();
 
  const today = new Date();
  const currentMonth = today.getMonth(); // 0-11
  const currentQuarter = Math.floor(currentMonth / 3);

  // Determine the first month of the *next* quarter
  const nextQuarterStartMonth = ((currentQuarter + 1) % 4) * 3;
  const nextQuarterYear = currentQuarter === 3 ? today.getFullYear() + 1 : today.getFullYear();
  logMessage(getCallStackTrace() + `: nextQuarterStartMonth = ${nextQuarterStartMonth}, nextQuarterYear = ${nextQuarterYear}`);

  // First day of the next quarter
  const firstDay = new Date(nextQuarterYear, nextQuarterStartMonth, 1);

  // Find the first Sunday (0 = Sunday)
  const dayOfWeek = firstDay.getDay();
  const offset = (7 - dayOfWeek) % 7;
  const firstSunday = new Date(firstDay);
  firstSunday.setDate(firstDay.getDate() + offset);
  logMessage(getCallStackTrace() + `: firstSunday = ${firstSunday}`);

  // Format as M/d/yyyy
  const formatted = Utilities.formatDate(firstSunday, Session.getScriptTimeZone(), "M/d/yyyy");
  logMessage(getCallStackTrace() + `: The 1st sunday of the next quarter is ${formatted} !!!`);
  src_sheet.getRange(src_dataRowStart + 1, src_dataColumnStart, 1, 1).setValue(formatted);

  logMessage(getCallStackTrace() + ": \"Worship Schedule\" sheet clean up all data, and set the next quarter sundays !!!");

  SpreadsheetApp.getActive().toast("Done, Clean up all the current \"Worship Schedule\" data, and set the next quarter sundays 👍 !!!");

}
