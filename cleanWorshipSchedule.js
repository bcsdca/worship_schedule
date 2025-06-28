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

  console.log(getCallStackTrace() + ": src_dataRowStart = %d, src_dataRowEnd = %d, src_dataColumnStart = %d, src_dataColumnEnd = %d, total delete rows = %d, and total delete columns = %d ", src_dataRowStart, src_dataRowEnd, src_dataColumnStart, src_dataColumnEnd, total_delete_rows, total_delete_columns);
  
  //return;

  //clean up the all the data, except 1st row with "Date", and starting column 5 "Speaker" column
  src_sheet.getRange(src_dataRowStart + 1, src_dataColumnStart + 4, total_delete_rows, total_delete_columns).clearContent();

  console.log(getCallStackTrace() + ": \"Worship Schedule\" sheet clean up all data !!!");
       
  SpreadsheetApp.getActive().toast("Done, Clean up all the current \"Worship Schedule\" data 👍 !!!");

}
