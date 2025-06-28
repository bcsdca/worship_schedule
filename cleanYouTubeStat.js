function cleanYouTubeStat() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("YouTube Stat")
  var lastRow = sheet.getLastRow()
  var lastColumn = sheet.getLastColumn();
  console.log("lastRow is %d, lastColumn is %d", lastRow, lastColumn);

  sheet.getRange(2, 1, lastRow, lastColumn).clearContent();

  SpreadsheetApp.getActive().toast("Done, Clean up all the old previous YouTube stats 👍 !!!");

}

