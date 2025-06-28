function buildPivotTable() {
  SpreadsheetApp.getActive().toast("Updating Worship Dashboard... Please wait.", "Processing...", -1); // Show a persistent toast
  var src_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var dest_sheet = SpreadsheetApp.getActive().getSheetByName("Pivot Table")
  var src_lastRow = src_sheet.getLastRow()
  var src_lastColumn = src_sheet.getLastColumn();
  logMessage(`${getCallStackTrace()}: src sheet lastRow is ${src_lastRow}, src sheet lastColumn is ${src_lastColumn}`);
  var src_dataRowStartFound = false;
  //look beyond at least 1 row beyond the last row for the row count, in order to find the actual end of the "Date" row
  for (var i = 1; i <= src_lastRow + 1; i++) {
    var srcRange = src_sheet.getRange(i, 1);
    var srcColumn1Value = srcRange.getValue();
    if (srcColumn1Value == "Date") {
      var srcHeaderRow = src_sheet.getRange(i, 1, 1, src_lastColumn).getValues();
      //srcHeaderRow[0] start at index 0
      for (k = 0; k <= srcHeaderRow[0].length -1; k++) { 
        if (srcHeaderRow[0][k] == "Worship Chairperson") {
          var srcTaskColumnStart = k+1;
          break;
        }
      }
      var destHeaderRow = src_sheet.getRange(i, srcTaskColumnStart, 1, src_lastColumn - (srcTaskColumnStart-1)).getValues();
      var srcHeaderRowNum = i;
      logMessage(`${getCallStackTrace()}: src Task Column Start = ${srcTaskColumnStart}`);
      logMessage(`${getCallStackTrace()}: src header row number = ${srcHeaderRowNum}`);
      logMessage(`${getCallStackTrace()}: src header row = ${JSON.stringify(srcHeaderRow)}`)
      logMessage(`${getCallStackTrace()}: dest header row = ${JSON.stringify(destHeaderRow)}`)
      //removing the speaker column
      destHeaderRow[0].splice(1, 1);
      //adding the the "name" column in the front
      destHeaderRow[0].splice(0, 0, 'Name');
      logMessage(`${getCallStackTrace()}: Modify dest header row = ${JSON.stringify(destHeaderRow)}`)
    }
    else if ((!src_dataRowStartFound) && (srcColumn1Value instanceof Date)) {
      src_dataRowStart = i;
      src_dataRowStartFound = true;
      logMessage(`${getCallStackTrace()}: src sheet dataRowStart is ${src_dataRowStart}`)
    }
    else if ((src_dataRowStartFound) && (!(srcColumn1Value instanceof Date))) {
      //one row before
      src_dataRowEnd = i - 1;
      logMessage(`${getCallStackTrace()}: src sheet dataRowEnd is ${src_dataRowEnd}`);
      break;
    }
  }
  //buidling up the transpose table

  var destRow = 1;
  var destCol = 2;
  var numOfRows = src_dataRowEnd - src_dataRowStart + 1;
  dest_sheet.clear();
  dest_sheet.appendRow(destHeaderRow[0]);
  destRow++;

  var srcRow = src_dataRowStart;
  //task column in source sheet start @5
  var srcCol = srcTaskColumnStart;

  for (k = 1; k <= destHeaderRow[0].length; k++) {
    //getting the name column from the src sheet
    var srcWorkerName = src_sheet.getRange(srcRow, srcCol, numOfRows, 1).getValues();
    //getting the task column from the src sheet
    var srcTaskName = src_sheet.getRange(srcHeaderRowNum, srcCol++, 1, 1).getValue();
    if (srcTaskName == "Speaker") {
      continue;
    } else {
      dest_sheet.getRange(destRow, 1, numOfRows, 1).setValues(srcWorkerName);
      dest_sheet.getRange(destRow, destCol++, numOfRows, 1).setValue(srcTaskName);
      destRow = destRow + numOfRows;
      logMessage(`${getCallStackTrace()}: Filling in the Pivot Table's Task Name ${srcTaskName}`);
    }
  }

  //sorting the Pivot Table based on the name column
  var dest_lastRow = dest_sheet.getLastRow()
  var dest_lastColumn = dest_sheet.getLastColumn();
  //starting the row 2 and column 1, and the rest of the rows and columns
  //avoiding row 1, the header row
  var dest_range = dest_sheet.getRange(2,1,dest_lastRow - 1,dest_lastColumn);
  //sorting base on column1 of the dest_range
  dest_range.sort(1);

  SpreadsheetApp.getActive().toast("Done, Just updated the Dashboard stats 👍 !!!", "Success.", 3);
  
}

