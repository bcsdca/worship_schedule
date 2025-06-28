//find the row and column start and end on the variable sheet passing in
//return 2 dimension arrays
//[headerRow] 
//[dataRowStart, dataRowEnd, dataColumnStart, dataColumnEnd]

//dataRowStart > pointing at the row with the "Date" in column 1
//dataRowEnd > pointing at the last row with the date format on column 1
//dataColumnStart > pointing at the 1st column with the "Date"]
//dataColumnEnd > pointing at the last column]

function findRowColStartEnd(sheet) {
  var returnArray = [];
  var headerRow = [];
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  //var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  //var sheetName = sheet.getName();
  var lastRow = sheet.getLastRow();
  var dataColumnEnd = sheet.getLastColumn();
  //logMessage("lastRow is %d, dataColumnEnd is %d", lastRow, dataColumnEnd)
  var dataRowStartFound = false;
  for (var i = 1; i <= lastRow; i++) {
    var tmpRange = sheet.getRange(i, 1);
    var tmpColumn1Value = tmpRange.getValue()
    //logMessage("tmpColumn1Value is %d", tmpColumn1Value)
    if (tmpColumn1Value == "Date") {
      for (var j = 1; j <= dataColumnEnd; j++) {
        var tmpxRange = sheet.getRange(i, j);
        var tmpColumnxValue = tmpxRange.getValue()
        headerRow.push(tmpColumnxValue);
        //looking for the 1st data column start, which is the "Worship Chairperson" column
        if (tmpColumnxValue == "Date") {
          var dataColumnStart = j;
          //break;
        }
      }
      
    }
    //dataRowStart including the header row
    else if ((!dataRowStartFound) && (tmpColumn1Value instanceof Date)) {
      dataRowStart = i - 1;
      dataRowStartFound = true;
      //logMessage("dataRowStart is %d", dataRowStart)
    }
    else if ((dataRowStartFound) && (i == lastRow) && (tmpColumn1Value instanceof Date)) {
      //this is the last row and column1 is still the date format
      dataRowEnd = i;
      //logMessage("dataRowEnd is %d", dataRowEnd)
      break;
    }
    else if ((dataRowStartFound) && (!(tmpColumn1Value instanceof Date))) {
      //one row before
      dataRowEnd = i - 1;
      //logMessage("dataRowEnd is %d", dataRowEnd)
      break;
    }
  }

  logMessage(getCallStackTrace() + `: dataRowStart = ${dataRowStart}, dataRowEnd = ${dataRowEnd}, dataColumnStart = ${dataColumnStart}, dataColumnEnd = ${dataColumnEnd}, for sheet \"${sheet.getName()}\"`);
  returnArray.push(headerRow);
  returnArray.push([dataRowStart, dataRowEnd, dataColumnStart, dataColumnEnd]);
  logMessage(getCallStackTrace() + `: return array for sheet \"${sheet.getName()}\" : ${JSON.stringify(returnArray)}`);
  return returnArray;

}
