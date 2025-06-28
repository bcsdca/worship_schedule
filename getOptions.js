function getOptions() {
  var sheet = SpreadsheetApp.getActive().getSheets()[0];
  var returnArray = findRowColStartEnd(sheet);
  const dataRowStart = returnArray[1][0];
  const dataRowEnd = returnArray[1][1];
  const dataColumnStart = returnArray[1][2];
  const dataColumnEnd = returnArray[1][3];
  const numColumns = dataColumnEnd - dataColumnStart + 1;
  logMessage(getCallStackTrace() + ": dataRowStart = %d, dataRowEnd = %d, dataColumnStart = %d, dataColumnEnd = %d", dataRowStart, dataRowEnd, dataColumnStart, dataColumnEnd);

  //column 5 is the "Speaker" column
  const startColumn = 5;
  
  //return returnArray[0]
  //reduce function just reduce 2 dimension array to 1 dimension array
  //filter is the filter out any non-string array element
  return sheet.getRange(dataRowStart, startColumn, 1, numColumns - startColumn + 1).getDisplayValues()
    .filter(String)
    .reduce(function (a, b) {
      return a.concat(b)
    })
}
