function getRowColStartEnd(sheet) {
  var returnArray = [];
  var headerRow = [];
  
  // Get the last row and column in the sheet
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  var dataRowStart = -1;
  var dataRowEnd = -1;
  var dataColumnStart = -1;
  var dataColumnEnd = lastColumn;

  var dataRowStartFound = false;

  for (var i = 1; i <= lastRow; i++) {
    var firstColumnValue = sheet.getRange(i, 1).getValue();
    
    if (firstColumnValue === "Date") {
      // Identify the header row and find the starting column for data
      for (var j = 1; j <= lastColumn; j++) {
        var columnValue = sheet.getRange(i, j).getValue();
        headerRow.push(columnValue);
        
        if (columnValue === "Date" && dataColumnStart === -1) {
          dataColumnStart = j;
        }
      }
    }
    else if (!dataRowStartFound && firstColumnValue instanceof Date) {
      // The first row where the first column contains a date marks the start of data
      dataRowStart = i;
      dataRowStartFound = true;
    }
    else if (dataRowStartFound && (!(firstColumnValue instanceof Date) || i === lastRow)) {
      // Identify the end of the data rows
      if (firstColumnValue instanceof Date && i === lastRow) {
        dataRowEnd = i;
      } else {
        dataRowEnd = i - 1;
      }
      break;
    }
  }

  // Log details and return the result
  logMessage(getCallStackTrace() + `: dataRowStart = ${dataRowStart}, dataRowEnd = ${dataRowEnd}, dataColumnStart = ${dataColumnStart}, dataColumnEnd = ${dataColumnEnd}, for sheet "${sheet.getName()}"`);
  
  returnArray.push(headerRow);
  returnArray.push([dataRowStart, dataRowEnd, dataColumnStart, dataColumnEnd]);
  
  //console.log(getCallStackTrace() + `: return array for sheet "${sheet.getName()}" : ${JSON.stringify(returnArray)}`);
  return returnArray;
}
