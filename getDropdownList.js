function getDropdownList(cellAddress) {
  var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  
  // Get the cell by address
  var cell = scheduleSheet.getRange(cellAddress);
 
  // Get the data validation rule for the cell
  var rule = cell.getDataValidation();
  
  var dropDownList = [];

  // Check if the cell has a data validation rule (dropdown)
  if (rule != null) {
    var criteria = rule.getCriteriaType();
    var args = rule.getCriteriaValues();

    // Log the dropdown options if criteria type is VALUE_IN_LIST
    if (criteria == SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      var dropDownList = args[0];
      logMessage(getCallStackTrace() + ": The value of the dropdown list in cell " + cellAddress + " is: " + JSON.stringify(dropDownList));
      return dropDownList;
    }
  } else {
    logMessage(getCallStackTrace() + "No dropdown options found for the cell " + cellAddress);
    return dropDownList;
  }
  
}

