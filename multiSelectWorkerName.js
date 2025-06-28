function multiSelectWorkerName() {
  // Open the sidebar when the user runs the function manually from a button in cell C8 in the qRMMTemplate tab
  // this function will select a worker name or names from the sheet0 and populate in B8 of the qRMMTemplate tab
  // to send a special email to ask them to review the tasks for the coming quarter
  const html = HtmlService.createHtmlOutputFromFile('htmlMultiSelectSideBar')
    .setTitle('Multi-Select for Names Input');
  SpreadsheetApp.getUi().showSidebar(html);
}

function populateCell(values) {
  // Get the active sheet and a target cell, populate with the selected values
  console.log(arguments.callee.name + ": The cell values user selected." + values);
  const templateSheet = SpreadsheetApp.getActive().getSheetByName("qRMMTemplate");
  const cell = templateSheet.getRange('B8');
  
  cell.setValue(values.join(", "));
}

// This function returns the checkbox options (you can modify this array)
// to return all the name to be selected in checkboxs
function getCheckboxOptions() {
  const excludedTasks = ["Speaker", "Usher/Welcome1", "Usher/Welcome2", "Usher/Welcome3"];
      
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const templateSheet = SpreadsheetApp.getActive().getSheetByName("qRMMTemplate");
    
  const rowColStartEnd = getRowColStartEnd(sheet);
  const headerRow = rowColStartEnd[0];
  const [dataRowStart, dataRowEnd, dataColumnStart, dataColumnEnd] = rowColStartEnd[1];

  const taskData = sheet.getRange(dataRowStart, dataColumnStart, dataRowEnd - dataRowStart + 1, dataColumnEnd - dataColumnStart + 1).getValues();

  const taskAssignments = processTaskAssignments(taskData, headerRow, excludedTasks);

  // Create an array to hold all names
  // Use a Set to remove duplicates
  const namesList = [...new Set(taskAssignments.map(({ name }) => name))];
  return namesList;
  //return ['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5']; // Example array
}

