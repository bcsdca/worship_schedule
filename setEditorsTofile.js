function setEditorsTofile() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());

  // Loop through each email and add as editor
  // editorEmailAddresses is defined in the global variable file "globalVarWorshipSchedule"
  editorEmailAddresses.forEach(email => {
    file.addEditor(email);
  });
  
  logMessage(getCallStackTrace() + ": Added editors via DriveApp: " + editorEmailAddresses.join(", "));
}
