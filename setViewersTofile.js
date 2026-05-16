function setViewersTofile() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());

  // Loop through each email and add as editor
  // viewerEmailAddresses is defined in the global variable file "globalVarWorshipSchedule"
  viewerEmailAddresses.forEach(email => {
    file.addViewer(email);
  });

  logMessage(getCallStackTrace() + ": Added viewer email address via DriveApp: " + viewerEmailAddresses.join(", "));
}
