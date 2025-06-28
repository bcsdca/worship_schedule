function setAnyoneWithLinkToViewer() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());

  // Set general access to "anyone with the link"
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  logMessage(getCallStackTrace() + ': General access changed to "Viewer" for anyone with the link.');
}