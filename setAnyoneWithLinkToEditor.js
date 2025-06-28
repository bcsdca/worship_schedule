function setAnyoneWithLinkToEditor() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());

  // Set general access to "anyone with the link" as Editor
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  logMessage(getCallStackTrace() + ': General access changed to "Editor" for anyone with the link.');
}
