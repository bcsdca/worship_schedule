function createUnavailableDatesTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const functionName = 'collectUnavailableDatesResponse';

  // Check if it already exists
  const alreadyExists = triggers.some(trigger => trigger.getHandlerFunction() === functionName);

  if (alreadyExists) {
    SpreadsheetApp.getUi().alert(': Monitoring is already enabled.');
    return;
  }

  ScriptApp.newTrigger(functionName)
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(': Monitoring has been enabled.');

  setAnyoneWithLinkToEditor();
}