function createUnavailableDatesTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const functionName = 'collectUnavailableDatesResponse';

  // Check if it already exists
  const alreadyExists = triggers.some(trigger => trigger.getHandlerFunction() === functionName);

  if (alreadyExists) {
    SpreadsheetApp.getUi().alert('Unavailable sheet Monitoring is already enabled.');
    return;
  }

  ScriptApp.newTrigger(functionName)
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert('Unavailable sheet Monitoring has been enabled.');

  //update the sheet permission to editor for everybody to update the "unavailable dates" tab for the 1st time
  //this permission will get remove on this function "removeUnavailableDatesTrigger"
  setAnyoneWithLinkToEditor();

  //enable the trigger for running the UpdateAllDropDowns function everyday at 12:30am
  toggleTriggerUpdateAllDropDowns("enable")
  
}