function removeUnavailableDatesTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach((trigger, index) => {
    logMessage(
      `${index + 1}. Function: ${trigger.getHandlerFunction()} | Event Type: ${trigger.getEventType()}`
    );
  });

  const functionName = 'collectUnavailableDatesResponse';

  var found = false;

  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      logMessage(getCallStackTrace() + `Found trigger: ${triggers[i].getHandlerFunction()} | Event Type: ${triggers[i].getEventType()} and deleted it`)
      ScriptApp.deleteTrigger(triggers[i]);
      found = true;
    }
  }

  const ui = SpreadsheetApp.getUi();
  if (found) {
    ui.alert('unavailable sheet monitoring was found ENABLED, we just disable it.');
  } else {
    ui.alert('unavailable sheet monitoring was found DISABLED, we don\'t do nothing.');
  }
 
  //running this function here is signify the start of the working on the new schedule for the new quarter !!!
  //probably the last week of the old quarter
  //update all the dropdowns upon the conclusion of updating the "unavailable dates" tab possibly a week before the new quarter
  updateAllDropDowns()

  //update the sheet permission to view only for everybody upon the conclusion of updating the "unavailable dates" tab possibly a week before the new quarter
  setAnyoneWithLinkToViewer();

  //update all the people that are allowed to edit this file, and they will receive an email notification for this.
  setEditorsTofile();

  //update all the people that are allowed to view this file, and they will receive an email notification for this.
  setViewersTofile();

  //remove the trigger for running the UpdateAllDropDowns function everyday at 12:30am
  toggleTriggerUpdateAllDropDowns("disable")
}