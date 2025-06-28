function removeUnavailableDatesTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const functionName = 'collectUnavailableDatesResponse';

  let found = false;

  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
      found = true;
    }
  }

  const ui = SpreadsheetApp.getUi();
  if (found) {
    ui.alert('Monitoring has been disabled.');
  } else {
    ui.alert('Monitoring was not enabled.');
  }

  //running this function here is signify the start of the working on the new schedule for the new quarter !!!
  //probably the last week of the old quarter
  //update all the dropdowns upon the conclusion of updating the "unavailable dates" tab possibly a week before the new quarter
  updateAllDropDowns()
  
  //update the sheet permission to view only for everybody upon the conclusion of updating the "unavailable dates" tab possibly a week before the new quarter
  setAnyoneWithLinkToViewer();
  
  //update all the people that are allowed to edit this file, before the schedule go prime time
  setEditorsTofile();
}