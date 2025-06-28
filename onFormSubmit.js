// If someone submit a form without any date selection, it is allowed no date selection in the google form for this purpose
// Effectively, this person just clean up all the date selections before, and new data selection is added
// Esscentially, removeRecord function just removed all the date selections existed for that person in "Exception Date" tab
// addRecord don't have any date selection to add
function onFormSubmit(e) {
  //logMessage(JSON.stringify(e));
  clearLogSheet();
  logMessage(getCallStackTrace() + `: Google form transmit trigger event = ${JSON.stringify(e)}`)

  record_array = []
  //this is the id for the form that could be used for editing, that show the response and setting at the top
  var form = FormApp.openById('1L1M6IZyFFbH8ZUYUhfPXexZFedwoSb2Xn2D2vf6LTGw'); // Form ID
  var formResponses = form.getResponses();
  var formCount = formResponses.length;

  var formResponse = formResponses[formCount - 1];
  var itemResponses = formResponse.getItemResponses();

  for (var j = 0; j < itemResponses.length; j++) {
    var itemResponse = itemResponses[j];
    var title = itemResponse.getItem().getTitle();
    var answer = itemResponse.getResponse();
    record_array.push(answer);
    logMessage(getCallStackTrace() + `: Google Form Response${j} for title \"${title}\", is \"${answer}\"`)
  }
  
  //record_array[0] is name
  //record_array[1] is date
  //record_array[2] is spouseSelection "Yes" or "No"
  logMessage(getCallStackTrace() + `: record_array = ${JSON.stringify(record_array)}`)
  var tab_name = "Exception Dates";

  //fixed on 9/21/2024
  //to make sure name is not "", winki submit a form without name on 9/18/2024
  if (record_array[0] != "") {
    //removed the previous record for this co-worker, using record_array[0]
    removeRecord(record_array[0], record_array[1], tab_name);

    //then add this latest record for this co-workder
    addRecord(record_array[0], record_array[1], tab_name);

    // Check if spouseSelection is "Yes"
    if (record_array[2] === "Yes") {
      // Retrieve the spouse's name from a mapping object
      const spouseName = spouseMap[record_array[0]]; //spouseMap is a globally available
      if (spouseName) {
        // Perform the same operations for the spouse
        logMessage(`Handling spouse ${spouseName} for ${record_array[0]}`);
        removeRecord(spouseName, record_array[1], tab_name);
        addRecord(spouseName, record_array[1], tab_name);
      } else {
        logMessage(`No spouse found for ${record_array[0]} in spouseMap.`);
      }
    }

    //then rebuild all the pull down list for the whole worship schedule
    //this can inmprove in the future to just update the affected sunday only
    updateAllDropDowns();
  } else {
    logMessage(getCallStackTrace() + ": Not doing anything, because someone just submitted the form without a name = " + JSON.stringify(record_array))
  }
  flushLogsToSheet();
}
