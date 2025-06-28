function set_change_email_reminder_trigger(option) {
  clearLogSheet();

  // removing all old change email reminder triggers 1st
  let remove_array = []
  var oldTrigger = ScriptApp.getScriptTriggers()
  //logMessage(oldTrigger.length);
  logMessage(getCallStackTrace() + ": The below triggers are the current running triggers !!!");
  for (var i = 0; i < oldTrigger.length; i++) {
    logMessage(getCallStackTrace() + ": " + ScriptApp.getScriptTriggers()[i].getHandlerFunction());
    if (ScriptApp.getScriptTriggers()[i].getHandlerFunction() == "run_change_email_reminder") {
      remove_array.push(oldTrigger[i]);

    }
  }
  remove_array.forEach(function (row) {
    //logMessage(row);
    ScriptApp.deleteTrigger(row);
    logMessage("set_change_email_reminder_trigger" + ": " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH:mm:ss') + ': Deleting the old change email reminder trigger ' + row + ' !!!');

  });

  if (option == "enable") {
    ScriptApp.newTrigger("run_change_email_reminder")
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    logMessage(getCallStackTrace() + ": The new a/v schedule change email reminder trigger was just created !!!",);
    SpreadsheetApp.getActive().toast("The new a/v schedule change email reminder trigger was just created 👍 !!!");
  } else {
    logMessage(getCallStackTrace() + ": The a/v schedule change email reminder trigger was just removed !!!",);
    SpreadsheetApp.getActive().toast("The a/v schedule change email reminder trigger was just removed 👍 !!!");
  }

  flushLogsToSheet();
  
  const sheet0 = SpreadsheetApp.getActive().getSheets()[0];
  SpreadsheetApp.getActive().setActiveSheet(sheet0); // keep the view at the 1st sheet
  //SpreadsheetApp.getActive().setActiveSheet(sheet); // keep the view at the worship
}
