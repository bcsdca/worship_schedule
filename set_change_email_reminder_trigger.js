function set_change_email_reminder_trigger(option) {
  clearLogSheet();

  const triggers = ScriptApp.getProjectTriggers();

  logMessage(getCallStackTrace() + ": Checking existing triggers...");

  let found = false;

  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();

    logMessage(getCallStackTrace() + ": Found trigger -> " + handler);

    if (handler === "run_change_email_reminder") {
      found = true;

      ScriptApp.deleteTrigger(trigger);

      logMessage(
        getCallStackTrace() +
        ": Deleted old trigger -> run_change_email_reminder"
      );
    }
  });

  if (option === "enable") {

    ScriptApp.newTrigger("run_change_email_reminder")
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    logMessage(
      getCallStackTrace() +
      ": Created new change email reminder trigger."
    );

    SpreadsheetApp.getActive().toast(
      "The new a/v schedule change email reminder trigger was just created 👍"
    );

  } else {

    if (found) {
      logMessage(
        getCallStackTrace() +
        ": Existing change email reminder trigger removed."
      );
    } else {
      logMessage(
        getCallStackTrace() +
        ": No existing change email reminder trigger found."
      );
    }

    SpreadsheetApp.getActive().toast(
      "The a/v schedule change email reminder trigger was removed 👍"
    );
  }

  flushLogsToSheet();
}
