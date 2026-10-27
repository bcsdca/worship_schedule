function removeGetYouTubeStatsTrigger() {
  clearLogSheet();

  const triggers = ScriptApp.getProjectTriggers();

  logMessage(
    getCallStackTrace() +
    ": Checking existing triggers..."
  );

  let found = false;

  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();

    logMessage(
      getCallStackTrace() +
      ": Found trigger -> " + handler
    );

    if (handler === "getYouTubeStats") {
      found = true;

      ScriptApp.deleteTrigger(trigger);

      logMessage(
        getCallStackTrace() +
        ": Deleted trigger -> getYouTubeStats"
      );
    }
  });

  if (found) {
    SpreadsheetApp.getActive().toast(
      "Done removing getYouTubeStats trigger 👍"
    );

    logMessage(
      getCallStackTrace() +
      ": Successfully removed getYouTubeStats trigger."
    );
  } else {

    SpreadsheetApp.getActive().toast(
      "No getYouTubeStats trigger found ⚠️"
    );

    logMessage(
      getCallStackTrace() +
      ": No getYouTubeStats trigger found."
    );
  }

  flushLogsToSheet();
}
