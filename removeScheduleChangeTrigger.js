function removeScheduleChangeTrigger() {
  //var option = "disable";
  //toggleScheduleChangeTrigger(option)
  //clearLogSheet();

  const triggers = ScriptApp.getProjectTriggers();

  logMessage(getCallStackTrace() + ": Checking existing triggers...");

  let found = false;

  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();

    logMessage(getCallStackTrace() + ": Found trigger -> " + handler);

    if (handler === "worshipScheduleChangeReminder") {
      found = true;

      ScriptApp.deleteTrigger(trigger);

      logMessage(
        getCallStackTrace() +
        ": Deleted old trigger -> worshipScheduleChangeReminder"
      );
    }
  });
}
