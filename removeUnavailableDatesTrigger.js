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
 
}