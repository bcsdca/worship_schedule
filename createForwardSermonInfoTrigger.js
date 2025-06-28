function createForwardSermonInfoTrigger() {
  //Create new trigger
  //check to make sure there is no createForwardSerminInfoTrigger already existed
  var find_it = false
  //logMessage(getCallStackTrace() + "got to the trigger function !!!");
  var oldTrigger = ScriptApp.getScriptTriggers()
  logMessage(getCallStackTrace() + ": The below triggers are the current running triggers !!!");
  //Logger.log(oldTrigger.length);
  for (var i = 0; i < oldTrigger.length; i++) {
    logMessage(getCallStackTrace() + ": Current running trigger is " + ScriptApp.getScriptTriggers()[i].getHandlerFunction());
    if (ScriptApp.getScriptTriggers()[i].getHandlerFunction() == "forwardSermonInfo") {
      find_it = true;
      break;
    }
  }

  if (!find_it) {
    ScriptApp.newTrigger('forwardSermonInfo').timeBased().everyMinutes(15).create();
    logMessage(getCallStackTrace() + ": " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH:mm:ss') + ': No existing forwardSermonInfo trigger, so creating forwardSermonInfo trigger!!!');
  } else {
    logMessage(getCallStackTrace() + ": " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH:mm:ss') + ': Found an existing forwardSermonInfo trigger, NOT creating forwardSermonInfo trigger !!!');
  }

  
}
