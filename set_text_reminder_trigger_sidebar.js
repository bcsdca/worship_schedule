function set_text_reminder_trigger_sidebar(d) {

  //var text_trigger_day = form.Day.toString();
  var text_trigger_day = d;
  Logger.log("The new text trigger day is set to \"%s\" !!!", text_trigger_day);
  // remove all old text reminder triggers
  let remove_array = []
  var oldTrigger = ScriptApp.getScriptTriggers()
  Logger.log("The above triggers are the current running triggers !!!");
  //Logger.log(oldTrigger.length);
  for (var i = 0; i < oldTrigger.length; i++) {
    Logger.log("Current running trigger is " + ScriptApp.getScriptTriggers()[i].getHandlerFunction());
    if (ScriptApp.getScriptTriggers()[i].getHandlerFunction() == "run_text_reminder") {
      remove_array.push(oldTrigger[i]);
    }
  }

  remove_array.forEach(function (row) {
    ScriptApp.deleteTrigger(row);
    Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH:mm:ss') + ': Deleting text reminder trigger ' + row + ' !!!');
  });

  switch (text_trigger_day) {
    case "MONDAY":
      ScriptApp.newTrigger('run_text_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
      Logger.log("Starting new trigger run_text_reminder on " + text_trigger_day + " !!!");
      break;
    case "TUESDAY":
      ScriptApp.newTrigger('run_text_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(9).create();
      Logger.log("Starting new trigger run_text_reminder on " + text_trigger_day + " !!!");
      break;
    case "WEDNESDAY":
      ScriptApp.newTrigger('run_text_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(9).create();
      Logger.log("Starting new trigger run_text_reminder on " + text_trigger_day + " !!!");
      break;
    case "THURSDAY":
      ScriptApp.newTrigger('run_text_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(9).create();
      Logger.log("Starting new trigger run_text_reminder on " + text_trigger_day + " !!!");
      break;
    case "FRIDAY":
      ScriptApp.newTrigger('run_text_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).create();
      Logger.log("Starting new trigger run_text_reminder on " + text_trigger_day + " !!!");
      break;
    case "SATURDAY":
      ScriptApp.newTrigger('run_text_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(9).create();
      break;
    default:
      ScriptApp.newTrigger('run_text_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(9).create();
      Logger.log("Don't recognize the day selected \"%s\", so starting the new trigger run_text_reminder on SATURDAY !!!", text_trigger_day);
      break;
  }
    
  //closeSidebar()
  console.log("%s was just set as the email reminder day",text_trigger_day)
  SpreadsheetApp.getActive().toast("was just set as the text reminder day 👍 !!!", text_trigger_day);
  
}
