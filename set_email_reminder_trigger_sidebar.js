function set_email_reminder_trigger_sidebar(d) {

  //var email_trigger_day = [form.Day].toString();
  var email_trigger_day = d;
  //SpreadsheetApp.getActiveSheet().appendRow(email_trigger_day)
  Logger.log("The new email trigger day is set to \"%s\" !!!", email_trigger_day);
  // removing all old email reminder triggers 1st
  let remove_array = []
  var oldTrigger = ScriptApp.getScriptTriggers()
  Logger.log("The above triggers are the current running triggers !!!");
  //Logger.log(oldTrigger.length);
  for (var i = 0; i < oldTrigger.length; i++) {
    Logger.log(ScriptApp.getScriptTriggers()[i].getHandlerFunction());
    if (ScriptApp.getScriptTriggers()[i].getHandlerFunction() == "run_email_reminder") {
      remove_array.push(oldTrigger[i]);

    }
  }
  remove_array.forEach(function (row) {
    //Logger.log(row);
    ScriptApp.deleteTrigger(row);
    Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH:mm:ss') + ': Deleting email reminder trigger ' + row + ' !!!');

  });

  switch (email_trigger_day) {
    case "MONDAY":
      ScriptApp.newTrigger('run_email_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(11).create();
      Logger.log("Starting new trigger run_email_reminder on " + email_trigger_day + " !!!");
      break;
    case "TUESDAY":
      ScriptApp.newTrigger('run_email_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(11).create();
      Logger.log("Starting new trigger run_email_reminder on " + email_trigger_day + " !!!");
      break;
    case "WEDNESDAY":
      ScriptApp.newTrigger('run_email_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(11).create();
      Logger.log("Starting new trigger run_email_reminder on " + email_trigger_day + " !!!");
      break;
    case "THURSDAY":
      ScriptApp.newTrigger('run_email_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(11).create();
      Logger.log("Starting new trigger run_email_reminder on " + email_trigger_day + " !!!");
      break;
    case "FRIDAY":
      ScriptApp.newTrigger('run_email_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(11).create();
      Logger.log("Starting new trigger run_email_reminder on " + email_trigger_day + " !!!");
      break;
    case "SATURDAY":
      ScriptApp.newTrigger('run_email_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(11).create();
      Logger.log("Starting new trigger run_email_reminder on " + email_trigger_day + " !!!");
      break;
    default:
      ScriptApp.newTrigger('run_email_reminder').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(11).create();
      Logger.log("Don't recognize the day selected \"%s\", so starting the new trigger run_email_reminder on TUESDAY !!!", email_trigger_day);
      break;
  }

  //closeSidebar()
  
  console.log("%s was just set as the email reminder day",email_trigger_day)
  SpreadsheetApp.getActive().toast("was just set as the email reminder day 👍 !!!", email_trigger_day);
}
