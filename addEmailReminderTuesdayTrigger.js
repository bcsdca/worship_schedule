function addEmailReminderTuesdayTrigger() {
  
  ScriptApp.newTrigger("emailReminderTuesday")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(11)
    .create();

  logMessage(
    getCallStackTrace() +
    ": Created new emailReminderTuesday trigger."
  );

  SpreadsheetApp.getActive().toast(
    "The new a/v emailReminderTuesday trigger was just created 👍")
}
