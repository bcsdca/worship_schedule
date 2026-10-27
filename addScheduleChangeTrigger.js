function addScheduleChangeTrigger() {
  //var option = "enable";
  //toggleScheduleChangeTrigger(option)
  ScriptApp.newTrigger("worshipScheduleChangeReminder")
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    logMessage(
      getCallStackTrace() +
      ": Created new worshipScheduleChangeReminder."
    );

    SpreadsheetApp.getActive().toast(
      "The new a/v worshipScheduleChangeReminder was just created 👍")
}
