function addGetYouTubeStatsTrigger() {
  ScriptApp.newTrigger('getYouTubeStats').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).nearMinute(50).create();
  logMessage(getCallStackTrace() + ": Creating getYouTubeStats trigger !!!")

  SpreadsheetApp.getActive().toast("Done, Creating youTube Stat collection trigger 👍 !!!");

}
