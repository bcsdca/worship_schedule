function test_forwardSermonInfo() {
  //run input parameter to forwardSermonInfo function is false
  //should only to Bill Chu
  logMessage(getCallStackTrace() + ": test mode, will forward email to Bill Chu !!!");
  forwardSermonInfo(true)
}
