function run_forwardSermonInfo() {
  //run input parameter to forwardSermonInfo function is true
  //should only to everbody
  logMessage(getCallStackTrace() + ": Actual running mode, will forward email to everbody !!!");
  forwardSermonInfo(true)

}
