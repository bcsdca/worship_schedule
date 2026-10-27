function test_worshipReminder() {
  
  const Modes = {
    testing: true,
    sendEmail: true,
    sendText: false
  };

  logMessage(getCallStackTrace() + ": Running worship reminder email in a testing mode !!!");
  worshipReminder(Modes);

} 