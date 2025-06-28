function run_test_email_reminder() {
  
  const Modes = {
    testing: true,
    sendEmail: true,
    sendText: false
  };

  logMessage(getCallStackTrace() + ": Running the email reminder in a testing mode !!!");
  worshipReminder(Modes);

} 
