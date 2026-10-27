function test_emailReminderTuesday() {
  
  const Modes = {
    testing: true,
    sendEmail: true,
    sendText: false
  };

  logMessage(getCallStackTrace() + `: Running the email reminder in a testing mode !!!`);
  worshipReminder(Modes);

} 
