function emailReminderTuesday() {
    const Modes = {
    testing: false,
    sendEmail: true,
    sendText: false
  };

  logMessage(getCallStackTrace() + `: Running the email reminder normal mode !!!`);
  worshipReminder(Modes);
}
