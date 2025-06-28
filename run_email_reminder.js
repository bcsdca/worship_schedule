function run_email_reminder() {
    const Modes = {
    testing: false,
    sendEmail: true,
    sendText: false
  };

  console.log("Running the email reminder normal mode !!!");
  worshipReminder(Modes);
}
