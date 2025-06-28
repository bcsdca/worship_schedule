function run_text_reminder() {
  const Modes = {
    testing: false,
    sendEmail: false,
    sendText: true
  };

  console.log("Running the text reminder in a normal mode !!!");
  worshipReminder(Modes);
}
