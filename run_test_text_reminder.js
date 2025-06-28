function run_test_text_reminder() {
  
  const Modes = {
    testing: true,
    sendEmail: false,
    sendText: true
  };

  console.log("Running the text reminder in a testing mode !!!");
  worshipReminder(Modes);
}
