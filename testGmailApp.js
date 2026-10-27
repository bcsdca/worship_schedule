function testGmailApp() {
  const recipient = Session.getActiveUser().getEmail(); // Sends to yourself
  const subject = "Google Apps Script Gmail Test";
  const body =
    "This is a test email sent by GmailApp.\n\n" +
    "Time: " + new Date().toString() + "\n\n" +
    "If you received this email, GmailApp is working correctly.";

  GmailApp.sendEmail(recipient, subject, body);

  Logger.log("Test email sent to: " + recipient);
  SpreadsheetApp.getUi().alert("Test email sent to: " + recipient);
}