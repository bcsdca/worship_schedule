function sendTextReminder(taskName, dutyContact, w_start_date, p_start_time, w_start_time, placeOfService, typeOfService, run_test) {

  const maps = mapContact();
  const email_collection = maps.emailMap;
  const phone_collection = maps.phoneMap;
  const wireless_carrier_collection = maps.carrierMap;

  let phone_number = phone_collection.get(dutyContact);
  let wireless_carrier = wireless_carrier_collection.get(dutyContact);
  let phone_number_no_dash = phone_number.replace(/-/g, "");

  let TextTo = phone_number_no_dash + "@" + wireless_carrier;
  let subject = "Cantonese Worship Task Reminder for this Sunday";
  let body;

  if (taskName === "Mandarin A/V helper") {
    body = `Dear ${dutyContact}, This is a friendly text reminder that you will be serving in Mandarin Worship Service for this coming Sunday(${w_start_date}) in the Fellowship Hall, as the ${taskName}, with the service start time at 9:15AM. Please arrive at Fellowship Hall at ** 8:45AM ** to setup a/v in fellowshp hall.`;
  } else if (taskName === "Power Point Preparation") {
    body = `Dear ${dutyContact}, This is a friendly text reminder that you will be serving in the ${typeOfService} Worship Service for this coming Sunday(${w_start_date}) in the ${placeOfService}, as the ${taskName}. Please submit your worship ppt to the Cantonese Worship account's google drive by 6:30pm of the Saturday night prior.`;
  } else {
    body = `Dear ${dutyContact}, This is a friendly text reminder that you will be serving in the ${typeOfService} Worship Service for this coming Sunday(${w_start_date}) in the ${placeOfService}, as the ${taskName}, with the service start time at ${w_start_time}. Please arrive at church at **${p_start_time}** to prepare our hearts to serve.`;
  }

  try {
    // Send the text
    if (run_test) {
      TextTo = "8587167471@tmomail.net";  // Testing
      logMessage(getCallStackTrace() + `: A reminder text was supposed to send to ${dutyContact} for ${taskName} at ${TextTo} !!!`);
      logMessage(getCallStackTrace() + ": But, re-directing to Bill Chu for testing!!!");
    } else {
      logMessage(getCallStackTrace() + `: Running in a normal mode, a reminder text will send to ${dutyContact} for ${taskName} at ${TextTo} !!!`);
    }
    GmailApp.sendEmail(TextTo, subject, body);

  } catch (e) {
    logMessage(getCallStackTrace() + "Error sending email to: " + dutyContact + " - " + e.message);
  }

}