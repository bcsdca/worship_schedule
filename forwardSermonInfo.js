function forwardSermonInfo(isTesting = false) {
     
  clearLogSheet();

  // If isTesting is not a boolean (e.g., it's an event object because of the time trigger), force it to false
  if (typeof isTesting !== 'boolean') {
    isTesting = false;
    logMessage(`${getCallStackTrace()}: This function was called by a time trigger, so force isTesting = ${isTesting}`)
  }

  //const sermonQuery = "is:inbox newer_than:2d subject:(sermon -mandarin)";
  const sermonQuery = "is:inbox newer_than:4d subject:(sermon -mandarin)";
  const taskRequireEmailForward = ["Worship Chairperson", "Song Leader", "On-Stage Translator"];
  const praiseTeamMembers = ["Stephen Wong", "Jessica Wan", "Jason Tong", "Kenneth Kong", "Emily Lieu", "Josiah Lee"];
  const praiseTeamLeaders = [
    {
      role: 'Song Leader',
      name: ['Josiah Lee']
      //name: ['Josiah Lee','Edmond Chan']
    },
    {
      role: 'Pianist',
      name: ['Jessica Wan']
      //name: ['Anna Lau']
    }
  ];
  const mandatoryRecipents = ["Andy Chu", "Bill Chu"];
  var emailFound = false;
  const today = new Date();
  let threads = [];

  try {
    // Search for Emily's sermon info email
    threads = GmailApp.search(sermonQuery, 0, 1);
    if (threads.length === 0) {
      throw new Error("No threads found.");
    } else {
      emailFound = true;
      //logMessage(getCallStackTrace() + `: Worship info email subject: "${threads[0].getMessages()[0].getSubject()}"`);
    }
  } catch (err) {
    console.error(getCallStackTrace() + ": 'Error retrieving sermon email: '" + err.message);
  }

  if (emailFound) {
    logMessage(getCallStackTrace() + `:  Worship info email subject: "${threads[0].getMessages()[0].getSubject()}"`);

    const emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    logMessage(getCallStackTrace() + `: Remaining email quota before: ${emailQuotaRemaining}`);

    const maps = mapContact(); // Assuming mapContact returns a map of email addresses
    const emailMap = maps.emailMap;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const [headerRow, [dataRowStart, dataRowEnd, dataColumnStart, dataColumnEnd]] = getRowColStartEnd(sheet);
    const data = sheet.getRange(dataRowStart, dataColumnStart, dataRowEnd - dataRowStart + 1, dataColumnEnd - dataColumnStart + 1).getValues();

    // Convert data to an object array using worship date as key (convert Date object to string format)
    const dataObjects = buildWorshipObjectArray(headerRow, data);

    const upcomingSundayStr = computeUpcomingSunday(today);

    const upcomingEntry = dataObjects[upcomingSundayStr];

    if (upcomingEntry) {
      logMessage(getCallStackTrace() + `:  Found the Upcoming Sunday entry (${upcomingSundayStr}) in the worship schedule !!!`);
      const forwardEmailArray = collectEmails(upcomingEntry, taskRequireEmailForward, emailMap);

      const forwardEmailList = finalizeEmailList(upcomingSundayStr, dataObjects, praiseTeamLeaders, forwardEmailArray, praiseTeamMembers, mandatoryRecipents, emailMap);

      logMessage(getCallStackTrace() + `:  The forward email list: ${forwardEmailList}`);

      const messages = threads[0].getMessages();
      const combinedMessageBody = combineMessageBodies(messages);
      const subject = messages[0].getSubject();
      const attachments = collectAttachments(messages); // Collect attachments

      if (isTesting) {
        sendTestEmail(emailMap, combinedMessageBody, subject, attachments);
        logMessage(getCallStackTrace() + `:  Running in test mode. Forwarding email to ${emailMap.get("Bill Chu")}`);
      } else {
        sendEmail(forwardEmailList, subject, combinedMessageBody, attachments);
        logMessage(getCallStackTrace() + `:  Running in normal mode. Forwarding email to ${JSON.stringify(forwardEmailList)}`);
      }

      logMessage(getCallStackTrace() + ": Email forwarded successfully.");

      logMessage(getCallStackTrace() + `: Remaining email quota after sending: ${MailApp.getRemainingDailyQuota()}`);

      deleteForwardSermonInfoTrigger(); // Clean up trigger
    } else {
      logMessage(getCallStackTrace() + `: No upcoming worship entry found for ${upcomingSundayStr}.`);
    }

    // Continue to check email if today is not Thursday net
  } else if (today.getDay() == 4) { // Thursday
    logMessage(getCallStackTrace() + ": It's Thursday. Stopping the trigger creation.");
    deleteForwardSermonInfoTrigger();
  } else {
    logMessage(getCallStackTrace() + `: Today is ${today.getDay()}. Continuing to look for sermon info email.`);
    createForwardSermonInfoTrigger();
  }

  flushLogsToSheet();
}

// Function to collect emails from the upcoming worship entry
function collectEmails(upcomingEntry, taskRequireEmailForward, emailMap) {
  let forwardEmailArray = [];

  for (let task in upcomingEntry) {
    if (taskRequireEmailForward.includes(task)) {
      const name = upcomingEntry[task];
      if (name) {
        forwardEmailArray.push(emailMap.get(name));
        logMessage(getCallStackTrace() + `:  Found "${name}" as "${task}"`);
      }
    }
  }

  return forwardEmailArray;
}

// Function to finalize the email list by adding mandatory recipients and praise team
function finalizeEmailList(selectedDate, worshipSchedule, praiseTeamLeaders, forwardEmailArray, praiseTeamMembers, mandatoryRecipents, emailMap) {
  mandatoryRecipents.forEach(name => forwardEmailArray.push(emailMap.get(name)));

  if (isPraiseTeamSunday(selectedDate, worshipSchedule, praiseTeamLeaders)) {
    // Add praise team members, checking for duplicates
    praiseTeamMembers.forEach(name => {
      const email = emailMap.get(name);
      if (email && !forwardEmailArray.includes(email)) {
        forwardEmailArray.push(email);
      }
    });
  }

  // Remove duplicates
  return [...new Set(forwardEmailArray)].join(",");
}

// Function to combine the message bodies from the email thread
function combineMessageBodies(messages) {
  let combinedMessageBody = "";

  combinedMessageBody = messages.reverse().map(msg => msg.getPlainBody()).join('\n');
  //messages.forEach(msg => combinedMessageBody = msg.getPlainBody() + "\n\n" + combinedMessageBody);

  logMessage(getCallStackTrace() + `:  Combined message content to forward: ${combinedMessageBody}, with the total of ${messages.length} messages`);

  return combinedMessageBody;
}

// Function to send the test email
function sendTestEmail(emailMap, combinedMessageBody, subject, attachments) {
  const testRecipient = emailMap.get("Bill Chu");
  GmailApp.sendEmail(testRecipient, subject, combinedMessageBody, {
    attachments: attachments,
  });
  logMessage(getCallStackTrace() + ": Test email sent to Bill Chu only with attachments (if any).");
}

// Function to send the actual email
function sendEmail(forwardEmailList, subject, combinedMessageBody, attachments) {
  GmailApp.sendEmail(forwardEmailList, subject, combinedMessageBody, {
    attachments: attachments,
  });
  logMessage(getCallStackTrace() + ": Email sent to everybody with attachments (if any).");
}

// Function to collect attachments from the email thread and log their file names
function collectAttachments(messages) {
  let attachments = [];

  messages.forEach((msg, index) => {
    const msgAttachments = msg.getAttachments();
    msgAttachments.forEach((attachment, idx) => {
      logMessage(getCallStackTrace() + `: Message ${index + 1}, Attachment ${idx + 1}: ${attachment.getName()}`);
    });
    attachments = attachments.concat(msgAttachments); // Combine all attachments
  });

  logMessage(getCallStackTrace() + `: Collected ${attachments.length} attachments in total.`);
  return attachments;
}
