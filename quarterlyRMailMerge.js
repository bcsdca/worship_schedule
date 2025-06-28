function quarterlyRMailMerge() {

  clearLogSheet();
  const emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  logMessage(`${getCallStackTrace()}: Remaining email quota_before: ${emailQuotaRemaining}`);
  
  const templateSheet = SpreadsheetApp.getActive().getSheetByName("qRMMTemplate");
  const emailSubject = templateSheet.getRange(1, 2).getDisplayValue();
  const emailContent = templateSheet.getRange(2, 2).getDisplayValue();
  const isTestMode = templateSheet.getRange(3, 2).getDisplayValue() === "Yes";
  const testRecipient = templateSheet.getRange(4, 2).getDisplayValue();
  const includeGif = templateSheet.getRange(5, 2).getDisplayValue() === "Yes";
  const sendGroupEmailOnly = templateSheet.getRange(6, 2).getDisplayValue() === "Yes"; // Group only flag
  const sendSpecialEmailOnly = templateSheet.getRange(7, 2).getDisplayValue() === "Yes"; // Special only flag
  const specialPeople = templateSheet.getRange(8, 2).getDisplayValue().split(',').map(value => value.trim()); // list of all people
  
  // Early exit: if both flags are false, no email is sent
  if (!sendGroupEmailOnly && !sendSpecialEmailOnly) {
    logMessage(`${getCallStackTrace()}: No emails are sent because both sendGroupEmailOnly and sendSpecialEmailOnly are false.`);
    SpreadsheetApp.getActive().toast("No emails sent. Both email flags are disabled.");
    return;
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy-HH:mm');
  const emailMap = mapContact().emailMap;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const sheetName = sheet.getName();

  const rowColStartEnd = getRowColStartEnd(sheet);
  const headerRow = rowColStartEnd[0];
  const [dataRowStart, dataRowEnd, dataColumnStart, dataColumnEnd] = rowColStartEnd[1];

  const taskData = sheet.getRange(dataRowStart, dataColumnStart, dataRowEnd - dataRowStart + 1, dataColumnEnd - dataColumnStart + 1).getValues();
  
  //excludedTasks  is a global variable to prevent sending email for this tasks defined in excludedTasks
  const taskAssignments = processTaskAssignments(taskData, headerRow, excludedTasks);
  
  // Filter assignments for general and special people
  const generalAssignments = taskAssignments.filter(assignment => !specialPeople.includes(assignment.name));
  const specialAssignments = taskAssignments.filter(assignment => specialPeople.includes(assignment.name));

  const msg3 = formatTaskAssignments(taskAssignments);

  // Initialize alert message (for missing emails/names)
  let alertMsg = ""; // This is the `msg_alert`

  // Prepare email lists based on flags
  let sendEmailList;
  let specialEmailList = "";

  if (sendGroupEmailOnly && !sendSpecialEmailOnly) {
    // If we are only sending general emails and NOT sending special emails, use the full taskAssignments
    logMessage(`${getCallStackTrace()}: sendGroupEmailOnly = ${sendGroupEmailOnly}, sendSpecialEmailOnly = ${sendSpecialEmailOnly}`);
    sendEmailList = prepareEmailList(taskAssignments, emailMap, isTestMode, testRecipient, alertMsg);
    sendGeneralEmail(sendEmailList, emailSubject, msg3, emailContent, url, sheetName, today, includeGif, alertMsg, isTestMode);
  } else if (sendGroupEmailOnly) {
    // The original logic: filter out special people if sending general emails
    logMessage(`${getCallStackTrace()}: sendGroupEmailOnly = ${sendGroupEmailOnly}`);
    sendEmailList = prepareEmailList(generalAssignments, emailMap, isTestMode, testRecipient, alertMsg);
    sendGeneralEmail(sendEmailList, emailSubject, msg3, emailContent, url, sheetName, today, includeGif, alertMsg, isTestMode);
  }

  if (sendSpecialEmailOnly) {
    // Prepare special email list
    logMessage(`${getCallStackTrace()}: sendSpecialEmailOnly = ${sendSpecialEmailOnly}`);
    specialEmailList = prepareEmailList(specialAssignments, emailMap, isTestMode, testRecipient, alertMsg);
    sendSpecialEmails(specialPeople, specialAssignments, emailSubject, emailContent, url, sheetName, today, includeGif, emailMap, alertMsg, isTestMode, testRecipient);
  }


  // Log remaining email quota
  const remainingQuotaAfter = MailApp.getRemainingDailyQuota();
  logMessage(`${getCallStackTrace()}: Remaining email quota_after: ${remainingQuotaAfter}`);
  flushLogsToSheet()
}

// Function to format task assignments for email
function formatTaskAssignments(assignments) {
  let lastName = "";
  return assignments.map(assignment => {
    const nameToDisplay = assignment.name !== lastName ? assignment.name : "";
    lastName = assignment.name;
    let taskLine = nameToDisplay ? `<div style="margin-top: 10px; font-weight: bold;">${nameToDisplay.trim()}</div>` : "";
    taskLine += `<div style="margin-left: 20px; font-size: 9pt;">&#8226; "${assignment.task}" on ${assignment.date}</div>`;
    return taskLine;
  }).join("\n");
}

// Function to prepare the email list
function prepareEmailList(assignments, emailMap, isTestMode, testRecipient, alertMsg) {
  const emailList = new Set();
  assignments.forEach(({ name }) => {
    const email = emailMap.get(name);
    if (email) {
      emailList.add(email);
    } else {
      alertMsg += `ALERT: No email address found for ${name}!\n`;
      logMessageError(`${getCallStackTrace()}: ALERT: No email address found for ${name}`);
    }
  });

  // Log the content of the emailList as an array
  logMessage(`${getCallStackTrace()}: This is actual non-test mode email list: ${Array.from(emailList)}`);
  
  if (isTestMode) {
    return emailMap.get(testRecipient) + ",";
  }
  return [...emailList].join(",") + ",";
}

// Function to send general email
function sendGeneralEmail(to, subject, msg3, msg1, url, sheetName, today, includeGif, alertMsg, isTestMode) {
  const msg2 = "\n\nLook forward to serving together for our Lord!\n\n-CEC Cantonese Worship Ministry\n"; // This is the `msg2`
  const templateFileName = includeGif ? 'htmlFileQrGIF' : 'htmlFileQr';

  const emailBody = HtmlService.createHtmlOutputFromFile(templateFileName)
    .getContent()
    .replace("msg1", msg1)   // Main email content
    .replace("msg3", msg3)   // Task assignments
    .replace("msg_alert", alertMsg)   // Warnings and alerts
    .replace("msg2", msg2)   // Closing message
    .replace("url", url);

  if (!isTestMode) {
    const recipientCount = to.split(",").length - 1;
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(`You are about to send ${recipientCount} emails. Is this correct?`, ui.ButtonSet.YES_NO);
    if (response === ui.Button.NO) return;
  }

  MailApp.sendEmail({
    to,
    subject: `${subject} (${sheetName}) as of ${today}`,
    name: "Cantonese Worship Team",
    htmlBody: emailBody,
    inlineImages: includeGif ? { special: DriveApp.getFileById("15lr4FWa8UTo3RrCTR037cMm-BQhBRJfn").getAs("image/gif") } : undefined
  });
  logMessage(`${getCallStackTrace()}: Send review email to ${to}`);
}

// Function to send special personalized emails
function sendSpecialEmails(specialPeople, specialAssignments, subject, msg1, url, sheetName, today, includeGif, emailMap, alertMsg, isTestMode, testRecipient) {
  const msg2 = "\n\nAbove is your personalized task assignment. Look forward to serving together for our Lord!\n\n-CEC Cantonese Worship Ministry\n"; // Special closing message
  const templateFileName = includeGif ? 'htmlFileQrGIF' : 'htmlFileQr';

  specialPeople.forEach(name => {
    const assignments = specialAssignments.filter(a => a.name === name);
    if (assignments.length === 0) {
      logMessage(`${getCallStackTrace()}: ${name} has no task assignment. No email sent.`);
      return;
    }

    let to = isTestMode ? emailMap.get(testRecipient) : emailMap.get(name);
    let cc = isTestMode ? emailMap.get(testRecipient) : `${emailMap.get("Andy Chu")},${emailMap.get("Bill Chu")},${emailMap.get("Edmond Chan")}`;

    const emailBody = HtmlService.createHtmlOutputFromFile(templateFileName)
      .getContent()
      .replace("msg1", msg1)   // Main email content
      .replace("msg3", formatTaskAssignments(assignments))   // Task assignments for special person
      .replace("msg_alert", alertMsg)   // Warnings and alerts
      .replace("msg2", msg2)   // Closing message
      .replace("url", url);

    MailApp.sendEmail({
      to,
      cc,
      subject: `Your personalized task assignment: ${subject} (${sheetName}) as of ${today}`,
      name: "Cantonese Worship Team",
      htmlBody: emailBody,
      inlineImages: includeGif ? { special: DriveApp.getFileById("15lr4FWa8UTo3RrCTR037cMm-BQhBRJfn").getAs("image/gif") } : undefined
    });
    logMessage(`${getCallStackTrace()}: Sent personalized review email to ${to}, and cc to ${cc}`);
  });
}



