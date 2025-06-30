// sendEmailList is based on the contact, and not all co-workers in the schedule
// test or normal run is based on the "ExceptionMMTemplate" tab of the worship Google Slide
/*
function sendGoogleFormExceptionDateEmail() {
  clearLogSheet();
  // Get remaining email quota
  let emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  logMessage(getCallStackTrace() + `Remaining email quota_before: ${emailQuotaRemaining}`);

  const exceptionMMTemplateSheet = SpreadsheetApp.getActive().getSheetByName("exceptionMMTemplate");

  // Fetch email subject, content, test mode, test person, and Google Form URL from the sheet
  const emailSubject = exceptionMMTemplateSheet.getRange(1, 2).getDisplayValue();
  const msg1 = exceptionMMTemplateSheet.getRange(2, 2).getDisplayValue();
  const isTestMode = exceptionMMTemplateSheet.getRange(3, 2).getDisplayValue() === "TRUE";
  const testPerson = exceptionMMTemplateSheet.getRange(4, 2).getDisplayValue();
  const googleFormURL = exceptionMMTemplateSheet.getRange(5, 2).getDisplayValue();

  // Fetch contact information from the "contact" sheet
  const contactSheet = SpreadsheetApp.getActive().getSheetByName("contact");
  const contactData = contactSheet.getDataRange().getDisplayValues();
  
  // dont send this email to the persons below
  const excludedNames = ["Alfred Ip", "Hilbert Chu", "Peggy Chu", "Winkie Zhang", "Wellington Hui"];
  
  // Define the array of tasks to exclude if they are the only tasks someone is performing, 
  // and not sending email to these person
  // it is lower case because it was used to compare with the force to lower case in isExcludedTask function
  const excludeSME = ["usher/welcome", "speaker"];

  // Prepare the AV email object
  const maps = mapContact();
  const emailMap = maps.emailMap;  // This is a Map object

  // Initial email array with specific people
  let emailArray = ["Andy Chu", "Bill Chu", "Edmond Chan"].map(name => emailMap.get(name));

  // Collect email addresses based on specific tasks
  contactData.forEach((row, index) => {
    if (index > 0) { // Skip header row
      const [name, email, , , smeColumn] = row;
      if (excludedNames.some(excludedName => excludedName.toLowerCase() === name.toLowerCase())) {
        return; // Skip this row if the name is excluded
      }

      if (email && !emailArray.includes(email)) {
        let smeArray = smeColumn.split(",");
        
        // Check if any SME should be excluded
        const shouldExclude = smeArray.some(sme => {
        const isExcluded = isExcludedTask(excludeSME, sme);
        const hasValidOtherSME = smeArray.some(otherSme =>
        otherSme !== sme && !isExcludedTask(excludeSME, otherSme)
        );

        // Log details for debugging
        if (isExcluded && !hasValidOtherSME) {
          logMessage(getCallStackTrace() + `: Excluding task: ${smeArray}, Email: ${email}, Row: ${index + 1}`);
        }

        return isExcluded && !hasValidOtherSME;
        });

        //logMessage(getCallStackTrace() + "Final shouldExclude:", shouldExclude);
        //logMessage(getCallStackTrace() + `Excluding task: Email: ${email}, Row: ${index + 1}`);

        if (!shouldExclude) {
          emailArray.push(email);
        }
      }
    }
  });

  let sendEmailList = emailArray.join(",");

  // Get worship schedule sheet data
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const sheetName = sheet.getName();
  const rowColStartEnd = findRowColStartEnd(sheet);
  const data = sheet.getRange(rowColStartEnd[1][0], rowColStartEnd[1][2], rowColStartEnd[1][1] - rowColStartEnd[1][0] + 1, rowColStartEnd[1][3] - rowColStartEnd[1][2] + 1).getValues();

  // Process data for task assignment
  let objectArray = [];
  processWorshipSchedule(data, objectArray);

  // Sort objectArray by name and date
  objectArray.sort((a, b) => a.name.localeCompare(b.name) || a.date.localeCompare(b.date));

  // Validate emails and prepare alert messages
  let msgAlert = validateEmailAddresses(objectArray, emailMap);

  let msg2 = "\nLook forward to serving together for our Lord!\n\n-CEC Cantonese Worship Ministry\n";
  logMessage(getCallStackTrace() + `: Email list for sending out this google form for exception report: ${sendEmailList}`);
  
  // Handle test mode
  if (isTestMode) {
    sendEmailList = emailMap.get(testPerson);
    logMessage(getCallStackTrace() + `: Test mode enabled. Sending email only to "${testPerson}" with this email address ${sendEmailList}`);
  } else {
    const numEmailRecipients = sendEmailList.split(",").length;
    logMessage(getCallStackTrace() + `: Normal mode. About to send ${numEmailRecipients} emails to co-workers.`);

    if (!confirmEmailSending(numEmailRecipients)) return;
  }

  // Prepare email body and send emails
  const emailBody = prepareEmailBody(msg1, msgAlert, msg2, googleFormURL);

  sendEmails(sendEmailList, emailSubject, emailBody, sheetName);

  logMessage(getCallStackTrace() + `Done. Exception date selection emails were sent.`);
  SpreadsheetApp.getActive().toast("Emails sent successfully!");

  emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  logMessage(getCallStackTrace() + `Remaining email quota_after: ${emailQuotaRemaining}`);
  flushLogsToSheet();
}

// Function to process worship schedule data and store in objectArray
function processWorshipSchedule(data, objectArray) {
  for (let i = 1; i < data.length; i++) { // Skip header row
    for (let j = 4; j < data[i].length; j++) {
      let name = data[i][j];
      if (name) {
        let task = data[0][j];
        if (!["Speaker", "Usher/Welcome1", "Usher/Welcome2", "Usher/Welcome3"].includes(task)) {
          objectArray.push({
            name: name,
            task: task,
            date: Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'MM/dd/yyyy')
          });
        }
      }
    }
  }
}

// Function to validate emails and create an alert message
function validateEmailAddresses(objectArray, emailMap) {
  let msgAlert = "";
  let listNoEmailAddress = "";
  let listNoName = "";

  objectArray.forEach(entry => {
    const tempName = entry.name;
    const email = emailMap.get(tempName);  // Use Map.get() here as well
    if (!email) {
      if (!listNoName.includes(tempName)) {
        msgAlert += `\nALERT!!! No email was sent to "${tempName}", as this name was not found in the contact list.`;
        listNoName += `${tempName},`;
      }
    }
  });

  return msgAlert;
}

// Function to prepare the email body
function prepareEmailBody(msg1, msgAlert, msg2, googleFormURL) {
  return HtmlService.createHtmlOutputFromFile('htmlFileException').getContent()
    .replace("msg1", msg1)
    .replace("msg_alert", msgAlert)
    .replace("msg2", msg2)
    .replace("url", googleFormURL);
}

// Function to confirm email sending with the user
function confirmEmailSending(numEmailRecipients) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(`You are about to send ${numEmailRecipients} emails. Are you sure?`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) {
    logMessageError(getCallStackTrace() + `User canceled the request.`);
    SpreadsheetApp.getUi().alert("Request canceled.");
    return false;
  }
  return true;
}

// Function to send emails
function sendEmails(sendEmailList, emailSubject, emailBody, sheetName) {
  MailApp.sendEmail({
    to: sendEmailList,
    //to: "shui.bill.chu@gmail.com",
    subject: `${emailSubject} (${sheetName})`,
    name: "Cantonese Worship Team",
    htmlBody: emailBody
  });
}


// Function to check if a task should be excluded
function isExcludedTask(excludeSME, task) {
  const trimmedTask = task.trim(); // Remove leading/trailing spaces
  return excludeSME.some(excludeTask =>
    trimmedTask.toLowerCase().startsWith(excludeTask)
  );
}
*/