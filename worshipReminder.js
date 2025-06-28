function worshipReminder(Modes) {

  clearLogSheet();
  
  //checking if Modes is defined or not
  //as well as others setting
  const run_test = Modes?.testing !== undefined ? Modes.testing : true;
  const run_email = Modes?.sendEmail !== undefined ? Modes.sendEmail : false;
  const run_text = Modes?.sendText !== undefined ? Modes.sendText : false;

  logMessage(getCallStackTrace() + `: run_test = ${run_test}, run_email = ${run_email}, run_text = ${run_text}`);

  if ((!run_email) && (!run_text)) {
    logMessage(getCallStackTrace() + `: Exiting, nothing to do, both run_text = ${run_text}, run_email = ${run_email} !!! `);
    return;
  }

  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  logMessage(getCallStackTrace() + "Remaining email quota_before: " + emailQuotaRemaining);

  // get the spreadsheet object
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // set the first sheet as active
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]);
  // fetch this sheet
  var sheet = spreadsheet.getActiveSheet();

  var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const rowColStartEnd = getRowColStartEnd(sheet);

  //actual_startRow is the row after the header row
  const actual_startRow = rowColStartEnd[1][0];
  const actual_lastRow = rowColStartEnd[1][1];
  var start_TaskColumn = rowColStartEnd[1][2];
  const last_TaskColumn = rowColStartEnd[1][3];
  const actual_numRows = actual_lastRow - actual_startRow + 1;

  logMessage(getCallStackTrace() + `: actual_startRow = ${actual_startRow}, actual_lastRow = ${actual_lastRow}, and actual_numRows = ${actual_numRows}`);

  // grab column 1 (the 'date' column) 
  const date_values = getColumnValues(sheet, actual_startRow, start_TaskColumn++, actual_lastRow);
  logMessage(getCallStackTrace() + ": date_values = " + JSON.stringify(date_values));

  // grab column 2 (the 'Type of Service' column)
  const type_of_service = getColumnValues(sheet, actual_startRow, start_TaskColumn++, actual_lastRow);
  logMessage(getCallStackTrace() + ": type_of_service = " + JSON.stringify(type_of_service));

  // grab column 3 (the 'Place of Service' column)
  const place_of_service = getColumnValues(sheet, actual_startRow, start_TaskColumn++, actual_lastRow);
  logMessage(getCallStackTrace() + ": place_of_service = " + JSON.stringify(place_of_service));

  // grab column 4 (the 'Service Start Time' column)
  const service_start_time = getColumnValues(sheet, actual_startRow, start_TaskColumn++, actual_lastRow);
  logMessage(getCallStackTrace() + ": service_start_time = " + JSON.stringify(service_start_time));

  // computing the actual # of TaskColumns
  actual_TaskColumns = last_TaskColumn - start_TaskColumn + 1;
  logMessage(getCallStackTrace() + ": start_TaskColumn = " + start_TaskColumn + "; last_TaskColumn = " + last_TaskColumn + "; actual_TaskColumns = " + actual_TaskColumns);

  // grab all the task name on (actual row - 1), column 5(start_TaskColumn), for actual_TaskColumns columns to the right
  // task name has only 1 row of actual_TaskColumn column of data
  const task_name = getRowValues(sheet, actual_startRow - 1, start_TaskColumn, actual_TaskColumns);
  logMessage(getCallStackTrace() + ": task_name = " + JSON.stringify(task_name));

  // grab contact name columns columns 5 (actual_TaskColumn),
  //staring column #5 (actual_TaskColumn) and for actual_TaskColumns columns after
  const duty_contacts = getRangeValues(sheet, actual_startRow, start_TaskColumn, actual_lastRow - actual_startRow + 1, actual_TaskColumns);
  logMessage(getCallStackTrace() + ": duty_contacts = " + JSON.stringify(duty_contacts));

  const maps = mapContact();
  const email_collection = maps.emailMap;
  const phone_collection = maps.phoneMap;
  const wireless_carrier_collection = maps.carrierMap;

  const today = new Date();

  let emailAdd = email_collection.get("Andy Chu") + ","
    + email_collection.get("Bill Chu") + ","
    + email_collection.get("Sarah Hui") + ","
    + email_collection.get("Pastor Joseph Liang") + ","
    + email_collection.get("Cheong Yik") + ",";

  for (let i = 0; i < actual_numRows; i++) {
    const date = new Date(date_values[i][0]);

    const w_start_date = formatDate(date_values[i][0]);
    const worshipDate = new Date(date_values[i][0]);

    const days_diff = (worshipDate - today) / (1000 * 3600 * 24);

    const w_start_time = formatTime(service_start_time[i][0]);

    const p_start_time = formatTime(service_start_time[i][0] - 30 * 60000);

    logMessage(getCallStackTrace() + `: Today=${today}, Worship Date=${worshipDate}, Worship Time=${w_start_time}, Prayer Time=${p_start_time}, Days Diff=${days_diff}`);

    if (run_email) {
      if (isTodayToSend(days_diff)) {
        //trying find the closest entry(sunday) in spreadsheet, within 7 days from today to send email reminder day
        logMessage(getCallStackTrace() + ": ENTRY FOUND, Sending out email reminder today!!! Today=" + today + ",Worship date= " + worshipDate + ",Worship time= " + w_start_time + ",Prayer time= " + p_start_time);

        const dutyMessages = constructDutyMessages(duty_contacts[i], task_name[0], email_collection);
        logMessage(getCallStackTrace() + ": dutyMessages= " + JSON.stringify(dutyMessages));

        emailAdd = constructEmailList(duty_contacts[i], task_name[0], email_collection, emailAdd);
        logMessage(getCallStackTrace() + ": emailAdd= " + emailAdd);

        //construction of all the email messages
        const htmlReplacements = constructAllMessages(type_of_service[i][0], place_of_service[i][0], dutyMessages, w_start_date, w_start_time, p_start_time, url);
        logMessage(getCallStackTrace() + ": All the email htmlReplacements = " + JSON.stringify(htmlReplacements));

        var htmlBody = HtmlService.createHtmlOutputFromFile('htmlWorshipReminder').getContent()

        for (const [key, value] of Object.entries(htmlReplacements)) {
          htmlBody = htmlBody.replace(key, value);
          logMessage(getCallStackTrace() + `: Key: ${key}, Value: ${value}`);
        }

        //emailImages = {}
        const emailImages = renderImages();

        const emailOptions = {
          htmlBody: htmlBody,
          name: "Cantonese Worship Team",
          inlineImages: emailImages
        };

        if (run_test) {
          GmailApp.sendEmail(email_collection.get("Bill Chu"), `${type_of_service[i][0]} Sunday Service(${w_start_date}) Reminder!!!`, htmlBody, emailOptions);
          logMessage(getCallStackTrace() + ": Sending only test email to Bill Chu!");
        } else {
          GmailApp.sendEmail(emailAdd, `${type_of_service[i][0]} Sunday Service(${w_start_date}) Reminder!!!`, htmlBody, emailOptions);
          //GmailApp.sendEmail(email_collection.get("Bill Chu"), `${type_of_service[i][0]} Sunday Service(${w_start_date}) Reminder!!!`, htmlBody, emailOptions);
          logMessage(getCallStackTrace() + ": Sending email to everybody!");
        }
        break;
      }
    } else if (run_text) {
      if (isTodayToSend(days_diff)) {
        //trying find the closest entry(sunday) in spreadsheet, within 7 days from today to send text reminder day
        logMessage(getCallStackTrace() + ": ENTRY FOUND, Sending out text reminder today!!! Today=" + today + ",Worship date= " + worshipDate + ",Worship time= " + w_start_time + ",Prayer time= " + p_start_time);

        let textReminders = [];

        for (let k = 0; k < actual_TaskColumns; k++) {
          let dutyContact = duty_contacts[i][k];
          let taskName = task_name[0][k];

          if (!dutyContact) continue;

          if (taskName === "Speaker") {
            logMessage(getCallStackTrace() + `: No need to send reminder text to speaker ${dutyContact} !!!`);
            continue;
          }

          if (!phone_collection.has(dutyContact) || !phone_collection.get(dutyContact)) {
            logMessage(getCallStackTrace() + `: No valid phone number found for ${dutyContact} !!!`);
            continue;
          }

          if (!wireless_carrier_collection.has(dutyContact) || !wireless_carrier_collection.get(dutyContact)) {
            logMessage(getCallStackTrace() + `: No Wireless carrier found for ${dutyContact} !!!`);
            continue;
          }

          // Add the valid reminder task to the textReminders array
          textReminders.push({
            taskName: taskName,
            dutyContact: dutyContact,
            w_start_date: w_start_date,
            p_start_time: p_start_time,
            w_start_time: w_start_time,
            place: place_of_service[i][0],
            serviceType: type_of_service[i][0],
            run_test: run_test
          });
        }

        logMessage(getCallStackTrace() + ": textReminders array =" + JSON.stringify(textReminders));

        // Store textReminders in script properties
        let scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.setProperty('textReminders', JSON.stringify(textReminders));
        scriptProperties.setProperty('currentIndex', '0');  // Start from the first text reminder

        //First delete all the existing sendTextReminderTrigger trigger
        let triggers = ScriptApp.getProjectTriggers();
        for (let trigger of triggers) {
          if (trigger.getHandlerFunction() === 'sendTextReminderTrigger') {
            ScriptApp.deleteTrigger(trigger);
          }
        }

        logMessage(getCallStackTrace() + ": Deleting all the previous sendTextReminderTrigger if it exists !!!");

        //Then set up a trigger to send textReminders every 5 minutes
        ScriptApp.newTrigger('sendTextReminderTrigger')
          .timeBased()
          .everyMinutes(5)
          .create();

        logMessage(getCallStackTrace() + ": Done with creating sendTextReminderTrigger, and will wait for the trigger to happen to send out reminder text !!!");
        break;
      }

    }
  }
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  logMessage(getCallStackTrace() + ": Remaining email quota after sending reminder email/text: " + emailQuotaRemaining);
  flushLogsToSheet();

  if ((!run_test) && (!run_text)) {
    //set up the trigger to look for the sermon info email and will try to forward to all the appropriate personels
    //if not testing mode and not run_text mode
    createForwardSermonInfoTrigger()
  }
}

function getColumnValues(sheet, startRow, startColumn, lastRow) {
  return sheet.getRange(startRow, startColumn, lastRow - startRow + 1, 1).getValues();
}

function getRowValues(sheet, startRow, startColumn, numColumns) {
  return sheet.getRange(startRow, startColumn, 1, numColumns).getValues();
}

function getRangeValues(sheet, startRow, startColumn, numRows, numColumns) {
  return sheet.getRange(startRow, startColumn, numRows, numColumns).getValues();
}

function formatDate(date) {
  return Utilities.formatDate(date, timeZone, "MM/dd/yyyy");
}

function formatTime(time) {
  const hours = Utilities.formatDate(new Date(time), timeZone, "HH");
  const minutes = Utilities.formatDate(new Date(time), timeZone, "mm");
  const ampm = hours >= 12 ? 'PM' : 'AM';
  return `${hours}:${minutes} ${ampm}`;
}

function constructDutyMessages(dutyContacts, taskNames, emailCollection) {
  let worship_chairperson = "";
  let msg3 = "";
  let msg_alert = "";
  let msg_worship_chairperson1 = "", msg_worship_chairperson2 = "", msg_worship_chairperson3 = "";
  let msg_on_stage_translator1 = "", msg_on_stage_translator2 = "", msg_on_stage_translator3 = "";
  let msg_mandarin_av_helper1 = "", msg_mandarin_av_helper2 = "", msg_mandarin_av_helper3 = "";

  for (let k = 0; k < taskNames.length; k++) {
    if (!dutyContacts[k]) continue;

    msg3 += `\n${taskNames[k]} - ${dutyContacts[k]}`;
    const email = emailCollection.get(dutyContacts[k]);

    if (taskNames[k] == "Speaker") continue;

    if (!emailCollection.has(dutyContacts[k])) {
      //No dutyContacts[k] name in contact list
      msg_alert += `\nALERT!!! This name \"${dutyContacts[k]}\" was not found in the contact list, thus NO worship reminder email was sent to him/her. Please check spelling or update the contact list.`;
      logMessage(getCallStackTrace() + ` : This name \"${dutyContacts[k]}\" was not found in the contact list, thus NO worship reminder email was sent to him/her. Please check spelling or update the contact list.`);

    } else if (!email) {
      //dutyContacts[k] name was found in contact list, but no email address was found
      msg_alert += `\nALERT!!! ${dutyContacts[k]}\'s email address was not found in the contact list. No worship reminder email was sent to him/her. Please check spelling or update the contact list.`;
      logMessage(getCallStackTrace() + ` : ${dutyContacts[k]}\'s email address was not found in the contact list. No worship reminder email was sent to him/her. Please check spelling or update the contact list.`);

    } else {
      switch (taskNames[k]) {
        case "Worship Chairperson":
          worship_chairperson = dutyContacts[k];
          msg_worship_chairperson1 = "Dear ";
          msg_worship_chairperson2 = dutyContacts[k];
          msg_worship_chairperson3 = ", You are assigned to be the Worship Chairperson for this week. Please \"Reply All\" to this email with the invocation passage, as soon as it is selected. Please use one of the following example formats for your invocation passage. It will be easier if you can just copy and paste one of the example below and modify the content:\n***\nCall to worship: 詩篇 Psalm 1xx:1-5\nCall to worship: 詩篇 Psalm 1xx:11b,12-15\nCall to worship: 詩篇 Psalm 1xx:12-15; 提摩太後書 2Timothy 2:21\n***";
          break;
        case "On-Stage Translator":
          msg_on_stage_translator1 = "\nDear ";
          msg_on_stage_translator2 = dutyContacts[k];
          msg_on_stage_translator3 = ", You are assigned to be the On-Stage Translator for this week, please bring your own personal 1/8 inch (wired) in-ear headset if you have one, If not, please \"Reply All\" to this email thread, and a share one might be available from CEC."
          break;
        case "Mandarin A/V helper":
          msg_mandarin_av_helper1 = "\nDear ";
          msg_mandarin_av_helper2 = dutyContacts[k];
          msg_mandarin_av_helper3 = ", You are assigned to be the Mandarin A/V helper for this week, please arrive at Fellowship Hall by 8:45AM."
          break;
      }
    }
  }
  return { msg3, msg_alert, worship_chairperson, msg_worship_chairperson1, msg_worship_chairperson2, msg_worship_chairperson3, msg_on_stage_translator1, msg_on_stage_translator2, msg_on_stage_translator3, msg_mandarin_av_helper1, msg_mandarin_av_helper2, msg_mandarin_av_helper3 };
}

function constructEmailList(dutyContacts, taskNames, emailCollection, emailAdd) {
  for (let k = 0; k < taskNames.length; k++) {
    if (taskNames[k] !== "Speaker") {
      const email = emailCollection.get(dutyContacts[k]);
      if (email && !emailAdd.includes(email)) {
        emailAdd += `${email},`;
      }
    }
  }
  return emailAdd;
}

function constructAllMessages(typeOfService, placeOfService, dutyMessages, w_start_date, w_start_time, p_start_time, url) {
  //the following section might be changed for join service
  const msg1a = `Dear Brothers and Sisters,\n\nThis is a friendly reminder that you will be serving in the ${typeOfService} Sunday Worship Service for this coming Sunday(${w_start_date}) in the ${placeOfService} with the service start time @${w_start_time}.\n\n`;

  const msg1b_mc = dutyMessages.worship_chairperson;

  var msg1b = ``;
  var msg1c = `Please arrive @ `;
  var msg1c_pt = `${p_start_time}`;
  var msg1d = " to prepare our hearts to serve.";

  if (msg1b_mc) {
    msg1b = `Please note that MC, who is `;
    msg1c = ` for this week, will lead all the Cantonese co-workers, for prayer @`;
    msg1c_pt = `${p_start_time}`;
    msg1d = " to prepare our hearts to serve.";
  }

  const msg2 = "\n\nLook forward to serving together for our Lord!\n\n-CEC Cantonese Worship Ministry\n";

  const htmlReplacements = {
    msg1a: msg1a,
    msg1b: msg1b,
    msg1b_mc: msg1b_mc,
    msg1c: msg1c,
    msg1c_pt: msg1c_pt,
    msg1d: msg1d,
    msg3: dutyMessages.msg3,
    msg_worship_chairperson1: dutyMessages.msg_worship_chairperson1,
    msg_worship_chairperson2: dutyMessages.msg_worship_chairperson2,
    msg_worship_chairperson3: dutyMessages.msg_worship_chairperson3,
    msg_on_stage_translator1: dutyMessages.msg_on_stage_translator1,
    msg_on_stage_translator2: dutyMessages.msg_on_stage_translator2,
    msg_on_stage_translator3: dutyMessages.msg_on_stage_translator3,
    msg_mandarin_av_helper1: dutyMessages.msg_mandarin_av_helper1,
    msg_mandarin_av_helper2: dutyMessages.msg_mandarin_av_helper2,
    msg_mandarin_av_helper3: dutyMessages.msg_mandarin_av_helper3,
    msg2: msg2,
    msg_alert: dutyMessages.msg_alert,
    url: url
  };

  return htmlReplacements;
}

function isTodayToSend(days_diff) {
  //looking for the days difference within 1 week
  return ((days_diff > 0 && days_diff < 6.45));
}

