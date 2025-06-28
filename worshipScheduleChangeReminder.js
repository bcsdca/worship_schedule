//only send email when the toList is not undefined or there is oldValue or inputValue (4/4/2023)
//adding more messages if no email address was found to send adding assignment email (4/5/2023)
//only modification on sheet0 was be monitored..
//using a lot of functions to streamline the code
//adding 2 rows for description of changes and who is making the change
//change the style of the table
function worshipScheduleChangeReminder(e, test) {
  logMessage(getCallStackTrace() + ": " + JSON.stringify(e, null, 2));

  clearLogSheet();

  if (e.source.getSheetName() == SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getName()) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var sheetName = sheet.getName();
    const rowColStartEnd = findRowColStartEnd(sheet);

    const headerRow = rowColStartEnd[0];

    const dataRowStart = rowColStartEnd[1][0];
    const dataRowEnd = rowColStartEnd[1][1];
    //const dataColumnStart = rowColStartEnd[1][2];
    const dataColumnEnd = rowColStartEnd[1][3];

    const maps = mapContact();
    const emailCollection = maps.emailMap;

    var inputValue = e.value;
    var oldValue = e.oldValue;

    var modifyRow = e.range.getRow();
    var modifyDate = sheet.getRange(modifyRow, 1).getDisplayValue();
    var modifyColumn = e.range.getColumn();
    var taskName = headerRow[modifyColumn - 1];
    var toDay = new Date();
    logMessage(getCallStackTrace() + `: toDay = ${toDay}, modifyDate = ${modifyDate}, modifyRow = ${modifyRow}, modifyColumn = ${modifyColumn}, oldValue = ${oldValue}, and inputValue = ${inputValue}`)
    if (shouldProcessChange(oldValue, inputValue, modifyRow, dataRowStart, dataRowEnd, modifyColumn, dataColumnEnd, modifyDate, toDay)) {
      logMessage(getCallStackTrace() + ": Prompt the user for change confirmation");
      var ui = SpreadsheetApp.getUi();
      var response = handleUserConfirmation(ui, oldValue, inputValue, taskName, modifyDate);

      if (response == ui.Button.YES) {
        var descriptionResponse = ui.prompt('Please enter a short description of the change:');
        var userName = promptForUserName(ui);
        var changeDescription = descriptionResponse.getResponseText();

        if (oldValue) {
          handleAssignmentChange("Removed", emailCollection, oldValue, inputValue, taskName, modifyDate, modifyRow, modifyColumn, userName, changeDescription, toDay, test, sheetName);
        }

        if (inputValue) {
          handleAssignmentChange("Added", emailCollection, oldValue, inputValue, taskName, modifyDate, modifyRow, modifyColumn, userName, changeDescription, toDay, test, sheetName);
        }

        buildPivotTable();
        logMessage(getCallStackTrace() + ": Updates the Pivot Table and thus the dashboard !!!");
      } else {
        handleUserCancellation(sheet, modifyRow, modifyColumn, oldValue);
        flushLogsToSheet();
        return;
      }
    } else {
      logMessage(getCallStackTrace() + ": Do nothing because of one of the checking conditions not meet !!!");
      SpreadsheetApp.getActive().toast("Do nothing because of one of the checking conditions not meet!! 👎 ", e.source.getSheetName());
    }
  }

  flushLogsToSheet();
}

function handleUserConfirmation(ui, oldValue, inputValue, taskName, modifyDate) {
  if ((oldValue != "") && (inputValue == undefined)) {
    return ui.alert('Do you really want to remove assignment for "' + oldValue + '" as the "' + taskName + '" on "' + modifyDate + '\" and nobody to replace him or her ?', ui.ButtonSet.YES_NO);
  } else if ((oldValue == undefined) && (inputValue != "")) {
    return ui.alert('Do you really want to add assignment for "' + inputValue + '" as the "' + taskName + '" on "' + modifyDate + '\" ?', ui.ButtonSet.YES_NO);
  } else if ((oldValue != "") && (inputValue != "")) {
    return ui.alert('Do you really want to substitute "' + oldValue + '" with "' + inputValue + '" as the "' + taskName + '" on "' + modifyDate + '\" ?', ui.ButtonSet.YES_NO);
  } else {
    ui.alert("Illegal combination of old existing name and new input name !!!");
    return ui.Button.NO;
  }
}

function handleUserCancellation(sheet, modifyRow, modifyColumn, oldValue) {
  logMessage(getCallStackTrace() + ": The user canceled the request.");
  SpreadsheetApp.getUi().alert("The user has canceled the request !!!");
  SpreadsheetApp.getActive().toast("Done, The user has canceled the request, restore the original value and no email will be sent ⚠️ !!!");
  sheet.getRange(modifyRow, modifyColumn, 1, 1).setValue(oldValue);
}

function promptForUserName(ui) {
  var nameResponse = ui.prompt('Please select your name:', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() == ui.Button.OK) {
    return nameResponse.getResponseText();
  }
  return "";
}

function shouldProcessChange(oldValue, inputValue, modifyRow, dataRowStart, dataRowEnd, modifyColumn, dataColumnEnd, modifyDate, toDay) {
  return (oldValue != inputValue) &&
    (modifyRow <= dataRowEnd) &&
    (modifyRow >= dataRowStart) &&
    (4 < modifyColumn) &&
    (modifyColumn <= dataColumnEnd) &&
    (modifyColumn != 6) &&
    (new Date(modifyDate) > toDay);
}

function handleAssignmentChange(actionType, emailCollection, oldValue, inputValue, taskName, modifyDate, modifyRow, modifyColumn, userName, changeDescription, toDay, test, sheetName) {
  var affectedName = actionType === "Removed" ? oldValue : inputValue;

  // Load the HTML template
  var htmlTemplate = HtmlService.createTemplateFromFile('htmlChangeTableTemplate');

  // Set the values in the template
  htmlTemplate.changeType = actionType;
  htmlTemplate.changeTypeClass = actionType === "Added" ? "change-type-added" : "change-type-removed";
  htmlTemplate.userName = userName;
  htmlTemplate.changeDescription = changeDescription;
  htmlTemplate.taskName = taskName;
  htmlTemplate.affectedName = affectedName;
  htmlTemplate.modifyDate = modifyDate;
  htmlTemplate.toDay = toDay;

  // Evaluate the template to generate the final HTML
  var table = htmlTemplate.evaluate().getContent();

  var toList = test ? emailCollection.get("Bill Chu") : (emailCollection.get(affectedName) || "");
  var ccList = test ? emailCollection.get("Bill Chu") : emailCollection.get("Andy Chu") + "," + emailCollection.get("Bill Chu") + "," + emailCollection.get("Edmond Chan");

  if (affectedName && emailCollection.get(affectedName)) {
    MailApp.sendEmail({
      to: toList,
      cc: ccList,
      subject: `Assignment Change in "${sheetName}" for the week of ${modifyDate}`,
      htmlBody: table
    });
    logMessage(getCallStackTrace() + `: Sending email to ${affectedName} for ${actionType.toLowerCase()} assignment of ${taskName} for the worship date of ${modifyDate}, with email address ${emailCollection.get(affectedName)}`);
    SpreadsheetApp.getActive().toast(`Successful in sending ${actionType.toLowerCase()} assignment email to the above co-worker 👍 !!!`, affectedName);
  } else {
    logMessageError(getCallStackTrace() + `: No email address was found for ${affectedName}, No ${actionType.toLowerCase()} assignment email was sent.`);
    SpreadsheetApp.getActive().toast(`Failed in sending ${actionType.toLowerCase()} assignment email because no email address was found for ${affectedName} ❌ !!!`, affectedName);
  }
}
