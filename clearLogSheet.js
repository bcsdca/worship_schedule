function clearLogSheet() {
    
  var logSheet = getLogSheet();
  shiftLogSheets(); // Handle backups before clearing
  // Now clear the current runLog
  // Clear all rows except the header (row 1)
  logMessage(getCallStackTrace() + `: Clearing all the contents of runLog...`)
  if (logSheet.getLastRow() > 1) {
    logSheet.deleteRows(2, logSheet.getLastRow() - 1);
  }

  //const ss = SpreadsheetApp.getActive();
  //const sheet0 = ss.getSheets()[0];
  //ss.setActiveSheet(sheet0); // keep the view at the 1st sheet
  
}

function shiftLogSheets() {
  const ss = SpreadsheetApp.getActive(); // Get active spreadsheet
  const maxLogs = 3; // Maximum backup logs

  // Delete the oldest backup if runLog3 exists
  let oldestSheet = ss.getSheetByName(`runLog${maxLogs}`); //maxLogs = 3 defined in global variable function
  if (oldestSheet) ss.deleteSheet(oldestSheet);

  // Shift runLog2 → runLog3, runLog1 → runLog2
  for (let i = maxLogs - 1; i >= 1; i--) {
    let sheet = ss.getSheetByName(`runLog${i}`);
    if (sheet) sheet.setName(`runLog${i + 1}`);
    logMessage(getCallStackTrace() + `: Renaming runLog${i} to runLog${i + 1}...`)
  }

  // Duplicate runLog to runLog1
  let mainLog = ss.getSheetByName("runLog");
  //make sure runLog is not just created, which should have some content beyond row 1
  if ((mainLog) && (mainLog.getLastRow()) > 1) {
    logMessage(getCallStackTrace() + `: runLog is not empty, so coping all from runLog to runLog1...`)
    mainLog.copyTo(ss).setName("runLog1");
  } else {
    logMessage(getCallStackTrace() + `: runLog is empty, and was just created for the 1st time...`)
  }

}