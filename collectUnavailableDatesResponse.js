function collectUnavailableDatesResponse(e) {
  const LOCK_TIMEOUT_MS = 30000; // Adjustable timeout value, 30000 proven good
  
  logMessage(getCallStackTrace() + ": The trigger event = " + JSON.stringify(e, null, 2));

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  if (sheetName !== 'Unavailable Dates') {
    logMessage(getCallStackTrace() + ': Do nothing becasue edit occurred on a different sheet: ' + sheetName);
    return;
  }

  // Try to acquire the lock
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
    logMessage(getCallStackTrace() + ': Could not acquire lock — skipping the recording of this edit !!!');
    return;
  }

  logMessage(getCallStackTrace() + ': Acquire lock — saving this edit !!!');

  try {
    const sheet = e.source.getSheetByName('Unavailable Dates');
    const range = e.range;
    const editedValue = range.getValue();

    const date = sheet.getRange(range.getRow(), 1).getValue();   // Column A = Date
    const name = sheet.getRange(1, range.getColumn()).getValue(); // Row 1 = Name

    if (e.range.columnStart === 1) {
      logMessage(getCallStackTrace() + ': Editing column 1 — skipping this edit');
      return;
    }

    if (e.range.rowStart === 1) {
      logMessage(getCallStackTrace() + ': Editing row 1 — skipping this edit');
      return;
    }

    if (!name || !date) {
      logMessage(getCallStackTrace() + ': Missing name or date — skipping this edit');
      return;
    }
    
    const status = editedValue === true ? 'check' : 'uncheck';

    const responseSheet = e.source.getSheetByName('Unavailable Dates Response');
    responseSheet.appendRow([new Date(), date, name, status]);
    logMessage(getCallStackTrace() + ': Finished saving this edit to "Unavailable Date Response" tab !!!');

  } catch (err) {
    logMessage(getCallStackTrace() + ': Error during handleEdit: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}






