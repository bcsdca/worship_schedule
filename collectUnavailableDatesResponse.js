function collectUnavailableDatesResponse(e) {
  const LOCK_TIMEOUT_MS = 30000; // Max wait time to acquire lock

  logMessage(getCallStackTrace() + ": The trigger event = " + JSON.stringify(e, null, 2));

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  if (sheetName !== 'Unavailable Dates') {
    logMessage(getCallStackTrace() + ': Do nothing because edit occurred on a different sheet: ' + sheetName);
    return;
  }

  // Acquire a lock to ensure appendRow doesn't collide with other simultaneous edits
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT_MS); // Wait up to 30 seconds for the lock
  } catch (err) {
    logMessageError(getCallStackTrace() + ': Could not acquire lock in time — skip logging of this edit !!!');
    return;
  }

  logMessage(getCallStackTrace() + ': Acquire lock — saving this edit !!!');

  try {
    const sheet = e.source.getSheetByName('Unavailable Dates');
    const range = e.range;

    // Normalize e.value to boolean (true if checked, false if unchecked or cleared)
    const editedValue = (typeof e.value === 'string' && e.value.toUpperCase() === 'TRUE');

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

    const status = editedValue ? 'check' : 'uncheck';

    const responseSheet = e.source.getSheetByName('Unavailable Dates Response');
    responseSheet.appendRow([new Date(), date, name, status]);

    logMessage(getCallStackTrace() + `: This entry's e.value = ${e.value}, normalized editedValue = ${editedValue}, status = ${status}`);
    logMessage(getCallStackTrace() + ': Finished saving this edit to "Unavailable Dates Response" tab !!!');

  } catch (err) {
    logMessage(getCallStackTrace() + ': Error during handleEdit: ' + err.toString());
  } finally {
    // Always release the lock
    lock.releaseLock();
  }
}
