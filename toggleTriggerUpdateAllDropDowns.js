function toggleTriggerUpdateAllDropDowns(action) {
  const functionName = "updateAllDropDowns";

  if (action === "enable") {
    // Remove existing trigger first (avoid duplicates)
    removeTriggerUpdateAllDropDowns(functionName);

    // Create new trigger at 12:30 AM
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .atHour(0)       // midnight
      .nearMinute(30)  // ~12:30 AM
      .everyDays(1)
      .create();

    logMessage(getCallStackTrace() + `: Trigger enabled for updateAllDropDowns at 12:30 AM daily.`);
  } else if (action === "disable") {
    removeTriggerUpdateAllDropDowns(functionName);
    logMessage(getCallStackTrace() + `: Trigger disabled for updateAllDropDowns.`);
  } else {
    logMessageError(getCallStackTrace() + `: Invalid input. Use "enable" or "disable.`);
    throw new Error('Invalid input. Use "enable" or "disable".');
  }
}

function removeTriggerUpdateAllDropDowns(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(t);
    }
  }
}

