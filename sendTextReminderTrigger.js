function sendTextReminderTrigger() {
    let scriptProperties = PropertiesService.getScriptProperties();
    let textReminders = JSON.parse(scriptProperties.getProperty('textReminders'));
    let currentIndex = parseInt(scriptProperties.getProperty('currentIndex'), 10);

    // Check if we are done sending all textReminders
    if (currentIndex >= textReminders.length) {
      // No more textReminders to send, delete the trigger
      let triggers = ScriptApp.getProjectTriggers();
      for (let trigger of triggers) {
        if (trigger.getHandlerFunction() === 'sendTextReminderTrigger') {
          ScriptApp.deleteTrigger(trigger);
        }
      }
      logMessage(getCallStackTrace() + "All sendTextReminderTrigger deleted.");
      return;
    }

    // Get the current textReminder
    let textReminder = textReminders[currentIndex];

    logMessage(getCallStackTrace() + `: Starting to send the reminder text # ${currentIndex + 1}.`);

    // Generate a random delay between 0 to 1 minutes (in milliseconds)
    var delay = Math.random() * 60000; // Random delay between 0 and 1 min (in milliseconds)

    // Pause execution for the random delay time
    logMessage(getCallStackTrace() + ": Waiting for " + (delay / 60000) + " minutes before sending the actual reminder text.");
    Utilities.sleep(delay);

    // Send the textReminder text
    sendTextReminder(
      textReminder.taskName,
      textReminder.dutyContact,
      textReminder.w_start_date,
      textReminder.p_start_time,
      textReminder.w_start_time,
      textReminder.place,
      textReminder.serviceType,
      textReminder.run_test

    );

    // Update the current index
    scriptProperties.setProperty('currentIndex', (currentIndex + 1).toString());
  }