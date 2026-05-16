/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 * @see https://developers.google.com/apps-script/guides/triggers#onopene
 */

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  let i = 1;

  const mainMenu = ui.createMenu('🎉 Cantonese Worship Service Utilities 🎉');

  //
  // ---- STEP Preparation ----
  //
  const stepPreparation = ui.createMenu('🛠️ Step "Preparation" – Prepare Coming Quarter Setup(~1 month before the new quarter)')
    .addItem(`${i++}. Start new schedule preparation`, 'preparation')
    .addItem(`${i++}. Send "Unavailable Date to Serve" Email`, 'sendUnavailableDatesEmail');

  //
  // ---- STEP Task Assignment ----
  //
  const stepTaskAssignment = ui.createMenu('⚙️ Step "Task Assignment" – Task Assignment & Email Review(~2 weeks before the new quarter)')
    .addItem(`${i++}. Prepare Task Assignment`, 'prepareTaskassignement')
    .addItem(`${i++}. Automatic A/V Task Assignment`, 'autoTaskAssignSidebar')
    .addItem(`${i++}. Update Dashboard`, 'buildPivotTable')
    .addItem(`${i++}. Send out Review Task Assignment Email`, 'quarterlyRMailMerge');

  //
  // ---- STEP Go Live ----
  //
  const stepGoLive = ui.createMenu('📈 Step "Go Live" – Worship Schedule Go Live(Before the 1st Tuesday of the new quarter)')
    .addItem(`${i++}. Go Live !!!`, 'goLive');
    
  //
  // ---- STEP optional ----
  //
  const optionalSteps = ui.createMenu('🔹 Step "Optional" - Optional steps for managing worship schedule')
    .addItem(`${i++}. (optional) Stop A/V Schedule Change Monitoring`, 'removeScheduleChangeTrigger')
    .addItem(`${i++}. (optional) Stop YouTube Stat Collection`, 'removeGetYouTubeStatsTrigger')
    .addItem(`${i++}. (optional) Remove All Triggers immediately`, 'removeAllTriggers');

  //
  // Add steps to main menu
  //
  mainMenu
    .addSubMenu(stepPreparation)
    .addSubMenu(stepTaskAssignment)
    .addSubMenu(stepGoLive)
    .addSubMenu(optionalSteps)
    .addToUi();
}

function preparation() {
  buildHistoricalNormalizeSchedule(); //Save Old Worship Data to "Historical Normalized Data
  cleanWorshipSchedule(); //Delete All Data in "Cantonese Worship Schedule" (set next quarter date)
  cleanYouTubeStat(); //Delete All Data in "YouTube Stat"
  cleanUnavailableDates(); //Delete All Data + initialize checkboxes in "Unavailable Dates"
  createUnavailableDatesTrigger(); //start accepting updates in "Unavailable Dates"
}

function prepareTaskassignement() {
  removeUnavailableDatesTrigger(); //Stop accepting updates in "Unavailable Dates"
  updateAllDropDowns(); //Update all dropdowns
}

function goLive() {
  emailReminderSidebar(); //Schedule Email Reminder
  textReminderSidebar(); //Schedule Text Reminder
  addScheduleChangeTrigger(); //Start A/V Schedule Change Monitoring
  addGetYouTubeStatsTrigger(); //Start YouTube Stat Collection
  scheduleQuarterlyTriggerCleanup(); //Schedule End of Quarter cleanup
}

function emailReminderSidebar() {
  var widget = HtmlService.createHtmlOutputFromFile("htmlSelDay_email");
  widget.setTitle("Worship Email Reminder Day Selection").setWidth(300);
  SpreadsheetApp.getUi().showSidebar(widget);
}

function textReminderSidebar() {
  var widget = HtmlService.createHtmlOutputFromFile("htmlSelDay_text");
  widget.setTitle("Worship Text Reminder Day Selection").setWidth(300);
  SpreadsheetApp.getUi().showSidebar(widget);
}

function autoTaskAssignSidebar() {
  var widget = HtmlService.createTemplateFromFile("htmlAutoAssign").evaluate();
  widget.setTitle("Automatic Task Assignment For Co-workers");
  SpreadsheetApp.getUi().showSidebar(widget);
}

function closeSidebar() {
  var html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
  SpreadsheetApp.getUi().showSidebar(html);
}

