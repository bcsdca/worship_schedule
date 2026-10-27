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
  const stepPreparation = ui.createMenu('🛠️ Step "Preparation" – Coming Quarter\'s Schedule Preparation(~1 month before the new quarter)')
    .addItem(`${i++}. Start New Quarter's Schedule Preparation`, 'preparation')
    .addItem(`${i++}. Send "Unavailable Date to Serve" Email`, 'sendUnavailableDatesEmail');

  //
  // ---- STEP Task Assignment ----
  //
  const stepTaskAssignmentAndReview = ui.createMenu('⚙️ Step "Task Assignment/Review" – Task Assignment(~1 or 2 weeks before the new quarter)')
    .addItem(`${i++}. New Quarter's Task Assignment Preparation`, 'prepareTaskAssignement')
    .addItem(`${i++}. Automatic A/V Task Assignment(Run this step only after Andy assign all the Speaker/MC/Song Leader)`, 'autoTaskAssignSidebar')
    .addItem(`${i++}. Update Dashboard for Quick Check`, 'buildPivotTable')
    .addItem(`${i++}. Send out Review Task Assignment Email`, 'quarterlyRMailMerge');

  
  // ---- STEP Task Review ----
  
  //const stepTaskReview = ui.createMenu('📈 Step "Task Review" – Send out Review Task Assignment Email(~1 weeks before the new quarter)')
  //  .addItem(`${i++}. Send out Review Task Assignment Email`, 'quarterlyRMailMerge');

  //
  // ---- STEP Go Live ----
  //
  //const stepGoLive = ui.createMenu('📈 Step "Go Live" – Worship Schedule Go Live(Before the 1st Tuesday of the new quarter)')
  //  .addItem(`${i++}. Go Live !!!`, 'goLive');

  //
  // ---- STEP optional ----
  //
  const optionalSteps = ui.createMenu('🔹 Step "Optional" - Optional steps for managing worship schedule')
    .addItem(`${i++}. (optional) Update All DropDowns`, 'updateAllDropDowns')
    .addItem(`${i++}. (optional) Update Dashboard`, 'buildPivotTable')
    .addItem(`${i++}. (optional) Scheduling Email Reminder day...`, 'emailReminderSidebar')
    .addItem(`${i++}. (optional) Scheduling Text Reminder day...`, 'textReminderSidebar')
    .addItem(`${i++}. (optional) Start A/V Schedule Change Monitoring`, 'addScheduleChangeTrigger')
    .addItem(`${i++}. (optional) Stop A/V Schedule Change Monitoring`, 'removeScheduleChangeTrigger')
    .addItem(`${i++}. (optional) Start LiveStream YouTube Stat Collection`, 'addGetYouTubeStatsTrigger')
    .addItem(`${i++}. (optional) Stop Livestream YouTube Stat Collection`, 'removeGetYouTubeStatsTrigger')
    .addItem(`${i++}. (optional) Remove All Triggers immediately`, 'removeAllTriggers');

  //
  // Add steps to main menu
  //
  mainMenu
    .addSubMenu(stepPreparation)
    .addSubMenu(stepTaskAssignmentAndReview)
    //.addSubMenu(stepTaskReview)
    .addSubMenu(optionalSteps)
    .addToUi();
}

function preparation() {

  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Confirmation Required',
    'This step will delete all pertinent data on this schedule sheet and prepare for the next quarter setup.\n\nDo you want to continue?',
    ui.ButtonSet.YES_NO
  );

  // Exit if user clicks No
  if (response !== ui.Button.YES) {
    ui.alert('Preparation cancelled.');
    return;
  }

  // Continue only if user clicked Yes
  buildHistoricalNormalizeSchedule(); // Save Old Worship Data
  cleanWorshipSchedule();             // Delete All Data in "Cantonese Worship Schedule"
  cleanYouTubeStat();                 // Delete All Data in "YouTube Stat"
  cleanUnavailableDates();           // Delete All Data + initialize checkboxes
  createUnavailableDatesTrigger();   // Start accepting updates in "Unavailable Dates"
  //run goLive function at the 1st monday of each quarter at 11am
  addGoLiveTrigger();
  
  //update the sheet permission to editor for everybody to update the "unavailable dates" tab for the 1st time
  //this permission will get remove on this function "removeUnavailableDatesTrigger"
  setAnyoneWithLinkToEditor();

  //enable the trigger for running the UpdateAllDropDowns function everyday at 12:30am
  toggleTriggerUpdateAllDropDowns("enable")

  ui.alert('Preparation completed successfully.');
}

function prepareTaskAssignement() {
  removeUnavailableDatesTrigger(); //Stop accepting updates in "Unavailable Dates"
  //running this function here is signify the start of the working on the new schedule for the new quarter !!!
  //probably the last week of the old quarter
  //update all the dropdowns upon the conclusion of updating the "unavailable dates" tab possibly a week before the new quarter
  updateAllDropDowns()

  //update the sheet permission to view only for everybody upon the conclusion of updating the "unavailable dates" tab possibly a week before the new quarter
  setAnyoneWithLinkToViewer();

  //update all the people that are allowed to edit this file, and they will receive an email notification for this.
  setEditorsTofile();

  //update all the people that are allowed to view this file, and they will receive an email notification for this.
  setViewersTofile();

  //remove the trigger for running the UpdateAllDropDowns function everyday at 12:30am
  toggleTriggerUpdateAllDropDowns("disable")
  
  //remove the all the triggers at the end of next quarter
  addEndOfQuarterCleanupTrigger();
  
}

function goLive() {
  addEmailReminderTuesdayTrigger(); //Schedule Email Reminder
  //textReminderSidebar(); //Schedule Text Reminder
  addScheduleChangeTrigger(); //Start A/V Schedule Change Monitoring
  addGetYouTubeStatsTrigger(); //Start YouTube Stat Collection
  //addQuarterlyCleanupTrigger(); //Schedule End of Quarter cleanup
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

