/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 * @see https://developers.google.com/apps-script/guides/triggers#onopene
 */

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // Create the main custom menu
  //const mainMenu = ui.createMenu('Cantonese Worship Service Utilities');
  //const mainMenu = ui.createMenu('✨ Cantonese Worship Service Utilities⚡');
  //const mainMenu = ui.createMenu('💥 Cantonese Worship Service Utilities 💥');
  const mainMenu = ui.createMenu('🎉 Cantonese Worship Service Utilities 🎉');

  // Submenu: "Clean up old copied over worship schedule data and handling "Unavailable Date to Serve
  const PrepNewQuarterMenu = ui.createMenu('Clean up old worship schedule data and handling "Unavailable Date to Serve"')
    .addItem('1. Delete All Data in "Cantonese Worship Schedule" Sheet', 'cleanWorshipSchedule')
    .addItem('2. Delete All Data in "YouTube Stat"', 'cleanYouTubeStat')
    .addItem('3. Delete All Data and initialize check box and enable email protection in "Unavailable Dates" Sheet', 'cleanUnavailableDates')
    .addItem('4. Delete All Data in "Unavailable Dates Response" Sheet', 'cleanUnavailableDatesResponse')
    .addItem('5. Send "Unavailable Date to Serve" Email', 'sendUnavailableDatesEmail')
    .addItem('6. Enable Monitoring the change in "Unavailable Dates" sheet', 'createUnavailableDatesTrigger')
    .addItem('7. Disable Monitoring the change in "Unavailable Dates" sheet', 'removeUnavailableDatesTrigger')
    .addItem('8. Update All Drop Downs', 'updateAllDropDowns');

  // Submenu: Preliminary Auto Task assignement
  const autoTaskAssignMenu = ui.createMenu('Preliminary Automatic Task assignement')
    .addItem('9. Automatic Task Assignment for Co-workers', 'autoTaskAssignSidebar')
    .addItem('10. Send Preliminary Worship Schedule for Review', 'quarterlyRMailMerge')

  // Submenu: Dashboard Data Management
  const dashBoardMenu = ui.createMenu('Dashboard Management')
    .addItem('11. Update Cantonese Worship Dashboard', 'buildPivotTable')

  // Submenu: Reminder Settings
  const reminderMenu = ui.createMenu('Email/Text Worship Reminder Settings')
    .addItem('12. Scheduling Email Reminder day...', 'emailReminderSidebar')
    .addItem('13. Scheduling Text Reminder day...', 'textReminderSidebar');

  // Submenu: Schedule Change Monitoring
  const scheduleChangeMenu = ui.createMenu('Worship Schedule Change Monitoring')
    .addItem('14. Start A/V Schedule Change Monitoring', 'addScheduleChangeTrigger')
    .addItem('15. Stop A/V Schedule Change Monitoring', 'removeScheduleChangeTrigger');

  // Submenu: YouTube Stat Collection
  const youtubeMenu = ui.createMenu('Weekly Worship YouTube Stat Collection')
    .addItem('16. Start YouTube Stat Collection', 'addGetYouTubeStatsTrigger')
    .addItem('17. Stop YouTube Stat Collection', 'removeGetYouTubeStatsTrigger')

  // Submenu: Remove All Triggers Eue To End Of Quarter
  const removeAllriggersMenu = ui.createMenu('Remove All Triggers Eue To End Of Quarter')
    .addItem('18. Remove All Triggers for End of Quarter', 'removeAllTriggers');

  // Add all submenus to the main menu
  mainMenu
    .addSubMenu(PrepNewQuarterMenu)
    .addSubMenu(autoTaskAssignMenu)
    .addSubMenu(dashBoardMenu)
    .addSubMenu(reminderMenu)
    .addSubMenu(scheduleChangeMenu)
    .addSubMenu(youtubeMenu)
    .addSubMenu(removeAllriggersMenu)
    .addToUi();
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

