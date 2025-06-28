function setAddOnFormSubmitTrigger(option) {


  // removing all old change email reminder triggers 1st
  let remove_array = []
  var oldTrigger = ScriptApp.getScriptTriggers()
  //Logger.log(oldTrigger.length);
  Logger.log("The below triggers are the current running triggers !!!");
  for (var i = 0; i < oldTrigger.length; i++) {
    Logger.log(ScriptApp.getScriptTriggers()[i].getHandlerFunction());
    if (ScriptApp.getScriptTriggers()[i].getHandlerFunction() == "onFormSubmit") {
      remove_array.push(oldTrigger[i]);

    }
  }
  remove_array.forEach(function (row) {
    //Logger.log(row);
    ScriptApp.deleteTrigger(row);
    Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH:mm:ss') + ': Deleting the old \"on Form Submit\" trigger ' + row + ' !!!');

  });

  if (option == "enable") {
    ScriptApp.newTrigger("onFormSubmit")
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onFormSubmit()
      .create();

    Logger.log("The new \"on Form Submit\" trigger was just created !!!",);
    SpreadsheetApp.getActive().toast("The new \"on Form Submit\" trigger was just created 👍 !!!");
  } else {
    Logger.log("The \"on Form Submit\" trigger was just removed !!!",);
    SpreadsheetApp.getActive().toast("The \"on Form Submit\" trigger was just removed 👍 !!!");
  }
}
