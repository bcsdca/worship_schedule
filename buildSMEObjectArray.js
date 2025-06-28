/**
 * Builds an array of objects containing SME labels and contact names from the contact sheet.
 */
function buildSMEObjectArray(contactSheet, avSMEColumn) {

  //dont include these names in the drop down list
  //const excludedNames = ["Hilbert Chu", "Peggy Chu", "Winkie Zhang", "Wellington Hui"]; //move to global variable file for easily find and changed
  const lastRow = contactSheet.getLastRow();
  const objectArray = [];

  // Get all data in one call for performance
  const contactData = contactSheet.getRange(2, 1, lastRow - 1, avSMEColumn).getValues();

  contactData.forEach(row => {
    const contactName = row[0];
    const smeList = row[avSMEColumn - 1].split(",");
    //excludeNames was defined in the global variable file
    if (!excludedNames.includes(contactName)) {
      smeList.forEach(sme => {
        objectArray.push({
          smeLabel: sme.trim(),
          name: contactName
        });
      });
    }
  });

  return objectArray;
}