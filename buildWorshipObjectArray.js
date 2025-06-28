// The buildWorshipObjectArray function is updated to convert Date objects to string format.
function buildWorshipObjectArray(headerRow, data) {
  const dataObjects = {};
  
  data.forEach(row => {
    const worshipDate = row[0]; // Assuming row[0] is a Date object
    // Convert worshipDate object to "MM/DD/YYYY" string
    const formattedWorshipDate = Utilities.formatDate(new Date(worshipDate), Session.getScriptTimeZone(), 'MM/dd/yyyy');

    if (!dataObjects[formattedWorshipDate]) {
      // Create an object for each row using headerRow as keys
      dataObjects[formattedWorshipDate] = {};
      for (let i = 1; i < row.length; i++) {
        dataObjects[formattedWorshipDate][headerRow[i]] = row[i];
      }
    }
  });
  
  return dataObjects;
}