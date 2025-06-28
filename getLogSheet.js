function getLogSheet() {
  const ss = SpreadsheetApp.getActive();
  //const sheet0 = ss.getSheets()[0];
  //ss.setActiveSheet(sheet0); // keep the view at the 1st sheet

  let sheet = ss.getSheetByName("runLog");
    
  if (!sheet) {
    // Insert runLog at the last position
    const sheetCount = ss.getSheets().length; // Get total number of sheets
    sheet = ss.insertSheet("runLog", sheetCount); // Insert at the last index

    // Set the headers
    sheet.appendRow(["Timestamp", "Console.log/Console.error"]);

    // Format the header row: Bold, Calibri font, Font size 12
    const headerRange = sheet.getRange(1, 1, 1, 2); // First row, two columns
    headerRange.setFontWeight("bold");
    headerRange.setFontFamily("Calibri");
    headerRange.setFontSize(12);
    headerRange.setHorizontalAlignment("left"); // Left justify the text     
  }
  
  //ss.setActiveSheet(sheet); // keep the view to the runLog sheet now
  return sheet;
}
