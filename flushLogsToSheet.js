function flushLogsToSheet() {
  if (logBuffer.length === 0) return; // Nothing to write

  const logSheet = getLogSheet(); // Ensure sheet exists
  const lastRow = logSheet.getLastRow(); // Find the last row with content
  const numRows = logBuffer.length;

  // Append logs in bulk
  const logRange = logSheet.getRange(lastRow + 1, 1, numRows, 2);
  logRange.setValues(logBuffer);

  formatLogRange(logRange);
  
  // Auto-resize columns to fit content
  logSheet.autoResizeColumns(1, 2);

  logBuffer = []; // Clear log buffer after writing
  //const sheet0 = ss.getSheets()[0];
  //ss.setActiveSheet(sheet0); // keep the view at the 1st sheet
}

function formatLogRange(logRange) {
  
  // Apply text formatting
  logRange.setFontFamily("Calibri");
  logRange.setFontSize(11);
  logRange.setHorizontalAlignment("left"); // Left justify the text
  logRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);  // Prevent text wrapping

  // Remove bold formatting
  logRange.setFontWeight("normal");

  // Apply text color formatting only for errors
  const fontColors = logBuffer.map(row => row[1].toLowerCase().includes("error") ? ["red", "red"] : [null, null]); 
  logRange.setFontColors(fontColors);

}
