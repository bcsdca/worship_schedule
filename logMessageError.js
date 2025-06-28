function logMessageError(message) { 
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  console.log(message);
  // Combine "ERROR" and message
  const errorMessage = `🚨 ERROR: ${message}`;
  
  logBuffer.push([timestamp, errorMessage]); // Store in memory
  
}