function logMessage(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  console.log(message);
  logBuffer.push([timestamp, message]); // Store in memory
    
}