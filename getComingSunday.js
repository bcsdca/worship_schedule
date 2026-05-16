function getComingSunday() {
  const today = new Date();

  const offset = today.getDay() === 0 ? 0 : (7 - today.getDay());
  const comingSunday = new Date(today);
  comingSunday.setDate(today.getDate() + offset);

  const comingSundayStr = Utilities.formatDate(
    comingSunday,
    Session.getScriptTimeZone(),
    "MM/dd/yyyy"
  );

  logMessage(getCallStackTrace() + `: Coming Sunday = ${comingSundayStr}`);

  return comingSundayStr;
}
