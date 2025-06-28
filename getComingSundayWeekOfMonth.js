//this function is used on both slide_src and worship reminder project
function getComingSundayWeekOfMonth() {
    const functionName = getComingSundayWeekOfMonth.name;
    // Get the current date
    var today = new Date();

    // Calculate the number of days to add to get the coming Sunday
    var comingSundayOffset = 7 - today.getDay();

    // Get the date of the coming Sunday
    var comingSunday = new Date(today);
    comingSunday.setDate(today.getDate() + comingSundayOffset);

    // Get the day of the month for the coming Sunday
    const date = comingSunday.getDate();

    // Calculate the week of the month for the coming Sunday
    // Weeks start on Sunday, so we divide the date by 7 and round up
    var comingSundayWeekOfMonth = Math.ceil(date / 7);

    logMessage(getCallStackTrace() + `: Today = ${today}`);
    logMessage(getCallStackTrace() + `: Coming Sunday = ${comingSunday}`);
    logMessage(getCallStackTrace() + `: Coming Sunday's day of the month = ${date}`);
    logMessage(getCallStackTrace() + `: Returning the Coming Sunday's Week of Month = ${comingSundayWeekOfMonth}`);
    
    return comingSundayWeekOfMonth;
}
