function isPraiseTeamSunday(selectedDate, worshipSchedule, praiseTeamLeaders) {
  // Step 1: Extract songLeader and pianist from the worshipSchedule
  const songLeaderFromSchedule = worshipSchedule[selectedDate]["Song Leader"];
  const pianistFromSchedule = worshipSchedule[selectedDate].Pianist;
  logMessage(getCallStackTrace() + ": Song leader \"%s\", Pianst \"%s\" from the worship schedule of date \"%s\"", songLeaderFromSchedule, pianistFromSchedule, selectedDate);

  // Step 2: Find songLeader and pianist in the praiseTeamLeader array
  // Find all Song Leaders from the praiseTeamLeaders
  // filter will return an array, find will return the 1st object it found
  //const songLeaderFromPraiseTeam = praiseTeamLeaders.filter(member => member.role === "Song Leader");
  const songLeaderFromPraiseTeam = praiseTeamLeaders.find(member => member.role === "Song Leader");
  const pianistFromPraiseTeam = praiseTeamLeaders.find(member => member.role === "Pianist");
  logMessage(getCallStackTrace() + ": Song leader \"%s\" and Pianst \"%s\" from the praiseTeamLeaders of date \"%s\"", JSON.stringify(songLeaderFromPraiseTeam), JSON.stringify(pianistFromPraiseTeam), selectedDate);

  // Step 3: Compare the values
  // Check if there are any song leaders and if they match the one from the schedule
  const isSongLeaderMatch = songLeaderFromPraiseTeam && songLeaderFromPraiseTeam.name.includes(songLeaderFromSchedule);
  const isPianistMatch = pianistFromPraiseTeam && pianistFromPraiseTeam.name.includes(pianistFromSchedule);
  logMessage(getCallStackTrace() + ": isSongLeaderMatch = %s, and isPianistMatch = %s ", isSongLeaderMatch, isPianistMatch );
  
  // Output the results of the comparison
  if (isSongLeaderMatch && isPianistMatch) {
    logMessage(getCallStackTrace() + ": This is the praise team sunday \"%s\"", selectedDate);
    return true
  } else {
    logMessage(getCallStackTrace() + ": This is NOT the praise team sunday \"%s\"", selectedDate);
    return false
  }
}
