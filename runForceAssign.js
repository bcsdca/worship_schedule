//wipe out all the assignments on the sheet
//taking care of all the force assignment
//write out all force assignments to the sheet
//returning the intiial historical count

function runForceAssign(sheet) {
  
  clearTaskAssignments(sheet);

  const forcedAssignments = resolveForcedAssignments(sheet);

  writeForcedAssignments(forcedAssignments, sheet);
  
  //here is good break between force assignment and the AI or GAS assignment
  //return;

  const bounds = getScheduleRowBounds(sheet);

  if (!bounds) {
    aiLog(`first row and last row to assign task were not found`);
    return;
  }

  const firstRow = bounds.firstRow;
  const lastRow = bounds.lastRow;
  //const lastRow = 4;

  aiLog(`Will using AI to assign tasks between row ${firstRow} and row ${lastRow}`);

  const historicalCount = {};

  /**************************************************************
   * Initial historical count from forced assignments
   **************************************************************/
  forcedAssignments.forEach(a => {

    if (!historicalCount[a.worker]) {
      historicalCount[a.worker] = 0;
    }

    historicalCount[a.worker]++;
  });


  /**************************************************************
   * Include Worship Chairperson and Song Leader history
   **************************************************************/
  const headerInfo = buildHeaderMap(sheet);
  const headerMap = headerInfo.headerMap;

  const worshipChairCol =
    headerMap["Worship Chairperson"];

  const songLeaderCol =
    headerMap["Song Leader"];

  for (let r = firstRow; r <= lastRow; r++) {

    const date =
      sheet.getRange(`A${r}`).getValue();

    if (!(date instanceof Date)) {
      continue;
    }

    const worshipChair =
      sheet.getRange(`${worshipChairCol}${r}`).getValue();

    if (worshipChair) {
      historicalCount[worshipChair] =
        (historicalCount[worshipChair] || 0) + 1;
    }

    const songLeader =
      sheet.getRange(`${songLeaderCol}${r}`).getValue();

    if (songLeader) {
      historicalCount[songLeader] =
        (historicalCount[songLeader] || 0) + 1;
    }
  }

  logLargeText(
    `=== INITIAL HISTORICAL COUNT ===`,
    JSON.stringify(
      historicalCount,
      null,
      2
    )
  );

  return historicalCount;
  
}




