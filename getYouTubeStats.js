function getYouTubeStats(isFirstTime = true) {

  clearLogSheet();

  //check to see isFirstTime is not a boolean (e.g., it's an event object because of the time trigger), force it to true
  if (typeof isFirstTime !== 'boolean') {
    isFirstTime = true;
    logMessage(`${getCallStackTrace()}: This function was called by a time trigger, so force isFirstTime = ${isFirstTime}`)
  }

  try {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy-HH:mm:ss');
    var videoId = "";

    logMessage(`${getCallStackTrace()}: This function was run as isFirstTime = ${isFirstTime}`)

    if (isFirstTime) {
      videoId = getUpcomingLiveStreamVideoId();
      logMessage(getCallStackTrace() + `: Found upcoming live stream videoId: ${videoId} on ${today}`);
    } else {
      videoId = PropertiesService.getScriptProperties().getProperty("LAST_STREAM_ID");
      logMessage(getCallStackTrace() + `: Found the saved live stream videoId: ${videoId} on ${today}`);
    }

    if (!videoId) {
      logMessage(getCallStackTrace() + ": ❌ No valid video ID found, exiting function.");
      return;
    }

    var part = ['snippet', 'statistics', 'liveStreamingDetails', 'recordingDetails', 'status', 'contentDetails', 'localizations', 'player'];
    var params = { 'id': videoId };

    var response = YouTube.Videos.list(part, params);

    if (!response || !response.items || response.items.length === 0) {
      logMessage(getCallStackTrace() + ": ❌ No video details found, exiting.");
      return;
    }

    var video = response.items[0];

    logMessage(getCallStackTrace() + ": Video Object= " + JSON.stringify(video, null, 2));

    var videoViewCount = video.statistics?.viewCount || 0;
    var videoConcurrentViewers = video.liveStreamingDetails?.concurrentViewers || 0;
    var videoPrivacyStatus = video.status?.privacyStatus || "unknown";

    var actualStartTime = video.liveStreamingDetails?.actualStartTime
      ? Utilities.formatDate(new Date(video.liveStreamingDetails.actualStartTime), Session.getScriptTimeZone(), 'MM/dd/yyyy-HH:mm:ss')
      : "";

    var actualEndTime = video.liveStreamingDetails?.actualEndTime
      ? Utilities.formatDate(new Date(video.liveStreamingDetails.actualEndTime), Session.getScriptTimeZone(), 'MM/dd/yyyy-HH:mm:ss')
      : "";

    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName("YouTube Stat");
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    logMessage(getCallStackTrace() + `: lastRow: ${lastRow}, lastColumn: ${lastColumn}`);

    var currDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy-HH:mm:ss');
    var title = video.snippet?.title || "Unknown Title";
    var lbc = video.snippet?.liveBroadcastContent || "unknown";
    var id = video.id;

    // Adjust `actualStartTime` and `actualEndTime` based on `liveBroadcastContent`
    if (lbc === "upcoming") {
      actualStartTime = "";
      actualEndTime = "";
    } else if (lbc === "live") {
      actualEndTime = "";
    }

    var youtubeStatDisplay = [id, title, lbc, actualStartTime, currDate, videoConcurrentViewers, videoViewCount, actualEndTime, videoPrivacyStatus];

    if (lbc === "upcoming") {
      logMessage(getCallStackTrace() + `: Live Broadcast Content is "upcoming". Handling accordingly...`);

      if (isFirstTime || lastRow === 1) {
        lastRow++;
        logMessage(getCallStackTrace() + `: First time or unexpected lastRow=1 condition. Updating lastRow to ${lastRow}`);
      }

      sheet.getRange(lastRow, 1, 1, lastColumn).setValues([youtubeStatDisplay]);
      create_trigger_5m();
    } else if (lbc === "live") {
      logMessage(getCallStackTrace() + `: Live Broadcast Content is "live". Appending to last row ${lastRow}`);
      sheet.appendRow(youtubeStatDisplay);
      create_trigger_5m();
    } else if (lbc === "none") {
      logMessage(getCallStackTrace() + `: Live Broadcast Content is "none". Appending to last row ${lastRow} and deleting trigger.`);
      sheet.appendRow(youtubeStatDisplay);
      delete_trigger_5m();
    } else {
      logMessage(getCallStackTrace() + `: ⚠️ Unsupported liveBroadcastContent: "${lbc}". Appending to last row ${lastRow} and deleting trigger.`);
      sheet.appendRow(youtubeStatDisplay);
      delete_trigger_5m();
    }
  } catch (error) {
    logMessage(getCallStackTrace() + `: ❌ Error occurred - ${error.message}`);
  }
  flushLogsToSheet();
}

function create_trigger_5m() {
  //Create new trigger
  //check to make sure there is no trigger_5m already existed
  var find_it = false
  var oldTrigger = ScriptApp.getScriptTriggers()
  logMessage(getCallStackTrace() + ": The below triggers are the current running triggers !!!");
  //Logger.log(oldTrigger.length);
  for (var i = 0; i < oldTrigger.length; i++) {
    logMessage(getCallStackTrace() + ": Current running trigger is " + ScriptApp.getScriptTriggers()[i].getHandlerFunction());
    if (ScriptApp.getScriptTriggers()[i].getHandlerFunction() == "trigger_5m") {
      find_it = true;
      break;
    }
  }

  if (!find_it) {
    ScriptApp.newTrigger('trigger_5m').timeBased().everyMinutes(5).create();
    logMessage(getCallStackTrace() + ": " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH:mm:ss') + ': No existing trigger_5m, so creating trigger_5m !!!');
  } else {
    logMessage(getCallStackTrace() + ": " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH:mm:ss') + ': Found an existing trigger_5m, NOT creating trigger_5m !!!');
  }

}

function delete_trigger_5m() {
  let remove_array = []
  var oldTrigger = ScriptApp.getScriptTriggers()
  logMessage(getCallStackTrace() + ": The below triggers are the current running triggers !!!");
  //Logger.log(oldTrigger.length);
  for (var i = 0; i < oldTrigger.length; i++) {
    logMessage(getCallStackTrace() + ": Current running trigger is " + ScriptApp.getScriptTriggers()[i].getHandlerFunction());
    if (ScriptApp.getScriptTriggers()[i].getHandlerFunction() == "trigger_5m") {
      remove_array.push(oldTrigger[i]);
    }
  }

  remove_array.forEach(function (row) {
    ScriptApp.deleteTrigger(row);
    logMessage(getCallStackTrace() + ": " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH:mm:ss') + ': Deleting 5 min trigger !!!');

  });
}

function trigger_5m() {
  logMessage(getCallStackTrace() + ": Running getYouTubeStats after 5 min !!!")
  getYouTubeStats(false)
}

function getUpcomingLiveStreamVideoId() {
  try {
    //cecYouTubeChannelID and youTubeAPIKey are defined in globalVarWorshipSchedule
    const searchUrl = `https://www.googleapis.com/youtube/v3/search?part=snippet&channelId=${cecYouTubeChannelID}&type=video&eventType=upcoming&maxResults=10&key=${youTubeAPIKey}`;

    const searchResponse = UrlFetchApp.fetch(searchUrl, { muteHttpExceptions: true });
    if (searchResponse.getResponseCode() !== 200) {
      logMessageError(getCallStackTrace() + `: ❌ API Error (search): ${searchResponse.getContentText()}`);
      return null;
    }

    const searchData = JSON.parse(searchResponse.getContentText());
    logMessage(getCallStackTrace() + `: 🔍 Raw search results = ${JSON.stringify(searchData.items, null, 2)}`);
    const videoIds = (searchData.items || []).map(item => item.id.videoId).filter(Boolean);

    if (videoIds.length === 0) {
      logMessage(getCallStackTrace() + `: 🚫 No video IDs found.`);
      return null;
    }

    // Include contentDetails to get stream ID
    const detailsUrl = `https://www.googleapis.com/youtube/v3/videos?part=liveStreamingDetails,snippet,contentDetails&id=${videoIds.join(",")}&key=${youTubeAPIKey}`;
    const detailsResponse = UrlFetchApp.fetch(detailsUrl);
    const detailsData = JSON.parse(detailsResponse.getContentText());
    const videoDetails = detailsData.items || [];

    // Sort by scheduled start time
    const sortedStreams = videoDetails
      .filter(item => item.liveStreamingDetails && item.liveStreamingDetails.scheduledStartTime)
      .sort((a, b) =>
        new Date(a.liveStreamingDetails.scheduledStartTime) - new Date(b.liveStreamingDetails.scheduledStartTime)
      );

    // Only look at the 2 closest upcoming livestreams, exclude "english"
    const topTwo = sortedStreams
      .filter(item => {
        const title = item.snippet.title.toLowerCase();
        return !title.includes("english") && !title.includes("mandarin");
      })
      .slice(0, 2);

    logMessage(getCallStackTrace() + `: top 2 streams = ${JSON.stringify(topTwo, null, 2)}`);

    const keywords = ["cantonese", "combine", "join"];

    let selected = null;
    let matchKeyword = null;

    for (const item of topTwo) {
      const title = item.snippet.title.toLowerCase();
      matchKeyword = keywords.find(kw => title.includes(kw));
      if (matchKeyword) {
        logMessage(getCallStackTrace() + `: ✅ Matched keyword "${matchKeyword}" in title "${item.snippet.title}"`);
        selected = item;
        break;
      }
    }

    if (!selected) {
      logMessage(getCallStackTrace() + `: 🚫 No "Cantonese" nor "combine" nor "join" livestream found in the closest 2 upcoming streams`);
      return null;
    }

    const videoId = selected.id;
    const videoTitle = selected.snippet.title;

    // ✅ Get bound stream ID and log it
    const streamId = selected.contentDetails?.boundStreamId;
    if (streamId) {
      logMessage(getCallStackTrace() + `: 🔗 Bound Stream ID: ${streamId}`);
    } else {
      logMessage(getCallStackTrace() + `: ⚠️ No bound Stream ID found.`);
    }

    PropertiesService.getScriptProperties().setProperty("LAST_STREAM_ID", videoId);
    const savedId = PropertiesService.getScriptProperties().getProperty("LAST_STREAM_ID");

    if (savedId !== videoId) {
      logMessageError(getCallStackTrace() + `: ⚠️ Error: Stored ID mismatch (${savedId} vs ${videoId})`);
      return null;
    }

    logMessage(getCallStackTrace() + `: 🎯 Selected livestream: ${videoTitle} (Video ID: ${videoId}), (Stream ID: ${streamId})`);
    return videoId;

  } catch (error) {
    logMessageError(getCallStackTrace() + `: ❌ Error: ${error.message}`);
    return null;
  }
}



