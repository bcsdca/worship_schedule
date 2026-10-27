/**********************************************************************
 * callDeepSeekAssignments.gs
 *
 * DeepSeek API assignment provider.
 **********************************************************************/

function callDeepSeekAssignments(scheduleRows, options) {

  options = options || {};

  const temperature =
    options.temperature !== undefined
      ? options.temperature
      : 0;

  aiLog("Calling DeepSeek API");


  const payload = {

    model: deepSeekModel,

    temperature: temperature,

    response_format: {
      type: "json_object"
    },

    messages: [
      {
        role: "system",
        content: aiPrompt()
      },
      {
        role: "user",
        content: JSON.stringify(
          scheduleRows,
          null,
          2
        )
      }
    ]

  };



  const response =
    UrlFetchApp.fetch(
      deepSeekAPIUrl,
      {
        method: "post",
        contentType: "application/json",
        headers: {
          "Authorization":
            "Bearer " + deepSeekAPIKey
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

  //this will log all the AI detail internal reasoning, which is good, but there is a lot of lines
  //which is good for debugging in the future..
  //const text = response.getContentText();
  //aiLog(`HTTP ${response.getResponseCode()}`);
  //logLargeText("=== RAW RESPONSE ===", text);

  aiLog(`HTTP ${response.getResponseCode()}`);
  
  const data =
    JSON.parse(
      response.getContentText()
    );

  //logLargeText(
  //  "=== AI JSON ===",
  //  data.choices[0].message.content
  //);

  if (data.error) {

    aiError(
      `DeepSeek API error: ${JSON.stringify(data.error)}`
    );

  }


  if (
    !data.choices ||
    !data.choices[0] ||
    !data.choices[0].message
  ) {

    aiError(
      "DeepSeek response missing choices/message"
    );

  }


  return aiParser(
    data.choices[0].message.content
  );

}
