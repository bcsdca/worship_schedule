/**********************************************************************
 * callChatGPTAssignments.gs
 *
 * ChatGPT API assignment provider.
 **********************************************************************/

function callChatGPTAssignments(scheduleRows, options) {

  options = options || {};

  const temperature =
    options.temperature !== undefined
      ? options.temperature
      : 0;

  aiLog(`Calling ChatGPT API`);

  const payload = {
    model: openaiModel,
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


  const fetchOptions = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + openaiAPIKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };


  const response = UrlFetchApp.fetch(
    openaiAPIUrl,
    fetchOptions
  );


  const data =
    JSON.parse(
      response.getContentText()
    );


  if (data.error) {
    throw new Error(
      `ChatGPT API error: ${JSON.stringify(data.error)}`
    );
  }


  const textOut =
    data.choices[0].message.content;


  return aiParser(textOut);

}