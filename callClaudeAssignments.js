/**********************************************************************
 * callClaudeAssignments.gs
 *
 * Claude API assignment provider.
 **********************************************************************/

function callClaudeAssignments(scheduleRows, options) {

  options = options || {};

  const temperature =
    options.temperature !== undefined
      ? options.temperature
      : 0;


  aiLog(`Calling Claude API`);


  const payload = {
    model: ANTHROPIC_MODEL,
    max_tokens: 8000,
    temperature: temperature,
    system: aiPrompt(),
    messages: [
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
      "x-api-key": ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };


  const response = UrlFetchApp.fetch(
    "https://api.anthropic.com/v1/messages",
    fetchOptions
  );


  const data =
    JSON.parse(
      response.getContentText()
    );


  if (data.error) {

    throw new Error(
      `Claude API error: ${JSON.stringify(data.error)}`
    );

  }


  const textOut =
    data.content
      .map(x => x.text || "")
      .join("");


  return aiParser(textOut);

}
