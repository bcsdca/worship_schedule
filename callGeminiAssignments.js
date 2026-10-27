/**********************************************************************
 * callGeminiAssignments.gs
 *
 * Gemini API assignment provider.
 **********************************************************************/

function callGeminiAssignments(scheduleRows, options) {

  options = options || {};

  const temperature =
    options.temperature !== undefined
      ? options.temperature
      : 0;

  aiLog(`Calling Gemini API`);


  const payload = {
    system_instruction: {
      parts: [
        {
          text: aiPrompt()
        }
      ]
    },
    contents: [
      {
        role: "user",
        parts: [
          {
            text: JSON.stringify(
              scheduleRows,
              null,
              2
            )
          }
        ]
      }
    ],
    generationConfig: {
      temperature: temperature,
      responseMimeType: "application/json"
    }
  };
  
  
  const url = `${geminiAPIUrl}?key=${geminiAPIKey}`;
  
  const fetchOptions = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };


  const response = UrlFetchApp.fetch(
    url,
    fetchOptions
  );


  const data =
    JSON.parse(
      response.getContentText()
    );


  if (data.error) {
    throw new Error(
      `Gemini API error: ${JSON.stringify(data.error)}`
    );
  }


  const textOut =
    data.candidates[0]
      .content
      .parts[0]
      .text;


  return aiParser(textOut);

}