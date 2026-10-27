/**********************************************************************
 * callDeepSeekRepair.gs
 *
 * Call DeepSeek to repair invalid AI-generated assignments.
 **********************************************************************/


function callDeepSeekRepair(prompt) {


  aiLog("Calling DeepSeek Repair API");
  aiLog("=== REPAIR PROMPT ===");
  aiLog(prompt);


  const payload = {

    model: deepSeekModel,

    temperature: 0,

    messages: [

      {
        role: "system",
        content:
          `
You are a schedule repair assistant.

Your task is NOT to create a new schedule.

Your task is ONLY to repair invalid assignments
identified by the validator.

Rules:

1. Keep all valid assignments unchanged.

2. Modify only assignments that violate constraints.

3. Never rebalance the schedule.

4. Never optimize workload.

5. Never change valid worker assignments.

6. Return only valid JSON.
`
      },


      {
        role: "user",
        content: prompt
      }

    ]

  };



  const response = UrlFetchApp.fetch(
    deepSeekAPIUrl,
    {
      method: "post",

      contentType:
        "application/json",

      headers: {

        Authorization:
          "Bearer " + deepSeekAPIKey

      },

      payload:
        JSON.stringify(payload),

      muteHttpExceptions:
        true

    }
  );



  const responseCode =
    response.getResponseCode();


  const responseText =
    response.getContentText();



  aiLog(
    "DeepSeek Repair HTTP Status: " +
    responseCode
  );


  if (responseCode !== 200) {

    throw new Error(
      "DeepSeek Repair API failed:\n" +
      responseText
    );

  }



  const json =
    JSON.parse(responseText);



  if (
    !json.choices ||
    !json.choices[0] ||
    !json.choices[0].message
  ) {

    throw new Error(
      "Invalid DeepSeek Repair response:\n" +
      responseText
    );

  }



  const aiText =
    json.choices[0].message.content;



  aiLog("Parsing DeepSeek Repair response");



  return aiParser(aiText);

}
