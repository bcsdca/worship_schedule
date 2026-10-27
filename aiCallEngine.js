/**
 * ***************************************************************
 * aiCallEngine.gs
 *
 * Route ONE AI assignment request to the selected provider.
 *
 * runAIAssign() calls this function for ONE Sunday at a time.
 *
 * scheduleRows therefore normally contains exactly ONE row.
 *
 * Flow:
 *
 *   runAIAssign()
 *        ↓
 *   aiCallEngine()
 *        ↓
 *   selected AI provider
 *        ↓
 *   validate
 *        ↓
 *   repair if necessary
 *        ↓
 *   return validated assignments
 * ***************************************************************
 */

function aiCallEngine(
  aiTool,
  scheduleRows,
  options = {}
) {

  let result;

  /**************************************************************
   * Step 1 — Generate AI assignments
   **************************************************************/
  switch ((aiTool || "").toLowerCase()) {

    case "claude":
      result =
        callClaudeAssignments(
          scheduleRows,
          options
        );
      break;

    case "chatgpt":
      result =
        callChatGPTAssignments(
          scheduleRows,
          options
        );
      break;

    case "gemini":
      result =
        callGeminiAssignments(
          scheduleRows,
          options
        );
      break;

    case "deepseek":
      result =
        callDeepSeekAssignments(
          scheduleRows,
          options
        );
      break;

    default:
      throw new Error(
        "Unknown AI Tool: " + aiTool
      );
  }

  /**************************************************************
   * Step 2 — Validate AI assignments
   *
   * Because scheduleRows contains only the current Sunday,
   * validation is also limited to the current Sunday.
   **************************************************************/
  try {

    aiValidator(
      result.assignments,
      scheduleRows
    );

    aiLog(
      "AI assignments passed validation"
    );

    return result;

  } catch (error) {

    aiLog(
      "AI validation failed:"
    );

    aiLog(
      error.validationErrors
        ? error.validationErrors.join("\n")
        : error.message
    );

    /************************************************************
     * Step 3 — Repair invalid assignments
     ************************************************************/

    aiLog(
      "Attempting AI assignment repair"
    );

    const repairedResult =
      aiRepair(
        aiTool,
        result,
        scheduleRows,
        options,
        error.validationErrors ||
          [error.message]
      );

    /************************************************************
     * Step 4 — Validate repaired result
     ************************************************************/

    aiValidator(
      repairedResult.assignments,
      scheduleRows
    );

    aiLog(
      "AI repair successful"
    );

    return repairedResult;
  }
}
