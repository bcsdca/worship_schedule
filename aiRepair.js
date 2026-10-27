
/**
 * ***************************************************************
 * aiRepair.gs
 *
 * Repair invalid AI-generated worship schedule assignments.
 *
 * The AI assignment engine now processes ONE Sunday at a time.
 * Therefore scheduleRows normally contains exactly ONE row.
 *
 * Repair rules:
 *
 *   - Do NOT regenerate the schedule.
 *   - Do NOT change valid assignments.
 *   - ONLY repair assignments identified by validation errors.
 *   - Keep date and columnName unchanged.
 *   - Replace only the invalid worker.
 *   - Worker must be in the task's current options.
 *   - Worker must satisfy the current Sunday's blocking rules.
 * ***************************************************************
 */

function aiRepair(
  aiTool,
  result,
  scheduleRows,
  options = {},
  validationErrors = []
) {

  aiLog(
    "Starting AI assignment repair"
  );

  const prompt =
    buildRepairPrompt(
      result.assignments || [],
      validationErrors,
      scheduleRows
    );

  let repairedResult;

  switch (
    (aiTool || "").toLowerCase()
  ) {

    case "deepseek":

      repairedResult =
        callDeepSeekRepair(
          prompt,
          options
        );

      break;

    case "claude":

      repairedResult =
        callClaudeRepair(
          prompt,
          options
        );

      break;

    case "chatgpt":

      repairedResult =
        callChatGPTRepair(
          prompt,
          options
        );

      break;

    case "gemini":

      repairedResult =
        callGeminiRepair(
          prompt,
          options
        );

      break;

    default:

      throw new Error(
        "Unknown AI Tool for repair: " +
        aiTool
      );
  }

  if (
    !repairedResult ||
    !Array.isArray(
      repairedResult.assignments
    )
  ) {

    throw new Error(
      "AI repair returned an invalid result"
    );
  }

  aiLog(
    "Repair returned " +
    repairedResult.assignments.length +
    " assignments"
  );

  return repairedResult;
}


/**
 * ***************************************************************
 * Build repair prompt.
 *
 * The repair AI receives:
 *
 *   - original assignments
 *   - validation errors
 *   - current Sunday's schedule data
 *
 * It must repair ONLY invalid assignments.
 * ***************************************************************
 */

function buildRepairPrompt(
  assignments,
  validationErrors,
  scheduleRows
) {

  return `
You are repairing an AI-generated worship schedule.

The schedule contains assignments for ONE Sunday only.

The original assignments are mostly correct.

DO NOT regenerate the schedule.

DO NOT change assignments that are already valid.

ONLY repair assignments identified by the validation errors.

VALIDATION ERRORS:
${JSON.stringify(
  validationErrors,
  null,
  2
)}

ORIGINAL AI ASSIGNMENTS:
${JSON.stringify(
  assignments,
  null,
  2
)}

CURRENT SUNDAY SCHEDULE DATA:
${JSON.stringify(
  scheduleRows,
  null,
  2
)}

For each invalid assignment:

1. Keep the exact same:
   - date
   - columnName

2. Replace ONLY the worker.

3. The replacement worker MUST:
   - exist in that task's options array
   - satisfy the availability rules already reflected in options
   - satisfy the blocking rule
   - not create a duplicate assignment for the same task
   - not create another blocking assignment for a worker already
     occupying a blocking role on this Sunday

4. Do NOT modify any assignment that did not fail validation.

5. Do NOT remove valid assignments.

6. Do NOT add assignments for tasks that were not in the
   original assignments.

7. historicalCount and futureCount are informational only.
   They must NOT be used to override the task options or
   blocking rules.

8. Return the COMPLETE assignments array, including:
   - all original valid assignments unchanged
   - repaired assignments with only the worker changed

Return ONLY valid JSON.

No markdown.

No explanations.

{
  "summary": [
    "Fixed invalid assignments"
  ],
  "assignments": [
    {
      "date": "8/9/2026",
      "columnName": "PowerPoint Control",
      "worker": "James Tam",
      "reason": "Previous worker violated blocking rule"
    }
  ]
}
`;
}


