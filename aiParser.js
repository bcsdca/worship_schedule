
/**
 * ***************************************************************
 * aiParser.gs
 *
 * Parse AI JSON response.
 *
 * The AI now receives ONE Sunday at a time, but the parser
 * remains independent of the number of schedule rows.
 *
 * Expected response:
 *
 * {
 *   "summary": [...],
 *   "decisionSummary": [...],
 *   "assignments": [
 *     {
 *       "date": "8/9/2026",
 *       "columnName": "Camera Person",
 *       "worker": "James Tam",
 *       "reason": "..."
 *     }
 *   ]
 * }
 * ***************************************************************
 */

function aiParser(textOut) {

  aiLog(
    "Parsing AI response"
  );

  if (isBlank(textOut)) {
    aiError(
      "AI response is empty."
    );
  }

  logLargeText(
    "=== RAW AI RESPONSE ===",
    textOut
  );

  const jsonText =
    extractJsonObject(textOut);

  const result =
    parseJson(jsonText);

  /**************************************************************
   * Validate top-level assignments array.
   **************************************************************/
  if (!Array.isArray(result.assignments)) {

    aiError(
      "AI JSON missing assignments array."
    );
  }

  aiLog(
    `Parsed ${result.assignments.length} assignments`
  );

  /**************************************************************
   * Log AI strategy.
   **************************************************************/
  if (
    Array.isArray(result.summary) &&
    result.summary.length
  ) {

    aiLog(
      "AI Strategy:"
    );

    result.summary.forEach(
      s => aiLog(`  • ${s}`)
    );
  }

  /**************************************************************
   * Log AI key decisions.
   **************************************************************/
  if (
    Array.isArray(result.decisionSummary) &&
    result.decisionSummary.length
  ) {

    aiLog(
      "AI Key Decisions:"
    );

    result.decisionSummary.forEach(
      s => aiLog(`  • ${s}`)
    );
  }

  /**************************************************************
   * Validate assignment structure.
   *
   * This is only a basic parser-level check.
   *
   * Business rules such as:
   *   - dropdown validity
   *   - blocking rules
   *   - duplicate assignments
   *
   * are handled by aiValidator().
   **************************************************************/
  result.assignments.forEach(
    (a, index) => {

      if (
        !a ||
        !a.date ||
        !a.columnName ||
        !a.worker
      ) {

        aiError(
          `Invalid assignment at index ${index}: ` +
          "date, columnName, and worker are required."
        );
      }

      aiLog(
        `${a.date} | ${a.columnName} -> ${a.worker}` +
        (
          a.reason
            ? ` (${a.reason})`
            : ""
        )
      );

    }
  );

  return result;
}


/**
 * ***************************************************************
 * Extract first complete JSON object.
 * ***************************************************************
 */

function extractJsonObject(text) {

  let cleaned =
    text.trim();

  /**************************************************************
   * Remove common markdown JSON wrappers.
   **************************************************************/
  cleaned =
    cleaned
      .replace(/^```json/i, "")
      .replace(/^```/, "")
      .replace(/```$/, "")
      .trim();

  const start =
    cleaned.indexOf("{");

  if (start < 0) {

    aiError(
      "No JSON object found."
    );
  }

  let depth = 0;

  for (
    let i = start;
    i < cleaned.length;
    i++
  ) {

    if (cleaned[i] === "{") {
      depth++;
    }

    if (cleaned[i] === "}") {
      depth--;
    }

    if (depth === 0) {

      return cleaned.substring(
        start,
        i + 1
      );
    }
  }

  aiError(
    "Incomplete JSON object."
  );
}

