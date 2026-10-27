/**********************************************************************
 * AI Assignment Framework V2
 *
 * Common helper functions shared by all AI providers.
 *
 * All logging goes through aiLog().
 **********************************************************************/

/**
 * Standard logging for the AI framework.
 */

function aiLog(message) {
  logMessage(getCallStackTrace() + `: ${message}`);
}

/**
 * Throws an Error after logging it.
 */
function aiError(message) {
  aiLog("ERROR: " + message);
  throw new Error(message);
}

/**
 * Safe JSON.parse()
 */
function parseJson(text) {

  try {
    return JSON.parse(text);
  }
  catch (e) {

    aiLog("JSON Parse Failed");
    aiLog(text);

    throw new Error(
      "Unable to parse JSON.\n\n" +
      e.message
    );
  }

}

/**
 * Pretty JSON for logging.
 */
function prettyJson(obj) {

  return JSON.stringify(
    obj,
    null,
    2
  );

}

/**
 * Log long text in chunks.
 */
function logLargeText(title, text) {

  aiLog(title);

  const chunkSize = 800;

  for (let i = 0; i < text.length; i += chunkSize) {

    aiLog(
      text.substring(
        i,
        i + chunkSize
      )
    );

  }

}

/**
 * Returns true if string is empty.
 */
function isBlank(value) {

  return value === null ||
    value === undefined ||
    value === "";

}

/**
 * Deep clone.
 */
function deepCopy(obj) {

  return JSON.parse(
    JSON.stringify(obj)
  );

}

/**
 * Returns true if worker already occupied.
 */
function isWorkerOccupied(worker, occupiedWorkers) {

  return occupiedWorkers.indexOf(worker) >= 0;

}

