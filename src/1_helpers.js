/*
Improvements (if time):

i.e. user-proofing LMAO

- (fixed in fix A) **Column order:** Always read the sheet's column ordering (indexes) directly, instead of 
  relying on config's ordering - may be mismatched; i.e. remove use of mapValueToHeaders
  across code
  - verify presence of expected headers + return index map of actual headers -> for writes
  - accomplishes: validate schema (safety) + physical position (flexibility)
- **Adding new rows:**Refactor to always use findLastDataRowIndex instead of .appendRow to
  guarantee safe/expected appends to the immediate next row across code when there are
  trailing/hidden artifacts
     ex. a whitespace on row 800 or failed (blank) formula -> means next append goes to 801
     *skull* im gone

( fixed in fix A) developer-proofing:
- config headers match with sheet's (verified)
- BUT dev can still try to get a column with the wrong name - not checked against config
- will result in a unexplained fail (no column found)
*/



/**
 * --------------------------------------------------------------------------------------------
 * Shared Utilities
 * 
 * --------------------------------------------------------------------------------------------
 */

/**
 * Returns whether the current user is in the allowed users list.
 *
 * @param {{user?: {email: string}}} e - The event object from the triggering menu action
 */
function authorizeUser_(e) {

  // Get email of current user
  var userEmail = Session.getActiveUser().getEmail();
  // Fallback: get email from event object
  if (!userEmail && e && e.user) {
    userEmail = e.user.email;
  }

  // Check if user is an authorized user
  const isAuthorized = CONFIG.allowedUsers.find(email => 
    email.toLowerCase() === userEmail.toLowerCase()
  ) !== undefined // returns undefined if not found
  
  var ui = SpreadsheetApp.getUi();
  if (!isAuthorized) {
    ui.alert(
      "Access Denied", 
      "You are not authorized to run this tool.", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  };  

  Logger.log("[authorizeUser] Authorized access granted: " + userEmail);
  return true;
}

/**
 * Appends an entry to the shared processing log sheet (CONFIG.logSheet).
 * @param {Object<string, *>} data - entry details matching logSheet's headers (ex. "Type": "Error"|"Warning")
 * @returns {boolean} True if the entry was logged successfully
 */
function logIssue_(data) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.logSheet.name);
  if (!logSheet) return false;

  const entry = { "Timestamp": new Date(), ...data };
  logSheet.appendRow(mapValuesToHeaders_(entry, CONFIG.logSheet.headers));

  if (normalize_(data.Type) === normalize_("Error")) {
    logSheet.getRange(logSheet.getLastRow(), 1, 1, CONFIG.logSheet.headers.length).setBackground("#f4cccc");
  }
  return true;
}

/**
 * Returns whether a sheet exists and its header row contains the given headers.
 * Shows an alert and returns false if not.
 * 
 *
 * @param {string} sheetName
 * @param {string[]} headers
 * @returns {boolean}
 */
function isSheetReady_(sheetName, headers) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const isReady = !!sheet && findHeaderRowIndex_(headers, sheet) !== -1;

  if (!isReady) {
    SpreadsheetApp.getUi().alert(`⚠️Error: "${sheetName}" sheet not found or missing required columns:\n${headers}\n\nVerify that the header columns match exactly.`);
  }
  return isReady;
}



/**
 * --------------------------------------------------------------------------------------------
 * String-Processing Helpers
 * 
 * --------------------------------------------------------------------------------------------
 */

// Strict param -> return type pairing
/**
 * @overload @param {string} headers @returns {string}
 * @overload @param {string[]} headers @returns {string[]}
 */
/**
 * Replaces whitespace with single space, trims, and applies lowercase.
 * 
 * Accepts strings and string arrays.
 * 
 * @param {Array<*> | string[] | string} headers
 * @returns {string[] | string}
 */
function normalize_(headers) {
  if(!headers) return "";
  if(!Array.isArray(headers)) return String(headers || "").replace(/\s+/g, " ").trim().toLowerCase();
  return headers.map(header =>
    String(header || "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase()
  );
}

/**
 * Converts and formats a plain-text email body into its HTML equivalent
 * 
 * Supported:
 * - **bold** markdown-style markers
 * - {text|color} markers (ex. "{Active Member!|navy}"); Accepts named colors (ex. navy) and hexadecimals (ex. #FF5733)
 * - Dynamically width, instead of wrapping at fixed column widths (forced plain text formatting)
 *
 * @param {string} text
 * @returns {string}
 */
function plainTextToHtml_(text) {

  // Pre-processing: Escape special characters to display safely
  const escaped = String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  
  return escaped
    .replace(/\*\*(.+?)\*\*/g, "<b>$1</b>")
    .replace(/\{(.+?)\|(\w+|#[0-9A-Fa-f]{6})\}/g, '<span style="color:$2">$1</span>')
    .replace(/\n/g, "<br>");
}

/**
 * Extracts the first name from a full name string by returning the first word, if exists.
 *
 * @param {string} fullName
 * @returns {string}
 */
function getFirstName_(fullName) {
  if (!fullName) return "";
  return String(fullName).trim().split(/\s+/)[0];
}



/**
 * --------------------------------------------------------------------------------------------
 * Sheet-Processing Utilities
 * 
 * --------------------------------------------------------------------------------------------
 */

/**
 * Maps headers names to their indices in a sheet's header row.
 *
 * @param {string[]} targetHeaders - array of header names to map
 * @param {Array<*>} sourceRow - array of sheet's actual header row
 * @returns {Object<string, number>}
 */
function mapHeaderIndices_(targetHeaders, sourceRow) {
  const normRow = normalize_(sourceRow);
  /** @type {Object<string, number>} */
  const indexMap = {};
  targetHeaders.forEach(header => {
    indexMap[header] = normRow.findIndex(h => normalize_(h) === normalize_(header));
  });
  return indexMap;
}

/**
 * Returns index of header row with columns matching target header names
 * from a specific sheet (case-insensitive)
 * 
 * Returns -1 if no matching row found
 * 
 * @param {string[]} targetHeaders
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet
 * @returns {number | null}
 */
function findHeaderRowIndex_(targetHeaders, sourceSheet) {
  if(!targetHeaders || !sourceSheet) return null;
  const sheetData = sourceSheet.getDataRange().getValues();
  const normHeaders = normalize_(targetHeaders);
  try {
    const idx = sheetData.findIndex(row => {
      const normRow = normalize_(row);
      return normHeaders.every(header => normRow.includes(header));
    });
    if (idx === -1) console.log(`[findHeaderRowIndex] Header row: ${targetHeaders} not found.`);
    return idx;
  }
  catch (err) {
    console.log(`[findHeaderRowIndex] error while processing ${sourceSheet.getName()}.`)
    return -1;
  }
}

/**
 * Returns the index (0-based) of the last data row, stopping at the first fully-blank
 * row found after headerRow. 
 *
 * @param {Array<Array<*>>} sheetData - full getDataRange().getValues() result
 * @param {number} headerRow - 0-based index of the header row within sheetData
 * @returns {number} 0-based index of the last data row (equal to headerRow if there's no data)
 */
function findLastDataRowIndex_(sheetData, headerRow) {
  const firstBlankRow = sheetData.findIndex((row, idx) => idx > headerRow && row.every(cell => cell === "" || cell === null));
  return firstBlankRow === -1 ? sheetData.length - 1 : firstBlankRow - 1;
}

/**
 * Reads a sheet's data range and locates its header row, returning the parsed rows, a
 * column index map, and the last actual data row. Returns null if the header row can't
 * be found.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} headers
 * @returns {?{sheetData: Array<Array<*>>, colMap: Object<string, number>, headerRow: number, lastDataRow: number}}
 */
function readSheetRows_(sheet, headers) {
  const headerRow = findHeaderRowIndex_(headers, sheet);
  if (headerRow === null || headerRow === -1) return null;

  const sheetData = sheet.getDataRange().getValues();
  const colMap = mapHeaderIndices_(headers, sheetData[headerRow]);
  const lastDataRow = findLastDataRowIndex_(sheetData, headerRow);
  return { sheetData, colMap, headerRow, lastDataRow };
}

/**
 * Orders values to match ordering in `headers`.
 * 
 * Headers with no matching key are left blank.
 *
 * @param {Object<string, *>} values - values keyed by header name
 * @param {string[]} headers - target header order
 * @returns {Array<*>}
 */
function mapValuesToHeaders_(values, headers) {
  /** @type {Object<string, *>} */
  const normalizedValues = {};
  Object.keys(values).forEach(key => {
    normalizedValues[normalize_(key)] = values[key];
  });
  return headers.map(header => normalizedValues[normalize_(header)] ?? "");
}