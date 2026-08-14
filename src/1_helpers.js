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

// Strict param -> return type pairing
/**
 * @overload @param {string} headers @returns {string}
 * @overload @param {string[]} headers @returns {string[]}
 */
/**
 * Applies lowercase, replaces whitespace with single space, and trims.
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
    if (idx === -1) console.log(`Header row: ${targetHeaders} not found.`);
    return idx;
  }
  catch (err) {
    console.log(`[findHeaderRowIndex] error while processing ${sourceSheet.getName()}.`)
    return -1;
  }
}










// /**
//  * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
//  * @param {string} name
//  * @returns {GoogleAppsScript.Spreadsheet.Sheet}
//  */
// function getOrCreateSheet_(ss, name) {
//   return ss.getSheetByName(name) || ss.insertSheet(name);
// }