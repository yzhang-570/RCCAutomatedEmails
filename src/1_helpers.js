/**
 * Returns whether the current user is in the allowed users list.
 *
 * @param {{user?: {email: string}}} e - The event object from the triggering menu action
 * @param {string[]} allowedUsers - Email addresses permitted to run the tool.
 */
function authorizeUser_(e, allowedUsers) {
  var userEmail = Session.getActiveUser().getEmail();
  
  // Fallback check: try getting the email from the event object if available
  if (!userEmail && e && e.user) {
    userEmail = e.user.email;
  }

  // Check if user is an authorized user
  const isAuthorized = allowedUsers.find(email => 
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
 * Applies lowercase, replaces whitespace with single space, and trims 
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
 * Internal helper: Returns whether an array (headers) contains a target string (targetName = required header column)
 * 
 * @param {string[] | null} headers
 * @param {string} targetName
 * @returns {number | null}
 */
function findRequiredHeaderCol_(headers, targetName) {
  if(!headers || !targetName) return null;
  const idx = headers.findIndex(header => (normalize_(header)).includes(normalize_(targetName.toLowerCase())));
  if (idx === -1) throw new Error(`Column "${targetName}" not found. (case-insensitive)`);
  return idx;
}

/**
 * Returns index of header row with given column names (case-insensitive)
 * from a specific sheet
 * Returns -1 if any target columns weren't found
 * 
 * @param {string[] | null} targetHeaders
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number | null}
 */
function findHeaderRowIndex_(targetHeaders, sheet) {
  if(!targetHeaders || !sheet) return null;
  const rowData = sheet.getDataRange().getValues();
  try {
    const idx = rowData.findIndex(row => (header.toLowerCase()).includes(targetName.toLowerCase()));
    if (idx === -1) throw new Error(`Column "${targetName}" not found. (case-insensitive)`);
    return idx;
  }
  catch (err) {
    console.log(`[findHeaderRowIndex] err - for ${sheet.getName()}`)
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