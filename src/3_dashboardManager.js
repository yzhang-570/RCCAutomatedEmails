/**
 * Classifies an event tab by name into one of the 4 check-in types.
 *
 * @param {string} sheetName
 * @returns {"social"|"nonSocial"|"ambMeet"|"tabling"|null}
 */
function classifyEventType_(sheetName) {
  const name = normalize_(sheetName);
  if (name.includes(normalize_("Non-social"))) return "nonSocial"; // check first - "Social" is a substring of it
  if (name.includes(normalize_("Social"))) return "social";
  if (name.includes(normalize_("AmbMeet"))) return "ambMeet";
  if (name.includes(normalize_("Tabling"))) return "tabling";
  return null;
}

/**
 * Reads a roster-style sheet (including Full Name, SJSU Email) into an email -> name map.
 *
 * @param {?GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} headers - target (required) header configuration
 * @returns {Object<string, string>} - email -> name map
 */
function readEmailNamePairs_(sheet, headers) {
  /** @type {Object<string, string>} */
  const map = {};
  if (!sheet) return map;

  const parsed = readSheetRows_(sheet, headers);
  if (!parsed) return map;
  const { sheetData, colMap, headerRow, lastDataRow } = parsed;

  for (let i = headerRow + 1; i <= lastDataRow; i++) {
    const email = sheetData[i][colMap["SJSU Email"]];
    if (!email) continue;

    // Checks if key with same email already exists
    if (!Object.hasOwn(map, email)) map[email] = sheetData[i][colMap["Full Name"]] || "";
  }
  return map;
}

/**
 * @typedef {Object} MemberRecord
 * @property {string} name
 * @property {boolean} isAmbassador
 * @property {"Member"|"Non-member"} membershipStatus
 * @property {boolean} milestoneSent - "Active Milestone Sent?" carried over from before the rebuild
 * @property {number} social
 * @property {number} nonSocial
 * @property {number} tabling
 * @property {number} ambMeetAttended
 * @property {number} ambMeetExcused
 */

/**
 * Returns a new member record (stats)
 * 
 * @param {string} [fullName]
 * @returns {MemberRecord}
 */
function createMemberRecord_(fullName) {
    return {
      name: fullName || "",
      isAmbassador: false,
      membershipStatus: "Non-member",
      milestoneSent: false,
      social: 0,
      nonSocial: 0,
      tabling: 0,
      ambMeetAttended: 0,
      ambMeetExcused: 0
    };
}

/**
 * Builds a Members DB row (keyed by header name) for a single member record.
 *
 * Meetings Credited/Attended/Excused/Tabling Shifts/Req Met? only apply to ambassadors.
 *
 * @param {string} email
 * @param {MemberRecord} record
 * @param {number} meetingsRequired - total # of AmbMeet tabs found this run
 * @returns {Object<string, *>}
 */
function buildMemberRow_(email, record, meetingsRequired) {
  const activeMember = (record.social >= 1 && record.nonSocial >= 1) ? "Yes" : "No";

  /** @type {Object<string, *>} */
  const row = {
    "Full Name": record.name || "",
    "SJSU Email": email,
    "Membership Status": record.membershipStatus,
    "Active Member?": activeMember,
    "Event Total": record.social + record.nonSocial,
    "Social": record.social,
    "Non-social": record.nonSocial,
    "Active Milestone Sent?": record.milestoneSent,
  };

  // Tabling/Attended/Excused/Meetings Credited/Req Met? are ambassador-only
  // concepts - left blank ("") for everyone else via mapValuesToHeaders_'s default.
  if (record.isAmbassador) {
    const meetingsCredited = record.ambMeetAttended + record.ambMeetExcused;
    row["Ambassador?"] = "Yes"
    row["Tabling Shifts"] = record.tabling;
    row["Attended"] = record.ambMeetAttended;
    row["Excused"] = record.ambMeetExcused;
    row["Meetings Credited"] = meetingsCredited;
    row["Req Met?"] = activeMember === "Yes" && record.tabling >= 1 && meetingsCredited === meetingsRequired;
  }

  return row;
}

/**
 * Parses a single Members DB row back into a MemberRecord. Inverse of buildMemberRow_ -
 * only reconstructs the raw fields MemberRecord tracks; derived display columns (Active
 * Member?, Event Total, Req Met?) are recomputed from these where needed, not stored.
 *
 * @param {Array<*>} row - one row from a Members DB sheetData array
 * @param {Object<string, number>} colMap - header -> column index, from readSheetRows_
 * @returns {MemberRecord}
 */
function readMemberRecordFromRow_(row, colMap) {
  return {
    name: row[colMap["Full Name"]] || "",
    isAmbassador: row[colMap["Ambassador?"]] === "Yes",
    membershipStatus: row[colMap["Membership Status"]] === "Member" ? "Member" : "Non-member",
    milestoneSent: row[colMap["Active Milestone Sent?"]] === true,
    social: row[colMap["Social"]] || 0,
    nonSocial: row[colMap["Non-social"]] || 0,
    tabling: row[colMap["Tabling Shifts"]] || 0,
    ambMeetAttended: row[colMap["Attended"]] || 0,
    ambMeetExcused: row[colMap["Excused"]] || 0
  };
}

/**
 * Reads the Members DB sheet fresh and reconstructs a full email -> MemberRecord map.
 * Lets a consumer (e.g. emailSender) get its own up-to-date snapshot of member stats
 * without depending on any state left over from a prior updateDashboardTotals() run.
 *
 * @returns {Object<string, MemberRecord>} email -> MemberRecord
 */
function readMemberRecordsFromDashboard_() {
  /** @type {Object<string, MemberRecord>} */
  const memberRecords = {};

  const dashSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.membersSheet.name);
  if (!dashSheet) return memberRecords;

  const parsed = readSheetRows_(dashSheet, CONFIG.membersSheet.headers);
  if (!parsed) return memberRecords;
  const { sheetData, colMap, headerRow, lastDataRow } = parsed;
  const emailCol = colMap["SJSU Email"];

  for (let i = headerRow + 1; i <= lastDataRow; i++) {
    const email = sheetData[i][emailCol];
    if (!email) continue;
    memberRecords[email] = readMemberRecordFromRow_(sheetData[i], colMap);
  }
  return memberRecords;
}

/**
 * Finds a cell containing "Meetings Required:" anywhere in the sheet and writes
 * the given number into the cell immediately to its right.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} meetingsRequired
 */
function fillMeetingsRequiredLabel_(sheet, meetingsRequired) {
  const values = sheet.getDataRange().getValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (normalize_(values[r][c]).includes(normalize_("Meetings Required:"))) {
        sheet.getRange(r + 1, c + 2).setValue(meetingsRequired);
        return;
      }
    }
  }
}

/**
 * Reads back the "Meetings Required:" count written by fillMeetingsRequiredLabel_.
 * @returns {number}
 */
function getMeetingsRequiredCount_() {
  const dashSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.membersSheet.name);
  if (!dashSheet) return 0;

  const values = dashSheet.getDataRange().getValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (normalize_(values[r][c]).includes(normalize_("Meetings Required:"))) {
        return values[r][c + 1] || 0;
      }
    }
  }
  return 0;
}

/**
 * Rebuilds the Members DB roster (ambassadors -> members -> check-in-only) and recomputes
 * attendance stats, in a single pass over all event check-in tabs.
 *
 * @param {Object} e - Event object
 * @returns True when update is successful
 */
function updateDashboardTotals(e) {

  authorizeUser_(e);
  if (!isSheetReady_(CONFIG.membersSheet.name, CONFIG.membersSheet.headers)) return false;
  if (!isSheetReady_(CONFIG.ambassadorFormSheet.name, CONFIG.ambassadorFormSheet.headers)) return false;
  if (!isSheetReady_(CONFIG.membershipFormSheet.name, CONFIG.membershipFormSheet.headers)) return false;

  console.log("[updateDashboardTotals] rebuilding Members DB roster and stats");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(CONFIG.membersSheet.name);
  if (!dashSheet) return false;

  const parsed = readSheetRows_(dashSheet, CONFIG.membersSheet.headers);
  if (!parsed) return false;
  const { sheetData, colMap, headerRow: memberHeaderRow, lastDataRow: lastDashDataRow } = parsed;

  // Check for missing required columns
  const missingCols = CONFIG.membersSheet.headers.filter(header => colMap[header] === -1);
  if (missingCols.length > 0) {
    SpreadsheetApp.getUi().alert(`⚠️Row ${memberHeaderRow + 1} of "${CONFIG.membersSheet.name}" is missing header(s): ${missingCols.join(", ")}`);
    return false;
  }

  // Maps email -> engagement stats object
  /** @type {Object<string, MemberRecord>} */
  const memberRecords = {};

  /**
   * Merges every entry of a roster sheet into memberRecords, applying `applyFlag` to mark
   * each record's role (e.g. ambassador, member).
   *
   * @param {{name: string, headers: string[]}} sheetConfig - e.g. CONFIG.ambassadorFormSheet
   * @param {(record: MemberRecord) => void} applyFlag
   * @returns {number} the number of entries merged from the sheet
   */
  const mergeRoster = (sheetConfig, applyFlag) => {
    const map = readEmailNamePairs_(ss.getSheetByName(sheetConfig.name), sheetConfig.headers);
    Object.keys(map).forEach(email => {
      const name = map[email];
      // Create new record if not exists
      if (!memberRecords[email]) memberRecords[email] = createMemberRecord_(name);
      
      // Backfill name if missing
      else if (name && !memberRecords[email].name) memberRecords[email].name = name;

      // (Optional) Apply flag, ex. Ambassador? = true, if sheet processed was ambassador roster
      applyFlag(memberRecords[email]);
    });
    return Object.keys(map).length;
  };

  // 1. Create records from ambassador roster, preserve # ambassadors
  const ambassadorCount =  mergeRoster(CONFIG.ambassadorFormSheet, record => { record.isAmbassador = true; });

  // 2. Create records from membership form list
  mergeRoster(CONFIG.membershipFormSheet, record => { record.membershipStatus = "Member"; });
  let meetingsRequired = 0; // total # of AmbMeet tabs found this run

  // 3. Single pass over event tabs: Create check-in-only records + compile engagement stats for all
  ss.getSheets().forEach(sheet => {
    const type = classifyEventType_(sheet.getName());
    if (!type) return; // not an event tab

    // Create and update records for check-ins (ambassador meetings, social/non-social/tabling event)
    if (type === "ambMeet") meetingsRequired++;

    const headers = type === "ambMeet"
      ? CONFIG.ambMeetCheckInSheet.headers
      : CONFIG.checkInSheet.headers;
    const parsed = readSheetRows_(sheet, headers);
    if (!parsed) return;
    const { sheetData, colMap, headerRow, lastDataRow } = parsed;
    const emailCol = colMap["SJSU Email"];
    const nameCol = colMap["Full Name"];
    const attendedCol = colMap["Attended"]; // used only for ambMeet types
    const excusedCol = colMap["Excused"];  // used only for non-ambMeet types
    const seenInSheet = new Set();

    for (let i = headerRow + 1; i <= lastDataRow; i++) {
      const email = sheetData[i][emailCol];
      if (!email) continue;

      const name = sheetData[i][nameCol];
      // Create record if not exists
      if (!memberRecords[email]) memberRecords[email] = createMemberRecord_(name);

      // Backfill name if missing
      else if (name && !memberRecords[email].name) memberRecords[email].name = name;

      const record = memberRecords[email];
      if (seenInSheet.has(email)) continue; // prevent double counting within a single event
      seenInSheet.add(email);

      if (type === "ambMeet") {
        if (normalize_(sheetData[i][attendedCol]) === "yes") record.ambMeetAttended++;
        if (normalize_(sheetData[i][excusedCol]) === "yes") record.ambMeetExcused++;
        // else, increment unexcused
      } else if (type === "social") record.social++;
      else if (type === "nonSocial") record.nonSocial++;
      else if (type === "tabling") record.tabling++;
    }
  });

  fillMeetingsRequiredLabel_(dashSheet, meetingsRequired);

  // 4. Capture "Active Milestone Sent?", pre-update
  const emailCol = colMap["SJSU Email"];
  const milestoneSentCol = colMap["Active Milestone Sent?"];
  for (let i = memberHeaderRow + 1; i <= lastDashDataRow; i++) {
    const email = sheetData[i][emailCol];
    if (email && memberRecords[email]) memberRecords[email].milestoneSent = sheetData[i][milestoneSentCol] === true;
  }

  // 5. Build the ordered, deduped roster rows: ambassadors -> members -> check-in-only
  const orderedRows = Object.keys(memberRecords).map(
    email => buildMemberRow_(email, memberRecords[email], meetingsRequired)
  );

  // 6. Full rebuild: clear existing data rows, then write the updated roster
  /** @type {(header: string, numRows: number) => GoogleAppsScript.Spreadsheet.Range} */
  
  // Helper: get column's slice of data for specified header
  const getColumnRange = (columnHeader, numRows) => 
    // Using 1-indexed indices
    dashSheet.getRange(memberHeaderRow + 2, colMap[columnHeader] + 1, numRows, 1);

  // Clear existing rows
  const existingRowCount = lastDashDataRow - memberHeaderRow;
  if (existingRowCount > 0) {
    CONFIG.membersSheet.headers.forEach(header => getColumnRange(header, existingRowCount).clearContent());
  }

  // Write updated data
  if (orderedRows.length > 0) {

    // note: total # rows can change (with membership form submissions, ambassador onboarding, check-ins, ...)

    // Add checkboxes for Active Milestone Sent? and Req Met?
    getColumnRange("Active Milestone Sent?", orderedRows.length).insertCheckboxes();
    if (ambassadorCount > 0) getColumnRange("Req Met?", ambassadorCount).insertCheckboxes(); // Only fill for ambassadors

    // Write by column (assumes non-contiguous columns, but same column order)
    const rowsToWrite = orderedRows.map(row => mapValuesToHeaders_(row, CONFIG.membersSheet.headers));
    CONFIG.membersSheet.headers.forEach((header, i) => {
      getColumnRange(header, rowsToWrite.length).setValues(rowsToWrite.map(row => [row[i]]));
    });
  }

  return true;
}
