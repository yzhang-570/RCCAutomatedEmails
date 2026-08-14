/**
 * @param {Object} e - Event object
 */
function processActiveEventTab(e) {

  authorizeUser_(e, CONFIG.ALLOWED_USERS);

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get the (event) sheet the user is looking at and Email Queue sheet
  var activeTab = activeSpreadsheet.getActiveSheet();
  const queueSheet = activeSpreadsheet.getSheetByName("Email Queue");
  if(!queueSheet) return SpreadsheetApp.getUi().alert("⚠️Check-in processing failed: Email queue not found.");

  // 2. Get the check in data and event name
  const checkInData = activeTab.getDataRange().getValues();
  const eventName = activeTab.getName();

  // TODO: migrate registered members responses from it's own sheet to avoid bloat

  // Guards - return if the current tab:
  // 1. is a known non-event tab (restricted)
  // 2. TODO: doesn't match the required header pattern
  const restrictedSheets = CONFIG.RESERVED_SHEETS;
  const isRestricted = restrictedSheets.findIndex(sheetName => normalize_(sheetName).includes(normalize_(eventName))) !== -1
  if (isRestricted) return SpreadsheetApp.getUi().alert(`⚠️Check-in processing failed:\n"${eventName}" is not an event check-in sheet`);

  // required headers check


  Logger.log(`[processActiveEventTab] Processing event data for: ${eventName}` );
  
  // 3. Perform appends to email queue
  // Skip table header rows (i = 2)
  let successfulAppends = 0;
  for (let i = 2; i < checkInData.length; i++) {
    const fullName = checkInData[i][1];
    const sjsuEmail = checkInData[i][2];
    
    if (!sjsuEmail) continue; // Skip empty rows
    
    // Structure matches your queue: Date Generated | Email Sent? | Full Name | SJSU Email | Triggered By
    const newRow = [
      new Date(),       // Date Generated
      "",               // Email Sent? (Leave blank until sent)
      fullName ?? "Beloved RCC member",         // Full Name
      // add a find to default to ambassadors vs. member depending on ambassador? column
      sjsuEmail,        // SJSU Email
      eventName         // Triggered By (The name of the event tab)
    ];
    
    Logger.log("Appending row to queue: ", newRow);
    queueSheet.appendRow(newRow);
    successfulAppends += 1;
    Logger.log("Succesfully appended row");
  }
  
  // Exclude header rows
  SpreadsheetApp.getUi().alert("Successfully processed tab: " + activeTab.getName() + `\n\n${successfulAppends}/${checkInData.length - 2}
    emails were added to queue.`);
}