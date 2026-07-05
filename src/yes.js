// /**
//  * The event handler triggered when opening the spreadsheet.
//  * @param {Event} e The onOpen event.
//  */
// function onOpen(e) {
//   // Add a custom menu to the spreadsheet
//   SpreadsheetApp.getUi()
//     .createMenu("RCC Auto-Email")
//     .addItem("Send approved emails", "sendApprovedEmailsInQueue")
//     .addToUi();
// }
const SPREADSHEET_ID = "1jHgzfrLoI32UQvmIgy_dYVHC5TFTwkrnhZkosP-o2Uo";
function setupMenuTrigger() {
  var targetSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  ScriptApp.newTrigger("injectCustomMenu")
           .forSpreadsheet(targetSheet)
           .onOpen()
           .create();
}

function injectCustomMenu() {
  SpreadsheetApp.getUi()
    .createMenu("RCC Auto-Email")
    .addItem("Send approved emails", "sendApprovedEmailsInQueue")
    .addItem("another button that does the same", "sendApprovedEmailsInQueue")
    .addToUi();
}

function sendApprovedEmailsInQueue() {
  Logger.log("[sendApprovedEmailsInQueue] Sending approved emails");
  console.log("testing console.log");
  console.log('testsetstset');
}




















/*
future todos:
- use standalone script
- remove old triggers on run at top
*/
// function myFunction() {
//   // get doc template + get sheet by ID
//   // events - log users/emails that triggered event, changeType (insert row/insert_column, etc.)
//     // doable for menu-triggered (?)
//     // block unauthorized users 
// }
