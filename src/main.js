/**
 * Useful information:
 * 
 * Navigating this codebase - RCC Automated Emails
 * - Apps Script treats all .js files as a shared global scope and doesn't use ESM (no import/export statements)
 */


// connect this to a doc
// refactor - split emailBuilder into helpers
// add a function to find header row
  // throw error if expected header signature not found



// ----------
// any annoying logic -> ignore, just use placeholder/fixed

// check in form template - based on og template
// feature: check-in data processing - populate Email Queue
    // create links, "triggered by (details)", status changes
    // choose correct template based on status change -> ignore logic for now
    // get event type, template doc from Config tab
    // only process if Email status = not created -> else warning to proceed
// next (important): use generated email doc, WRITE to send status
    // use generated doc link (or doc id) for template
    // weave info into template docs
    // update queue status, gray out
    // update event config status, sent

// ------------
// ambassadors - email template logic
// sheet - syncing engagement dashboard tracking

// stretch - not a registered member?
  //  add a line to fill out registered member form to be eligible to become Active Member

// stretch - better check-in data syncing:
  // use url to sync check-in data
  // use ui to enter (social, nonsocial, etc.)
  // updates table
  // use a dropdown, ideally

// Runs upon spreadsheet being opened
function setupMenuTrigger() {
  var targetSheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
 
  // Attach custom menu
  ScriptApp.newTrigger("injectCustomMenu")
           .forSpreadsheet(targetSheet)
           .onOpen()
           .create();

  // TODO: run updateDashboardTotals once on open
  // TODO: run on membership form submit (auto-update)
  // TODO: make process active event tab run updateDashboardTotals once at start
}
function injectCustomMenu() {
  SpreadsheetApp.getUi()
    .createMenu("📧RCC Auto-Email")
    // stretch: add guard - only show option if on an event tab (known to be non-not event)
    // use on tab change in addition to on open, if exists
    .addItem("Process active event tab (auto-generate status update emails)", "processActiveEventTab")
    
    // stretch: add guard - only show option if on Email Queue tab
    .addItem("Email Queue: Send approved emails", "sendApprovedEmailsInQueue") // see 3_emailBuilder.js
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu('Administration')
    .addItem('🔄 Recalculate Dashboard Totals', 'updateDashboardTotals')
    .addToUi();
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
