/**
 * Useful information:
 * 
 * Navigating this codebase - RCC Automated Emails
 * - Apps Script treats all .js files as a shared global scope and doesn't use ESM (no import/export statements)
 */

// Runs upon spreadsheet being opened
function setupMenuTrigger() {
  var targetSheet = SpreadsheetApp.openById(CONFIG.mainSpreadsheetId);
  var existingHandlers = ScriptApp.getProjectTriggers().map(trigger => trigger.getHandlerFunction());

  // Attach custom menu
  if (existingHandlers.indexOf("injectCustomMenu") === -1) {
    ScriptApp.newTrigger("injectCustomMenu")
             .forSpreadsheet(targetSheet)
             .onOpen()
             .create();
  }

  // Attach welcome email trigger (see 5_welcomeHandler.js)
  if (existingHandlers.indexOf("onMembershipFormSubmit") === -1) {
    ScriptApp.newTrigger("onMembershipFormSubmit")
             .forSpreadsheet(targetSheet)
             .onFormSubmit()
             .create();
  }

// TODO: run updateDashboardTotals:
    // - (done) on open
    // - on email generation
    // - (done) before email send

  // Update dashboard stats on open
  // if (existingHandlers.indexOf("updateDashboardTotals") === -1) {
  //   ScriptApp.newTrigger("updateDashboardTotals")
  //            .forSpreadsheet(targetSheet)
  //            .onOpen()
  //            .create();
}

function injectCustomMenu() {
  SpreadsheetApp.getUi()
    .createMenu("📧RCC Auto-Email")

    .addItem("Generate emails: Process active event tab", "generateEmailsFromActiveTab")
    
    .addItem("Send queued emails: Process email queue", "sendEmailsInQueue") // see 3_emailBuilder.js
    .addItem("🗑️ Clear email queue", "clearEmailQueue")

    .addItem('🔄 Update dashboard totals: Process check-in data', 'updateDashboardTotals')
    .addToUi();
}
