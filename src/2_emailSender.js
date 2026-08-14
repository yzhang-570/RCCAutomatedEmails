/**
 * Builds and sends approved emails from the queue sheet.
 * @param {Object} e - Event object
 */
function sendApprovedEmailsInQueue(e) {

  var ui = SpreadsheetApp.getUi();
  authorizeUser_(e);

  Logger.log("[sendApprovedEmailsInQueue] Processing approved emails...");

  // Get sheet being viewed + sheet name
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  const sheetName = currentSheet.getName();
  const data = currentSheet.getDataRange().getValues();

  // Guard: User must be viewing Email Queue to send emails
  if(normalize_(sheetName) !== normalize_(CONFIG.queueSheet.name)) {
    return SpreadsheetApp.getUi().alert("⚠️Send failed: User must be viewing Email Queue to send emails.");
  }

  // Get header row of email queue
  const headerRow = findHeaderRowIndex_(CONFIG.queueSheet.headers, currentSheet);
  Logger.log(`[sendApprovedEmailsInQueue] Header row found in row: ${headerRow}`);
  if(headerRow === null) {
    return SpreadsheetApp.getUi().alert(`
      [sendApprovedEmailsInQueue] ⚠️Send failed: Queue sheet headers not found (may be mismatched)
      \nTarget headers: ${CONFIG.queueSheet.headers}
    `);
  }

  // Check if queue is empty
  if (data.length - (headerRow + 1) <= 0) {
    SpreadsheetApp.getUi().alert("Send failed: Queue is empty.");
    return;
  }

  // Get column indices of each header
  const cols = mapHeaderIndices_(CONFIG.queueSheet.headers, data[headerRow]);
  const dateCol = cols["Date Generated"];
  const sentCol = cols["Email Sent?"];
  const nameCol = cols["Full Name"];
  const emailCol = cols["SJSU Email"];
  const eventCol = cols["Triggered By (Event Name)"];

  let emailCount = 0;
  const readyEmailList = [];

  // Always ensure dashboard metrics are updated before sending emails
  updateDashboardTotals(e);

  for (let i = headerRow + 1; i < data.length; i++) {
    const shouldSend = data[i][sentCol] !== "Yes";
    if (!shouldSend) continue;

    const row = data[i];

    // determine email type/status updates
    // read dashboard to populate most updated stats

    // i.e. if !ActiveMilestoneSent and reached -> send

    const studentEmail = row[emailCol];
    const studentName = row[nameCol] || "there";
    const emailType = row[eventCol] || "default template - triggering event wasn't found";

    const subject = `${emailType} Update`;
    const body = `Hi ${studentName},

    This is a placeholder for the "${emailType}" email.

    Best,
    RCC`;

    readyEmailList.push({
      address: studentEmail,
      subject: subject,
      body: body
    })

    // currentSheet.getRange(i + 1, sentCol + 1).setValue("Ready to send - Gmail line currently commented out");
    currentSheet.getRange(i + 1, dateCol + 1).setValue(new Date());

    emailCount++;
  }

  // todo: if any emails are missing stuff, mark yellow, show a warning -> 
  // X emails not processed and have been marked "missing info". + add a details column
  // X email is missing column: "columnName"
  // 
  // continue?
  // if no -> cancels

  console.log('Ready email list: transforming to preview \n', readyEmailList);
  var recipientsPreviewString = readyEmailList.reduce((previewString, email, idx) => (
    previewString + `${idx}: ${email.address} - ${email.subject}\n`
  ), "");
  console.log('Recipients preview: \n', recipientsPreviewString); 

  var result = ui.alert(
    `💫Are you sure you want to send the following?\n(${emailCount} approved emails found):`,
    `Click Yes to confirm.\n\n ${recipientsPreviewString}`, 
    ui.ButtonSet.YES_NO);
  if (result === ui.Button.NO) return;

  // add a try catch in case
  // commented out to avoid hitting quota (100/day)
  // for any unsent emails, schedule send (i.e. regularly send queued unsent emails at 9am daily)
  readyEmailList.forEach(email => GmailApp.sendEmail(email.address, email.subject, email.body));
  // once sent, gray out bg + make text gray

  // update email status (sent)
  ui.alert(`✅Success!\n\n${emailCount} approved row(s) were processed and sent.`, ui.ButtonSet.OK);
}






  // determine email type/status updates
  /* 

    subject: RCC: Active Member Status
    RCC: Active Member + Ambassador Status
    WOOOOOOOOO You're an Active Member!

    IF just became active member - addon
    - special celebration message! congrats

    IF ambassador only - addons
    - general meeting
    - tabling: 
    - active member status: registered

    ALL:
    - active member update - current status, # social, # nonsocial

  */