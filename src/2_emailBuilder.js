/**
 * Sends approved emails from the queue sheet after user confirmation.
 * @param {Object} e - Event object
 */
function sendApprovedEmailsInQueue(e) {

  var ui = SpreadsheetApp.getUi();
  authorizeUser_(e, CONFIG.ALLOWED_USERS);

  Logger.log("[sendApprovedEmailsInQueue] Processing approved emails...");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getActiveSheet();
  const sheetName = queueSheet.getName();

  // Guard: return if user is not on Email Queue
  if(normalize_(sheetName) !== normalize_(CONFIG.QUEUE_SHEET_NAME)) {
    return SpreadsheetApp.getUi().alert("Send failed: User is not viewing email queue.");
  }

  const data = queueSheet.getDataRange().getValues();

  // edit: start calculating length (# rows starting from header row); -> adjusted for header
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("Send failed: No email queue rows found.");
    return;
  }

  // TODO: check for required/matching headers, replace hardcoded header idx w/ dynamic

  // const headers = normalize_(data[3]);

  // const dateCol = findRequiredHeaderCol_(headers, "date generated");
  // const sentCol = findRequiredHeaderCol_(headers, "email sent?");
  // const nameCol = findRequiredHeaderCol_(headers, "full name");
  // const emailCol = findRequiredHeaderCol_(headers, "sjsu email");
  // const templateCol = findRequiredHeaderCol_(headers, "template type");

  let emailCount = 0;
  const readyEmailList = [];

  // replace w/ getting text from chosen email template
  // populate placeholders

  // Always ensure dashboard metrics are updated first
  updateDashboardTotals();

  // read dashboard

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

  for (let i = 4; i < data.length; i++) {
    const shouldSend = data[i][sentCol] === "No";

    if (!shouldSend) continue;

    const studentEmail = data[i][emailCol];
    const studentName = data[i][nameCol] || "there";
    const emailType = data[i][templateCol] || "default template - template wasn't found";

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

    // queueSheet.getRange(i + 1, statusCol + 1).setValue("Ready to send - Gmail line currently commented out");
    queueSheet.getRange(i + 1, dateCol + 1).setValue(new Date());

    emailCount++;
  }

  // todo: if any emails are missing stuff, mark yellow, show a warning -> 
  // X emails not processed and have been marked "missing info". + add a details column
  // X email is missing column: "columnName"
  // 
  // continue?
  // if no -> cancels

  console.log('ready email list: transforming to preview \n', readyEmailList);
  var recipientsPreviewString = readyEmailList.reduce((previewString, email, idx) => (
    previewString + `${idx}: ${email.address} - ${email.subject}\n`
  ), "");
  console.log('recipientspreview: \n', recipientsPreviewString); 

  var result = ui.alert(
    `💫Are you sure you want to send the following?\n(${emailCount} approved emails found):`,
    `Click Yes to confirm.\n\n ${recipientsPreviewString}`, 
    ui.ButtonSet.YES_NO);

  if (result === ui.Button.NO) return;

  // add a try catch in case
  // commented out to avoid hitting quota (100/day)
  readyEmailList.forEach(email => GmailApp.sendEmail(email.address, email.subject, email.body));

  // update email status (sent)
  ui.alert(`✅Success!\n\n${emailCount} approved row(s) were processed and sent.`, ui.ButtonSet.OK);
}