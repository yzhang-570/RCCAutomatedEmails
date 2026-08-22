/**
 * Builds the dynamic "Your RCC Status Update" email for a member, based on their current
 * Members DB stats. Includes the non-member registration prompt and/or ambassador
 * progress section only when applicable.
 *
 * @param {MemberRecord} record
 * @param {number} meetingsRequired - total # of AmbMeet tabs found this run
 * @param {number} eventName
 * @returns {{subject: string, body: string}}
 */
function buildStatusUpdateEmail_(record, meetingsRequired, eventName) {
  const firstName = getFirstName_(record.name) || "there";
  const activeMember = record.social >= 1 && record.nonSocial >= 1;

  let body = `Hi ${firstName},

Thanks for being involved with RCC! We saw that you recently attended ${eventName}. Here’s a quick update on your progress this semester.

**Active Member Progress**
Social event: ${record.social}/1 fulfilled
Non-social event: ${record.nonSocial}/1 fulfilled
Status: ${
    record.membershipStatus !== "Member"
      ? "{Become a member to be eligible to be an Active Member. Register for RCC Membership:|gray}" + `${CONFIG.membershipFormLink}`
      : activeMember
        ? "{Active Member!|green}"
        : "{In Progress|orange}"
  }`;

  if (record.isAmbassador) {
    const meetingsCredited = record.ambMeetAttended + record.ambMeetExcused;
    const reqMet = activeMember && record.tabling >= 1 && meetingsCredited === meetingsRequired;
    body += `

**Ambassador Requirements**
General Meetings: ${meetingsCredited}/${meetingsRequired} required
Tabling: ${Math.min(record.tabling, 1)}/1 required
Active Member Requirement: ${activeMember ? "Complete" : "In Progress"}
Status: ${reqMet ? "{Fulfilled|green}" : "{In Progress|orange}"}`;
  }

  body += `
  
We hope to see you at our upcoming opportunities and events!

Best,
Responsible Computing Club
Email: ${CONFIG.socialLinks.email}
Discord: ${CONFIG.socialLinks.discord}
Instagram: ${CONFIG.socialLinks.instagram}
LinkedIn: ${CONFIG.socialLinks.linkedin}`;

  return { subject: "Your RCC Status Update", body };
}

/**
 * Builds and sends approved emails from the queue sheet.
 * @param {Object} e - Event object
 */
function sendEmailsInQueue(e) {

  authorizeUser_(e);
  if (!isSheetReady_(CONFIG.queueSheet.name, CONFIG.queueSheet.headers)) return;

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

  // Get header row for column mapping
  const headerRow = findHeaderRowIndex_(CONFIG.queueSheet.headers, currentSheet);
  if (headerRow === null || headerRow === -1) {
    return SpreadsheetApp.getUi().alert(`⚠️Send failed: couldn't find header row with columns:\n${CONFIG.queueSheet.headers}`);
  }
  Logger.log(`[sendApprovedEmailsInQueue] Header row found in row: ${headerRow}`);

  // Check if queue is empty
  if (data.length - (headerRow + 1) <= 0) {
    SpreadsheetApp.getUi().alert("Send failed: Queue is empty.");
    return;
  }

  // Get column indices of each header
  const cols = mapHeaderIndices_(CONFIG.queueSheet.headers, data[headerRow]);
  const missingCols = CONFIG.queueSheet.headers.filter(h => cols[h] === -1);
  if (missingCols.length > 0) {
    return SpreadsheetApp.getUi().alert(`⚠️Send failed: Email Queue is missing header(s): ${missingCols.join(", ")}`);
  }

  const dateCol = cols["Date Generated"];
  const sentCol = cols["Email Sent?"];
  const emailCol = cols["SJSU Email"];
  const eventCol = cols["Triggered By (Event Name)"];

  let emailCount = 0;
  const readyEmailList = [];

  // Always ensure dashboard metrics are updated before sending emails
  const success = updateDashboardTotals(e);
  if (!success) return;

  // Fresh snapshot of member stats (maps email -> stats) + ambassador meetings required
  const memberRecords = readMemberRecordsFromDashboard_();
  const meetingsRequired = getMeetingsRequiredCount_();

  for (let i = headerRow + 1; i < data.length; i++) {
    const shouldSend = data[i][sentCol] !== "Yes";
    if (!shouldSend) continue;

    const row = data[i];
    const studentEmail = row[emailCol];
    const eventName = row[eventCol];

    const record = memberRecords[studentEmail];
    if (!record) {
      logIssue_({
        "Row #": i + 1,
        "Source Location": CONFIG.queueSheet.name,
        "Type": "Error",
        "Issue": `Failed to send status update: ${studentEmail} not found in "${CONFIG.membersSheet.name}"`
      });
      continue;
    }

    const { subject, body } = buildStatusUpdateEmail_(record, meetingsRequired, eventName);

    readyEmailList.push({
      sheetRow: i + 1,
      address: studentEmail,
      subject: subject,
      body: body
    })

    // currentSheet.getRange(i + 1, sentCol + 1).setValue("Ready to send - Gmail line currently commented out");
    currentSheet.getRange(i + 1, dateCol + 1).setValue(new Date());

    emailCount++;
  }

  console.log('Ready email list: transforming to preview \n', readyEmailList);
  var recipientsPreviewString = readyEmailList.reduce((previewString, email, idx) => (
    previewString + `${idx}: ${email.address} - ${email.subject}\n`
  ), "");
  console.log('Recipients preview: \n', recipientsPreviewString); 

  const ui = SpreadsheetApp.getUi();
  let result = ui.alert(
    `💫Are you sure you want to send the following?\n(${emailCount} unsent emails found):`,
    `Click Yes to confirm.\n\n ${recipientsPreviewString}`, 
    ui.ButtonSet.YES_NO);
  if (result === ui.Button.NO) return;

  // commented out to avoid hitting quota (100/day)
  // for any unsent emails, schedule send (i.e. regularly send queued unsent emails at 9am daily)

  readyEmailList.forEach(email => {
    GmailApp.sendEmail(email.address, email.subject, email.body, { htmlBody: plainTextToHtml_(email.body) });

    // Update sent status to 'Yes'
    currentSheet.getRange(email.sheetRow, sentCol + 1).setValue("Yes");
  });

  ui.alert(`✅Success!\n\n${emailCount} approved row(s) were processed and sent.`, ui.ButtonSet.OK);
}

/**
 * Clears all data from the email queue.
 * @param {Object} e - Event object
 */
function clearEmailQueue(e) {

  authorizeUser_(e);
  if (!isSheetReady_(CONFIG.queueSheet.name, CONFIG.queueSheet.headers)) return;

  const ui = SpreadsheetApp.getUi();

  // Get sheet being viewed + sheet name
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  const sheetName = currentSheet.getName();

  // Guard: User must be viewing Email Queue to clear emails
  if(normalize_(sheetName) !== normalize_(CONFIG.queueSheet.name)) {
    return SpreadsheetApp.getUi().alert("⚠️Clear failed: User must be viewing Email Queue to send emails.");
  }

  const headerRow = findHeaderRowIndex_(CONFIG.queueSheet.headers, currentSheet);
  if (headerRow === null || headerRow === -1) return;

  const numRowsToClear = currentSheet.getLastRow() - (headerRow + 1);
  if (numRowsToClear <= 0) {
    return ui.alert("Email Queue is already empty.");
  }

  const result = ui.alert(
    "⚠️Clear Email Queue?",
    `This will delete ${numRowsToClear} row(s) from the Email Queue.\n\nClick Yes to confirm.`,
    ui.ButtonSet.YES_NO
  );
  if (result === ui.Button.NO) return;

  currentSheet.deleteRows(headerRow + 2, numRowsToClear);
  ui.alert(`✅Cleared ${numRowsToClear} row(s) from the Email Queue.`);
}






  // determine email type/status updates
  /* 

    subject: RCC Status Progress

    We're sharing this update, because you recently attended RCC's X event (event type -> should be standardized, ex. AmbMeeting -> ambassador meeting).
    for ambmeeting: you recently attended/were marked excused for the X Ambassador General Meeting.

    membership status: non-member, member, active member! (0/1 socials, 0/1 non-socials attended)
    Register to be a member [here]! RCC membership is free.
    You're so close! Attend 1 more X event to become an active member.
    -# (small text): Members who attend at least 1 social and 1 non-social event per semester

    IF non-member
    - send we noticed you're not a member yet. Register to be a member [here]! RCC membership is free.
    - members get X?

    IF ambassador only - addons
    Ambassador requirements: In progress
    - meetings credited: X/X required
    - tabling: X/1 required
    - be an active member: in progress vs. fulfilled



    subject: WOOOOOOOOO You're an Active Member!
    IF just became active member (activeMilestoneSent = false)
    - special celebration message! congrats
    then set activeMilestoneSent to true

    subject: Fill out the RCC Membership Form
  */