/**
 * Sends a welcome email to the new registrant.
 * 
 * Fires when the Membership Form is submitted.
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e - Form submit event object
 */
function onMembershipFormSubmit(e) {
  const sheet = e.range.getSheet();
  if (normalize_(sheet.getName()) !== normalize_(CONFIG.membershipFormSheet.name)) return;

  const submitterDetails= {
    name: e.namedValues["Full Name"]?.[0] || "",
    email: e.namedValues["SJSU Email"]?.[0] || ""
  };

  // Check if membership form header row is valid
  const headerRow = findHeaderRowIndex_(CONFIG.membershipFormSheet.headers, sheet);
  if (headerRow === null || headerRow === -1) {
    logIssue_({
      "Row #": e.range.getRow(),
      "Source Location": CONFIG.membershipFormSheet.name,
      "Type": "Error",
      "Issue": `Failed to send welcome email: required header columns not found\n
                ${CONFIG.membershipFormSheet.headers}`,
      "Details": `${JSON.stringify(submitterDetails)}`
    });
    return;
  }

  const headerValues = sheet.getRange(headerRow + 1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cols = mapHeaderIndices_(CONFIG.membershipFormSheet.headers, headerValues);

  const rowNum = e.range.getRow(); // row # of this submission
  const rowValues = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fullName = rowValues[cols["Full Name"]];
  const sjsuEmail = rowValues[cols["SJSU Email"]];

  // Return if sjsu email empty or not valid
  if (!sjsuEmail || !sjsuEmail.includes("@sjsu.edu")) {
    logIssue_({
      "Row #": rowNum,
      "Source Location": CONFIG.membershipFormSheet.name,
      "Type": "Error",
      "Issue": "Failed to send welcome email: Missing/invalid SJSU Email",
      "Details": `${JSON.stringify(submitterDetails)}`
    });
    return;
  }

  const studentName = getFirstName_(fullName) || placeholderName;
  const subject = "Welcome to RCC!";
  const body = 

  `Hi ${studentName},

We’re excited to have you here! RCC offers many opportunities to gain hands-on experience, meet new people, and grow professionally.

**How can I get involved?:**
- Leadership opportunities
- Consulting and student-led projects
- Industry events and workshops
- Networking and socials
- ...and more!

We hope to to see you at our upcoming RCC events!

Best,
Responsible Computing Club
Instagram: ${CONFIG.socialLinks.instagram}
Discord: ${CONFIG.socialLinks.discord}
LinkedIn: ${CONFIG.socialLinks.linkedin}`;

  GmailApp.sendEmail(sjsuEmail, subject, body, { htmlBody: plainTextToHtml_(body) });
  Logger.log(`[onMembershipFormSubmit] Welcome email sent to ${sjsuEmail}`);
}
