/*
------------Google Sheet Config----------- 
*  Update sheet names (name) and header schemas (headers) to keep consistent with naming in dashboard sheet
*/
const DASHBOARD_SHEET_ID = "1hYCE-vMWlFxNbZdxeRFTrnCf1ePDh6kp3kuBwdI9ANw" // Automated Emails (sheet name as of 8/13)

const MEMBERS_SHEET = {
  name: "Members DB",
  headers: [
    // Regular stats
    "Full Name", "SJSU Email", "Membership Status", "Active Member?", "Event Total", "Social", "Non-social",
    
    // Ambassador only stats
    "Ambassador?", "Req Met?", "Tabling Shifts", "Meetings Credited", "Attended", "Excused",
    
    // Flag for tracking emails
    "Active Milestone Sent?"
  ]
};

const QUEUE_SHEET = {
  name: "Email Queue",
  headers: ["Date Generated", "Email Sent?", "Full Name", "SJSU Email", "Triggered By (Event Name)"]
};

// Header shape for social, non-social, tabling check-ins
const CHECKIN_SHEET = {
  // sheet name - determined by event name (event type)
  headers: ["Timestamp", "Full Name", "SJSU Email"]
}

const AMB_MEET_CHECKIN_SHEET = {
  // "Attended" is a formula column: checks whether this ambassador's email appears in the mirrored check-in list
  // "Excused" is manually added by an admin for ambassadors who missed the meeting but were excused ahead of time
  headers: ["Timestamp", "Full Name", "SJSU Email", "Attended", "Excused"]
}

const MEMBERSHIP_FORM_SHEET = {
  name: "Membership Form",
  headers: ["Timestamp", "Full Name", "SJSU Email"]
};

const LOG_SHEET = {
  name: "Processing Log",
  headers: ["Timestamp", "Row #", "Source Location", "Type", "Issue", "Details"]
};

const AMBASSADOR_FORM_SHEET = {
  name: "Ambassador List",
  headers: ["Full Name", "SJSU Email"]
}

const ALLOWED_EMAILS = [
  "rcc.sjsu@gmail.com"
]

const RCC_SOCIALS = {
  email: "rcc.sjsu@gmail.com",
  instagram: "instagram.com/rcc.sjsu",
  discord: "discord.com/invite/pmv9GJTgjm",
  linkedin: "linkedin.com/company/rcc-sjsu/"
}

const MEMBERSHIP_FORM_LINK = "https://docs.google.com/forms/d/e/1FAIpQLSfvBCZ7UzQIE-GZfhI75Pv3uakOJFAzUQ1mMc6a-zOUSwD_Tg/viewform?usp=header"


/*
------------DERIVED----------- **do not modify**
*/
const CONFIG = {
  mainSpreadsheetId: DASHBOARD_SHEET_ID, // central automated emails sheet
  membersSheet: MEMBERS_SHEET,           // member dashboard
  queueSheet: QUEUE_SHEET,               // email queue
  checkInSheet: CHECKIN_SHEET,           // event check-in format
  ambMeetCheckInSheet: AMB_MEET_CHECKIN_SHEET, // ambassador meeting check-in format
  logSheet: LOG_SHEET,                   // error/warning logs
  membershipFormSheet: MEMBERSHIP_FORM_SHEET,     // RCC membership form
  ambassadorFormSheet: AMBASSADOR_FORM_SHEET,     // ambassador roster

  reservedSheets: [
    MEMBERS_SHEET.name,
    QUEUE_SHEET.name,
    LOG_SHEET.name,
    MEMBERSHIP_FORM_SHEET.name,
    AMBASSADOR_FORM_SHEET.name
  ],
  
  socialLinks: RCC_SOCIALS,
  membershipFormLink: MEMBERSHIP_FORM_LINK,

  allowedUsers: ALLOWED_EMAILS
};
