/**
 * [todo: add comment]
 */
function updateDashboardTotals() {
  console.log("[updateDashboardTotals] dashboard recalculation ran; updated member stats");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = ss.getSheetByName(CONFIG.MEMBERS_SHEET_NAME);
  if (!dashSheet) return;
  
  var lastDashRow = dashSheet.getLastRow();
  const memberHeaderRow = 5; // HARDCODED - row of headers -> todo: add to config
  const checkInHeaderRow = 2; // HARDCODED - row of headers -> todo: add to config
  if (lastDashRow < memberHeaderRow + 1) return; 
  
  // 1. Get all emails from the dashboard (Column B - aka. 2)
  var dashboardEmails = dashSheet.getRange(memberHeaderRow + 1, 2, lastDashRow - 1, 1).getValues();
  
  // 2. Initialize tracking dictionaries
  var socialCounts = {};
  var nonSocialCounts = {};
  
  // Pre-fill dictionaries with dashboard emails
  for (var i = 0; i < dashboardEmails.length; i++) {
    var email = normalize_(dashboardEmails[i][0]);
    if(!email || email.length === 0) continue;
    console.log("[dashboardUpdater] normalized email: ", email);
    if (email) {
      socialCounts[email] = 0;
      nonSocialCounts[email] = 0;
    }
  }
  
  // 3. Scan through all sheets to aggregate attendance data
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sheetName = sheets[s].getName();
    console.log("[dashboardUpdater] processing ", sheetName);
    var isSocial = sheetName.includes("Social") && !sheetName.includes("Non-social");
    var isNonSocial = sheetName.includes("Non-social");
    console.log("[dashboardUpdater]", sheetName, "is social: ", isSocial);
    console.log("[dashboardUpdater]", sheetName, "is non-social: ", isNonSocial);
    
    // Skip sheets that aren't event logs
    if (!isSocial && !isNonSocial) continue;
    
    var lastRow = sheets[s].getLastRow();
    if (lastRow < memberHeaderRow + 1) continue;
    
    // Get all emails from Column C of the event sheet
    var eventEmails = sheets[s].getRange(checkInHeaderRow + 1, 3, lastRow - 1, 1).getValues();
    var seenEmails = {};

    console.log("[dashboardUpdater] event emails for ", sheetName, ": ", eventEmails);
    
    for (var e = 0; e < eventEmails.length; e++) {
      var eventEmail = String(normalize_(eventEmails[e][0]) || "");
      if (!eventEmail) continue;

      // Pevent double counting for single event
      if (Object.hasOwn(seenEmails, eventEmail)) continue;
      seenEmails[eventEmail] = true; 
      
      if (isSocial && Object.hasOwn(socialCounts, eventEmail)) {
        socialCounts[eventEmail]++;
      } else if (isNonSocial && Object.hasOwn(nonSocialCounts, eventEmail)) {
        nonSocialCounts[eventEmail]++;
      }
    }
  }

  // 4. Prepare the final output objects to write back to the sheet

  // Master array to collect our formatted row blocks for the batch-write
  var batchRowsArray = [];

  // Loop through every student email in dashboard list
  for (var d = 0; d < dashboardEmails.length; d++) {
    var dashEmail = normalize_(dashboardEmails[d][0]);
    
    // Create an object for the member with explicit default values
    var memberRecord = {
      activeMember: "No",
      total: 0,
      social: 0,
      nonSocial: 0
    };

    if (!dashEmail) {
      // If the dashboard row has no email, convert an empty object structure to clean blanks
      batchRowsArray.push(["", "", "", ""]);
      console.log("[dashboardUpdater] warning: email not found in row ", memberHeaderRow + 1 + d);
    } else {
      // Pull counts from your tracking dictionaries (falling back to our object defaults if missing)
      memberRecord.social = socialCounts[dashEmail] || 0;
      memberRecord.nonSocial = nonSocialCounts[dashEmail] || 0;
      memberRecord.total = memberRecord.social + memberRecord.nonSocial;
      
      // ACTIVE MEMBER RULE: Must have at least 1 Social AND at least 1 Non-Social
      if (memberRecord.social >= 1 && memberRecord.nonSocial >= 1) {
        memberRecord.activeMember = "Yes";
      }
      
      // Flatten object properties into the exact column array format Google Sheets requires:
      // Layout columns D, E, F, G: [Active Member?, Event Total, Social, Non-social]
      batchRowsArray.push([
        memberRecord.activeMember, 
        memberRecord.total, 
        memberRecord.social, 
        memberRecord.nonSocial
      ]); 
    }
  }

  // 5. Batch Write: Drop all 4 columns into the sheet at once
  // Column 4 is Column D (Active Member?). Spans 4 columns wide (D, E, F, G).
  var dataRange = dashSheet.getRange(memberHeaderRow + 1, 4, batchRowsArray.length, 4);
  dataRange.setValues(batchRowsArray);


  // 4. Batch Dropdowns: Force dropdown options onto Column C over the printed values
  // var yesNoRule = SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"]).build();
  // var activeMemberRange = dashSheet.getRange(memberHeaderRow + 1, 4, batchRowsArray.length, 1);
  // activeMemberRange.setDataValidation(yesNoRule);
    




  // // 4. Prepare the final output arrays to write back to the sheet
  // var socialOutput = [];
  // var nonSocialOutput = [];
  // var totalOutput = [];
  // var activeMemberOutput = [];
  
  // for (var d = 0; d < dashboardEmails.length; d++) {
  //   var dashEmail = normalize_(dashboardEmails[d][0]);
  //   if (!dashEmail) { // email is empty
  //     socialOutput.push([""]);
  //     nonSocialOutput.push([""]);
  //     totalOutput.push([""]);
  //     console.log("[dashboardUpdater] warning: email not found in row ", memberHeaderRow + 1 + d);
  //   } else {
  //     socialOutput.push([socialCounts[dashEmail] || 0]);
  //     nonSocialOutput.push([nonSocialCounts[dashEmail] || 0]);
  //     totalOutput.push([socialCounts[dashEmail] + nonSocialCounts[dashEmail] || 0]);
  //   }
  // }

  // // TODO: add constant/config for social, non social column in members db
  
  // // 5. Batch-write the data back to the dashboard all at once (highly efficient)
  // dashSheet.getRange(memberHeaderRow + 1, 5, totalOutput.length, 1).setValues(totalOutput); // Column D (Event total)
  // dashSheet.getRange(memberHeaderRow + 1, 6, socialOutput.length, 1).setValues(socialOutput);     // Column E (Social)
  // dashSheet.getRange(memberHeaderRow + 1, 7, nonSocialOutput.length, 1).setValues(nonSocialOutput); // Column F (Non-social)
}