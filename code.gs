/**
 * Creator: Galbatorix
 * Contact the creator for support, feature requests, or issues.
 *
 * Version: 7.17
 * Date: 2025-05-30
 */

// Global Emails Configuration
const CONFIG = {
  REPORTING_EMAILS: ["reporint1@email.com", "reporting2@email.com"], 
  GM_EMAIL: "your_gm@email.com",
  HR_EMAIL: "your_hr@email.com",
  SPV_MAP: {
    "Marketing": "your_marketing@email.com",
    "Finance": "your_finance@email.com"
  }
};

// Column indices
const COLUMNS = {
  TIMESTAMP: 1,
  NAME: 2,
  DEPARTMENT: 3,
  LEAVE_TYPE: 4,
  START_DATE: 5,
  END_DATE: 6,
  REASON: 7,
  STATUS: 8,
  REQUESTER_EMAIL: 9,
  SPV_EMAIL: 10,
  HR_EMAIL: 11,
  GM_EMAIL: 12,
  STAGE: 13,
  DECISION: 14,
  NOTE: 15,
  DECISION_DATE: 16,
  SPV_DECISION: 17,
  HR_DECISION: 18,
  GM_DECISION: 19,
  SPV_TOKEN: 20,
  HR_TOKEN: 21,
  GM_TOKEN: 22
};

/**
 * Helper function to parse dates from "d-m-Y" format.
 * @param {string} dateString The date string in "d-m-Y".
 * @return {Date} A JavaScript Date object.
 */

function parseDMYDate(dateString) {
  if (!dateString || typeof dateString !== 'string') return null;
  const parts = dateString.split("-");
  if (parts.length === 3) {
    // new Date(year, monthIndex, day)
    return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
  }
  return null; // Or throw an error, or return an invalid Date
}

// Helper function to format date objects to string
function formatDate(dateObj) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
    // If it's already a string (potentially from sheet), return it, or handle error
    return (typeof dateObj === 'string') ? dateObj : "Invalid Date";
  }
  try {
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  } catch(e) {
    Logger.log("formatDate error: " + e.toString());
    // Fallback for environments where Session might not be available or date is weird
    const d = dateObj.getDate();
    const m = dateObj.getMonth() + 1;
    const y = dateObj.getFullYear();
    return (d < 10 ? '0' : '') + d + '-' + (m < 10 ? '0' : '') + m + '-' + y;
  }
}

function generateRandomToken(length = 8) {
  const chars = 'abcdefghijklmnopqrstuvwxyz1234567890~<>!@#$%^&*';
  let token = '';
  for (let i = 0; i < length; i++) {
    token += chars[Math.floor(Math.random() * chars.length)];
  }
  return token;
}

function showErrorTokenPage(title, message) {
  const html = `
    <html>
      <head>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
      </head>
      <body class="bg-light d-flex align-items-center justify-content-center vh-100">
        <div class="card shadow-lg" style="max-width: 500px;">
          <div class="card-body text-center">
            <h3 class="text-danger mb-3">⚠️ ${title}</h3>
            <p class="text-muted">${message}</p>
            <hr>
            <a href="${ScriptApp.getService().getUrl()}" class="btn btn-primary mt-3">Return to Form</a>
          </div>
        </div>
      </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle(title);
}

function doGet(e) {
  const page = e?.parameter?.page;
  const action = e.parameter.action;
  const row = parseInt(e.parameter.row, 10);
  const stage = e.parameter.stage;
  const note = e.parameter.note;

  // Process the actual rejection with notes here
  if (action === 'review' && row && stage && note !== undefined) {
    // return processRejectionWithNote(row, stage, note);  // Enable this must enable to function rejection on line 163
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const COLUMNS = Object.fromEntries(headers.map((h, i) => [h, i + 1]));

    const rowIndex = row;
    const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

    const name = values[COLUMNS["Name"] - 1];
    const department = values[COLUMNS["Department"] - 1];
    const leaveType = values[COLUMNS["Leave Type"] - 1];
    const startDate = values[COLUMNS["Start Date"] - 1];
    const endDate = values[COLUMNS["End Date"] - 1];
    const requester = values[COLUMNS["Requester Email"] - 1];
    const reasonText = values[COLUMNS["Reason"] - 1];
    const currentStage = stage;

    sheet.getRange(rowIndex, COLUMNS["Status"]).setValue('Rejected');
    sheet.getRange(rowIndex, COLUMNS["Stage"]).setValue('Rejected at ' + currentStage);

    const formattedStartDate = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
    const formattedEndDate = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
    const leaveDays = Math.ceil((new Date(endDate) - new Date(startDate)) / (1000 * 60 * 60 * 24)) + 1;

    const rejectionHtml = `
      <div style="font-family: Arial, sans-serif; color: #333;">
        <h3 style="color: #dc3545;">❌ Leave Request Rejected - ${name}</h3>
        <p><strong>Requester:</strong> ${name}</p>
        <p><strong>Department:</strong> ${department}</p>
        <p><strong>Leave Type:</strong> ${leaveType}</p>
        <p><strong>Dates:</strong> ${formattedStartDate} to ${formattedEndDate} (${leaveDays} day${leaveDays > 1 ? 's' : ''})</p>
        <p><strong>Reason:</strong> ${reasonText}</p>
        <p><strong>Rejected at Stage:</strong> ${currentStage}</p>
        <p><strong>Note:</strong> ${note}</p>
        <div style="margin-top: 20px; padding: 10px; background-color: #f8d7da; color: #721c24; border-left: 5px solid #f5c6cb;">
          <strong>Status:</strong> Rejected
        </div>
      </div>
    `;

    GmailApp.sendEmail({
      to: requester,
      subject: `Your Company Name Leave/WFH Request Rejected: ${name}`,
      htmlBody: rejectionHtml,
      name: "Your Company Name Approval System"
    });
    const resultTemplate = HtmlService.createTemplateFromFile('result');
    resultTemplate.status = "Rejected";
    resultTemplate.message = "Rejection submitted successfully with notes.";
    resultTemplate.color = "danger";
    return resultTemplate.evaluate().setTitle("Rejection Submitted");
  }

  if (action === 'review' && row && stage) {
    const template = HtmlService.createTemplateFromFile('rejectWithNotes');
    template.row = row;
    template.stage = stage;
    return template.evaluate().setTitle('Reject with Notes');
  }


  // If this function active, do not forget to uncomment the return processRejectionWithNote(row, stage, note); at line 95
  // And comment all code inside if (action === 'review' && row && stage && note !== undefined) except reject prosses!

  // This function is to use rejection with notes, but not today, so I just need to commented this feature
  // We are hold this feature b'coz there a bug while using this.
  // We'll be right back after get some nice salary hahahaha :v, kidding ~~

  // function processRejectionWithNote(rowIndex, stage, note) {
  //   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  //   const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  //   const COLUMNS = Object.fromEntries(headers.map((h, i) => [h, i + 1]));
  //   const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  //   const name = values[COLUMNS["Name"] - 1];
  //   const department = values[COLUMNS["Department"] - 1];
  //   const leaveType = values[COLUMNS["Leave Type"] - 1];
  //   const startDate = values[COLUMNS["Start Date"] - 1];
  //   const endDate = values[COLUMNS["End Date"] - 1];
  //   const requester = values[COLUMNS["Requester Email"] - 1];
  //   const reasonText = values[COLUMNS["Reason"] - 1];

  //   sheet.getRange(rowIndex, COLUMNS["Status"]).setValue('Rejected');
  //   sheet.getRange(rowIndex, COLUMNS["Stage"]).setValue('Rejected at ' + stage);
  //   sheet.getRange(rowIndex, COLUMNS["Note"]).setValue(note);

  //   const formattedStartDate = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
  //   const formattedEndDate = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
  //   const leaveDays = Math.ceil((new Date(endDate) - new Date(startDate)) / (1000 * 60 * 60 * 24)) + 1;

  //   const rejectionHtml = `
  //     <div style="font-family: Arial, sans-serif; color: #333;">
  //       <h3 style="color: #dc3545;">❌ Leave Request Rejected - ${name}</h3>
  //       <p><strong>Requester:</strong> ${name}</p>
  //       <p><strong>Department:</strong> ${department}</p>
  //       <p><strong>Leave Type:</strong> ${leaveType}</p>
  //       <p><strong>Dates:</strong> ${formattedStartDate} to ${formattedEndDate} (${leaveDays} day${leaveDays > 1 ? 's' : ''})</p>
  //       <p><strong>Reason:</strong> ${reasonText}</p>
  //       <p><strong>Rejected at Stage:</strong> ${stage}</p>
  //       <p><strong>Note:</strong> ${note}</p>
  //       <div style="margin-top: 20px; padding: 10px; background-color: #f8d7da; color: #721c24; border-left: 5px solid #f5c6cb;">
  //         <strong>Status:</strong> Rejected
  //       </div>
  //     </div>
  //   `;

  //   GmailApp.sendEmail({
  //     to: requester,
  //     subject: `Your Company Name Leave Request Rejected: ${name}`,
  //     htmlBody: rejectionHtml,
  //     name: "Your Company Name Approval System"
  //   });

  //   const resultTemplate = HtmlService.createTemplateFromFile('result');
  //   resultTemplate.status = "Rejected";
  //   resultTemplate.message = "Rejection submitted successfully with notes.";
  //   resultTemplate.color = "danger";
  //   return resultTemplate.evaluate().setTitle("Rejection Submitted");
  // }
  
  if (page === 'privacy') {
    return HtmlService.createHtmlOutputFromFile('privacy').setTitle('Privacy Policy');
  }

  if (page === 'terms') {
    return HtmlService.createHtmlOutputFromFile('terms').setTitle('Terms of Service');
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');

    if (!e || !e.parameter || !e.parameter.action || !e.parameter.row) { // Simpler check for form display
      return HtmlService.createTemplateFromFile('form').evaluate().setTitle("Leave Request Form");
    }

    const rowIndex = parseInt(e.parameter.row, 10);
    const action = e.parameter.action;
    // Note for rejection: If a note is crucial for rejection,
    // the email link should prompt the user or they should be instructed to add it to the URL.
    // e.g., &note=YourReasonHere. Or an intermediary HTML page could be used for a richer experience.
    const note = e.parameter.note || ''; // Approvers can manually add ?note=text to the URL if needed.
    // const stageFromParam = e.parameter.stage; // The stage link was clicked from

    if (isNaN(rowIndex) || rowIndex < 2 || !['approve', 'reject'].includes(action)) {
      return HtmlService.createHtmlOutput("Error: Invalid parameters. Please ensure the link is correct.").setTitle("Error");
    }

    // Lock to prevent concurrent modifications to the same row if possible (simple lock)
    const lock = LockService.getScriptLock();
    lock.waitLock(15000); // Wait up to 15 seconds for lock

    let nextStage = 'Final';
    let decision = action === 'approve' ? 'Approved' : 'Rejected';
    let currentStage = '';
    let name = '';
    let department = '';
    let leaveType = '';
    let requester = '';
    let startDate, endDate, reasonText;

    try {
        currentStage = sheet.getRange(rowIndex, COLUMNS.STAGE).getValue();
        const currentStatus = sheet.getRange(rowIndex, COLUMNS.STATUS).getValue();

        if (currentStatus !== 'Pending' && currentStatus !== '') { // '' might be for old entries before status was set
             // Allow processing if stage indicates it's waiting for this specific action, even if status was updated.
             // More robust check: has this specific stage already been decided?
             // For now, if status is Approved/Rejected, consider it final.
            if (currentStatus === 'Approved' || currentStatus === 'Rejected') {
                const html = HtmlService.createTemplateFromFile('result');
                html.action = currentStatus.toLowerCase();
                html.stage = currentStage;
                html.note = `This request was already processed as ${currentStatus}. ` + sheet.getRange(rowIndex, COLUMNS.NOTE).getValue();
                html.nextStage = "Final";
                return html.evaluate().setTitle("Request Processed");
            }
        }

        // Get the stageToken map
        const stageTokenColumnMap = {
          "SPV Approval": COLUMNS.SPV_TOKEN,
          "HR Review": COLUMNS.HR_TOKEN,
          "GM Review": COLUMNS.GM_TOKEN
        };

        // Get the `data` row and extract currentStage BEFORE using it
        const data = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

        // Now it's safe to use currentStage to look up token column
        const tokenColumn = stageTokenColumnMap[currentStage];

        // Get token from link
        const tokenToUse = e.parameter.token;
        const stageParam = e.parameter.stage; // spv/hr/gm

        // Validate inputs
        if (!tokenToUse || !tokenColumn || !stageParam) {
          return showErrorTokenPage("Unauthorized Access", "Hey Stranger, what are you doing here?");
        }

        // Normalize stage and validate token
        const validStage = {
          spv: COLUMNS.SPV_TOKEN,
          hr: COLUMNS.HR_TOKEN,
          gm: COLUMNS.GM_TOKEN
        }[stageParam.toLowerCase()];

        // Get email of current Google user
        const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();
        const spvEmail = (data[COLUMNS.SPV_EMAIL - 1] || "").toLowerCase();
        const hrEmail  = (data[COLUMNS.HR_EMAIL - 1] || "").toLowerCase();
        const gmEmail  = (data[COLUMNS.GM_EMAIL - 1] || "").toLowerCase();

        let expectedApprover = "";

        if (stageParam === "spv") expectedApprover = spvEmail;
        else if (stageParam === "hr") expectedApprover = hrEmail;
        else if (stageParam === "gm") expectedApprover = gmEmail;
        else return showErrorTokenPage("Invalid Stage", "Unknown stage: " + stageParam);

        if (currentUserEmail !== expectedApprover) {
          return showErrorTokenPage("Unauthorized Approver", `Whoopzz Whoopzz <code>${currentUserEmail}</code>, <br>You're not authorized to approve this request<br>for stage <b>${currentStage}</b>.`);
        }

        // Check saved token
        const savedToken = sheet.getRange(rowIndex, tokenColumn).getValue();

        if (validStage !== tokenColumn || savedToken !== tokenToUse || savedToken.endsWith("_used")) {
          return showErrorTokenPage("Access Denied", "Invalid or expired token. This action link has already been used or broken.");
        }

        // Block re-approval if already finalized
        if (["Approved", "Rejected"].includes(currentStatus)) {
          const html = HtmlService.createTemplateFromFile('result');
          html.action = currentStatus.toLowerCase();
          html.stage = currentStage;
          html.note = `This request was already processed as ${currentStatus}. ` + sheet.getRange(rowIndex, COLUMNS.NOTE).getValue();
          html.nextStage = "Final";
          return html.evaluate().setTitle("Request Processed");
        }

        // Invalidate the token after use
        sheet.getRange(rowIndex, tokenColumn).setValue(`${tokenToUse}_used`);

        name = sheet.getRange(rowIndex, COLUMNS.NAME).getValue();
        department = sheet.getRange(rowIndex, COLUMNS.DEPARTMENT).getValue();
        leaveType = sheet.getRange(rowIndex, COLUMNS.LEAVE_TYPE).getValue();
        requester = sheet.getRange(rowIndex, COLUMNS.REQUESTER_EMAIL).getValue();
        startDate = sheet.getRange(rowIndex, COLUMNS.START_DATE).getValue();
        endDate = sheet.getRange(rowIndex, COLUMNS.END_DATE).getValue();
        reasonText = sheet.getRange(rowIndex, COLUMNS.REASON).getValue();

        // Update decision details
        sheet.getRange(rowIndex, COLUMNS.DECISION).setValue(decision + " by " + currentStage); // Be more specific
        sheet.getRange(rowIndex, COLUMNS.NOTE).setValue(note);
        sheet.getRange(rowIndex, COLUMNS.DECISION_DATE).setValue(new Date());

        if (action === 'approve') {
          sheet.getRange(rowIndex, COLUMNS.STATUS).setValue('Pending'); // Keep pending until final approval

          switch(currentStage) {
            case "SPV Approval":
              if (leaveType === "Working From Home (WFH)") {
                finalizeRequest(rowIndex, decision, note, name, requester, "Approved by SPV");
              } else {
                nextStage = "HR Review";
                sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);

                // Generate a new token for HR
                const hrToken = generateRandomToken();

                // Save token in the sheet in HR_TOKEN column
                sheet.getRange(rowIndex, COLUMNS.HR_TOKEN).setValue(hrToken);

                // Pass same token to the email function
                sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.HR_EMAIL, nextStage, rowIndex, hrToken);
              }
              break;

            case "HR Review":
              // Best Use Cases Logic!
              // | **Submitter** | **Expected Approval Flow**            | **Expected Reject Flow**   | **Acc Flow** | **Reject Flow** |
              // | ------------- | ------------------------------------- | -------------------------- | ------------ | --------------- |
              // | Employee      | SPV → HR → Reporting (V)              | Rejected at the stage of   | Done         | Done            |
              // | SPV           | HR → GM → Reporting (V)               | ~~                         | Done         | Done            |
              // | HR            | GM → Reporting (V)                    |   ~~                       | Done         | Done            | 
              // | GM            | HR → Reporting (V)                    |     ~~                     | Done         | Done            |
              // | Unpaid Leave  | SPV(V) → HR(V) → GM(V) → Reporting (V)|       ~~                   | Done         | Done            | GM Unpaid needs to fix!
              // | WFH           | Emp -> SPV(V), SPV -> HR(V), HR -> GM(V), GM -> HR(V) |            | Done         | Done            |
              //
              //
              // GM Unpaid need to be fix ASAP == | GM Unpaid --> HR | not |GM Unpaid --> HR --> GM | , kyaaaaaaaaaaaa, mumet ndaseeee cuuookkk. 
              // Dahlah besok ajaa, capeee
              //
              // I don't know maybe this function still have bias on same stage, but whis is work fine.
              // Maybe for you the next person who saw this code, you can fix this better. Thanks
              // 
              // With Great Respect,
              // Galbatorix

              const allSpvEmails = Object.values(CONFIG.SPV_MAP);
              const isRequesterSPV = allSpvEmails.includes(requester);
              const isRequesterHR = requester.toLowerCase() === CONFIG.HR_EMAIL.toLowerCase();
              // const isRequesterGM = requester.toLowerCase() === CONFIG.GM_EMAIL.toLowerCase();
              const needsGM = (leaveType === "Unpaid Leave") || isRequesterSPV;

              // if (leaveType === "Unpaid Leave" && isRequesterGM) {
              //   // If GM submit Unpaid Leave it should be only end in HR Review, do not move to next stage!
              //   nextStage = "Final";
              //   sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
              //   finalizeRequest(rowIndex, decision, note, name, requester, "Approved by HR")
              // } else {
              //   finalizeRequest(rowIndex, decision, note, name, requester, "Approved by HR");
              // }

              // Need to fix GM flow is he/she pick unpaid-leave! GM only need to HR, Not HR --> GM(again LOL)!

              if (leaveType === "Working From Home (WFH)") {
                if (isRequesterHR) {
                  // WFH: HR submitted → GM next
                  nextStage = "GM Review";
                    sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
                    const gmToken = generateRandomToken();
                    sheet.getRange(rowIndex, COLUMNS.GM_TOKEN).setValue(gmToken);
                    sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.GM_EMAIL, nextStage, rowIndex, gmToken);
                } else {
                  // WFH: SPV submitted → HR → Reporting
                  finalizeRequest(rowIndex, decision, note, name, requester, "Approved by HR");
                }
              } else {
                // Regular flow
                nextStage = needsGM ? "GM Review" : "Final";
                sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
                const gmToken = generateRandomToken();
                sheet.getRange(rowIndex, COLUMNS.GM_TOKEN).setValue(gmToken);

                if (needsGM) {
                  sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.GM_EMAIL, nextStage, rowIndex, gmToken);
                } else {
                  finalizeRequest(rowIndex, decision, note, name, requester, "Approved by HR");
                }
              }
              break;

            case "GM Review":
              //const isRequesterHR = requester === CONFIG.HR_EMAIL;
              if (leaveType === "Working From Home (WFH)") {
                // WFH: HR submitted → GM → Reporting
                finalizeRequest(rowIndex, decision, note, name, requester, "Approved by GM");
              } else {
                nextStage = "Final";
                sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage); // Though it's final, setting stage to 'Final' or 'Completed'
                finalizeRequest(rowIndex, decision, note, name, requester, "Approved by GM");
              }
              break;

            default:
              // Should not happen if flow is correct
              nextStage = "Error in Workflow";
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage); 
              finalizeRequest(rowIndex, "Error", "Workflow error at stage: " + stage, name, requester, "Workflow Error");
              break;
          }
        } else { // Action is 'reject'
          sheet.getRange(rowIndex, COLUMNS.STATUS).setValue('Rejected');
          sheet.getRange(rowIndex, COLUMNS.STAGE).setValue('Rejected at ' + currentStage); // More specific
          const rejectionNote = note || `Rejected by ${currentStage}.`;

          const formattedStartDate = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
          const formattedEndDate = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
          const leaveDays = Math.ceil((new Date(endDate) - new Date(startDate)) / (1000 * 60 * 60 * 24)) + 1;

          const rejectionHtml = `
            <div style="font-family: Arial, sans-serif; color: #333;">
              <h3 style="color: #dc3545;">❌ Leave/WFH Request Rejected - ${name}</h3>
              <p><strong>Requester:</strong> ${name}</p>
              <p><strong>Department:</strong> ${department}</p>
              <p><strong>Leave Type:</strong> ${leaveType}</p>
              <p><strong>Dates:</strong> ${formattedStartDate} to ${formattedEndDate} (${leaveDays} day${leaveDays > 1 ? 's' : ''})</p>
              <p><strong>Reason:</strong> ${reasonText}</p>
              <p><strong>Rejected at Stage:</strong> ${currentStage}</p>
              <p><strong>Note:</strong> ${rejectionNote}</p>
              <div style="margin-top: 20px; padding: 10px; background-color: #f8d7da; color: #721c24; border-left: 5px solid #f5c6cb;">
                <strong>Status:</strong> Rejected
              </div>
            </div>
          `;
          // Rejection email to requester
          GmailApp.sendEmail(
            requester,
            `Your Company Name Leave/WFH Request Rejected: ${name}`,
            '',
            { 
              htmlBody: rejectionHtml,
              name: 'Your Company Name Approval System'
            }
          );

          // Rejection notification to reporting team
          //CONFIG.REPORTING_EMAILS.forEach(email => {
          //  GmailApp.sendEmail(email, `Leave Request Rejected - ${name}`, '', {
          //    htmlBody: rejectionHtml,
          //    name: 'Your Company Name Approval System'
          //  });
          //});
          nextStage = 'Final (Rejected)';
        }
    } finally {
        lock.releaseLock();
    }

    const html = HtmlService.createTemplateFromFile('result');
    html.action = action;
    html.stage = currentStage; // The stage that just made the decision
    html.note = note || (action === 'reject' ? "Rejected by " + currentStage : "Approved by " + currentStage);
    html.nextStage = nextStage;
    return html.evaluate().setTitle("Request Processed");

  } catch (err) {
    Logger.log("Error in doGet: " + err.toString() + "\nStack: " + err.stack);
    // Return a user-friendly error page
    let errorHtml = '<h1>Oops! Something went wrong.</h1>';
    errorHtml += '<p>We encountered an error while processing your request. This could be due to the request being outdated, already processed, or a temporary issue.</p>';
    errorHtml += '<p>Please try again later or contact support if the problem persists.</p>';
    errorHtml += '<h1><small>Error details (for support): ' + escapeHtml(err.toString()) + '</small></h1>';
    errorHtml += '<h2><small>Contact Support (for support): Please screenshoot this error and send it to Your Support Name</small></h2>';
    return HtmlService.createHtmlOutput(errorHtml).setTitle("Processing Error");
  }
}

function submitRequest(name, email, department, leaveType, startDate, endDate, reason) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("Requests");
    const spvEmail = CONFIG.SPV_MAP[department] || CONFIG.GM_EMAIL; // Default to GM if SPV not found
    const hrEmail = CONFIG.HR_EMAIL;
    const gmEmail = CONFIG.GM_EMAIL;

    const firstDate = parseDMYDate(startDate);
    const lastDate = parseDMYDate(endDate);

    if (!firstDate || !lastDate) {
      Logger.log("Invalid date format received: " + startDate + ", " + endDate);
      throw new Error("Invalid date format. Please use DD-MM-YYYY.");
    }

    if (lastDate < firstDate) {
      throw new Error("Are you blind or what?<br>Have you make a coffee today?<br>Last day of leave cannot be before the first day.");
    }

    // Determine if the requester is the SPV of their department
    // Javascript?? this is Jawascript hahahahaha
    const ikiEmailnyaSPVngertiOra = (email.toLowerCase() === spvEmail.toLowerCase());
    const ikiEmailnyaGMngertiOra = (email.toLowerCase() === gmEmail.toLowerCase());
    const ikiEmailnyaHRngertiOra = (email.toLowerCase() === hrEmail.toLowerCase());

    let stage;
    let approvalEmail;

    const saumpaniHRprei = ["Annual Leave", "Bereavement Leave", "Career Leave", "Ceremony Leave", "Sick Leave", "Unpaid Leave", "Other", "Working From Home (WFH)"]; // If SPV as a HR(Dyah Retno) submit all leave it must set nextstage to GM!

    if (ikiEmailnyaSPVngertiOra) {
      // This is case when your HRGA is also the SPV of HRGA, you know what I mean, chiiizzzzz ~xD
      // So if your HRGA is on different posisition you're not need this
      // That's mean you gonna change the flow from here
      if (saumpaniHRprei.includes(leaveType) && ikiEmailnyaHRngertiOra) {
        stage = "GM Review";
        approvalEmail = gmEmail;
      } else {
        stage = "HR Review";
        approvalEmail = hrEmail;
      }
    } else if (ikiEmailnyaGMngertiOra) {
      stage = "HR Review";
      approvalEmail = hrEmail;
    } else {
      if (leaveType === "Working From Home (WFH)") {
        stage = "SPV Approval"; // employee WFH should go to SPV
        approvalEmail = spvEmail;
      } else {
        stage = "SPV Approval";
        approvalEmail = spvEmail;
      }
    }

    const spvToken = generateRandomToken();
    const hrToken = generateRandomToken();
    const gmToken = generateRandomToken();

    let tokenToUse;
      if (stage === "SPV Approval") {
        tokenToUse = spvToken;
      } else if (stage === "HR Review") {
        tokenToUse = hrToken;
      } else if (stage === "GM Review") {
        tokenToUse = gmToken;
      }

    // Append the request to the sheet
    const newRow = sheet.appendRow([
    new Date(), name, department, leaveType,
    firstDate, lastDate, reason,
    "Pending", email, spvEmail, hrEmail, gmEmail, stage,
    "", "", new Date(), // Decision, Note, Decision Date
    "", "", "",         // SPV_DECISION, HR_DECISION, GM_DECISION
    spvToken, hrToken, gmToken
  ]);

    const rowIndex = sheet.getLastRow();

    // Send confirmation to requester
    sendSubmissionConfirmation(email, name, leaveType, firstDate, lastDate, reason, stage);

    // Send approval request to next stage
    sendApprovalEmail(name, leaveType, firstDate, lastDate, reason, approvalEmail, stage, rowIndex, tokenToUse);

    return "Success";

  } catch (e) {
    Logger.log("Submit error: " + e.toString() + " Stack: " + e.stack);
    return e.message || "Submission failed due to a server error. Please try again.";
  }
}

function escapeHtml(text) {
  if (!text) return "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function finalizeRequest(row, decision, note, name, requesterEmail, finalApprovalStageNote) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Requests");

  // FINAL NOTE: Handle fallback early before any usage
  const finalNote = note || finalApprovalStageNote || "No Notes";

  // Update status and stage
  sheet.getRange(row, COLUMNS.STATUS).setValue(decision); // 'Approved' or 'Rejected'
  sheet.getRange(row, COLUMNS.STAGE).setValue('Completed');
  sheet.getRange(row, COLUMNS.NOTE).setValue(finalNote);

  // Get detailed request data
  const leaveType = sheet.getRange(row, COLUMNS.LEAVE_TYPE).getValue();
  const startDate = sheet.getRange(row, COLUMNS.START_DATE).getValue();
  const endDate = sheet.getRange(row, COLUMNS.END_DATE).getValue();
  const reason = sheet.getRange(row, COLUMNS.REASON).getValue();
  const spvStatus = sheet.getRange(row, COLUMNS.SPV_DECISION).getValue();
  const hrStatus = (COLUMNS.HR_DECISION) ? sheet.getRange(row, COLUMNS.HR_DECISION).getValue() : '';
  const gmStatus = (COLUMNS.GM_DECISION) ? sheet.getRange(row, COLUMNS.GM_DECISION).getValue() : '';

  const days = (new Date(endDate) - new Date(startDate)) / (1000 * 60 * 60 * 24) + 1;

  // Final Notification to Requester
  const template = HtmlService.createTemplateFromFile('finalNotification');
  template.name = name;
  template.leaveType = leaveType;
  template.startDate = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
  template.endDate = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
  template.totalDays = days;
  template.reason = reason;
  template.spvStatus = spvStatus;
  template.hrStatus = hrStatus;
  template.gmStatus = gmStatus;
  template.finalDecision = decision;
  template.finalNote = finalNote;

  const htmlBody = template.evaluate().getContent();

  // Final decision email to requester (Approved/Rejected)
  GmailApp.sendEmail(requesterEmail, `Your Company Name Leave Request ${decision}: ${name}`, '', {
    htmlBody: htmlBody,
    name: 'Your Company Name Approval System'
  });

  // Notify Reporting Team
  const department = sheet.getRange(row, COLUMNS.DEPARTMENT).getValue();
  const calendarTitle = `${name} - ${leaveType}`;
  const calendarDescription = `Leave Request\nRequester: ${name}\nDepartment: ${department}\nType: ${leaveType}\nDecision: ${decision}\nNote: ${finalNote}`;
  const calendarLink = generateCalendarLink(calendarTitle, startDate, endDate, calendarDescription);

  const formattedStart = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
  const formattedEnd = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "dd-MM-yyyy");  

  const reportingHtml = `
    <div style="font-family: Arial, sans-serif; padding: 20px; background-color: #f8f9fa; color: #212529; border: 1px solid #dee2e6; border-radius: .25rem;">
      <h3 style="margin-top:0;">Leave/WFH Request <span style="color:${decision === 'Approved' ? '#28a745' : '#dc3545'};">${decision}</span> - ${name}</h3>
      <p><strong>Requester:</strong> ${name}</p>
      <p><strong>Department:</strong> ${department}</p>
      <p><strong>Leave Type:</strong> ${leaveType}</p>
      <p><strong>Dates:</strong> ${formattedStart} to ${formattedEnd} (${days} day${days > 1 ? 's' : ''})</p>
      <p><strong>Reason:</strong> ${reason || 'N/A'}</p>
      <p><strong>Note:</strong> ${finalNote}</p>
      <div style="margin-top: 20px;">
        <a href="${calendarLink}" target="_blank" style="display:inline-block; padding:10px 20px; background-color:#007bff; color:white; text-decoration:none; border-radius:.25rem;">➕     Add to Google Calendar</a>
      </div>
    </div>
  `;

  // Final decision email to reporting team
  CONFIG.REPORTING_EMAILS.forEach(email => {
    GmailApp.sendEmail(email, `Leave/WFH Request ${decision} - ${name}`, '', { 
      htmlBody: reportingHtml,
      name: 'Your Company Name Approval System'
    });
  });
}

function generateCalendarLink(title, startDate, endDate, description) {
  const formatDate = date => Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyyMMdd");

  const start = formatDate(startDate);
  const end = formatDate(new Date(new Date(endDate).getTime() + 86400000)); // +1 day to make it inclusive

  const details = encodeURIComponent(description);
  const calTitle = encodeURIComponent(title);

  return `https://www.google.com/calendar/render?action=TEMPLATE&text=${calTitle}&dates=${start}/${end}&details=${details}`;
}

function sendApprovalEmail(name, leaveType, startDate, endDate, reason, approverEmail, stage, row, tokenToUse) {
  const baseUrl = ScriptApp.getService().getUrl();

  const template = HtmlService.createTemplateFromFile("emailtemplate");
  template.name = name;
  template.leaveType = leaveType;
  template.startDate = formatDate(startDate);
  template.endDate = formatDate(endDate);
  template.reason = reason;
  template.stage = stage;
  template.baseUrl = baseUrl;
  template.row = row;

  // 🧠 Extract normalized stage (lowercase) → Used for URL param
  const shortStage = stage.toLowerCase().includes("spv") ? "spv"
                   : stage.toLowerCase().includes("hr") ? "hr"
                   : stage.toLowerCase().includes("gm") ? "gm"
                   : "unknown"; // fallback

  if (!tokenToUse || shortStage === "unknown") {
    Logger.log(`❗ Missing token or invalid stage in sendApprovalEmail for stage: ${stage}, row: ${row}`);
    return;
  }

  // ✅ Safe & encoded approval/rejection URLs
  const encodedToken = encodeURIComponent(tokenToUse);
  const noteText = `Approved at ${stage}`;
  template.approveUrl = `${baseUrl}?action=approve&stage=${shortStage}&row=${row}&token=${encodedToken}&note=${encodeURIComponent(noteText)}`;
  template.rejectUrl  = `${baseUrl}?action=reject&stage=${shortStage}&row=${row}&token=${encodedToken}&note=`; // will be filled in by user

  try {
    MailApp.sendEmail({
      to: approverEmail,
      subject: `[Action Required] Leave/WFH Request: ${name} (${stage})`,
      htmlBody: template.evaluate().getContent(),
      name: "Your Company Name Approval System"
    });
    Logger.log(`✅ Approval email sent to ${approverEmail} for ${name} at ${stage}`);
  } catch (e) {
    Logger.log(`❌ Failed to send approval email to ${approverEmail} for row ${row}. Error: ${e}`);
  }
}

function sendSubmissionConfirmation(email, name, leaveType, startDate, endDate, reason, stage) {
  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:600px; margin:auto; border:1px solid #ddd; padding:20px;">
      <h2 style="color:#2c3e50; border-bottom:1px solid #eee; padding-bottom:10px;">Leave Request Submitted</h2>
      <p>Dear ${name},</p>
      <p>Your leave request has been successfully submitted and is now pending <strong>${stage}</strong>.</p>
      
      <table style="width:100%;border-collapse:collapse;margin:20px 0;">
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;width:30%;"><strong>Leave Type</strong></td><td style="padding:8px;border:1px solid #ddd;">${leaveType}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;"><strong>Dates</strong></td><td style="padding:8px;border:1px solid #ddd;">${formatDate(startDate)} to ${formatDate(endDate)}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;"><strong>Reason</strong></td><td style="padding:8px;border:1px solid #ddd;">${escapeHtml(reason)}</td></tr>
      </table>
      
      <p>You will receive further email notifications as your request is processed.</p>
      <p style="color:#7f8c8d;font-size:0.9em;margin-top:20px;">Thank you,<br/>Your Company Name Approval System</p>
    </div>
  `;
  try {
    MailApp.sendEmail({
      to: email,
      subject: "Your Company Name Leave/WFH Request Submission Confirmation",
      htmlBody: htmlBody,
      name: "Your Company Name Approval System" 
    });
  } catch (e) {
    Logger.log("Failed to send confirmation email to " + email + ". Error: " + e.toString());
  }
}