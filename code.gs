// Global Emails Configuration
const CONFIG = {
  REPORTING_EMAILS: ["reporting.1@yourcompany.com", "reporting.2@yourcompany.com"],
  GM_EMAIL: "gm.email@yourcompany.com",
  HR_EMAIL: "hr.email@yourcompany.com",
  SPV_MAP: {
    "Management Division": "management@yourcompany.com",
    "Finance": "finance@yourcompany.com",
    "Helper": "helper@yourcompany.com",
    "HR": "hr.leader@yourcompany.com",
    "Security Pos": "security.pos@yourcompany.com",
    "IT": "it.department@yourcompany.com",
    // You can add your all supervisor and leader here.
    //Feel free to add them.
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
  STAGE: 12,
  DECISION: 13,
  NOTE: 14,
  DECISION_DATE: 15,
  SPV_DECISION: 16,
  HR_DECISION: 17,
  GM_DECISION: 18
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

function doGet(e) {
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
            case 'SPV Approval':
              nextStage = 'HR Review';
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
              sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.HR_EMAIL, nextStage, rowIndex);
              break;

            case 'HR Review':
              // If HR department staff requests leave (even not unpaid), it should go to GM.
              // If HR member submits for someone else, it follows normal flow.
              // Assuming 'department' is the requester's department.
              const needsGM = (leaveType === 'Unpaid Leave') || (department === 'HR');
              nextStage = needsGM ? 'GM Review' : 'Final';
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
              
              if (needsGM) {
                sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.GM_EMAIL, nextStage, rowIndex);
              } else {
                finalizeRequest(rowIndex, decision, note, name, requester, "Approved by HR");
              }
              break;

            case 'GM Review':
              nextStage = 'Final';
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage); // Though it's final, setting stage to 'Final' or 'Completed'
              finalizeRequest(rowIndex, decision, note, name, requester, "Approved by GM");
              break;
            default:
              // Should not happen if flow is correct
              nextStage = 'Error in Workflow';
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
              finalizeRequest(rowIndex, "Error", "Workflow error at stage: " + currentStage, name, requester, "Workflow Error");
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
              <h3 style="color: #dc3545;">❌ Leave Request Rejected - ${name}</h3>
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
          // Send to requester
          GmailApp.sendEmail(
            requester,
            `ONEderland Leave Request Rejected: ${name}`,
            '',
            { htmlBody: rejectionHtml }
          );
          // Notify reporting team
          CONFIG.REPORTING_EMAILS.forEach(email => {
            GmailApp.sendEmail(email, `Leave Request Rejected - ${name}`, '', {
              htmlBody: rejectionHtml
            });
          });
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
    errorHtml += '<p><small>Error details (for support): ' + escapeHtml(err.toString()) + '</small></p>';
    return HtmlService.createHtmlOutput(errorHtml).setTitle("Processing Error");
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

  // Final Notification to Requester (Styled)
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

  GmailApp.sendEmail(requesterEmail, `ONEderland Leave Request ${decision}: ${name}`, '', {
    htmlBody: htmlBody
  });

  // Notify Reporting Team with Bootstrap-like styled email
  const department = sheet.getRange(row, COLUMNS.DEPARTMENT).getValue();
  const calendarTitle = `${name} - ${leaveType}`;
  const calendarDescription = `Leave Request\nRequester: ${name}\nDepartment: ${department}\nType: ${leaveType}\nDecision: ${decision}\nNote: ${finalNote}`;
  const calendarLink = generateCalendarLink(calendarTitle, startDate, endDate, calendarDescription);

  const formattedStart = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
  const formattedEnd = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "dd-MM-yyyy");

  const reportingHtml = `
    <div style="font-family: Arial, sans-serif; padding: 20px; background-color: #f8f9fa; color: #212529; border: 1px solid #dee2e6; border-radius: .25rem;">
      <h3 style="margin-top:0;">Leave Request <span style="color:${decision === 'Approved' ? '#28a745' : '#dc3545'};">${decision}</span> - ${name}</h3>
      <p><strong>Requester:</strong> ${name}</p>
      <p><strong>Department:</strong> ${department}</p>
      <p><strong>Leave Type:</strong> ${leaveType}</p>
      <p><strong>Dates:</strong> ${formattedStart} to ${formattedEnd} (${days} day${days > 1 ? 's' : ''})</p>
      <p><strong>Reason:</strong> ${reason || 'N/A'}</p>
      <p><strong>Note:</strong> ${finalNote}</p>
      <div style="margin-top: 20px;">
        <a href="${calendarLink}" target="_blank" style="display:inline-block; padding:10px 20px; background-color:#007bff; color:white; text-decoration:none; border-radius:.25rem;">➕ Add to Google Calendar</a>
      </div>
    </div>
  `;

  CONFIG.REPORTING_EMAILS.forEach(email => {
    GmailApp.sendEmail(email, `Leave Request ${decision} - ${name}`, '', { htmlBody: reportingHtml });
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
      throw new Error("Last day of leave cannot be before the first day.");
    }

    // Determine if the requester is the SPV of their department
    const isSPVSubmitting = (email.toLowerCase() === spvEmail.toLowerCase());

    let stage;
    let approvalEmail;

    if (isSPVSubmitting) {
      // If SPV is submitting their own request: skip HR, go directly to GM
      stage = "GM Review";
      approvalEmail = gmEmail;
    } else {
      // Normal flow: non-SPV submits → goes to SPV
      stage = "SPV Approval";
      approvalEmail = spvEmail;
    }

    // Append the request to the sheet
    const newRow = sheet.appendRow([
      new Date(), name, department, leaveType,
      firstDate, lastDate, reason,
      "Pending", email, spvEmail, hrEmail, stage,
      "", "", new Date() // Decision, Note, Decision Date (initially submission timestamp)
    ]);

    const rowIdx = sheet.getLastRow();

    // Send confirmation to requester
    sendSubmissionConfirmation(email, name, leaveType, firstDate, lastDate, reason, stage);

    // Send approval request to next stage
    sendApprovalEmail(name, leaveType, firstDate, lastDate, reason, approvalEmail, stage, rowIdx);

    return "Success";

  } catch (e) {
    Logger.log("Submit error: " + e.toString() + " Stack: " + e.stack);
    return e.message || "Submission failed due to a server error. Please try again.";
  }
}


function sendApprovalEmail(name, leaveType, startDate, endDate, reason, approverEmail, stage, row) {
  const baseUrl = ScriptApp.getService().getUrl();
  
  const template = HtmlService.createTemplateFromFile("emailtemplate"); // Ensure this is 'emailtemplate' not 'EmailTemplate' if filename is lowercase
  template.name = name;
  template.leaveType = leaveType;
  template.startDate = formatDate(startDate);
  template.endDate = formatDate(endDate);
  template.reason = reason;
  template.stage = stage;
  // Note: For rejections, if a detailed note is required, the approver would currently
  // need to manually add it to the URL by replacing the end of '&note='
  // or an intermediary page/prompt system would be needed.
  template.approveUrl = `${baseUrl}?action=approve&stage=${encodeURIComponent(stage)}&row=${row}&note=ApprovedAt${encodeURIComponent(stage)}`; // Basic approval note
  template.rejectUrl = `${baseUrl}?action=reject&stage=${encodeURIComponent(stage)}&row=${row}&note=`; // Blank note, expect user to fill if desired

  try {
    MailApp.sendEmail({
      to: approverEmail,
      subject: `[Action Required] Leave Request: ${name} (${stage})`,
      htmlBody: template.evaluate().getContent(),
      name: "ONEderland Approval System"
    });
  } catch (e) {
    Logger.log("Failed to send approval email to " + approverEmail + " for row " + row + ". Error: " + e.toString());
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
      <p style="color:#7f8c8d;font-size:0.9em;margin-top:20px;">Thank you,<br/>ONEderland Approval System</p>
    </div>
  `;
  try {
    MailApp.sendEmail({
      to: email,
      subject: "ONEderland Leave Request Submission Confirmation",
      htmlBody: htmlBody,
      name: "ONEderland Approval System" 
    });
  } catch (e) {
    Logger.log("Failed to send confirmation email to " + email + ". Error: " + e.toString());
  }
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