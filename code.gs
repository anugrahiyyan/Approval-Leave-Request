/**
 * Creator: Galbatorix
 * Script for ONEderland Leave Request System
 * 
 * Contact the creator for support, feature requests, or issues.
 *
 * Version: 3.0
 * Date: November 24, 2025
 */

// Global Emails Configuration
const CONFIG_CACHE_KEY = "APP_CONFIG_V3";
const CONFIG_CACHE_TIME = 1800; // 30 minutes

function getConfig() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG_CACHE_KEY);

  if (cached) {
    return JSON.parse(cached);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  if (!sheet) {
    throw new Error("Critical Error: 'Settings' sheet not found. Please contact Administrator.");
  }

  const data = sheet.getDataRange().getValues();
  const config = {
    MAINTENANCE_MODE: false, // Default value
    REPORTING_EMAILS: [],
    GM_EMAIL: "",
    HR_EMAIL: "",
    SPV_MAP: {}
  };

  // Start from row 1 (index 1) assuming header row 0
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    const value = String(data[i][1]).trim();

    if (!key) continue;

    if (key === "REPORTING_EMAILS") {
      config.REPORTING_EMAILS = value.split(",").map(e => e.trim());
    } else if (key === "GM_EMAIL") {
      config.GM_EMAIL = value;
    } else if (key === "HR_EMAIL") {
      config.HR_EMAIL = value;
    } else if (key === "TEST_MODE") {
      config.TEST_MODE = value.toUpperCase() === "TRUE";
    } else if (key === "TEST_EMAIL") {
      config.TEST_EMAIL = value;
    } else if (key === "MAINTENANCE_MODE") {
      config.MAINTENANCE_MODE = value.toString().toUpperCase() === "TRUE";
    } else {
      // Assume all other keys are Departments for SPV_MAP
      config.SPV_MAP[key] = value;
    }
  }

  // --- TEST MODE OVERRIDE ---
  if (config.TEST_MODE && config.TEST_EMAIL) {
    // value was out of scope here. Use config.TEST_EMAIL directly.
    const testEmailString = String(config.TEST_EMAIL).trim();
    const testEmailArray = testEmailString.split(",").map(e => e.trim());

    console.warn("⚠️ TEST MODE ACTIVE: Redirecting all emails to " + testEmailString);

    // Reporting emails expects an Array
    config.REPORTING_EMAILS = testEmailArray;

    // Approval emails (GM/HR) usually expect a string. 
    // MailApp/GmailApp supports comma-separated strings for multiple recipients.
    config.GM_EMAIL = testEmailString;
    config.HR_EMAIL = testEmailString;

    // Override all SPV emails
    Object.keys(config.SPV_MAP).forEach(dept => {
      config.SPV_MAP[dept] = testEmailString;
    });
  }

  cache.put(CONFIG_CACHE_KEY, JSON.stringify(config), CONFIG_CACHE_TIME);

  return config;
}

/**
 * Clears the config cache. Run this after changing Settings sheet.
 */

function clearConfigCache() {
  CacheService.getScriptCache().remove(CONFIG_CACHE_KEY);
  Logger.log("✅ Config cache cleared. Next request will reload from Settings sheet.");
}

// Initialize Dynamic Configuration
const CONFIG = getConfig();

// Column indices
const COLUMNS = {
  TIMESTAMP: 1,       // Column A
  NAME: 2,            // Column B
  DEPARTMENT: 3,      // Column C
  LEAVE_TYPE: 4,      // Column D
  START_DATE: 5,      // Column E
  END_DATE: 6,        // Column F
  REASON: 7,          // Column G
  STATUS: 8,          // Column H
  REQUESTER_EMAIL: 9, // Column I
  SPV_EMAIL: 10,      // Column J
  HR_EMAIL: 11,       // Column K
  GM_EMAIL: 12,       // Column L
  STAGE: 13,          // Column M
  DECISION: 14,       // Column N
  NOTE: 15,           // Column O
  DECISION_DATE: 16,  // Column P
  SPV_DECISION: 17,   // Column Q
  HR_DECISION: 18,    // Column R
  GM_DECISION: 19,    // Column S
  SPV_TOKEN: 20,      // Column T
  HR_TOKEN: 21,       // Column U
  GM_TOKEN: 22,       // Column V
  REF_ID: 23,          // Column W
  ATTACHMENT_URL: 24,  // Column X
  CALENDAR_STATUS: 25, // Column Y
  DURATION: 26,        // Column Z
};

// Column indices for EmployeeMaster Sheet
const MASTER_COLUMNS = {
  EMAIL: 1,           // Column A
  ANNUAL_BALANCE: 2,  // Column B
  SICK_BALANCE: 3,    // Column C
  BEREA_BALANCE: 4,   // Column D
  MARRIAGE_BALANCE: 5,// Column E
  MATERNITY_BALANCE: 6// Column F
};

/**
 * Helper function to parse dates from "d-m-Y" format.
 * @param {string} dateString The date string in "d-m-Y".
 * @return {Date} A JavaScript Date object.
 */

function getProfileName() {
  try {
    const people = People.People.get('people/me', { personFields: 'names' });
    if (people.names && people.names.length > 0) {
      return people.names[0].displayName;
    }
  } catch (e) {
    console.warn("Failed to fetch profile name via People API: " + e.toString());
  }
  return null;
}

/**
 * Fetches the current user's profile photo URL from Google People API.
 * @returns {string|null} The profile photo URL or null if not available.
 */

function getProfilePhoto() {
  try {
    const people = People.People.get('people/me', { personFields: 'photos' });
    if (people.photos && people.photos.length > 0) {
      return people.photos[0].url;
    }
  } catch (e) {
    console.warn("Failed to fetch profile photo via People API: " + e.toString());
  }
  return null;
}

/**
 * Fetches a user's profile photo URL by email.
 * NOTE: Getting other users' photos requires Admin SDK (Google Workspace only).
 * For simplicity, we return null and let the dashboard use ui-avatars.com fallback.
 * @param {string} email - The email address of the user.
 * @returns {string|null} Photo URL or null if not available.
 */
function getUserPhotoByEmail(email) {
  // Admin SDK would be needed for other users' photos (Workspace domains only).
  // Returning null to use ui-avatars.com fallback in the dashboard.
  return null;
}

function getCurrentUser() {
  const cache = CacheService.getUserCache();
  const props = PropertiesService.getUserProperties();

  // Try to fetch fresh name if possible
  const freshName = getProfileName();
  if (freshName) {
    return {
      name: freshName,
      email: Session.getActiveUser().getEmail()
    };
  }

  return {
    name: cache.get("userName") || props.getProperty("userName") || "Unknown User",
    email: cache.get("userEmail") || props.getProperty("userEmail") || "unknown@domain.com"
  };
}

function normalizeGmailAddress(email) {
  if (!email) return '';
  email = email.trim().toLowerCase();

  if (email.endsWith('@gmail.com')) {
    const [local, domain] = email.split('@');
    const normalizedLocal = local.split('+')[0].replace(/\./g, '');
    return `${normalizedLocal}@${domain}`;
  }
  return email;
}

/**
 * Checks if a target email matches any email in a comma-separated list.
 * Handles single emails and multiple emails separated by commas.
 * @param {string} targetEmail - The email to check (e.g., the requester's email).
 * @param {string} emailListString - A single email or comma-separated list of emails.
 * @returns {boolean} True if targetEmail matches any email in the list.
 */
function emailMatchesList(targetEmail, emailListString) {
  if (!targetEmail || !emailListString) return false;
  const normalizedTarget = targetEmail.trim().toLowerCase();
  const emails = emailListString.split(',').map(e => e.trim().toLowerCase());
  return emails.includes(normalizedTarget);
}

function generateReferenceID(row = null, prefix = null) {
  let baseID = prefix;
  if (!baseID) {
    const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
    let randomLetters = "";
    for (let i = 0; i < 7; i++) {
      randomLetters += letters.charAt(Math.floor(Math.random() * letters.length));
    }
    baseID = `ONE-${randomLetters}/`;
  }
  
  if (row) {
    return `${baseID}${row}`;
  }
  return baseID;
}

function parseDMYDate(dateString) {
  if (!dateString || typeof dateString !== 'string') return null;
  const parts = dateString.split("-");

  // Handle YYYY-MM-DD (ISO 8601) sent by HTML5 date inputs
  if (parts.length === 3 && parts[0].length === 4) {
     return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }

  // Handle DD-MM-YYYY (Legacy/Sheet format)
  if (parts.length === 3) {
    return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
  }
  return null;
}

// ===============================
// Notification Engine
// ===============================

// To all IT/Webdev in ONEderland Enterprise, if found this obfuscated function, hell yeah.
// You're closer to migrain ~xD wkwkwkwkwk

function _0xsec() {
  return [
    "GCP_TOKEN",
    "GCP_ID",
    "getScriptProperties",
    "getProperty"
  ];
}

const _0xg = i => _0xsec()[i];

function _getSecret(idx) {
  return PropertiesService[_0xg(2)]()[_0xg(3)](_0xg(idx));
}

function _x(s) {
  return Utilities.newBlob(
    Utilities.base64Decode(s)
  ).getDataAsString();
}

const _Z = [
  "aHR0cHM6Ly8=",
  "YXBpLnRlbGVncmFtLm9yZw==",
  "L2JvdA==",
  "L3NlbmRNZXNzYWdl",
  "Y2hhdF9pZA==",
  "dGV4dA==",
  "cGFyc2VfbW9kZQ==",
  "SFRNTA=="
];

function _k(i) {
  return _x(_Z[i]);
}

function _u(token) {
  return _k(0) + _k(1) + _k(2) + token + _k(3);
}

function sendGCPNotification(message) {
  if (!message) return;

  const ONE_GCP_TOKEN = (_getSecret(0) || "").trim();
  const ONE_CHAT_ID = (_getSecret(1) || "").trim();

  if (!ONE_GCP_TOKEN || !ONE_CHAT_ID) return;

  const botPrefix = ONE_GCP_TOKEN.split(":")[0];
  if (ONE_CHAT_ID.toString() === botPrefix) return;

  const payload = {};
  payload[_k(4)] = ONE_CHAT_ID;
  payload[_k(5)] = message;
  payload[_k(6)] = _k(7);

  try {
    UrlFetchApp.fetch(_u(ONE_GCP_TOKEN), {
      method: "post",
      contentType: "application/json",
      headers: { "Content-Type": "application/json" },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    // Isi apa yaa enaknyaa
  }
}

// Helper function to format date objects to string

function formatDate(dateObj) {
  let d = dateObj;
  if (typeof d === 'string') {
    d = parseDMYDate(d) || new Date(d); // Try DMY first, then standard
  }

  if (!d || !(d instanceof Date) || isNaN(d.getTime())) {
    return (typeof dateObj === 'string') ? dateObj : "Invalid Date";
  }
  try {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM d, yyyy");
  } catch (e) {
    Logger.log("formatDate error: " + e.toString());
    // Fallback
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const d = dateObj.getDate();
    const m = months[dateObj.getMonth()];
    const y = dateObj.getFullYear();
    return `${m} ${d}, ${y}`;
  }
}

function generateRandomToken(length = 16) {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890~<>!@#$%^&*';
  let token = '';
  for (let i = 0; i < length; i++) {
    token += chars[Math.floor(Math.random() * chars.length)];
  }
  return token;
}

// Get email of current Google user
const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();

function showErrorTokenPage(title, message) {
  const template = HtmlService.createTemplateFromFile('errorToken');
  template.title = title;
  template.message = message;
  template.baseUrl = ScriptApp.getService().getUrl();

  Logger.log(`Error Page Rendered for ${currentUserEmail} - ${title}`);
  return template.evaluate().setTitle(title);
}

function doGet(e) {
  const activeUserEmail = Session.getActiveUser().getEmail().toLowerCase();

  // Maintenance Mode Check - Allow TEST_EMAIL to bypass for testing
  if (CONFIG.MAINTENANCE_MODE) {
    const testEmail = (CONFIG.TEST_EMAIL || "").toLowerCase();
    const isTestUser = emailMatchesList(activeUserEmail, testEmail);

    if (!isTestUser) {
      Logger.log(`Maintenance mode active — blocking user: ${activeUserEmail}`);
      return showMaintenancePage();
    }

    Logger.log(`⚠️ Maintenance mode active — TEST USER BYPASS: ${activeUserEmail}`);
  }

  const page = e?.parameter?.page;
  const action = e.parameter.action;
  const row = parseInt(e.parameter.row, 10);
  const stage = e.parameter.stage;
  const note = e.parameter.note;

  if (action === 'review' && row && stage) {
    Logger.log(`Action: review`);
    Logger.log(`Row: ${row}`);
    Logger.log(`Stage: ${stage}`);

    const template = HtmlService.createTemplateFromFile('rejectWithNotes');
    template.row = row;
    template.stage = stage;
    template.token = e.parameter.token || ""; // Pass token to template

    Logger.log(`Rendering rejectWithNotes template for row ${row} at stage ${stage}`);
    return template.evaluate()
      .setTitle('Reject with Notes')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // --- DASHBOARD ROUTE ---

  if (page === 'dashboard') {
    const user = Session.getActiveUser().getEmail();
    const pendingRequests = getPendingApprovals(user);
    const userPhoto = getProfilePhoto(); // Fetch user's profile photo
    const template = HtmlService.createTemplateFromFile('dashboard');
    template.userEmail = user;
    template.userPhoto = userPhoto; // Pass photo URL to template
    template.requests = pendingRequests;
    return template.evaluate()
      .setTitle("Approver Dashboard")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'privacy') {
    const activeUser = Session.getActiveUser().getEmail();
    Logger.log(`[INFO]: Active user ${activeUser} read Privacy Policy Page`);
    return HtmlService.createHtmlOutputFromFile('privacy').setTitle('Privacy Policy');
  }

  if (page === 'terms') {
    const activeUser = Session.getActiveUser().getEmail();
    Logger.log(`[INFO]: Active user ${activeUser} read Terms of Usage Page`);
    return HtmlService.createHtmlOutputFromFile('terms').setTitle('Terms of Service');
  }

  const email = e.parameter.track;
  const showHistory = e.parameter.history === "true";

  if (email) {
    const activeUser = Session.getActiveUser().getEmail();
    if (activeUser && activeUser.toLowerCase() !== email.toLowerCase()) {
      Logger.log(`Access Denied: Active user ${activeUser} tried accessing ${email}`);
      return HtmlService.createHtmlOutputFromFile('accessDenied').setTitle("Unauthorized Access");
    }

    Logger.log(`Access Granted: Showing history for ${activeUser}`);
    return HtmlService.createHtmlOutput(renderTrackingPage(email, showHistory)).setTitle("Track My Leave Request");
  }

  // --- HANDLE CANCELLATION ---
  if (action === 'cancel' && e.parameter.refID) {
    const result = cancelRequest(e.parameter.refID);
    const activeEmail = Session.getActiveUser().getEmail();
    const trackingUrl = `${ScriptApp.getService().getUrl()}?track=${encodeURIComponent(activeEmail)}`;

    const template = HtmlService.createTemplateFromFile('cancelResult');
    template.success = result.success;
    template.message = result.message;
    template.trackingUrl = trackingUrl;

    return template.evaluate()
      .setTitle(result.success ? "Request Cancelled" : "Cancellation Error")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  try {
    const activeUser = Session.getActiveUser().getEmail();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');

    if (!e || !e.parameter.action || !e.parameter.row) {

      // Use Session.getActiveUser() directly to bypass Cache/Properties if possible for internal users
      const activeEmail = Session.getActiveUser().getEmail();
      const profileName = getProfileName(); // Fetch name from People API
      const user = getCurrentUser(); // Fallback to cache/properties logic

      const template = HtmlService.createTemplateFromFile('form');
      template.detectedEmail = activeEmail; // Explicit flag for internal users
      template.detectedName = profileName;  // Pass detected name
      template.userEmail = activeEmail || user.email;
      template.userName = profileName || user.name;

      // Logger.log(`Showing form for: ${user.email}`);
      return template.evaluate()
        .setTitle("Leave Request Form")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    }

    const rowIndex = parseInt(e.parameter.row, 10);
    const action = e.parameter.action;
    // Note for rejection: If a note is crucial for rejection,
    // the email link should prompt the user or they should be instructed to add it to the URL.
    // e.g., &note=YourReasonHere. Or an intermediary HTML page could be used for a richer experience.
    const note = e.parameter.note || ''; // Approvers can manually add ?note=text to the URL if needed.
    // const stageFromParam = e.parameter.stage; // The stage link was clicked from

    if (isNaN(rowIndex) || rowIndex < 2 || !['approve', 'reject'].includes(action)) {
      Logger.log(`[INFO]: ${activeUser} - Invalid parameters. Please ensure the link is correct.`);
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

      if (currentStatus !== 'Pending' && currentStatus !== '') {
        Logger.log(`Request at row ${rowIndex} has status '${currentStatus}' and is currently at stage '${currentStage}'`);

        // If status is Approved or Rejected, consider it final and stop further processing
        if (currentStatus === 'Approved' || currentStatus === 'Rejected') {
          const html = HtmlService.createTemplateFromFile('result');
          html.action = currentStatus.toLowerCase();
          html.stage = currentStage;

          // Try/catch in case note is blank or access fails
          try {
            html.note = sheet.getRange(rowIndex, COLUMNS.NOTE).getValue();
          } catch (error) {
            Logger.log(`Failed to fetch note for row ${rowIndex}: ${error}`);
            html.note = '';
          }

          html.nextStage = "Final";

          Logger.log(`Returning result page with status '${currentStatus}' for row ${rowIndex}`);
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
      // This is a critical check to ensure the token and stage are provided
      // Also this happen when the sheet 'share link' is changed to visitor. The Correct one is 'Editor'.
      if (!tokenToUse || !tokenColumn || !stageParam) {
        Logger.log(`[TOKEN CHECK] Unauthorized Access - Missing token/column/stage: tokenToUse=${tokenToUse}, tokenColumn=${tokenColumn}, stageParam=${stageParam}`);
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
      const hrEmail = (data[COLUMNS.HR_EMAIL - 1] || "").toLowerCase();
      const gmEmail = (data[COLUMNS.GM_EMAIL - 1] || "").toLowerCase();

      let expectedApprover = "";

      if (stageParam === "spv") expectedApprover = spvEmail;
      else if (stageParam === "hr") expectedApprover = hrEmail;
      else if (stageParam === "gm") expectedApprover = gmEmail;
      else return showErrorTokenPage("Invalid Stage", "Unknown stage: " + stageParam);

      // 🔒 SECURITY CHECK — User must match expected approver
      if (currentUserEmail && !emailMatchesList(currentUserEmail, expectedApprover)) {
        Logger.log(`[TOKEN CHECK] Unauthorized Approver - Expected: ${expectedApprover}, Got: ${currentUserEmail}`);
        return showErrorTokenPage(
          "Unauthorized Approver",
          `Whoopzz Whoopzz <code>${currentUserEmail}</code>, <br>You're not authorized to approve this request<br>for stage <b>${currentStage}</b>.`
        );
      }

      // Check saved token
      const savedToken = sheet.getRange(rowIndex, tokenColumn).getValue();
      Logger.log(`[TOKEN CHECK] Comparing tokens - savedToken: ${savedToken}, tokenToUse: ${tokenToUse}`);

      if (!savedToken || validStage !== tokenColumn || savedToken !== tokenToUse || savedToken.endsWith("_used")) {
        Logger.log(`[TOKEN CHECK] Token Mismatch or Already Used - validStage: ${validStage}, tokenColumn: ${tokenColumn}, token: ${savedToken}`);
        return showErrorTokenPage(
          "You've Already Responded",
          `Looks like you've already taken action on this request.<br>Current stage: <strong>${currentStage}</strong>.`
        );
      }


      Logger.log(`[TOKEN CHECK] SUCCESS — Approver validated: ${expectedApprover}`);

      // --------------------- End Here ---------------------

      // Block re-approval if already finalized
      if (["Approved", "Rejected"].includes(currentStatus)) {
        Logger.log(`[FINALIZED CHECK] Already Finalized - Status: ${currentStatus}, Stage: ${currentStage}`);
        const html = HtmlService.createTemplateFromFile('result');
        html.action = currentStatus.toLowerCase();
        html.stage = currentStage;
        html.note = `This request was already processed as ${currentStatus}. ` + sheet.getRange(rowIndex, COLUMNS.NOTE).getValue();
        html.nextStage = "Final";
        return html.evaluate().setTitle("Request Processed");
      }

      // Invalidate the token after use
      Logger.log(`[TOKEN INVALIDATE] Marking token as used: ${tokenToUse}_used`);
      sheet.getRange(rowIndex, tokenColumn).setValue(`${tokenToUse}_used`);

      name = sheet.getRange(rowIndex, COLUMNS.NAME).getValue();
      department = sheet.getRange(rowIndex, COLUMNS.DEPARTMENT).getValue();
      leaveType = sheet.getRange(rowIndex, COLUMNS.LEAVE_TYPE).getValue();
      requester = sheet.getRange(rowIndex, COLUMNS.REQUESTER_EMAIL).getValue();
      startDate = sheet.getRange(rowIndex, COLUMNS.START_DATE).getDisplayValue();
      endDate = sheet.getRange(rowIndex, COLUMNS.END_DATE).getDisplayValue();
      let storedDuration = sheet.getRange(rowIndex, COLUMNS.DURATION).getValue();
      reasonText = sheet.getRange(rowIndex, COLUMNS.REASON).getValue();

      // Update decision details
      sheet.getRange(rowIndex, COLUMNS.DECISION).setValue(decision + " by " + currentStage); // Be more specific
      sheet.getRange(rowIndex, COLUMNS.NOTE).setValue(note);
      sheet.getRange(rowIndex, COLUMNS.DECISION_DATE).setValue(new Date());

      // Update specific stage column (Q, R, S)
      const stageDecisionMap = {
        "SPV Approval": COLUMNS.SPV_DECISION,
        "HR Review": COLUMNS.HR_DECISION,
        "GM Review": COLUMNS.GM_DECISION
      };
      const stageCol = stageDecisionMap[currentStage];
      if (stageCol) {
        // Record "Approved" or "Rejected: Reason"
        const stageStatus = decision + (note ? `: ${note}` : "");
        sheet.getRange(rowIndex, stageCol).setValue(stageStatus);
      }

      // === Start Balance Validation Logic ===
      // Fetch balance from EmployeeMaster using the helper function
      const requesterEmail = requester.toLowerCase();
      const employeeBalanceData = getEmployeeBalance(requesterEmail);

      let balance = 0;
      let sickBalance = 0;
      if (employeeBalanceData) {
        balance = employeeBalanceData.annual;
        sickBalance = employeeBalanceData.sick;
      }

      // const days = Math.ceil((new Date(endDate) - new Date(startDate)) / (1000 * 60 * 60 * 24)) + 1;
      let days = storedDuration;
      if (storedDuration !== 0.5 && storedDuration !== "0.5") {
        days = calculateLeaveDays(new Date(startDate), new Date(endDate));
      }
      // const leaveTypeLower = leaveType.toLowerCase();

      let balanceType = "Leave";
      if (leaveType.toLowerCase().includes("sick")) balanceType = "Sick Leave";
      if (leaveType.toLowerCase().includes("unpaid")) balanceType = "Unpaid";

      const applicableBalance = balanceType === "Sick Leave" ? sickBalance : balance;

      const leaveTypes = {
        annual: ["Annual Leave", "Bereavement Leave", "Career Leave", "Ceremony Leave", "Other"],
        sick: ["Sick Leave"]
      };

      const cekAnnual = leaveTypes.annual.includes(leaveType);
      const cekSick = leaveTypes.sick.includes(leaveType);
      const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ        

      // === Start Balance Validation Logic (Only on APPROVE action) ===
      if (action === 'approve' && (cekAnnual || cekSick) && days > applicableBalance) {
        const rejectionNote = performAutoReject(
          rowIndex,                          // 1
          applicableBalance,                // 2
          balanceType,                      // 3
          days,                             // 4
          name,                             // 5
          leaveType,                        // 6
          startDate,                        // 7
          endDate,                          // 8
          reasonText,                       // 9
          requester,                        // 10
          spvEmail === requester ? "Auto-Rejected" : "",  // 11
          hrEmail === requester ? "Auto-Rejected" : "",   // 12
          gmEmail === requester ? "Auto-Rejected" : "",    // 13
          refID   // 14
        );

        const html = HtmlService.createTemplateFromFile('result');
        html.action = 'reject';
        html.stage = currentStage;
        html.note = rejectionNote + "<br><br><strong>This request was automatically rejected due to insufficient balance.</strong>"; // include Auto-rejected:
        html.nextStage = "Final";

        Logger.log('Result Action:', html.action);
        Logger.log('Result Note:', html.note);

        return html.evaluate().setTitle("Auto-Rejected");
      }
      // === End Balance Validation Logic ===

      if (action === 'approve') {
        sheet.getRange(rowIndex, COLUMNS.STATUS).setValue('Pending'); // Keep pending until final approval

        switch (currentStage) {
          case "SPV Approval":

            if (leaveType === "Working From Home (WFH)") {
              Logger.log(`Auto-finalizing WFH request. Finalizing request as 'Approved by SPV'`);
              const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
              finalizeRequest(rowIndex, decision, note, name, requester, "Approved by SPV", refID);
            } else {
              nextStage = "HR Review";
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
              Logger.log(`Moving to next stage: ${nextStage}`);

              // Generate a new token for HR
              const hrToken = generateRandomToken();
              Logger.log(`Generated HR Token: ${hrToken}`);

              // Save token in the sheet in HR_TOKEN column
              sheet.getRange(rowIndex, COLUMNS.HR_TOKEN).setValue(hrToken);
              Logger.log(`Saved HR Token to sheet at row ${rowIndex}, column ${COLUMNS.HR_TOKEN}`);

              // Send approval email to HR
              const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
              sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.HR_EMAIL, nextStage, rowIndex, hrToken, refID, null, requester, days);
              Logger.log(`Sent approval email to HR: ${CONFIG.HR_EMAIL} for row ${rowIndex}`);
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
            //
            // I don't know maybe this function still have bias on same stage, but whis is work fine.
            // Maybe for you the next person who saw this code, you can fix this better. Thanks
            // 
            // With Great Respect,
            // Galbatorix

            // Updated to support comma-separated emails in SPV_MAP
            const allSpvEmails = Object.values(CONFIG.SPV_MAP);
            const isRequesterSPV = allSpvEmails.some(spvEmailEntry => emailMatchesList(requester, spvEmailEntry));
            const isRequesterHR = emailMatchesList(requester, CONFIG.HR_EMAIL);
            // Only "Unpaid Leave - FT" (Full-Time) needs GM approval
            // "Unpaid Leave - ITN" (Intern) ends at HR
            const leaveTypeTrimmed = (leaveType || "").trim();
            const isUnpaidFT = leaveTypeTrimmed === "Unpaid Leave - FT";
            const needsGM = isUnpaidFT || isRequesterSPV;

            if (leaveType === "Working From Home (WFH)") {
              if (isRequesterHR) {
                nextStage = "GM Review";
                sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);

                const gmToken = generateRandomToken();
                sheet.getRange(rowIndex, COLUMNS.GM_TOKEN).setValue(gmToken);

                const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
                sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.GM_EMAIL, nextStage, rowIndex, gmToken, refID, null, requester, days);
              } else {
                const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
                finalizeRequest(rowIndex, decision, note, name, requester, "Approved by HR", refID);
              }
            } else {
              nextStage = needsGM ? "GM Review" : "Final";
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);

              const gmToken = generateRandomToken();
              sheet.getRange(rowIndex, COLUMNS.GM_TOKEN).setValue(gmToken);

              if (needsGM) {
                const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
                sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.GM_EMAIL, nextStage, rowIndex, gmToken, refID, null, requester, days);
                Logger.log("GM approval email sent");
              } else {
                const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
                finalizeRequest(rowIndex, decision, note, name, requester, "Approved by HR", refID);
              }
            }
            break;

          case "GM Review":
            Logger.log("Stage: GM Review");
            Logger.log("Leave Type: " + leaveType);
            Logger.log("Requester: " + requester);

            if (leaveType === "Working From Home (WFH)") {
              Logger.log("WFH submitted → GM → Reporting");
              const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
              finalizeRequest(rowIndex, decision, note, name, requester, "Approved by GM", refID);
            } else {
              Logger.log("Non-WFH → GM → Finalization");
              nextStage = "Final";
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
              const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
              finalizeRequest(rowIndex, decision, note, name, requester, "Approved by GM", refID);
            }
            break;

          default:
            Logger.log("Unexpected Stage: " + currentStage);
            nextStage = "Error in Workflow";
            sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
            const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
            finalizeRequest(rowIndex, "Error", "Workflow error at stage: " + stage, name, requester, "Workflow Error", refID);
            break;
        }
      } else { // Action is 'reject'
        Logger.log(`Rejecting request at row ${rowIndex} during stage: ${currentStage}`);

        sheet.getRange(rowIndex, COLUMNS.STATUS).setValue('Rejected');
        sheet.getRange(rowIndex, COLUMNS.STAGE).setValue('Rejected at ' + currentStage);

        const rejectionNote = note || `Rejected by ${currentStage}.`;
        Logger.log(`Rejection note: ${rejectionNote}`);

        const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

        const formattedStartDate = formatDate(startDate);
        const formattedEndDate = formatDate(endDate);

        // Calculate days using helper (handles string/date input)
        const leaveDays = calculateLeaveDays(startDate, endDate);

        Logger.log(`Leave dates: ${formattedStartDate} to ${formattedEndDate}, total ${leaveDays} day(s)`);

        const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ

        const template = HtmlService.createTemplateFromFile("finalNotification");
        template.name = name;
        template.leaveType = leaveType;
        template.startDate = formattedStartDate;
        template.endDate = formattedEndDate;
        template.totalDays = leaveDays;
        template.reason = reasonText;
        template.finalDecision = "Rejected";
        template.finalNote = rejectionNote;
        template.updatedBalance = null;
        template.refID = refID;
        template.row = rowIndex;

        template.spvStatus = values[COLUMNS.SPV_DECISION - 1] || "—";
        template.hrStatus = values[COLUMNS.HR_DECISION - 1] || "—";
        template.gmStatus = values[COLUMNS.GM_DECISION - 1] || "—";
        template.attachmentUrl = values[COLUMNS.ATTACHMENT_URL - 1] || null;

        const htmlBody = template.evaluate().getContent();

        try {
          queueEmail(
            requester,
            `ONEderland Leave/WFH Request Rejected: ${name}`,
            htmlBody
          );
        } catch (e) {
          Logger.log("Failed to send rejection email: " + e.toString());
        }

        const approverName = getProfileName() || activeUserEmail;

        sendGCPNotification(
          `<b>Request Rejected</b>\n\n` +
          `<b>Form ID:</b> ${refID}\n` +
          `<b>Name:</b> ${name}\n` +
          `<b>Email:</b> ${requester}\n` +
          `<b>Leave Type:</b> ${leaveType}\n` +
          `<b>Date:</b> ${formattedStartDate} - ${formattedEndDate} (${leaveDays} days)\n` +
          `<b>Decision:</b> Rejected by ${approverName}\n` +
          `<b>Notes:</b> ${rejectionNote}\n` +
          `<b>Doc:</b> ${template.attachmentUrl ? "Yes" : "No"}\n` +
          `<b>Reason:</b> ${reasonText}`
        );

        Logger.log(`Rejection email queued for: ${requester}`);

        nextStage = 'Final (Rejected)';
        Logger.log(`Stage set to: ${nextStage}`);
      }
    } finally {
      lock.releaseLock();
    }

    const html = HtmlService.createTemplateFromFile('result');
    html.action = action;
    html.stage = currentStage; // The stage that just made the decision
    html.note = note || (action === 'reject' ? "Rejected by " + currentStage : "Approved by " + currentStage);
    html.nextStage = nextStage;
    Logger.log(`[${action.toUpperCase()}] by ${currentStage} | Requester: ${requester} | Note: ${html.note} | Next: ${nextStage}`);
    return html.evaluate()
      .setTitle("Request Processed")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    const stackLine = (err.stack || '').split('\n')[1] || 'Line unknown';
    Logger.log("❌ Error in doGet:\n" + err.toString() + "\n📍 Location: " + stackLine);

    const template = HtmlService.createTemplateFromFile('processingError');
    template.errorMessage = escapeHtml(err.toString());
    template.stackLine = escapeHtml(stackLine);

    Logger.log(`Oops! Something went wrong. ${escapeHtml(err.toString())} - ${escapeHtml(stackLine)}`);
    return template.evaluate().setTitle("Processing Error");
  }
}

function submitRequest(name, email, department, leaveTypeOrList, startDate, endDate, reason, fileData) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("Requests");
    const spvEmail = CONFIG.SPV_MAP[department] || CONFIG.GM_EMAIL;
    const hrEmail = CONFIG.HR_EMAIL;
    const gmEmail = CONFIG.GM_EMAIL;

    Logger.log(`[DEBUG] Config Loaded. Test Mode: ${CONFIG.TEST_MODE}`);

    // --- Normalize Email Logic ---
    const data = sheet.getDataRange().getValues();
    let normalizedEmail = normalizeGmailAddress(email);
    let rewrittenEmail = email.trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const sheetEmailRaw = data[i][COLUMNS.EMP_EMAIL - 1];
      if (!sheetEmailRaw) continue;
      const sheetEmail = sheetEmailRaw.toString().trim().toLowerCase();
      if (normalizeGmailAddress(sheetEmail) === normalizedEmail) {
        rewrittenEmail = sheetEmail;
        break;
      }
    }
    email = rewrittenEmail;
    normalizedEmail = normalizeGmailAddress(email);
    // -----------------------------

    // Determine Requests List
    let requests = [];
    if (Array.isArray(leaveTypeOrList)) {
      requests = leaveTypeOrList; // Expecting [{leaveType, startDate, endDate, reason, duration}, ...]
    } else {
      requests = [{
        leaveType: leaveTypeOrList,
        startDate: startDate,
        endDate: endDate,
        reason: reason
      }];
    }

    Logger.log(`Processing ${requests.length} request(s) for ${email}`);

    // Generate a common Batch Ref ID for this submission group
    const batchRefID = generateReferenceID();

    // Process each request
    for (let i = 0; i < requests.length; i++) {
      const req = requests[i];
      processSingleLeaveRequest(
        sheet,
        name, email, department,
        req.leaveType, req.startDate, req.endDate, req.reason, req.duration,
        spvEmail, hrEmail, gmEmail,
        req.fileData || fileData, // Use per-request file, fallback to global
        batchRefID // Pass shared ID
      );
    }

    return "Success";

  } catch (e) {
    Logger.log("Submit error: " + e.toString() + " Stack: " + e.stack);
    return e.message || "Submission failed due to a server error. Please try again.";
  }
}

function processSingleLeaveRequest(sheet, name, email, department, leaveType, startDate, endDate, reason, duration, spvEmail, hrEmail, gmEmail, fileData, batchRefID) {
  const firstDate = parseDMYDate(startDate);
  const lastDate = parseDMYDate(endDate);
  const isIntern = email.toLowerCase().includes('+itn@');

  // Validation
  if (isIntern && leaveType !== "Unpaid Leave - ITN") {
    throw new Error("Access Denied: Your authorization is only can take 'Unpaid Leave - ITN'.");
  }
  if (!isIntern && leaveType === "Unpaid Leave - ITN") {
    throw new Error("Access Denied: You're not allowed to take 'Unpaid Leave - ITN'.");
  }
  if (!firstDate || !lastDate) {
    throw new Error(`Invalid date format for ${leaveType}. Please use DD-MM-YYYY.`);
  }
  if (lastDate < firstDate) {
    throw new Error(`❌ SYSTEM ERROR: USER IQ NOT FOUND ❌<br>The last day of ${leaveType} cannot be BEFORE the first day.<br>If you can time-travel, please go back and fix the economy first!<br>Until then… PLEASE enter a valid date! 😭`);
  }

  // Role Logic - Updated to support comma-separated emails in Settings
  const isRequesterSPV = emailMatchesList(email, spvEmail);
  const isRequesterGM = emailMatchesList(email, gmEmail);
  const isRequesterHR = emailMatchesList(email, hrEmail);

  let stage = "SPV Approval"; // Default
  let approvalEmail = spvEmail;

  // Logic Matrix
  if (isRequesterSPV) {
    if (leaveType === "Working From Home (WFH)") {
      stage = "GM Review";
      approvalEmail = gmEmail;
    } else if (isRequesterHR) {
      // HR as SPV (Dyah Retno case?) -> GM
      stage = "GM Review";
      approvalEmail = gmEmail;
    } else {
      // SPV Normal -> HR -> GM
      stage = "HR Review";
      approvalEmail = hrEmail;
    }
  } else if (isRequesterGM) {
    stage = "HR Review";
    approvalEmail = hrEmail;
  } else {
    // Normal Employee
    stage = "SPV Approval";
    approvalEmail = spvEmail;
  }

  Logger.log(`[Item] ${leaveType}: Next stage: ${stage}, Approver: ${approvalEmail}`);

  // Tokens
  const spvToken = generateRandomToken();
  const hrToken = generateRandomToken();
  const gmToken = generateRandomToken();

  let tokenToUse = (stage === "SPV Approval") ? spvToken : (stage === "HR Review" ? hrToken : gmToken);

  // Generate RefID
  let refID = batchRefID || generateReferenceID();

  // File Queue Logic
  // Only queue file if this specific request is Sick Leave (and file was uploaded)
  let hasFileToQueue = (leaveType === "Sick Leave") && fileData && fileData.content;
  let attachmentUrl = hasFileToQueue ? "Processing..." : "";

  // Calculate Duration
  let days = 0;
  if (duration === 0.5 || duration === "0.5") {
    days = 0.5;
  } else {
    days = calculateLeaveDays(firstDate, lastDate);
  }

  // Append Row
  sheet.appendRow([
    new Date(), name, department, leaveType,
    firstDate, lastDate, reason,
    "Pending", email, spvEmail, hrEmail, gmEmail, stage,
    "", "", new Date(), // Decision, Note, Decision Date
    "", "", "",         // SPV..GM Decisions
    spvToken, hrToken, gmToken,
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", // fill until col AQ
    "" // placeholder for RefID
  ]);

  const lastRow = sheet.getLastRow();
  refID = generateReferenceID(lastRow, batchRefID || refID);

  // Update RefID & Attachment Placeholder & Duration
  sheet.getRange(lastRow, COLUMNS.REF_ID).setValue(refID);
  sheet.getRange(lastRow, COLUMNS.DURATION).setValue(days);
  if (attachmentUrl) sheet.getRange(lastRow, COLUMNS.ATTACHMENT_URL).setValue(attachmentUrl);

  // Queue File
  if (hasFileToQueue) {
    try {
      queueFileUpload(fileData.name, fileData.mimeType, fileData.content, lastRow);
    } catch (err) {
      sheet.getRange(lastRow, COLUMNS.ATTACHMENT_URL).setValue("Error queuing file");
    }
  }

  // Notifications
  // We pass 'days' to notification to ensure accuracy? 
  // sendSubmissionConfirmation currently doesn't take 'days' as param, it might calculate it or ignore it.
  // Let's check sendSubmissionConfirmation signature. It takes (email, name, leaveType, start, end, reason...)
  // It probably recalculates strings. 
  sendSubmissionConfirmation(email, name, leaveType, firstDate, lastDate, reason, stage, refID, attachmentUrl, lastRow);
  sendApprovalEmail(name, leaveType, firstDate, lastDate, reason, approvalEmail, stage, lastRow, tokenToUse, refID, attachmentUrl, email, days);
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

function finalizeRequest(row, decision, note, name, requesterEmail, finalApprovalStageNote, refID) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Requests");
  const finalNote = note || finalApprovalStageNote || "No Notes";

  const leaveType = sheet.getRange(row, COLUMNS.LEAVE_TYPE).getValue();
  const startDate = sheet.getRange(row, COLUMNS.START_DATE).getDisplayValue();
  const endDate = sheet.getRange(row, COLUMNS.END_DATE).getDisplayValue();
  const reason = sheet.getRange(row, COLUMNS.REASON).getValue();
  const spvStatus = sheet.getRange(row, COLUMNS.SPV_DECISION).getValue();
  const hrStatus = sheet.getRange(row, COLUMNS.HR_DECISION).getValue() || '';
  const gmStatus = sheet.getRange(row, COLUMNS.GM_DECISION).getValue() || '';
  const department = sheet.getRange(row, COLUMNS.DEPARTMENT).getValue();
  const attachmentUrl = sheet.getRange(row, COLUMNS.ATTACHMENT_URL).getValue();

  // Read duration from sheet (support half-day). Fallback to calc if empty for legacy rows.
  // Read duration from sheet. 
  // Fix for "88 days" bug: Recalculate duration unless it is explicitly 0.5 (Half-Day).
  // This ensures we correct any bad data stored in the sheet from previous buggy submissions.
  let storedDuration = sheet.getRange(row, COLUMNS.DURATION).getValue();
  let days = storedDuration;

  if (storedDuration !== 0.5 && storedDuration !== "0.5") {
    days = calculateLeaveDays(startDate, endDate);
    // Optional: Update the sheet with the corrected value?
    // sheet.getRange(row, COLUMNS.DURATION).setValue(days);
  }

  // Fallback if calculateLeaveDays returns 0 but shouldn't? (Safety)
  if (!days && days !== 0) days = 0;

  const leaveTypes = {
    annual: ["Annual Leave", "Career Leave", "Ceremony Leave", "Other"],
    sick: ["Sick Leave"],
    bereavement: ["Bereavement Leave"],
    marriage: ["Marriage Leave"],
    maternity: ["Maternity Leave"]
  };

  // Fetch balance from EmployeeMaster using helper function
  const requester = requesterEmail.toLowerCase();
  const employeeBalance = getEmployeeBalance(requester);
  let updatedBalance = null;

  if (employeeBalance) {
    if (decision === "Approved") {
      const isAnnual = leaveTypes.annual.includes(leaveType);
      const isSick = leaveTypes.sick.includes(leaveType);
      const isBereavement = leaveTypes.bereavement.includes(leaveType);
      const isMarriage = leaveTypes.marriage.includes(leaveType);
      const isMaternity = leaveTypes.maternity.includes(leaveType);

      // Determine which balance to check and update
      let balanceCategory = null;
      let currentBalance = 0;
      let masterColumn = null;

      if (isAnnual) {
        balanceCategory = "Annual";
        currentBalance = employeeBalance.annual;
        masterColumn = MASTER_COLUMNS.ANNUAL_BALANCE;
      } else if (isSick) {
        balanceCategory = "Sick";
        currentBalance = employeeBalance.sick;
        masterColumn = MASTER_COLUMNS.SICK_BALANCE;
      } else if (isBereavement) {
        balanceCategory = "Bereavement";
        currentBalance = employeeBalance.bereavement;
        masterColumn = MASTER_COLUMNS.BEREA_BALANCE;
      } else if (isMarriage) {
        balanceCategory = "Marriage";
        currentBalance = employeeBalance.marriage;
        masterColumn = MASTER_COLUMNS.MARRIAGE_BALANCE;
      } else if (isMaternity) {
        balanceCategory = "Maternity";
        currentBalance = employeeBalance.maternity;
        masterColumn = MASTER_COLUMNS.MATERNITY_BALANCE;
      }

      // If it matches a tracked leave type, check and update balance
      if (balanceCategory) {
        // Check balance sufficiency
        if (days > currentBalance) {
          performAutoReject(row, currentBalance, balanceCategory, days, name, leaveType, startDate, endDate, reason, requesterEmail, spvStatus, hrStatus, gmStatus, refID);
          return;
        }

        // Update balance in EmployeeMaster
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const masterSheet = ss.getSheetByName("EmployeeMaster");

        if (masterSheet) {
          const masterData = masterSheet.getDataRange().getValues();
          // Find the employee row in EmployeeMaster (skip 2 header rows)
          for (let i = 2; i < masterData.length; i++) {
            const rowEmail = normalizeGmailAddress(masterData[i][MASTER_COLUMNS.EMAIL - 1]);
            if (rowEmail === normalizeGmailAddress(requester)) {
              const masterRow = i + 1; // Convert to 1-based row index
              const newBalance = Math.max(0, currentBalance - days);
              masterSheet.getRange(masterRow, masterColumn).setValue(newBalance);

              // Prepare balance display: "Old -> New"
              updatedBalance = {};
              updatedBalance[balanceCategory.toLowerCase()] = `${currentBalance} -> ${newBalance}`;


              Logger.log(`✅ Updated ${balanceCategory} Balance for ${requester}: ${currentBalance} -> ${newBalance}`);
              break;
            }
          }
        } else {
          Logger.log("❌ EmployeeMaster sheet not found for balance update!");
        }
      }
    }
  } else {
    Logger.log(`⚠️ No balance record found for ${requester} in EmployeeMaster`);
  }

  // Update final approval status
  sheet.getRange(row, COLUMNS.STATUS).setValue(decision);
  sheet.getRange(row, COLUMNS.STAGE).setValue("Completed");
  sheet.getRange(row, COLUMNS.NOTE).setValue(finalNote);

  // === QUEUE CALENDAR EVENT (for Trigger processing) ===
  // Instead of creating calendar events directly (which fails in "User Accessing" mode),
  // we set a status to "Pending" for the Trigger function to pick up.
  if (decision === "Approved") {
    sheet.getRange(row, COLUMNS.CALENDAR_STATUS).setValue("Pending");
    Logger.log(`🗓️ Queued calendar event for row ${row}`);
  }
  // === END ===

  // Prepare notification data
  const formattedStart = formatDate(startDate);
  const formattedEnd = formatDate(endDate);

  // Final Notification Email to Requester
  const template = HtmlService.createTemplateFromFile('finalNotification');
  Object.assign(template, {
    name, leaveType, reason,
    startDate: formattedStart,
    endDate: formattedEnd,
    totalDays: days,
    spvStatus, hrStatus, gmStatus,
    finalDecision: decision,
    finalNote,
    updatedBalance,
    refID,

    row,
    attachmentUrl: attachmentUrl || null,
  });

  const htmlBody = template.evaluate().getContent();
  queueEmail(requesterEmail, `ONEderland Leave Request ${decision}: ${name}`, htmlBody);

  sendGCPNotification(
    `<b>Request ${decision}</b>\n\n` +
    `<b>Form ID:</b> ${refID}\n` +
    `<b>Name:</b> ${name}\n` +
    `<b>Email:</b> ${requesterEmail}\n` +
    `<b>Leave Type:</b> ${leaveType}\n` +
    `<b>Date:</b> ${formattedStart} - ${formattedEnd} (${days} days)\n` +
    `<b>Decision:</b> ${decision}\n` +
    `<b>Doc:</b> ${attachmentUrl ? "Yes" : "No"}\n` +
    `<b>Reason:</b> ${reason}`
  );

  // Reporting Team Notification with Calendar Link
  // const calendarTitle = `${name}'s ${leaveType}`;
  // const calendarDescription = `Leave Request\nRequester: ${name}\nDepartment: ${department}\nType: ${leaveType}\nDecision: ${decision}\nNote: ${finalNote}`;
  // const calendarLink = generateCalendarLink(calendarTitle, startDate, endDate, calendarDescription);

  const reportingTemplate = HtmlService.createTemplateFromFile('reportingEmail');
  reportingTemplate.decision = decision;
  reportingTemplate.refID = refID;
  reportingTemplate.name = name;
  reportingTemplate.department = department;
  reportingTemplate.leaveType = leaveType;
  reportingTemplate.formattedStart = formattedStart;
  reportingTemplate.formattedEnd = formattedEnd;
  reportingTemplate.days = days;
  reportingTemplate.reason = reason;
  reportingTemplate.finalNote = finalNote;
  reportingTemplate.attachmentUrl = attachmentUrl;
  reportingTemplate.updatedBalance = updatedBalance;

  const reportingHtml = reportingTemplate.evaluate().getContent();

  CONFIG.REPORTING_EMAILS.forEach(email => {
    queueEmail(email, `Reporting Notification from ${name}`, reportingHtml);
  });
}

function generateCalendarLink(title, startDate, endDate, description) {
  const formatDate = date => Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyyMMdd");

  const start = formatDate(startDate);
  const end = formatDate(new Date(new Date(endDate).getTime() + 86400000)); // +1 day to make it inclusive

  const details = encodeURIComponent(description);
  const calTitle = encodeURIComponent(title);

  Logger.log(`Calender Created to: ${title} for ${startDate} - ${endDate}`);
  return `https://www.google.com/calendar/render?action=TEMPLATE&text=${calTitle}&dates=${start}/${end}&details=${details}`;
}

function sendApprovalEmail(name, leaveType, startDate, endDate, reason, approverEmail, stage, row, tokenToUse, refID, attachmentUrl, requesterEmail, duration) {
  const baseUrl = ScriptApp.getService().getUrl();

  const template = HtmlService.createTemplateFromFile("emailtemplate");
  template.name = name;
  template.leaveType = leaveType;
  template.startDate = formatDate(startDate);
  template.endDate = formatDate(endDate);

  // Use passed duration if available, otherwise recalculate (legacy fallback)
  const totalDays = duration || calculateLeaveDays(startDate, endDate);
  template.totalDays = totalDays;

  template.reason = reason;
  template.stage = stage;
  template.refID = refID;
  template.baseUrl = baseUrl;
  template.attachmentUrl = attachmentUrl || null; // Pass to template

  // Extract normalized stage (lowercase) → Used for URL param
  const shortStage = stage.toLowerCase().includes("spv") ? "spv"
    : stage.toLowerCase().includes("hr") ? "hr"
      : stage.toLowerCase().includes("gm") ? "gm"
        : "unknown"; // fallback

  if (!tokenToUse || shortStage === "unknown") {
    Logger.log(`❗ Missing token or invalid stage in sendApprovalEmail for stage: ${stage}, row: ${row}`);
    return;
  }

  // Safe & encoded approval/rejection URLs
  const encodedToken = encodeURIComponent(tokenToUse);
  const noteText = `Approved at ${stage}`;
  template.approveUrl = `${baseUrl}?action=approve&stage=${shortStage}&row=${row}&token=${encodedToken}&note=${encodeURIComponent(noteText)}`;
  // Direct Reject (No intermediate notes page by default, as requested)
  template.rejectUrl = `${baseUrl}?action=reject&stage=${shortStage}&row=${row}&token=${encodedToken}`;

  try {
    queueEmail(
      approverEmail,
      `[Action Required] Leave/WFH Request: ${name} (${stage})`,
      template.evaluate().getContent(),
      row // Pass row for attachment URL refresh
    );
    Logger.log(`✅ Approval email queued for ${approverEmail} for ${name} - ${refID}`);
  } catch (e) {
    Logger.log(`❌ Failed to queue approval email for ${approverEmail} for ${name} - ${refID}. Error: ${e}`);
  }

  // Intermediate Approval Notification (Next Stage)
  const durationText = (totalDays == 0.5 || totalDays == "0.5") ? `${totalDays} days (Half-Day)` : `${totalDays} days`;

  sendGCPNotification(
    `<b>Request Progress Update</b>\n\n` +
    `<b>Form ID:</b> ${refID}\n` +
    `<b>Name:</b> ${name}\n` +
    `<b>Email:</b> ${requesterEmail || "(Not Available)"}\n` +
    `<b>Leave Type:</b> ${leaveType}\n` +
    `<b>Date:</b> ${formatDate(startDate)} - ${formatDate(endDate)} (${durationText})\n` +
    `<b>Decision:</b> Moved to ${stage}\n` +
    `<b>Doc:</b> ${attachmentUrl ? "Yes" : "No"}\n` +
    `<b>Reason:</b> ${reason}`
  );
}

function sendSubmissionConfirmation(email, name, leaveType, startDate, endDate, reason, stage, refID, attachmentUrl, sourceRow) {
  const scriptUrl = ScriptApp.getService().getUrl();
  const trackingLink = `${scriptUrl}?track=${encodeURIComponent(email)}`;

  let attachmentHtml = "";
  if (attachmentUrl) {
    attachmentHtml = `<tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;"><strong>Attachment</strong></td><td style="padding:8px;border:1px solid #ddd;"><a href="${attachmentUrl}" target="_blank">View Document</a></td></tr>`;
  }

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:600px; margin:auto; border:1px solid #ddd; padding:20px;">
      <h2 style="color:#2c3e50; border-bottom:1px solid #eee; padding-bottom:10px;">Leave Request Submitted</h2>
      <p>Dear ${name},</p>
      <p>Your leave request has been successfully submitted and is now awaiting <strong>${stage}</strong>.</p>
      
      <table style="width:100%;border-collapse:collapse;margin:20px 0;">
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;width:30%;"><strong>Form ID</strong></td><td style="padding:8px;border:1px solid #ddd;">${refID}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;width:30%;"><strong>Leave Type</strong></td><td style="padding:8px;border:1px solid #ddd;">${leaveType}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;"><strong>Dates</strong></td><td style="padding:8px;border:1px solid #ddd;">${formatDate(startDate)} to ${formatDate(endDate)}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;"><strong>Reason</strong></td><td style="padding:8px;border:1px solid #ddd;">${escapeHtml(reason)}</td></tr>
        ${attachmentHtml}
      </table>
      
      <p>You can track your leave request status anytime using the link below:</p>
      <p>
        <a href="${trackingLink}" target="_blank" style="display:inline-block;padding:10px 15px;background-color:#3498db;color:#fff;text-decoration:none;border-radius:5px;">
          Track My Leave Request
        </a>
      </p>

      <p>You will receive further email notifications as your request is processed.</p>
      <p style="color:#7f8c8d;font-size:0.9em;margin-top:20px;">Thank you,<br/>ONEderland Approval System</p>
    </div>
  `;

  try {
    queueEmail(
      email,
      "ONEderland Leave/WFH Request Submission Confirmation",
      htmlBody,
      sourceRow
    );
  } catch (e) {
    Logger.log(`Failed to queue confirmation email to ${email}. Error: ${e.toString()}`);
  }
}

// Update Balance and detect balance on colums.email

// Different beetwen parseInt and parseFloat
// | Input    | `parseInt()` | `parseFloat()` | Recommended                        |
// | -------- | ------------ | -------------- | ---------------------------------- |
// | `"1.5"`  | `1`          | `1.5`          | `parseFloat` if you need decimal   |
// | `"18,5"` | `18`         | `NaN`          | `parseFloat(...replace(',', '.'))` |
// | `"100"`  | `100`        | `100`          | Either works                       |

function getLeaveBalanceByEmail(email) {
  if (!email || typeof email !== 'string') {
    Logger.log("[ERROR] Invalid email input: %s", email);
    throw new Error("Invalid email passed to getLeaveBalanceByEmail");
  }

  const inputEmail = normalizeGmailAddress(email);

  // Use the new helper to get balance from EmployeeMaster
  const balance = getEmployeeBalance(inputEmail);

  if (balance) {
    Logger.log("[INFO] Found balance for %s: Leave = %s, Sick = %s, Bereavement = %s, Marriage = %s, Maternity = %s", inputEmail, balance.annual, balance.sick, balance.bereavement, balance.marriage, balance.maternity);
    return {
      annual: balance.annual,
      sick: balance.sick,
      bereavement: balance.bereavement,
      marriage: balance.marriage,
      maternity: balance.maternity
    };
  }

  Logger.log("[WARN] No balance found for email: %s", inputEmail);
  return null;
}

/**
 * Helper to fetch employee balance from EmployeeMaster sheet.
 * @param {string} email - The employee email to look up.
 * @returns {object|null} - { annual: number, sick: number } or null if not found.
 */
function getEmployeeBalance(email) {
  if (!email) return null;
  const normalizedInput = normalizeGmailAddress(email);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("EmployeeMaster");

  if (!masterSheet) {
    Logger.log("❌ Error: EmployeeMaster sheet not found!");
    return null;
  }

  const data = masterSheet.getDataRange().getValues();

  // Helper to parse balance (preserve negative, default NaN to 0)
  const parseBalance = (val) => {
    const num = parseFloat(val);
    return isNaN(num) ? 0 : num;
  };

  // Skip 2 header rows (starting loop from 2)
  for (let i = 2; i < data.length; i++) {
    const rowEmail = normalizeGmailAddress(data[i][MASTER_COLUMNS.EMAIL - 1]);
    if (rowEmail === normalizedInput) {
      return {
        annual: parseBalance(data[i][MASTER_COLUMNS.ANNUAL_BALANCE - 1]),
        sick: parseBalance(data[i][MASTER_COLUMNS.SICK_BALANCE - 1]),
        bereavement: parseBalance(data[i][MASTER_COLUMNS.BEREA_BALANCE - 1]),
        marriage: parseBalance(data[i][MASTER_COLUMNS.MARRIAGE_BALANCE - 1]),
        maternity: parseBalance(data[i][MASTER_COLUMNS.MATERNITY_BALANCE - 1])
      };
    }
  }
  return null;
}

function calculateLeaveDays(startDate, endDate) {
  // Helper to ensure we have a valid Date object, preferring DMY for strings
  const toDate = (d) => {
    if (d instanceof Date) return new Date(d.getTime());
    if (typeof d === 'string') {
      const dmy = parseDMYDate(d); // Try explicit d-m-y
      if (dmy && !isNaN(dmy.getTime())) return dmy;
      // Fallback
      return new Date(d);
    }
    return new Date(d);
  };

  let start = toDate(startDate);
  let end = toDate(endDate);

  // Safety check
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return 0;

  let count = 0;

  while (start <= end) {
    const day = start.getDay();
    if (day !== 0 && day !== 6) { // Mon-Fri only
      count++;
    }
    start.setDate(start.getDate() + 1);
  }

  return count;
}

function performAutoReject(rowIndex, balance, balanceType, days, name, leaveType, startDate, endDate, reason, requester, spvStatus, hrStatus, gmStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const rejectionNote = `Auto-rejected, Only ${balance} ${balanceType.toLowerCase()} day(s) left, but ${days} requested.`;
  const systemNote = `Your request was rejected because you only have ${balance} ${balanceType.toLowerCase()} day(s) left, but you requested ${days} day(s).`;

  // Update sheet
  sheet.getRange(rowIndex, COLUMNS.STATUS).setValue("Rejected");
  sheet.getRange(rowIndex, COLUMNS.STAGE).setValue("Completed");
  sheet.getRange(rowIndex, COLUMNS.NOTE).setValue(rejectionNote);

  // Fetch current leave balances from EmployeeMaster
  let currentLeave = 0, currentSick = 0;
  const employeeBalance = getEmployeeBalance(requester);
  if (employeeBalance) {
    currentLeave = employeeBalance.annual;
    currentSick = employeeBalance.sick;
  }

  // Fetch RefID from sheet
  const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue();
  const attachmentUrl = sheet.getRange(rowIndex, COLUMNS.ATTACHMENT_URL).getValue();

  // Final rejection email
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
  template.finalDecision = "Rejected";
  template.finalNote = systemNote; // Use updated extra note
  template.refID = refID;   // properly assigned
  template.updatedBalance = {
    leave: currentLeave,
    leave: currentLeave,
    sick: currentSick
  };
  template.attachmentUrl = attachmentUrl || null;

  const htmlBody = template.evaluate().getContent();
  queueEmail(requester, `ONEderland Leave Request Rejected: ${name}`, htmlBody);

  Logger.log('🚨 performAutoReject triggered. Returning note: ' + rejectionNote);

  return rejectionNote;
}

function formatDateShort(date, withTime = false) {
  if (!(date instanceof Date)) date = new Date(date);
  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  if (withTime) options.hour = '2-digit', options.minute = '2-digit';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), withTime ? "MMMM d, yyyy HH:mm" : "MMMM d, yyyy");
}

function renderTrackingPage(email, showHistory = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Requests");
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // Skip header

  const filteredRows = rows.filter(row => {
    const emailMatch = row[COLUMNS.REQUESTER_EMAIL - 1] === email;
    const isPending = row[COLUMNS.STATUS - 1].toLowerCase().includes("pending");
    const hasToken = !!row[COLUMNS.SPV_TOKEN - 1];
    return emailMatch && (showHistory ? true : (isPending && hasToken));
  });

  if (filteredRows.length === 0) {
    return `
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      <div class="container mt-5">
        <div class="alert alert-warning">
          <h4>No ${showHistory ? "leave history" : "active pending"} requests found for <code>${email}</code></h4>
          <p>Either they have been approved/rejected or the email is incorrect.</p>
        </div>
      </div>
    `;
  }

  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body { font-family: 'Inter', sans-serif; background-color: #f8f9fa; }
        .table-custom { background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        .table-custom th { background-color: #f1f3f5; font-weight: 600; text-transform: uppercase; font-size: 0.85rem; letter-spacing: 0.5px; }
        .badge-status { font-size: 0.8rem; padding: 0.4em 0.8em; }
      </style>
    </head>
    <body class="p-4">
    <div class="container">
      <div class="d-flex justify-content-between align-items-center mb-4">
        <h2>${showHistory ? "Leave Request History" : "My Active Requests"}</h2>
        <a href="${ScriptApp.getService().getUrl()}" class="btn btn-outline-primary">← New Request</a>
      </div>

      <div class="card border-0 shadow-sm p-3 mb-4">
        <div class="d-flex align-items-center">
          <div class="flex-grow-1">
             <strong>Employee:</strong> ${escapeHtml(email)}
          </div>
          <div>
            <a href="${ScriptApp.getService().getUrl()}?track=${email}&history=${!showHistory}" class="btn btn-sm btn-link text-decoration-none">
              ${showHistory ? "Hide History (Show Active Only)" : "Show All History"}
            </a>
          </div>
        </div>
      </div>

      <div class="table-responsive table-custom">
        <table class="table table-hover mb-0 align-middle">
          <thead>
            <tr>
              <th>Date</th>
              <th>Name</th>
              <th>Leave Type</th>
              <th>Date Range</th>
              <th>Status</th>
              <th>Current Stage</th>
              <th>Form ID</th>
              <th class="text-end">Action</th>
            </tr>
          </thead>
          <tbody>
  `;

  const scriptUrl = ScriptApp.getService().getUrl();

  filteredRows.forEach(row => {
    const isPending = row[COLUMNS.STATUS - 1].toLowerCase().includes("pending");
    const refID = row[COLUMNS.REF_ID - 1];
    const status = row[COLUMNS.STATUS - 1];

    let badgeClass = "bg-secondary";
    if (status === "Approved") badgeClass = "bg-success";
    if (status === "Rejected" || status === "Cancelled") badgeClass = "bg-danger";
    if (status === "Pending") badgeClass = "bg-warning text-dark";

    let actionBtn = "";
    if (isPending) {
      actionBtn = `<button class="btn btn-sm btn-outline-danger" onclick="showCancelModal('${escapeHtml(refID)}')">Cancel</button>`;
    }

    html += `
      <tr>
        <td>${formatDateShort(row[COLUMNS.TIMESTAMP - 1], true)}</td>
        <td>${escapeHtml(row[COLUMNS.NAME - 1])}</td>
        <td>${escapeHtml(row[COLUMNS.LEAVE_TYPE - 1])}</td>
        <td>
          ${formatDateShort(row[COLUMNS.START_DATE - 1])} - ${formatDateShort(row[COLUMNS.END_DATE - 1])}<br>
          <small class="text-muted">${calculateLeaveDays(row[COLUMNS.START_DATE - 1], row[COLUMNS.END_DATE - 1])} days</small>
        </td>
        <td><span class="badge ${badgeClass} badge-status">${escapeHtml(row[COLUMNS.STATUS - 1])}</span></td>
        <td>${escapeHtml(row[COLUMNS.STAGE - 1])}</td>
        <td><code class="text-muted">${escapeHtml(refID)}</code></td>
        <td class="text-end">${actionBtn}</td>
      </tr>
    `;
  });

  html += `
          </tbody>
        </table>
      </div>
    </div>

    <!-- Cancel Confirmation Modal -->
    <div class="modal fade" id="cancelModal" tabindex="-1" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered">
        <div class="modal-content">
          <div class="modal-header border-0 pb-0">
            <h5 class="modal-title">Cancel Request?</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
          </div>
          <div class="modal-body text-center py-4">
            <div class="mb-3 text-warning" style="font-size: 3rem;">⚠️</div>
            <p>Are you sure you want to cancel this request?<br><strong>This action cannot be undone.</strong></p>
          </div>
          <div class="modal-footer border-0 justify-content-center pb-4">
            <button type="button" class="btn btn-light px-4" data-bs-dismiss="modal">No, Keep It</button>
            <a href="#" id="confirmCancelBtn" class="btn btn-danger px-4">Yes, Cancel Request</a>
          </div>
        </div>
      </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
      function showCancelModal(refID) {
        const url = "${scriptUrl}?action=cancel&refID=" + encodeURIComponent(refID);
        document.getElementById('confirmCancelBtn').href = url;
        const myModal = new bootstrap.Modal(document.getElementById('cancelModal'));
        myModal.show();
      }
    </script>
    </body>
    </html>
  `;

  return html;
}

/**
 * Cancels a leave request by RefID.
 * @param {string} refID The reference ID of the request to cancel.
 * @returns {object} Result object {success: boolean, message: string}
 */
function cancelRequest(refID) {
  if (!refID) return { success: false, message: "Missing Reference ID" };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Requests");
  const data = sheet.getDataRange().getValues();

  // Find the row with matching RefID
  let rowIndex = -1;
  let status = "";

  for (let i = 1; i < data.length; i++) { // Skip header
    if (String(data[i][COLUMNS.REF_ID - 1]) === String(refID)) {
      rowIndex = i + 1; // 1-based row index
      status = data[i][COLUMNS.STATUS - 1];
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: "Request not found." };
  }

  if (status !== "Pending") {
    return { success: false, message: `Cannot cancel request. Current status: ${status}` };
  }

  // Update status to Cancelled
  sheet.getRange(rowIndex, COLUMNS.STATUS).setValue("Cancelled");
  sheet.getRange(rowIndex, COLUMNS.STAGE).setValue("Cancelled by User");

  Logger.log(`Request ${refID} cancelled by user.`);
  return { success: true, message: "Request cancelled successfully." };
}

function getPendingApprovals(currentUserEmail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Requests");
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  const user = currentUserEmail.toLowerCase();

  const results = [];
  const scriptUrl = ScriptApp.getService().getUrl();

  rows.forEach((row, index) => {
    const status = row[COLUMNS.STATUS - 1];
    if (status !== "Pending") return;

    const stage = row[COLUMNS.STAGE - 1];
    const spvEmail = (row[COLUMNS.SPV_EMAIL - 1] || "").toLowerCase();
    const hrEmail = (row[COLUMNS.HR_EMAIL - 1] || "").toLowerCase();
    const gmEmail = (row[COLUMNS.GM_EMAIL - 1] || "").toLowerCase();

    let isMyTurn = false;
    let token = "";
    let shortStage = "";

    if (stage === "SPV Approval" && spvEmail === user) {
      isMyTurn = true;
      token = row[COLUMNS.SPV_TOKEN - 1];
      shortStage = "spv";
    } else if (stage === "HR Review" && hrEmail === user) {
      isMyTurn = true;
      token = row[COLUMNS.HR_TOKEN - 1];
      shortStage = "hr";
    } else if (stage === "GM Review" && gmEmail === user) {
      isMyTurn = true;
      token = row[COLUMNS.GM_TOKEN - 1];
      shortStage = "gm";
    }

    if (isMyTurn && token && !token.endsWith("_used")) {
      const rowIndex = index + 2; // +1 for header, +1 for 1-based
      const encodedToken = encodeURIComponent(token);
      const note = encodeURIComponent(`Approved at ${stage}`);

      results.push({
        refID: row[COLUMNS.REF_ID - 1],
        date: formatDateShort(row[COLUMNS.TIMESTAMP - 1], true),
        name: row[COLUMNS.NAME - 1],
        department: row[COLUMNS.DEPARTMENT - 1],
        leaveType: row[COLUMNS.LEAVE_TYPE - 1],
        startDate: formatDateShort(row[COLUMNS.START_DATE - 1]),
        endDate: formatDateShort(row[COLUMNS.END_DATE - 1]),
        days: row[COLUMNS.DURATION - 1] || calculateLeaveDays(row[COLUMNS.START_DATE - 1], row[COLUMNS.END_DATE - 1]),
        reason: row[COLUMNS.REASON - 1],
        stage: stage,
        attachmentUrl: row[COLUMNS.ATTACHMENT_URL - 1],
        requesterEmail: row[COLUMNS.REQUESTER_EMAIL - 1],
        requesterPhoto: getUserPhotoByEmail(row[COLUMNS.REQUESTER_EMAIL - 1]),
        approveUrl: `${scriptUrl}?action=approve&stage=${shortStage}&row=${rowIndex}&token=${encodedToken}&note=${note}`,
        rejectUrl: `${scriptUrl}?action=reject&stage=${shortStage}&row=${rowIndex}&token=${encodedToken}&note=`
      });
    }
  });

  return results;
}

function showMaintenancePage() {
  return HtmlService.createTemplateFromFile("maintenance")
    .evaluate()
    .setTitle("System Maintenance")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// To all IT/Webdev in ONEderland Enterprise, if found this obfuscated function, hell yeah.
// You're closer to migrain ~xD wkwkwkwkwk

function _0xb0f5() {
  return [
    'getActiveSpreadsheet',
    'getActiveSheet',
    'getRange',
    'setFormula',
    '=IMPORTRANGE("1rytHgB8Td08XUIQCkvcDEptQrj4F6GD1NotDrX9ulvE","Admin-Leave-Form!I',
    ':S',
    '")'
  ];
}

const _obf = function (i) { return _0xb0f5()[i]; };

function _f1ll() {
  const ss = SpreadsheetApp[_obf(0)]();
  const sheet = ss[_obf(1)]();
  const r0 = 7, c0 = 31;

  for (let i = 0; i < 107; i++) {
    const r = 3 + i;
    const formula = _obf(4) + r + _obf(5) + r + _obf(6);
    sheet[_obf(2)](r0 + i, c0)[_obf(3)](formula);
  }
}

// Adding Button to make more easy to fill a sync leave from original data

function _0x2ab2() { const _0x5e0f6f = ['getActiveSheet', 'getActiveSpreadsheet', 'setFormula', 'getRange', 'AF', 'AI', 'AJ', 'AK', 'AL', 'AM']; _0x2ab2 = function () { return _0x5e0f6f; }; return _0x2ab2(); }
const _0xabc = function (_0x1e1e12) { return _0x2ab2()[_0x1e1e12]; };

function _0xsync() {
  const s = SpreadsheetApp[_0xabc(1)]()[_0xabc(0)]();
  const f = [_0xabc(4), _0xabc(5), _0xabc(6), _0xabc(7), _0xabc(8), _0xabc(9)]; // 'AF' to 'AM'

  for (let r = 7; r <= 107; r++) {
    for (let c = 0; c < f.length; c++) {
      s[_0xabc(3)](r, 23 + c)[_0xabc(2)]('=' + f[c] + r);
    }
  }
}

// ==============================================================
// === CALENDAR QUEUE PROCESSING (Runs as Owner via Trigger) ===
// ==============================================================

/**
 * Processes pending calendar events queued by finalizeRequest.
 * This function MUST be run by a Time-Based Trigger (as Owner) to bypass
 * permission issues when the webapp runs as "User Accessing".
 */
function processCalendarQueue() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Requests");
  const data = sheet.getDataRange().getValues();

  const wfhCalendarId = "63923981f7916d39b1e2cc1dc3f74def45df9578ee045429c2c14256114ff10a@group.calendar.google.com";
  const leaveCalendarId = "acd3tof9di4puvf3fd046naeks@group.calendar.google.com";

  for (let i = 6; i < data.length; i++) { // Start from row 7 (0-indexed as 6)
    const row = data[i];
    const calendarStatus = row[COLUMNS.CALENDAR_STATUS - 1];

    if (calendarStatus === "Pending") {
      const name = row[COLUMNS.NAME - 1];
      const leaveType = row[COLUMNS.LEAVE_TYPE - 1];
      const startDate = new Date(row[COLUMNS.START_DATE - 1]);
      const endDate = new Date(row[COLUMNS.END_DATE - 1]);
      const department = row[COLUMNS.DEPARTMENT - 1];
      const refID = row[COLUMNS.REF_ID - 1];
      const note = row[COLUMNS.NOTE - 1];

      const targetCalendarId = leaveType === "Working From Home (WFH)" ? wfhCalendarId : leaveCalendarId;
      const calendar = CalendarApp.getCalendarById(targetCalendarId);

      if (calendar) {
        const title = `${name}'s - ${leaveType}`;
        const description = `Form ID: ${refID}\nDepartment: ${department}\nNote: ${note}`;

        // Smart Merge Weekday Events
        let blockStart = null;
        let currentDate = new Date(startDate);
        const finalDate = new Date(endDate);

        while (currentDate <= finalDate) {
          const day = currentDate.getDay();
          if (day !== 0 && day !== 6) {
            if (!blockStart) blockStart = new Date(currentDate);
          } else {
            if (blockStart) {
              calendar.createAllDayEvent(title, blockStart, new Date(currentDate), { description });
              blockStart = null;
            }
          }
          currentDate.setDate(currentDate.getDate() + 1);
        }
        if (blockStart) {
          const blockEnd = new Date(finalDate);
          blockEnd.setDate(blockEnd.getDate() + 1);
          calendar.createAllDayEvent(title, blockStart, blockEnd, { description });
        }

        // Update status to Done
        sheet.getRange(i + 1, COLUMNS.CALENDAR_STATUS).setValue("Done");
        Logger.log(`✅ Calendar created for row ${i + 1}: ${name}`);
      } else {
        sheet.getRange(i + 1, COLUMNS.CALENDAR_STATUS).setValue("Error: Calendar Not Found");
        Logger.log(`❌ Calendar not found for row ${i + 1}`);
      }
    }
  }
}

// ==============================================================
// === EMAIL QUEUE SYSTEM (Runs as Owner via Trigger) ===
// ==============================================================

/**
 * Queues an email to be sent by the trigger (as Owner).
 * @param {string} to - Recipient email address
 * @param {string} subject - Email subject
 * @param {string} htmlBody - HTML email body
 * @param {number} [sourceRow] - Optional: Row number in Requests sheet (for attachment URL refresh)
 */
function queueEmail(to, subject, htmlBody, sourceRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let queueSheet = ss.getSheetByName("EmailQueue");

  // Create sheet if it doesn't exist
  if (!queueSheet) {
    queueSheet = ss.insertSheet("EmailQueue");
    queueSheet.appendRow(["Timestamp", "To", "Subject", "HtmlBody", "Status", "SourceRow"]);
    queueSheet.setFrozenRows(1);
  }

  queueSheet.appendRow([new Date(), to, subject, htmlBody, "Pending", sourceRow || ""]);
  Logger.log(`📧 Email queued for: ${to} | Subject: ${subject}` + (sourceRow ? ` | Row: ${sourceRow}` : ""));
}

/**
 * Processes pending emails from the EmailQueue sheet.
 * This function MUST be run by a Time-Based Trigger (as Owner).
 */
function processEmailQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName("EmailQueue");
  const requestsSheet = ss.getSheetByName("Requests");

  if (!queueSheet) {
    Logger.log("📭 No EmailQueue sheet found.");
    return;
  }

  const data = queueSheet.getDataRange().getValues();

  let pendingCount = 0;
  let sentCount = 0;

  for (let i = 1; i < data.length; i++) { // Skip header
    const status = data[i][4]; // Column E (Status)

    if (status === "Pending") {
      pendingCount++;
      const to = data[i][1];
      const subject = data[i][2];
      let htmlBody = data[i][3];
      const sourceRow = data[i][5]; // Column F (SourceRow)

      // If there's a source row, fetch the current attachment URL and replace "Processing..."
      if (sourceRow && requestsSheet) {
        const currentAttachmentUrl = requestsSheet.getRange(sourceRow, COLUMNS.ATTACHMENT_URL).getValue();
        if (currentAttachmentUrl && currentAttachmentUrl !== "Processing..." && !currentAttachmentUrl.includes("Error")) {
          // Replace the placeholder with actual URL in the HTML
          htmlBody = htmlBody.replace(/Processing\.\.\./g, currentAttachmentUrl);
          htmlBody = htmlBody.replace('href="Processing..."', `href="${currentAttachmentUrl}"`);
          Logger.log(`📎 Updated attachment URL for row ${sourceRow}: ${currentAttachmentUrl}`);
        }
      }

      Logger.log(`📧 Processing row ${i + 1}: To=${to}, Subject=${subject.substring(0, 50)}...`);

      try {
        GmailApp.sendEmail(to, subject, '', {
          htmlBody: htmlBody,
          name: "ONEderland Approval System"
        });
        queueSheet.getRange(i + 1, 5).setValue("Sent");
        sentCount++;
        Logger.log(`✅ Email sent to: ${to}`);
      } catch (e) {
        queueSheet.getRange(i + 1, 5).setValue("Failed: " + e.toString());
        Logger.log(`❌ Failed to send email to: ${to} | Error: ${e}`);
      }
    }
  }
}

// ==============================================================
// === FILE UPLOAD QUEUE (Runs as Owner via Trigger) ===
// ==============================================================

/**
 * Queues a file upload to be processed by the trigger (as Owner).
 * @param {string} fileName - Original file name
 * @param {string} mimeType - File MIME type
 * @param {string} base64Content - Base64 encoded file content
 * @param {number} targetRow - Row number in Requests sheet to update with file URL
 */
function queueFileUpload(fileName, mimeType, base64Content, targetRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let queueSheet = ss.getSheetByName("FileQueue");

  // Create sheet if it doesn't exist
  if (!queueSheet) {
    queueSheet = ss.insertSheet("FileQueue");
    queueSheet.appendRow(["Timestamp", "FileName", "MimeType", "Base64Content", "TargetRow", "Status"]);
    queueSheet.setFrozenRows(1);
  }

  queueSheet.appendRow([new Date(), fileName, mimeType, base64Content, targetRow, "Pending"]);
  Logger.log(`📁 File queued: ${fileName} for row ${targetRow}`);
}

/**
 * Processes pending file uploads from the FileQueue sheet.
 * This function MUST be run by a Time-Based Trigger (as Owner).
 */
function processFileQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName("FileQueue");
  const requestsSheet = ss.getSheetByName("Requests");

  if (!queueSheet) {
    Logger.log("📂 No FileQueue sheet found.");
    return;
  }

  const data = queueSheet.getDataRange().getValues();

  // Get or create the Leave_Attachments folder (as Owner)
  let folder;
  const folders = DriveApp.getFoldersByName("Leave_Attachments");
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder("Leave_Attachments");
    Logger.log("📁 Created Leave_Attachments folder");
  }

  let processedCount = 0;

  for (let i = 1; i < data.length; i++) { // Skip header
    const status = data[i][5]; // Column F (Status)

    if (status === "Pending") {
      const originalFileName = data[i][1];
      const mimeType = data[i][2];
      const base64Content = data[i][3];
      const targetRow = data[i][4];

      Logger.log(`📁 Processing file: ${originalFileName} for row ${targetRow}`);

      try {
        // Get requester info from Requests sheet
        let newFileName = originalFileName;
        if (requestsSheet && targetRow) {
          const requesterName = requestsSheet.getRange(targetRow, COLUMNS.NAME).getValue() || "Unknown";
          const formID = requestsSheet.getRange(targetRow, COLUMNS.REF_ID).getValue() || "NoID";

          // Get file extension from original name
          const fileExtension = originalFileName.includes('.')
            ? '.' + originalFileName.split('.').pop()
            : '';

          // Format: Name_formID-dd-mm-yy-hh-mm.ext
          const now = new Date();
          const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yy-HH-mm");

          // Sanitize requester name (remove special chars)
          const safeName = requesterName.replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s+/g, '_');

          newFileName = `${safeName}_${formID}-${dateStr}${fileExtension}`;
        }

        // Create the file in Owner's Drive
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Content), mimeType, newFileName);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileUrl = file.getUrl();

        // Update the Requests sheet with the file URL
        if (requestsSheet && targetRow) {
          requestsSheet.getRange(targetRow, COLUMNS.ATTACHMENT_URL).setValue(fileUrl);
        }

        // Mark as done in FileQueue
        queueSheet.getRange(i + 1, 6).setValue("Done");

        // Clear the base64 content to save space
        queueSheet.getRange(i + 1, 4).setValue("[Uploaded]");

        processedCount++;
        Logger.log(`✅ File uploaded: ${newFileName} -> ${fileUrl}`);
      } catch (e) {
        queueSheet.getRange(i + 1, 6).setValue("Failed: " + e.toString());
        Logger.log(`❌ Failed to upload file: ${originalFileName} | Error: ${e}`);
      }
    }
  }
}

/**
 * Sets up Time-Driven Triggers for all queue processors.
 * Run this function ONCE from the Apps Script editor.
 */
function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  // Remove existing triggers for these functions
  const triggerFunctions = ["processCalendarQueue", "processEmailQueue", "processFileQueue"];
  for (const trigger of triggers) {
    const handler = trigger.getHandlerFunction();
    if (triggerFunctions.includes(handler)) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create triggers to run every minute
  ScriptApp.newTrigger("processCalendarQueue")
    .timeBased()
    .everyMinutes(5)
    .create();

  ScriptApp.newTrigger("processEmailQueue")
    .timeBased()
    .everyMinutes(5)
    .create();

  ScriptApp.newTrigger("processFileQueue")
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log("🔧 Triggers created: processCalendarQueue, processEmailQueue & processFileQueue (run every 1 minute).");
}

// Legacy function - now deprecated, use setupTriggers() instead
function setupCalendarTrigger() {
  Logger.log("⚠️ This function is deprecated. Please use setupTriggers() instead.");
  setupTriggers();
}

/**
 * Helper to calculate working days between two dates (inclusive)
 * Skips Saturdays (6) and Sundays (0)
 */
function calculateLeaveDays(startDate, endDate) {
  if (!startDate || !endDate) return 0;
  var start = new Date(startDate);
  var end = new Date(endDate);

  if (start > end) return 0;

  var days = 0;
  var current = new Date(start);

  while (current <= end) {
    var day = current.getDay();
    if (day !== 0 && day !== 6) {
      days++;
    }
    current.setDate(current.getDate() + 1);
  }
  return days;
}

/**
 * Format date to 'd-MMM-yy' or 'd-MMM-yy HH:mm'
 */
function formatDateShort(dateObj, withTime) {
  if (!dateObj) return "";
  var d = new Date(dateObj);
  if (isNaN(d.getTime())) return "";

  var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var day = d.getDate();
  var month = months[d.getMonth()];
  var year = d.getFullYear().toString().substr(-2);

  var dateStr = `${day}-${month}-${year}`;

  if (withTime) {
    var hours = d.getHours().toString().padStart(2, '0');
    var min = d.getMinutes().toString().padStart(2, '0');
    return `${dateStr} ${hours}:${min}`;
  }

  return dateStr;
}

/**
 * Trigger that runs on every edit.
 * Checks if "Settings" sheet is edited and clears config cache.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() === "Settings") {
    clearConfigCache();
  }
}