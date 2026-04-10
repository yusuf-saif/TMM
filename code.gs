/**
 * TMM Community Web App (Google Apps Script)
 * ---------------------------------------------------------
 * ✅ Saves submissions to Google Sheet
 * ✅ Sends a beautiful welcome email (logo header + CTA + footer)
 * ✅ Optionally notifies admin on every submission
 * ✅ Reusable Campaign Engine for Batch Emailing!
 *
 * IMPORTANT SETUP
 * 1) CONFIG.SPREADSHEET_ID must be ONLY the ID (not the full URL)
 * 2) CONFIG.COMMUNITY_LOGO_URL must be a PUBLIC direct image URL
 * 3) Deploy as Web App:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4) Run testAuth() once from the editor to authorize Sheets + Mail + Drive
 * 5) After any change: Deploy > Manage deployments > Edit > New version
 */

/**
 * =========================================================
 * HOW TO ADD A NEW EMAIL CAMPAIGN (DEVELOPER GUIDE)
 * =========================================================
 * Follow these steps to create a new campaign:
 * 
 * 1. Add Config/Assets: Add any needed Drive IDs to the CONFIG object.
 * 2. Create Template: Create a new email template function in "D. EMAIL TEMPLATE FUNCTIONS"
 *    e.g. function sendWebinarEmail_(email, name, context) { ... }
 * 3. Register Template: Add it to the EMAIL_TEMPLATES routing map in Section D:
 *    "webinar": sendWebinarEmail_
 * 4. Create Runner: Create a runner function in "F. CAMPAIGN RUNNER FUNCTIONS":
 *    function sendWebinarBatch() {
 *      sendCampaignEmailBatch_({
 *        baseColumnName: "Webinar", // Automatically creates "Webinar Sent", "Webinar Sent At", "Webinar Error"
 *        batchLimit: 100, // keep at 100 to avoid timeouts
 *        resend: false,
 *        templateName: "webinar",
 *        fileIds: { logoId: CONFIG.LOGO_FILE_ID } 
 *      });
 *    }
 * 5. Test & Run: Select your runner function in the Apps Script menu and click "Run".
 * =========================================================
 */

/** =========================
 * A. CONFIG (EDIT THESE)
 * ========================== */
const CONFIG = {
  // Spreadsheet ID & Tab
  SPREADSHEET_ID: "1BPncD_AAmAR09s05OkyKsiH-U1Q6CPO2EJIkKFZndKI",
  SHEET_NAME: "TMM Community",

  // Core Variables
  ADMIN_EMAIL: "themarketmasterng@gmail.com",
  WHATSAPP_GROUP_LINK: "https://chat.whatsapp.com/IT19ZTxvK90HLUqY1AKmTB",
  WEBSITE_URL: "https://themarketmasters.com.ng",

  // Brand Styling & Assets
  FROM_NAME: "The Market Masters (TMM)",
  BRAND_PRIMARY: "#4664E8",
  BRAND_INK: "#2C1B66",
  BRAND_SOFT: "#F6F7FF",
  COMMUNITY_LOGO_URL: "https://themarketmasters.com.ng/assets/img/logo.jpg", // Public Image 

  // Google Drive File IDs (For inline images or attachments)
  LOGO_FILE_ID: "1zo79XPAA4GJyUlMYnNdEQyH921_oYa27",
  FLYER_FILE_ID: "1byzkUnNKIeW4ShFOmrqn8KWRil1FsPoF",

  // Context Links (Session 1)
  SESSION_1_RESOURCE_LINK: "https://docs.google.com/presentation/d/1pvTNSbN0IanNj2n1sqih4et5yKga_pTH/edit?usp=sharing&ouid=112076881544159926854&rtpof=true&sd=true",
  SESSION_1_RECORDING_LINK: "https://drive.google.com/file/d/1OrJzxuw2lPNTYCWkxDy5zuT90lzwhAFp/view?usp=drivesdk", // Official recording link

  // Context Content (Events)
  EVENT_TITLE: "Excel For Business Mastery",
  EVENT_DATE_LABEL: "TODAY",
  EVENT_EMAIL_SUBJECT: "TODAY: TMM Community Meetup is Happening! 🚀",
  EVENT_JOIN_LINK: "https://meet.google.com/ocd-fome-omq",
  
  // Context Content (Next Meeting)
  NEXT_MEETING_TITLE: "Excel For Business Mastery - Session 2",
  NEXT_MEETING_DATE_LABEL: "TOMORROW at 10:00 AM prompt",
  NEXT_MEETING_EMAIL_SUBJECT: "TOMORROW: Excel For Business Mastery, Session 2 is Happening! 🚀"
};

/** =========================
 * B. WEB APP ENDPOINTS
 * ========================== */

function doPost(e) {
  try {
    const body = parseJson_(e);

    // Honeypot anti-bot
    if (body.website && String(body.website).trim() !== "") {
      return json_({ ok: true, ignored: true });
    }

    // Validate required fields
    const required = [
      "fullName", "phone", "email", "location", "businessName",
      "industry", "businessStage", "registered", "operationMode",
      "offerings", "goals", "referralSource", "consent"
    ];

    const missing = required.filter((k) => !body[k] || String(body[k]).trim() === "");
    if (missing.length) {
      return json_({ ok: false, error: "Missing required fields: " + missing.join(", ") });
    }

    // Sheet setup
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
    ensureHeader_(sheet);
    
    // Write row explicitly without missing tracking column space
    const rowNum = appendRow_(sheet, body);

    // Send the Welcome Email immediately, but handle its success/fail properly
    try {
      sendWelcomeEmail_(body.email, body.fullName);
      
      const sentCol = getColumnIndexByName_(sheet, "Welcome Email Sent") + 1;
      const timeCol = getColumnIndexByName_(sheet, "Welcome Email Sent At") + 1;
      
      sheet.getRange(rowNum, sentCol).setValue("YES");
      sheet.getRange(rowNum, timeCol).setValue(new Date());
    } catch(err) {
      const sentCol = getColumnIndexByName_(sheet, "Welcome Email Sent") + 1;
      const errCol = getColumnIndexByName_(sheet, "Welcome Email Error") + 1;
      
      sheet.getRange(rowNum, sentCol).setValue("FAILED");
      sheet.getRange(rowNum, errCol).setValue(String(err.message));
    }

    // Alert Admin
    if (CONFIG.ADMIN_EMAIL) {
      notifyAdmin_(body);
    }

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.stack ? err.stack : err) });
  }
}

function doGet() {
  return json_({ ok: true, message: "TMM Community API running.", time: new Date().toISOString() });
}

/** =========================
 * C. SHEET HELPERS
 * ========================== */

function ensureHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow([
    "Timestamp", "Full Name", "Phone", "Email", "Location", "Age Range",
    "Business Name", "Industry", "Business Stage", "Registered",
    "Business Link", "Staff Count", "Operation Mode", "Offerings",
    "Revenue Range", "Challenges", "Challenges (Other)", "Support Needed",
    "Support Format", "Goals (6–12 months)", "Success Definition",
    "Heard About Us", "Comments", "Consent",
    "Welcome Email Sent", "Welcome Email Sent At", "Welcome Email Error",
    "Session 1 Thank You Sent", "Session 1 Thank You Sent At", "Session 1 Thank You Error",
    "Event Invite Sent", "Event Invite Sent At", "Event Invite Error",
    "Next Meeting Reminder Sent", "Next Meeting Reminder Sent At", "Next Meeting Reminder Error",
    "Session 2 Reminder Sent", "Session 2 Reminder Sent At", "Session 2 Reminder Error",
    "Session 2 Thank You Sent", "Session 2 Thank You Sent At", "Session 2 Thank You Error",
    "Session 3 Reminder Sent", "Session 3 Reminder Sent At", "Session 3 Reminder Error",
    "Session 3 Thank You Sent", "Session 3 Thank You Sent At", "Session 3 Thank You Error",
    "Session 4 Reminder Sent", "Session 4 Reminder Sent At", "Session 4 Reminder Error",
    "Session 4 Thank You Sent", "Session 4 Thank You Sent At", "Session 4 Thank You Error",
    "Session 5 Reminder Sent", "Session 5 Reminder Sent At", "Session 5 Reminder Error",
    "Session 5 Thank You Sent", "Session 5 Thank You Sent At", "Session 5 Thank You Error",
    "Session 6 Reminder Sent", "Session 6 Reminder Sent At", "Session 6 Reminder Error",
    "Session 6 Thank You Sent", "Session 6 Thank You Sent At", "Session 6 Thank You Error"
  ]);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function appendRow_(sheet, body) {
  sheet.appendRow([
    new Date(), body.fullName, body.phone, body.email, body.location, body.ageRange || "",
    body.businessName, body.industry, body.businessStage, body.registered,
    body.businessLink || "", body.staffCount || "", body.operationMode, body.offerings,
    body.revenueRange || "", body.challenges || "", body.challengesOther || "",
    body.support || "", body.supportFormat || "", body.goals, body.successDefinition || "",
    body.referralSource, body.comments || "", body.consent
  ]);
  return sheet.getLastRow();
}

/** Helper to find a column index by name. Creates it if missing. */
function getColumnIndexByName_(sheet, colName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let idx = headers.indexOf(colName);
  
  if (idx === -1) {
    idx = headers.length; // Append at the end (0-based)
    sheet.getRange(1, idx + 1).setValue(colName).setFontWeight("bold");
  }
  return idx;
}

/** =========================
 * D. EMAIL TEMPLATE FUNCTIONS
 * ========================== */

/** 1. Welcome Email Template (Used in doPost) */
function sendWelcomeEmail_(toEmail, fullName) {
  const safeName = escapeHtml_(fullName);
  const subject = "Welcome to the TMM Community ✅";

  const textBody =
    `Hi ${fullName},\n\nWelcome to The Market Masters (TMM) Community!\n\n` +
    `Join the WhatsApp Community here:\n${CONFIG.WHATSAPP_GROUP_LINK}\n\n` +
    `Warm regards,\n${CONFIG.FROM_NAME}`;

  const htmlBody = `
  <div style="background:#f3f5ff;padding:24px 0;">
    <table align="center" width="600" style="width:600px;background:#ffffff;border-radius:16px;border:1px solid rgba(15,23,42,.10);">
      <tr>
        <td style="background:linear-gradient(135deg, ${CONFIG.BRAND_PRIMARY}, ${CONFIG.BRAND_INK});padding:18px 22px;">
          <table width="100%">
            <tr>
              <td width="56" valign="middle"><img src="${CONFIG.COMMUNITY_LOGO_URL}" width="46" style="border-radius:12px;background:#fff;padding:6px;"></td>
              <td valign="middle" style="padding-left:12px;">
                <div style="color:#ffffff;font-family:Arial,sans-serif;font-weight:800;font-size:16px;">The Market Masters (TMM)</div>
                <div style="color:rgba(255,255,255,.86);font-family:Arial,sans-serif;font-size:12.5px;">Community Onboarding</div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td style="padding:26px 22px 10px 22px;font-family:Arial,sans-serif;font-size:15px;color:#0f172a;line-height:1.7;">
          <p>Hi <strong>${safeName}</strong>,</p>
          <p>Welcome to the <strong>TMM Community</strong> 🎉 We’re excited to have you in a network built for growth.</p>
          <div style="margin:16px 0;padding:14px;border-radius:14px;background:${CONFIG.BRAND_SOFT};">
            <div style="font-weight:800;">Next step</div>
            <div>Join the WhatsApp community using the button below.</div>
          </div>
        </td>
      </tr>
      <tr>
        <td align="center" style="padding:0 22px 22px 22px;">
          <a href="${CONFIG.WHATSAPP_GROUP_LINK}" style="display:inline-block;background:${CONFIG.BRAND_PRIMARY};padding:12px 18px;color:#fff;text-decoration:none;border-radius:999px;font-family:Arial,sans-serif;font-weight:800;">Join WhatsApp Community →</a>
        </td>
      </tr>
      <tr>
        <td style="padding:16px 22px 22px 22px;font-family:Arial,sans-serif;font-size:12.5px;color:rgba(15,23,42,.70);border-top:1px solid rgba(15,23,42,.10);">
          Need help? <a href="mailto:${CONFIG.ADMIN_EMAIL}" style="color:${CONFIG.BRAND_INK};font-weight:800;">${CONFIG.ADMIN_EMAIL}</a><br><br>
          <strong>${CONFIG.FROM_NAME}</strong> • <a href="${CONFIG.WEBSITE_URL}" style="color:${CONFIG.BRAND_INK};">${CONFIG.WEBSITE_URL}</a>
        </td>
      </tr>
    </table>
  </div>`;

  MailApp.sendEmail({ to: toEmail, subject: subject, name: CONFIG.FROM_NAME, body: textBody, htmlBody: htmlBody });
}

/** 2. Session 1 Thank You Template */
function sendSession1ThankYouEmail_(toEmail, fullName, context) {
  const safeName = escapeHtml_(fullName);
  const subject = "Thank you for attending Session 1! ✨";
  
  const hasRecording = CONFIG.SESSION_1_RECORDING_LINK && String(CONFIG.SESSION_1_RECORDING_LINK).trim() !== "";

  const textBody =
    `Hi ${fullName},\n\nThank you for joining us for the first session of Excel for Business Mastery. We appreciate the time, energy, and attention you brought to the class.\n\n` +
    `You can access today's resource slide deck here:\n${CONFIG.SESSION_1_RESOURCE_LINK}\n\n` +
    (hasRecording ? `Watch the session recording here:\n${CONFIG.SESSION_1_RECORDING_LINK}\n\n` : "") +
    `Keep an eye on the WhatsApp group for the next upcoming modules!\n\n` +
    `Warm regards,\n${CONFIG.FROM_NAME}`;

  const htmlBody = `
  <div style="background:#f3f5ff;padding:24px 0;">
    <table align="center" width="600" style="width:600px;background:#ffffff;border-radius:16px;border:1px solid rgba(15,23,42,.10);">
      <tr>
        <td style="background:linear-gradient(135deg, ${CONFIG.BRAND_PRIMARY}, ${CONFIG.BRAND_INK});padding:18px 22px;">
          <table width="100%">
            <tr>
              <td width="56" valign="middle"><img src="cid:logo_image" width="46" style="border-radius:12px;background:#fff;padding:6px;"></td>
              <td valign="middle" style="padding-left:12px;">
                <div style="color:#ffffff;font-family:Arial,sans-serif;font-weight:800;font-size:16px;">The Market Masters (TMM)</div>
                <div style="color:rgba(255,255,255,.86);font-family:Arial,sans-serif;font-size:12.5px;">Session 1 Recap</div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td style="padding:26px 22px 10px 22px;font-family:Arial,sans-serif;font-size:15px;color:#0f172a;line-height:1.7;">
          <p>Hi <strong>${safeName}</strong>,</p>
          <p>Thank you for joining us for the first session of <strong>Excel for Business Mastery</strong>. We appreciate the time, energy, and attention you brought to the class.</p>
          <div style="margin:16px 0;padding:14px;border-radius:14px;background:${CONFIG.BRAND_SOFT};border:1px solid rgba(70,100,232,.20);">
            <div style="font-weight:800;margin-bottom:6px;">Today's Resources</div>
            <div style="color:rgba(15,23,42,.86);margin-bottom:12px;">Access the session slide deck ${hasRecording ? "and recording" : ""} using the button(s) below.</div>
            <a href="${CONFIG.SESSION_1_RESOURCE_LINK}" target="_blank" rel="noreferrer" style="display:inline-block;background:${CONFIG.BRAND_PRIMARY};padding:10px 16px;color:#fff;margin-right:8px;margin-bottom:8px;text-decoration:none;border-radius:999px;font-family:Arial,sans-serif;font-weight:800;font-size:13px;">View Presentation →</a>
            ${hasRecording ? `<a href="${CONFIG.SESSION_1_RECORDING_LINK}" target="_blank" rel="noreferrer" style="display:inline-block;background:${CONFIG.BRAND_INK};padding:10px 16px;color:#fff;margin-bottom:8px;text-decoration:none;border-radius:999px;font-family:Arial,sans-serif;font-weight:800;font-size:13px;">Watch Recording →</a>` : ""}
          </div>
          <div style="margin:16px 0;padding:14px;border-radius:14px;background:${CONFIG.BRAND_SOFT};">
            <div style="font-weight:800;margin-bottom:6px;">What's next?</div>
            <div style="color:rgba(15,23,42,.86);">Keep an eye on the WhatsApp group for the next upcoming modules and resources.</div>
          </div>
        </td>
      </tr>
      <tr>
        <td style="padding:16px 22px 22px 22px;font-family:Arial,sans-serif;font-size:12.5px;color:rgba(15,23,42,.70);border-top:1px solid rgba(15,23,42,.10);">
          Need help? <a href="mailto:${CONFIG.ADMIN_EMAIL}" style="color:${CONFIG.BRAND_INK};font-weight:800;">${CONFIG.ADMIN_EMAIL}</a><br><br>
          <strong>${CONFIG.FROM_NAME}</strong> • <a href="${CONFIG.WEBSITE_URL}" style="color:${CONFIG.BRAND_INK};">${CONFIG.WEBSITE_URL}</a>
        </td>
      </tr>
    </table>
  </div>`;

  MailApp.sendEmail({ 
    to: toEmail, subject: subject, name: CONFIG.FROM_NAME, body: textBody, htmlBody: htmlBody,
    inlineImages: context.logoBlob ? { logo_image: context.logoBlob } : undefined 
  });
}

/** 3. Event Reminder/Invite Template */
function sendEventMail_(toEmail, fullName, context) {
  const safeName = escapeHtml_(fullName);
  const subject = CONFIG.EVENT_EMAIL_SUBJECT;

  const textBody =
    `Hi ${fullName},\n\nOur special event, ${CONFIG.EVENT_TITLE}, is happening ${CONFIG.EVENT_DATE_LABEL}! We can't wait to see you.\n\n` +
    `Join here: ${CONFIG.EVENT_JOIN_LINK}\n\n` +
    `Please find the programme flyer attached.\n\n` +
    `Warm regards,\n${CONFIG.FROM_NAME}`;

  const htmlBody = `
  <div style="background:#f3f5ff;padding:24px 0;">
    <table align="center" width="600" style="width:600px;background:#ffffff;border-radius:16px;border:1px solid rgba(15,23,42,.10);">
      <tr>
        <td style="background:linear-gradient(135deg, ${CONFIG.BRAND_PRIMARY}, ${CONFIG.BRAND_INK});padding:18px 22px;">
          <table width="100%">
            <tr>
              <td width="56" valign="middle"><img src="cid:logo_image" width="46" style="border-radius:12px;background:#fff;padding:6px;"></td>
              <td valign="middle" style="padding-left:12px;">
                <div style="color:#ffffff;font-family:Arial,sans-serif;font-weight:800;font-size:16px;">The Market Masters (TMM)</div>
                <div style="color:rgba(255,255,255,.86);font-family:Arial,sans-serif;font-size:12.5px;">Event Registration</div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td style="padding:26px 22px 10px 22px;font-family:Arial,sans-serif;font-size:15px;color:#0f172a;line-height:1.7;">
          <p>Hi <strong>${safeName}</strong>,</p>
          <p>Our special event, <strong>${CONFIG.EVENT_TITLE}</strong>, is happening <strong>${CONFIG.EVENT_DATE_LABEL}</strong>! We can't wait to see you there. 
             Please find the programme flyer attached to this email.</p>
          <div style="margin:16px 0;padding:14px;border-radius:14px;background:${CONFIG.BRAND_SOFT};">
            <div style="font-weight:800;">How to join</div>
            <div>Click the button below to enter the Google Meet instantly.</div>
          </div>
        </td>
      </tr>
      <tr>
        <td align="center" style="padding:0 22px 22px 22px;">
          <a href="${CONFIG.EVENT_JOIN_LINK}" style="display:inline-block;background:${CONFIG.BRAND_PRIMARY};padding:12px 18px;color:#fff;text-decoration:none;border-radius:999px;font-family:Arial,sans-serif;font-weight:800;">Join The Event →</a>
        </td>
      </tr>
      <tr>
        <td align="center" style="padding:0 22px 22px 22px;font-family:Arial,sans-serif;font-size:12px;color:rgba(15,23,42,.65);">
          If the button doesn’t work, join via this link:<br>
          <a href="${CONFIG.EVENT_JOIN_LINK}" style="color:${CONFIG.BRAND_INK};">${CONFIG.EVENT_JOIN_LINK}</a>
        </td>
      </tr>
      <tr>
        <td style="padding:16px 22px 22px 22px;font-family:Arial,sans-serif;font-size:12.5px;color:rgba(15,23,42,.70);border-top:1px solid rgba(15,23,42,.10);">
          Need help? <a href="mailto:${CONFIG.ADMIN_EMAIL}" style="color:${CONFIG.BRAND_INK};font-weight:800;">${CONFIG.ADMIN_EMAIL}</a><br><br>
          <strong>${CONFIG.FROM_NAME}</strong> • <a href="${CONFIG.WEBSITE_URL}" style="color:${CONFIG.BRAND_INK};">${CONFIG.WEBSITE_URL}</a>
        </td>
      </tr>
    </table>
  </div>`;

  const mailOptions = {
    to: toEmail, subject: subject, name: CONFIG.FROM_NAME, body: textBody, htmlBody: htmlBody
  };
  
  if (context.logoBlob) mailOptions.inlineImages = { logo_image: context.logoBlob };
  if (context.flyerBlob) mailOptions.attachments = [context.flyerBlob];

  MailApp.sendEmail(mailOptions);
}

/** 4. Next Meeting Reminder Template */
function sendNextMeetingReminderMail_(toEmail, fullName, context) {
  const safeName = escapeHtml_(fullName);
  const subject = CONFIG.NEXT_MEETING_EMAIL_SUBJECT;
  
  const hasRecording = CONFIG.SESSION_1_RECORDING_LINK && String(CONFIG.SESSION_1_RECORDING_LINK).trim() !== "";

  const textBody =
    `Hi ${fullName},\n\nThis is a quick reminder that our next session for ${CONFIG.NEXT_MEETING_TITLE} is happening ${CONFIG.NEXT_MEETING_DATE_LABEL}!\n\n` +
    `Join here tomorrow: ${CONFIG.EVENT_JOIN_LINK}\n\n` +
    `In case you missed the previous session, you can catch up with the resources below:\n` +
    `Resources: ${CONFIG.SESSION_1_RESOURCE_LINK}\n` +
    (hasRecording ? `Recording: ${CONFIG.SESSION_1_RECORDING_LINK}\n\n` : "\n") +
    `Warm regards,\n${CONFIG.FROM_NAME}`;

  const htmlBody = `
  <div style="background:#f3f5ff;padding:24px 0;">
    <table align="center" width="600" style="width:600px;background:#ffffff;border-radius:16px;border:1px solid rgba(15,23,42,.10);">
      <tr>
        <td style="background:linear-gradient(135deg, ${CONFIG.BRAND_PRIMARY}, ${CONFIG.BRAND_INK});padding:18px 22px;">
          <table width="100%">
            <tr>
              <td width="56" valign="middle"><img src="cid:logo_image" width="46" style="border-radius:12px;background:#fff;padding:6px;"></td>
              <td valign="middle" style="padding-left:12px;">
                <div style="color:#ffffff;font-family:Arial,sans-serif;font-weight:800;font-size:16px;">The Market Masters (TMM)</div>
                <div style="color:rgba(255,255,255,.86);font-family:Arial,sans-serif;font-size:12.5px;">Meeting Reminder</div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td style="padding:26px 22px 10px 22px;font-family:Arial,sans-serif;font-size:15px;color:#0f172a;line-height:1.7;">
          <p>Hi <strong>${safeName}</strong>,</p>
          <p>This is a quick reminder that our next virtual meetup: <strong>${CONFIG.NEXT_MEETING_TITLE}</strong> is happening <strong>${CONFIG.NEXT_MEETING_DATE_LABEL}</strong>! We're excited to learn and grow together.</p>
          
          <div style="margin:16px 0;padding:14px;border-radius:14px;background:${CONFIG.BRAND_SOFT};">
            <div style="font-weight:800;margin-bottom:6px;">Join Tomorrow's Session</div>
            <div>Click the button below to enter the Google Meet at the scheduled time.</div>
             <a href="${CONFIG.EVENT_JOIN_LINK}" style="display:inline-block;background:${CONFIG.BRAND_PRIMARY};padding:10px 16px;color:#fff;margin-top:10px;text-decoration:none;border-radius:999px;font-family:Arial,sans-serif;font-weight:800;font-size:13px;">Join The Meeting →</a>
          </div>

          <div style="margin:16px 0;padding:14px;border-radius:14px;background:${CONFIG.BRAND_SOFT};border:1px solid rgba(70,100,232,.20);">
            <div style="font-weight:800;margin-bottom:6px;">Previous Session Catch-up</div>
            <div style="color:rgba(15,23,42,.86);margin-bottom:12px;">In case you missed the previous session or want to review, here are the resources:</div>
            <a href="${CONFIG.SESSION_1_RESOURCE_LINK}" target="_blank" rel="noreferrer" style="display:inline-block;background:${CONFIG.BRAND_PRIMARY};padding:10px 16px;color:#fff;margin-right:8px;margin-bottom:8px;text-decoration:none;border-radius:999px;font-family:Arial,sans-serif;font-weight:800;font-size:13px;">View Presentation →</a>
            ${hasRecording ? `<a href="${CONFIG.SESSION_1_RECORDING_LINK}" target="_blank" rel="noreferrer" style="display:inline-block;background:${CONFIG.BRAND_INK};padding:10px 16px;color:#fff;margin-bottom:8px;text-decoration:none;border-radius:999px;font-family:Arial,sans-serif;font-weight:800;font-size:13px;">Watch Recording →</a>` : ""}
          </div>
        </td>
      </tr>
      <tr>
        <td style="padding:16px 22px 22px 22px;font-family:Arial,sans-serif;font-size:12.5px;color:rgba(15,23,42,.70);border-top:1px solid rgba(15,23,42,.10);">
          Need help? <a href="mailto:${CONFIG.ADMIN_EMAIL}" style="color:${CONFIG.BRAND_INK};font-weight:800;">${CONFIG.ADMIN_EMAIL}</a><br><br>
          <strong>${CONFIG.FROM_NAME}</strong> • <a href="${CONFIG.WEBSITE_URL}" style="color:${CONFIG.BRAND_INK};">${CONFIG.WEBSITE_URL}</a>
        </td>
      </tr>
    </table>
  </div>`;

  const mailOptions = {
    to: toEmail, subject: subject, name: CONFIG.FROM_NAME, body: textBody, htmlBody: htmlBody
  };
  
  if (context.logoBlob) mailOptions.inlineImages = { logo_image: context.logoBlob };
  if (context.flyerBlob) mailOptions.attachments = [context.flyerBlob];

  MailApp.sendEmail(mailOptions);
}

/** 
 * 5. TEMPLATE ROUTING MAP
 * Maps template names to their respective functions above.
 */
const EMAIL_TEMPLATES = {
  "session1": sendSession1ThankYouEmail_,
  "event": sendEventMail_,
  "next_meeting": sendNextMeetingReminderMail_
};


/** =========================
 * E. GENERIC BATCH SENDER ENGINE
 * ========================== */

function sendCampaignEmailBatch_(options) {
  const { baseColumnName, batchLimit = 100, resend = false, templateName, fileIds = {} } = options;
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`Critical Error: Sheet "${CONFIG.SHEET_NAME}" not found. Please verify your CONFIG.SHEET_NAME.`);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  
  const headers = values[0];
  const emailIdx = headers.indexOf("Email");
  const nameIdx = headers.indexOf("Full Name");
  
  if (emailIdx === -1 || nameIdx === -1) {
    throw new Error("Critical Error: Could not find 'Email' or 'Full Name' columns in the dataset.");
  }
  
  // Create / Verify Tracking Columns
  const sentColIdx = getColumnIndexByName_(sheet, baseColumnName + " Sent");
  const timeColIdx = getColumnIndexByName_(sheet, baseColumnName + " Sent At");
  const errColIdx = getColumnIndexByName_(sheet, baseColumnName + " Error");
  
  // Verify Template Exists
  const templateFn = EMAIL_TEMPLATES[templateName];
  if (!templateFn) {
    throw new Error(`Template Router Error: Unknown template name "${templateName}" requested.`);
  }
  
  // Load Drive files exactly once per batch to prevent execution timeouts
  const context = {};
  if (fileIds.logoId || fileIds.attachId) {
    try {
      if (fileIds.logoId) context.logoBlob = DriveApp.getFileById(fileIds.logoId).getBlob().setName("logo_image");
      if (fileIds.attachId) context.flyerBlob = DriveApp.getFileById(fileIds.attachId).getBlob();
    } catch (err) {
      throw new Error(`Permissions Error: Could not access Drive files for ${templateName} template. Error: ${err.message}. Have you run testAuth()?`);
    }
  }
  
  let sentCount = 0;
  
  for (let r = 1; r < values.length; r++) {
    if (sentCount >= batchLimit) {
      Logger.log("Batch limit of " + batchLimit + " reached. Run script again to send in next cycle.");
      break;
    }
    
    const rowNum = r + 1;
    const email = String(values[r][emailIdx] || "").trim();
    const name = String(values[r][nameIdx] || "").trim();
    const deliveryStatus = String(values[r][sentColIdx] || "").trim().toUpperCase();
    
    // Skip out if no email, or already processed (Unless using resend)
    if (!email || (deliveryStatus === "YES" && !resend)) continue;
    
    try {
      // Direct routing execution
      templateFn(email, name, context);
      
      // Update successful statuses
      sheet.getRange(rowNum, sentColIdx + 1).setValue("YES");
      sheet.getRange(rowNum, timeColIdx + 1).setValue(new Date());
      sheet.getRange(rowNum, errColIdx + 1).setValue(""); // clear history if resend
      sentCount++;
    } catch (err) {
      Logger.log("Delivery Failed for " + email + ": " + err.message);
      sheet.getRange(rowNum, sentColIdx + 1).setValue("FAILED");
      sheet.getRange(rowNum, errColIdx + 1).setValue(err.message);
    }
  }
  
  Logger.log(`Batch Complete: Sent ${sentCount} '${templateName}' emails.`);
}


/** =========================
 * F. CAMPAIGN RUNNER FUNCTIONS
 * ========================== */

/** RUN THIS: Session 1 Follow Up */
function sendSession1ThankYouBatch() {
  sendCampaignEmailBatch_({
    baseColumnName: "Session 1 Thank You", // Will target "Session 1 Thank You Sent", "Session 1 Thank You Sent At", ...
    batchLimit: 100, // Important to prevent GAS runtime limits
    resend: false, 
    templateName: "session1",
    fileIds: { logoId: CONFIG.LOGO_FILE_ID } 
  });
}

/** RUN THIS: Today's Event Reminder */
function sendTodayEventMailBatch() {
  sendCampaignEmailBatch_({
    baseColumnName: "Event Invite", // Will target "Event Invite Sent", "Event Invite Sent At", ...
    batchLimit: 100,
    resend: false,
    templateName: "event",
    fileIds: { logoId: CONFIG.LOGO_FILE_ID, attachId: CONFIG.FLYER_FILE_ID }
  });
}

/** RUN THIS: Tomorrow's Meeting Reminder */
function sendNextMeetingReminderBatch() {
  sendCampaignEmailBatch_({
    baseColumnName: "Next Meeting Reminder", 
    batchLimit: 100,
    resend: false,
    templateName: "next_meeting",
    fileIds: { logoId: CONFIG.LOGO_FILE_ID, attachId: CONFIG.FLYER_FILE_ID } // includes flyer too just in case
  });
}

/** 
 * TEST FUNCTION: Send a test of the Next Meeting Reminder to the Admin Email 
 * Run this to preview how the email looks before sending the batch.
 */
function testNextMeetingReminder() {
  const adminEmail = Session.getActiveUser().getEmail() || CONFIG.ADMIN_EMAIL;
  Logger.log("Sending test email to: " + adminEmail);
  
  const context = {};
  if (CONFIG.LOGO_FILE_ID) {
    try {
      context.logoBlob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob().setName("logo_image");
    } catch(e) { Logger.log("Could not load logo: " + e.message); }
  }
  if (CONFIG.FLYER_FILE_ID) {
    try {
      context.flyerBlob = DriveApp.getFileById(CONFIG.FLYER_FILE_ID).getBlob();
    } catch(e) { Logger.log("Could not load flyer: " + e.message); }
  }

  sendNextMeetingReminderMail_(adminEmail, "Test Admin User", context);
  Logger.log("Test email sent successfully! Check your inbox.");
}


/** =========================
 * G. UTILITY HELPERS
 * ========================== */

/** RUN THIS: Update Sheet Headers Manually */
function updateSheetHeaders() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;
  
  const columns = [
    "Next Meeting Reminder",
    "Session 2 Reminder", "Session 2 Thank You",
    "Session 3 Reminder", "Session 3 Thank You",
    "Session 4 Reminder", "Session 4 Thank You",
    "Session 5 Reminder", "Session 5 Thank You",
    "Session 6 Reminder", "Session 6 Thank You"
  ];
  
  columns.forEach(c => {
    getColumnIndexByName_(sheet, c + " Sent");
    getColumnIndexByName_(sheet, c + " Sent At");
    getColumnIndexByName_(sheet, c + " Error");
  });
  
  Logger.log("Sheet tracking columns have been added!");
}

function notifyAdmin_(body) {
  const subject = "New Community Registration — " + body.fullName;
  const msg =
    "New registration:\n\n" +
    `Name: ${body.fullName}\nPhone: ${body.phone}\nEmail: ${body.email}\n` +
    `Location: ${body.location}\nBusiness: ${body.businessName}\n` +
    `Industry: ${body.industry}\nStage: ${body.businessStage}\n`;
  MailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, msg);
}

/**
 * Parses JSON safely from the request payload.
 */
function parseJson_(e) {
  if (!e) {
    throw new Error("No event data received (e is undefined). This endpoint expects a POST request.");
  }
  if (!e.postData || !e.postData.contents) {
    throw new Error("Empty POST payload. Ensure you're sending a physical JSON body in the fetch() request.");
  }
  
  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    throw new Error("Malformed JSON received: " + err.message);
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml_(str) {
  return String(str || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

/** RUN THIS ONCE AFTER PASTING CODE */
function testAuth() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
  ensureHeader_(sheet);
  
  DriveApp.getFileById(CONFIG.LOGO_FILE_ID);
  if (CONFIG.FLYER_FILE_ID) DriveApp.getFileById(CONFIG.FLYER_FILE_ID);
  
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "TMM Auth Test", "Authorization and Drive permissions work.");
}
