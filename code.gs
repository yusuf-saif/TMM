/**
 * TMM Community Web App (Google Apps Script)
 * ---------------------------------------------------------
 * ✅ Saves submissions to Google Sheet
 * ✅ Sends a beautiful welcome email (logo header + CTA + footer)
 * ✅ Optionally notifies admin on every submission
 *
 * IMPORTANT SETUP
 * 1) CONFIG.SPREADSHEET_ID must be ONLY the ID (not the full URL)
 * 2) CONFIG.COMMUNITY_LOGO_URL must be a PUBLIC direct image URL
 * 3) Deploy as Web App:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4) Run testAuth() once from the editor to authorize Sheets + Mail
 * 5) After any change: Deploy > Manage deployments > Edit > New version
 *
 * FRONTEND (your community.html)
 * - Posts JSON to your /exec URL:
 *   fetch(WEB_APP_URL, { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload) })
 */

/** =========================
 * CONFIG (EDIT THESE)
 * ========================== */
const CONFIG = {
  // ✅ Spreadsheet ID only (looks like: 1BPncD_AAmAR09s05OkyKsiH-U1Q6CPO2EJIkKFZndKI)
  SPREADSHEET_ID: "1BPncD_AAmAR09s05OkyKsiH-U1Q6CPO2EJIkKFZndKI",

  // ✅ Tab name inside the sheet
  SHEET_NAME: "TMM Community",

  // ✅ Admin email (optional)
  ADMIN_EMAIL: "themarketmasterng@gmail.com",

  // ✅ WhatsApp invite link
  WHATSAPP_GROUP_LINK: "https://chat.whatsapp.com/IT19ZTxvK90HLUqY1AKmTB",

  // ✅ Sender display name
  FROM_NAME: "The Market Masters (TMM)",

  // ✅ Brand styling for the email
  BRAND_PRIMARY: "#4664E8",
  BRAND_INK: "#2C1B66",
  BRAND_SOFT: "#F6F7FF",

  // ✅ Website (footer)
  WEBSITE_URL: "https://themarketmasters.com.ng",

  // ✅ MUST be a public direct image URL (host it on your site)
  COMMUNITY_LOGO_URL: "https://themarketmasters.com.ng/assets/img/logo.jpg",

  // ✅ EVENT ASSET IDs (From Google Drive Links)
  LOGO_FILE_ID: "1zo79XPAA4GJyUlMYnNdEQyH921_oYa27",
  FLYER_FILE_ID: "1byzkUnNKIeW4ShFOmrqn8KWRil1FsPoF",
  EVENT_JOIN_LINK: "https://meet.google.com/ocd-fome-omq"
};

/** =========================
 * WEB APP ENDPOINTS
 * ========================== */

/**
 * POST /exec
 * Receives JSON payload from the website.
 */
function doPost(e) {
  try {
    const body = parseJson_(e);

    // Honeypot anti-bot (ignore silently)
    if (body.website && String(body.website).trim() !== "") {
      return json_({ ok: true, ignored: true });
    }

    // Validate required fields (matches your community.html wizard)
    const required = [
      "fullName",
      "phone",
      "email",
      "location",
      "businessName",
      "industry",
      "businessStage",
      "registered",
      "operationMode",
      "offerings",
      "goals",
      "referralSource",
      "consent"
    ];

    const missing = required.filter((k) => !body[k] || String(body[k]).trim() === "");
    if (missing.length) {
      return json_({ ok: false, error: "Missing required fields: " + missing.join(", ") });
    }

    // Ensure sheet exists + header
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    ensureHeader_(sheet);

    // Save data to sheet
    appendRow_(sheet, body);

    // Send welcome email to user
    sendWelcomeEmail_(body.email, body.fullName);

    // Optional admin alert
    if (CONFIG.ADMIN_EMAIL) {
      notifyAdmin_(body);
    }

    return json_({ ok: true });
  } catch (err) {
    // Helpful error message (also check "Executions" in Apps Script)
    return json_({ ok: false, error: String(err && err.stack ? err.stack : err) });
  }
}

/**
 * GET /exec
 * Useful for quick checks in browser.
 */
function doGet() {
  return json_({
    ok: true,
    message: "TMM Community API running.",
    time: new Date().toISOString()
  });
}

/** =========================
 * SHEET HELPERS
 * ========================== */

function ensureHeader_(sheet) {
  // Only add header if the sheet is empty
  if (sheet.getLastRow() > 0) return;

  sheet.appendRow([
    "Timestamp",
    "Full Name",
    "Phone",
    "Email",
    "Location",
    "Age Range",
    "Business Name",
    "Industry",
    "Business Stage",
    "Registered",
    "Business Link",
    "Staff Count",
    "Operation Mode",
    "Offerings",
    "Revenue Range",
    "Challenges",
    "Challenges (Other)",
    "Support Needed",
    "Support Format",
    "Goals (6–12 months)",
    "Success Definition",
    "Heard About Us",
    "Comments",
    "Consent"
  ]);

  // Optional: make header bold + freeze header row
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function appendRow_(sheet, body) {
  sheet.appendRow([
    new Date(),
    body.fullName,
    body.phone,
    body.email,
    body.location,
    body.ageRange || "",
    body.businessName,
    body.industry,
    body.businessStage,
    body.registered,
    body.businessLink || "",
    body.staffCount || "",
    body.operationMode,
    body.offerings,
    body.revenueRange || "",
    body.challenges || "",
    body.challengesOther || "",
    body.support || "",
    body.supportFormat || "",
    body.goals,
    body.successDefinition || "",
    body.referralSource,
    body.comments || "",
    body.consent
  ]);
}

/** =========================
 * EMAILS
 * ========================== */

function sendWelcomeEmail_(toEmail, fullName) {
  const safeName = escapeHtml_(fullName);
  const subject = "Welcome to the TMM Community ✅";

  // Plain-text fallback
  const textBody =
    `Hi ${fullName},\n\n` +
    `Welcome to The Market Masters (TMM) Community!\n\n` +
    `Join the WhatsApp Community here:\n${CONFIG.WHATSAPP_GROUP_LINK}\n\n` +
    `If you have questions, reply to this email.\n\n` +
    `Warm regards,\n${CONFIG.FROM_NAME}\n`;

  // Table-based HTML (best compatibility for Gmail + mobile clients)
  const htmlBody = `
  <div style="margin:0;padding:0;background:#f3f5ff;">
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background:#f3f5ff;padding:24px 0;">
      <tr>
        <td align="center">
          <table role="presentation" cellpadding="0" cellspacing="0" width="600"
                 style="width:600px;max-width:600px;background:#ffffff;border-radius:16px;overflow:hidden;border:1px solid rgba(15,23,42,.10);">

            <!-- Header -->
            <tr>
              <td style="background:linear-gradient(135deg, ${CONFIG.BRAND_PRIMARY}, ${CONFIG.BRAND_INK});padding:18px 22px;">
                <table role="presentation" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td style="width:56px;" valign="middle">
                      <img src="${CONFIG.COMMUNITY_LOGO_URL}" width="46" height="46" alt="TMM Community Logo"
                           style="display:block;border-radius:12px;object-fit:contain;background:#ffffff;padding:6px;">
                    </td>
                    <td valign="middle" style="padding-left:12px;">
                      <div style="font-family:Arial,sans-serif;color:#ffffff;font-weight:800;font-size:16px;line-height:1.2;">
                        The Market Masters (TMM)
                      </div>
                      <div style="font-family:Arial,sans-serif;color:rgba(255,255,255,.86);font-size:12.5px;line-height:1.3;margin-top:2px;">
                        Community Onboarding
                      </div>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Body -->
            <tr>
              <td style="padding:26px 22px 10px 22px;">
                <div style="font-family:Arial,sans-serif;color:#0f172a;font-size:15px;line-height:1.7;">
                  <p style="margin:0 0 10px 0;">Hi <strong>${safeName}</strong>,</p>

                  <p style="margin:0 0 12px 0;">
                    Welcome to the <strong>TMM Community</strong> 🎉
                    We’re excited to have you in a network built for growth, visibility, and real opportunities.
                  </p>

                  <div style="margin:16px 0 14px 0;padding:14px 14px;border-radius:14px;background:${CONFIG.BRAND_SOFT};border:1px solid rgba(70,100,232,.20);">
                    <div style="font-weight:800;margin-bottom:6px;">Next step</div>
                    <div style="color:rgba(15,23,42,.86);">
                      Join the WhatsApp community using the button below.
                    </div>
                  </div>
                </div>
              </td>
            </tr>

            <!-- CTA -->
            <tr>
              <td align="center" style="padding:0 22px 22px 22px;">
                <table role="presentation" cellpadding="0" cellspacing="0">
                  <tr>
                    <td align="center" style="border-radius:999px;background:${CONFIG.BRAND_PRIMARY};">
                      <a href="${CONFIG.WHATSAPP_GROUP_LINK}" target="_blank" rel="noreferrer"
                         style="display:inline-block;padding:12px 18px;font-family:Arial,sans-serif;font-size:14px;font-weight:800;color:#ffffff;text-decoration:none;border-radius:999px;">
                        Join WhatsApp Community →
                      </a>
                    </td>
                  </tr>
                </table>

                <div style="font-family:Arial,sans-serif;color:rgba(15,23,42,.65);font-size:12px;line-height:1.5;margin-top:10px;">
                  If the button doesn’t work, copy and paste this link:<br>
                  <a href="${CONFIG.WHATSAPP_GROUP_LINK}" target="_blank" rel="noreferrer"
                     style="color:${CONFIG.BRAND_INK};word-break:break-all;">
                    ${CONFIG.WHATSAPP_GROUP_LINK}
                  </a>
                </div>
              </td>
            </tr>

            <!-- Divider -->
            <tr>
              <td style="padding:0 22px;">
                <div style="height:1px;background:rgba(15,23,42,.10);"></div>
              </td>
            </tr>

            <!-- Footer -->
            <tr>
              <td style="padding:16px 22px 22px 22px;">
                <div style="font-family:Arial,sans-serif;color:rgba(15,23,42,.70);font-size:12.5px;line-height:1.6;">
                  <div style="margin-bottom:10px;">
                    Need help? Reply to this email or reach us at
                    <a href="mailto:${CONFIG.ADMIN_EMAIL}" style="color:${CONFIG.BRAND_INK};text-decoration:none;font-weight:800;">
                      ${CONFIG.ADMIN_EMAIL}
                    </a>.
                  </div>

                  <div style="margin-top:10px;color:rgba(15,23,42,.55);">
                    You’re receiving this email because you registered on the TMM Community page.
                  </div>

                  <div style="margin-top:10px;">
                    <span style="font-weight:800;color:#0f172a;">${CONFIG.FROM_NAME}</span>
                    <span style="color:rgba(15,23,42,.45);"> • </span>
                    <a href="${CONFIG.WEBSITE_URL}" target="_blank" rel="noreferrer"
                       style="color:${CONFIG.BRAND_INK};text-decoration:none;">
                      ${CONFIG.WEBSITE_URL}
                    </a>
                  </div>
                </div>
              </td>
            </tr>

          </table>
        </td>
      </tr>
    </table>
  </div>`;

  MailApp.sendEmail({
    to: toEmail,
    subject: subject,
    name: CONFIG.FROM_NAME,
    body: textBody,
    htmlBody: htmlBody
  });
}

function notifyAdmin_(body) {
  const subject = "New Community Registration — " + body.fullName;

  const msg =
    "New registration:\n\n" +
    `Name: ${body.fullName}\n` +
    `Phone: ${body.phone}\n` +
    `Email: ${body.email}\n` +
    `Location: ${body.location}\n` +
    `Business: ${body.businessName}\n` +
    `Industry: ${body.industry}\n` +
    `Stage: ${body.businessStage}\n` +
    `Registered: ${body.registered}\n` +
    `Mode: ${body.operationMode}\n` +
    `Challenges: ${body.challenges || ""} ${body.challengesOther ? "(" + body.challengesOther + ")" : ""}\n` +
    `Support: ${body.support || ""}\n` +
    `Goals: ${body.goals}\n`;

  MailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, msg);
}

/** =========================
 * UTILITIES
 * ========================== */

/**
 * Parses JSON safely from the request.
 * Gives a useful error if the payload is missing or invalid.
 */
function parseJson_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("No POST body received. Ensure you're sending JSON in fetch().");
  }
  return JSON.parse(e.postData.contents);
}

/**
 * JSON response helper
 */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Escape HTML for safe injection into HTML email templates
 */
function escapeHtml_(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/** =========================
 * AUTH TEST (RUN ONCE)
 * ========================== */

/**
 * Run this ONCE in Apps Script editor to authorize:
 * - SpreadsheetApp
 * - MailApp
 *
 * Then redeploy your Web App as a new version.
 */
function testAuth() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);

  ensureHeader_(sheet);
  
  // Trigger Drive permissions
  DriveApp.getFileById(CONFIG.LOGO_FILE_ID);
  DriveApp.getFileById(CONFIG.FLYER_FILE_ID);
  
  sheet.appendRow([new Date(), "AUTH TEST", "000", "test@example.com"]);

  MailApp.sendEmail(Session.getActiveUser().getEmail(), "TMM Auth Test", "Authorization works.");
}

/** =========================
 * TODAY'S EVENT BATCH SENDER
 * ========================== */

/**
 * 📢 Run this manually to send batch event emails.
 * Sent in groups of 100 to avoid Google Apps Script limits.
 * Run once, wait until done, then repeat if more emails are needed.
 */
function sendTodayEventMailBatch() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length < 2) return;
  
  const headers = values[0];
  let sentColIdx = headers.indexOf("Event Invite Sent");
  
  // If column doesn't exist, create it
  if (sentColIdx === -1) {
    sentColIdx = headers.length;
    sheet.getRange(1, sentColIdx + 1).setValue("Event Invite Sent").setFontWeight("bold");
  }
  
  const emailIdx = headers.indexOf("Email");
  const nameIdx = headers.indexOf("Full Name");
  
  if (emailIdx === -1 || nameIdx === -1) {
    throw new Error("Could not find 'Email' or 'Full Name' columns.");
  }
  
  // Try to load auth for logo and flyer once per batch to avoid repeated slow fetches
  let logoBlob, flyerBlob;
  try {
    logoBlob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob().setName("logo_image");
    flyerBlob = DriveApp.getFileById(CONFIG.FLYER_FILE_ID).getBlob();
  } catch (err) {
    throw new Error("Could not access Drive files (Logo/Flyer). Did you run testAuth()?");
  }
  
  let sentCount = 0;
  const BATCH_LIMIT = 100;
  
  for (let r = 1; r < values.length; r++) {
    if (sentCount >= BATCH_LIMIT) {
      Logger.log("Batch limit of " + BATCH_LIMIT + " reached. Run script again for next batch.");
      break;
    }
    
    const rowNum = r + 1;
    const email = String(values[r][emailIdx] || "").trim();
    const name = String(values[r][nameIdx] || "").trim();
    const alreadySent = String(values[r][sentColIdx] || "").trim().toUpperCase() === "YES";
    
    if (!email || alreadySent) continue;
    
    try {
      sendEventMail_(email, name, logoBlob, flyerBlob);
      sheet.getRange(rowNum, sentColIdx + 1).setValue("YES");
      sentCount++;
    } catch (err) {
      Logger.log("Failed to send to " + email + ": " + err.message);
      sheet.getRange(rowNum, sentColIdx + 1).setValue("FAILED");
    }
  }
  
  Logger.log("Finished sending " + sentCount + " event invitations in this batch.");
}

function sendEventMail_(toEmail, fullName, logoBlob, flyerBlob) {
  const safeName = escapeHtml_(fullName);
  const subject = "TODAY: TMM Community Meetup is Happening! 🚀";
  
  const textBody = 
    `Hi ${fullName},\n\n` +
    `Our event "Excel For Business Mastery" is happening TODAY! We are so excited to see you by 10 AM .\n\n` +
    `Join us here: ${CONFIG.EVENT_JOIN_LINK}\n\n` +
    `The full programme flyer is attached to this email.\n\n` +
    `Warm regards,\n${CONFIG.FROM_NAME}`;
    
  const htmlBody = `
  <div style="margin:0;padding:0;background:#f3f5ff;">
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background:#f3f5ff;padding:24px 0;">
      <tr>
        <td align="center">
          <table role="presentation" cellpadding="0" cellspacing="0" width="600"
                 style="width:600px;max-width:600px;background:#ffffff;border-radius:16px;overflow:hidden;border:1px solid rgba(15,23,42,.10);">

            <!-- Header -->
            <tr>
              <td style="background:linear-gradient(135deg, ${CONFIG.BRAND_PRIMARY}, ${CONFIG.BRAND_INK});padding:18px 22px;">
                <table role="presentation" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td style="width:56px;" valign="middle">
                      <img src="cid:logo_image" width="46" height="46" alt="TMM Community Logo"
                           style="display:block;border-radius:12px;object-fit:contain;background:#ffffff;padding:6px;">
                    </td>
                    <td valign="middle" style="padding-left:12px;">
                      <div style="font-family:Arial,sans-serif;color:#ffffff;font-weight:800;font-size:16px;line-height:1.2;">
                        The Market Masters (TMM)
                      </div>
                      <div style="font-family:Arial,sans-serif;color:rgba(255,255,255,.86);font-size:12.5px;line-height:1.3;margin-top:2px;">
                        Event Today!
                      </div>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Body -->
            <tr>
              <td style="padding:26px 22px 10px 22px;">
                <div style="font-family:Arial,sans-serif;color:#0f172a;font-size:15px;line-height:1.7;">
                  <p style="margin:0 0 10px 0;">Hi <strong>${safeName}</strong>,</p>

                  <p style="margin:0 0 12px 0;">
                    Our special event is happening <strong>TODAY</strong>! We can't wait to see you there. 
                    Please find the programme flyer attached to this email.
                  </p>

                  <div style="margin:16px 0 14px 0;padding:14px 14px;border-radius:14px;background:${CONFIG.BRAND_SOFT};border:1px solid rgba(70,100,232,.20);">
                    <div style="font-weight:800;margin-bottom:6px;">How to join</div>
                    <div style="color:rgba(15,23,42,.86);">
                      Click the button below to enter the Google Meet instantly.
                    </div>
                  </div>
                </div>
              </td>
            </tr>

            <!-- CTA -->
            <tr>
              <td align="center" style="padding:0 22px 22px 22px;">
                <table role="presentation" cellpadding="0" cellspacing="0">
                  <tr>
                    <td align="center" style="border-radius:999px;background:${CONFIG.BRAND_PRIMARY};">
                      <a href="${CONFIG.EVENT_JOIN_LINK}" target="_blank" rel="noreferrer"
                         style="display:inline-block;padding:12px 18px;font-family:Arial,sans-serif;font-size:14px;font-weight:800;color:#ffffff;text-decoration:none;border-radius:999px;">
                        Join The Event →
                      </a>
                    </td>
                  </tr>
                </table>

                <div style="font-family:Arial,sans-serif;color:rgba(15,23,42,.65);font-size:12px;line-height:1.5;margin-top:10px;">
                  If the button doesn’t work, join via this link:<br>
                  <a href="${CONFIG.EVENT_JOIN_LINK}" target="_blank" rel="noreferrer"
                     style="color:${CONFIG.BRAND_INK};word-break:break-all;">
                    ${CONFIG.EVENT_JOIN_LINK}
                  </a>
                </div>
              </td>
            </tr>
            <!-- Divider -->
            <tr>
              <td style="padding:0 22px;">
                <div style="height:1px;background:rgba(15,23,42,.10);"></div>
              </td>
            </tr>

            <!-- Footer -->
            <tr>
              <td style="padding:16px 22px 22px 22px;">
                <div style="font-family:Arial,sans-serif;color:rgba(15,23,42,.70);font-size:12.5px;line-height:1.6;">
                  <div style="margin-bottom:10px;">
                    Need help? Reply to this email or reach us at
                    <a href="mailto:${CONFIG.ADMIN_EMAIL}" style="color:${CONFIG.BRAND_INK};text-decoration:none;font-weight:800;">
                      ${CONFIG.ADMIN_EMAIL}
                    </a>.
                  </div>
                  <div style="margin-top:10px;">
                    <span style="font-weight:800;color:#0f172a;">${CONFIG.FROM_NAME}</span>
                    <span style="color:rgba(15,23,42,.45);"> • </span>
                    <a href="${CONFIG.WEBSITE_URL}" target="_blank" rel="noreferrer"
                       style="color:${CONFIG.BRAND_INK};text-decoration:none;">
                      ${CONFIG.WEBSITE_URL}
                    </a>
                  </div>
                </div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </div>`;

  MailApp.sendEmail({
    to: toEmail,
    subject: subject,
    name: CONFIG.FROM_NAME,
    body: textBody,
    htmlBody: htmlBody,
    inlineImages: { logo_image: logoBlob },
    attachments: [flyerBlob]
  });
}
