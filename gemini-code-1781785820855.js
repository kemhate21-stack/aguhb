// ============================================================
//  KlexDigirents – Digicam Rental Automation [CYBER-Y2K ED.]
//  Google Apps Script for Google Forms + Google Sheets
// ============================================================

// ─────────────────────────────────────────────
//  CONFIGURATION
// ─────────────────────────────────────────────
const CONFIG = {
  OWNER_EMAIL: "Kemhate21@gmail.com",          // Your Gmail address (owner)
  OWNER_NAME:  "Klex",                           // Your display name
  BUSINESS_NAME: "KlexDigirents",
  GCASH_NUMBER: "09516427351",

  // How many hours before rent START to send the reminder
  REMINDER_HOURS_BEFORE: 24,

  // Column indices (1-based) matching your sheet headers
  COL: {
    TIMESTAMP:        1,
    EMAIL_ADDRESS:    2,   // First auto-collected Email Address
    NAME_DIGICAM:     3,   // "Enter your first name and the digicam..."
    EMAIL_MANUAL:     4,   // Second Email Address column
    FIRST_NAME:       5,   // First name
    LAST_NAME:        6,   // Last name
    ADDRESS:          7,
    CONTACT:          8,
    DIGICAM:          9,
    RENT_START:       10,
    RENT_END:         11,
    RENT_TIME:        12,
    MEETUP:           13,
    ID_FRONT:         14,
    ID_BACK:          15,
    PAYMENT_AMOUNT:   16,  // Amount paid
    PAYMENT_SS:       17,  // Screenshot of payment
    EXTRA_BATTERY:    18,
    AGREEMENT:        19,
    // Internal tracking columns
    BOOKING_ID:       20,
    STATUS:           21,
    REMINDER_SENT:    22,
  }
};

// ─────────────────────────────────────────────
//  ENTRY POINT – triggered on form submit
// ─────────────────────────────────────────────
function onFormSubmit(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const row   = e.range.getRow();
    const data  = getRowData(sheet, row);

    // Generate a unique booking ID
    const bookingId = generateBookingId(data);
    sheet.getRange(row, CONFIG.COL.BOOKING_ID).setValue(bookingId);
    sheet.getRange(row, CONFIG.COL.STATUS).setValue("CONFIRMED");
    sheet.getRange(row, CONFIG.COL.REMINDER_SENT).setValue("NO");

    // 1 · Send confirmation to renter
    sendRenterConfirmation(data, bookingId);

    // 2 · Notify owner (Klex)
    sendOwnerNotification(data, bookingId);

    // 3 · Schedule reminder
    scheduleReminder(data, bookingId, row);

    // 4 · Create calendar event & invite renter
    createCalendarEvent(data, bookingId);

    Logger.log(`[KlexDigirents] Booking ${bookingId} processed for ${data.fullName}`);
  } catch (err) {
    Logger.log("ERROR in onFormSubmit: " + err.message);
    MailApp.sendEmail(
      CONFIG.OWNER_EMAIL,
      "Error processing new booking",
      "An error occurred while processing a new form submission.\n\nError: " + err.message +
      "\n\nPlease check the Google Sheet manually."
    );
  }
}

// ─────────────────────────────────────────────
//  READ ROW DATA
// ─────────────────────────────────────────────
function getRowData(sheet, row) {
  const C = CONFIG.COL;
  const get = (col) => sheet.getRange(row, col).getValue();

  let emailRaw = String(get(C.EMAIL_MANUAL) || get(C.EMAIL_ADDRESS)).trim();
  
  if (emailRaw.indexOf("@") === -1) {
    emailRaw = String(get(C.EMAIL_ADDRESS)).trim();
  }

  const fName = String(get(C.FIRST_NAME)).trim();
  const lName = String(get(C.LAST_NAME)).trim();
  const combinedName = (fName + " " + lName).trim() || "Valued Customer";

  const rentStart   = new Date(get(C.RENT_START));
  const rentEnd     = new Date(get(C.RENT_END));
  const rentTime    = String(get(C.RENT_TIME)).trim();
  const extraBat    = String(get(C.EXTRA_BATTERY)).trim().toUpperCase();

  return {
    timestamp:     get(C.TIMESTAMP),
    email:         emailRaw,
    nameDigicam:   String(get(C.NAME_DIGICAM)).trim(),
    fullName:      combinedName,
    address:       String(get(C.ADDRESS)).trim(),
    contact:       String(get(C.CONTACT)).trim(),
    digicam:       String(get(C.DIGICAM)).trim(),
    rentStart:     rentStart,
    rentEnd:       rentEnd,
    rentTime:      rentTime,
    meetup:        String(get(C.MEETUP)).trim(),
    paymentAmount: get(C.PAYMENT_AMOUNT),
    extraBattery:  extraBat === "YES" || extraBat === "TRUE" || extraBat === "1",
    agreement:     String(get(C.AGREEMENT)).trim(),
    row:           row,
  };
}

// ─────────────────────────────────────────────
//  1 · RENTER CONFIRMATION EMAIL [CYBER-Y2K]
// ─────────────────────────────────────────────
function sendRenterConfirmation(data, bookingId) {
  if (!data.email || data.email.indexOf("@") === -1) return;

  const subject = `Booking Confirmed // ${CONFIG.BUSINESS_NAME} [Ref: ${bookingId}]`;
  const html    = buildRenterEmailHtml(data, bookingId);

  GmailApp.sendEmail(data.email, subject, stripHtml(html), {
    name:     CONFIG.BUSINESS_NAME,
    replyTo:  CONFIG.OWNER_EMAIL,
    htmlBody: html,
  });
}

function buildRenterEmailHtml(data, bookingId) {
  const depositNote  = "PHP 150 security deposit (refunded on timely return in good condition)";
  const lateNote     = "Late return - half-day fee of PHP 100 applies";
  const extraBatNote = data.extraBattery ? "Yes - requested" : "No - not requested";

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  body{font-family:'Courier New', Courier, monospace, Arial, sans-serif;background:#0d0e15;margin:0;padding:0;color:#e0e0e6;}
  .wrapper{max-width:600px;margin:30px auto;background:#161722;border:2px solid #00f0ff;border-radius:0px;overflow:hidden;box-shadow:0 0 15px rgba(0,240,255,0.3);}
  .header{background:linear-gradient(135deg,#0a0b10,#12131a);color:#fff;padding:35px 24px;text-align:center;border-bottom:2px solid #ff007f;}
  .header h1{margin:0;font-size:28px;letter-spacing:3px;color:#00f0ff;text-transform:uppercase;text-shadow:0 0 8px rgba(0,240,255,0.5);}
  .header p{margin:6px 0 0;opacity:.8;font-size:12px;letter-spacing:2px;color:#ff007f;text-transform:uppercase;}
  .badge{display:inline-block;background:transparent;color:#00f0ff;border:1px solid #00f0ff;padding:5px 16px;font-size:11px;font-weight:600;margin-top:14px;letter-spacing:1.5px;}
  .body{padding:28px 24px;}
  .greeting{font-size:16px;color:#ffffff;margin-bottom:20px;border-bottom:1px dashed #34364c;padding-bottom:10px;}
  .card{background:#1c1d2a;border-left:4px solid #ff007f;padding:16px 18px;margin-bottom:18px;}
  .card h3{margin:0 0 12px;font-size:13px;text-transform:uppercase;letter-spacing:1px;color:#ff007f;}
  table{width:100%;border-collapse:collapse;font-size:13px;}
  td{padding:7px 4px;vertical-align:top;color:#c4c5d4;}
  td:first-child{font-weight:600;color:#00f0ff;width:42%;white-space:nowrap;text-transform:uppercase;font-size:12px;}
  .highlight{background:#112233;border:1px solid #00f0ff;padding:14px 18px;margin-bottom:18px;font-size:13px;color:#e0e0e6;}
  .rules{font-size:13px;color:#a5a7c2;line-height:1.7;}
  .footer{background:#0a0b10;color:#656785;padding:20px 24px;text-align:center;font-size:11px;border-top:1px solid #232435;letter-spacing:1px;}
  a{color:#00f0ff;text-decoration:none;}
</style>
</head>
<body>
<div class="wrapper">
  <div class="header">
    <h1>${CONFIG.BUSINESS_NAME}</h1>
    <p>// SYSTEM OVERRIDE: ACCESS GRANTED</p>
    <div class="badge">[ BOOKING_CONFIRMED ]</div>
  </div>
  <div class="body">
    <p class="greeting">INITIALIZING DATA FOR: <strong>${data.fullName}</strong>...</p>

    <div class="card">
      <h3>GEN-Y2K RENTAL LOGS</h3>
      <table>
        <tr><td>SYSTEM_REF</td><td><strong>${bookingId}</strong></td></tr>
        <tr><td>HARDWARE</td><td>${data.digicam}</td></tr>
        <tr><td>AUX_BATTERY</td><td>${extraBatNote}</td></tr>
        <tr><td>LAUNCH_DATE</td><td>${formatDate(data.rentStart)}</td></tr>
        <tr><td>RETURN_DATE</td><td>${formatDate(data.rentEnd)}</td></tr>
        <tr><td>TIME_STAMP</td><td>${data.rentTime}</td></tr>
        <tr><td>NODE_LOCATION</td><td>${data.meetup}</td></tr>
        <tr><td>CREDITS_PAID</td><td>PHP ${data.paymentAmount}</td></tr>
      </table>
    </div>

    <div class="highlight">
      <strong style="color:#ff007f;">[SEC_DEPOSIT]:</strong> ${depositNote}<br>
      <strong style="color:#ff007f;">[LATE_RETURN]:</strong> ${lateNote}
    </div>

    <div class="card">
      <h3>HARDWARE INCLUSIONS</h3>
      <div class="rules">
        :: Kodak Easyshare Digital Camera &nbsp;:: Battery &nbsp;:: Memory Card<br>
        :: Card Reader &nbsp;:: Cyber Pouch &nbsp;:: Strap &nbsp;:: Charger
        ${data.extraBattery ? "<br>:: <strong>Extra Battery Pack [ENGAGED]</strong>" : ""}
      </div>
    </div>

    <div class="card">
      <h3>CORE PROTOCOLS</h3>
      <div class="rules">
        * Return all gear in its original network configuration/condition<br>
        * Return time matrix is exact 24-hour cycles based on your rental time<br>
        * No unauthorized network sub-renting or hardware modding allowed
      </div>
    </div>

    <p style="font-size:12px;color:#8587a3;text-align:center;margin-top:25px;">
      Inquiries? Ping terminal at 
      <a href="mailto:${CONFIG.OWNER_EMAIL}">${CONFIG.OWNER_EMAIL}</a> <br>
      or secure line: <strong>${CONFIG.GCASH_NUMBER}</strong>
    </p>
  </div>
  <div class="footer">
    &lt; KLEXDIGIRENTS // CORE_v2.0 // SYSTEM_YEAR_${new Date().getFullYear()} &gt;
  </div>
</div>
</body>
</html>`;
}

// ─────────────────────────────────────────────
//  2 · OWNER NOTIFICATION EMAIL [CYBER-Y2K]
// ─────────────────────────────────────────────
function sendOwnerNotification(data, bookingId) {
  const subject = `[NEW_INBOUND] ${data.digicam} // ${data.fullName} // ${bookingId}`;
  const html    = buildOwnerEmailHtml(data, bookingId);

  GmailApp.sendEmail(CONFIG.OWNER_EMAIL, subject, stripHtml(html), {
    name:     CONFIG.BUSINESS_NAME + " HQ",
    htmlBody: html,
  });
}

function buildOwnerEmailHtml(data, bookingId) {
  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body{font-family:'Courier New',monospace;background:#0d0e15;margin:0;padding:0;color:#e0e0e6;}
  .wrapper{max-width:600px;margin:20px auto;background:#161722;border:2px solid #ff007f;}
  .header{background:#0a0b10;color:#00f0ff;padding:20px 24px;border-bottom:1px solid #ff007f;}
  .header h2{margin:0;font-size:18px;letter-spacing:1px;}
  .header span{font-size:11px;color:#ff007f;}
  .body{padding:24px;}
  table{width:100%;border-collapse:collapse;font-size:13px;}
  th{background:#1c1d2a;text-align:left;padding:8px 10px;font-size:11px;color:#00f0ff;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid #34364c;}
  td{padding:9px 10px;border-bottom:1px solid #1c1d2a;vertical-align:top;}
  td:first-child{font-weight:600;color:#ff007f;width:38%;font-size:12px;}
  .status{display:inline-block;background:transparent;color:#00f0ff;border:1px solid #00f0ff;padding:2px 8px;font-size:11px;}
  .alert{background:#112233;border:1px solid #00f0ff;padding:12px 16px;margin-top:16px;font-size:12px;line-height:1.6;}
</style>
</head>
<body>
<div class="wrapper">
  <div class="header">
    <h2>[NEW_TRANSACTION_LOGGED]</h2>
    <span>SYS_TIME: ${new Date().toLocaleString("en-PH")}</span>
  </div>
  <div class="body">
    <table>
      <tr><th colspan="2">// SYSTEM METRICS</th></tr>
      <tr><td>REF_ID</td><td><strong>${bookingId}</strong> <span class="status">CONFIRMED</span></td></tr>
      <tr><td>HARDWARE</td><td>${data.digicam}</td></tr>
      <tr><td>AUX_BATT</td><td>${data.extraBattery ? "YES" : "NO"}</td></tr>
      <tr><th colspan="2">// RENTER MANIFEST</th></tr>
      <tr><td>CLIENT_NAME</td><td>${data.fullName}</td></tr>
      <tr><td>EMAIL_NODE</td><td>${data.email}</td></tr>
      <tr><td>COMMS_LINE</td><td>${data.contact}</td></tr>
      <tr><td>GRID_ADDR</td><td>${data.address}</td></tr>
      <tr><th colspan="2">// TIME GRID</th></tr>
      <tr><td>START_CYCLE</td><td>${formatDate(data.rentStart)}</td></tr>
      <tr><td>END_CYCLE</td><td>${formatDate(data.rentEnd)}</td></tr>
      <tr><td>MATRIX_TIME</td><td>${data.rentTime}</td></tr>
      <tr><td>RENDEZVOUS</td><td>${data.meetup}</td></tr>
      <tr><th colspan="2">// FINANCIALS</th></tr>
      <tr><td>ESCROW_AMT</td><td>PHP ${data.paymentAmount}</td></tr>
      <tr><td>PROT_AGREE</td><td>${data.agreement ? "SIGNED" : "VOID"}</td></tr>
    </table>
    <div class="alert">
      <strong>ACTION REQUIRED:</strong> Stage the hardware <strong>${data.digicam}</strong> for dispatch at node <strong>${data.meetup}</strong> on cycle <strong>${formatDate(data.rentStart)} @ ${data.rentTime}</strong>.
    </div>
  </div>
</div>
</body>
</html>`;
}

// ─────────────────────────────────────────────
//  3 · REMINDER SCHEDULER
// ─────────────────────────────────────────────
function scheduleReminder(data, bookingId, row) {
  if (!data.email || data.email.indexOf("@") === -1) return;

  const reminderTime = new Date(data.rentStart);
  reminderTime.setHours(reminderTime.getHours() - CONFIG.REMINDER_HOURS_BEFORE);

  const reminder = {
    bookingId:    bookingId,
    renterEmail:  data.email,
    renterName:   data.fullName,
    digicam:      data.digicam,
    rentStart:    data.rentStart.toISOString(),
    rentTime:     data.rentTime,
    meetup:       data.meetup,
    triggerTime:  reminderTime.toISOString(),
    row:          row,
    sent:         false,
  };

  const props = PropertiesService.getScriptProperties();
  const key   = "reminder_" + bookingId;
  props.setProperty(key, JSON.stringify(reminder));
  Logger.log(`Reminder scheduled for ${reminderTime} (booking ${bookingId})`);
}

function processReminders() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  const now   = new Date();

  Object.keys(allProps).forEach(key => {
    if (!key.startsWith("reminder_")) return;

    let reminder;
    try { reminder = JSON.parse(allProps[key]); } catch (e) { return; }
    if (reminder.sent) return;

    const triggerTime = new Date(reminder.triggerTime);
    if (now >= triggerTime) {
      sendReminderEmail(reminder);
      sendOwnerReminderNotice(reminder);

      reminder.sent = true;
      props.setProperty(key, JSON.stringify(reminder));

      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        sheet.getRange(reminder.row, CONFIG.COL.REMINDER_SENT).setValue("YES");
      } catch (e) { /* sheet update is best-effort */ }

      Logger.log(`Reminder sent for booking ${reminder.bookingId}`);
    }
  });
}

function sendReminderEmail(reminder) {
  if (!reminder.renterEmail || reminder.renterEmail.indexOf("@") === -1) return;

  const rentDate  = new Date(reminder.rentStart);
  const subject   = `[ALERT] Cyber-Rental Liftoff Tomorrow! // Ref: ${reminder.bookingId}`;
  const html      = `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body{font-family:'Courier New',monospace;background:#0d0e15;margin:0;padding:0;color:#e0e0e6;}
  .wrapper{max-width:560px;margin:30px auto;background:#161722;border:2px solid #00f0ff;box-shadow:0 0 10px rgba(0,240,255,0.2);}
  .header{background:linear-gradient(135deg,#ff007f,#0a0b10);color:#fff;padding:24px;text-align:center;}
  .header h2{margin:0;font-size:20px;letter-spacing:1px;}
  .body{padding:24px;}
  .detail{background:#1c1d2a;border-left:4px solid #00f0ff;padding:16px;margin:16px 0;font-size:13px;line-height:1.8;}
  .detail strong{color:#ff007f;}
  .footer{text-align:center;padding:14px;font-size:11px;color:#656785;}
</style>
</head>
<body>
<div class="wrapper">
  <div class="header">
    <h2>// AUTO_REMINDER_PULSE</h2>
    <p style="margin:4px 0 0;font-size:11px;color:#00f0ff;letter-spacing:1px;">${CONFIG.BUSINESS_NAME.toUpperCase()}</p>
  </div>
  <div class="body">
    <p>Attention <strong>${reminder.renterName}</strong>,</p>
    <p>Your hardware rental schedule is entering immediate orbit in <strong>${CONFIG.REMINDER_HOURS_BEFORE} hours</strong>. Sync parameters below:</p>
    <div class="detail">
      SYSTEM_UNIT: ${reminder.digicam}<br>
      CYCLE_DATE: ${formatDate(rentDate)}<br>
      MATRIX_TIME: ${reminder.rentTime}<br>
      NODE_MEETUP: ${reminder.meetup}<br>
      TRACKING_ID: ${reminder.bookingId}
    </div>
    <p style="font-size:12px;color:#a5a7c2;">
      Ready your <strong>PHP 150 backup escrow (security deposit)</strong> in physical credits.<br>
      System issues? Ping terminal at <a href="mailto:${CONFIG.OWNER_EMAIL}" style="color:#00f0ff;">${CONFIG.OWNER_EMAIL}</a>
    </p>
  </div>
  <div class="footer">&lt; SYSTEM_WARN_DISPATCH_AUTOMATION &gt;</div>
</div>
</body>
</html>`;

  GmailApp.sendEmail(reminder.renterEmail, subject, stripHtml(html), {
    name:     CONFIG.BUSINESS_NAME,
    htmlBody: html,
    replyTo:  CONFIG.OWNER_EMAIL,
  });
}

function sendOwnerReminderNotice(reminder) {
  const rentDate = new Date(reminder.rentStart);
  const subject  = `[DISPATCH_WARN] ${reminder.digicam} -> ${reminder.renterName} tomorrow!`;
  const body     =
    `// UPCOMING DISPATCH CYCLE – ${CONFIG.BUSINESS_NAME.toUpperCase()}\n\n` +
    `TRACKING_ID : ${reminder.bookingId}\n` +
    `CLIENT_NODE : ${reminder.renterName}\n` +
    `HARDWARE    : ${reminder.digicam}\n` +
    `CYCLE_DATE  : ${formatDate(rentDate)}\n` +
    `MATRIX_TIME : ${reminder.rentTime}\n` +
    `NODE_MEETUP : ${reminder.meetup}\n\n` +
    `Pre-stage and clean system hardware immediately.`;

  MailApp.sendEmail(CONFIG.OWNER_EMAIL, subject, body);
}

// ─────────────────────────────────────────────
//  4 · GOOGLE CALENDAR EVENT + RENTER INVITE
// ─────────────────────────────────────────────
function createCalendarEvent(data, bookingId) {
  const calendar = CalendarApp.getDefaultCalendar();

  const startDt = combineDateAndTime(data.rentStart, data.rentTime);
  const endDt   = new Date(startDt.getTime() + 60 * 60 * 1000); 

  const title   = `[${CONFIG.BUSINESS_NAME}] ${data.digicam} -> ${data.fullName}`;
  const desc    = [
    `Booking Ref : ${bookingId}`,
    `Camera      : ${data.digicam}`,
    `Extra Batt  : ${data.extraBattery ? "Yes" : "No"}`,
    ``,
    `Renter      : ${data.fullName}`,
    `Email       : ${data.email}`,
    `Contact     : ${data.contact}`,
    ``,
    `Rent Period : ${formatDate(data.rentStart)} -> ${formatDate(data.rentEnd)}`,
    `Time        : ${data.rentTime}`,
    `Meet-up     : ${data.meetup}`,
    ``,
    `Payment     : PHP ${data.paymentAmount}`,
    `Deposit     : PHP 150 (to collect on meet-up)`,
  ].join("\n");

  const hasValidEmail = data.email && data.email.indexOf("@") !== -1;

  const eventOptions = {
    description: desc,
    location:    data.meetup,
  };
  if (hasValidEmail) {
    eventOptions.guests = data.email;
    eventOptions.sendInvites = true;
  }

  calendar.createEvent(title, startDt, endDt, eventOptions);

  // Return event
  const returnStart = combineDateAndTime(data.rentEnd, data.rentTime);
  const returnEnd   = new Date(returnStart.getTime() + 60 * 60 * 1000);

  const returnOptions = {
    description: `Return of ${data.digicam}\nRef: ${bookingId}\nRenter: ${data.fullName} (${data.contact})\nMeet-up: ${data.meetup}`,
    location:    data.meetup,
  };
  if (hasValidEmail) {
    returnOptions.guests = data.email;
    returnOptions.sendInvites = true;
  }

  calendar.createEvent(
    `[${CONFIG.BUSINESS_NAME}] RETURN - ${data.digicam} from ${data.fullName}`,
    returnStart,
    returnEnd,
    returnOptions
  );

  Logger.log(`Calendar events created for booking ${bookingId}`);
}

// ─────────────────────────────────────────────
//  UTILITY HELPERS
// ─────────────────────────────────────────────
function combineDateAndTime(dateObj, timeStr) {
  const d = new Date(dateObj);
  d.setSeconds(0, 0);

  if (!timeStr) { d.setHours(10, 0); return d; }

  const upperTime = timeStr.toUpperCase().trim();
  const amPmMatch = upperTime.match(/(\d+):?(\d{0,2})\s*(AM|PM)/);
  const h24Match  = upperTime.match(/^(\d{1,2}):(\d{2})$/);

  if (amPmMatch) {
    let h  = parseInt(amPmMatch[1]);
    const m = parseInt(amPmMatch[2] || "0");
    const period = amPmMatch[3];
    if (period === "PM" && h !== 12) h += 12;
    if (period === "AM" && h === 12) h = 0;
    d.setHours(h, m);
  } else if (h24Match) {
    d.setHours(parseInt(h24Match[1]), parseInt(h24Match[2]));
  } else {
    d.setHours(10, 0);
  }
  return d;
}

function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date)) return "N/A";
  return date.toLocaleDateString("en-PH", {
    weekday: "long", year: "numeric", month: "long", day: "numeric"
  });
}

function stripHtml(html) {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]+>/g, "")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function generateBookingId(data) {
  const prefix = "KDR";
  const date   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const rand   = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${prefix}-${date}-${rand}`;
}

// ─────────────────────────────────────────────
//  TRIGGER SETUP
// ─────────────────────────────────────────────
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("processReminders")
    .timeBased()
    .everyMinutes(30)
    .create();

  Logger.log("Triggers set up successfully!");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("KlexDigirents")
    .addItem("Setup Triggers", "setupTrigger")
    .addItem("Run Reminders Now", "processReminders")
    .addToUi();
}