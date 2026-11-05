/**
 * Paste this into Extensions > Apps Script on the Pledges sheet.
 * Deploy > New deployment > Web app.
 * Execute as: Me
 * Who has access: Anyone
 * Copy the web app URL into js/config.js as scriptUrl.
 * Project Settings, Script properties: TURNSTILE_SECRET = the Turnstile secret key.
 *
 * Verification mail is read from the Gmail account that owns this sheet.
 * Run installPledgeMailTrigger once. It checks the Gmail label "Pledges" every hour.
 * A row becomes Verified only when the pledge ID, the code, and the From address all match.
 *
 * The sheet stays private. This script returns totals and public names only.
 */

var HEADERS = [
  "Id",
  "First Name",
  "Last Name",
  "Age",
  "Email",
  "Country",
  "Amount you would commit",
  "Would you pay in one payment or in installments?",
  "Do you want to show your pledge publicly?",
  "Aggreement",
  "Display Name",
  "Time",
  "Country Code",
  "Phone",
  "Status",
  "Verification code"
];

function doGet(e) {
  try {
    return respond_(e, { ok: true, stats: readStats_() });
  } catch (err) {
    return respond_(e, { ok: false, error: "The pledge sheet is not ready." });
  }
}

function doPost(e) {
  try {
    var raw = e.postData && e.postData.contents ? e.postData.contents : "{}";
    var data = JSON.parse(raw);
    if (data.company) return respond_(e, { ok: false, error: "Rejected." });
    var turnstileError = verifyTurnstile_(data.turnstileToken);
    if (turnstileError) return respond_(e, { ok: false, error: turnstileError });
    var clean = validate_(data);
    if (clean.error) return respond_(e, { ok: false, error: clean.error });
    var saved = append_(clean.value);
    return respond_(e, { ok: true, id: saved.id, code: saved.code, stats: readStats_() });
  } catch (err) {
    return respond_(e, { ok: false, error: "Could not save the pledge." });
  }
}

function respond_(e, obj) {
  var text = JSON.stringify(obj);
  var callback = e && e.parameter ? e.parameter.callback : "";
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + "(" + text + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyTurnstile_(token) {
  var secret = PropertiesService.getScriptProperties().getProperty("TURNSTILE_SECRET");
  if (!secret) return "The pledge check is not ready.";
  if (!token) return "Confirm you are a person and try again.";
  var response = UrlFetchApp.fetch("https://challenges.cloudflare.com/turnstile/v0/siteverify", {
    method: "post",
    payload: { secret: secret, response: String(token) },
    muteHttpExceptions: true
  });
  var body = {};
  try {
    body = JSON.parse(response.getContentText() || "{}");
  } catch (err) {
    return "Confirm you are a person and try again.";
  }
  if (!body.success) return "Confirm you are a person and try again.";
  return "";
}

function validate_(data) {
  var first = text_(data.firstName);
  var last = text_(data.lastName);
  var email = text_(data.email).toLowerCase();
  var phone = text_(data.phone);
  var countryCode = text_(data.countryCode);
  var country = text_(data.country);
  var displayName = text_(data.displayName);
  var payment = text_(data.payment);
  var age = Number(data.age);
  var amount = Number(data.amount);
  var isPublic = data.isPublic === true || data.isPublic === "Yes";
  var agreed = data.agreed === true || data.agreed === "Yes";

  if (!first || !last) return { error: "Enter your first and last name." };
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { error: "Enter a valid email address." };
  if (!/^\+\d{1,3}$/.test(countryCode)) return { error: "Choose a country code." };
  if (!phone) return { error: "Enter a phone number." };
  if (phone.indexOf("-") !== -1) return { error: "Dashes are not allowed. Use digits only." };
  if (!/^\d+$/.test(phone)) return { error: "Use digits only." };
  if (phone.length < 6 || phone.length > 12) return { error: "Enter 6 to 12 digits." };
  if (!isFinite(age) || age < 1 || age > 120) return { error: "Enter your age." };
  if (!country) return { error: "Choose a country." };
  if (!isFinite(amount) || amount < 1 || amount > 50000) return { error: "Enter an amount from $1 to $50,000." };
  if (payment !== "One payment" && payment !== "Installments") return { error: "Choose how you would pay." };
  if (!agreed) return { error: "The agreement is required." };

  return {
    value: {
      firstName: first,
      lastName: last,
      age: Math.round(age),
      email: email,
      phone: phone,
      countryCode: countryCode,
      country: country,
      amount: Math.round(amount * 100) / 100,
      payment: payment,
      isPublic: isPublic ? "Yes" : "No",
      agreed: "Yes",
      displayName: displayName
    }
  };
}

function append_(row) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var id = Utilities.getUuid();
    var code = secretCode_();
    var time = Utilities.formatDate(new Date(), "Etc/UTC", "yyyy-MM-dd HH:mm:ss 'UTC'");
    var map = headerMap_(sheet);
    var line = new Array(Math.max(sheet.getLastColumn(), HEADERS.length));
    line[map.id] = id;
    line[map.first] = row.firstName;
    line[map.last] = row.lastName;
    line[map.age] = row.age;
    line[map.email] = row.email;
    line[map.country] = row.country;
    line[map.amount] = row.amount;
    line[map.payment] = row.payment;
    line[map.pub] = row.isPublic;
    line[map.agree] = row.agreed;
    line[map.display] = row.displayName;
    line[map.time] = time;
    line[map.code] = row.countryCode;
    line[map.phone] = row.phone;
    line[map.status] = "Pending";
    line[map.verify] = code;
    sheet.appendRow(line);
    return { id: id, code: code };
  } finally {
    lock.releaseLock();
  }
}

function readStats_() {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var map = headerMap_(sheet);
  var total = 0;
  var supporters = 0;
  var countries = {};
  var publicRows = [];

  for (var i = 1; i < values.length; i++) {
    if (!isVerified_(values[i][map.status])) continue;
    var amount = Number(values[i][map.amount]);
    if (!isFinite(amount) || amount <= 0) continue;
    supporters++;
    total += amount;
    var country = text_(values[i][map.country]);
    if (country) countries[country] = (countries[country] || 0) + 1;
    if (isYes_(values[i][map.pub])) {
      var name = text_(values[i][map.display]);
      if (!name) continue;
      publicRows.push({
        name: name,
        country: country,
        amount: amount,
        time: text_(values[i][map.time])
      });
    }
  }

  var recent = publicRows.slice().reverse().slice(0, 12);
  var largest = publicRows.slice().sort(function (a, b) {
    return b.amount - a.amount;
  }).slice(0, 8);

  return {
    pledged: Math.round(total * 100) / 100,
    supporters: supporters,
    countries: countries,
    recent: recent,
    largest: largest
  };
}

function getSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var first = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var blank = first.join("") === "";
  if (blank) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function headerMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || HEADERS.length).getValues()[0];
  var index = {};
  for (var i = 0; i < headers.length; i++) index[norm_(headers[i])] = i;
  function col(names) {
    for (var n = 0; n < names.length; n++) {
      if (index[names[n]] !== undefined) return index[names[n]];
    }
    throw new Error("Missing column");
  }
  return {
    id: col(["id"]),
    first: col(["first name"]),
    last: col(["last name"]),
    age: col(["age"]),
    email: col(["email"]),
    country: col(["country"]),
    amount: col(["amount you would commit"]),
    payment: col(["would you pay in one payment or in installments?"]),
    pub: col(["do you want to show your pledge publicly?"]),
    agree: col(["aggreement", "agreement"]),
    display: col(["display name"]),
    time: col(["time"]),
    code: col(["country code"]),
    phone: col(["phone", "phone number"]),
    status: col(["status"]),
    verify: col(["verification code"])
  };
}

function isYes_(value) {
  return /^(yes|true|y|1)$/i.test(text_(value));
}

function isVerified_(value) {
  return text_(value).toLowerCase() === "verified";
}

function secretCode_() {
  var alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var out = "";
  for (var i = 0; i < 8; i++) {
    out += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  }
  return out;
}

function text_(value) {
  return String(value == null ? "" : value).trim();
}

function norm_(value) {
  return text_(value).toLowerCase();
}

var PLEDGE_MAIL_QUERY = 'label:pledges -label:"pledge checked" newer_than:14d';
var OWN_MAILBOXES = {
  "hello@bringunitedhome.org": true,
  "pledge@bringunitedhome.org": true,
  "bringunitedhome@gmail.com": true
};

function installPledgeMailTrigger() {
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === "checkPledgeMail") {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  ScriptApp.newTrigger("checkPledgeMail").timeBased().everyHours(1).create();
}

function checkPledgeMail() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log("Pledge mail check skipped because the sheet is busy.");
    return;
  }
  try {
    var sheet = getSheet_();
    var map = headerMap_(sheet);
    var values = sheet.getDataRange().getValues();
    var pending = {};
    for (var i = 1; i < values.length; i++) {
      if (text_(values[i][map.status]).toLowerCase() !== "pending") continue;
      var id = text_(values[i][map.id]).toLowerCase();
      if (!id) continue;
      pending[id] = {
        row: i + 1,
        email: text_(values[i][map.email]).toLowerCase(),
        code: text_(values[i][map.verify]).toLowerCase()
      };
    }

    var checked = gmailLabel_("Pledge checked");
    var threadsRead = 0;
    var verifiedCount = 0;
    for (var page = 0; page < 4; page++) {
      var threads = GmailApp.search(PLEDGE_MAIL_QUERY, 0, 25);
      if (!threads.length) break;
      for (var t = 0; t < threads.length; t++) {
        var messages = threads[t].getMessages();
        for (var m = 0; m < messages.length; m++) {
          if (verifyFromMessage_(messages[m], pending, sheet, map)) verifiedCount++;
        }
        threads[t].addLabel(checked);
        threadsRead++;
      }
    }
    Logger.log("Threads read: " + threadsRead + ". Newly verified: " + verifiedCount);
  } finally {
    lock.releaseLock();
  }
}

function verifyFromMessage_(message, pending, sheet, map) {
  var id = pledgeIdFrom_(message.getSubject()) || pledgeIdFrom_(message.getPlainBody());
  if (!id) return false;
  var row = pending[id];
  if (!row) {
    Logger.log("No pending row for " + id);
    return false;
  }
  var code = codeFrom_(message.getPlainBody());
  if (!code || code !== row.code) {
    Logger.log("Code did not match for " + id);
    return false;
  }
  var from = senderAddress_(message);
  if (!from || from !== row.email) {
    Logger.log("From did not match for " + id + ". Sender seen: " + from);
    return false;
  }
  sheet.getRange(row.row, map.status + 1).setValue("Verified");
  delete pending[id];
  Logger.log("Verified " + id);
  return true;
}

function pledgeIdFrom_(value) {
  var match = text_(value).match(/pledge\s+([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i);
  if (!match) {
    match = text_(value).match(/\b([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\b/i);
  }
  return match ? match[1].toLowerCase() : "";
}

function codeFrom_(value) {
  var match = text_(value).match(/verification code:\s*([a-hj-np-z2-9]{8})/i);
  return match ? match[1].toLowerCase() : "";
}

function senderAddress_(message) {
  var from = emailAddress_(message.getFrom());
  if (from && !OWN_MAILBOXES[from]) return from;
  var reply = emailAddress_(message.getReplyTo());
  if (reply && !OWN_MAILBOXES[reply]) return reply;
  var wrapped = text_(message.getPlainBody()).match(/from:\s*(?:.*<)?([^\s<>]+@[^\s<>]+)/i);
  var inner = wrapped ? text_(wrapped[1]).toLowerCase() : "";
  if (inner && !OWN_MAILBOXES[inner]) return inner;
  return from;
}

function emailAddress_(value) {
  var raw = text_(value);
  var match = raw.match(/<([^>]+)>/);
  return text_(match ? match[1] : raw).toLowerCase();
}

function gmailLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}
