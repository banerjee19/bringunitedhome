/**
 * Paste this into Extensions > Apps Script on the Pledges sheet.
 * Deploy > New deployment > Web app.
 * Execute as: Me
 * Who has access: Anyone
 * Copy the web app URL into js/config.js as scriptUrl.
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
  "Time"
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
    var clean = validate_(data);
    if (clean.error) return respond_(e, { ok: false, error: clean.error });
    var id = append_(clean.value);
    return respond_(e, { ok: true, id: id, stats: readStats_() });
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

function validate_(data) {
  var first = text_(data.firstName);
  var last = text_(data.lastName);
  var email = text_(data.email).toLowerCase();
  var country = text_(data.country);
  var displayName = text_(data.displayName);
  var payment = text_(data.payment);
  var age = Number(data.age);
  var amount = Number(data.amount);
  var isPublic = data.isPublic === true || data.isPublic === "Yes";
  var agreed = data.agreed === true || data.agreed === "Yes";

  if (!first || !last) return { error: "Enter your first and last name." };
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { error: "Enter a valid email address." };
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
    var time = Utilities.formatDate(new Date(), "Etc/UTC", "yyyy-MM-dd HH:mm:ss 'UTC'");
    var map = headerMap_(sheet);
    var line = new Array(HEADERS.length);
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
    sheet.appendRow(line);
    return id;
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
    time: col(["time"])
  };
}

function isYes_(value) {
  return /^(yes|true|y|1)$/i.test(text_(value));
}

function text_(value) {
  return String(value == null ? "" : value).trim();
}

function norm_(value) {
  return text_(value).toLowerCase();
}
