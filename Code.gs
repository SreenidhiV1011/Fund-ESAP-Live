/**
 * ESG IC Form endpoint — Apps Script web app
 *
 * What it does, on every submission from the Fund ESAP Tracker:
 *   1. Checks the caller is signed in with a @blume.vc Google account
 *      (verifies the Firebase ID token the tracker sends).
 *   2. Re-checks the required fields and the three eligibility confirmations.
 *   3. Appends one row to the "IC Form Responses" tab of the ESG Checklist sheet.
 *   4. Emails you (REVIEWER) a review notice with a link to the row.
 *
 * Deploy: Deploy > New deployment > Web app
 *         Execute as: Me   |   Who has access: Anyone
 * Then paste the web app URL into ENDPOINT at the top of esg-ic.js.
 */

var SHEET_ID        = '1myqEzSSK6b6ZpaO6tWvRq9gBhHksn-lH_WuN8NoYMPY';
var TAB_NAME        = 'IC Form Responses';
var REVIEWER        = 'sreenidhi@blume.vc';
var ALLOWED_DOMAIN  = '@blume.vc';
var FIREBASE_API_KEY = 'AIzaSyB8J-qZ9g76-_GDqitqX5M5dVaogFMQMFs'; // public web key from the tracker

var HEADERS = [
  'Timestamp', 'Submitted by', 'Company Name', 'Sector', 'Brief Business description',
  'Operates in India', 'Operates outside India', 'Exposure countries',
  'Male founders', 'Female founders', 'Website',
  'Eligibility confirmations (all three ticked)', 'Operating in a Developing Country',
  'Mandatory condition', 'Secondary condition', 'SDG alignment',
  'Permanent employees', 'Part-time employees', 'Contractual employees',
  'Repeat submission', 'Review status', 'Reviewed on'
];

function doGet() {
  return json_({ ok: true, message: 'ESG IC Form endpoint is live' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var body = JSON.parse(e.postData.contents);
    var email = verifyToken_(body.idToken);
    if (!email) return json_({ ok: false, error: 'Not signed in with a Blume account' });

    var d = body.data || {};
    var problem = validate_(d);
    if (problem) return json_({ ok: false, error: problem });

    lock.waitLock(20000);
    locked = true;
    var sh = getSheet_();
    var repeat = isRepeat_(sh, d.company);
    sh.appendRow(toRow_(email, d, repeat));
    var row = sh.getLastRow();
    notify_(email, d, repeat, row, sh);
    return json_({ ok: true, row: row, repeat: repeat });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (x) {} }
  }
}

/* ── identity ────────────────────────────────────────────────────────── */
function verifyToken_(idToken) {
  if (!idToken) return null;
  var res = UrlFetchApp.fetch(
    'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY,
    { method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  var users = JSON.parse(res.getContentText()).users;
  if (!users || !users.length) return null;
  var u = users[0];
  var mail = String(u.email || '').toLowerCase();
  if (!u.emailVerified || mail.slice(-ALLOWED_DOMAIN.length) !== ALLOWED_DOMAIN) return null;
  return mail;
}

/* ── validation (mirrors the form, so a hand-built request can't skip it) ── */
function validate_(d) {
  var must = ['company', 'sectorLabel', 'description', 'sdg'];
  for (var i = 0; i < must.length; i++) {
    if (!String(d[must[i]] || '').trim()) return 'Missing field: ' + must[i];
  }
  var c = d.confirmations || {};
  if (!(c.prohibition && c.debarred && c.myanmar)) return 'All three eligibility confirmations are required';
  if (!d.operatesIndia && !d.operatesOutside) return 'Operating geography is required';
  if (d.operatesOutside && !(d.exposureCountries && d.exposureCountries.length)) return 'Exposure countries are required';
  if (d.developingCountry !== 'Yes' && d.developingCountry !== 'No') return 'Developing Country answer is required';
  return '';
}

/* ── sheet ───────────────────────────────────────────────────────────── */
function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_NAME);
  if (!sh) {
    sh = ss.insertSheet(TAB_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Reviewed'], true).build();
    sh.getRange(2, HEADERS.indexOf('Review status') + 1, 2000, 1).setDataValidation(rule);
  }
  return sh;
}

function isRepeat_(sh, company) {
  var last = sh.getLastRow();
  if (last < 2) return false;
  var names = sh.getRange(2, 3, last - 1, 1).getValues();
  var key = String(company).trim().toLowerCase();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim().toLowerCase() === key) return true;
  }
  return false;
}

function toRow_(email, d, repeat) {
  return [
    new Date(), email, d.company, d.sectorLabel, d.description,
    d.operatesIndia ? 'Yes' : 'No', d.operatesOutside ? 'Yes' : 'No',
    (d.exposureCountries || []).join(', '),
    d.maleFounders, d.femaleFounders, d.website || '',
    'Yes', d.developingCountry,
    (d.mandatory || []).join(', '), (d.secondary || []).join(', '), d.sdg,
    d.permanent, d.partTime, d.contractual,
    repeat ? 'Yes' : 'No', 'Pending', ''
  ];
}

/* ── review email ────────────────────────────────────────────────────── */
function notify_(email, d, repeat, row, sh) {
  var flags = [];
  if (repeat) flags.push('Repeat submission for this Company');
  if (d.developingCountry === 'No') flags.push('Marked as not operating in a Developing Country');
  var geo = [d.operatesIndia ? 'India' : '',
             d.operatesOutside ? 'Outside India (' + (d.exposureCountries || []).join(', ') + ')' : '']
            .filter(String).join('; ');
  var link = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit#gid=' + sh.getSheetId() + '&range=A' + row;
  var html =
    '<p><b>' + esc_(d.company) + '</b> was submitted on the ESG IC Form by ' + esc_(email) + '.</p>' +
    (flags.length ? '<p style="color:#b45309"><b>Check:</b> ' + esc_(flags.join('. ')) + '.</p>' : '') +
    '<table cellpadding="4" style="border-collapse:collapse">' +
    '<tr><td><b>Sector</b></td><td>' + esc_(d.sectorLabel) + '</td></tr>' +
    '<tr><td><b>Geography</b></td><td>' + esc_(geo) + '</td></tr>' +
    '<tr><td><b>Developing Country</b></td><td>' + esc_(d.developingCountry) + '</td></tr>' +
    '<tr><td valign="top"><b>SDG alignment</b></td><td>' + esc_(d.sdg).replace(/\n/g, '<br>') + '</td></tr></table>' +
    '<p><a href="' + link + '">Open the row</a> and set Review status to Reviewed when done.</p>';
  MailApp.sendEmail({
    to: REVIEWER,
    subject: 'ESG IC Form: ' + d.company + ' (' + d.sectorLabel + ')' + (repeat ? ' [repeat]' : ''),
    htmlBody: html
  });
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
