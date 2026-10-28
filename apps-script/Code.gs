/**
 * Event Attendance System — Google Apps Script
 *
 * Flow:
 *   1. Student submits the Google Form.
 *   2. onFormSubmitHandler() gives them a random token and emails a QR code.
 *      If the daily email quota is used up, the row is marked PENDING and
 *      sendPending() (runs hourly) sends it later.
 *   3. Students can also open the web app URL to look up their QR code.
 *   4. On event day, the scanner page POSTs the scanned token to doPost(),
 *      which records attendance in the "Attendance" sheet.
 */

// ===== EDIT THESE TO MATCH YOUR FORM =====
const CONFIG = {
  EVENT_NAME: 'Our Event',
  EVENT_DETAILS: 'Date / time / place here',

  RESPONSES_SHEET: 'Form Responses 1', // tab created by Google Forms
  ATTENDANCE_SHEET: 'Attendance',

  // Column headers = your form question titles (case-insensitive)
  COL_EMAIL: 'Email Address',
  COL_NAME: 'Full Name',
  COL_STUDENT_ID: 'Student ID',

  // Columns added by the script
  COL_TOKEN: 'Token',
  COL_STATUS: 'QR Email Status',
};
// =========================================

const QR_API = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&margin=10&data=';

/** Run once from the editor: adds columns, Attendance tab and triggers. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getResponsesSheet_();

  const map = getHeaderMap_(sheet);
  [CONFIG.COL_TOKEN, CONFIG.COL_STATUS].forEach(function (h) {
    if (!map[h.toLowerCase()]) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
    }
  });

  if (!ss.getSheetByName(CONFIG.ATTENDANCE_SHEET)) {
    const att = ss.insertSheet(CONFIG.ATTENDANCE_SHEET);
    att.appendRow(['Check-in Time', 'Token', 'Name', 'Student ID', 'Email', 'Method']);
    att.setFrozenRows(1);
  }

  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmitHandler').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('sendPending').timeBased().everyHours(1).create();

  // Gives tokens/emails to anyone who registered before setup was run
  sendPending();
}

/** Set the PIN staff must enter in the scanner page. Edit and run once. */
function setStaffPin() {
  PropertiesService.getScriptProperties().setProperty('STAFF_PIN', 'change-me-1234');
}

function onFormSubmitHandler(e) {
  const sheet = getResponsesSheet_();
  const row = e.range.getRow();
  ensureToken_(sheet, row);
  sendQrEmail_(sheet, row);
}

/** Sends QR emails to every row not yet SENT, while quota remains. */
function sendPending() {
  const sheet = getResponsesSheet_();
  const map = getHeaderMap_(sheet);
  const last = sheet.getLastRow();
  for (let row = 2; row <= last; row++) {
    ensureToken_(sheet, row);
    const status = sheet.getRange(row, col_(map, CONFIG.COL_STATUS)).getValue();
    if (String(status).indexOf('SENT') === 0) continue;
    if (MailApp.getRemainingDailyQuota() < 1) break;
    sendQrEmail_(sheet, row);
  }
}

// ---------- Web app ----------

/** GET: the "Get my QR code" page for students. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Lookup')
    .setTitle(CONFIG.EVENT_NAME + ' — My QR code')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called from Lookup.html. Both email and student ID must match. */
function lookupQr(email, studentId) {
  email = String(email || '').trim().toLowerCase();
  studentId = String(studentId || '').trim().toLowerCase();
  if (!email || !studentId) return { ok: false, message: 'Please fill in both fields.' };

  const sheet = getResponsesSheet_();
  const map = getHeaderMap_(sheet);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const r = data[i];
    if (String(r[col_(map, CONFIG.COL_EMAIL) - 1]).trim().toLowerCase() === email &&
        String(r[col_(map, CONFIG.COL_STUDENT_ID) - 1]).trim().toLowerCase() === studentId) {
      const token = ensureToken_(sheet, i + 1);
      return {
        ok: true,
        name: r[col_(map, CONFIG.COL_NAME) - 1],
        event: CONFIG.EVENT_NAME,
        qr: 'data:image/png;base64,' + Utilities.base64Encode(fetchQr_(token).getBytes()),
      };
    }
  }
  return { ok: false, message: 'No registration found. Check your email and student ID.' };
}

/**
 * POST from the scanner page. Body (JSON, sent as text/plain):
 *   { pin, token }        — scanned QR
 *   { pin, studentId }    — manual check-in
 */
function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ status: 'error', message: 'Bad request' });
  }

  const pin = PropertiesService.getScriptProperties().getProperty('STAFF_PIN');
  if (!pin || req.pin !== pin) return json_({ status: 'bad_pin', message: 'Wrong staff PIN' });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return json_(checkIn_(req));
  } finally {
    lock.releaseLock();
  }
}

function checkIn_(req) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getResponsesSheet_();
  const map = getHeaderMap_(sheet);
  const data = sheet.getDataRange().getValues();

  const token = String(req.token || '').trim();
  const sid = String(req.studentId || '').trim().toLowerCase();
  const tokenCol = col_(map, CONFIG.COL_TOKEN) - 1;
  const sidCol = col_(map, CONFIG.COL_STUDENT_ID) - 1;

  let found = null;
  for (let i = data.length - 1; i >= 1; i--) {
    if ((token && String(data[i][tokenCol]) === token) ||
        (!token && sid && String(data[i][sidCol]).trim().toLowerCase() === sid)) {
      found = data[i];
      break;
    }
  }
  if (!found) return { status: 'not_found', message: 'Not registered' };

  const person = {
    token: String(found[tokenCol]),
    name: found[col_(map, CONFIG.COL_NAME) - 1],
    studentId: found[sidCol],
    email: found[col_(map, CONFIG.COL_EMAIL) - 1],
  };

  const att = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  const attData = att.getDataRange().getValues();
  for (let j = 1; j < attData.length; j++) {
    if (String(attData[j][1]) === person.token) {
      return Object.assign({ status: 'duplicate', message: 'Already checked in',
        time: formatTime_(attData[j][0]) }, person);
    }
  }

  const now = new Date();
  att.appendRow([now, person.token, person.name, person.studentId, person.email,
    token ? 'QR' : 'Manual']);
  return Object.assign({ status: 'ok', message: 'Checked in', time: formatTime_(now) }, person);
}

// ---------- Helpers ----------

function ensureToken_(sheet, row) {
  const map = getHeaderMap_(sheet);
  const cell = sheet.getRange(row, col_(map, CONFIG.COL_TOKEN));
  let token = cell.getValue();
  if (!token) {
    token = 'ATT-' + Utilities.getUuid();
    cell.setValue(token);
  }
  return String(token);
}

function sendQrEmail_(sheet, row) {
  const map = getHeaderMap_(sheet);
  const statusCell = sheet.getRange(row, col_(map, CONFIG.COL_STATUS));
  if (MailApp.getRemainingDailyQuota() < 1) {
    statusCell.setValue('PENDING (daily email limit)');
    return;
  }
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const email = values[col_(map, CONFIG.COL_EMAIL) - 1];
  const name = values[col_(map, CONFIG.COL_NAME) - 1];
  const token = values[col_(map, CONFIG.COL_TOKEN) - 1];
  try {
    const qr = fetchQr_(token).setName('qr.png');
    MailApp.sendEmail({
      to: email,
      subject: 'Your QR code for ' + CONFIG.EVENT_NAME,
      htmlBody:
        '<p>Hi ' + escapeHtml_(name) + ',</p>' +
        '<p>Thank you for registering for <b>' + escapeHtml_(CONFIG.EVENT_NAME) + '</b>.<br>' +
        escapeHtml_(CONFIG.EVENT_DETAILS) + '</p>' +
        '<p>Please show this QR code at the entrance:</p>' +
        '<p><img src="cid:qr" width="250" height="250" alt="QR code"></p>' +
        '<p style="color:#666;font-size:12px">Do not share this code. It is unique to you.</p>',
      inlineImages: { qr: qr },
      attachments: [qr],
    });
    statusCell.setValue('SENT ' + formatTime_(new Date()));
  } catch (err) {
    statusCell.setValue('ERROR: ' + err.message);
  }
}

function fetchQr_(token) {
  return UrlFetchApp.fetch(QR_API + encodeURIComponent(token)).getBlob();
}

function getResponsesSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.RESPONSES_SHEET);
  if (!sheet) throw new Error('Sheet "' + CONFIG.RESPONSES_SHEET + '" not found. Check CONFIG.RESPONSES_SHEET.');
  return sheet;
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function (h, i) {
    map[String(h).trim().toLowerCase()] = i + 1;
  });
  return map;
}

function col_(map, header) {
  const c = map[header.toLowerCase()];
  if (!c) throw new Error('Column "' + header + '" not found. Check CONFIG or run setup().');
  return c;
}

function formatTime_(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
