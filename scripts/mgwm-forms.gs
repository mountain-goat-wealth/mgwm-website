/**
 * mgwm-forms.gs — Google Apps Script web app
 * Single backend for every form on mountaingoatwealth.com. The browser never
 * talks to MailerLite directly; it POSTs here and this script holds the key.
 *
 * For each submission, in order:
 *   1. Append a row to the form's tab in the bound Sheet (must succeed — a lead
 *      is never lost even if everything after it fails).
 *   2. Roadmap only: email a plain-text alert to LEAD_ALERT_RECIPIENTS.
 *   3. Upsert the subscriber into MailerLite and record the result in the
 *      row's "MailerLite status" column.
 *
 * ── DEPLOYMENT ──
 *
 * 1. Create a Google Sheet (e.g. "MGWM Form Leads"), then open
 *    Extensions → Apps Script. This binds the script to that Sheet.
 *
 * 2. Paste this whole file into the editor (replace any existing code).
 *
 * 3. Project Settings (gear icon) → Script properties → add:
 *      ML_API_KEY             MailerLite API token (Integrations → API)
 *      LEAD_ALERT_RECIPIENTS  comma-separated emails, e.g. a@x.com,b@y.com
 *    Secrets live ONLY here. Never paste them into this file or the website.
 *
 * 4. Fill in FORMS.roadmap.groupId below with the Retirement Roadmap
 *    MailerLite group ID (MailerLite → Subscribers → Groups → open the group;
 *    the ID is in the URL).
 *
 * 5. Deploy → New deployment → type "Web app":
 *      Execute as:      Me
 *      Who has access:  Anyone
 *    Authorize when prompted. Copy the web app URL (ends in /exec) and paste it
 *    into FORMS_ENDPOINT in /assets/js/forms.js.
 *
 * 6. To update later: Deploy → Manage deployments → pencil (Edit) →
 *    Version: "New version" → Deploy. The URL stays the same.
 *    (Creating a *new deployment* instead gives a new URL and breaks the site.)
 *
 * Optional test: run testRoadmap() from the editor and check the Sheet,
 * your inbox and MailerLite.
 */

// ── CONFIG ──
// fields: the keys accepted from the browser, in Sheet column order for a new
// tab. Anything not listed is ignored. Rows are written by header name, so a
// field added later gets its own column appended to an existing tab.
var FORMS = {
  roadmap: {
    groupId: 'REPLACE_ME', // MailerLite group: Retirement Roadmap
    sheet:   'Roadmap',
    fields:  ['name', 'email', 'phone', 'state', 'age_range', 'retire_timeline',
              'lead_source', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_term',
              'gclid', 'page_url', 'st']
  },
  assessment: {
    groupId: '188300274002036208', // MailerLite group: Retirement Self-Assessment
    sheet:   'Assessment',
    fields:  ['name', 'email', 'source', 'page_url']
  },
  infographic: {
    groupId: '188300284572730636', // MailerLite group: 9-Step Infographic
    sheet:   'Infographic',
    fields:  ['email', 'page_url']
  }
};

// Which submitted fields are sent to MailerLite as subscriber fields.
// Custom fields (age_range, retire_timeline, lead_source) must already exist
// in MailerLite → Subscribers → Fields, or MailerLite drops them.
var ML_FIELDS = {
  roadmap:     ['name', 'phone', 'state', 'age_range', 'retire_timeline', 'lead_source'],
  assessment:  ['name'],
  infographic: []
};

var ML_ENDPOINT  = 'https://connect.mailerlite.com/api/subscribers';
var STATUS_COL   = 'MailerLite status';
var MAX_FIELD_LEN = 500;

function doPost(e) {
  var data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'invalid JSON' });
  }

  // Honeypot: bots fill the hidden "website" field. Pretend success, do nothing.
  if (data.website) return json_({ ok: true });

  var formName = String(data.form || '');
  var config = FORMS.hasOwnProperty(formName) ? FORMS[formName] : null;
  if (!config) return json_({ ok: false, error: 'unknown form' });

  var clean = {};
  config.fields.forEach(function (k) {
    clean[k] = data[k] == null ? '' : String(data[k]).trim().slice(0, MAX_FIELD_LEN);
  });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean.email)) {
    return json_({ ok: false, error: 'invalid email' });
  }

  // 1. Sheet row — must succeed before anything else.
  var sheet, row;
  try {
    var saved = appendRow_(formName, config, clean);
    sheet = saved.sheet;
    row = saved.row;
  } catch (err) {
    console.error('Sheet append failed', formName, err);
    return json_({ ok: false, error: 'could not save submission' });
  }

  // 2. Roadmap alert email. A failure here must not lose the lead.
  if (formName === 'roadmap') {
    try {
      sendRoadmapAlert_(clean);
    } catch (err) {
      console.error('Alert email failed', err);
    }
  }

  // 3. MailerLite upsert. Failures are recorded on the row, not surfaced.
  var status;
  try {
    status = mlUpsert_(formName, config, clean);
  } catch (err) {
    status = 'error: ' + err;
  }
  try {
    var col = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf(STATUS_COL) + 1;
    if (col > 0) sheet.getRange(row, col).setValue(status);
  } catch (err) {
    console.error('Could not write MailerLite status', err);
  }

  return json_({ ok: true });
}

function appendRow_(formName, config, clean) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(config.sheet);
    var headers = ['Timestamp', 'Form'].concat(config.fields, [STATUS_COL]);
    if (!sheet) {
      sheet = ss.insertSheet(config.sheet);
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    // Match values to the tab's actual headers; append any missing ones.
    var current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach(function (h) {
      if (current.indexOf(h) !== -1) return;
      current.push(h);
      sheet.getRange(1, current.length).setValue(h).setFontWeight('bold');
    });
    var record = { Timestamp: new Date(), Form: formName };
    config.fields.forEach(function (k) { record[k] = sheetSafe_(clean[k]); });
    record[STATUS_COL] = 'pending';
    var values = current.map(function (h) {
      return record.hasOwnProperty(h) ? record[h] : '';
    });
    sheet.appendRow(values);
    var row = sheet.getLastRow(); // safe: we hold the lock
    SpreadsheetApp.flush();
    return { sheet: sheet, row: row };
  } finally {
    lock.releaseLock();
  }
}

// Stop user input like "=IMPORTXML(...)" from being evaluated as a formula.
function sheetSafe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function sendRoadmapAlert_(d) {
  var recipients = (PropertiesService.getScriptProperties().getProperty('LEAD_ALERT_RECIPIENTS') || '')
    .split(',').map(function (s) { return s.trim(); }).filter(String);
  if (!recipients.length) {
    console.warn('LEAD_ALERT_RECIPIENTS is not set; skipping alert email.');
    return;
  }

  var subject = 'New Roadmap lead: ' + (d.name || '(no name)') +
                ' (' + (d.state || '?') + ', ' + (d.age_range || '?') + ')';
  var body = [
    'New Retirement Roadmap request.',
    '',
    'Name:             ' + d.name,
    'Email:            ' + d.email,
    'Phone:            ' + d.phone,
    'Call:             tel:+1' + d.phone.replace(/\D/g, '').replace(/^1(?=\d{10}$)/, ''),
    '',
    'State:            ' + d.state,
    'Age Range:        ' + d.age_range,
    'Retire Timeline:  ' + d.retire_timeline,
    'Lead Source:      ' + d.lead_source,
    'Photo State (st): ' + d.st,
    '',
    'UTM Source:       ' + d.utm_source,
    'UTM Medium:       ' + d.utm_medium,
    'UTM Campaign:     ' + d.utm_campaign,
    'UTM Term:         ' + d.utm_term,
    'GCLID:            ' + d.gclid,
    '',
    'Page URL:         ' + d.page_url,
    'Submitted:        ' + new Date().toLocaleString('en-US', { timeZone: 'America/Denver' }) + ' MT'
  ].join('\n');

  MailApp.sendEmail({ to: recipients.join(','), subject: subject, body: body });
}

function mlUpsert_(formName, config, clean) {
  var key = PropertiesService.getScriptProperties().getProperty('ML_API_KEY');
  if (!key) return 'skipped: ML_API_KEY not set';
  if (!config.groupId || config.groupId === 'REPLACE_ME') return 'skipped: group ID not set';

  var fields = {};
  (ML_FIELDS[formName] || []).forEach(function (k) {
    if (clean[k]) fields[k] = clean[k];
  });

  var res = UrlFetchApp.fetch(ML_ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key, Accept: 'application/json' },
    payload: JSON.stringify({ email: clean.email, fields: fields, groups: [config.groupId] }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code >= 200 && code < 300) return 'ok (' + code + ')';
  return 'error ' + code + ': ' + res.getContentText().slice(0, 300);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Manual test: run from the Apps Script editor ──
function testRoadmap() {
  var res = doPost({ postData: { contents: JSON.stringify({
    form: 'roadmap', website: '',
    name: 'Test', email: 'test@example.com', phone: '(385) 555-0100',
    state: 'Utah', age_range: '50-59', retire_timeline: 'Within 5 years',
    lead_source: 'test', st: 'none', page_url: 'apps-script-editor'
  }) } });
  console.log(res.getContent());
}
