// Google Apps Script — proxy used by url-checker.html
// Page fetching (CORS), OpenAI calls and fluff feedback stored in a Google Sheet.
//
// This file holds no secret: everything sensitive lives in
// Project Settings → Script Properties:
//   OPENAI_KEY         OpenAI API key
//   FEEDBACK_SHEET_ID  ID of the Google Sheet that stores fluff feedback
//                      (the long string in the sheet URL between /d/ and /edit)
//   ACCESS_CODES       one code per person, "name:code" separated by commas,
//                      e.g. johann:3f9c…,ed:a71b…  — remove an entry to revoke someone
//   PROXY_PASSWORD     optional: the old shared password, still accepted during the switch.
//                      Delete this property once the new page is live.
//
// Deploy: Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone
// After updating this code: Deploy → Manage deployments → ✏️ → Version: New version → Deploy
// (editing the existing deployment keeps the same /exec URL)

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Returns the name attached to the code, or null when the code is unknown
function checkAccess_(code) {
  if (!code) return null;
  var props = PropertiesService.getScriptProperties();
  var entries = (props.getProperty('ACCESS_CODES') || '').split(',');
  for (var i = 0; i < entries.length; i++) {
    var parts = entries[i].split(':');
    var c = parts.pop().trim();
    if (c && c === code) return parts.join(':').trim() || 'user';
  }
  var legacy = props.getProperty('PROXY_PASSWORD');
  if (legacy && code === legacy) return 'legacy';
  return null;
}

// ========== FEEDBACK SHEET ==========
// Creates or gets the "FluffFeedback" tab in the sheet set in FEEDBACK_SHEET_ID
function getFeedbackSheet() {
  var id = PropertiesService.getScriptProperties().getProperty('FEEDBACK_SHEET_ID');
  if (!id) return null;
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName('FluffFeedback');
  if (!sheet) {
    sheet = ss.insertSheet('FluffFeedback');
    sheet.appendRow(['Timestamp', 'URL', 'Domain', 'Quote', 'Tests Failed', 'Verdict', 'Reviewer']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ========== GET ==========
function doGet(e) {
  var p = e.parameter;
  if (!checkAccess_(p.pw)) return json_({error: 'unauthorized'});

  // Read feedback data
  if (p.action === 'feedback') {
    var sheet = getFeedbackSheet();
    if (!sheet) return json_({error: 'FEEDBACK_SHEET_ID not set in Script Properties'});
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j].toLowerCase().replace(/\s+/g, '_')] = data[i][j];
      }
      rows.push(row);
    }
    return json_({feedback: rows, count: rows.length});
  }

  // Default: CORS proxy (http/https only)
  var url = p.url;
  if (!url) return ContentService.createTextOutput('Missing url param');
  if (!/^https?:\/\//i.test(url)) return json_({error: 'invalid url'});
  var resp = UrlFetchApp.fetch(url, {muteHttpExceptions: true, followRedirects: true});
  return ContentService.createTextOutput(resp.getContentText())
    .setMimeType(ContentService.MimeType.TEXT);
}

// ========== POST ==========
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { body = {}; }
  if (!checkAccess_(body._pw)) return json_({error: 'unauthorized'});
  // The code is only for this script: OpenAI rejects unknown fields
  delete body._pw;

  // Write feedback row
  if (body.action === 'feedback') {
    var sheet = getFeedbackSheet();
    if (!sheet) return json_({error: 'FEEDBACK_SHEET_ID not set in Script Properties'});
    sheet.appendRow([
      new Date().toISOString(),
      body.url || '',
      body.domain || '',
      body.quote || '',
      body.tests_failed || '',
      body.verdict || '',
      body.reviewer || ''
    ]);
    return json_({ok: true});
  }

  // Default: OpenAI proxy
  var key = PropertiesService.getScriptProperties().getProperty('OPENAI_KEY');
  if (!key) return json_({error: 'OPENAI_KEY not set in Script Properties'});

  var resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + key
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  return ContentService.createTextOutput(resp.getContentText())
    .setMimeType(ContentService.MimeType.JSON);
}
