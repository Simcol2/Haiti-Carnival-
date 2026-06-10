// ─── Ayiti Cheri Carnival — Google Sheets Logger ─────────────────────────────
//
// SETUP (do this once):
//   1. Open the Google Sheet → Extensions → Apps Script
//   2. Paste this entire file, click Save
//   3. Run setupSheet() once (creates headers + column widths)
//   4. Deploy as Web App:
//        Deploy → New deployment → Web app
//        Execute as: Me
//        Who has access: Anyone
//      Copy the Web App URL
//   5. In index.html, replace SHEETS_LOGGER_URL placeholder with that URL
//
// IMPORT EXISTING SUBMISSIONS (if on Formspree Gold plan):
//   1. Project Settings → Script Properties → Add property:
//        Name:  FORMSPREE_API_KEY
//        Value: (your API key from formspree.io/account)
//   2. Run fetchFromFormspree()
//
// FREE PLAN — manual import:
//   1. Log into formspree.io → your form → Submissions → Export CSV
//   2. Run manualImportCSV() and paste the CSV text when prompted
// ─────────────────────────────────────────────────────────────────────────────

var SPREADSHEET_ID = '1-1tXYk2blhMCHk8k8mxfirz4oE0z5wv9xcPAkSnK94w';
var SHEET_NAME     = 'Registrations';

var HEADERS = [
  'Timestamp',
  'Masquerader Count',
  'Masqueraders',
  'Parent Name',
  'Phone',
  'Email',
  'Instagram',
  'TikTok',
  'Parent Apparel',
  'Notes',
  'Estimated Total',
  'Deposit Due'
];

// ── Receives every new form submission ───────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    appendRow_(data);
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Also handle GET (some deployments send query params) ─────────────────────
function doGet(e) {
  try {
    var p = e.parameter || {};
    appendRow_(p);
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Appends one row to the sheet ─────────────────────────────────────────────
function appendRow_(data) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  var now = Utilities.formatDate(
    new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm:ss'
  );

  sheet.appendRow([
    now,
    data.masqueraderCount || '',
    data.masqueraders     || '',
    data.parentName       || '',
    data.phone            || '',
    data.email            || '',
    data.instagram        || '',
    data.tiktok           || '',
    data.apparel          || '',
    data.notes            || '',
    data.estimatedTotal   || '',
    data.depositDue       || ''
  ]);
}

// ── Run once to create headers & format the sheet ────────────────────────────
function setupSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (!sheet.getRange(1, 1).getValue()) {
    var hdrRange = sheet.getRange(1, 1, 1, HEADERS.length);
    hdrRange.setValues([HEADERS]);
    hdrRange.setBackground('#880e0e');
    hdrRange.setFontColor('#ffffff');
    hdrRange.setFontWeight('bold');
    hdrRange.setFontSize(10);
    sheet.setFrozenRows(1);
  }

  var widths = [160, 80, 360, 160, 140, 210, 130, 130, 220, 220, 130, 130];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  SpreadsheetApp.getUi().alert('Sheet ready! Now deploy as a Web App and copy the URL into index.html.');
}

// ── Import existing submissions via Formspree API (Gold plan required) ───────
// Add Script Property FORMSPREE_API_KEY before running.
function fetchFromFormspree() {
  var props  = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('FORMSPREE_API_KEY');
  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      'Missing API key.\n\n' +
      'Go to Project Settings → Script Properties\n' +
      'Add: FORMSPREE_API_KEY = <your key from formspree.io/account>'
    );
    return;
  }

  var FORM_ID  = 'mgoqlgry';
  var imported = 0;
  var after    = null;

  while (true) {
    var url  = 'https://formspree.io/api/0/forms/' + FORM_ID + '/submissions' +
               (after ? '?after=' + encodeURIComponent(after) : '');
    var resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert(
        'Formspree API error ' + resp.getResponseCode() + ':\n' + resp.getContentText() +
        '\n\nIf you are on the free plan, use the manual CSV import instead (run manualImportCSV).'
      );
      return;
    }

    var json        = JSON.parse(resp.getContentText());
    var submissions = json.submissions || [];
    if (submissions.length === 0) break;

    submissions.forEach(function(sub) {
      var d = sub.data || sub;
      appendRow_({
        masqueraderCount: d.masqueraderCount || '',
        masqueraders:     d.masqueraders     || '',
        parentName:       d.parentName       || '',
        phone:            d.phone            || '',
        email:            d.email            || d._replyto || '',
        instagram:        d.instagram        || '',
        tiktok:           d.tiktok           || '',
        apparel:          d.apparel          || '',
        notes:            d.notes            || '',
        estimatedTotal:   d.estimatedTotal   || '',
        depositDue:       d.depositDue       || ''
      });
      imported++;
    });

    after = json.next || null;
    if (!after) break;
  }

  SpreadsheetApp.getUi().alert('Done! Imported ' + imported + ' submissions from Formspree.');
}

// ── Manual CSV import (free Formspree plan) ───────────────────────────────────
// 1. formspree.io → your form → Submissions → Export CSV
// 2. Run this function, paste the CSV text in the prompt
function manualImportCSV() {
  var ui  = SpreadsheetApp.getUi();
  var res = ui.prompt(
    'Paste CSV from Formspree',
    'Go to formspree.io → your form → Submissions → Export CSV.\nPaste the full CSV content here:',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  var csv  = res.getResponseText();
  var rows = Utilities.parseCsv(csv);
  if (rows.length < 2) { ui.alert('No data found in CSV.'); return; }

  var headers = rows[0].map(function(h) { return h.toLowerCase().trim(); });

  function col(name) { return headers.indexOf(name); }

  var imported = 0;
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    if (!row || row.length === 0) continue;

    appendRow_({
      masqueraderCount: row[col('masqueradercount')]  || row[col('masquerader count')]  || '',
      masqueraders:     row[col('masqueraders')]       || '',
      parentName:       row[col('parentname')]         || row[col('parent name')]         || '',
      phone:            row[col('phone')]              || '',
      email:            row[col('email')]              || '',
      instagram:        row[col('instagram')]          || '',
      tiktok:           row[col('tiktok')]             || '',
      apparel:          row[col('apparel')]            || '',
      notes:            row[col('notes')]              || '',
      estimatedTotal:   row[col('estimatedtotal')]     || row[col('estimated total')]     || '',
      depositDue:       row[col('depositdue')]         || row[col('deposit due')]         || ''
    });
    imported++;
  }

  ui.alert('Done! Imported ' + imported + ' rows from CSV.');
}
