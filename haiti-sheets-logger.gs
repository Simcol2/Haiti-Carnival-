/**
 * Ayiti Cheri — Carnival Registration Logger
 * Google Apps Script Web App (doGet)
 *
 * Deployment:
 *   1. Open https://script.google.com and paste this file
 *   2. Deploy > New Deployment > Web App
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   3. Copy the deployment URL into index.html → SHEETS_LOGGER_URL
 *
 * To import existing registrations from Gmail:
 *   Run importFromGmail() — it finds all Formspree notification emails
 *   and pulls the data into the sheet automatically.
 */

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById('1-1tXYk2blhMCHk8k8mxfirz4oE0z5wv9xcPAkSnK94w');
    const sheet = ss.getActiveSheet();
    const p     = e.parameter;

    const HEADERS = [
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
      'Deposit Due',
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);

      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#880e0e');
      headerRange.setFontColor('#ffffff');

      sheet.setColumnWidth(3, 420);
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(4, 160);
      sheet.setColumnWidth(6, 210);
      sheet.setColumnWidth(9, 220);
      sheet.setColumnWidth(10, 220);
    }

    sheet.appendRow([
      new Date(),
      p.masqueraderCount || '1',
      p.masqueraders     || '',
      p.parentName       || '',
      p.phone            || '',
      p.email            || '',
      p.instagram        || 'N/A',
      p.tiktok           || 'N/A',
      p.apparel          || 'None',
      p.notes            || 'None',
      p.estimatedTotal   || '',
      p.depositDue       || '',
    ]);

    return ContentService.createTextOutput('OK');

  } catch (err) {
    return ContentService.createTextOutput('Error: ' + err.toString());
  }
}

// ── Import all existing registrations from Gmail ─────────────────────────────
// Formspree emails you every submission — this reads those emails and logs them.
// Run this once. It will ask permission to access Gmail the first time.
function importFromGmail() {
  const ui = SpreadsheetApp.getUi();

  const threads = GmailApp.search('subject:"Ayiti Cheri Registration"', 0, 500);

  if (threads.length === 0) {
    ui.alert(
      'No emails found.\n\n' +
      'Make sure the emails are in your inbox (not spam).\n' +
      'The script looks for emails with "Ayiti Cheri Registration" in the subject.'
    );
    return;
  }

  const ss    = SpreadsheetApp.openById('1-1tXYk2blhMCHk8k8mxfirz4oE0z5wv9xcPAkSnK94w');
  const sheet = ss.getActiveSheet();

  // Helper: extract a field value from Formspree email body
  function get(body, key) {
    // Formspree formats fields as "key: value" — try a few casing variants
    const pattern = new RegExp(
      '(?:^|\\n)' + key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*:\\s*([^\\n]+)',
      'i'
    );
    const m = body.match(pattern);
    return m ? m[1].trim() : '';
  }

  let imported = 0;

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      const body = msg.getPlainBody();
      const date = Utilities.formatDate(msg.getDate(), 'America/Toronto', 'yyyy-MM-dd HH:mm:ss');

      sheet.appendRow([
        date,
        get(body, 'masqueraderCount') || get(body, 'masquerader count') || '1',
        get(body, 'masqueraders')     || '',
        get(body, 'parentName')       || get(body, 'parent name')       || '',
        get(body, 'phone')            || '',
        get(body, 'email')            || '',
        get(body, 'instagram')        || 'N/A',
        get(body, 'tiktok')           || 'N/A',
        get(body, 'apparel')          || 'None',
        get(body, 'notes')            || 'None',
        get(body, 'estimatedTotal')   || get(body, 'estimated total')   || '',
        get(body, 'depositDue')       || get(body, 'deposit due')       || '',
      ]);
      imported++;
    });
  });

  ui.alert('Done! Imported ' + imported + ' registrations from Gmail.');
}
