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
 * To update after code changes:
 *   Deploy > Manage Deployments > edit the existing deployment
 *
 * To import existing Formspree submissions:
 *   1. formspree.io → your form → Submissions → Export CSV
 *   2. Run importFromFormspreeCSV() and paste the CSV text when prompted
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

    // Write header row on first use
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);

      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#880e0e');
      headerRange.setFontColor('#ffffff');

      sheet.setColumnWidth(3, 420);   // Masqueraders
      sheet.setColumnWidth(1, 160);   // Timestamp
      sheet.setColumnWidth(4, 160);   // Parent Name
      sheet.setColumnWidth(6, 210);   // Email
      sheet.setColumnWidth(9, 220);   // Parent Apparel
      sheet.setColumnWidth(10, 220);  // Notes
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

// ── Run this once to import existing submissions from Formspree ───────────────
// 1. formspree.io → your form → Submissions → Export CSV
// 2. Select all the CSV text, copy it
// 3. Run this function and paste when prompted
function importFromFormspreeCSV() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Import Existing Registrations',
    'Paste the full CSV text exported from Formspree:',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const rows = Utilities.parseCsv(res.getResponseText());
  if (rows.length < 2) { ui.alert('No data found.'); return; }

  const headers = rows[0].map(h => h.toLowerCase().trim());
  const col = name => headers.indexOf(name);

  const ss    = SpreadsheetApp.openById('1-1tXYk2blhMCHk8k8mxfirz4oE0z5wv9xcPAkSnK94w');
  const sheet = ss.getActiveSheet();

  let imported = 0;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.every(c => !c)) continue;

    sheet.appendRow([
      row[col('date')]           || row[col('timestamp')] || row[col('created_at')] || '',
      row[col('masqueradercount')] || row[col('masquerader count')] || '1',
      row[col('masqueraders')]   || '',
      row[col('parentname')]     || row[col('parent name')] || row[col('name')] || '',
      row[col('phone')]          || '',
      row[col('email')]          || '',
      row[col('instagram')]      || 'N/A',
      row[col('tiktok')]         || 'N/A',
      row[col('apparel')]        || 'None',
      row[col('notes')]          || 'None',
      row[col('estimatedtotal')] || row[col('estimated total')] || '',
      row[col('depositdue')]     || row[col('deposit due')]     || '',
    ]);
    imported++;
  }

  ui.alert('Done! Imported ' + imported + ' registrations.');
}
