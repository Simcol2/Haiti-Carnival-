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
 * To import existing registrations from the Formspree CSV:
 *   1. Upload the downloaded CSV anywhere in your Google Drive
 *   2. Run importFromDriveCSV() — it will find the file automatically
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

// ── Import from Formspree CSV uploaded to Google Drive ───────────────────────
// 1. Upload your downloaded CSV anywhere in Google Drive
// 2. Select importFromDriveCSV from the dropdown and click Run
function importFromDriveCSV() {
  const ui = SpreadsheetApp.getUi();

  // Search Drive for any CSV file with "formspree" in the name
  const files = DriveApp.searchFiles('title contains "formspree" and mimeType = "text/csv"');

  let file = null;
  if (files.hasNext()) {
    file = files.next();
  } else {
    // Fallback: look for any CSV uploaded recently (last 7 days)
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 7);
    const recent = DriveApp.searchFiles('mimeType = "text/csv" and modifiedDate > "' + cutoff.toISOString() + '"');
    if (recent.hasNext()) file = recent.next();
  }

  if (!file) {
    ui.alert(
      'CSV not found in Google Drive.\n\n' +
      'Upload your Formspree CSV to Google Drive (anywhere), then run this again.\n' +
      'Make sure the filename contains the word "formspree".'
    );
    return;
  }

  const csvText = file.getBlob().getDataAsString();
  const rows    = Utilities.parseCsv(csvText);

  if (rows.length < 2) {
    ui.alert('The CSV file appears to be empty.');
    return;
  }

  const headers = rows[0].map(h => h.toLowerCase().trim());
  const col = name => headers.indexOf(name);

  const ss    = SpreadsheetApp.openById('1-1tXYk2blhMCHk8k8mxfirz4oE0z5wv9xcPAkSnK94w');
  const sheet = ss.getActiveSheet();

  let imported = 0;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.every(c => !c)) continue;

    sheet.appendRow([
      row[col('date')]               || row[col('timestamp')]        || row[col('created_at')] || '',
      row[col('masqueradercount')]   || row[col('masquerader count')] || '1',
      row[col('masqueraders')]       || '',
      row[col('parentname')]         || row[col('parent name')]       || row[col('name')] || '',
      row[col('phone')]              || '',
      row[col('email')]              || '',
      row[col('instagram')]          || 'N/A',
      row[col('tiktok')]             || 'N/A',
      row[col('apparel')]            || 'None',
      row[col('notes')]              || 'None',
      row[col('estimatedtotal')]     || row[col('estimated total')]   || '',
      row[col('depositdue')]         || row[col('deposit due')]       || '',
    ]);
    imported++;
  }

  ui.alert('Done! Imported ' + imported + ' registrations from "' + file.getName() + '".');
}
