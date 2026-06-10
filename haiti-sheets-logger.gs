/**
 * Ayiti Cheri — Carnival Registration Logger
 * Google Apps Script Web App (doGet)
 * One row per masquerader.
 *
 * Deployment:
 *   1. Open your Google Sheet → Extensions → Apps Script
 *   2. Paste this file, click Save
 *   3. Deploy > New Deployment > Web App
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   4. Copy the deployment URL into index.html → SHEETS_LOGGER_URL
 *
 * To import existing registrations:
 *   Upload your Formspree CSV to Google Drive (filename must contain "formspree")
 *   then run importFromDriveCSV()
 */

var SPREADSHEET_ID = '1-1tXYk2blhMCHk8k8mxfirz4oE0z5wv9xcPAkSnK94w';

var HEADERS = [
  'Timestamp',
  'Masq #',
  'First Name',
  'Last Name',
  'Age',
  'Gender',
  'Costume',
  'Child Size',
  'Adult Size',
  'Add-ons',
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

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getActiveSheet();
    const p     = e.parameter;

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      const hr = sheet.getRange(1, 1, 1, HEADERS.length);
      hr.setFontWeight('bold');
      hr.setBackground('#880e0e');
      hr.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1,  160);  // Timestamp
      sheet.setColumnWidth(3,  130);  // First Name
      sheet.setColumnWidth(4,  130);  // Last Name
      sheet.setColumnWidth(7,  160);  // Costume
      sheet.setColumnWidth(10, 160);  // Add-ons
      sheet.setColumnWidth(11, 150);  // Parent Name
      sheet.setColumnWidth(13, 200);  // Email
      sheet.setColumnWidth(16, 200);  // Parent Apparel
      sheet.setColumnWidth(17, 200);  // Notes
    }

    sheet.appendRow([
      new Date(),
      p.masqNum        || '1',
      p.firstName      || '',
      p.lastName       || '',
      p.age            || '',
      p.gender         || '',
      p.costume        || '',
      p.childSize      || '',
      p.adultSize      || '',
      p.addons         || '',
      p.parentName     || '',
      p.phone          || '',
      p.email          || '',
      p.instagram      || '',
      p.tiktok         || '',
      p.apparel        || '',
      p.notes          || '',
      p.estimatedTotal || '',
      p.depositDue     || '',
    ]);

    return ContentService.createTextOutput('OK');

  } catch (err) {
    return ContentService.createTextOutput('Error: ' + err.toString());
  }
}

// ── Import existing Formspree CSV from Google Drive ──────────────────────────
// 1. Upload your Formspree CSV to Google Drive (filename must contain "formspree")
// 2. Open this script from the sheet via Extensions > Apps Script
// 3. Select importFromDriveCSV and click Run
function importFromDriveCSV() {
  const files = DriveApp.searchFiles('title contains "formspree" and mimeType = "text/csv"');

  let file = null;
  if (files.hasNext()) {
    file = files.next();
  } else {
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 7);
    const recent = DriveApp.searchFiles('mimeType = "text/csv" and modifiedDate > "' + cutoff.toISOString() + '"');
    if (recent.hasNext()) file = recent.next();
  }

  if (!file) {
    Logger.log('CSV not found. Upload your Formspree CSV to Google Drive with "formspree" in the filename.');
    return;
  }

  Logger.log('Found: ' + file.getName());

  const rows = Utilities.parseCsv(file.getBlob().getDataAsString());
  if (rows.length < 2) { Logger.log('File is empty.'); return; }

  const headers = rows[0].map(h => h.toLowerCase().trim());
  const col = name => headers.indexOf(name);

  Logger.log('Columns: ' + headers.join(', '));

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getActiveSheet();

  let imported = 0;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.every(c => !c)) continue;

    // Formspree CSV has a combined "masqueraders" field — split into one row each
    const masqText = row[col('masqueraders')] || '';
    const masqLines = masqText.split(/\n\n|\n/).filter(l => l.trim());

    if (masqLines.length > 0) {
      masqLines.forEach((line, idx) => {
        const name    = (line.match(/^Masquerader \d+:\s*(.+?)\s*,\s*Age/) || [])[1] || '';
        const age     = (line.match(/Age\s+(\d+)/) || [])[1] || '';
        const gender  = (line.match(/,\s*(Girl|Boy)\s*,/) || [])[1] || '';
        const costume = (line.match(/(?:Girl|Boy),\s*(.+?)(?:\s*\(\$[\d]+\))?,\s*Child Size/) || [])[1] || '';
        const child   = (line.match(/Child Size:\s*([^,]+)/) || [])[1]?.trim() || '';
        const adult   = (line.match(/Adult Size:\s*([^,]+)/) || [])[1]?.trim() || '';
        const addons  = (line.match(/Add-ons:\s*(.+)$/) || [])[1]?.trim() || '';

        sheet.appendRow([
          row[col('date')] || row[col('timestamp')] || row[col('created_at')] || '',
          idx + 1,
          name.split(' ')[0] || '',
          name.split(' ').slice(1).join(' ') || '',
          age,
          gender,
          costume,
          child,
          adult,
          addons,
          row[col('parentname')] || row[col('parent name')] || row[col('name')] || '',
          row[col('phone')]      || '',
          row[col('email')]      || '',
          row[col('instagram')]  || '',
          row[col('tiktok')]     || '',
          row[col('apparel')]    || '',
          row[col('notes')]      || '',
          row[col('estimatedtotal')] || row[col('estimated total')] || '',
          row[col('depositdue')]     || row[col('deposit due')]     || '',
        ]);
        imported++;
      });
    } else {
      // Fallback: no masquerader breakdown, just log the whole row
      sheet.appendRow([
        row[col('date')] || '',
        '1', '', '', '', '', masqText, '', '', '',
        row[col('parentname')] || row[col('name')] || '',
        row[col('phone')]  || '',
        row[col('email')]  || '',
        row[col('instagram')] || '',
        row[col('tiktok')]    || '',
        row[col('apparel')]   || '',
        row[col('notes')]     || '',
        row[col('estimatedtotal')] || '',
        row[col('depositdue')]     || '',
      ]);
      imported++;
    }
  }

  Logger.log('Done! Imported ' + imported + ' rows.');
}
