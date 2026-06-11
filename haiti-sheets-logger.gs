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

// Column layout (matches sheet):
// A: Timestamp  B: #  C: First Name  D: Last Name  E: Age  F: Gender
// G: Costume  H: Child Size  I: Adult Size  J: Add-ons  K: Parent Name
// L: Phone  M: Email  N: Instagram  O: TikTok  P: Parent Apparel
// Q: Notes  R: Estimated Total  S: Deposit Due

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getActiveSheet();
    const p     = e.parameter;

    // Auto-increment # — count existing data rows (row 1 = header)
    const rowNum = Math.max(sheet.getLastRow(), 1);

    sheet.appendRow([
      new Date(),           // A: Timestamp
      rowNum,               // B: #
      p.firstName      || '', // C: First Name
      p.lastName       || '', // D: Last Name
      p.age            || '', // E: Age
      p.gender         || '', // F: Gender
      p.costume        || '', // G: Costume
      p.childSize      || '', // H: Child Size
      p.adultSize      || '', // I: Adult Size
      p.addons         || '', // J: Add-ons
      p.parentName     || '', // K: Parent Name
      p.phone          || '', // L: Phone
      p.email          || '', // M: Email
      p.instagram      || '', // N: Instagram
      p.tiktok         || '', // O: TikTok
      p.apparel        || '', // P: Parent Apparel
      p.notes          || '', // Q: Notes
      p.estimatedTotal || '', // R: Estimated Total
      p.depositDue     || '', // S: Deposit Due
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
