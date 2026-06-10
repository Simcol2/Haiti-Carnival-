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
