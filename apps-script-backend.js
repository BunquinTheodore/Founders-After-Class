/**
 * Founders After Class — registration backend.
 *
 * Setup:
 * 1. Create a new Google Sheet (this will hold your registrations).
 * 2. Extensions > Apps Script. Delete any starter code, paste this whole file in.
 * 3. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the deployment URL (ends in /exec).
 * 5. Paste that URL into the "formEndpoint" tweak on the registration page, then re-export/redeploy.
 *
 * Each submission becomes one row in a "Registrations" sheet tab, with the
 * payment screenshot saved to a Drive folder ("Founders After Class Payments")
 * and linked from the row so you can open it directly.
 */

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registrations');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Registrations');
    sheet.appendRow(['Timestamp', 'Full Name', 'Age', 'Contact', 'Project', 'GCash Username', 'Amount', 'Payment Screenshot']);
  }

  var screenshotUrl = '';
  if (data.screenshot) {
    var match = String(data.screenshot).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
    if (match) {
      var folder = getOrCreateFolder_('Founders After Class Payments');
      var blob = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], (data.screenshotName || 'payment') + '-' + Date.now());
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      screenshotUrl = file.getUrl();
    }
  }

  sheet.appendRow([
    data.submittedAt || new Date().toISOString(),
    data.fullName || '',
    data.age || '',
    data.contact || '',
    data.project || '',
    data.gcashUsername || '',
    data.amount || '',
    screenshotUrl
  ]);

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateFolder_(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}
