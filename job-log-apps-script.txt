/**
 * 246 Home Care: job log
 * Paste this into Extensions > Apps Script in the "246 Home Care Jobs" Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Every booking from 246homecare.com is added as a new row on the "Jobs" tab.
 */
const SHEET_NAME = 'Jobs';
const HEADERS = ['Received', 'Job ref', 'Status', 'Service', 'Job type', 'Estimate', 'Name', 'Phone', 'Parish',
  'Address', 'Drop-off', 'Payment', 'Notes', 'Details', 'Appointment', 'Quoted (BBD)', 'Deposit paid', 'Assigned to', 'Completed', 'Full message'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const sheet = getSheet_();
    sheet.appendRow([
      new Date(), clean_(d.ref), 'New', clean_(d.service), clean_(d.job), clean_(d.estimate), clean_(d.name), clean_(d.phone),
      clean_(d.parish), clean_(d.address), clean_(d.dropoff), clean_(d.payment), clean_(d.notes), clean_(d.details),
      '', '', '', '', '', clean_(d.message)
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Stops spreadsheet formulas being injected through form fields
function clean_(v) {
  const t = String(v == null ? '' : v).slice(0, 5000);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#00267F').setFontColor('#FFFFFF');
    const status = SpreadsheetApp.newDataValidation()
      .requireValueInList(['New', 'Quote sent', 'Deposit paid', 'Scheduled', 'In progress', 'Done', 'Cancelled'], true).build();
    sheet.getRange('C2:C1000').setDataValidation(status);
  }
  return sheet;
}

// Run once from the editor (select "setup" and press Run) to create the tab and headers
function setup() { getSheet_(); }
