/**
 * 246 Home Care: job log (version 2)
 * Jobs tab: every website booking.
 * Contractors tab: every "Work with 246 Home Care" sign-up.
 * Invoices tab: every invoice made in invoice.html.
 * Review requests: set a job's Status to "Done" and a one-tap WhatsApp review link appears in that row.
 *
 * To update: paste this over the old code, Save, then Deploy > Manage deployments > pencil (Edit)
 * > Version: New version > Deploy. That keeps the same web app link.
 * Then choose "setup" in the toolbar and press Run once.
 */

// Paste your Google review link here (Google Business Profile > Ask for reviews > copy link)
const REVIEW_URL = 'PASTE_YOUR_GOOGLE_REVIEW_LINK_HERE';

const TABS = {
  booking: {
    name: 'Jobs',
    headers: ['Received', 'Job ref', 'Status', 'Service', 'Job type', 'Estimate', 'Name', 'Phone', 'Parish',
      'Address', 'Drop-off', 'Payment', 'Notes', 'Details', 'Appointment', 'Quoted (BBD)', 'Deposit paid', 'Assigned to',
      'Completed', 'Full message', 'Review request'],
    status: ['New', 'Quote sent', 'Deposit paid', 'Scheduled', 'In progress', 'Done', 'Cancelled'],
    row: d => [new Date(), d.ref, 'New', d.service, d.job, d.estimate, d.name, d.phone, d.parish, d.address, d.dropoff,
      d.payment, d.notes, d.details, '', '', '', '', '', d.message, '']
  },
  contractor: {
    name: 'Contractors',
    headers: ['Received', 'Status', 'Name', 'Phone', 'Parish', 'Experience', 'Work', 'Days', 'Own transport', 'Own tools',
      'Police certificate', 'Notes', 'Full message'],
    status: ['New', 'Called', 'Trial booked', 'Approved', 'Not suitable'],
    row: d => [new Date(), 'New', d.name, d.phone, d.parish, d.experience, d.work, d.days, d.transport, d.tools,
      d.police, d.notes, d.message]
  },
  invoice: {
    name: 'Invoices',
    headers: ['Date', 'Invoice #', 'Job ref', 'Client', 'Phone', 'Address', 'Items', 'Total (BBD)', 'Deposit received',
      'Balance due', 'Status', 'Paid date'],
    status: ['Sent', 'Part paid', 'Paid', 'Cancelled'],
    statusCol: 11,
    row: d => [d.date || new Date(), d.invoice, d.ref, d.client, d.phone, d.address, d.items, num_(d.total),
      num_(d.deposit), num_(d.balance), d.status || 'Sent', '']
  }
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const kind = TABS[d.type] ? d.type : 'booking';
    const sheet = getSheet_(kind);
    const row = TABS[kind].row(d).map(v => (v instanceof Date || typeof v === 'number') ? v : clean_(v));
    if (kind === 'invoice' && d.invoice && findRow_(sheet, 2, d.invoice)) return ContentService.createTextOutput('ok');
    sheet.appendRow(row);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Status set to Done on the Jobs tab: put a one-tap WhatsApp review request in that row
function onJobEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TABS.booking.name || e.range.getColumn() !== 3 || e.range.getRow() < 2) return;
  const row = e.range.getRow();
  const reviewCell = sheet.getRange(row, TABS.booking.headers.indexOf('Review request') + 1);
  if (String(e.value) !== 'Done') return;
  const v = sheet.getRange(row, 1, 1, 10).getValues()[0];
  const name = String(v[6] || '').split(' ')[0] || 'there', service = String(v[4] || v[3] || 'job').toLowerCase();
  const phone = waNumber_(v[7]);
  if (!phone) { reviewCell.setValue('No valid phone number'); return; }
  const msg = 'Hi ' + name + ', thank you for choosing 246 Home Care! If you were happy with the ' + service +
    ', would you leave us a quick Google review? It takes 30 seconds: ' + REVIEW_URL;
  const url = 'https://wa.me/' + phone + '?text=' + encodeURIComponent(msg);
  reviewCell.setFormula('=HYPERLINK("' + url.replace(/"/g, '""') + '","📲 Send review request")');
}

// Barbados numbers: 7 digits become 1246xxxxxxx; 10 digits starting 246 get the 1
function waNumber_(p) {
  let n = String(p || '').replace(/\D/g, '');
  if (n.length === 7) n = '1246' + n;
  else if (n.length === 10 && n.startsWith('246')) n = '1' + n;
  return n.length >= 10 ? n : '';
}

function findRow_(sheet, col, value) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const vals = sheet.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(value)) return i + 2;
  return 0;
}

// Stops spreadsheet formulas being injected through form fields
function clean_(v) {
  const t = String(v == null ? '' : v).slice(0, 5000);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}
function num_(v) { const n = parseFloat(String(v == null ? '' : v).replace(/[^0-9.\-]/g, '')); return isNaN(n) ? '' : n; }

function getSheet_(kind) {
  const t = TABS[kind];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(t.name);
  if (!sheet) sheet = ss.insertSheet(t.name);
  if (sheet.getLastRow() === 0) sheet.appendRow(t.headers);
  // Keep headers up to date (adds new columns such as "Review request" to an existing tab)
  if (sheet.getLastColumn() < t.headers.length) sheet.getRange(1, 1, 1, t.headers.length).setValues([t.headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, t.headers.length).setFontWeight('bold').setBackground('#00267F').setFontColor('#FFFFFF');
  const col = t.statusCol || (t.headers.indexOf('Status') + 1);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(t.status, true).build();
  sheet.getRange(2, col, 999, 1).setDataValidation(rule);
  return sheet;
}

// Run once: creates/updates all three tabs and installs the review-request trigger
function setup() {
  Object.keys(TABS).forEach(getSheet_);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onJobEdit').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onJobEdit').forSpreadsheet(ss).onEdit().create();
}
