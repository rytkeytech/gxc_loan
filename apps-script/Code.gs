/**
 * GENXCHOICE QUICK LOAN — application receiver
 *
 * Paste this into the Google Sheet "GENXCHOICE Quick Loan Applications":
 *   Extensions → Apps Script → replace everything → Save
 *   Deploy → New deployment → Web app
 *     Execute as: Me
 *     Who has access: Anyone
 *   Copy the Web app URL (ends in /exec) and send it to Claude.
 *
 * Each application becomes a new row in the "Applications" tab.
 * Photos and signatures are saved in a private Google Drive folder
 * called "GENXCHOICE Loan Uploads", with links in the row.
 */

const SHEET_NAME = 'Applications';
const FOLDER_NAME = 'GENXCHOICE Loan Uploads';

// Column numbers in the Applications tab (A = 1)
const COL_REF = 1;        // A  Ref, B Submitted, C Status
const COL_DATA = 8;       // H  Full Name ... AI Notes
const TEXT_COLS = [9, 10, 13, 22]; // Phone, WhatsApp, National ID, Kin Phone

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);

    // Spam trap: real people never fill this hidden field
    if (d.website) return reply({ ok: true });

    if (!d.ref || !d.fullName || !d.phone || !d.loanAmount) {
      return reply({ ok: false, error: 'Missing required details.' });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const lastRow = lastFilledRow(sheet);

    // Don't save the same application twice (e.g. double tap)
    if (lastRow > 1) {
      const refs = sheet.getRange(2, COL_REF, lastRow - 1, 1).getValues().map(r => String(r[0]));
      if (refs.indexOf(String(d.ref)) !== -1) return reply({ ok: true, ref: d.ref, duplicate: true });
    }

    const row = lastRow + 1;
    const folder = getFolder();
    const photoUrl = saveImage(folder, d.photo, d.ref + ' - applicant');
    const collateralUrl = saveImage(folder, d.collateralPhoto, d.ref + ' - collateral');
    const signatureUrl = saveImage(folder, d.signature, d.ref + ' - signature');

    sheet.getRange(row, COL_REF, 1, 3).setValues([[d.ref, new Date(), 'Pending']]);

    TEXT_COLS.forEach(c => sheet.getRange(row, c).setNumberFormat('@'));

    sheet.getRange(row, COL_DATA, 1, 28).setValues([[
      clean(d.fullName),
      clean(d.phone),
      clean(d.whatsapp),
      toDate(d.dob),
      clean(d.gender),
      clean(d.nationalId),
      clean(d.address),
      clean(d.gpsAddress),
      clean(d.employmentStatus),
      clean(d.employerName),
      toNumber(d.monthlyIncome),
      clean(d.existingLoans),
      clean(d.kinName),
      clean(d.kinRelationship),
      clean(d.kinPhone),
      toNumber(d.loanAmount),
      toNumber(d.months),
      toNumber(d.repaymentAmount),
      clean(d.loanPurpose),
      clean(d.disbursementMethod),
      clean(d.payoutDetails),
      clean(d.collateralDescription),
      toNumber(d.collateralValue),
      toDate(d.signDate),
      photoUrl,
      collateralUrl,
      signatureUrl || (d.sigMethod === 'paper' ? 'Will sign printed copy' : ''),
      ''
    ]]);

    return reply({ ok: true, ref: d.ref });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Open the Web app URL in a browser to check it's running
function doGet() {
  return reply({ ok: true, message: 'GENXCHOICE loan receiver is running.' });
}

function lastFilledRow(sheet) {
  const values = sheet.getRange(1, COL_REF, sheet.getMaxRows(), 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '' && values[i][0] !== null) return i + 1;
  }
  return 1;
}

function getFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  const folder = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function saveImage(folder, dataUrl, name) {
  if (!dataUrl || typeof dataUrl !== 'string') return '';
  const match = dataUrl.match(/^data:(image\/[a-zA-Z+]+);base64,(.+)$/);
  if (!match) return '';
  const ext = match[1] === 'image/png' ? '.png' : '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], name + ext);
  return folder.createFile(blob).getUrl();
}

function clean(v) {
  if (v === undefined || v === null) return '';
  let s = String(v).trim().slice(0, 1000);
  // Stop text being run as a spreadsheet formula
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function toNumber(v) {
  const n = parseFloat(v);
  return isNaN(n) ? '' : n;
}

function toDate(v) {
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : '';
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
