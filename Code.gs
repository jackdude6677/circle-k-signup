/**
 * Circle K Club sign-up: the Google Sheets side.
 *
 * Lives inside the roster spreadsheet (Extensions > Apps Script) and is
 * deployed as a web app. The sign-up page posts each sign-up to this web
 * app's address, and it writes one row per person, one answer per column.
 */

const SHEET_NAME = 'Members';
const HEADERS = ['Role', 'Last Name', 'First Name', 'Date of Birth', 'Grade', 'Email', 'Phone Number', 'Signed Up'];
// '@' keeps a column as plain text, so a phone number or a name is never
// reinterpreted as a number, a date or a formula.
const FORMATS = ['@', '@', '@', 'mmm d, yyyy', '@', '@', '@', 'mmm d, yyyy h:mm am/pm'];

const ROLES = ['President', 'Vice President', 'Secretary', 'Treasurer', 'Member'];
const GRADES = [
  '11th grade', '12th grade',
  'College freshman', 'College sophomore', 'College junior', 'College senior',
  'Other'
];

const NAME_RE = /^[\p{L}][\p{L} .'’\-]{0,49}$/u;
const EMAIL_RE = /^[a-z0-9][a-z0-9._%+\-]*@[a-z0-9][a-z0-9.\-]*\.[a-z]{2,}$/;

/** Opening the web app address in a browser shows that it is running. */
function doGet() {
  return ContentService.createTextOutput('Circle K sign-up is running.');
}

/** The sign-up page posts here. Always answers with JSON. */
function doPost(e) {
  let result;
  try {
    result = submitMember(JSON.parse(e.postData.contents));
  } catch (err) {
    console.error(err);
    result = { ok: false, error: 'not-saved' };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Re-checks every answer (the page's own checks can be bypassed), then adds
 * one row to the Members sheet.
 * Returns { ok: true } or { ok: false, errors: { fieldName: message } }.
 */
function submitMember(data) {
  data = data || {};
  // Hidden field real people never see. Bots fill it in; drop those quietly.
  if (clean_(data.website)) return { ok: true };

  const errors = {};

  const role = ROLES.indexOf(data.role) === -1 ? 'Member' : data.role;

  const lastName = clean_(data.lastName);
  if (!NAME_RE.test(lastName)) errors.lastName = 'Enter your last name using letters only.';

  const firstName = clean_(data.firstName);
  if (!NAME_RE.test(firstName)) errors.firstName = 'Enter your first name using letters only.';

  const dob = parseDate_(data.dob);
  if (!dob) errors.dob = 'Choose a real date of birth that is not in the future.';

  const grade = clean_(data.grade);
  if (GRADES.indexOf(grade) === -1) errors.grade = 'Choose your grade.';

  const email = clean_(data.email).toLowerCase();
  if (email.length > 120 || !EMAIL_RE.test(email)) errors.email = 'Enter an email address like name@example.com.';

  let phone = String(data.phone || '').replace(/\D/g, '');
  if (phone.length === 11 && phone.charAt(0) === '1') phone = phone.slice(1);
  if (phone.length !== 10) errors.phone = 'Enter a 10-digit phone number, like (555) 123-4567.';
  phone = '(' + phone.slice(0, 3) + ') ' + phone.slice(3, 6) + '-' + phone.slice(6);

  if (Object.keys(errors).length) return { ok: false, errors: errors };

  // One writer at a time, so two people pressing the button together
  // cannot land on the same row.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    const last = sheet.getLastRow();
    // Same person sent twice in a row (a retry after a dropped connection):
    // the first one is already saved, so do not add a second row.
    if (last > 1) {
      const prev = sheet.getRange(last, 1, 1, 7).getDisplayValues()[0];
      if (prev[0] === role && prev[1] === lastName && prev[2] === firstName && prev[5] === email && prev[6] === phone) {
        return { ok: true };
      }
    }
    const row = last + 1;
    if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 50);
    const range = sheet.getRange(row, 1, 1, HEADERS.length);
    range.setNumberFormats([FORMATS]);
    range.setValues([[role, lastName, firstName, dob, grade, email, phone, new Date()]]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/** Finds the Members tab, creating it and its header row if they are missing. */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const first = ss.getSheets()[0];
    sheet = first.getLastRow() === 0 ? first.setName(SHEET_NAME) : ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function clean_(value) {
  return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
}

/** Turns "2009-03-14" into a Date, or null if it is not a real past date. */
function parseDate_(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ''));
  if (!m) return null;
  const year = Number(m[1]), month = Number(m[2]), day = Number(m[3]);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  const now = new Date();
  if (date > now || year < now.getFullYear() - 100) return null;
  return date;
}
