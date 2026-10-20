/**
 * JCA Sports Reservation API
 * Google Sheets is the shared database.
 *
 * Setup:
 * 1. Create a Google Sheet.
 * 2. Open Extensions > Apps Script and paste this file.
 * 3. Run setupDatabase() once and authorize it.
 * 4. Deploy > New deployment > Web app.
 *    Execute as: Me
 *    Who has access: Anyone (or the access level your school requires)
 * 5. Copy the Web app URL ending in /exec into the website's API_URL.
 */

const SHEET_NAME = 'Reservations';
const HEADERS = [
  'Booking ID', 'Full Name', 'Facility', 'Reservation Date',
  'Start Hour', 'Start Time', 'Status', 'Created At'
];
const FACILITIES = ['Seed Dome', 'Skyline Arch', 'Sky Arch'];
const START_HOUR = 8;
const END_HOUR = 19; // final session starts at 7:00 PM

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Create/open the Google Sheet, then open Apps Script from Extensions.');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, HEADERS.length);
  }
  return 'Database ready: ' + ss.getUrl();
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  try {
    if (action === 'availability') {
      const facility = String(e.parameter.facility || '');
      const date = String(e.parameter.date || '');
      if (!FACILITIES.includes(facility) || !isValidDate_(date)) {
        return json_({ ok: false, error: 'Invalid facility or date.' });
      }
      const bookings = getBookings_();
      return json_({
        ok: true,
        bookings: bookings.filter(b => b.facility === facility && b.date === date),
        reservedHours: bookings.filter(b => b.facility === facility && b.date === date).map(b => b.hour)
      });
    }
    if (action === 'list') {
      return json_({ ok: true, bookings: getBookings_() });
    }
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return json_({ ok: false, error: 'Unable to read reservations. ' + err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (payload.action !== 'create') {
      return json_({ ok: false, error: 'Unknown action.' });
    }

    const name = String(payload.name || '').trim();
    const facility = String(payload.facility || '');
    const date = String(payload.date || '');
    const hour = Number(payload.hour);

    if (!name || name.length > 80) return json_({ ok: false, error: 'Enter a valid full name (up to 80 characters).' });
    if (!FACILITIES.includes(facility)) return json_({ ok: false, error: 'Choose a valid facility.' });
    if (!isValidDate_(date)) return json_({ ok: false, error: 'Choose a valid reservation date.' });
    if (!Number.isInteger(hour) || hour < START_HOUR || hour > END_HOUR) {
      return json_({ ok: false, error: 'Choose a valid one-hour timeslot.' });
    }
    if (date < today_()) return json_({ ok: false, error: 'Past dates cannot be booked.' });

    lock.waitLock(15000);
    const existing = getBookings_();
    const conflict = existing.some(b =>
      b.facility === facility && b.date === date && Number(b.hour) === hour
    );
    if (conflict) return json_({ ok: false, error: 'That timeslot has already been reserved. Please choose another.' });

    const sheet = getSheet_();
    const id = Utilities.getUuid();
    const createdAt = new Date();
    sheet.appendRow([
      id, safeCell_(name), facility, date, hour, formatHour_(hour), 'Confirmed', createdAt
    ]);
    return json_({
      ok: true,
      booking: {
        id: id, name: name, facility: facility, date: date, hour: hour,
        status: 'Confirmed', createdAt: createdAt.toISOString()
      }
    });
  } catch (err) {
    return json_({ ok: false, error: 'Could not save reservation. ' + err.message });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function getBookings_() {
  const sheet = getSheet_();
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  return rows.map(r => ({
    id: String(r[0]),
    name: String(r[1]),
    facility: String(r[2]),
    date: normalizeDate_(r[3]),
    hour: Number(r[4]),
    time: String(r[5]),
    status: String(r[6] || 'Confirmed'),
    createdAt: r[7] instanceof Date ? r[7].toISOString() : String(r[7] || '')
  })).filter(b => b.id && b.name && FACILITIES.includes(b.facility) && isValidDate_(b.date));
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet. Open Apps Script from the Google Sheet.');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function isValidDate_(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const d = new Date(value + 'T12:00:00Z');
  return !isNaN(d.getTime()) && d.toISOString().slice(0, 10) === value;
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normalizeDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value || '').slice(0, 10);
}

function formatHour_(hour) {
  const start = new Date(2000, 0, 1, hour);
  const end = new Date(2000, 0, 1, hour + 1);
  return Utilities.formatDate(start, Session.getScriptTimeZone(), 'h:mm a') + ' – ' +
         Utilities.formatDate(end, Session.getScriptTimeZone(), 'h:mm a');
}

function safeCell_(value) {
  // Prevent spreadsheet formula injection from names beginning with formula characters.
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
