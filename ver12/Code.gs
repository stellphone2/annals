// Annals cloud receiver. Paste into Extensions > Apps Script of your Google Sheet.
//
// SETUP
// 1) Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
// 2) When you change this code later: Deploy > Manage deployments > pencil > Version: New version > Deploy.
//    The URL stays the same. Put that URL in CLOUD_URL at the top of the app's main.py.
//
// WHO MAY DO WHAT is decided by the worksheet named "Access" in this spreadsheet:
//   Community | Emails                    | Password                | Days with notes | Edit from  | View
//   PLT       | anna@gmail.com, ravi@...   | (type the password in C2) | 41  <- script  | 01-10-2026 | Dum Dum
//   Dum Dum   | ravi@gmail.com                                       | 17              | (empty)    | all
// - the Community column is the list shown in the app (and the name of that community's worksheet);
// - Emails = the Gmail addresses that may ADD notes to that community (and of course view it);
// - VIEW (column F) = other communities the people on that row may VIEW (separate names with commas);
//   the word "all" (any capitals) means every community. Empty = only their own community.
//   Someone listed on several rows gets the union. Communities not allowed are not sent to the app
//   and a read request for them is refused;
// - PASSWORD = the first non-empty cell of column C below the heading. Change it there at any time;
//   the app asks again as soon as it stops matching. If column C is EMPTY, no password is needed
//   and the app's password field can be left empty;
// - DAYS WITH NOTES (column D) is written by this script: the number of dates that have a note in that
//   community's worksheet. All rows are refreshed every time somebody taps Sync in the app; one row is
//   also refreshed after a save and after a manual edit of that community's worksheet. To refresh all rows
//   by hand, run "updateDayCounts" in the Apps Script editor;
// - EDIT FROM (column E, optional) = a date written dd-mm-yyyy (e.g. 01-10-2026; the script formats the cells that way).
//   Notes dated on or after it
//   can be edited freely even after they were posted (the whole text can be replaced).
//   Before that date (or if column E is EMPTY) a posted note can only be EXTENDED: text may be added
//   after it, the posted text itself never changes;
// - edit the Access sheet and the apps pick up the change by themselves.
//
// SIGN-IN IS SIMPLE: the app sends the typed Gmail address and the password with every request.
// Nothing proves the person owns that address, so anyone who knows a listed address, the password
// AND this web app's URL could act as that person. Fine for a trusted group.
//
// Rules enforced here, whatever the app sends:
// - the password must match (when one is set) and the caller's address must be listed in Access;
// - one row per date; a new date is added as a new row;
// - an existing date's text can only be EXTENDED (the old text must stay at the beginning),
//   EXCEPT for dates on or after the community's "Edit from" date (column E), whose text may be replaced;
// - nothing is ever deleted.

const SPREADSHEET_ID = '1K6AqnsuByrzWLapsyaoV1lk2EiEINii46JUG9p41VRs';
const ACCESS_SHEET = 'Access';
const HEADING_WORDS = ['community', 'communities', 'name', 'worksheet']; // what A1 may say when row 1 is a heading
const DAYS_COLUMN = 4; // column D of the Access sheet: number of days that have notes, one row per community
const EDIT_COLUMN = 5; // column E of the Access sheet: "Edit from" date (empty = notes cannot be changed once posted)
const VIEW_COLUMN = 6; // column F of the Access sheet: other communities this row's people may view ("all" = every community)

// Open the web app URL in a browser to see which version of this script is really live.
function doGet() {
  return ContentService.createTextOutput('Annals script v8 - simple sign-in (Gmail address + optional password from column C of the Access sheet), day counts in column D refreshed on every Sync, "Edit from" date in column E, "View" list in column F.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const data = JSON.parse(e.postData.contents);

    // 1) password (only checked when column C of the Access sheet holds one)
    const pw = currentPassword_();
    if (pw && String(data.password || '').trim() !== pw) {
      return reply_({ ok: false, auth: false, error: 'Wrong password (it must match column C of the Access sheet).' });
    }

    // 2) the typed Gmail address, which must be listed in Access
    const email = String(data.email || '').trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      return reply_({ ok: false, auth: false, error: 'Not a valid Gmail address.' });
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tz = ss.getSpreadsheetTimeZone();
    const access = accessList_(ss, tz);
    if (!access.some(function (c) { return c.emails.indexOf(email) >= 0; })) {
      return reply_({ ok: false, error: email + ' is not listed in the "' + ACCESS_SHEET + '" sheet.' });
    }

    // the communities this person may see: their own rows plus whatever column F allows ("all" = every one)
    const viewable = viewableFor_(access, email);

    // ACCESS: the communities for the dropdown, and which of them this person may write to.
    // When the person taps Sync the app sets "sync": true, and every community's day count is rewritten.
    if (data.action === 'access') {
      if (data.sync === true) writeAllDayCounts_(ss, tz, access);
      return reply_({
        ok: true,
        email: email,
        communities: viewable.map(function (c) {
          return { name: c.name, write: c.emails.indexOf(email) >= 0, edit_from: c.editFrom };
        })
      });
    }

    // READ: only the communities this person may view (own rows + column F).
    if (data.action === 'read') {
      const readName = String(data.community || '').trim();
      if (!readName) return reply_({ ok: false, error: 'No community name was sent.' });
      if (!findCommunity_(viewable, readName)) {
        return reply_({ ok: false, error: '"' + readName + '" is not on your list of communities.' });
      }
      const found = ss.getSheetByName(readName);
      const names = ss.getSheets().map(function (sh) { return sh.getName(); });
      if (!found || found.getLastRow() < 1) {
        return reply_({ ok: true, rows: [], found: !!found, sheets: names });
      }
      const rows = found.getRange(1, 1, found.getLastRow(), 2).getValues().map(function (r) {
        return [text_(r[0], tz), text_(r[1], tz)];
      });
      return reply_({ ok: true, rows: rows, found: true, sheets: names });
    }

    // SAVE: only for the people listed for that community. Add new dates, extend existing ones.
    if (data.action === 'save') {
      const name = String(data.sheet || '').trim();
      if (!name) return reply_({ ok: false, error: 'No community name was sent.' });
      const community = findCommunity_(access, name);
      if (!community) return reply_({ ok: false, error: '"' + name + '" is not on the Access list.' });
      if (community.emails.indexOf(email) < 0) {
        return reply_({ ok: false, error: email + ' may not add notes to "' + name + '".' });
      }
      const incoming = data.rows || [];
      if (!incoming.length) return reply_({ ok: true, saved: [], skipped: [], count: 0 });

      lock.waitLock(20000); // one upload at a time
      const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
      if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, 2).setValues([['Date', 'Note']]).setFontWeight('bold');
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 100);
        sheet.setColumnWidth(2, 600);
      }

      const existing = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
      const rowsOf = {}; // date -> row numbers that carry it
      existing.forEach(function (r, i) {
        const d = text_(r[0], tz);
        (rowsOf[d] = rowsOf[d] || []).push(i + 1);
      });

      const saved = [], skipped = [], toAppend = [];
      incoming.forEach(function (r) {
        const d = String(r[0]), full = String(r[1]);
        const hits = rowsOf[d] || [];
        if (hits.length === 0) {
          toAppend.push([d, full]);
          saved.push(d);
        } else if (hits.length > 1) {
          skipped.push({ date: d, reason: 'several rows for this date in the sheet' });
        } else {
          const current = text_(existing[hits[0] - 1][1], tz);
          if (norm_(full) === norm_(current)) {
            saved.push(d); // already there
          } else if (community.editFrom && d >= community.editFrom && norm_(full)) {
            // on/after the "Edit from" date (column E): the posted text may be replaced
            const cell = sheet.getRange(hits[0], 2);
            cell.setNumberFormat('@');
            cell.setValue(full);
            cell.setWrap(true);
            saved.push(d);
          } else if (norm_(full).indexOf(norm_(current)) === 0) {
            const cell = sheet.getRange(hits[0], 2);
            cell.setNumberFormat('@'); // keep text as text (no dates or formulas)
            cell.setValue(full);
            cell.setWrap(true);
            saved.push(d);
          } else {
            skipped.push({ date: d, reason: 'the sheet text for this date changed; sync again' });
          }
        }
      });

      if (toAppend.length) {
        const range = sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 2);
        range.setNumberFormat('@');
        range.setValues(toAppend);
        range.setVerticalAlignment('top');
        range.offset(0, 1, toAppend.length, 1).setWrap(true);
      }
      writeDayCount_(ss, tz, access, name); // column D of the Access sheet
      return reply_({ ok: true, saved: saved, skipped: skipped, count: saved.length });
    }

    return reply_({ ok: false, error: 'Unknown action. Update the app and the script together.' });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// The password typed in column C of the Access sheet ('' when none is set).
function currentPassword_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(ACCESS_SHEET);
  if (sh && sh.getLastRow() >= 1) {
    const vals = sh.getRange(1, 3, sh.getLastRow(), 1).getDisplayValues();
    const a1 = String(sh.getRange(1, 1).getDisplayValue()).trim().toLowerCase();
    for (var i = 0; i < vals.length; i++) {
      const v = String(vals[i][0]).trim();
      if (!v) continue;
      if (i === 0 && (v.toLowerCase() === 'password' || HEADING_WORDS.indexOf(a1) >= 0)) continue; // row 1 is a heading row
      return v;
    }
  }
  return '';
}

// CHECK: in the Apps Script editor pick "checkPassword" in the function list and press Run, then open
// Execution log. It shows which password the web app will accept (the first non-empty cell in column C).
function checkPassword() {
  const pw = currentPassword_();
  Logger.log(pw ? 'The web app expects the password: [' + pw + '] (' + pw.length + ' characters)'
                : 'No password in column C of the "' + ACCESS_SHEET + '" sheet, so none is needed (leave the app\'s password field empty).');
}

// Column E value -> 'yyyy-MM-dd' (the format the app uses), or '' when empty / not a date.
// The sheet uses dd-mm-yyyy (a real date cell, or text such as 01-10-2026; "/" and "." also work).
function editFrom_(v, tz) {
  const t = text_(v, tz).trim().slice(0, 10); // a date cell arrives here as yyyy-MM-dd
  var y, mo, d, m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/.exec(t);
  if (m) { d = +m[1]; mo = +m[2]; y = +m[3]; }
  else if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t))) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else return '';
  const check = new Date(y, mo - 1, d);
  if (check.getFullYear() !== y || check.getMonth() !== mo - 1 || check.getDate() !== d) return ''; // e.g. 31-02-2026
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}

// Reads the Access worksheet: [{ name, emails: [lowercase emails], editFrom: 'yyyy-MM-dd' or '' }]
function accessList_(ss, tz) {
  const sh = ss.getSheetByName(ACCESS_SHEET);
  if (!sh || sh.getLastRow() < 1) return [];
  const vals = sh.getRange(1, 1, sh.getLastRow(), VIEW_COLUMN).getValues();
  const out = [];
  vals.forEach(function (r, i) {
    const name = text_(r[0], tz).trim();
    if (!name) return;
    if (i === 0 && HEADING_WORDS.indexOf(name.toLowerCase()) >= 0) return;
    if (name.toLowerCase() === ACCESS_SHEET.toLowerCase()) return;
    if (out.some(function (c) { return c.name === name; })) return; // first row for a name wins
    out.push({
      name: name,
      row: i + 1, // the row of this community in the Access sheet
      emails: text_(r[1], tz).toLowerCase().split(/[\s,;]+/).filter(Boolean),
      editFrom: editFrom_(r[EDIT_COLUMN - 1], tz), // column E: notes from this date on stay editable
      // column F: names of other communities (or "all"), lower-cased, separated by commas / semicolons / new lines
      view: text_(r[VIEW_COLUMN - 1], tz).toLowerCase().split(/[,;\n]+/)
        .map(function (s) { return s.trim(); }).filter(Boolean)
    });
  });
  return out;
}

// The communities a person may see: the rows that list their address, plus every name in those rows'
// column F; the word "all" in any of them opens every community.
function viewableFor_(access, email) {
  const mine = access.filter(function (c) { return c.emails.indexOf(email) >= 0; });
  var all = false;
  const names = {};
  mine.forEach(function (c) {
    names[c.name.toLowerCase()] = true;
    c.view.forEach(function (v) { if (v === 'all') all = true; else names[v] = true; });
  });
  return access.filter(function (c) { return all || names[c.name.toLowerCase()]; });
}

// The number of days that have a note in a community's worksheet (one row per date, blank notes do not count).
function countDays_(ss, name, tz) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 1) return 0;
  const days = {};
  sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (r) {
    const d = text_(r[0], tz).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(d) && String(r[1]).trim()) days[d] = true;
  });
  return Object.keys(days).length;
}

// Writes that number into column D of the community's row in the Access sheet.
// Does nothing for a worksheet that is not on the Access list. Never throws: a failed count must not
// fail a save or block somebody's edit.
function writeDayCount_(ss, tz, access, name) {
  try {
    const c = findCommunity_(access, name);
    if (!c) return;
    const sh = ss.getSheetByName(ACCESS_SHEET);
    if (!sh) return;
    const heading = sh.getRange(1, 1).getDisplayValue().trim().toLowerCase();
    if (HEADING_WORDS.indexOf(heading) >= 0 && !String(sh.getRange(1, DAYS_COLUMN).getDisplayValue()).trim()) {
      sh.getRange(1, DAYS_COLUMN).setValue('Days with notes').setFontWeight('bold');
    }
    sh.getRange(c.row, EDIT_COLUMN).setNumberFormat('dd-mm-yyyy'); // column E shows dates as dd-mm-yyyy
    const cell = sh.getRange(c.row, DAYS_COLUMN);
    const days = countDays_(ss, name, tz);
    if (String(cell.getDisplayValue()).trim() !== String(days)) cell.setValue(days); // write only when it changed
  } catch (err) {}
}

// Rewrites column D for every community on the Access list.
function writeAllDayCounts_(ss, tz, access) {
  access.forEach(function (c) { writeDayCount_(ss, tz, access, c.name); });
}

// REFRESH ALL BY HAND: in the Apps Script editor pick "updateDayCounts" in the function list and press Run.
// (The app does the same on every Sync.)
function updateDayCounts() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const access = accessList_(ss, tz);
  writeAllDayCounts_(ss, tz, access);
  Logger.log('Day counts written for ' + access.length + ' communities.');
}

function findCommunity_(access, name) {
  for (var i = 0; i < access.length; i++) if (access[i].name === name) return access[i];
  return null;
}

// Line endings and surrounding blanks do not count as a change.
function norm_(s) {
  return String(s).replace(/\r\n?/g, '\n').trim();
}

// A date typed by hand into the sheet comes back as a Date object; turn it into yyyy-MM-dd.
function text_(v, tz) {
  return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v);
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ON EDIT does two things (a "simple trigger": it runs by itself on manual edits, nothing to set up):
// 1) when a community worksheet is edited by hand, its "days with notes" count in the Access sheet is refreshed;
// 2) RENAME: when a community name is edited in the Access sheet (one cell in column A), the worksheet that
// carries the old name is renamed to the new one, so its notes follow. The rename does nothing when:
// - several cells were changed at once (paste, fill, delete), because Google then gives no old value;
// - there is no worksheet with the old name, or one with the new name already exists (nothing is overwritten).
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== ACCESS_SHEET) { // a note edited by hand in a community worksheet: refresh its count
      const book = e.source;
      const bookTz = book.getSpreadsheetTimeZone();
      writeDayCount_(book, bookTz, accessList_(book, bookTz), sh.getName());
      return;
    }
    if (e.range.getColumn() !== 1 || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
    const oldName = String(e.oldValue === undefined ? '' : e.oldValue).trim();
    const newName = String(e.value === undefined ? '' : e.value).trim();
    if (!oldName || !newName || oldName === newName) return;
    if (e.range.getRow() === 1 && HEADING_WORDS.indexOf(oldName.toLowerCase()) >= 0) return;
    if (oldName.toLowerCase() === ACCESS_SHEET.toLowerCase() || newName.toLowerCase() === ACCESS_SHEET.toLowerCase()) return;
    const ss = e.source;
    const target = ss.getSheetByName(oldName);
    if (!target) return;
    if (ss.getSheetByName(newName)) {
      ss.toast('A worksheet named "' + newName + '" already exists, so "' + oldName + '" was not renamed.', 'Annals', 8);
      return;
    }
    target.setName(newName);
    ss.toast('Worksheet "' + oldName + '" renamed to "' + newName + '".', 'Annals', 5);
  } catch (err) {
    // a failed rename must never block the person's edit
  }
}
