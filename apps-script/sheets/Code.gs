// Personal / WORK sheets. The same code runs in both files.
// Works straight away (no permission needed):
//   Pick a Priority -> rows re-order (New, High, Medium, blank, Low).
//   Tick Done?/Bought? -> row goes to the bottom, ticked (the Planner app's Sunday-night clear-out deletes it).
//   Untick it -> row goes back up to its priority.
//   Pick a tab in Move to -> row moves to that tab.
//   Any change in a row -> Date modified shows today.
// Needs "Extras > Turn on extras" once in each file (Google asks you to Allow):
//   Tick the notes box -> pop-up for longer notes, emails and links.
//   Move to Personal / WORK -> row moves to the other file.
//   Personal file only: 10am email of High items + timed reminders (done rows are skipped).

const FILES = {
  Personal: '1_V6vrgl6jfVMnTlC40_zy7Sn7__FdjIFXh8avcVB3_A',
  WORK: '1ueRmxGq1M2oQX3TkOfiEPdLRHQ-Yw9h5eAbsH-C5tOc',
};
const EMAIL_TO = 'wells247@hotmail.com';
const TZ = 'Europe/London';

const LAYOUT = {
  'To-do':         {prio:1, text:2, det:3, area:4, done:5, move:6, date:7, remind:8, repeat:9, width:9},
  'To buy':        {prio:1, text:2, det:3, area:4, done:5, move:6, date:7, width:7},
  'Remember':      {prio:1, text:2, det:3, area:4, move:5, date:6, width:6},
  'Notes & ideas': {prio:1, text:2, det:3, area:4, move:5, date:6, width:6},
  'To film':       {prio:1, text:2, det:3, area:4, date:8, width:9},  // no Done/Move: Captured? keeps the row
};
const RANK = {'New': 0, 'High': 1, 'Medium': 2, '': 3, 'Low': 4};
const REPEAT_HRS = {'Every 1 hr': 1, 'Every 3 hrs': 3, 'Every 6 hrs': 6};
const REPEAT_START_HOUR = 9;   // a repeat set with a date but no time starts at 9am
const HAS_DETAILS = '#9fc5e8';

function thisFileName_(ss) {
  return ss.getId() === FILES.WORK ? 'WORK' : 'Personal';
}
function otherFileName_(ss) {
  return thisFileName_(ss) === 'WORK' ? 'Personal' : 'WORK';
}

// ---------- Simple trigger: runs for every edit, no permission needed ----------
function onEdit(e) {
  const sh = e.range.getSheet();
  const L = LAYOUT[sh.getName()];
  if (!L || e.range.getRow() < 2 || e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
  const col = e.range.getColumn();
  const r = e.range.getRow();
  if (col === L.det) return;                                   // handled by the pop-up
  if (col === L.move && FILES[e.value]) return;                // handled by the cross-file move
  if (L.done && col === L.done) return doneChanged_(e, sh, L, r);
  if (col === L.move) return moveRow_(e, sh, L, r);
  if (col !== L.date && col <= L.width) stampDate_(sh, L, r);
  // Sort only once the row has text, so picking a priority before typing (or typing after) both work.
  const hasText = String(sh.getRange(r, L.text).getValue()).trim() !== '';
  const hasPrio = String(sh.getRange(r, L.prio).getValue()).trim() !== '';
  if (hasText && (col === L.prio || (col === L.text && hasPrio))) sortByPriority_(sh, L);
}

// Last row that has text in it (empty tick-boxes below don't count).
function lastDataRow_(sh, L) {
  const n = sh.getLastRow();
  if (n < 2) return 1;
  const t = sh.getRange(2, L.text, n - 1, 1).getValues();
  for (let i = t.length - 1; i >= 0; i--) if (String(t[i][0]).trim() !== '') return i + 2;
  return 1;
}

function stampDate_(sh, L, r) {
  sh.getRange(r, L.date).setValue(Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy'));
}

// Ticked: the row drops below everything not done. Unticked: it goes back up to its priority.
function doneChanged_(e, sh, L, r) {
  if (String(sh.getRange(r, L.text).getValue()).trim() === '') return;
  stampDate_(sh, L, r);
  sortByPriority_(sh, L);
  if (String(e.value).toUpperCase() === 'TRUE') e.source.toast('Done. Moved to the bottom; cleared on Sunday night.');
}

const isDone_ = (L, v) => !!L.done && v[L.done - 1] === true;

// Sorts with the sheet's own sort, so notes, colours and tick-boxes move with their row.
// Order: New, High, Medium, blank, Low, then done rows (in the same order), then empty rows.
function sortByPriority_(sh, L) {
  const last = lastDataRow_(sh, L);
  if (last < 3) return;
  const n = last - 1;
  const helper = L.width + 1;
  const rows = sh.getRange(2, 1, n, Math.max(2, L.done || 0)).getValues();
  const keys = rows.map((p, i) => {
    let k = RANK[String(p[L.prio - 1]).trim()];
    if (k === undefined) k = 3;
    if (isDone_(L, p)) k += 10;                               // done rows sit below the rest
    if (String(p[L.text - 1]).trim() === '') k = 99;         // empty rows go to the bottom
    return [k * 100000 + i];
  });
  const hr = sh.getRange(2, helper, n, 1);
  hr.setValues(keys);
  sh.getRange(2, 1, n, helper).sort({column: helper, ascending: true});
  hr.clearContent();
}

// Move to another tab in the same file.
function moveRow_(e, sh, from, r) {
  const targetName = e.value;
  const to = LAYOUT[targetName];
  if (!to || targetName === sh.getName()) return;
  e.range.clearContent();
  const v = sh.getRange(r, 1, 1, from.width).getValues()[0];
  if (v[from.text - 1] === '') return;
  const tgt = e.source.getSheetByName(targetName);
  const start = addRow_(tgt, to);
  // Priority, text, notes box (with its pop-up notes and colour) and area sit in the same columns on every tab.
  sh.getRange(r, 1, 1, 4).copyTo(tgt.getRange(start, 1, 1, 4));
  tgt.getRange(start, to.det).setValue(false);
  tgt.getRange(start, to.date).setValue(Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy'));
  sh.deleteRow(r);
  sortByPriority_(tgt, to);
  e.source.toast('Moved to ' + targetName);
}

// Uses the first empty row under the list. Dropdowns and tick-boxes already run down every column.
function addRow_(tgt, L) {
  const start = lastDataRow_(tgt, L) + 1;
  const dst = tgt.getRange(start, 1, 1, L.width);
  dst.clearContent();
  dst.clearNote();
  dst.setBackground(null);
  tgt.getRange(start, L.text).setWrap(true);
  tgt.getRange(start, L.det).setValue(false);
  if (L.done) tgt.getRange(start, L.done).setValue(false);
  return start;
}

// ---------- Installable trigger: needs "Turn on extras" once ----------
function onEditExtras(e) {
  const sh = e.range.getSheet();
  const L = LAYOUT[sh.getName()];
  if (!L || e.range.getRow() < 2 || e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
  const col = e.range.getColumn();
  const r = e.range.getRow();
  if (col === L.det && String(e.value).toUpperCase() === 'TRUE') return openDetails_(sh, L, r, e.range);
  if (col === L.move && FILES[e.value] && e.value !== thisFileName_(e.source)) return moveToOtherFile_(e, sh, L, r);
}

function openDetails_(sh, L, r, cell) {
  cell.setValue(false);
  const text = String(sh.getRange(r, L.text).getValue());
  const t = HtmlService.createTemplateFromFile('Details');
  t.sheetName = sh.getName();
  t.row = r;
  t.task = text;
  t.details = cell.getNote();
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(640).setHeight(520), 'Details');
}

// Called from the pop-up. Finds the row again by its text, in case rows moved while it was open.
function saveDetails(sheetName, row, task, details) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const L = LAYOUT[sheetName];
  let r = row;
  if (String(sh.getRange(r, L.text).getValue()) !== task) {
    const texts = sh.getRange(2, L.text, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
    const i = texts.findIndex(x => String(x[0]) === task);
    if (i < 0) throw new Error('Could not find the row. It may have been moved or deleted.');
    r = i + 2;
  }
  const cell = sh.getRange(r, L.det);
  const clean = String(details || '').trim();
  cell.setNote(clean);
  cell.setBackground(clean ? HAS_DETAILS : null);
  stampDate_(sh, L, r);
  return true;
}

function moveToOtherFile_(e, sh, L, r) {
  const targetFile = e.value;
  e.range.clearContent();
  const v = sh.getRange(r, 1, 1, L.width).getValues()[0];
  if (v[L.text - 1] === '') return;
  const note = sh.getRange(r, L.det).getNote();
  const other = SpreadsheetApp.openById(FILES[targetFile]);
  const tgt = other.getSheetByName(sh.getName());
  const to = LAYOUT[sh.getName()];
  const start = addRow_(tgt, to);
  // Area is left blank: the two files have different areas.
  tgt.getRange(start, 1, 1, 2).setValues([[v[L.prio - 1], v[L.text - 1]]]);
  if (note) tgt.getRange(start, to.det).setNote(note).setBackground(HAS_DETAILS);
  tgt.getRange(start, to.date).setValue(Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy'));
  if (to.remind) tgt.getRange(start, to.remind, 1, 2).setValues([[v[L.remind - 1], v[L.repeat - 1]]]);
  sh.deleteRow(r);
  sortByPriority_(tgt, to);
  e.source.toast('Moved to ' + targetFile + ' > ' + sh.getName());
}

// ---------- Menu and one-time setup ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Extras')
    .addItem('Turn on extras (one time)', 'turnOnExtras')
    .addItem('Open notes box for selected row', 'openDetailsForSelection')
    .addItem('Send me the High priority email now', 'sendDailyEmail')
    .addToUi();
}

function openDetailsForSelection() {
  const sh = SpreadsheetApp.getActiveSheet();
  const L = LAYOUT[sh.getName()];
  const r = sh.getActiveRange().getRow();
  if (!L || r < 2) return SpreadsheetApp.getUi().alert('Select a row on To-do, To buy, Remember or Notes & ideas first.');
  openDetails_(sh, L, r, sh.getRange(r, L.det));
}

function turnOnExtras() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onEditExtras').forSpreadsheet(ss).onEdit().create();
  if (thisFileName_(ss) === 'Personal') {
    ScriptApp.newTrigger('sendDailyEmail').timeBased().atHour(10).nearMinute(0).everyDays(1).inTimezone(TZ).create();
    ScriptApp.newTrigger('checkReminders').timeBased().everyMinutes(15).create();
    ss.toast('Extras on: notes pop-up, move to WORK, 10am email and reminders.', 'Done', 8);
  } else {
    ss.toast('Extras on: notes pop-up and move to Personal.', 'Done', 8);
  }
}

// ---------- Emails (run from the Personal file) ----------
function sendDailyEmail() {
  const now = new Date();
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  let high = '', due = '';
  const toClear = [];
  ['Personal', 'WORK'].forEach(f => {
    const ss = SpreadsheetApp.openById(FILES[f]);
    ['To-do'].forEach(tab => {                       // morning email: To-do tabs only
      const L = LAYOUT[tab];
      const sh = ss.getSheetByName(tab);
      if (!sh || sh.getLastRow() < 2) return;
      const n = sh.getLastRow() - 1;
      const vals = sh.getRange(2, 1, n, L.width).getValues();
      const notes = sh.getRange(2, L.det, n, 1).getNotes();
      vals.forEach((v, i) => {
        const text = v[L.text - 1];
        if (!text || isDone_(L, v)) return;            // done rows aren't in the email
        const extra = notes[i][0] ? '<br><span style="color:#666">' + esc_(notes[i][0]).replace(/\n/g, '<br>') + '</span>' : '';
        if (v[L.prio - 1] === 'High') high += '<li><b>' + f + ' &gt; ' + tab + ':</b> ' + esc_(text) + extra + '</li>';
        if (L.remind && v[L.remind - 1] instanceof Date) {
          const d = v[L.remind - 1];
          const dateOnly = d.getHours() === 0 && d.getMinutes() === 0;
          if (dateOnly && !REPEAT_HRS[v[L.repeat - 1]] && Utilities.formatDate(d, TZ, 'yyyy-MM-dd') <= today) {
            due += '<li><b>' + f + ':</b> ' + esc_(text) + extra + '</li>';
            toClear.push(sh.getRange(i + 2, L.remind));
          }
        }
      });
    });
  });
  let body = '<h3>High priority to-dos</h3>' + (high ? '<ul>' + high + '</ul>' : '<p>No to-dos marked High.</p>');
  if (due) body += '<h3>Reminders for today</h3><ul>' + due + '</ul>';
  // Top things to film (WORK > To film): High items not yet captured; if none, the first 3 open ones.
  let film = '';
  const fsh = SpreadsheetApp.openById(FILES.WORK).getSheetByName('To film');
  if (fsh && fsh.getLastRow() > 1) {
    const FL = LAYOUT['To film'];
    const fv = fsh.getRange(2, 1, fsh.getLastRow() - 1, FL.width).getValues();
    const open = fv.filter(v => String(v[1]).trim() !== '' && v[4] !== true);
    let top = open.filter(v => v[0] === 'High');
    if (!top.length) top = open.filter(v => v[0] !== 'Low').slice(0, 3);
    film = top.map(v => '<li>' + esc_(v[1]) + ' <span style="color:#666">(' + esc_(v[3]) + ', ' + esc_(v[5]) + (v[8] ? ', file name ' + esc_(v[8]) : '') + ')</span></li>').join('');
  }
  body += '<h3>Top things to film</h3>' + (film ? '<ul>' + film + '</ul>' : '<p>Nothing on the To film list.</p>');
  body += '<p><a href="https://docs.google.com/spreadsheets/d/' + FILES.Personal + '/edit">Open Personal</a> · <a href="https://docs.google.com/spreadsheets/d/' + FILES.WORK + '/edit">Open WORK</a></p>';
  MailApp.sendEmail({to: EMAIL_TO, subject: 'Today: High priority to-dos (' + Utilities.formatDate(now, TZ, 'EEE d MMM') + ')', htmlBody: body});
  toClear.forEach(c => c.clearContent());
}

// Every 15 minutes: emails timed reminders that are due, then moves repeating ones forward.
// Repeats only run on the reminder's own day (from 9am if no time was set), and stop when the task is done.
function checkReminders() {
  const now = new Date();
  ['Personal', 'WORK'].forEach(f => {
    const ss = SpreadsheetApp.openById(FILES[f]);
    const L = LAYOUT['To-do'];
    const sh = ss.getSheetByName('To-do');
    if (!sh || sh.getLastRow() < 2) return;
    const n = sh.getLastRow() - 1;
    const vals = sh.getRange(2, 1, n, L.width).getValues();
    const notes = sh.getRange(2, L.det, n, 1).getNotes();
    vals.forEach((v, i) => {
      let d = v[L.remind - 1];
      const text = v[L.text - 1];
      if (!(d instanceof Date) || !text || isDone_(L, v)) return;
      const hrs = REPEAT_HRS[v[L.repeat - 1]];
      const dateOnly = d.getHours() === 0 && d.getMinutes() === 0;
      if (dateOnly && !hrs) return;                      // date-only goes in the 10am email
      if (dateOnly) d = new Date(d.getTime() + REPEAT_START_HOUR * 3600000);
      if (d > now) return;
      const cell = sh.getRange(i + 2, L.remind);
      const day = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
      if (hrs && day < Utilities.formatDate(now, TZ, 'yyyy-MM-dd')) { cell.clearContent(); return; }  // that day has passed
      const body = '<p><b>' + esc_(text) + '</b></p>' + (notes[i][0] ? '<p>' + esc_(notes[i][0]).replace(/\n/g, '<br>') + '</p>' : '') +
        '<p><a href="https://docs.google.com/spreadsheets/d/' + FILES[f] + '/edit">Open ' + f + '</a></p>';
      MailApp.sendEmail({to: EMAIL_TO, subject: 'Reminder: ' + text, htmlBody: body});
      if (hrs) {
        let next = new Date(d.getTime());
        while (next <= now) next = new Date(next.getTime() + hrs * 3600000);
        if (Utilities.formatDate(next, TZ, 'yyyy-MM-dd') === day) cell.setValue(next);
        else cell.clearContent();                        // no more repeats after that day
      } else {
        cell.clearContent();
      }
    });
  });
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
