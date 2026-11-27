/**
 * Planner app: the link between the phone app and the Personal and WORK Google Sheets.
 * Paste this whole file into a new project at script.google.com (see SETUP.md).
 *
 * It does the same things the sheets do when you edit them by hand, because sheet
 * scripts don't react to changes made by another script:
 *   - new notes go in at the top, marked New, with today's Date modified
 *   - changing Priority moves the row into place (New, High, Medium, Low)
 *   - Done? / Bought? deletes the row; Captured? (To film) ticks it and keeps it
 *   - voice notes from the Voice notes app wait in a hidden "Voice inbox" tab in Personal
 *     until the 8pm job sorts them into the right sheet and tab
 */

// ---- Fill these two in during setup (SETUP.md tells you when) ----
const API_URL = 'https://script.google.com/macros/s/AKfycbw6u2uNPcxY8Khygi9ljhHHYDcZxvqS7CBcfDM8-rLj8EAZqkDCSwF5OcWy57PUHwHY/exec';  // the Web app URL (from Deploy)
const APP_URL = 'https://mw2472012-lgtm.github.io/planner/';  // the phone app's web address (GitHub Pages)
// -------------------------------------------------------------------

const BOOKS = {
  personal: { id: '1_V6vrgl6jfVMnTlC40_zy7Sn7__FdjIFXh8avcVB3_A', tabs: ['To-do', 'To buy', 'Remember', 'Notes & ideas'] },
  work:     { id: '1ueRmxGq1M2oQX3TkOfiEPdLRHQ-Yw9h5eAbsH-C5tOc', tabs: ['To-do', 'To buy', 'Remember', 'Notes & ideas', 'To film'] },
};
const RANK = { New: 0, High: 1, Medium: 2, Low: 3 };
const INBOX_TAB = 'Voice inbox';   // hidden tab in Personal: Note | Recorded | ID

/** Run this once after deploying: it emails you the link to open on your phone. */
function sendMeTheLink() {
  if (!API_URL || !APP_URL) throw new Error('Fill in API_URL and APP_URL at the top first.');
  const link = APP_URL.replace(/\/?$/, '/') + '#setup=' + encodeURIComponent(JSON.stringify({ api: API_URL, key: key_() }));
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Your Planner app link',
    'Open this on your phone (in Chrome):\n\n' + link + '\n\nThen tap Chrome\'s menu and choose "Install app" or "Add to Home screen".\nKeep this email private: the link lets anyone who has it read and change your lists.');
  Logger.log('Sent. Check your email on your phone.');
}

/** Run this if the link ever leaks: it makes a new key, so the old link stops working. */
function resetKey() {
  PropertiesService.getScriptProperties().deleteProperty('KEY');
  sendMeTheLink();
}

function key_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('KEY');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('KEY', k); }
  return k;
}

function doGet(e) { return respond_(e.parameter || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return respond_(p);
}

function respond_(p) {
  let out;
  try {
    if (!p.key || p.key !== key_()) throw new Error('Not allowed. Use the link from your email again.');
    const fn = ACTIONS[p.action];
    if (!fn) throw new Error('Unknown action');
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { out = { ok: true, data: fn(p) }; } finally { lock.releaseLock(); }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

const ACTIONS = {
  /** Everything in both sheets, for the app's lists. */
  all() {
    const result = {};
    Object.keys(BOOKS).forEach(b => {
      const ss = SpreadsheetApp.openById(BOOKS[b].id);
      result[b] = { url: ss.getUrl(), tabs: BOOKS[b].tabs.map(t => readTab_(ss, t)).filter(Boolean) };
    });
    result.voice = inboxRead_();
    return result;
  },

  /** Voice notes waiting for the 8pm sort. */
  inboxList() { return inboxRead_(); },

  /** Save a voice note. The app gives each note an id, so a retry after lost signal never doubles it. */
  inboxAdd(p) {
    const text = String(p.text || '').trim();
    if (!text) throw new Error('Nothing to save');
    const sh = inbox_();
    if (inboxFind_(sh, p.id) < 0) sh.appendRow([text, p.at ? new Date(p.at) : new Date(), String(p.id || Utilities.getUuid())]);
    return true;
  },

  /** Change the wording of a waiting voice note. */
  inboxEdit(p) {
    const sh = inbox_();
    const r = inboxFind_(sh, p.id);
    if (r < 0) throw new Error('That note has already been sorted into your lists.');
    sh.getRange(r, 1).setValue(String(p.text || '').trim());
    return true;
  },

  /** Delete a waiting voice note. */
  inboxRemove(p) {
    const sh = inbox_();
    const r = inboxFind_(sh, p.id);
    if (r >= 0) sh.deleteRow(r);
    return true;
  },

  /** Add a new row at the top of a tab, marked New. */
  add(p) {
    addRow_(p.book, p.tab, p.text, p.area);
    return true;
  },

  /** Move a row to another list (or the same list in the other sheet), keeping its priority. */
  move(p) {
    const src = sheet_(p.book, p.tab);
    const r = find_(src, p);
    const c = cols_(src);
    const v = src.getRange(r, 1, 1, src.getLastColumn()).getValues()[0];
    const priority = c.pri >= 0 ? String(v[c.pri]) : '';
    const area = c.area >= 0 ? String(v[c.area]) : '';
    const dst = addRow_(p.toBook, p.toTab, String(v[c.text]), area);
    const dc = cols_(dst);
    if (dc.area >= 0 && area) {   // Personal and Work have different areas: clear it if it doesn't fit
      const dv = dst.getRange(2, dc.area + 1).getDataValidation();
      const ok = dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST
        ? dv.getCriteriaValues()[0].indexOf(area) >= 0 : p.book === p.toBook;
      if (!ok) dst.getRange(2, dc.area + 1).setValue('');
    }
    if (priority in RANK && priority !== 'New') {
      dst.getRange(2, dc.pri + 1).setValue(priority);
      placeRow_(dst, 2, dc, priority);
    }
    if (p.book !== p.toBook || p.tab !== p.toTab) src.deleteRow(find_(src, { row: r, text: p.text }));
    return true;
  },
  /** Tick Done? / Bought? (row is deleted) or Captured? (row is kept, ticked). */
  done(p) {
    const sh = sheet_(p.book, p.tab);
    const r = find_(sh, p);
    const c = cols_(sh);
    if (p.tab === 'To film') {
      sh.getRange(r, c.done + 1).setValue(true);
      touch_(sh, r, c);
    } else {
      sh.deleteRow(r);
    }
    return true;
  },

  /** Delete a row (for Remember and Notes & ideas, which have no Done box). */
  remove(p) {
    const sh = sheet_(p.book, p.tab);
    sh.deleteRow(find_(sh, p));
    return true;
  },

  /** Change the wording, area and/or priority of a row. */
  update(p) {
    const sh = sheet_(p.book, p.tab);
    const r = find_(sh, p);
    const c = cols_(sh);
    if (typeof p.newText === 'string' && p.newText.trim()) sh.getRange(r, c.text + 1).setValue(p.newText.trim());
    if (c.area >= 0 && typeof p.area === 'string') sh.getRange(r, c.area + 1).setValue(p.area);
    if (c.remind >= 0 && typeof p.remind === 'string') setRemind_(sh.getRange(r, c.remind + 1), p.remind);
    if (c.repeat >= 0 && typeof p.repeat === 'string') sh.getRange(r, c.repeat + 1).setValue(p.repeat);
    touch_(sh, r, c);
    if (c.pri >= 0 && p.priority && p.priority in RANK) {
      sh.getRange(r, c.pri + 1).setValue(p.priority);
      placeRow_(sh, r, c, p.priority);
    }
    return true;
  },
};

/** Put a new row at the top of a tab, marked New, copying the look of the row below. */
function addRow_(book, tab, text, area) {
  const sh = sheet_(book, tab);
  const n = sh.getLastColumn();
  const c = cols_(sh);
  sh.insertRowBefore(2);
  const dst = sh.getRange(2, 1, 1, n);
  if (sh.getLastRow() >= 3) {
    const src = sh.getRange(3, 1, 1, n);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
  const row = new Array(n).fill('');
  if (c.pri >= 0) row[c.pri] = 'New';
  row[c.text] = String(text || '').trim();
  if (c.area >= 0) row[c.area] = area || '';
  if (c.note >= 0) row[c.note] = false;
  if (c.done >= 0) row[c.done] = false;
  if (c.date >= 0) row[c.date] = today_();
  dst.setValues([row]);
  if (c.note >= 0) sh.getRange(2, c.note + 1).setBackground(null); // a new row has no 📝 note yet
  return sh;
}

function readTab_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) return null;
  const last = sh.getLastRow();
  const n = sh.getLastColumn();
  const c = cols_(sh);
  let areas = [];
  if (c.area >= 0 && sh.getMaxRows() >= 2) {
    const dv = sh.getRange(2, c.area + 1).getDataValidation();
    if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) areas = dv.getCriteriaValues()[0];
  }
  const items = [];
  if (last >= 2) {
    const values = sh.getRange(2, 1, last - 1, n).getValues();
    const noteBg = c.note >= 0 ? sh.getRange(2, c.note + 1, last - 1, 1).getBackgrounds() : null;
    const tz = ss.getSpreadsheetTimeZone();
    values.forEach((v, i) => {
      const text = String(v[c.text] || '').trim();
      if (!text) return;
      items.push({
        row: i + 2,
        text: text,
        priority: c.pri >= 0 ? String(v[c.pri] || '') : '',
        area: c.area >= 0 ? String(v[c.area] || '') : '',
        done: c.done >= 0 ? v[c.done] === true : false,
        hasNote: noteBg ? !/^#?(ffffff|fff)$/i.test(noteBg[i][0]) : false,
        modified: fmt_(v[c.date], tz),
        remind: fmt_(v[c.remind], tz, true),
        repeat: c.repeat >= 0 ? String(v[c.repeat] || '') : '',
      });
    });
  }
  return {
    name: name,
    gid: sh.getSheetId(),
    areaLabel: c.area >= 0 ? String(sh.getRange(1, c.area + 1).getValue()) : '',
    doneLabel: c.done >= 0 ? String(sh.getRange(1, c.done + 1).getValue()).replace('?', '') : '',
    areas: areas,
    items: items,
    hasRemind: c.remind >= 0,
    repeatOptions: c.repeat >= 0 && sh.getMaxRows() >= 2 ? listOf_(sh.getRange(2, c.repeat + 1)) : [],
  };
}

function cols_(sh) {
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  const f = re => h.findIndex(x => re.test(x));
  return {
    pri: f(/^Priority$/i),
    text: 1,
    note: f(/📝/),
    area: f(/^(Area|Project)$/i),
    done: f(/^(Done|Bought|Captured)\?$/i),
    date: f(/^Date modified$/i),
    remind: f(/Remind me on/i),
    repeat: f(/^Repeat$/i),
  };
}

/** The hidden Voice inbox tab in Personal (made the first time it's needed). */
function inbox_() {
  const ss = SpreadsheetApp.openById(BOOKS.personal.id);
  let sh = ss.getSheetByName(INBOX_TAB);
  if (!sh) {
    sh = ss.insertSheet(INBOX_TAB, ss.getSheets().length);
    sh.getRange(1, 1, 1, 3).setValues([['Note', 'Recorded', 'ID']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 500);
    sh.hideSheet();
  }
  return sh;
}

function inboxRead_() {
  const ss = SpreadsheetApp.openById(BOOKS.personal.id);
  const sh = ss.getSheetByName(INBOX_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  const tz = ss.getSpreadsheetTimeZone();
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(v => String(v[0]).trim())
    .map(v => ({ id: String(v[2]), text: String(v[0]).trim(), at: v[1] instanceof Date ? v[1].getTime() : 0,
      when: v[1] instanceof Date ? Utilities.formatDate(v[1], tz, 'dd/MM HH:mm') : '' }));
}

function inboxFind_(sh, id) {
  if (!id || sh.getLastRow() < 2) return -1;
  const ids = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function sheet_(book, tab) {
  if (!BOOKS[book] || BOOKS[book].tabs.indexOf(tab) < 0) throw new Error('Unknown list');
  const sh = SpreadsheetApp.openById(BOOKS[book].id).getSheetByName(tab);
  if (!sh) throw new Error('Tab not found: ' + tab);
  return sh;
}

/** Find the row: trust the row number if the words still match, otherwise look for the words. */
function find_(sh, p) {
  const want = String(p.text || '').trim();
  const last = sh.getLastRow();
  if (p.row >= 2 && p.row <= last && String(sh.getRange(p.row, 2).getValue()).trim() === want) return p.row;
  if (last >= 2) {
    const col = sh.getRange(2, 2, last - 1, 1).getValues();
    for (let i = 0; i < col.length; i++) if (String(col[i][0]).trim() === want) return i + 2;
  }
  throw new Error('That item has changed in the sheet. Pull down to refresh.');
}

/** Move row r so it sits with the others of the same priority (New, High, Medium, Low). */
function placeRow_(sh, r, c, priority) {
  const last = sh.getLastRow();
  if (last < 3) return;
  const data = sh.getRange(2, 1, last - 1, Math.max(c.pri, c.text) + 1).getValues();
  const mine = RANK[priority];
  let dest = -1, lastFilled = 1;
  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    if (!String(data[i][c.text]).trim()) continue;
    lastFilled = row;
    if (row === r) continue;
    const rank = String(data[i][c.pri]) in RANK ? RANK[String(data[i][c.pri])] : 9;
    if (dest < 0 && rank > mine) dest = row;
  }
  if (dest < 0) dest = lastFilled + 1;
  if (dest === r || dest === r + 1) return;
  sh.moveRows(sh.getRange(r + ':' + r), dest);
}

/** The choices in a cell's dropdown, if it has one. */
function listOf_(cell) {
  const dv = cell.getDataValidation();
  return dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST ? dv.getCriteriaValues()[0] : [];
}

/** Remind me on: '' clears it, 'yyyy-mm-dd' is a date, 'yyyy-mm-ddTHH:MM' is a date and time. */
function setRemind_(cell, v) {
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/);
  if (!m) { cell.setValue(''); return; }
  const d = new Date(+m[1], +m[2] - 1, +m[3], m[4] ? +m[4] : 0, m[5] ? +m[5] : 0);
  cell.setValue(d).setNumberFormat(m[4] && (+m[4] || +m[5]) ? 'dd/mm/yyyy hh:mm' : 'dd/mm/yyyy');
}

function touch_(sh, r, c) { if (c.date >= 0) sh.getRange(r, c.date + 1).setValue(today_()); }

function today_() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }

function fmt_(v, tz, withTime) {
  if (!(v instanceof Date)) return v ? String(v) : '';
  const hasTime = v.getHours() || v.getMinutes();
  return Utilities.formatDate(v, tz, withTime && hasTime ? 'dd/MM/yyyy HH:mm' : 'dd/MM/yyyy');
}
