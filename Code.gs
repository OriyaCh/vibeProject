/**
 * מצרכים בבית – צד שרת (Google Apps Script, מחובר לגיליון).
 * הגיליון הוא מקור האמת: לשונית _data שומרת את כל הנתונים,
 * ולשוניות קריאות (קטלוג, מלאי, משפחה, חשבוניות, יומן, רשימה שבועית) נכתבות מחדש בכל שמירה.
 * שינויים שנעשים ידנית בלשונית הרשימה השבועית ובעמודת הכמות בלשונית המלאי נקלטים בחזרה באפליקציה.
 */

var DATA_SHEET = '_data';
var CHUNK = 45000;
var CATS = {veg:'ירקות ופירות',bread:'לחם ומאפים',dairy:'חלב וביצים',meat:'בשר ודגים',frozen:'קפואים',pantry:'מזווה ויבשים',snacks:'חטיפים ומתוקים',drinks:'משקאות',clean:'ניקיון',hyg:'היגיינה',baby:'תינוקות',other:'אחר'};
var CAT_ORDER = Object.keys(CATS);

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('מצרכים בבית')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function props_() { return PropertiesService.getDocumentProperties(); }
function version_() { return Number(props_().getProperty('version') || 0); }

function sheet_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) s = ss_().insertSheet(name);
  return s;
}

function readJson_() {
  var s = ss_().getSheetByName(DATA_SHEET);
  if (!s || s.getLastRow() === 0) return null;
  var vals = s.getRange(1, 1, s.getLastRow(), 1).getValues();
  var txt = vals.map(function (r) { return r[0]; }).join('');
  return txt ? JSON.parse(txt) : null;
}

function writeJson_(state) {
  var s = sheet_(DATA_SHEET);
  var txt = JSON.stringify(state);
  var rows = [];
  for (var i = 0; i < txt.length; i += CHUNK) rows.push([txt.substr(i, CHUNK)]);
  s.clear();
  s.getRange(1, 1, rows.length, 1).setNumberFormat('@').setValues(rows);
  s.hideSheet();
}

function whoAmI() { return Session.getActiveUser().getEmail(); }
function getSheetUrl() { return ss_().getUrl(); }

/** מחזיר את המצב, אחרי קליטת שינויים ידניים מהגיליון. */
function getState(clientVersion) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var state = readJson_();
    if (!state) return { empty: true, version: 0 };
    var changed = absorbSheetEdits_(state);
    var v = version_();
    if (changed) {
      v++;
      writeJson_(state);
      props_().setProperty('version', String(v));
      writeTabs_(state);
    }
    if (!changed && Number(clientVersion) === v) return { version: v, same: true, json: JSON.stringify(state) };
    return { version: v, changed: changed, json: JSON.stringify(state) };
  } finally {
    lock.releaseLock();
  }
}

/** שמירה עם בדיקת גרסה (מונע דריסה של עדכון מקביל). */
function saveState(json, baseVersion) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var v = version_();
    var current = readJson_();
    if (current && Number(baseVersion) !== v) {
      return { conflict: true, version: v, json: JSON.stringify(current) };
    }
    var state = JSON.parse(json);
    v++;
    writeJson_(state);
    props_().setProperty('version', String(v));
    writeTabs_(state);
    return { version: v };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- לשוניות קריאות ---------- */
function dm_(d) { var x = new Date(d); return x.getDate() + '.' + (x.getMonth() + 1); }
function listTabName_(state) { return 'רשימה ' + dm_(state.list.week || new Date()); }
function active_(a) { return (a || []).filter(function (x) { return !x.del; }); }
function byId_(state, id) { for (var i = 0; i < state.catalog.length; i++) if (state.catalog[i].id === id) return state.catalog[i]; return null; }
function fmt_(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function liQty_(state, li) {
  var c = byId_(state, li.item);
  if (!c) return li.qtyText || '1';
  return c.pack === 1 ? fmt_(li.packs) + ' ' + c.unit : fmt_(li.packs) + ' × ' + fmt_(c.pack) + ' ' + c.unit;
}
function hash_(vals) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(vals), Utilities.Charset.UTF_8));
}

function writeTable_(name, header, rows, opts) {
  var s = sheet_(name);
  s.clear();
  s.setRightToLeft(true);
  var data = [header].concat(rows);
  s.getRange(1, 1, data.length, header.length).setValues(data);
  s.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#E2F1E8');
  s.setFrozenRows(1);
  if (opts && opts.checkboxCol && rows.length) s.getRange(2, opts.checkboxCol, rows.length, 1).insertCheckboxes();
  if (opts && opts.hideFirst) s.hideColumns(1);
  return s;
}

function writeTabs_(state) {
  var items = active_(state.catalog);
  // קטלוג
  writeTable_('קטלוג', ['מזהה', 'שם', 'קטגוריה', 'יחידה', 'גודל אריזה', 'צריכה שבועית למבוגר', 'מינימום ידני', 'מיקום', 'תוקף ממוצע (ימים)', 'כינויים בחשבוניות', 'מוצר מועדף', 'חלופות', 'מחיר אחרון'],
    items.map(function (c) {
      var p = c.prices && c.prices.length ? c.prices[c.prices.length - 1].p : '';
      return [c.id, c.name, CATS[c.cat] || c.cat, c.unit, c.pack, c.weekly, c.customMin == null ? '' : c.customMin, c.loc, c.shelf || '', (c.aliases || []).join(', '), c.preferred || '', c.alts || '', p];
    }), { hideFirst: true });
  // מלאי – עמודת הכמות ניתנת לעריכה ונקלטת חזרה
  var inv = items.map(function (c) { return [c.id, c.name, c.loc, Number(c.qty) || 0, c.unit, c.expiry || '', c.by || '', c.qu ? new Date(c.qu) : '']; });
  writeTable_('מלאי', ['מזהה', 'פריט', 'מיקום', 'כמות', 'יחידה', 'תוקף', 'עודכן על ידי', 'עודכן'], inv, { hideFirst: true });
  props_().setProperty('invHash', hash_(inv.map(function (r) { return [r[0], r[3]]; })));
  // משפחה
  var roles = { admin: 'מנהל/ת פרופיל', shopper: 'אחראי/ת קניות', member: 'בן/בת בית' };
  writeTable_('משפחה', ['שם', 'תאריך לידה', 'תפקיד', 'מייל'], active_(state.family).map(function (m) { return [m.name, m.birth, roles[m.role] || m.role, m.email || '']; }));
  // חשבוניות
  writeTable_('חשבוניות', ['תאריך', 'חנות', 'סכום', 'סטטוס', 'שורות', 'קובץ מקור'], active_(state.receipts).map(function (r) {
    return [r.date, r.store, r.total, r.status === 'approved' ? 'אושרה' : 'ממתינה', (r.lines || []).length, r.fileUrl || ''];
  }));
  // יומן (300 אחרונים)
  writeTable_('יומן', ['זמן', 'אירוע', 'שינוי', 'משתמש'], (state.log || []).slice(-300).reverse().map(function (l) {
    return [new Date(l.t), l.name + (l.type === 'out' ? ' – נגמר' : ''), l.item ? (fmt_(l.prev) + ' ← ' + fmt_(l.qty)) : '', l.by || ''];
  }));
  // רשימה שבועית – לשונית לכל שבוע, לפי סדר המעברים
  var li = active_(state.list.items).filter(function (i) { return !i.removed; });
  li.sort(function (a, b) { return CAT_ORDER.indexOf(a.cat) - CAT_ORDER.indexOf(b.cat); });
  var rows = li.map(function (i) {
    var c = byId_(state, i.item);
    var p = c && c.prices && c.prices.length ? fmt_(c.prices[c.prices.length - 1].p * (i.packs || 1)) : '';
    return [i.id, !!i.done, i.name, liQty_(state, i), CATS[i.cat] || '', i.src === 'auto' ? 'אוטומטי' : 'ידני', i.reason || '', i.who || '', p, i.note || ''];
  });
  var tab = listTabName_(state);
  var s = writeTable_(tab, ['מזהה', '✓', 'פריט', 'כמות', 'קטגוריה', 'מקור', 'סיבה', 'הוסיף/ה', 'מחיר משוער ₪', 'הערה'], rows, { checkboxCol: 2, hideFirst: true });
  s.getRange(1, 12).setValue('סטטוס: ' + ({ draft: 'טיוטה', approved: 'מאושרת', done: 'הושלמה' }[state.list.status] || ''));
  if (ss_().getSheets()[0].getName() !== tab) { ss_().setActiveSheet(s); ss_().moveActiveSheet(1); }
  props_().setProperty('listTab', tab);
  props_().setProperty('listHash', hash_(readListRows_(s)));
}

function readListRows_(s) {
  if (s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 10).getValues().filter(function (r) { return r[2] !== ''; });
}

/** קליטת עריכות ידניות מהגיליון. מחזיר true אם היה שינוי. */
function absorbSheetEdits_(state) {
  var changed = false, now = Date.now();
  // רשימת קניות
  var tab = props_().getProperty('listTab');
  var s = tab && ss_().getSheetByName(tab);
  if (s && tab === listTabName_(state)) {
    var rows = readListRows_(s);
    if (hash_(rows) !== props_().getProperty('listHash')) {
      var seen = {};
      rows.forEach(function (r) {
        var id = r[0], done = r[1] === true, name = String(r[2]).trim();
        var it = null;
        for (var i = 0; i < state.list.items.length; i++) if (state.list.items[i].id === id) it = state.list.items[i];
        if (it) {
          seen[id] = true;
          if (!!it.done !== done) { it.done = done; it.u = now; changed = true; }
          if (!byId_(state, it.item) && String(r[3]) !== String(it.qtyText || '')) { it.qtyText = String(r[3]); it.u = now; changed = true; }
          if (String(r[9] || '') !== String(it.note || '')) { it.note = String(r[9] || ''); it.u = now; changed = true; }
        } else if (name) {
          var match = null;
          active_(state.catalog).forEach(function (c) { if (c.name === name) match = c; });
          state.list.items.push({ id: 'g' + now.toString(36) + Math.random().toString(36).slice(2, 6), item: match ? match.id : null, name: name, cat: match ? match.cat : 'other', packs: 1, qtyText: String(r[3] || '1'), src: 'manual', who: 'גיליון', done: done, u: now });
          changed = true;
        }
      });
      active_(state.list.items).forEach(function (i) {
        if (!i.removed && !seen[i.id] && i.u < now) { i.removed = true; i.u = now; changed = true; }
      });
      if (changed) state.list.u = now;
    }
  }
  // מלאי – עמודת כמות
  var inv = ss_().getSheetByName('מלאי');
  if (inv && inv.getLastRow() > 1) {
    var vals = inv.getRange(2, 1, inv.getLastRow() - 1, 4).getValues();
    if (hash_(vals.map(function (r) { return [r[0], r[3]]; })) !== props_().getProperty('invHash')) {
      vals.forEach(function (r) {
        var c = byId_(state, r[0]);
        var q = Number(r[3]);
        if (c && !isNaN(q) && q !== Number(c.qty)) {
          state.log.push({ id: 'g' + Math.random().toString(36).slice(2), t: now, item: c.id, name: c.name, type: q === 0 ? 'out' : q < c.qty ? 'use' : 'fix', delta: q - c.qty, prev: c.qty, qty: q, by: 'גיליון', u: now });
          c.qty = q; c.qu = now; c.by = 'גיליון'; c.u = now; changed = true;
        }
      });
    }
  }
  return changed;
}

/* ---------- התראות וארכיון ---------- */
function notify(subject, body) {
  var state = readJson_();
  var emails = active_(state ? state.family : []).map(function (m) { return m.email; }).filter(String);
  if (!emails.length) return false;
  MailApp.sendEmail({ to: emails.join(','), subject: 'מצרכים בבית: ' + subject, htmlBody: '<div dir="rtl" style="font-family:Arial">' + body + '<br><br><a href="' + ScriptApp.getService().getUrl() + '">לפתיחת האפליקציה</a></div>' });
  return true;
}

function saveReceiptImage(dataUrl, name) {
  var folders = DriveApp.getFoldersByName('חשבוניות מצרכים');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('חשבוניות מצרכים');
  var b64 = String(dataUrl).split(',')[1];
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name));
  return file.getUrl();
}

/** טריגר שבועי (אופציונלי): תזכורת במייל ביום הפקת הרשימה. הפעילו פעם אחת את installTrigger. */
function weeklyReminder() {
  var state = readJson_();
  if (!state || !state.settings.notif.ready) return;
  if (new Date().getDay() !== Number(state.settings.listDay)) return;
  notify('הגיע הזמן לרשימה השבועית', 'פתחו את האפליקציה – הרשימה החדשה תחושב אוטומטית לפי המלאי והפרופיל.');
}
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'weeklyReminder') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('weeklyReminder').timeBased().everyDays(1).atHour(6).create();
}
