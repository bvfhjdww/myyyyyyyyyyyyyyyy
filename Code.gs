/**
 * دفتر ملاحظات HSE — SEDRA · Google Sheets backend (v3)
 * ------------------------------------------------------------------
 * 1) افتح الشيت ← Extensions ← Apps Script ← احذف الكود القديم والصق هذا الكود ← Save
 * 2) اختر الدالة setup من الأعلى واضغط Run مرة وحدة ← وافق على الصلاحيات (Sheets + Drive)
 * 3) عندك نشر قديم؟  Deploy ← Manage deployments ← ✎ Edit ← Version: New version ← Deploy  (الرابط يبقى نفسه)
 *    أول مرة؟        Deploy ← New deployment ← Web app ← Execute as: Me ← Who has access: Anyone ← Deploy
 * الكود يضيف تبويبات: Observations · Custom Templates · Summary — ولا يلمس تبويبات Villas / Templates القديمة.
 */
var SHEET_ID = '1GVG6mwqdIGwz3Friq5JOcIc94LC7kMtJ02wIWjZ2JZE';
var APP = 'sedra-hse';
var VERSION = 3;
var OBS_TAB = 'Observations';
var TPL_TAB = 'Custom Templates';
var SUM_TAB = 'Summary';
var PHOTO_FOLDER = 'SEDRA HSE Photos';

var OBS_COLS = [
  ['id', 'Ref'], ['date', 'Date'], ['time', 'Time'], ['phase', 'Phase'], ['area', 'Area'], ['loc', 'Location'],
  ['category', 'Category'], ['risk', 'Risk'], ['text', 'Observation'], ['ar', 'Arabic summary'],
  ['status', 'Status'], ['closedAt', 'Closed on'], ['closeNote', 'Close-out note'], ['repeated', 'Repeated'],
  ['photos', 'Photos'], ['preview', 'Preview'], ['gps', 'GPS'], ['items', 'Template IDs'],
  ['device', 'Device'], ['createdAt', 'Created'], ['updatedAt', 'Updated']
];
var TPL_COLS = [['id', 'ID'], ['cat', 'Category'], ['risk', 'Risk'], ['ar', 'Arabic'], ['np', 'English'],
  ['act', 'Action'], ['hz', 'Hazards'], ['updatedAt', 'Updated'], ['deleted', 'Deleted']];
var RISK_NAME = {H: 'High', M: 'Medium', L: 'Low', G: 'Good practice'};
var RISK_CODE = {'High': 'H', 'Medium': 'M', 'Low': 'L', 'Good practice': 'G'};

/* ---------------- entry points ---------------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    var ss = ss_();
    ensure_(ss);
    var base = {ok: true, app: APP, version: VERSION, sheetName: ss.getName(), sheetUrl: ss.getUrl()};
    if (p.action === 'list') {
      base.observations = listObs_(ss, p.since || '', Number(p.limit) || 1500);
      base.templates = listTpl_(ss);
    } else {
      base.time = new Date().toISOString();
    }
    return out_(base);
  } catch (err) {
    return out_({ok: false, app: APP, error: msg_(err)});
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (x) { return out_({ok: false, app: APP, error: 'Sheet busy, retrying'}); }
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!body.ops) return out_({ok: false, app: APP, legacy: true, error: 'Old page version. Please update the HSE page.'});
    var ss = ss_();
    var sh = ensure_(ss);
    var results = body.ops.map(function (op) {
      try {
        if (op.op === 'upsertObs') return upsertObs_(sh, op.obs);
        if (op.op === 'deleteObs') return deleteObs_(sh, op.id);
        if (op.op === 'photo') return photo_(sh, op);
        if (op.op === 'upsertTpl') return upsertTpl_(ss, op.tpl);
        return {ok: false, error: 'Unknown operation'};
      } catch (err) {
        return {ok: false, error: msg_(err)};
      }
    });
    try { refreshSummary_(ss); } catch (x) {}
    return out_({ok: true, app: APP, version: VERSION, results: results});
  } catch (err) {
    return out_({ok: false, app: APP, error: msg_(err)});
  } finally {
    lock.releaseLock();
  }
}

/** شغّلها مرة وحدة من المحرر لتجهيز التبويبات وأخذ الصلاحيات */
function setup() {
  var ss = ss_();
  ensure_(ss);
  folder_();
  refreshSummary_(ss);
  Logger.log('Ready: ' + ss.getUrl());
  return 'Ready';
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('HSE')
      .addItem('تحديث الملخص', 'refreshSummary')
      .addItem('تجهيز / إصلاح التبويبات', 'setup')
      .addToUi();
  } catch (x) {}
}

function refreshSummary() { refreshSummary_(ss_()); }

/** أي تعديل يدوي في تبويب Observations يحدّث وقت التعديل عشان ينتقل للجوال */
function onEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() !== OBS_TAB) return;
    var r0 = Math.max(2, e.range.getRow());
    var r1 = e.range.getRow() + e.range.getNumRows() - 1;
    if (r1 < 2) return;
    var rows = r1 - r0 + 1;
    var updCol = col_(OBS_COLS, 'updatedAt') + 1;
    if (e.range.getColumn() === updCol && e.range.getNumColumns() === 1) return;
    var now = new Date().toISOString();
    var upd = [];
    for (var i = 0; i < rows; i++) upd.push([now]);
    sh.getRange(r0, updCol, rows, 1).setNumberFormat('@').setValues(upd);
    var stCol = col_(OBS_COLS, 'status') + 1, clCol = col_(OBS_COLS, 'closedAt') + 1;
    if (e.range.getColumn() <= stCol && e.range.getLastColumn() >= stCol) {
      var tz = sh.getParent().getSpreadsheetTimeZone();
      var st = sh.getRange(r0, stCol, rows, 1).getValues();
      var cl = sh.getRange(r0, clCol, rows, 1).getValues();
      for (var j = 0; j < rows; j++) {
        if (String(st[j][0]) === 'Closed' && !cl[j][0]) cl[j][0] = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
        if (String(st[j][0]) === 'Open') cl[j][0] = '';
      }
      sh.getRange(r0, clCol, rows, 1).setNumberFormat('@').setValues(cl);
    }
  } catch (x) {}
}

/* ---------------- sheets ---------------- */
function ss_() {
  try {
    var a = SpreadsheetApp.getActiveSpreadsheet();
    if (a) return a;
  } catch (x) {}
  return SpreadsheetApp.openById(SHEET_ID);
}

function ensure_(ss) {
  var sh = ss.getSheetByName(OBS_TAB);
  if (!sh) {
    sh = ss.insertSheet(OBS_TAB, 0);
    var n = OBS_COLS.length;
    sh.getRange(1, 1, 1, n).setValues([OBS_COLS.map(function (c) { return c[1]; })])
      .setFontWeight('bold').setBackground('#1b1d20').setFontColor('#ffffff').setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 30);
    var widths = {id: 150, date: 92, time: 56, phase: 56, area: 56, loc: 90, category: 190, risk: 84, text: 440, ar: 300,
      status: 80, closedAt: 128, closeNote: 220, repeated: 74, photos: 210, preview: 130, gps: 190, items: 120,
      device: 80, createdAt: 160, updatedAt: 160};
    Object.keys(widths).forEach(function (k) { sh.setColumnWidth(col_(OBS_COLS, k) + 1, widths[k]); });
    var max = sh.getMaxRows();
    sh.getRange(2, 1, max - 1, n).setVerticalAlignment('top');
    sh.getRange(2, col_(OBS_COLS, 'text') + 1, max - 1, 2).setWrap(true);
    sh.getRange(2, col_(OBS_COLS, 'closeNote') + 1, max - 1, 1).setWrap(true);
    var stRange = sh.getRange(2, col_(OBS_COLS, 'status') + 1, max - 1, 1);
    stRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Open', 'Closed'], true).setAllowInvalid(true).build());
    var rkRange = sh.getRange(2, col_(OBS_COLS, 'risk') + 1, max - 1, 1);
    var rule = function (txt, bg, fg, rng) {
      return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(txt).setBackground(bg).setFontColor(fg).setRanges([rng]).build();
    };
    sh.setConditionalFormatRules([
      rule('High', '#f8d7d4', '#b3261e', rkRange), rule('Medium', '#fde5cc', '#9a4f00', rkRange),
      rule('Low', '#d6e7f8', '#1d5fa3', rkRange), rule('Good practice', '#d4f0e1', '#1b7a4b', rkRange),
      rule('Open', '#fff2bf', '#5c4a00', stRange), rule('Closed', '#d4f0e1', '#1b7a4b', stRange)
    ]);
  }
  var t = ss.getSheetByName(TPL_TAB);
  if (!t) {
    t = ss.insertSheet(TPL_TAB);
    t.getRange(1, 1, 1, TPL_COLS.length).setValues([TPL_COLS.map(function (c) { return c[1]; })])
      .setFontWeight('bold').setBackground('#1b1d20').setFontColor('#ffffff');
    t.setFrozenRows(1);
    t.getRange(1, 1, t.getMaxRows(), TPL_COLS.length).setNumberFormat('@');
    t.setColumnWidth(4, 240); t.setColumnWidth(5, 340); t.setColumnWidth(6, 340);
  }
  if (!ss.getSheetByName(SUM_TAB)) ss.insertSheet(SUM_TAB, 1);
  return sh;
}

function col_(cols, key) {
  for (var i = 0; i < cols.length; i++) if (cols[i][0] === key) return i;
  return -1;
}

function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (!id || last < 2) return 0;
  var hit = sh.getRange(2, 1, last - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function ensureRows_(sh, row) {
  var max = sh.getMaxRows();
  if (row > max) sh.insertRowsAfter(max, Math.max(200, row - max));
}

function mergeLinks_(cur, add) {
  var list = String(cur || '').split(/\s+/).filter(function (x) { return x; });
  (add || []).forEach(function (u) { if (u && list.indexOf(u) < 0) list.push(u); });
  return list.join('\n');
}

/* ---------------- observations ---------------- */
function upsertObs_(sh, o) {
  if (!o || !o.id) return {ok: false, error: 'Missing observation id'};
  var n = OBS_COLS.length;
  var row = findRow_(sh, o.id);
  var cur = row ? sh.getRange(row, 1, 1, n).getValues()[0] : null;
  if (cur) {
    var curUpd = String(cur[col_(OBS_COLS, 'updatedAt')] || '');
    if (curUpd && o.updatedAt && curUpd > String(o.updatedAt)) return {ok: true, id: o.id, skipped: 'newer-in-sheet'};
  } else {
    row = sh.getLastRow() + 1;
    ensureRows_(sh, row);
  }
  var v = {
    id: o.id, date: o.date || '', time: o.time || '', phase: o.phase || '', area: o.area || '', loc: o.loc || '',
    category: o.catNames || '', risk: RISK_NAME[o.risk] || o.risk || '', text: o.text || '', ar: o.ar || '',
    status: o.status === 'closed' ? 'Closed' : 'Open', closedAt: o.closedAt || '', closeNote: o.closeNote || '',
    repeated: o.repeated ? 'Yes' : '', photos: mergeLinks_(cur ? cur[col_(OBS_COLS, 'photos')] : '', o.photoUrls),
    gps: o.gps ? 'https://maps.google.com/?q=' + o.gps : '', items: (o.items || []).join(','),
    device: o.device || '', createdAt: o.createdAt || '', updatedAt: o.updatedAt || new Date().toISOString()
  };
  var pv = col_(OBS_COLS, 'preview');
  var left = OBS_COLS.slice(0, pv).map(function (c) { return v[c[0]]; });
  var right = OBS_COLS.slice(pv + 1).map(function (c) { return v[c[0]]; });
  // نص صريح (@) عشان الشيت ما يحوّل التاريخ/الوقت/رقم المرحلة تلقائيًا
  sh.getRange(row, 1, 1, left.length).setNumberFormat('@').setValues([left]);
  sh.getRange(row, pv + 2, 1, right.length).setNumberFormat('@').setValues([right]);
  return {ok: true, id: o.id, row: row};
}

function deleteObs_(sh, id) {
  var row = findRow_(sh, id);
  if (row) sh.deleteRow(row);
  return {ok: true, id: id, deleted: !!row};
}

function photo_(sh, op) {
  if (!op || !op.id || !op.data) return {ok: false, error: 'Empty photo'};
  var folder = folder_();
  var name = op.name || (op.id + '_' + (op.idx || 1) + '.jpg');
  var file;
  var it = folder.getFilesByName(name);
  if (it.hasNext()) {
    file = it.next();
  } else {
    file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(op.data), op.mime || 'image/jpeg', name));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (x) {}
  }
  var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';
  var row = findRow_(sh, op.id);
  if (row) {
    var pc = col_(OBS_COLS, 'photos') + 1;
    sh.getRange(row, pc).setValue(mergeLinks_(sh.getRange(row, pc).getValue(), [url]));
    var prev = sh.getRange(row, col_(OBS_COLS, 'preview') + 1);
    if (!prev.getFormula()) {
      prev.setFormula('=IMAGE("https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400")');
      sh.setRowHeight(row, 96);
    }
  }
  return {ok: true, id: op.id, url: url};
}

function folder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (x) {}
  }
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  var f = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
  props.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}

function listObs_(ss, since, limit) {
  var sh = ss.getSheetByName(OBS_TAB);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, OBS_COLS.length).getValues();
  var tz = ss.getSpreadsheetTimeZone();
  var out = [];
  for (var i = vals.length - 1; i >= 0 && out.length < limit; i--) {
    var r = vals[i], o = {};
    for (var c = 0; c < OBS_COLS.length; c++) {
      var val = r[c];
      if (val instanceof Date) val = Utilities.formatDate(val, tz, OBS_COLS[c][0] === 'time' ? 'HH:mm' : 'yyyy-MM-dd');
      o[OBS_COLS[c][0]] = val === null || val === undefined ? '' : String(val);
    }
    if (!o.id) continue;
    if (since && o.date < since) continue;
    out.push({
      id: o.id, date: o.date, time: o.time, phase: o.phase, area: o.area, loc: o.loc,
      category: o.category, risk: RISK_CODE[o.risk] || 'M', text: o.text, ar: o.ar,
      status: o.status.toLowerCase() === 'closed' ? 'closed' : 'open', closedAt: o.closedAt, closeNote: o.closeNote,
      repeated: o.repeated.toLowerCase() === 'yes',
      photoUrls: o.photos.split(/\s+/).filter(function (x) { return x; }),
      gps: o.gps.replace('https://maps.google.com/?q=', ''),
      items: o.items.split(',').filter(function (x) { return x; }),
      device: o.device, createdAt: o.createdAt, updatedAt: o.updatedAt
    });
  }
  return out;
}

/* ---------------- templates ---------------- */
function upsertTpl_(ss, t) {
  if (!t || !t.id) return {ok: false, error: 'Missing template id'};
  var sh = ss.getSheetByName(TPL_TAB);
  var n = TPL_COLS.length;
  var row = findRow_(sh, t.id);
  if (row) {
    var cur = sh.getRange(row, 1, 1, n).getValues()[0];
    if (String(cur[7] || '') > String(t.updatedAt || '')) return {ok: true, id: t.id, skipped: 'newer-in-sheet'};
  } else {
    row = sh.getLastRow() + 1;
    ensureRows_(sh, row);
  }
  sh.getRange(row, 1, 1, n).setNumberFormat('@').setValues([[
    t.id, t.cat || '', t.risk || '', t.ar || '', t.np || '', t.act || '', t.hz || '',
    t.updatedAt || new Date().toISOString(), t.deleted ? 'Yes' : ''
  ]]);
  return {ok: true, id: t.id};
}

function listTpl_(ss) {
  var sh = ss.getSheetByName(TPL_TAB);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, TPL_COLS.length).getValues().map(function (r) {
    return {id: String(r[0]), cat: String(r[1]), risk: String(r[2]), ar: String(r[3]), np: String(r[4]), act: String(r[5]),
      hz: String(r[6]), updatedAt: String(r[7]), deleted: String(r[8]).toLowerCase() === 'yes'};
  }).filter(function (t) { return t.id; });
}

/* ---------------- summary tab ---------------- */
function refreshSummary_(ss) {
  var sh = ss.getSheetByName(OBS_TAB);
  var sum = ss.getSheetByName(SUM_TAB) || ss.insertSheet(SUM_TAB);
  var tz = ss.getSpreadsheetTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var last = sh.getLastRow();
  var vals = last > 1 ? sh.getRange(2, 1, last - 1, OBS_COLS.length).getValues() : [];
  var iD = col_(OBS_COLS, 'date'), iA = col_(OBS_COLS, 'area'), iC = col_(OBS_COLS, 'category'),
      iR = col_(OBS_COLS, 'risk'), iS = col_(OBS_COLS, 'status');
  var all = {n: 0, open: 0, H: 0}, tod = {n: 0, open: 0, H: 0}, cats = {}, areas = {}, days = {};
  vals.forEach(function (r) {
    if (!r[0]) return;
    var d = r[iD] instanceof Date ? Utilities.formatDate(r[iD], tz, 'yyyy-MM-dd') : String(r[iD]);
    var open = String(r[iS]) !== 'Closed', high = String(r[iR]) === 'High';
    all.n++; if (open) all.open++; if (high) all.H++;
    if (d === today) { tod.n++; if (open) tod.open++; if (high) tod.H++; }
    String(r[iC] || 'Other').split(' / ').forEach(function (c) { bump_(cats, c || 'Other', open); });
    bump_(areas, String(r[iA] || '—'), open);
    bump_(days, d, open);
  });
  var rate = function (o) { return o.n ? Math.round((o.n - o.open) / o.n * 100) + '%' : '—'; };
  var out = [
    ['SEDRA HSE — Observation Register', '', ''],
    ['Updated ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'), '', ''],
    ['', '', ''],
    ['Metric', 'All time', 'Today'],
    ['Total observations', all.n, tod.n],
    ['Open', all.open, tod.open],
    ['Closed', all.n - all.open, tod.n - tod.open],
    ['Close-out rate', rate(all), rate(tod)],
    ['High risk', all.H, tod.H],
    ['', '', ''],
    ['Category', 'Total', 'Open']
  ];
  sorted_(cats).forEach(function (k) { out.push([k, cats[k].n, cats[k].open]); });
  out.push(['', '', ''], ['Area', 'Total', 'Open']);
  sorted_(areas).forEach(function (k) { out.push([k, areas[k].n, areas[k].open]); });
  out.push(['', '', ''], ['Date (last 14 days)', 'Total', 'Open']);
  Object.keys(days).sort().reverse().slice(0, 14).forEach(function (k) { out.push([k, days[k].n, days[k].open]); });
  sum.clear();
  sum.getRange(1, 1, out.length, 3).setValues(out);
  sum.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  sum.getRange(2, 1).setFontColor('#666666');
  out.forEach(function (r, i) {
    if (r[0] === 'Metric' || r[0] === 'Category' || r[0] === 'Area' || r[0] === 'Date (last 14 days)') {
      sum.getRange(i + 1, 1, 1, 3).setFontWeight('bold').setBackground('#1b1d20').setFontColor('#ffffff');
    }
  });
  sum.setColumnWidth(1, 300);
  sum.getRange(1, 2, out.length, 2).setHorizontalAlignment('center');
}

function bump_(m, k, open) {
  if (!m[k]) m[k] = {n: 0, open: 0};
  m[k].n++;
  if (open) m[k].open++;
}

function sorted_(m) {
  return Object.keys(m).sort(function (a, b) { return m[b].n - m[a].n; });
}

/* ---------------- helpers ---------------- */
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function msg_(err) {
  return String((err && err.message) || err);
}
