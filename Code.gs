// ─────────────────────────────────────────────
//  Waze Traffic Archive  |  Code.gs
//  - Accumulates JSON snapshots into a raw_data archive
//  - Auto-prunes data older than 30 days
//  - Filter sidebar drives 6 analysis sheets
// ─────────────────────────────────────────────

var RAW_SHEET    = 'raw_data';
var SOURCES_SHEET = 'מקור';
var FILTER_SHEET  = '_filter';   // hidden, stores last applied filter
var RETENTION_DAYS = 30;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚦 Waze')
    .addItem('פתח סרגל צד...', 'showSidebar')
    .addSeparator()
    .addItem('מחק את כל הנתונים', 'clearAllData')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Waze — בקרה')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── Called from sidebar ──────────────────────
function processWazeJSON(jsonString) {
  var data;
  try { data = JSON.parse(jsonString); }
  catch(e) { return { ok: false, error: 'JSON לא תקין: ' + e.message }; }

  var jams = data.jams || [];
  if (!jams.length) return { ok: false, error: 'הקובץ לא מכיל פקקים (jams)' };

  var snapshotTs = data.startTime ? new Date(data.startTime) : new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Dedupe: reject if a snapshot with the same timestamp already exists
  if (_snapshotExists(ss, snapshotTs)) {
    return {
      ok: false,
      duplicate: true,
      error: 'דגימה זו כבר בארכיון (' +
             Utilities.formatDate(snapshotTs, 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm') +
             ') — לא נוספה שוב',
    };
  }

  try {
    _appendToRawData(ss, jams, snapshotTs);
    _logSource(ss, snapshotTs, jams.length);
    _pruneOld(ss);
    _rebuildAggregation(ss);
  } catch(e) {
    return { ok: false, error: e.message };
  }

  return {
    ok: true,
    ts: Utilities.formatDate(snapshotTs, 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm'),
    jams: jams.length,
    totalSnapshots: _countSnapshots(ss),
  };
}

function _snapshotExists(ss, snapshotTs) {
  var src = ss.getSheetByName(SOURCES_SHEET);
  if (!src || src.getLastRow() < 2) return false;
  var existing = src.getRange(2, 1, src.getLastRow() - 1, 1).getValues();
  var target = snapshotTs.getTime();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i][0] && new Date(existing[i][0]).getTime() === target) return true;
  }
  return false;
}

function applyFilter(filter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _saveFilter(ss, filter);
  try {
    var raw = _readFiltered(ss, filter);
    if (!raw.length) return { ok: false, error: 'אין נתונים בטווח שנבחר' };
    var allRaw = _readFiltered(ss, null);          // baselines use FULL archive
    var baselines = _computeBaselines(allRaw);
    _sheet0_dashboard(ss, raw, filter, baselines);
    _sheet1_summary(ss, raw, filter, baselines);
    _sheet2_timebins(ss, raw, baselines);
    _sheet3_directions(ss, raw);
    _sheet4_anomalies(ss, raw);
    _sheet5_detail(ss, raw);
    _sheet6_legend(ss);
    ss.getSheetByName('🎯 לוח מחוונים').activate();
    return { ok: true, jams: raw.length, archiveJams: allRaw.length };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Build baselines from full archive: per route × dir, and per route × dir × tbin
function _computeBaselines(allRaw) {
  var byRD = {};      // route::dir → { n, sumDelay }
  var byRDT = {};     // route::dir::tbin → { n, sumDelay }
  ROUTES.forEach(function(route) {
    [{ ix:1, streets:route.streets_dir1 },
     { ix:2, streets:route.streets_dir2 }].forEach(function(d) {
      if (!d.streets.length) return;
      var jl = _getJamsForStreets(allRaw, d.streets);
      var k = route.name + '::' + d.ix;
      byRD[k] = { n: jl.length, sumDelay: _sum(jl, 'delay_s') };
      jl.forEach(function(j) {
        var tk = k + '::' + j.tbin;
        var b = byRDT[tk] = byRDT[tk] || { n: 0, sumDelay: 0 };
        b.n++; b.sumDelay += j.delay_s || 0;
      });
    });
  });
  return {
    // Avg delay per jam, route × dir, in seconds. null if no archive samples
    avgPerJam: function(routeName, dirIx) {
      var b = byRD[routeName + '::' + dirIx];
      return b && b.n > 0 ? b.sumDelay / b.n : null;
    },
    // Avg delay per jam, route × dir × tbin. Returns { avg, n } only if n>=3
    avgPerJamForTbin: function(routeName, dirIx, tbin) {
      var b = byRDT[routeName + '::' + dirIx + '::' + tbin];
      return b && b.n >= 3 ? { avg: b.sumDelay / b.n, n: b.n } : null;
    },
  };
}

function getStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName(RAW_SHEET);
  if (!raw || raw.getLastRow() < 2) {
    return { snapshots: 0, jams: 0, dateRange: null, filter: _loadFilter(ss) };
  }
  var sources = ss.getSheetByName(SOURCES_SHEET);
  var snapCount = sources && sources.getLastRow() > 1 ? sources.getLastRow() - 1 : 0;
  var jamCount = raw.getLastRow() - 1;

  // Date range from raw data
  var dates = raw.getRange(2, 1, jamCount, 1).getValues().map(function(r){ return r[0]; });
  var minD = dates.reduce(function(a,b){ return a<b?a:b; }, dates[0]);
  var maxD = dates.reduce(function(a,b){ return a>b?a:b; }, dates[0]);

  return {
    snapshots: snapCount,
    jams: jamCount,
    dateRange: {
      from: Utilities.formatDate(new Date(minD), 'Asia/Jerusalem', 'yyyy-MM-dd'),
      to:   Utilities.formatDate(new Date(maxD), 'Asia/Jerusalem', 'yyyy-MM-dd'),
    },
    filter: _loadFilter(ss),
  };
}

// ─── Fetch JSON from a URL ────────────────────
function fetchFromUrl(url, headersJson) {
  if (!url) return { ok: false, error: 'יש להזין URL' };
  var options = { muteHttpExceptions: true, followRedirects: true };
  if (headersJson && headersJson.trim()) {
    try { options.headers = JSON.parse(headersJson); }
    catch(e) { return { ok: false, error: 'Headers JSON לא תקין: ' + e.message }; }
  }
  var resp;
  try { resp = UrlFetchApp.fetch(url, options); }
  catch(e) { return { ok: false, error: 'שגיאת רשת: ' + e.message }; }

  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    return { ok: false, error: 'HTTP ' + code + ': ' + resp.getContentText().substring(0, 200) };
  }
  return processWazeJSON(resp.getContentText());
}

// ─── Config storage in the spreadsheet itself ──
// Values stored in a hidden _config sheet so they persist across users/devices
function _getConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('_config');
  if (!s) {
    s = ss.insertSheet('_config');
    try { s.hideSheet(); } catch(e) {}
    s.getRange(1, 1, 1, 2).setValues([['key','value']])
     .setBackground('#1F3864').setFontColor('#fff').setFontWeight('bold');
  }
  return s;
}
function _setConfig(key, value) {
  var s = _getConfigSheet();
  var n = s.getLastRow() - 1;
  if (n > 0) {
    var data = s.getRange(2, 1, n, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === key) { s.getRange(i + 2, 2).setValue(value); return; }
    }
  }
  s.appendRow([key, value]);
}
function _getConfig(key) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('_config');
  if (!s || s.getLastRow() < 2) return '';
  var data = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) return String(data[i][1] || '');
  }
  return '';
}
function _currentUserEmail() {
  try { return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || ''; }
  catch(e) { return ''; }
}
function _hasOwnTrigger() {
  return ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === '_scheduledFetch';
  });
}
function _createTrigger(intervalKey) {
  var b = ScriptApp.newTrigger('_scheduledFetch').timeBased();
  if      (intervalKey === '30min') b.everyMinutes(30);
  else if (intervalKey === '1h')    b.everyHours(1);
  else if (intervalKey === '4h')    b.everyHours(4);
  else if (intervalKey === '1d')    b.everyDays(1).atHour(7);
  else throw new Error('תזמון לא מוכר');
  b.create();
}
function _deleteOwnTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === '_scheduledFetch') ScriptApp.deleteTrigger(t);
  });
}

// ─── Schedule auto-fetch ──────────────────────
function setSchedule(url, headersJson, intervalKey) {
  _setConfig('fetch_url', url || '');
  _setConfig('fetch_headers', headersJson || '');
  _setConfig('fetch_interval', intervalKey || 'off');

  _deleteOwnTriggers();

  if (!intervalKey || intervalKey === 'off') {
    _setConfig('trigger_owner', '');
    return { ok: true, scheduled: false };
  }

  try { _createTrigger(intervalKey); }
  catch(e) { return { ok: false, error: e.message }; }

  var email = _currentUserEmail();
  _setConfig('trigger_owner', email);
  return { ok: true, scheduled: true, owner: email };
}

// Re-activate the saved schedule on the current account (e.g., when switching computers/users)
function activateTriggerHere() {
  var interval = _getConfig('fetch_interval');
  if (!interval || interval === 'off') {
    return { ok: false, error: 'אין תזמון מוגדר. הגדר תזמון תחילה ושמור.' };
  }
  _deleteOwnTriggers();
  try { _createTrigger(interval); }
  catch(e) { return { ok: false, error: e.message }; }
  var email = _currentUserEmail();
  _setConfig('trigger_owner', email);
  return { ok: true, owner: email };
}

function _scheduledFetch() {
  var url = _getConfig('fetch_url');
  if (!url) return;
  var headers = _getConfig('fetch_headers') || '';
  var result = fetchFromUrl(url, headers);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var log = ss.getSheetByName('_fetch_log');
    if (!log) {
      log = ss.insertSheet('_fetch_log'); log.hideSheet();
      log.appendRow(['זמן','משתמש','סטטוס','פקקים','שגיאה']);
    }
    log.appendRow([
      new Date(),
      _currentUserEmail(),
      result.ok ? 'הצלחה' : (result.duplicate ? 'כפילות' : 'שגיאה'),
      result.jams || '',
      result.error || '',
    ]);
  } catch(e) {}
}

function getScheduleConfig() {
  return {
    url:           _getConfig('fetch_url'),
    headers:       _getConfig('fetch_headers'),
    interval:      _getConfig('fetch_interval') || 'off',
    owner:         _getConfig('trigger_owner'),
    currentUser:   _currentUserEmail(),
    hasTriggerHere: _hasOwnTrigger(),
  };
}

function clearAllData() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('מחיקה מלאה',
    'למחוק את כל הנתונים והגיליונות? לא ניתן לבטל.',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return { ok: false };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var keep = ss.insertSheet('_tmp_' + Date.now());
  ss.getSheets().forEach(function(s) {
    if (s.getName() !== keep.getName()) ss.deleteSheet(s);
  });
  keep.setName('Sheet1');
  return { ok: true };
}

// ═══ DATA LAYER ═══════════════════════════════

var RAW_COLS = ['snapshot_ts','pub_ts','date','day','hour','tbin',
                'street','city','level','length_m','delay_s','speed_kmh',
                'start_node','end_node','tt_min'];

function _ensureRawSheet(ss) {
  var s = ss.getSheetByName(RAW_SHEET);
  if (!s) {
    s = ss.insertSheet(RAW_SHEET);
    s.getRange(1, 1, 1, RAW_COLS.length).setValues([RAW_COLS])
     .setBackground('#1F3864').setFontColor('#fff').setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function _ensureSourcesSheet(ss) {
  var s = ss.getSheetByName(SOURCES_SHEET);
  if (!s) {
    s = ss.insertSheet(SOURCES_SHEET);
    try { s.setRightToLeft(true); } catch(e) {}
    s.getRange(1, 1, 1, 4).setValues([['חותמת זמן','תאריך','שעה','מס\' פקקים']])
     .setBackground('#1F3864').setFontColor('#fff').setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function _appendToRawData(ss, jams, snapshotTs) {
  var s = _ensureRawSheet(ss);
  var days = ['שני','שלישי','רביעי','חמישי','שישי','שבת','ראשון'];
  var rows = jams.map(function(j) {
    var pm = j.pubMillis || snapshotTs.getTime();
    var dt = new Date(pm);
    var spd = j.speedKMH || 0;
    var ln  = j.length   || 0;
    var hour = dt.getHours();
    return [
      snapshotTs,                                       // snapshot_ts (Date)
      dt,                                                // pub_ts
      Utilities.formatDate(dt, 'Asia/Jerusalem', 'yyyy-MM-dd'),
      days[dt.getDay() === 0 ? 6 : dt.getDay() - 1],
      hour,
      _tbin(hour),
      j.street || '', j.city || '',
      j.level || 0, ln,
      j.delay || 0, spd,
      j.startNode || '', j.endNode || '',
      spd > 0 ? Math.round((ln / 1000) / spd * 60 * 100) / 100 : '',
    ];
  });
  if (rows.length) {
    s.getRange(s.getLastRow() + 1, 1, rows.length, RAW_COLS.length).setValues(rows);
  }
}

function _logSource(ss, snapshotTs, jamCount) {
  var s = _ensureSourcesSheet(ss);
  s.appendRow([
    snapshotTs,
    Utilities.formatDate(snapshotTs, 'Asia/Jerusalem', 'yyyy-MM-dd'),
    Utilities.formatDate(snapshotTs, 'Asia/Jerusalem', 'HH:mm'),
    jamCount,
  ]);
}

function _pruneOld(ss) {
  var cutoff = new Date(Date.now() - RETENTION_DAYS * 86400000);
  // Prune raw_data
  var raw = ss.getSheetByName(RAW_SHEET);
  if (raw && raw.getLastRow() > 1) {
    var data = raw.getRange(2, 1, raw.getLastRow() - 1, 1).getValues();
    var firstKeep = -1;
    for (var i = 0; i < data.length; i++) {
      if (new Date(data[i][0]) >= cutoff) { firstKeep = i; break; }
    }
    if (firstKeep > 0) {
      raw.deleteRows(2, firstKeep);
    } else if (firstKeep === -1) {
      // all old
      raw.deleteRows(2, data.length);
    }
  }
  // Prune sources
  var src = ss.getSheetByName(SOURCES_SHEET);
  if (src && src.getLastRow() > 1) {
    var data2 = src.getRange(2, 1, src.getLastRow() - 1, 1).getValues();
    var firstKeep2 = -1;
    for (var j = 0; j < data2.length; j++) {
      if (new Date(data2[j][0]) >= cutoff) { firstKeep2 = j; break; }
    }
    if (firstKeep2 > 0) src.deleteRows(2, firstKeep2);
    else if (firstKeep2 === -1) src.deleteRows(2, data2.length);
  }
}

function _countSnapshots(ss) {
  var s = ss.getSheetByName(SOURCES_SHEET);
  return s && s.getLastRow() > 1 ? s.getLastRow() - 1 : 0;
}

function _saveFilter(ss, filter) {
  var s = ss.getSheetByName(FILTER_SHEET);
  if (!s) { s = ss.insertSheet(FILTER_SHEET); s.hideSheet(); }
  s.getRange('A1').setValue(JSON.stringify(filter));
}

function _loadFilter(ss) {
  var s = ss.getSheetByName(FILTER_SHEET);
  if (!s) return null;
  try { return JSON.parse(s.getRange('A1').getValue()); } catch(e) { return null; }
}

// Read raw_data, apply filter, return array of jam-objects
function _readFiltered(ss, filter) {
  var s = ss.getSheetByName(RAW_SHEET);
  if (!s || s.getLastRow() < 2) return [];
  var values = s.getRange(2, 1, s.getLastRow() - 1, RAW_COLS.length).getValues();

  var fromD = filter && filter.fromDate ? new Date(filter.fromDate) : null;
  var toD   = filter && filter.toDate   ? new Date(filter.toDate + 'T23:59:59') : null;
  var fromH = filter && typeof filter.fromHour === 'number' ? filter.fromHour : 0;
  var toH   = filter && typeof filter.toHour   === 'number' ? filter.toHour   : 23;
  var days  = filter && filter.days && filter.days.length ? filter.days : null;

  return values.filter(function(r) {
    var pubTs = new Date(r[1]);
    if (fromD && pubTs < fromD) return false;
    if (toD && pubTs > toD) return false;
    if (r[4] < fromH || r[4] > toH) return false;
    if (days && days.indexOf(r[3]) === -1) return false;
    return true;
  }).map(function(r) {
    return {
      snapshot_ts: r[0], pub_ts: r[1],
      date: r[2], day: r[3], hour: r[4], tbin: r[5],
      street: r[6], city: r[7], level: r[8],
      length_m: r[9], delay_s: r[10], speed: r[11],
      sn: r[12], en: r[13], tt_min: r[14],
    };
  });
}

// ═══ ROUTES ═══════════════════════════════════
var ROUTES = [
  {section:'דרום',name:'כביש 10',from:'ניצנה',to:'אילת',distance_km:200,free_flow_min:140,streets_dir1:['10'],streets_dir2:[],dir1_label:'שני הכיוונים',dir2_label:''},
  {section:'דרום',name:'כביש 25',from:'באר שבע',to:'שדרות',distance_km:35,free_flow_min:30,streets_dir1:['25 מזרח'],streets_dir2:['25 מערב'],dir1_label:'שדרות → באר שבע (מזרח)',dir2_label:'באר שבע → שדרות (מערב)'},
  {section:'דרום',name:'כביש 40',from:'באר שבע',to:'תל אביב',distance_km:110,free_flow_min:70,streets_dir1:['40 צפון'],streets_dir2:['40 דרום'],dir1_label:'באר שבע → ת"א (צפון)',dir2_label:'ת"א → באר שבע (דרום)'},
  {section:'מרכז',name:'כביש 1',from:'תל אביב',to:'ירושלים',distance_km:62,free_flow_min:45,streets_dir1:['1 מזרח'],streets_dir2:['1 מערב'],dir1_label:'תל אביב → ירושלים',dir2_label:'ירושלים → תל אביב'},
  {section:'מרכז',name:'כביש 4',from:'אשדוד',to:'חיפה',distance_km:120,free_flow_min:80,streets_dir1:['4 צפון'],streets_dir2:['4 דרום'],dir1_label:'אשדוד → חיפה (צפון)',dir2_label:'חיפה → אשדוד (דרום)'},
  {section:'מרכז',name:'כביש 5',from:'ת"א',to:'אריאל',distance_km:40,free_flow_min:30,streets_dir1:['5 מזרח'],streets_dir2:['5 מערב'],dir1_label:'ת"א → אריאל (מזרח)',dir2_label:'אריאל → ת"א (מערב)'},
  {section:'מרכז',name:'כביש 6',from:'קריית גת',to:'חדרה',distance_km:130,free_flow_min:65,streets_dir1:['6 צפון'],streets_dir2:['6 דרום'],dir1_label:'ק. גת → חדרה (צפון)',dir2_label:'חדרה → ק. גת (דרום)'},
  {section:'מרכז',name:'כביש 41',from:'אשדוד',to:'ראשל"צ',distance_km:20,free_flow_min:18,streets_dir1:['41 מזרח'],streets_dir2:['41 מערב'],dir1_label:'אשדוד → ראשל"צ (מזרח)',dir2_label:'ראשל"צ → אשדוד (מערב)'},
  {section:'מרכז',name:'כביש 44',from:'אשדוד',to:'מודיעין',distance_km:40,free_flow_min:35,streets_dir1:['44 צפון'],streets_dir2:['44 דרום'],dir1_label:'אשדוד → מודיעין (צפון)',dir2_label:'מודיעין → אשדוד (דרום)'},
  {section:'מרכז',name:'כביש 60',from:'באר שבע',to:'נצרת',distance_km:200,free_flow_min:180,streets_dir1:['60'],streets_dir2:[],dir1_label:'שני הכיוונים (מעורב)',dir2_label:''},
  {section:'מרכז',name:'כביש 444',from:'ראש העין',to:'נחשונים',distance_km:15,free_flow_min:12,streets_dir1:['444 צפון'],streets_dir2:['444 דרום'],dir1_label:'נחשונים → ר"ע (צפון)',dir2_label:'ר"ע → נחשונים (דרום)'},
  {section:'מרכז',name:'כביש 461',from:'אור יהודה',to:'יהוד',distance_km:8,free_flow_min:8,streets_dir1:['461 מזרח'],streets_dir2:['461 מערב'],dir1_label:'אור יהודה → יהוד (מזרח)',dir2_label:'יהוד → אור יהודה (מערב)'},
  {section:'צפון',name:'כביש 22',from:'חיפה מפרץ',to:'חיפה כרמל',distance_km:12,free_flow_min:12,streets_dir1:['22 צפון'],streets_dir2:['22 דרום'],dir1_label:'דרום → צפון',dir2_label:'צפון → דרום'},
  {section:'צפון',name:'כביש 57',from:'נתניה',to:'טול כרם',distance_km:22,free_flow_min:20,streets_dir1:['57 מזרח'],streets_dir2:['57 מערב'],dir1_label:'נתניה → מזרח',dir2_label:'מזרח → נתניה'},
  {section:'צפון',name:'כביש 65',from:'חדרה',to:'עפולה',distance_km:55,free_flow_min:45,streets_dir1:['65 מזרח'],streets_dir2:['65 מערב'],dir1_label:'חדרה → עפולה (מזרח)',dir2_label:'עפולה → חדרה (מערב)'},
  {section:'צפון',name:'כביש 66',from:'מגידו',to:'עפולה',distance_km:20,free_flow_min:18,streets_dir1:['66 צפון'],streets_dir2:['66 דרום'],dir1_label:'מגידו → עפולה (צפון)',dir2_label:'עפולה → מגידו (דרום)'},
  {section:'צפון',name:'כביש 70',from:'חיפה',to:'עכו',distance_km:25,free_flow_min:20,streets_dir1:['70 צפון'],streets_dir2:['70 דרום'],dir1_label:'חיפה → עכו (צפון)',dir2_label:'עכו → חיפה (דרום)'},
  {section:'צפון',name:'כביש 75',from:'חיפה',to:'נצרת',distance_km:30,free_flow_min:30,streets_dir1:['75'],streets_dir2:[],dir1_label:'שני הכיוונים',dir2_label:''},
  {section:'צפון',name:'כביש 89',from:'צפת',to:'נהריה',distance_km:42,free_flow_min:40,streets_dir1:['89'],streets_dir2:[],dir1_label:'שני הכיוונים (מעורב)',dir2_label:''},
  {section:'צפון',name:'כביש 90',from:'אילת',to:'מטולה',distance_km:480,free_flow_min:360,streets_dir1:['90'],streets_dir2:[],dir1_label:'שני הכיוונים (מעורב)',dir2_label:''},
  {section:'צפון',name:'כביש 98',from:'תל קציר',to:'מבוא חמה (גולן)',distance_km:25,free_flow_min:25,streets_dir1:['98'],streets_dir2:[],dir1_label:'שני הכיוונים',dir2_label:''},
  {section:'צפון',name:'כביש 781',from:'כפר ביאליק',to:'קריות',distance_km:5,free_flow_min:6,streets_dir1:['781'],streets_dir2:[],dir1_label:'שני הכיוונים',dir2_label:''},
  {section:'צפון',name:'כביש 804',from:'ראס אל-עין',to:'פקיעין',distance_km:12,free_flow_min:15,streets_dir1:['804'],streets_dir2:[],dir1_label:'שני הכיוונים',dir2_label:''},
  {section:'צפון',name:'כביש 807',from:'מגדל',to:'רביד (כינרת)',distance_km:10,free_flow_min:12,streets_dir1:['807'],streets_dir2:[],dir1_label:'שני הכיוונים',dir2_label:''},
  {section:'צפון',name:'כביש 866',from:'כרמיאל',to:'צפת',distance_km:30,free_flow_min:35,streets_dir1:['866'],streets_dir2:[],dir1_label:'שני הכיוונים (מעורב)',dir2_label:''},
  {section:'צפון',name:'כביש 899',from:'יפתח',to:'גבול לבנון',distance_km:8,free_flow_min:10,streets_dir1:['899'],streets_dir2:[],dir1_label:'שני הכיוונים',dir2_label:''},
  {section:'צפון',name:'כביש 8655',from:'פקיעין',to:'פקיעין חדשה',distance_km:5,free_flow_min:6,streets_dir1:['8655'],streets_dir2:[],dir1_label:'שני הכיוונים',dir2_label:''},
  {section:'צפון',name:'כביש 8697',from:'רמות (גולן)',to:'רמות',distance_km:5,free_flow_min:5,streets_dir1:['8697'],streets_dir2:[],dir1_label:'שני הכיוונים',dir2_label:''},
];

// ═══ STYLING ══════════════════════════════════
var C_HDR_BG = '#1F3864', C_HDR_FG = '#FFFFFF';
var C_SEC_BG = '#D6E4F0', C_SEC_FG = '#1F3864';
var C_ALT    = '#F2F7FB';
var C_RED_BG = '#FCE4EC', C_RED_FG = '#C00000';
var C_ORG_BG = '#FFF3E0', C_ORG_FG = '#BF6000';
var C_YEL_BG = '#FFF8E1', C_YEL_FG = '#996600';
var C_GRN_BG = '#E8F5E9', C_GRN_FG = '#006100';

function _newSheet(ss, name, tabColor) {
  var s = ss.getSheetByName(name);
  if (s) { s.clear(); s.clearFormats(); }
  else   { s = ss.insertSheet(name); }
  s.setTabColor(tabColor);
  try { s.setRightToLeft(true); } catch(e) {}
  return s;
}

function _hdrRow(sheet, cols, row) {
  row = row || 1;
  var range = sheet.getRange(row, 1, 1, cols.length);
  range.setValues([cols])
       .setBackground(C_HDR_BG).setFontColor(C_HDR_FG)
       .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
}
function _dataRow(sheet, row, vals, alt) {
  var r = sheet.getRange(row, 1, 1, vals.length);
  r.setValues([vals]).setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center');
  if (alt) r.setBackground(C_ALT);
}
function _secRow(sheet, row, ncols, text) {
  sheet.getRange(row, 1, 1, ncols).merge().setValue(text)
       .setBackground(C_SEC_BG).setFontColor(C_SEC_FG)
       .setFontWeight('bold').setHorizontalAlignment('right');
}
function _colorEst(sheet, row, col, est, ff) {
  if (typeof est !== 'number' || ff <= 0) return;
  var ratio = est / ff;
  var cell = sheet.getRange(row, col);
  if      (ratio > 1.5)  cell.setBackground(C_RED_BG).setFontColor(C_RED_FG).setFontWeight('bold');
  else if (ratio > 1.2)  cell.setBackground(C_ORG_BG).setFontColor(C_ORG_FG).setFontWeight('bold');
  else if (ratio > 1.05) cell.setBackground(C_YEL_BG).setFontColor(C_YEL_FG).setFontWeight('bold');
  else                   cell.setBackground(C_GRN_BG).setFontColor(C_GRN_FG).setFontWeight('bold');
}
function _colorStatus(sheet, row, col, st) {
  var cell = sheet.getRange(row, col);
  if      (st === 'חריג מאוד') cell.setBackground(C_RED_BG).setFontColor(C_RED_FG).setFontWeight('bold');
  else if (st === 'עמוס')      cell.setBackground(C_ORG_BG).setFontColor(C_ORG_FG).setFontWeight('bold');
  else if (st === 'מתון')      cell.setBackground(C_YEL_BG).setFontColor(C_YEL_FG);
  else if (st === 'תקין')      cell.setBackground(C_GRN_BG).setFontColor(C_GRN_FG);
}
function _autoWidth(sheet, nc) {
  for (var c = 1; c <= nc; c++) sheet.autoResizeColumn(c);
}

// Returns { label, pct } — % deviation of current avg-per-jam vs historical avg-per-jam
function _deviationLabel(currentJams, histAvgSec) {
  if (!currentJams.length) return { label: '— (אין פקקים)', pct: undefined };
  if (histAvgSec == null || histAvgSec === 0) return { label: '— (אין היסטוריה)', pct: undefined };
  var curAvg = _sum(currentJams, 'delay_s') / currentJams.length;
  var pct = _round1((curAvg - histAvgSec) / histAvgSec * 100);
  return { label: (pct >= 0 ? '+' : '') + pct + '%', pct: pct };
}

function _colorDeviation(sheet, row, col, pct) {
  if (typeof pct !== 'number') return;
  var cell = sheet.getRange(row, col);
  if      (pct > 50)  cell.setBackground(C_RED_BG).setFontColor(C_RED_FG).setFontWeight('bold');
  else if (pct > 25)  cell.setBackground(C_ORG_BG).setFontColor(C_ORG_FG).setFontWeight('bold');
  else if (pct > 10)  cell.setBackground(C_YEL_BG).setFontColor(C_YEL_FG).setFontWeight('bold');
  else if (pct >= -10) cell.setBackground(C_GRN_BG).setFontColor(C_GRN_FG).setFontWeight('bold');
  else                cell.setBackground('#DBEAFE').setFontColor('#1E40AF').setFontWeight('bold');
}

function _getJamsForStreets(jams, streets) {
  if (!streets || !streets.length) return [];
  return jams.filter(function(j) { return streets.indexOf(j.street) !== -1; });
}
function _sum(arr, key) { return arr.reduce(function(s, v) { return s + (v[key] || 0); }, 0); }
function _mean(arr) { return arr.length ? arr.reduce(function(s,v){return s+v;},0) / arr.length : 0; }
function _stdev(arr) {
  if (arr.length < 2) return 0;
  var m = _mean(arr);
  return Math.sqrt(arr.reduce(function(s,x){return s+Math.pow(x-m,2);},0) / (arr.length-1));
}
function _tbin(h) { var b = Math.floor(h/4)*4; return _pad(b)+':00-'+_pad(b+4)+':00'; }
function _pad(n) { return n < 10 ? '0'+n : ''+n; }
function _round1(v) { return Math.round(v*10)/10; }

// ═══ ANALYSIS SHEETS ══════════════════════════

function _filterLabel(filter) {
  if (!filter) return 'כל הנתונים';
  var p = [];
  if (filter.fromDate || filter.toDate) p.push((filter.fromDate||'…') + ' עד ' + (filter.toDate||'…'));
  if (filter.fromHour !== undefined && filter.toHour !== undefined)
    p.push('שעות ' + _pad(filter.fromHour) + ':00–' + _pad(filter.toHour) + ':59');
  if (filter.days && filter.days.length) p.push('ימים: ' + filter.days.join(', '));
  return p.length ? p.join(' | ') : 'כל הנתונים';
}

// ─── Sheet 0: Dashboard (one-pager TL;DR) ─────
function _sheet0_dashboard(ss, jams, filter, baselines) {
  var ws = _newSheet(ss, '🎯 לוח מחוונים', '#0EA5E9');

  // Build per-route × dir summary with status + deviation
  var summary = [];
  ROUTES.forEach(function(route) {
    [{ label: route.dir1_label, streets: route.streets_dir1, ix: 1 },
     { label: route.dir2_label, streets: route.streets_dir2, ix: 2 }].forEach(function(d) {
      if (!d.label || !d.streets.length) return;
      var jl = _getJamsForStreets(jams, d.streets);
      if (!jl.length) return;
      var curAvg = _sum(jl, 'delay_s') / jl.length;
      var genBase = baselines && baselines.avgPerJam(route.name, d.ix);
      var devPct, src, status;
      if (genBase != null && genBase > 0) {
        devPct = _round1((curAvg - genBase) / genBase * 100);
        src = 'היסטורי';
      } else {
        var totalMin = _sum(jl, 'delay_s') / 60;
        devPct = route.free_flow_min > 0 ? _round1((totalMin / route.free_flow_min) * 100) : 0;
        src = 'תיאורטי';
      }
      status = devPct > 50 ? 'חריג מאוד' :
               devPct > 25 ? 'עמוס' :
               devPct > 10 ? 'מתון' :
               devPct < -10 ? 'טוב מהרגיל' : 'תקין';
      summary.push({
        section: route.section, route: route.name, dir: d.label,
        jams: jl.length, devPct: devPct, status: status, src: src,
        totalDelayMin: _round1(_sum(jl, 'delay_s') / 60),
      });
    });
  });

  var ncW = 6;

  // ─── Title ───
  ws.getRange(1, 1, 1, ncW).merge()
    .setValue('🎯 לוח מחוונים — מצב התנועה')
    .setFontSize(20).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#1F3864').setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  ws.setRowHeight(1, 48);

  ws.getRange(2, 1, 1, ncW).merge()
    .setValue('מסנן: ' + _filterLabel(filter) +
              '  •  ' + jams.length + ' פקקים בטווח  •  ' +
              summary.length + ' מסלולים פעילים')
    .setFontSize(11).setFontColor('#475569')
    .setBackground('#F1F5F9').setHorizontalAlignment('center');

  // ─── KPI strip (status counts) ───
  var counts = { 'תקין':0,'מתון':0,'עמוס':0,'חריג מאוד':0,'טוב מהרגיל':0 };
  summary.forEach(function(r) { counts[r.status]++; });

  var kpis = [
    { label: '🟢 תקין',         count: counts['תקין'],        color: '#10B981' },
    { label: '🟡 מתון',         count: counts['מתון'],        color: '#F59E0B' },
    { label: '🟠 עמוס',         count: counts['עמוס'],        color: '#F97316' },
    { label: '🔴 חריג מאוד',    count: counts['חריג מאוד'],   color: '#EF4444' },
    { label: '🔵 טוב מהרגיל',   count: counts['טוב מהרגיל'],  color: '#3B82F6' },
  ];

  for (var i = 0; i < kpis.length; i++) {
    var col = i + 1;
    ws.getRange(4, col).setValue(kpis[i].label)
      .setBackground(kpis[i].color).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(5, col).setValue(kpis[i].count)
      .setBackground('#FFFFFF').setFontColor(kpis[i].color)
      .setFontWeight('bold').setFontSize(28)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  ws.setRowHeight(4, 32);
  ws.setRowHeight(5, 64);

  // Empty 6th column when fewer than 6 KPIs
  if (kpis.length < ncW) {
    for (var k = kpis.length + 1; k <= ncW; k++) {
      ws.getRange(4, k).setBackground('#F1F5F9');
      ws.getRange(5, k).setBackground('#F1F5F9');
    }
  }

  var row = 7;

  // ─── Top 10 worst routes ───
  ws.getRange(row, 1, 1, ncW).merge()
    .setValue('📍 10 המסלולים הגרועים ביותר')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#1F3864').setFontColor('#FFFFFF')
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  ws.setRowHeight(row, 28); row++;

  ws.getRange(row, 1, 1, ncW).setValues([['#','מסלול','כיוון','אזור','פקקים','סטייה']])
    .setBackground('#475569').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  row++;

  var top10 = summary.slice().sort(function(a, b) { return b.devPct - a.devPct; }).slice(0, 10);
  top10.forEach(function(r, i) {
    var dev = (r.devPct >= 0 ? '+' : '') + r.devPct + '%';
    var rng = ws.getRange(row, 1, 1, ncW)
      .setValues([[i + 1, r.route, r.dir, r.section, r.jams, dev]])
      .setFontSize(11).setHorizontalAlignment('center');
    if (i % 2 === 0) rng.setBackground('#F8FAFC');
    _colorDeviation(ws, row, 6, r.devPct);
    row++;
  });

  row++;

  // ─── Best 5 routes (least delayed) ───
  ws.getRange(row, 1, 1, ncW).merge()
    .setValue('✅ 5 המסלולים הטובים ביותר')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#10B981').setFontColor('#FFFFFF')
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  ws.setRowHeight(row, 28); row++;

  ws.getRange(row, 1, 1, ncW).setValues([['#','מסלול','כיוון','אזור','פקקים','סטייה']])
    .setBackground('#475569').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  row++;

  var best5 = summary.slice().sort(function(a, b) { return a.devPct - b.devPct; }).slice(0, 5);
  best5.forEach(function(r, i) {
    var dev = (r.devPct >= 0 ? '+' : '') + r.devPct + '%';
    var rng = ws.getRange(row, 1, 1, ncW)
      .setValues([[i + 1, r.route, r.dir, r.section, r.jams, dev]])
      .setFontSize(11).setHorizontalAlignment('center');
    if (i % 2 === 0) rng.setBackground('#F8FAFC');
    _colorDeviation(ws, row, 6, r.devPct);
    row++;
  });

  row++;

  // ─── Section breakdown ───
  ws.getRange(row, 1, 1, ncW).merge()
    .setValue('🗺️ פילוח לפי אזור')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#1F3864').setFontColor('#FFFFFF')
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  ws.setRowHeight(row, 28); row++;

  ws.getRange(row, 1, 1, ncW)
    .setValues([['אזור','🟢 תקין','🟡 מתון','🟠 עמוס','🔴 חריג','🔵 טוב מהרגיל']])
    .setBackground('#475569').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  row++;

  ['דרום','מרכז','צפון'].forEach(function(sec) {
    var c = { 'תקין':0,'מתון':0,'עמוס':0,'חריג מאוד':0,'טוב מהרגיל':0 };
    summary.filter(function(r){ return r.section === sec; }).forEach(function(r) { c[r.status]++; });
    ws.getRange(row, 1, 1, ncW)
      .setValues([[sec, c['תקין'], c['מתון'], c['עמוס'], c['חריג מאוד'], c['טוב מהרגיל']]])
      .setFontSize(11).setHorizontalAlignment('center');
    ws.getRange(row, 2).setBackground('#E8F5E9').setFontColor('#006100').setFontWeight('bold');
    ws.getRange(row, 3).setBackground('#FFF8E1').setFontColor('#996600').setFontWeight('bold');
    ws.getRange(row, 4).setBackground('#FFF3E0').setFontColor('#BF6000').setFontWeight('bold');
    ws.getRange(row, 5).setBackground('#FCE4EC').setFontColor('#C00000').setFontWeight('bold');
    ws.getRange(row, 6).setBackground('#DBEAFE').setFontColor('#1E40AF').setFontWeight('bold');
    row++;
  });

  row++;

  // ─── Quality of analysis ───
  var hist = summary.filter(function(r){ return r.src === 'היסטורי'; }).length;
  var theo = summary.filter(function(r){ return r.src === 'תיאורטי'; }).length;
  ws.getRange(row, 1, 1, ncW).merge()
    .setValue('📊 איכות הניתוח: ' + hist + ' מסלולים מבוססים על היסטוריה  •  ' +
              theo + ' מסלולים על בסיס תיאורטי' +
              (theo > 0 ? '  ⚠️ העלה עוד דגימות לדיוק' : ''))
    .setFontSize(10).setFontColor('#475569')
    .setBackground('#F8FAFC').setHorizontalAlignment('center')
    .setFontStyle('italic');

  // Column widths
  ws.setColumnWidth(1, 50);
  ws.setColumnWidth(2, 130);
  ws.setColumnWidth(3, 200);
  ws.setColumnWidth(4, 80);
  ws.setColumnWidth(5, 80);
  ws.setColumnWidth(6, 100);
}

function _sheet1_summary(ss, jams, filter, baselines) {
  var ws = _newSheet(ss, 'סיכום מסלולים', '#1F3864');
  var cols = ['אזור','מסלול','מוצא','יעד','מרחק (ק"מ)','זמן חופשי (דק\')',
              'כיוון 1','כיוון 2','פקקים כ1','פקקים כ2',
              'השהיה כ1 (דק\')','השהיה כ2 (דק\')','זמן משוער כ1','זמן משוער כ2',
              'סטייה כ1 vs ממוצע כללי','סטייה כ2 vs ממוצע כללי'];
  var nc = cols.length;

  ws.getRange(1,1,1,nc).merge()
    .setValue('ניתוח זמני נסיעה — סיכום אגרגטיבי').setFontSize(15)
    .setFontWeight('bold').setFontColor('#1F3864').setHorizontalAlignment('center');
  ws.getRange(2,1,1,nc).merge()
    .setValue('מסנן: ' + _filterLabel(filter) + ' | פקקים: ' + jams.length)
    .setFontSize(10).setFontColor('#666').setHorizontalAlignment('center');

  _hdrRow(ws, cols, 4);

  var row = 5, curSec = '';
  ROUTES.forEach(function(route) {
    if (route.section !== curSec) {
      curSec = route.section;
      _secRow(ws, row, nc, ({דרום:'🔽 דרום',מרכז:'🔶 מרכז',צפון:'🔼 צפון'}[curSec]||curSec));
      row++;
    }
    var j1 = _getJamsForStreets(jams, route.streets_dir1);
    var j2 = _getJamsForStreets(jams, route.streets_dir2);
    var d1 = _round1(_sum(j1, 'delay_s') / 60);
    var d2 = _round1(_sum(j2, 'delay_s') / 60);
    var e1 = _round1(route.free_flow_min + d1);
    var e2 = route.dir2_label ? _round1(route.free_flow_min + d2) : '—';

    // Per-jam current vs historical
    var dev1 = _deviationLabel(j1, baselines && baselines.avgPerJam(route.name, 1));
    var dev2 = route.dir2_label
      ? _deviationLabel(j2, baselines && baselines.avgPerJam(route.name, 2))
      : '—';

    var vals = [route.section, route.name, route.from, route.to,
                route.distance_km, route.free_flow_min,
                route.dir1_label, route.dir2_label || '—',
                j1.length, route.dir2_label ? j2.length : '—',
                d1, route.dir2_label ? d2 : '—', e1, e2,
                dev1.label, dev2.label || dev2];
    _dataRow(ws, row, vals, row % 2 === 0);
    _colorEst(ws, row, 13, e1, route.free_flow_min);
    if (typeof e2 === 'number') _colorEst(ws, row, 14, e2, route.free_flow_min);
    _colorDeviation(ws, row, 15, dev1.pct);
    if (dev2 && dev2.pct !== undefined) _colorDeviation(ws, row, 16, dev2.pct);
    row++;
  });
  ws.setFrozenRows(4);
  _autoWidth(ws, nc);
}

function _sheet2_timebins(ss, jams, baselines) {
  var ws = _newSheet(ss, 'פירוט לפי מרווח זמן', '#2E75B6');
  var TIME_BINS = ['00:00-04:00','04:00-08:00','08:00-12:00','12:00-16:00','16:00-20:00','20:00-24:00'];
  var cols = ['אזור','מסלול','כיוון','מרווח זמן','מס\' פקקים','אורך (ק"מ)',
              'השהיה (דק\')','מהירות ממוצעת','רמת פקק ממוצעת',
              'השהיה לפקק (דק\')','היסטורי באותה שעה','היסטורי כללי',
              'סטייה משעה','סטייה מכללי','מקור השוואה','סטטוס'];
  _hdrRow(ws, cols, 1);

  var row = 2;
  ROUTES.forEach(function(route) {
    [{label:route.dir1_label, streets:route.streets_dir1, ix:1},
     {label:route.dir2_label, streets:route.streets_dir2, ix:2}].forEach(function(dir) {
      if (!dir.label || !dir.streets.length) return;
      var jl = _getJamsForStreets(jams, dir.streets);
      if (!jl.length) return;
      TIME_BINS.forEach(function(tb) {
        var jj = jl.filter(function(j){return j.tbin === tb;});
        if (!jj.length) return;
        var tl = _round1(_sum(jj,'length_m')/1000);
        var td = _round1(_sum(jj.filter(function(j){return j.delay_s>0;}),'delay_s')/60);
        var spds = jj.filter(function(j){return j.speed>0;}).map(function(j){return j.speed;});
        var asp = spds.length ? _round1(_mean(spds)) : 0;
        var alv = _round1(_mean(jj.map(function(j){return j.level;})));

        var curAvg = jj.length ? _sum(jj, 'delay_s') / jj.length : 0;
        var hourBase = baselines && baselines.avgPerJamForTbin(route.name, dir.ix, tb);
        var genBase  = baselines && baselines.avgPerJam(route.name, dir.ix);

        var hourAvgMin = hourBase ? _round1(hourBase.avg / 60) : '—';
        var genAvgMin  = (genBase != null) ? _round1(genBase / 60) : '—';
        var devHour = (hourBase && hourBase.avg > 0)
          ? _round1((curAvg - hourBase.avg) / hourBase.avg * 100) : null;
        var devGen  = (genBase != null && genBase > 0)
          ? _round1((curAvg - genBase) / genBase * 100) : null;

        // Smart status: prefer hour-baseline; fallback to general; fallback to free-flow vs total delay
        var st, source;
        if (devHour !== null) {
          st = devHour>50?'חריג מאוד':devHour>25?'עמוס':devHour>10?'מתון':'תקין';
          source = 'שעה (n='+hourBase.n+')';
        } else if (devGen !== null) {
          st = devGen>50?'חריג מאוד':devGen>25?'עמוס':devGen>10?'מתון':'תקין';
          source = 'כללי';
        } else {
          var pa = route.free_flow_min > 0 ? Math.round(td/route.free_flow_min*1000)/10 : 0;
          st = pa>50?'חריג מאוד':pa>25?'עמוס':pa>10?'מתון':'תקין';
          source = 'תיאורטי';
        }

        _dataRow(ws, row, [route.section, route.name, dir.label, tb,
                           jj.length, tl, td, asp, alv,
                           _round1(curAvg/60), hourAvgMin, genAvgMin,
                           devHour!==null?devHour+'%':'—',
                           devGen!==null?devGen+'%':'—',
                           source, st],
                 row % 2 === 0);
        if (devHour !== null) _colorDeviation(ws, row, 13, devHour);
        if (devGen  !== null) _colorDeviation(ws, row, 14, devGen);
        _colorStatus(ws, row, 16, st);
        row++;
      });
    });
  });
  ws.setFrozenRows(1);
  _autoWidth(ws, cols.length);
}

function _sheet3_directions(ss, jams) {
  var ws = _newSheet(ss, 'השוואת כיוונים', '#ED7D31');
  var cols = ['אזור','מסלול','מרחק','זמן חופשי','כיוון 1','כיוון 2',
              'השהיה כ1','השהיה כ2','זמן כ1','זמן כ2','הפרש','כיוון עמוס','יחס'];
  _hdrRow(ws, cols, 1);
  var row = 2;
  ROUTES.forEach(function(route) {
    if (!route.dir2_label) return;
    var j1 = _getJamsForStreets(jams, route.streets_dir1);
    var j2 = _getJamsForStreets(jams, route.streets_dir2);
    var d1 = _round1(_sum(j1,'delay_s')/60);
    var d2 = _round1(_sum(j2,'delay_s')/60);
    var e1 = _round1(route.free_flow_min + d1);
    var e2 = _round1(route.free_flow_min + d2);
    var diff = _round1(Math.abs(e1-e2));
    var busier = e1>e2 ? route.dir1_label : route.dir2_label;
    var minE = Math.min(e1,e2);
    var ratio = minE>0 ? Math.round((Math.max(e1,e2)/minE)*100)/100 : 1;
    _dataRow(ws, row, [route.section, route.name, route.distance_km, route.free_flow_min,
                       route.dir1_label, route.dir2_label, d1, d2, e1, e2, diff, busier, ratio+'x'],
             row % 2 === 0);
    var bc = e1 > e2 ? 9 : 10;
    ws.getRange(row, bc).setBackground(C_RED_BG).setFontColor(C_RED_FG).setFontWeight('bold');
    row++;
  });
  ws.setFrozenRows(1);
  _autoWidth(ws, cols.length);
}

function _sheet4_anomalies(ss, jams) {
  var ws = _newSheet(ss, 'חריגות', '#C00000');
  var cols = ['#','אזור','מסלול','כיוון','קטע','עיר','תאריך','יום','שעה','מרווח',
              'מהירות','אורך (מ\')','השהיה (דק\')','ממוצע (דק\')','חריגה %','רמת פקק','חומרה'];
  _hdrRow(ws, cols, 1);
  var anoms = [];
  ROUTES.forEach(function(route) {
    [{label:route.dir1_label,streets:route.streets_dir1},
     {label:route.dir2_label,streets:route.streets_dir2}].forEach(function(dir) {
      if (!dir.label || !dir.streets.length) return;
      var jl = _getJamsForStreets(jams, dir.streets);
      if (!jl.length) return;
      var dels = jl.filter(function(j){return j.delay_s>0;}).map(function(j){return j.delay_s;});
      if (!dels.length) return;
      var avgd = _mean(dels), stdd = _stdev(dels);
      var thr = stdd>0 ? avgd+1.5*stdd : avgd*2;
      jl.forEach(function(j) {
        if (j.delay_s > thr || j.speed < 5 || j.level >= 4) {
          var dp = avgd>0 ? Math.round(((j.delay_s-avgd)/avgd)*1000)/10 : 0;
          var sv = (j.speed<3 || dp>200) ? 'קריטי' : (j.level>=4 || dp>100) ? 'גבוה' : 'בינוני';
          anoms.push({sec:route.section, rt:route.name, dir:dir.label,
            seg:j.sn+' → '+j.en, city:j.city, date:j.date, day:j.day,
            hour:j.hour!==null?_pad(j.hour)+':00':'', tb:j.tbin,
            spd:j.speed, ln:j.length_m,
            dm:_round1(j.delay_s/60), am:_round1(avgd/60),
            dp:dp, lv:j.level, sv:sv});
        }
      });
    });
  });
  anoms.sort(function(a,b){
    var sv = {קריטי:3,גבוה:2,בינוני:1};
    return (sv[b.sv]||0) - (sv[a.sv]||0) || b.dm - a.dm;
  });
  var row = 2;
  anoms.forEach(function(a, i) {
    _dataRow(ws, row, [i+1, a.sec, a.rt, a.dir, a.seg, a.city, a.date, a.day, a.hour, a.tb,
                       a.spd, a.ln, a.dm, a.am, a.dp+'%', a.lv, a.sv],
             row % 2 === 0);
    _colorStatus(ws, row, 17, a.sv==='קריטי'?'חריג מאוד':a.sv==='גבוה'?'עמוס':'מתון');
    row++;
  });
  ws.setFrozenRows(1);
  _autoWidth(ws, cols.length);
}

function _sheet5_detail(ss, jams) {
  var ws = _newSheet(ss, 'פירוט פקקים', '#548235');
  var cols = ['#','אזור','מסלול','כיוון','קטע','עיר','תאריך','יום','שעה','מרווח',
              'מהירות','אורך (מ\')','השהיה (שנ\')','זמן נסיעה (דק\')','רמת פקק'];
  _hdrRow(ws, cols, 1);
  var row = 2, idx = 1;
  ROUTES.forEach(function(route) {
    [{label:route.dir1_label,streets:route.streets_dir1},
     {label:route.dir2_label,streets:route.streets_dir2}].forEach(function(dir) {
      if (!dir.label || !dir.streets.length) return;
      var jl = _getJamsForStreets(jams, dir.streets);
      jl.sort(function(a,b){return new Date(a.pub_ts) - new Date(b.pub_ts);});
      jl.forEach(function(j) {
        _dataRow(ws, row, [idx++, route.section, route.name, dir.label,
                           j.sn+' → '+j.en, j.city, j.date, j.day,
                           j.hour!==null?_pad(j.hour)+':00':'', j.tbin,
                           j.speed, j.length_m, j.delay_s, j.tt_min, j.level],
                 row % 2 === 0);
        row++;
      });
    });
  });
  ws.setFrozenRows(1);
  _autoWidth(ws, cols.length);
}

function _sheet6_legend(ss) {
  var ws = _newSheet(ss, 'מקרא', '#666666');
  var legend = [
    ['שדה','הסבר'],
    ['זמן חופשי','זמן נסיעה תיאורטי ללא עומסים (קבוע, לא תלוי שעה)'],
    ['השהיה מצטברת','סך ההשהיות מכלל הפקקים במסלול בטווח שנבחר'],
    ['השהיה לפקק','השהיה ממוצעת לפקק יחיד בטווח הנבחר (דקות)'],
    ['היסטורי באותה שעה','ממוצע השהיה לפקק עבור אותו מסלול × כיוון × מרווח שעות, מכל הארכיון'],
    ['היסטורי כללי','ממוצע השהיה לפקק עבור אותו מסלול × כיוון בכל שעות היום, מכל הארכיון'],
    ['סטייה משעה','(נוכחי − היסטורי שעה) ÷ היסטורי שעה × 100. נדרשות לפחות 3 דגימות לאותה שעה'],
    ['סטייה מכללי','(נוכחי − היסטורי כללי) ÷ היסטורי כללי × 100. עובד גם עם דגימה אחת בארכיון'],
    ['מקור השוואה','שעה — היה מספיק היסטוריה. כללי — fallback. תיאורטי — אין היסטוריה כלל'],
    ['רמת פקק','1=זרימה, 2=מתון, 3=בינוני, 4=כבד, 5=עצירה'],
    ['',''],
    ['סטטוס','קריטריונים (לפי מקור ההשוואה הזמין)'],
    ['תקין','סטייה עד 10%'],
    ['מתון','10%–25%'],
    ['עמוס','25%–50%'],
    ['חריג מאוד','מעל 50%'],
    ['',''],
    ['קוד צבעים בעמודות סטייה',''],
    ['ירוק','בין 10%- ל-+10% (תקין)'],
    ['צהוב','+10% עד +25%'],
    ['כתום','+25% עד +50%'],
    ['אדום','מעל +50%'],
    ['כחול','מתחת ל-10%- (פחות עומס מהרגיל)'],
  ];
  legend.forEach(function(r, i) {
    var range = ws.getRange(i+1, 1, 1, 2);
    range.setValues([r]);
    if (i === 0) range.setBackground(C_HDR_BG).setFontColor(C_HDR_FG).setFontWeight('bold');
    else ws.getRange(i+1, 1).setFontWeight('bold');
  });
  _autoWidth(ws, 2);
}

// ═══ AGGREGATION (always uses ALL raw data) ═══
function _rebuildAggregation(ss) {
  var raw = _readFiltered(ss, null);
  if (!raw.length) return;

  var ws = _newSheet(ss, 'אגרגציה לאורך זמן', '#7030A0');
  ws.getRange(1,1,1,8).merge()
    .setValue('מגמות לאורך זמן (כל הנתונים בארכיון)')
    .setFontSize(15).setFontWeight('bold').setFontColor('#1F3864').setHorizontalAlignment('center');

  // Section A: Avg delay by route × hour
  var cols = ['מסלול','כיוון','שעה','מס\' דגימות','השהיה ממוצעת (דק\')','מהירות ממוצעת','רמה ממוצעת'];
  _hdrRow(ws, cols, 3);
  var row = 4;
  ROUTES.forEach(function(route) {
    [{label:route.dir1_label, streets:route.streets_dir1},
     {label:route.dir2_label, streets:route.streets_dir2}].forEach(function(dir) {
      if (!dir.label || !dir.streets.length) return;
      var rj = _getJamsForStreets(raw, dir.streets);
      if (!rj.length) return;
      var byHour = {};
      rj.forEach(function(j) {
        var h = j.hour;
        (byHour[h] = byHour[h] || []).push(j);
      });
      Object.keys(byHour).sort(function(a,b){return +a-+b;}).forEach(function(h) {
        var arr = byHour[h];
        var avgD = _round1(_mean(arr.map(function(j){return j.delay_s;}))/60);
        var avgS = _round1(_mean(arr.filter(function(j){return j.speed>0;}).map(function(j){return j.speed;})));
        var avgL = _round1(_mean(arr.map(function(j){return j.level;})));
        _dataRow(ws, row, [route.name, dir.label, _pad(+h)+':00',
                           arr.length, avgD, avgS, avgL], row%2===0);
        _colorEst(ws, row, 5, route.free_flow_min + avgD, route.free_flow_min);
        row++;
      });
    });
  });
  ws.setFrozenRows(3);
  _autoWidth(ws, cols.length);
}
