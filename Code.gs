// ─────────────────────────────────────────────
//  Waze Traffic Archive  |  Code.gs
//  - Accumulates JSON snapshots into a raw_data archive
//  - Auto-prunes data older than 30 days
//  - Filter sidebar drives 6 analysis sheets
// ─────────────────────────────────────────────

var RAW_SHEET    = 'raw_data';
var SOURCES_SHEET = 'מקור';
var FILTER_SHEET  = '_filter';   // hidden, stores last applied filter
var BASELINE_ARCHIVE = '_baseline_archive';   // hidden, permanent aggregated counters
var RETENTION_DAYS = 30;
// Safety cap: Google Sheets enforces ~10M cells per spreadsheet. With 18 RAW_COLS
// that's ~555K rows before raw_data alone hits the wall — BUT Sheets allocates
// a small column buffer (typically 2 trailing empty cols), pushing the effective
// width to 20. To leave breathing room for the buffer + ~200K cells used by the
// other sheets, trigger emergency prune well below the theoretical max.
// At 20 cols/row: 450K × 20 = 9M cells, leaving ~1M cells of headroom.
var MAX_RAW_ROWS = 450000;

// Resumable archive job: bounded work per execution, resumed via time trigger.
// Smaller chunks + tighter budget = more frequent checkpoints and a safer exit
// well before Apps Script's 6-min wall. A single chunk includes a Drive.createFile
// call, which can take 30-90s — must leave room for one full chunk after the budget
// check fires.
var ARCHIVE_CHUNK_ROWS    = 2000;
var ARCHIVE_TIME_BUDGET_MS = 3 * 60 * 1000;
var ARCHIVE_JOB_PROP       = 'archiveJob';
var ARCHIVE_TRIGGER_FN     = '_continueArchiveJob';
var ARCHIVE_LAST_TS_PROP   = 'archiveLastTs';   // ISO ts of latest snapshot already archived
var ARCHIVE_LOG_SHEET      = 'ארכיון נפרד';

// Daily auto-archive: runs exportFullArchive (incremental) + _pruneOld every day at 03:00.
var DAILY_ARCHIVE_TRIGGER_FN  = '_dailyArchive';
var DAILY_ARCHIVE_CONFIG_KEY  = 'daily_archive';
// Legacy weekly trigger handler — kept ONLY so stale triggers don't error out;
// migration in onOpen removes them.
var LEGACY_WEEKLY_TRIGGER_FN  = '_weeklyArchive';

// Self-healing fetch: when an import fails because raw_data hit the cell cap,
// auto-trigger a prune and reschedule the fetch to retry after a short delay.
// Bounded retry budget with exponential backoff — gives up after MAX attempts
// so a stuck recovery (prune can't free cells) can't loop forever.
var FETCH_RETRY_TRIGGER_FN    = '_retryScheduledFetch';
var FETCH_RETRY_COUNT_PROP    = 'fetchRetryCount';
var FETCH_RETRY_MAX_ATTEMPTS  = 5;
var FETCH_RETRY_BACKOFF_MS    = [2, 5, 10, 20, 30].map(function(m){ return m * 60 * 1000; });

// Background monitor: every 6 hours, check whether an archive job exists in
// PropertiesService without a continuation trigger and resume it. Eliminates
// the "stuck job" pattern that previously required clicking "המשך ייצוא תקוע".
var JOB_MONITOR_TRIGGER_FN    = '_monitorArchiveJob';

// Auto-chained prune: fires once after a 'manual' archive completes if
// raw_data is still over the cap. Lets manual export → prune happen with no
// user clicks in between.
var POST_ARCHIVE_PRUNE_FN     = '_postArchivePrune';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var advanced = ui.createMenu('מתקדם')
    .addItem('💾 ייצא ארכיון מלא ל-Drive', 'menuExportArchiveFull')
    .addItem('▶️ המשך ייצוא תקוע', 'menuResumeArchiveJob')
    .addItem('🛑 בטל ייצוא תקוע', 'menuCancelArchiveJob')
    .addItem('🔄 העבר נתונים קיימים לארכיון אגרגטיבי', 'menuMigrateToArchive')
    .addSeparator()
    .addItem('📊 בדוק שימוש בתאים', 'menuDiagnoseCellUsage')
    .addItem('🗜️ דחס תאים ריקים', 'menuCompactSheets');

  ui.createMenu('🚦 Waze')
    .addItem('פתח סרגל צד...', 'showSidebar')
    .addSeparator()
    .addItem('📥 ייצא חדשים ל-Drive', 'menuExportArchive')
    .addItem('🗓️ הפעל/כבה גיבוי יומי', 'menuToggleDailyArchive')
    .addItem('🧹 קצץ raw_data עכשיו', 'menuPruneNow')
    .addSeparator()
    .addSubMenu(advanced)
    .addSeparator()
    .addItem('מחק את כל הנתונים', 'clearAllData')
    .addToUi();

  // Migration: remove any legacy weekly-archive triggers from before the
  // switch to daily. Safe no-op if none exist.
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === LEGACY_WEEKLY_TRIGGER_FN) {
        ScriptApp.deleteTrigger(t);
      }
    });
  } catch(_) {}

  // Auto-install daily archive trigger on first open; skipped if user disabled.
  try {
    if (_getConfig(DAILY_ARCHIVE_CONFIG_KEY) !== 'off' && !_hasDailyArchiveTrigger()) {
      _installDailyArchiveTrigger();
      _setConfig(DAILY_ARCHIVE_CONFIG_KEY, 'on');
    }
  } catch(_) {}

  // Auto-install the archive-job monitor (resumes orphaned chunked jobs).
  try {
    if (!_hasJobMonitorTrigger()) _installJobMonitorTrigger();
  } catch(_) {}
}

// Manually run the next slice of a stuck archive job. Use when the auto-trigger
// failed to fire (trigger quota exhausted, silent failure, etc.).
function menuResumeArchiveJob() {
  var ui = SpreadsheetApp.getUi();
  var job = _getArchiveJob();
  if (!job) {
    ui.alert('אין ייצוא פעיל', 'לא נמצא ג\'וב ייצוא פעיל לחידוש.', ui.ButtonSet.OK);
    return;
  }
  // Watchdog: if this synchronous slice is hard-killed at the 6-min wall before
  // the budget-exit code can schedule a continuation, this trigger revives it
  // at +7 min. Clean exits (done / budget exit) overwrite it.
  _scheduleArchiveWatchdog();
  try {
    var r = _runArchiveSlice();
    if (r && r.done) {
      ui.alert('הצלחה',
        'הייצוא הושלם!\n\nשם קובץ: ' + r.filename + '\nמספר פקקים: ' + r.count + '\n\nקישור: ' + r.url,
        ui.ButtonSet.OK);
    } else {
      var prog = (r && r.progress) ? r.progress : { processed: '?', total: '?' };
      ui.alert('המשך ריצה',
        'בוצע סלייס נוסף.\n\nהתקדמות: ' + prog.processed + ' מתוך ' + prog.total + '\n\n' +
        'אם הריצה עוצרת שוב, פתח שוב את התפריט וסמן "▶️ המשך ייצוא תקוע".',
        ui.ButtonSet.OK);
    }
  } catch(e) {
    ui.alert('שגיאה', 'הריצה נכשלה: ' + e.message + '\n\nאפשר לנסות שוב, או לבטל את הג\'וב.', ui.ButtonSet.OK);
  }
}

// Abort a stuck archive job — clears state and removes triggers. Partial CSV files in Drive remain.
function menuCancelArchiveJob() {
  var ui = SpreadsheetApp.getUi();
  var job = _getArchiveJob();
  if (!job) {
    ui.alert('אין ייצוא פעיל', 'לא נמצא ג\'וב ייצוא פעיל לביטול.', ui.ButtonSet.OK);
    return;
  }
  var resp = ui.alert('ביטול ייצוא',
    'לבטל את הג\'וב? נכון לעכשיו: ' + (job.nextRow - 1) + ' מתוך ' + job.totalRows + ' שורות עובדו.\n\n' +
    'ה-' + job.partFileIds.length + ' קבצי CSV החלקיים יישארו ב-Drive (תיקיית Waze Archive) — תוכל למחוק ידנית.\n\n' +
    'במקרה של prune-job, raw_data לא יימחק.',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  _clearArchiveJob();
  _deleteArchiveContinuationTriggers();
  ui.alert('בוטל', 'הג\'וב בוטל. אפשר להתחיל ייצוא חדש מהתפריט.', ui.ButtonSet.OK);
}

function menuMigrateToArchive() {
  var ui = SpreadsheetApp.getUi();
  var r = migrateToBaselineArchive();
  if (!r.ok) { ui.alert('שגיאה', r.error, ui.ButtonSet.OK); return; }
  ui.alert('הצלחה',
    'הועברו ' + r.processed + ' שורות מ-raw_data לארכיון אגרגטיבי קבוע.\n\n' +
    'הארכיון נמצא בלשונית מוסתרת _baseline_archive ומשמש לחישוב baseline היסטוריים.',
    ui.ButtonSet.OK);
}

function menuExportArchive() {
  var ui = SpreadsheetApp.getUi();
  // Watchdog: synchronous slice — see menuPruneNow for rationale.
  _scheduleArchiveWatchdog();
  var r = exportFullArchive();
  if (!r.ok) { ui.alert('שגיאה', r.error, ui.ButtonSet.OK); return; }
  if (r.done) {
    ui.alert('הצלחה',
      'הארכיון נשמר ל-Google Drive\n\n' +
      'שם קובץ: ' + r.filename + '\n' +
      'מספר פקקים: ' + r.count + '\n\n' +
      'פתח ב: ' + r.url, ui.ButtonSet.OK);
  } else {
    ui.alert('בתהליך',
      'הגיבוי גדול מדי לריצה אחת — ממשיך ברקע אוטומטית.\n\n' +
      'התקדמות: ' + r.progress.processed + ' מתוך ' + r.progress.total + '\n' +
      'בדוק שוב בעוד מספר דקות (תיקייה: Waze Archive).', ui.ButtonSet.OK);
  }
}

// Full export of raw_data — ignores the incremental high-water mark.
// Identical to the legacy "ייצא ארכיון מלא" behavior; useful when you want a
// complete snapshot regardless of what was archived before.
function menuExportArchiveFull() {
  var ui = SpreadsheetApp.getUi();
  // Watchdog: synchronous slice — see menuPruneNow for rationale.
  _scheduleArchiveWatchdog();
  var r = exportFullArchiveAll();
  if (!r.ok) { ui.alert('שגיאה', r.error, ui.ButtonSet.OK); return; }
  if (r.done) {
    ui.alert('הצלחה',
      'הארכיון נשמר ל-Google Drive\n\n' +
      'שם קובץ: ' + r.filename + '\n' +
      'מספר פקקים: ' + r.count + '\n\n' +
      'פתח ב: ' + r.url, ui.ButtonSet.OK);
  } else {
    ui.alert('בתהליך',
      'הייצוא גדול מדי לריצה אחת — ממשיך ברקע אוטומטית.\n\n' +
      'התקדמות: ' + r.progress.processed + ' מתוך ' + r.progress.total + '\n' +
      'בדוק שוב בעוד מספר דקות (תיקייה: Waze Archive).', ui.ButtonSet.OK);
  }
}

// Trigger _pruneOld immediately — archives + deletes rows older than
// RETENTION_DAYS, plus the oldest rows if raw_data exceeds MAX_RAW_ROWS.
// Useful when raw_data is approaching Google Sheets' cell cap.
function menuPruneNow() {
  var ui = SpreadsheetApp.getUi();
  if (_getArchiveJob()) {
    ui.alert('ייצוא כבר פעיל',
      'יש ג\'וב ייצוא בתהליך. המתן לסיומו לפני קציצה נוספת.',
      ui.ButtonSet.OK);
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName(RAW_SHEET);
  if (!raw || raw.getLastRow() < 2) {
    ui.alert('אין מה לקצץ', 'raw_data ריק.', ui.ButtonSet.OK);
    return;
  }
  var before = raw.getLastRow() - 1;
  // Watchdog: _pruneOld → _startArchiveJob → _runArchiveSlice runs synchronously.
  // If a chunk straddles the 6-min wall the execution dies; this trigger revives
  // the job at +7 min so no manual click is needed.
  _scheduleArchiveWatchdog();
  _pruneOld(ss);
  var job = _getArchiveJob();
  if (!job) {
    // No archive started — clear the watchdog we pre-scheduled.
    _deleteArchiveContinuationTriggers();
    ui.alert('אין שורות לקציצה',
      'כל ' + before + ' השורות ב-raw_data בתוך טווח ה-retention (' + RETENTION_DAYS + ' ימים) ' +
      'וגם מתחת לתקרת ' + MAX_RAW_ROWS + ' שורות. אין מה לקצץ כרגע.',
      ui.ButtonSet.OK);
    return;
  }
  ui.alert('קציצה החלה',
    job.totalRows + ' שורות (מתוך ' + before + ') מיוצאות ל-Drive ואז יימחקו מ-raw_data.\n\n' +
    'הריצה תמשיך אוטומטית ברקע — שם הקובץ: ' + job.filename + '\n' +
    'אם נתקע, השתמש ב-"▶️ המשך ייצוא תקוע".',
    ui.ButtonSet.OK);
}

// Show a per-sheet breakdown of cell usage so you can see what's eating the
// 10M-cell budget. Highlights any sheet whose "allocated" cells (max rows ×
// max cols, including empty trailing space) far exceed its data cells —
// those are candidates for menuCompactSheets.
function menuDiagnoseCellUsage() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = [];
  var totalAlloc = 0, totalData = 0;
  ss.getSheets().forEach(function(s) {
    var maxR = s.getMaxRows(), maxC = s.getMaxColumns();
    var lastR = s.getLastRow(), lastC = s.getLastColumn();
    var alloc = maxR * maxC;
    var data = lastR * lastC;
    totalAlloc += alloc;
    totalData += data;
    rows.push({ name: s.getName(), alloc: alloc, data: data, maxR: maxR, lastR: lastR });
  });
  rows.sort(function(a, b) { return b.alloc - a.alloc; });

  var lines = rows.map(function(r) {
    var slack = r.alloc - r.data;
    var slackPct = r.alloc > 0 ? Math.round(slack / r.alloc * 100) : 0;
    return r.name + ': ' + r.lastR + '/' + r.maxR + ' שורות, מוקצים ' +
      r.alloc.toLocaleString() + ' תאים' +
      (slackPct > 30 ? ' (' + slackPct + '% ריקים — מועמד לדחיסה)' : '');
  });
  lines.push('');
  lines.push('סה"כ data: ' + totalData.toLocaleString());
  lines.push('סה"כ מוקצים: ' + totalAlloc.toLocaleString() + ' / 10,000,000');
  lines.push('שוליים: ' + (10000000 - totalAlloc).toLocaleString());

  ui.alert('שימוש בתאים בגיליון', lines.join('\n'), ui.ButtonSet.OK);
}

// Trim empty trailing rows and columns from every sheet to reclaim cells.
// Leaves a small headroom buffer so the very next append doesn't immediately
// force a resize. Safe — uses getLastRow/getLastColumn which only return the
// last row/col with actual content.
function menuCompactSheets() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reclaimed = _compactAllSheets(ss);
  ui.alert('דחיסה הושלמה',
    'שוחררו ' + reclaimed.toLocaleString() + ' תאים מלשוניות עם שורות/עמודות ריקות.\n\n' +
    'הרץ "📊 בדוק שימוש בתאים" כדי לראות את המצב החדש.',
    ui.ButtonSet.OK);
}

function _compactAllSheets(ss) {
  var reclaimed = 0;
  ss.getSheets().forEach(function(s) {
    try {
      var name = s.getName();
      var maxR = s.getMaxRows();
      var lastR = Math.max(1, s.getLastRow());
      var keepRows = lastR + 100;  // buffer for next append
      if (maxR > keepRows) {
        var trimR = maxR - keepRows;
        s.deleteRows(keepRows + 1, trimR);
        reclaimed += trimR * s.getMaxColumns();
      }
      var maxC = s.getMaxColumns();
      // For known-schema sheets, trust the schema constant — getLastColumn can
      // be fooled by stray formatting/empty-string values in trailing cols,
      // leaving phantom "buffer" cols that quietly burn 500K+ cells each.
      var schemaC = null;
      if (name === RAW_SHEET) schemaC = RAW_COLS.length;
      else if (name === BASELINE_ARCHIVE) schemaC = BASELINE_COLS.length;
      var lastC = schemaC !== null ? schemaC : Math.max(1, s.getLastColumn());
      // For wide sheets (lots of rows), each buffer column costs `lastR` cells.
      // Above 50K rows we drop the buffer entirely — schema additions are rare
      // enough that auto-expand on setValues is cheaper than the standing cost.
      var keepCols = lastR > 50000 ? lastC : lastC + 2;
      if (maxC > keepCols) {
        var trimC = maxC - keepCols;
        s.deleteColumns(keepCols + 1, trimC);
        reclaimed += trimC * s.getMaxRows();
      }
    } catch(e) {
      Logger.log('Compact failed for ' + s.getName() + ': ' + e.message);
    }
  });
  return reclaimed;
}

// Toggle the daily auto-archive trigger (03:00 → exportFullArchive + _pruneOld).
function menuToggleDailyArchive() {
  var ui = SpreadsheetApp.getUi();
  if (_hasDailyArchiveTrigger()) {
    _deleteDailyArchiveTrigger();
    _setConfig(DAILY_ARCHIVE_CONFIG_KEY, 'off');
    ui.alert('🗓️ גיבוי יומי כובה',
      'הטריגר היומי הוסר. אפשר להפעיל שוב מאותו פריט בתפריט.',
      ui.ButtonSet.OK);
    return;
  }
  try {
    _installDailyArchiveTrigger();
    _setConfig(DAILY_ARCHIVE_CONFIG_KEY, 'on');
    ui.alert('🗓️ גיבוי יומי הופעל',
      'ייצוא אינקרמנטלי + קציצה ירוצו אוטומטית כל יום בשעה 03:00.',
      ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('שגיאה', 'התקנת הטריגר נכשלה: ' + e.message, ui.ButtonSet.OK);
  }
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
    // Prune-first: free space (if any rows are eligible) BEFORE appending, so
    // we don't crash against the cell cap. Idempotent — no-op if no rows match.
    _pruneOld(ss);
    _appendToRawData(ss, jams, snapshotTs);
    _logSource(ss, snapshotTs, jams.length);
    _rebuildAggregation(ss);
  } catch(e) {
    // Self-healing: if we hit the cell cap, make actual progress on freeing
    // space — start a new prune job, or if one already exists, force its
    // continuation trigger to fire NOW (in case it was orphaned). Callers
    // detect this via result.autoRecovering and react accordingly.
    if (_isCellCapError(e.message)) {
      var existingJob = _getArchiveJob();
      var initialSource = existingJob ? existingJob.source : 'prune';
      try {
        if (existingJob) {
          // Existing job — clear any stale continuation trigger and schedule
          // a fresh one that fires in ~1 sec. Pushes the stuck job forward.
          _deleteArchiveContinuationTriggers();
          _scheduleArchiveContinuation();
        } else {
          _pruneOld(ss);
        }
      } catch(_) {}
      return {
        ok: false, autoRecovering: true,
        initialSource: initialSource,
        error: existingJob && existingJob.source !== 'prune'
          ? 'ייצוא ישן בתהליך — דוחפים אותו קדימה, וכשיסיים תרוץ קציצה אוטומטית. שליפה חוזרת תרוץ בעוד 2 דקות.'
          : 'raw_data התמלא — פינוי אוטומטי החל. תוכל לעקוב כאן אחר ההתקדמות; שליפה חוזרת תרוץ בעוד 2 דקות.',
      };
    }
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
    // Remove legacy/renamed sheets that the current version no longer produces
    ['פירוט לפי מרווח זמן', 'מקרא'].forEach(function(legacyName) {
      var legacy = ss.getSheetByName(legacyName);
      if (legacy) ss.deleteSheet(legacy);
    });

    var raw = _readFiltered(ss, filter);
    if (!raw.length) return { ok: false, error: 'אין נתונים בטווח שנבחר' };
    var baselines = _computeBaselines(ss);          // reads from permanent _baseline_archive
    _sheet0_dashboard(ss, raw, filter, baselines);
    _sheet1_summary(ss, raw, filter, baselines);
    _sheet2_timebins(ss, raw, baselines);
    _sheet3_directions(ss, raw);
    _sheet4_anomalies(ss, raw);
    _sheet5_detail(ss, raw);
    _sheet6_legend(ss);
    ss.getSheetByName('🎯 לוח מחוונים').activate();
    var arch = ss.getSheetByName(BASELINE_ARCHIVE);
    var archiveCells = arch && arch.getLastRow() > 1 ? arch.getLastRow() - 1 : 0;
    return { ok: true, jams: raw.length, archiveJams: archiveCells };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Build baselines from the PERMANENT archive (_baseline_archive sheet).
// Keys by route × dir × hour × daytype. Falls back to ±1, ±2 neighbor hours if exact cell is sparse.
// Min sample size for a cell to be usable: n >= 3.
function _computeBaselines(ss) {
  var s = ss.getSheetByName(BASELINE_ARCHIVE);
  var byRDH = {};   // route::dir::hour::daytype → { n, sumDelay }
  var byRD  = {};   // route::dir → { n, sumDelay } (all hours combined — for summary view)

  if (s && s.getLastRow() > 1) {
    var rows = s.getRange(2, 1, s.getLastRow() - 1, BASELINE_COLS.length).getValues();
    rows.forEach(function(r) {
      var route = r[0], dir = r[1], date = r[2], hour = r[3];
      var n = +r[4] || 0, sumDelay = +r[5] || 0;
      if (!route || dir === '' || hour === '' || !n) return;
      var dt = _dayTypeFromDate(date);
      var k = route + '::' + dir + '::' + hour + '::' + dt;
      var b = byRDH[k] = byRDH[k] || { n: 0, sumDelay: 0 };
      b.n += n; b.sumDelay += sumDelay;
      var k2 = route + '::' + dir;
      var b2 = byRD[k2] = byRD[k2] || { n: 0, sumDelay: 0 };
      b2.n += n; b2.sumDelay += sumDelay;
    });
  }

  function lookupExact(routeName, dirIx, hour, daytype) {
    var b = byRDH[routeName + '::' + dirIx + '::' + hour + '::' + daytype];
    return (b && b.n >= 3) ? { avg: b.sumDelay / b.n, n: b.n } : null;
  }

  return {
    // Returns { avg, n, source } or null if even ±2 doesn't yield n>=3.
    avgPerJamAtHour: function(routeName, dirIx, hour, daytype) {
      var exact = lookupExact(routeName, dirIx, hour, daytype);
      if (exact) return { avg: exact.avg, n: exact.n, source: 'שעה זו' };
      for (var w = 1; w <= 2; w++) {
        var totN = 0, totSum = 0;
        for (var dh = -w; dh <= w; dh++) {
          var h = ((hour + dh) % 24 + 24) % 24;
          var bb = byRDH[routeName + '::' + dirIx + '::' + h + '::' + daytype];
          if (bb) { totN += bb.n; totSum += bb.sumDelay; }
        }
        if (totN >= 3) {
          return { avg: totSum / totN, n: totN, source: '±'+w+' שעות' };
        }
      }
      return null;
    },
    // Route × dir, all hours/daytypes combined — for "סיכום מסלולים" sheet only.
    // Returns null if no archive data.
    avgPerJam: function(routeName, dirIx) {
      var b = byRD[routeName + '::' + dirIx];
      return (b && b.n >= 3) ? b.sumDelay / b.n : null;
    }
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
  // Persist the URL/headers so an auto-recovery retry (or any other scheduled
  // fetch) has something to use — even if the user never explicitly clicked
  // "שמור" on the auto-fetch panel.
  try {
    if (_getConfig('fetch_url') !== url) _setConfig('fetch_url', url);
    if (headersJson && _getConfig('fetch_headers') !== headersJson) {
      _setConfig('fetch_headers', headersJson);
    }
  } catch(_) {}
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
  var result = processWazeJSON(resp.getContentText());
  // processWazeJSON already kicked off a prune if it detected the cell cap.
  // For the fetch path we ALSO schedule a one-shot retry, so the next attempt
  // happens automatically once raw_data has room.
  if (result && result.autoRecovering) {
    var scheduled = _scheduleFetchRetry();
    result.error = scheduled
      ? 'raw_data התמלא — פינוי אוטומטי החל ושליפה חוזרת תרוץ בעוד דקות ספורות. אין צורך לעשות כלום.'
      : 'raw_data התמלא ופינוי אוטומטי לא הצליח לאחר ' + FETCH_RETRY_MAX_ATTEMPTS + ' ניסיונות. הרץ "📊 בדוק שימוש בתאים" מהתפריט "מתקדם" כדי לראות מה תופס תאים, ואז "🗜️ דחס תאים ריקים" או מחק ידנית. אחרי שליפה מוצלחת המנגנון יתאפס.';
  } else if (result && result.ok) {
    _resetFetchRetryBudget();
  }
  return result;
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

  // New schedule = explicit user intervention. Reset the self-healing retry budget
  // so the next failure starts fresh.
  _resetFetchRetryBudget();
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
                'start_node','end_node','tt_min','route_name','dir_ix','archived'];

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
  // Hard guard against Sheets' ~555K-row wall (10M cells / 18 cols). If we're
  // close to it and a prune job is mid-flight, fail loudly with a friendly
  // pointer — silently swallowing the setValues error confuses users.
  var nCurrent = Math.max(0, s.getLastRow() - 1);
  var hardCeiling = Math.min(540000, Math.floor(MAX_RAW_ROWS * 1.08));
  if (nCurrent + jams.length > hardCeiling) {
    var jobMsg = _getArchiveJob()
      ? 'יש ייצוא+קציצה בתהליך — המתן לסיומו ונסה שוב.'
      : 'הרץ "🧹 קצץ raw_data עכשיו" מהתפריט והמתן לסיום הייצוא.';
    throw new Error('raw_data כמעט מלא (' + nCurrent + ' שורות, ' +
      'תקרת Sheets ~555K). ' + jobMsg);
  }
  var days = ['שני','שלישי','רביעי','חמישי','שישי','שבת','ראשון'];
  var jamObjsForArchive = [];
  var rows = jams.map(function(j) {
    var pm = j.pubMillis || snapshotTs.getTime();
    var dt = new Date(pm);
    var spd = j.speedKMH || 0;
    var ln  = j.length   || 0;
    var hour = dt.getHours();
    var street = j.street || '';
    var dateStr = Utilities.formatDate(dt, 'Asia/Jerusalem', 'yyyy-MM-dd');
    var routeInfo = _routeForStreet(street);
    var routeName = routeInfo ? routeInfo.routeName : '';
    var dirIx = routeInfo ? routeInfo.dirIx : '';
    var delay = j.delay || 0;
    var level = j.level || 0;

    if (routeInfo) {
      jamObjsForArchive.push({
        routeName: routeName, dirIx: dirIx,
        date: dateStr, hour: hour,
        delay_s: delay, speed: spd, level: level
      });
    }

    return [
      snapshotTs,                                       // snapshot_ts (Date)
      dt,                                                // pub_ts
      dateStr,
      days[dt.getDay() === 0 ? 6 : dt.getDay() - 1],
      hour,
      _tbin(hour),
      street, j.city || '',
      level, ln,
      delay, spd,
      j.startNode || '', j.endNode || '',
      spd > 0 ? Math.round((ln / 1000) / spd * 60 * 100) / 100 : '',
      routeName, dirIx,
      true,                                              // archived = true (already aggregated)
    ];
  });
  if (rows.length) {
    s.getRange(s.getLastRow() + 1, 1, rows.length, RAW_COLS.length).setValues(rows);
  }
  // Aggregate into permanent baseline archive (jams already marked archived=true above)
  if (jamObjsForArchive.length) {
    _upsertBaselineArchive(ss, jamObjsForArchive);
  }
}

var BASELINE_COLS = ['route','dir','date','hour','n','sum_delay_s','sum_speed','sum_level','last_updated'];

function _ensureBaselineArchive(ss) {
  var s = ss.getSheetByName(BASELINE_ARCHIVE);
  if (!s) {
    s = ss.insertSheet(BASELINE_ARCHIVE);
    try { s.hideSheet(); } catch(e) {}
    s.getRange(1, 1, 1, BASELINE_COLS.length).setValues([BASELINE_COLS])
     .setBackground('#1F3864').setFontColor('#fff').setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

// Upsert aggregated counters from a list of jam-objects.
// Each jam must have: routeName, dirIx, date, hour, delay_s, speed, level.
// Caller is responsible for ensuring jams aren't double-counted (use archived flag).
function _upsertBaselineArchive(ss, jamObjs) {
  if (!jamObjs || !jamObjs.length) return 0;
  var s = _ensureBaselineArchive(ss);
  var n = s.getLastRow() - 1;
  var existing = n > 0 ? s.getRange(2, 1, n, BASELINE_COLS.length).getValues() : [];

  // index existing rows by composite key for O(1) lookup
  var idx = {};
  for (var i = 0; i < existing.length; i++) {
    var r = existing[i];
    idx[r[0] + '::' + r[1] + '::' + r[2] + '::' + r[3]] = i;
  }

  // bucket incoming jams
  var buckets = {};
  jamObjs.forEach(function(j) {
    if (j.hour == null || !j.routeName) return;
    var dateStr = j.date;
    if (dateStr instanceof Date) {
      dateStr = Utilities.formatDate(dateStr, 'Asia/Jerusalem', 'yyyy-MM-dd');
    }
    var k = j.routeName + '::' + j.dirIx + '::' + dateStr + '::' + j.hour;
    var b = buckets[k] = buckets[k] || {
      route: j.routeName, dir: j.dirIx, date: dateStr, hour: j.hour,
      n: 0, sumDelay: 0, sumSpeed: 0, sumLevel: 0
    };
    b.n++;
    b.sumDelay += (+j.delay_s) || 0;
    b.sumSpeed += (+j.speed) || 0;
    b.sumLevel += (+j.level) || 0;
  });

  var now = new Date();
  var toAppend = [];
  Object.keys(buckets).forEach(function(k) {
    var b = buckets[k];
    if (idx[k] !== undefined) {
      var i = idx[k], r = existing[i];
      r[4] = (+r[4] || 0) + b.n;
      r[5] = (+r[5] || 0) + b.sumDelay;
      r[6] = (+r[6] || 0) + b.sumSpeed;
      r[7] = (+r[7] || 0) + b.sumLevel;
      r[8] = now;
      s.getRange(i + 2, 1, 1, BASELINE_COLS.length).setValues([r]);
    } else {
      toAppend.push([b.route, b.dir, b.date, b.hour, b.n, b.sumDelay, b.sumSpeed, b.sumLevel, now]);
    }
  });

  if (toAppend.length) {
    s.getRange(s.getLastRow() + 1, 1, toAppend.length, BASELINE_COLS.length).setValues(toAppend);
  }
  return Object.keys(buckets).length;
}

// Catch-up upsert: for rows in raw_data that aren't yet marked archived=true,
// aggregate them into _baseline_archive and flip the flag. Idempotent (only touches unarchived rows).
function _catchUpBaselineArchive(ss, raw, maxRows) {
  if (!raw || raw.getLastRow() < 2) return 0;
  var nRows = Math.min(maxRows || (raw.getLastRow() - 1), raw.getLastRow() - 1);
  var values = raw.getRange(2, 1, nRows, RAW_COLS.length).getValues();
  var jamObjs = [];
  var rowsToFlag = [];

  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    var archivedFlag = r[17];
    var isArchived = archivedFlag === true || archivedFlag === 'TRUE' || archivedFlag === 'true';
    if (isArchived) continue;

    var street = r[6];
    var routeName = r[15];
    var dirIx = r[16];
    if (!routeName || dirIx === '' || dirIx === null) {
      var info = _routeForStreet(street);
      if (info) { routeName = info.routeName; dirIx = info.dirIx; }
    }
    if (!routeName) continue;

    var dateStr = r[2];
    if (dateStr instanceof Date) {
      dateStr = Utilities.formatDate(dateStr, 'Asia/Jerusalem', 'yyyy-MM-dd');
    }
    jamObjs.push({
      routeName: routeName, dirIx: dirIx,
      date: dateStr, hour: r[4],
      delay_s: r[10] || 0, speed: r[11] || 0, level: r[8] || 0
    });
    rowsToFlag.push({ sheetRow: i + 2, routeName: routeName, dirIx: dirIx });
  }

  if (!jamObjs.length) return 0;

  _upsertBaselineArchive(ss, jamObjs);

  // Flag rows as archived (and write back resolved route_name/dir_ix if missing)
  rowsToFlag.forEach(function(rf) {
    raw.getRange(rf.sheetRow, 16, 1, 3).setValues([[rf.routeName, rf.dirIx, true]]);
  });
  return jamObjs.length;
}

// One-time migration: walk existing raw_data and populate _baseline_archive.
// Safe to re-run — only processes rows not yet marked archived.
function migrateToBaselineArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName(RAW_SHEET);
  if (!raw || raw.getLastRow() < 2) {
    return { ok: false, error: 'אין נתונים ב-raw_data' };
  }
  _ensureBaselineArchive(ss);
  var count = _catchUpBaselineArchive(ss, raw);
  return { ok: true, processed: count };
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
  // If a chunked archive job is mid-flight, skip — it will resume via trigger and
  // delete its own pruned rows on completion.
  if (_getArchiveJob()) return;

  var cutoff = new Date(Date.now() - RETENTION_DAYS * 86400000);

  var raw = ss.getSheetByName(RAW_SHEET);
  if (raw && raw.getLastRow() > 1) {
    var nRows = raw.getLastRow() - 1;
    // Read just the date column to find the boundary — avoids loading the whole sheet.
    var dates = raw.getRange(2, 1, nRows, 1).getValues();
    var firstKeep = -1;
    for (var i = 0; i < dates.length; i++) {
      if (new Date(dates[i][0]) >= cutoff) { firstKeep = i; break; }
    }
    var pruneCount = 0;
    if (firstKeep > 0) pruneCount = firstKeep;
    else if (firstKeep === -1) pruneCount = nRows;

    // Safety cap: even if all rows are within the date-retention window, prune
    // the oldest if raw_data is approaching Google Sheets' cell limit. Brings
    // the count down to 80% of MAX_RAW_ROWS to leave breathing room.
    if (nRows > MAX_RAW_ROWS) {
      var capPrune = nRows - Math.floor(MAX_RAW_ROWS * 0.8);
      if (capPrune > pruneCount) pruneCount = capPrune;
    }

    if (pruneCount > 0) {
      // Safety net: catch-up upsert any rows not yet aggregated to _baseline_archive.
      // Normally _appendToRawData has already marked them archived=true, so this is a no-op.
      _catchUpBaselineArchive(ss, raw, pruneCount);
      // Start a chunked archive job; deleteRows happens in _finalizeArchive after success.
      _startArchiveJob(ss, raw, 2, pruneCount, 'prune');
    } else {
      // No row pruning needed — but if we were called from auto-recovery, the
      // cell cap was hit despite raw_data being under our row cap. Reclaim
      // cells from empty trailing rows/columns in OTHER sheets as a fallback.
      try { _compactAllSheets(ss); } catch(_) {}
    }
  }

  // Prune sources log (small, safe to do inline)
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

// ─── Long-term archive to Google Drive (chunked + resumable) ────────
function _archiveFolder(ss) {
  var name = 'Waze Archive — ' + ss.getName();
  var iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

function _archiveProps() { return PropertiesService.getDocumentProperties(); }
function _getArchiveJob() {
  var s = _archiveProps().getProperty(ARCHIVE_JOB_PROP);
  return s ? JSON.parse(s) : null;
}
function _setArchiveJob(job) { _archiveProps().setProperty(ARCHIVE_JOB_PROP, JSON.stringify(job)); }
function _clearArchiveJob() { _archiveProps().deleteProperty(ARCHIVE_JOB_PROP); }

function _scheduleArchiveContinuation() {
  // Always replace any existing continuation trigger. A budget-exit's fresh
  // 1-sec trigger must override a long-delay watchdog set by a menu handler.
  _deleteArchiveContinuationTriggers();
  ScriptApp.newTrigger(ARCHIVE_TRIGGER_FN).timeBased().after(1000).create();
}

// Watchdog: scheduled before a synchronous slice runs in a menu handler. If the
// slice exits cleanly (done or budget exit), _finalizeArchive / _scheduleArchiveContinuation
// removes this. If the calling execution is hard-killed at the 6-min wall before
// the budget-exit code can run, this trigger fires ~7 min later and resumes
// the job — no manual click needed.
function _scheduleArchiveWatchdog() {
  _deleteArchiveContinuationTriggers();
  ScriptApp.newTrigger(ARCHIVE_TRIGGER_FN).timeBased().after(7 * 60 * 1000).create();
}

function _deleteArchiveContinuationTriggers() {
  var trigs = ScriptApp.getProjectTriggers();
  for (var i = 0; i < trigs.length; i++) {
    if (trigs[i].getHandlerFunction() === ARCHIVE_TRIGGER_FN) {
      ScriptApp.deleteTrigger(trigs[i]);
    }
  }
}

function _pad(n, w) {
  var s = String(n);
  while (s.length < w) s = '0' + s;
  return s;
}

function _formatRowCSV(r) {
  var out = [];
  for (var k = 0; k < r.length; k++) {
    var c = r[k];
    var s;
    if (c === null || c === undefined) s = '';
    else if (c instanceof Date) s = Utilities.formatDate(c, 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm:ss');
    else s = String(c);
    if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) {
      s = '"' + s.replace(/"/g, '""') + '"';
    }
    out.push(s);
  }
  return out.join(',');
}

function _chunkToCSV(rows, includeHeader) {
  var lines = includeHeader ? [RAW_COLS.join(',')] : [];
  for (var i = 0; i < rows.length; i++) lines.push(_formatRowCSV(rows[i]));
  return lines.join('\n');
}

function _logArchiveFile(ss, filename, url, jamCount, firstDate, lastDate) {
  var name = ARCHIVE_LOG_SHEET;
  var s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.setTabColor('#7C3AED');
    try { s.setRightToLeft(true); } catch(e) {}
    s.getRange(1, 1, 1, 6).setValues([['נשמר ב-','שם קובץ','מתאריך','עד תאריך','מס\' פקקים','קישור']])
     .setBackground('#1F3864').setFontColor('#fff').setFontWeight('bold')
     .setHorizontalAlignment('center');
    s.setFrozenRows(1);
  }
  var lastRow = s.getLastRow() + 1;
  s.getRange(lastRow, 1, 1, 6).setValues([[
    new Date(),
    filename,
    Utilities.formatDate(new Date(firstDate), 'Asia/Jerusalem', 'yyyy-MM-dd'),
    Utilities.formatDate(new Date(lastDate),  'Asia/Jerusalem', 'yyyy-MM-dd'),
    jamCount,
    '=HYPERLINK("' + url + '", "פתח ב-Drive")',
  ]]).setHorizontalAlignment('center').setFontFamily('Arial').setFontSize(10);
  for (var c = 1; c <= 6; c++) s.autoResizeColumn(c);
}

// Per-slice diagnostic log. One row per slice exit (budget cut-off or completion)
// so we can see whether time budget / chunk size are tuned right. Hidden sheet.
function _logArchiveSlice(ss, job, chunksThisSlice, elapsedMs, lastChunkMs, reason) {
  try {
    var name = '_archive_slice_log';
    var s = ss.getSheetByName(name);
    if (!s) {
      s = ss.insertSheet(name);
      try { s.hideSheet(); } catch(e) {}
      s.getRange(1, 1, 1, 9).setValues([[
        'time','filename','reason','chunks_this_slice','elapsed_ms','last_chunk_ms',
        'processed','total','total_parts'
      ]]).setBackground('#1F3864').setFontColor('#fff').setFontWeight('bold');
      s.setFrozenRows(1);
    }
    s.appendRow([
      new Date(),
      job.filename,
      reason,
      chunksThisSlice,
      elapsedMs,
      lastChunkMs,
      job.nextRow - 1,
      job.totalRows,
      job.partFileIds.length,
    ]);
  } catch(e) {
    Logger.log('_logArchiveSlice failed: ' + e.message);
  }
}

// Incremental export: dump only rows newer than the last archived snapshot_ts.
// On first run (no high-water mark), falls back to a full export.
// Returns { ok, done, ... } — when done=false a background trigger continues the job.
function exportFullArchive() {
  var existing = _getArchiveJob();
  if (existing) {
    return {
      ok: true, done: false, inProgress: true,
      progress: { processed: existing.nextRow - 1, total: existing.totalRows },
      filename: existing.filename,
    };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName(RAW_SHEET);
  if (!raw || raw.getLastRow() < 2) {
    return { ok: false, error: 'אין נתונים לייצוא' };
  }

  var nRows = raw.getLastRow() - 1;
  var lastTs = _archiveLastTs(ss);
  if (!lastTs) {
    // First-ever export — write everything and seed the high-water mark.
    return _startArchiveJob(ss, raw, 2, nRows, 'manual');
  }

  var cutoffMs = lastTs.getTime();
  var dates = raw.getRange(2, 1, nRows, 1).getValues();
  var startOffset = -1;
  for (var i = 0; i < dates.length; i++) {
    var ts = dates[i][0];
    var t = (ts instanceof Date) ? ts.getTime() : new Date(ts).getTime();
    if (t > cutoffMs) { startOffset = i; break; }
  }
  if (startOffset === -1) {
    return { ok: false, error: 'אין שורות חדשות מאז הגיבוי האחרון' };
  }
  return _startArchiveJob(ss, raw, startOffset + 2, nRows - startOffset, 'manual');
}

// Full re-export: ignores the high-water mark; exports all of raw_data.
function exportFullArchiveAll() {
  var existing = _getArchiveJob();
  if (existing) {
    return {
      ok: true, done: false, inProgress: true,
      progress: { processed: existing.nextRow - 1, total: existing.totalRows },
      filename: existing.filename,
    };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName(RAW_SHEET);
  if (!raw || raw.getLastRow() < 2) {
    return { ok: false, error: 'אין נתונים לייצוא' };
  }
  return _startArchiveJob(ss, raw, 2, raw.getLastRow() - 1, 'manual');
}

function getArchiveJobStatus() {
  var job = _getArchiveJob();
  var lastError = '';
  try { lastError = _archiveProps().getProperty('archiveJobLastError') || ''; } catch(_) {}
  if (!job) return { active: false, lastError: lastError };

  // Check whether a continuation trigger actually exists; if not, the auto-resume is broken.
  var hasTrigger = false;
  try {
    hasTrigger = ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === ARCHIVE_TRIGGER_FN;
    });
  } catch(_) {}

  return {
    active: true,
    processed: job.nextRow - 1,
    total: job.totalRows,
    filename: job.filename,
    source: job.source,
    parts: job.partFileIds ? job.partFileIds.length : 0,
    hasTrigger: hasTrigger,
    lastError: lastError,
  };
}

function _startArchiveJob(ss, raw, startSheetRow, rowCount, source) {
  // Snapshot the date range from the rows we're about to archive.
  var firstDate = raw.getRange(startSheetRow, 1).getValue();
  var lastDate = raw.getRange(startSheetRow + rowCount - 1, 1).getValue();
  var fmt = function(d) { return Utilities.formatDate(new Date(d), 'Asia/Jerusalem', 'yyyy-MM-dd'); };
  var stamp = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyyMMdd-HHmmss');
  var filename = 'waze-archive_' + fmt(firstDate) + '_to_' + fmt(lastDate) + '_' + stamp + '.csv';

  var job = {
    filename: filename,
    startSheetRow: startSheetRow,                    // first sheet row to read (>= 2)
    totalRows: rowCount,                             // number of rows to export from startSheetRow
    nextRow: 1,                                      // 1-indexed within slice; sheet row = startSheetRow + nextRow - 1
    firstDate: new Date(firstDate).toISOString(),
    lastDate: new Date(lastDate).toISOString(),
    partFileIds: [],
    source: source,                                  // 'manual' | 'prune'
    pruneRowCount: source === 'prune' ? rowCount : 0,
  };
  _setArchiveJob(job);
  return _runArchiveSlice();
}

function _continueArchiveJob() {
  // Triggered by a one-shot time trigger; clear it first so a thrown error won't loop.
  _deleteArchiveContinuationTriggers();
  try { _runArchiveSlice(); }
  catch(e) {
    Logger.log('Archive continuation failed: ' + e.message);
    // Persist the error to the doc so the user can see it (Logger output isn't visible in spreadsheet UI).
    try {
      _archiveProps().setProperty('archiveJobLastError',
        new Date().toISOString() + ' — ' + (e.message || String(e)));
    } catch(_) {}
    // Reschedule once so a transient error gets a retry; user can also re-run manually.
    _scheduleArchiveContinuation();
  }
}

function _runArchiveSlice() {
  var job = _getArchiveJob();
  if (!job) return { ok: false, error: 'אין ג\'וב פעיל' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName(RAW_SHEET);
  if (!raw) {
    _clearArchiveJob();
    return { ok: false, error: 'גיליון raw_data לא נמצא' };
  }
  var folder = _archiveFolder(ss);
  var startedAt = Date.now();
  var lastChunkMs = 0;
  var chunksThisSlice = 0;
  Logger.log('Archive slice start: ' + job.filename + ' — at row ' + job.nextRow + '/' + job.totalRows +
             ' (budget ' + (ARCHIVE_TIME_BUDGET_MS / 1000) + 's, chunk ' + ARCHIVE_CHUNK_ROWS + ' rows)');

  while (job.nextRow <= job.totalRows) {
    // Cooperative cancel check: if the job was cleared externally
    // (menuCancelArchiveJob), abort without overwriting state.
    var current = _getArchiveJob();
    if (!current || current.filename !== job.filename) {
      Logger.log('Archive slice aborted: job was cancelled externally');
      return { ok: false, cancelled: true };
    }
    // Forward-looking budget check: if the previous chunk's duration suggests
    // the NEXT one would finish past the budget, exit now rather than starting
    // a chunk that might run us into the 6-min Apps Script wall. On the first
    // iteration lastChunkMs=0, so we always attempt at least one chunk.
    var elapsed = Date.now() - startedAt;
    if (elapsed + lastChunkMs > ARCHIVE_TIME_BUDGET_MS) {
      _setArchiveJob(job);
      _scheduleArchiveContinuation();
      Logger.log('Archive slice budget exit: processed ' + (job.nextRow - 1) + '/' + job.totalRows +
                 ' (' + chunksThisSlice + ' chunks this slice, elapsed ' + Math.round(elapsed / 1000) +
                 's, lastChunkMs=' + lastChunkMs + ') — continuation scheduled');
      _logArchiveSlice(ss, job, chunksThisSlice, elapsed, lastChunkMs, 'budget');
      return {
        ok: true, done: false,
        progress: { processed: job.nextRow - 1, total: job.totalRows },
        filename: job.filename,
      };
    }
    var chunkStart = Date.now();
    var remaining = job.totalRows - (job.nextRow - 1);
    var size = Math.min(ARCHIVE_CHUNK_ROWS, remaining);
    var sheetRow = (job.startSheetRow || 2) + (job.nextRow - 1);
    var rows = raw.getRange(sheetRow, 1, size, RAW_COLS.length).getValues();
    var includeHeader = (job.partFileIds.length === 0);
    var csv = _chunkToCSV(rows, includeHeader);
    var partName = job.filename.replace(/\.csv$/, '') + '__part' + _pad(job.partFileIds.length + 1, 3) + '.csv';
    var partFile = folder.createFile(partName, csv, 'text/csv');
    job.partFileIds.push(partFile.getId());
    job.nextRow += size;
    _setArchiveJob(job);
    lastChunkMs = Date.now() - chunkStart;
    chunksThisSlice++;
    Logger.log('Chunk ' + job.partFileIds.length + ': ' + size + ' rows in ' + lastChunkMs +
               'ms — cumulative ' + (job.nextRow - 1) + '/' + job.totalRows);
  }

  Logger.log('Archive slice done: all ' + job.totalRows + ' rows processed in ' + chunksThisSlice +
             ' chunks this slice (total ' + job.partFileIds.length + ' chunks across all slices)');
  _logArchiveSlice(ss, job, chunksThisSlice, Date.now() - startedAt, lastChunkMs, 'done');

  return _finalizeArchive(ss, folder, job);
}

function _finalizeArchive(ss, folder, job) {
  var finalFile = null;
  var keptParts = false;

  // Try to merge parts into one file. If the merged blob would exceed Apps
  // Script's 50 MB createFile cap (or fails for any other reason), keep the
  // __partNNN files as-is so the export is never lost.
  try {
    var pieces = [];
    for (var i = 0; i < job.partFileIds.length; i++) {
      pieces.push(DriveApp.getFileById(job.partFileIds[i]).getBlob().getDataAsString());
    }
    finalFile = folder.createFile(job.filename, pieces.join('\n'), 'text/csv');
    for (var j = 0; j < job.partFileIds.length; j++) {
      try { DriveApp.getFileById(job.partFileIds[j]).setTrashed(true); } catch(e) {}
    }
  } catch(mergeErr) {
    Logger.log('Archive merge failed, keeping parts: ' + mergeErr.message);
    keptParts = true;
  }

  var url, displayName;
  if (finalFile) {
    url = finalFile.getUrl();
    displayName = job.filename;
  } else {
    url = folder.getUrl();
    displayName = job.filename.replace(/\.csv$/, '') + ' (' + job.partFileIds.length + ' parts)';
  }

  _logArchiveFile(ss, displayName, url, job.totalRows,
    new Date(job.firstDate), new Date(job.lastDate));

  // Advance the incremental high-water mark so the next export skips these rows.
  try { _setArchiveLastTs(new Date(job.lastDate)); } catch(_) {}

  if (job.source === 'prune' && job.pruneRowCount > 0) {
    var raw = ss.getSheetByName(RAW_SHEET);
    if (raw) raw.deleteRows(2, job.pruneRowCount);
  }

  _clearArchiveJob();
  _deleteArchiveContinuationTriggers();
  try { _archiveProps().deleteProperty('archiveJobLastError'); } catch(_) {}

  // Auto-chain: if this was an export (not a prune) and raw_data is still
  // over the cap, schedule a follow-up prune. Runs via a one-shot trigger so
  // the current execution doesn't exceed its 6-min budget.
  if (job.source !== 'prune') {
    try {
      var raw = ss.getSheetByName(RAW_SHEET);
      if (raw && (raw.getLastRow() - 1) > MAX_RAW_ROWS) {
        Logger.log('Auto-chaining prune after ' + job.source + ' archive: ' + (raw.getLastRow() - 1) + ' rows > cap');
        // Clear any existing post-archive prune trigger first so we don't pile up.
        ScriptApp.getProjectTriggers().forEach(function(t) {
          if (t.getHandlerFunction() === POST_ARCHIVE_PRUNE_FN) ScriptApp.deleteTrigger(t);
        });
        ScriptApp.newTrigger(POST_ARCHIVE_PRUNE_FN).timeBased().after(1000).create();
      }
    } catch(_) {}
  }

  return {
    ok: true, done: true,
    count: job.totalRows,
    url: url,
    filename: displayName,
    keptParts: keptParts,
    source: job.source,
  };
}

// Triggered ~1 sec after a manual/export job completes. Starts a prune if
// raw_data is still over the cap.
function _postArchivePrune() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === POST_ARCHIVE_PRUNE_FN) ScriptApp.deleteTrigger(t);
  });
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (_getArchiveJob()) return;  // another job started; let it run
    _pruneOld(ss);
  } catch(e) {
    Logger.log('Post-archive prune failed: ' + e.message);
  }
}

// High-water mark: ISO timestamp of the latest snapshot_ts that has been archived.
// Used by exportFullArchive to skip rows already exported.
function _archiveLastTs(ss) {
  var p = _archiveProps().getProperty(ARCHIVE_LAST_TS_PROP);
  if (p) {
    var d = new Date(p);
    if (!isNaN(d.getTime())) return d;
  }
  // Bootstrap from existing log sheet if present (max value in 'עד תאריך' column).
  var s = ss.getSheetByName(ARCHIVE_LOG_SHEET);
  if (!s || s.getLastRow() < 2) return null;
  var n = s.getLastRow() - 1;
  var col = s.getRange(2, 4, n, 1).getValues();
  var maxMs = 0;
  for (var i = 0; i < col.length; i++) {
    var v = col[i][0];
    if (!v) continue;
    var ms;
    if (v instanceof Date) ms = v.getTime();
    else {
      var m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (!m) continue;
      // Treat the logged "yyyy-MM-dd" as end-of-day in Israel (+03:00 DST is
      // worst-case; off-by-an-hour is acceptable since duplicates are harmless).
      ms = new Date(m[1] + '-' + m[2] + '-' + m[3] + 'T23:59:59+03:00').getTime();
    }
    if (ms > maxMs) maxMs = ms;
  }
  return maxMs ? new Date(maxMs) : null;
}

function _setArchiveLastTs(d) {
  if (!d) return;
  var newMs = (d instanceof Date) ? d.getTime() : new Date(d).getTime();
  if (isNaN(newMs)) return;
  // Monotonic: never move the high-water mark backward, so a prune that
  // archives old rows doesn't make us re-export newer ones on the next incremental.
  var existing = _archiveProps().getProperty(ARCHIVE_LAST_TS_PROP);
  if (existing) {
    var existingMs = new Date(existing).getTime();
    if (!isNaN(existingMs) && existingMs >= newMs) return;
  }
  _archiveProps().setProperty(ARCHIVE_LAST_TS_PROP, new Date(newMs).toISOString());
}

// ─── Daily auto-archive trigger ────────────────────────────────────
function _hasDailyArchiveTrigger() {
  return ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === DAILY_ARCHIVE_TRIGGER_FN;
  });
}

function _installDailyArchiveTrigger() {
  if (_hasDailyArchiveTrigger()) return;
  ScriptApp.newTrigger(DAILY_ARCHIVE_TRIGGER_FN)
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

function _deleteDailyArchiveTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === DAILY_ARCHIVE_TRIGGER_FN) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// Daily trigger handler: incremental export to Drive, then prune old/excess
// rows. Both are no-ops when there's nothing to do, so it's safe to run nightly.
// Errors are logged but never throw — next night's run will retry.
function _dailyArchive() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Export new rows first (sets the archive job if any); prune skips when a
    // job is mid-flight, so it'll be a no-op this run if exportFullArchive
    // started one — that's fine, prune runs again on the next daily tick.
    exportFullArchive();
    if (!_getArchiveJob()) _pruneOld(ss);
  } catch(e) {
    Logger.log('Daily archive failed: ' + e.message);
    try {
      _archiveProps().setProperty('archiveJobLastError',
        new Date().toISOString() + ' — daily: ' + (e.message || String(e)));
    } catch(_) {}
  }
}

// ─── Archive-job watchdog ───────────────────────────────────────────
function _hasJobMonitorTrigger() {
  return ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === JOB_MONITOR_TRIGGER_FN;
  });
}

function _installJobMonitorTrigger() {
  if (_hasJobMonitorTrigger()) return;
  ScriptApp.newTrigger(JOB_MONITOR_TRIGGER_FN)
    .timeBased()
    .everyHours(6)
    .create();
}

// Resume archive jobs that lost their continuation trigger (e.g., trigger
// quota exhaustion, a thrown error that didn't reschedule). Runs every 6h.
function _monitorArchiveJob() {
  var job = _getArchiveJob();
  if (!job) return;
  var hasContinuation = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === ARCHIVE_TRIGGER_FN;
  });
  if (!hasContinuation) {
    Logger.log('Monitor: resuming orphaned archive job ' + job.filename);
    _scheduleArchiveContinuation();
  }
}

// ─── Self-healing fetch ─────────────────────────────────────────────
function _isCellCapError(msg) {
  if (!msg) return false;
  return /10000000 cells|cells in the workbook|raw_data כמעט מלא/.test(msg);
}

function _scheduleFetchRetry() {
  var props = PropertiesService.getDocumentProperties();
  var count = parseInt(props.getProperty(FETCH_RETRY_COUNT_PROP) || '0', 10);
  // Replace any existing pending retry to coalesce repeated failures.
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === FETCH_RETRY_TRIGGER_FN) {
      ScriptApp.deleteTrigger(t);
    }
  });
  // If a chunked archive job is actively in flight, recovery IS making progress —
  // skip the budget check and use the max backoff so we don't pile up retries
  // while the archive grinds through its chunks. The 6-hour _monitorArchiveJob
  // revives the job if it stalls.
  var hasActiveJob = !!_getArchiveJob();
  if (!hasActiveJob && count >= FETCH_RETRY_MAX_ATTEMPTS) {
    return false;
  }
  var delay = hasActiveJob
    ? FETCH_RETRY_BACKOFF_MS[FETCH_RETRY_BACKOFF_MS.length - 1]
    : FETCH_RETRY_BACKOFF_MS[Math.min(count, FETCH_RETRY_BACKOFF_MS.length - 1)];
  ScriptApp.newTrigger(FETCH_RETRY_TRIGGER_FN)
    .timeBased()
    .after(delay)
    .create();
  if (!hasActiveJob) props.setProperty(FETCH_RETRY_COUNT_PROP, String(count + 1));
  return true;
}

function _resetFetchRetryBudget() {
  try { PropertiesService.getDocumentProperties().deleteProperty(FETCH_RETRY_COUNT_PROP); } catch(_) {}
}

// One-shot trigger handler that retries the scheduled fetch.
function _retryScheduledFetch() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === FETCH_RETRY_TRIGGER_FN) {
      ScriptApp.deleteTrigger(t);
    }
  });
  _scheduledFetch();
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
      routeName: r[15] || '', dirIx: r[16] || '',
      archived: r[17] === true || r[17] === 'TRUE' || r[17] === 'true',
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

// Short explanatory intro at the top of a sheet — one row, merged across columns.
function _tabIntro(sheet, ncols, text, row) {
  row = row || 1;
  sheet.getRange(row, 1, 1, ncols).merge()
    .setValue('ℹ️ ' + text)
    .setBackground('#E0F2FE').setFontColor('#075985')
    .setFontSize(10).setFontStyle('italic')
    .setHorizontalAlignment('right').setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(row, 36);
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

// Day-of-week → 'weekday' | 'weekend'. Israeli convention: ו'+ש' are weekend.
function _dayType(dayName) {
  return (dayName === 'שישי' || dayName === 'שבת') ? 'weekend' : 'weekday';
}
function _dayTypeFromDate(dateStr) {
  var d = new Date(dateStr).getDay();   // 0=Sun..6=Sat
  return (d === 5 || d === 6) ? 'weekend' : 'weekday';
}

// Lookup table built lazily: street → { routeName, dirIx }.
// Streets that appear in multiple routes will resolve to the FIRST match — should be rare.
var _STREET_INDEX = null;
function _buildStreetIndex() {
  if (_STREET_INDEX) return _STREET_INDEX;
  _STREET_INDEX = {};
  for (var i = 0; i < ROUTES.length; i++) {
    var r = ROUTES[i];
    (r.streets_dir1 || []).forEach(function(s) {
      if (!_STREET_INDEX[s]) _STREET_INDEX[s] = { routeName: r.name, dirIx: 1 };
    });
    (r.streets_dir2 || []).forEach(function(s) {
      if (!_STREET_INDEX[s]) _STREET_INDEX[s] = { routeName: r.name, dirIx: 2 };
    });
  }
  return _STREET_INDEX;
}
function _routeForStreet(street) {
  return _buildStreetIndex()[street] || null;
}

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
      // Aggregate deviation across all (hour, daytype) groups of this route×dir.
      // Each group contributes a baseline lookup; we average the per-group deviations
      // weighted by jam count, and tag the dominant source.
      var groups = {};
      jl.forEach(function(j) {
        if (j.hour == null) return;
        var dt = _dayType(j.day);
        var k = j.hour + '::' + dt;
        (groups[k] = groups[k] || { hour: j.hour, daytype: dt, jams: [] }).jams.push(j);
      });

      var weightedDev = 0, totalJamsWithBase = 0, sourceCounts = {};
      Object.keys(groups).forEach(function(k) {
        var g = groups[k];
        var avg = _sum(g.jams, 'delay_s') / g.jams.length;
        var base = baselines && baselines.avgPerJamAtHour(route.name, d.ix, g.hour, g.daytype);
        if (base && base.avg > 0) {
          var dv = (avg - base.avg) / base.avg * 100;
          weightedDev += dv * g.jams.length;
          totalJamsWithBase += g.jams.length;
          sourceCounts[base.source] = (sourceCounts[base.source] || 0) + g.jams.length;
        }
      });

      var devPct, src, status;
      if (totalJamsWithBase > 0) {
        devPct = _round1(weightedDev / totalJamsWithBase);
        // pick the dominant source label
        var best = null, bestN = 0;
        Object.keys(sourceCounts).forEach(function(s) {
          if (sourceCounts[s] > bestN) { best = s; bestN = sourceCounts[s]; }
        });
        src = best || 'היסטורי';
        status = devPct > 50 ? 'חריג מאוד'
               : devPct > 25 ? 'עמוס'
               : devPct > 10 ? 'מתון'
               : devPct < -10 ? 'טוב מהרגיל'
               : 'תקין';
      } else {
        devPct = null;
        src = 'אין נתונים';
        status = 'אין מספיק נתונים';
      }

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

  ws.getRange(3, 1, 1, ncW).merge()
    .setValue('ℹ️ סקירה כללית. KPIs סופרים מסלולים לפי סטטוס · "10 הגרועים" / "5 הטובים" ממוינים לפי % סטייה מההיסטוריה · "איכות הניתוח" אומר על איזה baseline נשענת ההשוואה (שעה זו = מדויק).')
    .setBackground('#E0F2FE').setFontColor('#075985')
    .setFontSize(10).setFontStyle('italic')
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setWrap(true);
  ws.setRowHeight(3, 36);

  // ─── KPI strip (status counts) ───
  var counts = { 'תקין':0,'מתון':0,'עמוס':0,'חריג מאוד':0,'טוב מהרגיל':0,'אין מספיק נתונים':0 };
  summary.forEach(function(r) { counts[r.status] = (counts[r.status] || 0) + 1; });

  var kpis = [
    { label: '🟢 תקין',         count: counts['תקין'],              color: '#10B981' },
    { label: '🟡 מתון',         count: counts['מתון'],              color: '#F59E0B' },
    { label: '🟠 עמוס',         count: counts['עמוס'],              color: '#F97316' },
    { label: '🔴 חריג מאוד',    count: counts['חריג מאוד'],         color: '#EF4444' },
    { label: '🔵 טוב מהרגיל',   count: counts['טוב מהרגיל'],        color: '#3B82F6' },
    { label: '⚪ אין נתונים',    count: counts['אין מספיק נתונים'],  color: '#94A3B8' },
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

  var withDev = summary.filter(function(r) { return r.devPct !== null; });
  var top10 = withDev.slice().sort(function(a, b) { return b.devPct - a.devPct; }).slice(0, 10);
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

  var best5 = withDev.slice().sort(function(a, b) { return a.devPct - b.devPct; }).slice(0, 5);
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
    summary.filter(function(r){ return r.section === sec; }).forEach(function(r) {
      if (c[r.status] !== undefined) c[r.status]++;
    });
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
  var srcExact  = summary.filter(function(r){ return r.src === 'שעה זו'; }).length;
  var srcW1     = summary.filter(function(r){ return r.src === '±1 שעות'; }).length;
  var srcW2     = summary.filter(function(r){ return r.src === '±2 שעות'; }).length;
  var srcNone   = summary.filter(function(r){ return r.src === 'אין נתונים'; }).length;
  ws.getRange(row, 1, 1, ncW).merge()
    .setValue('📊 איכות הניתוח: ' +
              srcExact + ' שעה זו  •  ' +
              srcW1 + ' ±1 שעות  •  ' +
              srcW2 + ' ±2 שעות  •  ' +
              srcNone + ' אין נתונים' +
              (srcNone > 0 ? '  ⚠️ העלה עוד דגימות' : ''))
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

  // intro row 3 will be written AFTER _autoWidth at the end of this function,
  // so the merged intro text doesn't influence column auto-sizing.
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
  _autoWidth(ws, nc);
  _tabIntro(ws, nc, 'שורה לכל מסלול. שני הכיוונים מוצגים זה ליד זה. עמודות "סטייה כ1/כ2" משוות מול ממוצע היסטורי כללי של הכיוון (לא ספציפי לשעה — לרזולוציה לפי שעה ראה "פירוט לפי שעה").', 3);
  ws.setFrozenRows(4);
}

function _sheet2_timebins(ss, jams, baselines) {
  var ws = _newSheet(ss, 'פירוט לפי שעה', '#2E75B6');
  var cols = ['אזור','מסלול','כיוון','שעה','סוג יום',
              'מס\' פקקים','אורך (ק"מ)','השהיה (דק\')','מהירות ממוצעת','רמת פקק ממוצעת',
              'השהיה לפקק (דק\')','ממוצע היסטורי (דק\')','מקור השוואה','n',
              'סטייה %','סטטוס'];
  _hdrRow(ws, cols, 2);

  var row = 3;
  ROUTES.forEach(function(route) {
    [{label:route.dir1_label, streets:route.streets_dir1, ix:1},
     {label:route.dir2_label, streets:route.streets_dir2, ix:2}].forEach(function(dir) {
      if (!dir.label || !dir.streets.length) return;
      var jl = _getJamsForStreets(jams, dir.streets);
      if (!jl.length) return;

      // Group by (hour, daytype). Most jams carry day name; daytype derives from day.
      var groups = {};
      jl.forEach(function(j) {
        if (j.hour == null) return;
        var dt = _dayType(j.day);
        var k = j.hour + '::' + dt;
        (groups[k] = groups[k] || { hour: j.hour, daytype: dt, jams: [] }).jams.push(j);
      });

      var keys = Object.keys(groups).sort(function(a, b) {
        var ga = groups[a], gb = groups[b];
        if (ga.daytype !== gb.daytype) return ga.daytype < gb.daytype ? -1 : 1;
        return ga.hour - gb.hour;
      });

      keys.forEach(function(k) {
        var g = groups[k];
        var jj = g.jams;
        var tl = _round1(_sum(jj,'length_m')/1000);
        var td = _round1(_sum(jj.filter(function(j){return j.delay_s>0;}),'delay_s')/60);
        var spds = jj.filter(function(j){return j.speed>0;}).map(function(j){return j.speed;});
        var asp = spds.length ? _round1(_mean(spds)) : 0;
        var alv = _round1(_mean(jj.map(function(j){return j.level;})));
        var curAvg = jj.length ? _sum(jj, 'delay_s') / jj.length : 0;

        var base = baselines && baselines.avgPerJamAtHour(route.name, dir.ix, g.hour, g.daytype);

        var histMin, source, n, dev, status;
        if (base) {
          histMin = _round1(base.avg / 60);
          source = base.source;
          n = base.n;
          dev = base.avg > 0 ? _round1((curAvg - base.avg) / base.avg * 100) : 0;
          status = dev > 50 ? 'חריג מאוד'
                 : dev > 25 ? 'עמוס'
                 : dev > 10 ? 'מתון'
                 : dev < -10 ? 'טוב מהרגיל'
                 : 'תקין';
        } else {
          histMin = '—';
          source = '—';
          n = '—';
          dev = null;
          status = 'אין מספיק נתונים';
        }

        var hourLabel = _pad(g.hour) + ':00';
        var dtLabel = g.daytype === 'weekend' ? 'סופ"ש' : 'חול';

        _dataRow(ws, row, [route.section, route.name, dir.label, hourLabel, dtLabel,
                           jj.length, tl, td, asp, alv,
                           _round1(curAvg/60), histMin, source, n,
                           dev !== null ? (dev >= 0 ? '+' : '') + dev + '%' : '—',
                           status],
                 row % 2 === 0);
        if (dev !== null) _colorDeviation(ws, row, 15, dev);
        _colorStatus(ws, row, 16, status);
        row++;
      });
    });
  });
  _autoWidth(ws, cols.length);
  _tabIntro(ws, cols.length, 'הלב של הניתוח. שורה לכל מסלול × כיוון × שעה × סוג יום (חול/סופ"ש). "סטייה %" אומרת אם המצב כעת גרוע (חיובי) או טוב (שלילי) ביחס להיסטוריה באותה שעה. "מקור השוואה": "שעה זו" = השוואה מדויקת; "±1/±2 שעות" = הורחב כי אין מספיק דגימות; "—" = אין מספיק היסטוריה.', 1);
  ws.setFrozenRows(2);
}

function _sheet3_directions(ss, jams) {
  var ws = _newSheet(ss, 'השוואת כיוונים', '#ED7D31');
  var cols = ['אזור','מסלול','מרחק','זמן חופשי','כיוון 1','כיוון 2',
              'השהיה כ1','השהיה כ2','זמן כ1','זמן כ2','הפרש','כיוון עמוס','יחס'];
  _hdrRow(ws, cols, 2);
  var row = 3;
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
  _autoWidth(ws, cols.length);
  _tabIntro(ws, cols.length, 'השוואת שני כיוונים זה מול זה — רק מסלולים דו-כיווניים. "כיוון עמוס" מסומן באדום; "יחס" מציג כמה פעמים אחד עמוס מהשני (1.5x = פי 1.5).', 1);
  ws.setFrozenRows(2);
}

function _sheet4_anomalies(ss, jams) {
  var ws = _newSheet(ss, 'חריגות', '#C00000');
  var cols = ['#','אזור','מסלול','כיוון','קטע','עיר','תאריך','יום','שעה','מרווח',
              'מהירות','אורך (מ\')','השהיה (דק\')','ממוצע (דק\')','חריגה %','רמת פקק','חומרה'];
  _hdrRow(ws, cols, 2);
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
  var row = 3;
  anoms.forEach(function(a, i) {
    _dataRow(ws, row, [i+1, a.sec, a.rt, a.dir, a.seg, a.city, a.date, a.day, a.hour, a.tb,
                       a.spd, a.ln, a.dm, a.am, a.dp+'%', a.lv, a.sv],
             row % 2 === 1);
    _colorStatus(ws, row, 17, a.sv==='קריטי'?'חריג מאוד':a.sv==='גבוה'?'עמוס':'מתון');
    row++;
  });
  _autoWidth(ws, cols.length);
  _tabIntro(ws, cols.length, 'פקקים בודדים שבולטים מאוד בתוך הפילטר הנוכחי. הקריטריון מקומי (1.5σ מעל הממוצע, או speed<5, או level≥4) — לא היסטורי. ממוין לפי חומרה ואז גודל ההשהיה.', 1);
  ws.setFrozenRows(2);
}

function _sheet5_detail(ss, jams) {
  var ws = _newSheet(ss, 'פירוט פקקים', '#548235');
  var cols = ['#','אזור','מסלול','כיוון','קטע','עיר','תאריך','יום','שעה','מרווח',
              'מהירות','אורך (מ\')','השהיה (שנ\')','זמן נסיעה (דק\')','רמת פקק'];
  _hdrRow(ws, cols, 2);
  var row = 3, idx = 1;
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
                 row % 2 === 1);
        row++;
      });
    });
  });
  _autoWidth(ws, cols.length);
  _tabIntro(ws, cols.length, 'יומן מלא של כל הפקקים בפילטר הנוכחי — שורה לכל פקק יחיד. ממוין לפי מסלול → כיוון → זמן. שימושי לחפירה לעומק.', 1);
  ws.setFrozenRows(2);
}

function _sheet6_legend(ss) {
  var ws = _newSheet(ss, 'מקרא ומתודולוגיה', '#666666');
  // Position as the first tab — entry point for users
  try { ss.setActiveSheet(ws); ss.moveActiveSheet(1); } catch(e) {}

  var SEC = 'section', HDR = 'header', ROW = 'row', BLANK = 'blank';
  var items = [
    { type: SEC, text: '§1. מדריך לשוניות — מה רואים בכל גיליון' },
    { type: HDR, a: 'לשונית', b: 'מה היא מציגה ואיך לקרוא אותה' },

    { type: ROW, a: '🎯 לוח מחוונים',
      b: 'תצוגה אחת על דף — סיכום מהיר של מצב התנועה. בראש: שש כרטיסיות KPI שסופרות כמה מסלולים-כיוונים נמצאים בכל סטטוס (תקין/מתון/עמוס/חריג מאוד/טוב מהרגיל/אין נתונים). אחרי זה: 10 המסלולים הגרועים ביותר (לפי % סטייה מההיסטוריה), 5 המסלולים הטובים, ופילוח לפי אזור (דרום/מרכז/צפון). בתחתית: שורת איכות הניתוח שאומרת כמה מסלולים נשענים על baseline מדויק vs. הרחבת חלון. הלשונית הזו טובה ל"מה קורה עכשיו" — ולא לחפירה לעומק.' },

    { type: ROW, a: 'סיכום מסלולים',
      b: 'שורה לכל מסלול (כביש 1, כביש 6 וכו\'), מקובץ לפי אזור. כל שורה מציגה את שני הכיוונים זה ליד זה — כמה פקקים בכל כיוון, סך השהיה בדקות, וזמן נסיעה משוער (זמן חופשי + השהיה). שתי העמודות האחרונות הן "סטייה כ1/כ2 vs ממוצע כללי" — איך שתי המספרים האלה משווים את כל הפקקים בכיוון לממוצע ההיסטורי הכללי של אותו כיוון (מכל השעות). זה תצוגה גסה — לרזולוציה אמיתית של "האם זה גרוע מהרגיל באותה שעה" לך ללשונית הבאה.' },

    { type: ROW, a: 'פירוט לפי שעה',
      b: 'הלב של הניתוח. שורה לכל שילוב של (מסלול × כיוון × שעה × סוג יום). מציגה: כמה פקקים בשעה הזו, ההשהיה הממוצעת לפקק כעת, ההשהיה הממוצעת ההיסטורית באותה שעה, ועמודת "סטייה %" — האם המצב כעת גרוע (חיובי) או טוב (שלילי) ביחס להיסטוריה. עמודת "מקור השוואה" אומרת איזה חלון שימש: "שעה זו" = השוואה מדויקת; "±1 שעות" / "±2 שעות" = הורחבנו לחלון כי אין מספיק דגימות; "—" = אין מספיק היסטוריה בכלל. עמודת n = כמה פקקים היסטוריים השתתפו ב-baseline.' },

    { type: ROW, a: 'השוואת כיוונים',
      b: 'רק מסלולים דו-כיווניים. מציגה שני כיוונים זה מול זה: השהיה כ1 vs כ2, זמן משוער כ1 vs כ2, ההפרש ביניהם, הכיוון העמוס יותר (מסומן באדום), ויחס (1.5x = כיוון אחד פי 1.5 מהשני). מועיל לתשובה: "אם אני יוצא עכשיו, באיזה כיוון יהיה פקק יותר גדול?"' },

    { type: ROW, a: 'חריגות',
      b: 'רשימת פקקים בודדים שבולטים מאוד מהאחרים. הקריטריון אינו השוואה היסטורית אלא סטטיסטיקה מקומית: ההשהיה גבוהה ב-1.5 סטיות תקן מעל הממוצע של אותו מסלול×כיוון בפילטר הנוכחי, או שהמהירות נמוכה מ-5 קמ"ש, או שרמת הפקק היא 4-5 (כבד/עצירה). ממוין לפי חומרה (קריטי → גבוה → בינוני) ואז לפי גודל ההשהיה. שימושי לזהות "מה הפקק הכי חריג בשבוע האחרון?" — בלי קשר אם זה דפוס רגיל לכביש הזה או לא.' },

    { type: ROW, a: 'פירוט פקקים',
      b: 'יומן גולמי — שורה לכל פקק יחיד בפילטר הנוכחי. רואים בדיוק איזה קטע (start → end), עיר, תאריך, שעה, מהירות, אורך הפקק במטרים, השהיה בשניות, וזמן נסיעה. ממוין לפי מסלול × כיוון × זמן. שימושי לחפירה — "מה בדיוק קרה ביום שלישי ב-08:00?"' },

    { type: ROW, a: 'אגרגציה לאורך זמן',
      b: 'תצוגה ישירה של הארכיון הקבוע (_baseline_archive) — כל ההיסטוריה, לא רק 30 הימים האחרונים. שורה לכל (תאריך × מסלול × כיוון × שעה) עם מספר הפקקים שהיו וההשהיה/מהירות/רמה הממוצעת. ממוין מהחדש לישן. שימושי למגמות לאורך זמן: "האם התנועה בכביש 1 ב-17:00 גרועה יותר בחורף או בקיץ?", "האם נובמבר 2026 גרוע מנובמבר 2025?". מוגבל ל-5000 שורות אחרונות בתצוגה (אם הארכיון גדול מזה — סנן ב-Sheets פילטר).' },

    { type: ROW, a: 'מקרא ומתודולוגיה',
      b: 'הלשונית הזו. הסבר על השדות, איך מחושב הניתוח, מה משמעות כל סטטוס וצבע, ומה המגבלות.' },

    { type: BLANK },
    { type: SEC, text: '§2. שדות הנתונים' },
    { type: HDR, a: 'שדה', b: 'הסבר' },
    { type: ROW, a: 'זמן חופשי (free_flow_min)', b: 'זמן נסיעה תיאורטי ללא עומסים (קבוע ידני לכל מסלול). משמש לעמודת ייחוס בלבד, לא לחישוב סטייה.' },
    { type: ROW, a: 'השהיה (delay_s)', b: 'כמה שניות נוספו לזמן הנסיעה בגלל הפקק (לפי Waze).' },
    { type: ROW, a: 'השהיה לפקק', b: 'ממוצע ה-delay_s לפקק בודד בטווח הנוכחי (דקות).' },
    { type: ROW, a: 'רמת פקק (level)', b: '1=זרימה · 2=מתון · 3=בינוני · 4=כבד · 5=עצירה.' },
    { type: ROW, a: 'מהירות (speed_kmh)', b: 'מהירות נסיעה ממוצעת באזור הפקק.' },
    { type: ROW, a: 'tbin', b: 'רצועת 4 שעות (לתאימות לאחור). הניתוח החדש משתמש בשעה בודדת.' },
    { type: ROW, a: 'route_name + dir_ix', b: 'שיוך הפקק למסלול מוגדר ולכיוון (1 או 2). מחושב בעת ה-upload.' },
    { type: ROW, a: 'archived', b: 'דגל: האם הפקק כבר נספר ל-_baseline_archive (מונע ספירה כפולה).' },

    { type: BLANK },
    { type: SEC, text: '§3. ארכיטקטורת הנתונים — סוגי לשוניות' },
    { type: HDR, a: 'לשונית', b: 'מתי מתעדכנת ומה היא צוברת' },

    { type: ROW, a: '🔄 7 לשוניות הניתוח',
      b: 'לוח מחוונים · סיכום מסלולים · פירוט לפי שעה · השוואת כיוונים · חריגות · פירוט פקקים · מקרא ומתודולוגיה — **נמחקות ונבנות מחדש בכל "החל סינון"**. הפילטר חל על raw_data, אבל ההשוואות מול _baseline_archive (כל ההיסטוריה).' },

    { type: ROW, a: '🔁 אגרגציה לאורך זמן',
      b: 'נבנית מחדש **רק בעת upload חדש**, לא בעת סינון. תצוגה של _baseline_archive ממוין מהחדש לישן.' },

    { type: ROW, a: '📥 raw_data',
      b: 'גלוי. שורה לכל פקק בכל snapshot — מקור גרעיני. עמודות: snapshot_ts, pub_ts, date, day, hour, tbin, street, city, level, length_m, delay_s, speed_kmh, start_node, end_node, tt_min, route_name, dir_ix, archived. **נשמר 30 יום ואז עובר prune אוטומטי** (השורות הישנות מיוצאות ל-Drive כ-CSV לפני המחיקה).' },

    { type: ROW, a: '💎 _baseline_archive',
      b: 'מוסתר. **לעולם לא נמחק** — מקור האמת לכל baseline היסטורי. רשומה מצטברת פר (route, dir, date, hour) עם n, sum_delay_s, sum_speed, sum_level, last_updated. מתעדכן upsert בכל upload חדש. גם אחרי שה-raw_data של פקק מסוים נמחק, המונה שלו ממשיך להשפיע על ה-baseline.' },

    { type: ROW, a: '📋 מקור',
      b: 'גלוי. יומן הופעות snapshot — חותמת זמן, תאריך, שעה, מספר פקקים בכל upload. משמש לזיהוי כפילויות (אם startTime זהה כבר קיים, ה-upload נדחה). נשמר 30 יום.' },

    { type: ROW, a: '📁 ארכיון נפרד',
      b: 'גלוי. יומן של הייצואים ל-Drive (שם קובץ, טווח תאריכים, קישור). מתעדכן בכל ייצוא של 30 הימים שעוברים prune או ייצוא ידני.' },

    { type: ROW, a: '🔒 _config (מוסתר)',
      b: 'תצורת auto-fetch: fetch_url, fetch_headers, fetch_interval, trigger_owner. נשמר ברמת הגיליון כדי לשרוד החלפת מחשבים/חשבונות.' },

    { type: ROW, a: '🔒 _fetch_log (מוסתר)',
      b: 'יומן ריצות של ה-auto-fetch — חותמת זמן, משתמש, סטטוס (הצלחה/כפילות/שגיאה), מספר פקקים, הודעת שגיאה. שימושי לדיבאג של ה-cron.' },

    { type: ROW, a: '🔒 _filter (מוסתר)',
      b: 'הפילטר האחרון שהוחל (JSON blob: fromDate, toDate, fromHour, toHour, days). נדרס בכל "החל סינון".' },

    { type: ROW, a: 'עיקרון מפתח',
      b: 'הסינון מסנן את raw_data בלבד. ה-baseline-ים תמיד מגיעים מ-_baseline_archive המלא, ללא קשר לפילטר. כך "מה שאתה רואה בטווח שבחרת" מושווה תמיד מול "כל מה שהמערכת יודעת מההיסטוריה".' },

    { type: BLANK },
    { type: SEC, text: '§4. איך מחושב baseline היסטורי' },
    { type: ROW, a: 'מפתח', b: 'מסלול × כיוון × שעה × סוג-יום (חול=א\'–ה\' · סופ"ש=ו\'–ש\').' },
    { type: ROW, a: 'מקור', b: 'סכימה מצטברת מתוך _baseline_archive (כל ההיסטוריה, לא רק 30 יום).' },
    { type: ROW, a: 'מינימום דגימות', b: 'n ≥ 3 לתא. תאים עם פחות נחשבים "ריקים".' },

    { type: BLANK },
    { type: SEC, text: '§5. מקור ההשוואה (Fallback)' },
    { type: ROW, a: '"שעה זו"', b: 'יש n ≥ 3 בתא המדויק (אותה שעה ואותו סוג יום). השוואה מדויקת.' },
    { type: ROW, a: '"±1 שעות"', b: 'התא המדויק רֵיק. הורחב לחלון של 3 שעות סביב (H−1, H, H+1) באותו סוג יום.' },
    { type: ROW, a: '"±2 שעות"', b: 'גם ±1 לא הספיק. הורחב לחלון של 5 שעות (H−2..H+2).' },
    { type: ROW, a: '"אין מספיק נתונים"', b: 'אפילו ±2 לא נתן 3 דגימות. אין השוואה, אין סטטוס. (במקום fallback כללי שהיה יוצר false positives.)' },
    { type: ROW, a: 'הימנעות מהצלבה', b: 'ימי חול לעולם לא משווים מול דגימות סופ"ש, גם בהרחבה.' },

    { type: BLANK },
    { type: SEC, text: '§6. חישוב הסטייה' },
    { type: ROW, a: 'נוסחה', b: 'סטייה% = (ממוצע_נוכחי − ממוצע_היסטורי) ÷ ממוצע_היסטורי × 100' },
    { type: ROW, a: 'יחידות', b: 'הממוצע מחושב בשניות (delay_s) לפקק בודד. ההמרה לדקות נעשית רק לתצוגה.' },
    { type: ROW, a: 'משמעות', b: '+30% = הפקקים גרועים ב-30% מהמצב הרגיל באותה שעה. −15% = טובים ב-15%.' },

    { type: BLANK },
    { type: SEC, text: '§7. ספי סטטוס' },
    { type: ROW, a: '🟢 תקין', b: '|סטייה| ≤ 10%' },
    { type: ROW, a: '🟡 מתון', b: 'סטייה +10% עד +25%' },
    { type: ROW, a: '🟠 עמוס', b: 'סטייה +25% עד +50%' },
    { type: ROW, a: '🔴 חריג מאוד', b: 'סטייה > +50%' },
    { type: ROW, a: '🔵 טוב מהרגיל', b: 'סטייה < −10%' },
    { type: ROW, a: '⚪ אין מספיק נתונים', b: 'אין baseline היסטורי לתא — לא ניתן לקבוע סטטוס.' },

    { type: BLANK },
    { type: SEC, text: '§8. זיהוי חריגות (לשונית "חריגות")' },
    { type: ROW, a: 'שיטה', b: 'שונה מ-baseline היסטורי: סטטיסטיקה מקומית **בתוך** הפילטר הנוכחי.' },
    { type: ROW, a: 'קריטריון', b: 'פקק נחשב חריג אם: delay > mean + 1.5σ (לאותו מסלול×כיוון), או speed < 5, או level ≥ 4.' },
    { type: ROW, a: 'חומרה', b: 'קריטי (speed<3 או סטייה>200%) · גבוה (level≥4 או סטייה>100%) · בינוני (השאר).' },
    { type: ROW, a: 'למה שתי שיטות', b: 'ה-baseline שואל: "האם זה גרוע מהרגיל?". החריגות שואלות: "אילו פקקים בולטים מבין מה שראינו השבוע?". משלימות.' },

    { type: BLANK },
    { type: SEC, text: '§9. סולם צבעים בעמודות סטייה' },
    { type: ROW, a: '🟢 ירוק', b: 'בין −10% ל-+10% (תקין)' },
    { type: ROW, a: '🟡 צהוב', b: '+10% עד +25%' },
    { type: ROW, a: '🟠 כתום', b: '+25% עד +50%' },
    { type: ROW, a: '🔴 אדום', b: 'מעל +50%' },
    { type: ROW, a: '🔵 כחול', b: 'מתחת ל-−10% (פחות עומס מהרגיל)' },

    { type: BLANK },
    { type: SEC, text: '§10. מגבלות המתודולוגיה' },
    { type: ROW, a: 'חודש ראשון', b: 'רוב התאים יתפסו ב-±1 או ±2 שעות. ככל שהארכיון מתבגר, יותר תאים יקבלו "שעה זו". זה צפוי.' },
    { type: ROW, a: 'חגים ואירועים', b: 'הארכיון לא מבחין בין יום רגיל ליום חג. דגימת חג הופכת לדגימה רגילה בסטטיסטיקה.' },
    { type: ROW, a: 'משקלול לפי עדכניות', b: 'אין. דגימה מלפני שנה שווה בערכה לדגימה מהשבוע.' },
    { type: ROW, a: 'free_flow_min', b: 'הערכה ידנית של זמן נסיעה ללא פקקים. משמש לעמודת ייחוס בלבד, לא לחישוב סטייה.' },

    { type: BLANK },
    { type: SEC, text: '§11. רענון הנתונים' },
    { type: ROW, a: 'בכל upload', b: 'מתווסף ל-raw_data ו-upsert ל-_baseline_archive בו זמנית.' },
    { type: ROW, a: 'pruning', b: 'raw_data מנקה אוטומטית שורות > 30 יום (אחרי שמייצאות ל-Drive כ-CSV). _baseline_archive תמיד נשמר.' },
    { type: ROW, a: 'מיגרציה', b: 'תפריט "🚦 Waze → 🔄 העבר נתונים קיימים לארכיון" — קולט שורות ישנות שעדיין לא נספרו לארכיון.' },

    { type: BLANK },
    { type: SEC, text: '§12. מילון מונחים מקוצר' },
    { type: ROW, a: 'daytype', b: 'סוג יום: weekday (חול) או weekend (סופ"ש).' },
    { type: ROW, a: 'jam', b: 'פקק יחיד באירוע יחיד — שורה אחת ב-raw_data.' },
    { type: ROW, a: 'snapshot', b: 'תמונת מצב אחת מ-Waze (עשרות-מאות פקקים).' },
    { type: ROW, a: 'baseline', b: 'הממוצע ההיסטורי שנגדו משווים.' },
    { type: ROW, a: 'upsert', b: 'update-or-insert: אם המפתח קיים — מצטבר ל-counters; אחרת — שורה חדשה.' },
    { type: ROW, a: 'prune', b: 'מחיקת שורות ישנות מ-raw_data (אחרי 30 יום).' },
  ];

  var SEC_BG = '#1F3864', SEC_FG = '#FFFFFF';
  var HDR_BG = '#475569', HDR_FG = '#FFFFFF';
  var ALT_BG = '#F8FAFC';

  var r = 1;
  items.forEach(function(item, i) {
    if (item.type === SEC) {
      ws.getRange(r, 1, 1, 2).merge().setValue(item.text)
        .setBackground(SEC_BG).setFontColor(SEC_FG)
        .setFontWeight('bold').setFontSize(13)
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
      ws.setRowHeight(r, 30);
    } else if (item.type === HDR) {
      ws.getRange(r, 1, 1, 2).setValues([[item.a, item.b]])
        .setBackground(HDR_BG).setFontColor(HDR_FG)
        .setFontWeight('bold').setHorizontalAlignment('center');
    } else if (item.type === ROW) {
      var rng = ws.getRange(r, 1, 1, 2).setValues([[item.a, item.b]])
        .setFontSize(11).setVerticalAlignment('top').setWrap(true);
      ws.getRange(r, 1).setFontWeight('bold').setHorizontalAlignment('right');
      ws.getRange(r, 2).setHorizontalAlignment('right');
      if (i % 2 === 0) rng.setBackground(ALT_BG);
    } else if (item.type === BLANK) {
      ws.setRowHeight(r, 12);
    }
    r++;
  });

  ws.setColumnWidth(1, 220);
  ws.setColumnWidth(2, 520);
  ws.setFrozenRows(0);
}

// ═══ AGGREGATION VIEW (reads from permanent _baseline_archive) ═══
// Shows trends across the FULL history (not the 30-day raw_data window).
function _rebuildAggregation(ss) {
  var arch = ss.getSheetByName(BASELINE_ARCHIVE);
  if (!arch || arch.getLastRow() < 2) return;

  var ws = _newSheet(ss, 'אגרגציה לאורך זמן', '#7030A0');
  var ncW = 9;
  ws.getRange(1,1,1,ncW).merge()
    .setValue('מגמות לאורך זמן — מבוסס על הארכיון הקבוע (כל ההיסטוריה)')
    .setFontSize(15).setFontWeight('bold').setFontColor('#1F3864').setHorizontalAlignment('center');
  ws.setRowHeight(1, 32);

  // intro row 2 written AFTER _autoWidth at the end

  var rows = arch.getRange(2, 1, arch.getLastRow() - 1, BASELINE_COLS.length).getValues();

  // resolve dir labels
  var dirLabel = {};
  ROUTES.forEach(function(r) {
    dirLabel[r.name + '::1'] = r.dir1_label || '';
    dirLabel[r.name + '::2'] = r.dir2_label || '';
  });

  // Sort: date desc, route asc, dir asc, hour asc
  rows.sort(function(a, b) {
    var da = a[2] instanceof Date ? a[2].getTime() : new Date(a[2]).getTime();
    var db = b[2] instanceof Date ? b[2].getTime() : new Date(b[2]).getTime();
    if (da !== db) return db - da;
    if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
    if (a[1] !== b[1]) return a[1] - b[1];
    return a[3] - b[3];
  });

  var cols = ['תאריך','סוג יום','מסלול','כיוון','שעה','מס\' פקקים',
              'השהיה ממוצעת (דק\')','מהירות ממוצעת','רמה ממוצעת'];
  _hdrRow(ws, cols, 3);

  var row = 4;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var route=r[0], dir=r[1], date=r[2], hour=r[3];
    var n=+r[4]||0, sumD=+r[5]||0, sumS=+r[6]||0, sumL=+r[7]||0;
    if (!n) continue;
    var dateStr = date instanceof Date
      ? Utilities.formatDate(date, 'Asia/Jerusalem', 'yyyy-MM-dd')
      : String(date);
    var dt = _dayTypeFromDate(dateStr);
    var avgD = _round1((sumD / n) / 60);
    var avgS = _round1(sumS / n);
    var avgL = _round1(sumL / n);
    _dataRow(ws, row, [
      dateStr, dt === 'weekend' ? 'סופ"ש' : 'חול',
      route, dirLabel[route+'::'+dir] || ('כיוון '+dir),
      _pad(hour)+':00',
      n, avgD, avgS, avgL
    ], row % 2 === 0);
    row++;
    if (row > 5000) break;   // protect against huge archive — show most recent 5K rows
  }

  _autoWidth(ws, cols.length);
  _tabIntro(ws, cols.length, 'תצוגה ישירה של הארכיון הקבוע _baseline_archive (כל ההיסטוריה, לא רק 30 הימים האחרונים). שורה לכל תאריך × מסלול × כיוון × שעה. ממוין מהחדש לישן. מציג עד 5,000 שורות.', 2);
  ws.setFrozenRows(3);
}
