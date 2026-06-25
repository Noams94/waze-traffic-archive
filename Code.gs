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
// Safety cap: Google Sheets enforces ~10M cells per spreadsheet. With 19 RAW_COLS
// that's ~525K rows before raw_data alone hits the wall — BUT Sheets allocates
// a small column buffer (typically 2 trailing empty cols), pushing the effective
// width to 21. To leave breathing room for the buffer + ~200K cells used by the
// other sheets, trigger emergency prune well below the theoretical max.
// At 21 cols/row: 450K × 21 = 9.45M cells, leaving ~350K cells of headroom.
var MAX_RAW_ROWS = 450000;

// ─── National Traffic Index (NCI) ─────────────
// One national number per rush window = jam-weighted average of the per-route
// "deviation %" (current avg-delay-per-jam vs the permanent _baseline_archive),
// with hours bucketed into the windows below. + = heavier than routine, − = lighter.
var NCI_WINDOWS = [
  { key: 'בוקר', hours: [6, 7, 8, 9], runAtHour: 10 },   // 06:00–10:00 → run ~10:00
  { key: 'ערב',  hours: [16, 17, 18], runAtHour: 20 },   // 16:00–19:00 → run ~20:00
];
var NCI_SHEET           = '🚦 מדד ארצי';
var NCI_TIMEFRAME_SHEET = 'פירוט לפי מסגרת זמן';
var NCI_HISTORY_SHEET   = '_nci_history';   // hidden, permanent — survives the 30-day prune
var NCI_HISTORY_COLS    = ['run_ts','date','window','daytype','index_pct','n_jams','n_routes','period'];
var NCI_TRIGGER_FN      = '_scheduledNCI';
var NCI_TREND_POINTS    = 14;   // how many recent readings the slide's trend shows
var NCI_EMAIL_CONFIG_KEY = 'nci_email';      // 'on' (default) | 'off' — email the index after each run
var NCI_EMAIL_TO_KEY     = 'nci_email_to';   // comma/semicolon-separated recipients; blank = sheet owner
var NCI_EMAIL_TREND      = 7;   // recent same-window readings shown in the email trend
var NCI_MAP_KEY_CONFIG_KEY = 'nci_geoapify_key';   // Geoapify Static Maps key (free, no card); blank = email sent without the map image
// Published /exec URL of the anonymous "Map" web-app deployment. Hardcoded because
// ScriptApp.getService().getUrl() can return the owner-only /dev URL (multi-login
// fails) and the bare /exec can serve a stale edge-cached response. _nciMapUrl()
// appends a cache-busting token so menu/email links always load fresh.
var NCI_MAP_WEBAPP_URL   = 'https://script.google.com/macros/s/AKfycby0PoDWYHg2iMSnmBZN1BEm8LRZRRUQ4fbn-3qPrWmfBB2-oP3Dba-XaK6B0pawFQDR_g/exec';

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
    .addItem('🔴 סווג חירום/שגרה לנתונים קיימים', 'menuBackfillPeriod')
    .addSeparator()
    .addItem('📊 בדוק שימוש בתאים', 'menuDiagnoseCellUsage')
    .addItem('🗜️ דחס תאים ריקים', 'menuCompactSheets')
    .addItem('🔃 תקן סדר raw_data (לפי זמן)', 'menuRepairRawDataOrder');

  ui.createMenu('🚦 Waze')
    .addItem('פתח סרגל צד...', 'showSidebar')
    .addSeparator()
    .addItem('📥 ייצא חדשים ל-Drive', 'menuExportArchive')
    .addItem('🗓️ הפעל/כבה גיבוי יומי', 'menuToggleDailyArchive')
    .addItem('🧹 קצץ raw_data עכשיו', 'menuPruneNow')
    .addSeparator()
    .addItem('🚦 חשב מדד ארצי עכשיו', 'menuRunNCINow')
    .addItem('⏰ התקן תזמון מדד ארצי (בוקר+ערב)', 'menuInstallNCITriggers')
    .addItem('📧 הפעל/כבה מייל מדד ארצי', 'menuToggleNCIEmail')
    .addItem('✉️ הגדר נמעני מייל מדד...', 'menuSetNCIEmailRecipients')
    .addItem('📤 שלח מייל מדד עכשיו (בדיקה)', 'menuSendNCIEmailTest')
    .addSeparator()
    .addItem('🗺️ פתח מפת חום אינטראקטיבית', 'menuOpenCongestionMap')
    .addItem('🔑 הגדר מפתח מפת חום למייל (Geoapify)...', 'menuSetNCIMapApiKey')
    .addItem('🧪 בדוק תמונת מפה למייל', 'menuTestNCIMapImage')
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

  // Auto-install the national-index triggers (morning+evening); skipped if disabled.
  try {
    if (_getConfig('nci_schedule') !== 'off' && !_hasNCITrigger()) {
      installNCITriggers();
      _setConfig('nci_schedule', 'on');
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

// Backfill the 'period' column for raw_data rows imported before the column
// existed. Re-derives חירום/שגרה from each row's pub_ts. Idempotent — safe to
// re-run (e.g. after declaring a new emergency window in EMERGENCY_WINDOWS).
function menuBackfillPeriod() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('סיווג חירום/שגרה',
    'למלא את עמודת "period" (חירום/שגרה) לכל השורות הקיימות ב-raw_data לפי זמן הפקק?\n\n' +
    'בטוח להרצה חוזרת. ירוץ גם על שורות שכבר סווגו (לעדכון אחרי שינוי חלוני חירום).',
    ui.ButtonSet.OK_CANCEL);
  if (resp !== ui.Button.OK) return;
  try {
    var n = _backfillRawPeriod(SpreadsheetApp.getActiveSpreadsheet());
    ui.alert('הושלם', 'סווגו ' + n + ' שורות ב-raw_data.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('שגיאה', 'הסיווג נכשל:\n' + e.message, ui.ButtonSet.OK);
  }
}
function _backfillRawPeriod(ss) {
  var s = ss.getSheetByName(RAW_SHEET);
  if (!s) return 0;
  var col = RAW_COLS.indexOf('period') + 1;   // 1-based; column may not exist yet on old sheets
  // Ensure the header cell exists/labelled (old sheets froze at 18 cols).
  s.getRange(1, col).setValue('period')
   .setBackground('#1F3864').setFontColor('#fff').setFontWeight('bold');
  var n = s.getLastRow() - 1;
  if (n < 1) return 0;
  // pub_ts lives in column 2; derive period from it in one batched read/write.
  var ts = _retry(function() { return s.getRange(2, 2, n, 1).getValues(); });
  var out = ts.map(function(r) { return [ _periodFromTs(r[0]) ]; });
  _retry(function() { s.getRange(2, col, n, 1).setValues(out); });
  return n;
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

// Restore raw_data to chronological order (oldest → newest) by sorting on
// snapshot_ts. The whole pipeline assumes new rows are appended at the bottom,
// so a manual column sort (or any reordering) silently breaks the index and the
// prune boundary. This re-establishes the invariant in one pass.
function menuRepairRawDataOrder() {
  var ui = SpreadsheetApp.getUi();
  if (_getArchiveJob()) {
    ui.alert('ייצוא פעיל', 'יש ג\'וב ייצוא/קציצה בתהליך. המתן לסיומו ואז נסה שוב.', ui.ButtonSet.OK);
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName(RAW_SHEET);
  if (!raw || raw.getLastRow() < 3) {
    ui.alert('אין מה לתקן', 'raw_data ריק או קצר מדי.', ui.ButtonSet.OK);
    return;
  }
  var n = raw.getLastRow() - 1;
  var resp = ui.alert('תיקון סדר raw_data',
    'ימוין מחדש ' + n + ' שורות לפי זמן (snapshot_ts) בסדר עולה, כך שהנתון החדש ביותר חוזר לתחתית.\n\n' +
    'הנתונים עצמם לא משתנים — רק הסדר. להמשיך?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  try {
    // Sort the data range (excluding the header) by column 1 ascending.
    raw.getRange(2, 1, n, RAW_COLS.length).sort({ column: 1, ascending: true });
    var last = _ymd(raw.getRange(raw.getLastRow(), 3).getValue());
    ui.alert('הסדר תוקן', 'raw_data מוין לפי זמן. התאריך האחרון כעת: ' + last +
      '\n\nכעת הרץ "🚦 חשב מדד ארצי עכשיו" כדי לרענן את המדד.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('התיקון נכשל', e.message +
      '\n\nאפשר גם למיין ידנית: בחר את raw_data ← נתונים ← מיון טווח ← לפי עמודה A, עולה.', ui.ButtonSet.OK);
  }
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
    // National index tabs — filter-independent (latest complete window vs the
    // permanent baseline). Rendered here too so they refresh on every manual rebuild.
    // Isolated in its own try so a failure here can't break the core filter flow.
    try {
      var nci = _nciData(ss, baselines);
      if (nci) {
        _sheetTimeframes(ss, nci, baselines);
        _sheetNCI(ss, nci, _readNCIHistory(ss, NCI_TREND_POINTS));
      }
    } catch (nciErr) {}
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
      // Emergency periods are anomalous by definition — keep them out of the
      // "routine" baseline so deviation % is measured against normal traffic.
      if (_periodFromDateHour(date, hour) === PERIOD_EMERGENCY) return;
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
  var result;
  try {
    result = fetchFromUrl(url, headers);
  } catch (e) {
    // fetchFromUrl converts almost every failure into a result object, but a
    // throw here (e.g. a post-fetch helper) would otherwise leave NO row in
    // _fetch_log and surface only in the Apps Script execution log. Capture it
    // as a real error row so the log never looks clean when a run actually failed.
    result = { ok: false, error: 'חריגה לא צפויה: ' + ((e && e.message) || e) };
  }
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
  } catch (e) {
    // Even writing the log failed (sheet locked, quota, etc.). Don't swallow it
    // silently — make it visible in the execution log so a stalled schedule is
    // diagnosable instead of looking idle.
    console.error('_scheduledFetch: failed to write _fetch_log: ' + ((e && e.message) || e));
  }
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
                'start_node','end_node','tt_min','route_name','dir_ix','archived',
                'period'];   // חירום/שגרה — appended last so positional indices 0..17 are unchanged

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
      _periodFromTs(dt),                                 // period: חירום/שגרה by jam time
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
  {section:'דרום',name:'כביש 41',from:'אשדוד',to:'ראשל"צ',distance_km:20,free_flow_min:18,streets_dir1:['41 מזרח'],streets_dir2:['41 מערב'],dir1_label:'אשדוד → ראשל"צ (מזרח)',dir2_label:'ראשל"צ → אשדוד (מערב)'},
  {section:'מרכז',name:'כביש 44',from:'אשדוד',to:'מודיעין',distance_km:40,free_flow_min:35,streets_dir1:['44 צפון'],streets_dir2:['44 דרום'],dir1_label:'אשדוד → מודיעין (צפון)',dir2_label:'מודיעין → אשדוד (דרום)'},
  {section:'מרכז',name:'כביש 60',from:'באר שבע',to:'נצרת',distance_km:200,free_flow_min:180,streets_dir1:['60'],streets_dir2:[],dir1_label:'שני הכיוונים (מעורב)',dir2_label:''},
  {section:'מרכז',name:'כביש 444',from:'ראש העין',to:'נחשונים',distance_km:15,free_flow_min:12,streets_dir1:['444 צפון'],streets_dir2:['444 דרום'],dir1_label:'נחשונים → ר"ע (צפון)',dir2_label:'ר"ע → נחשונים (דרום)'},
  {section:'מרכז',name:'כביש 461',from:'אור יהודה',to:'יהוד',distance_km:8,free_flow_min:8,streets_dir1:['461 מזרח'],streets_dir2:['461 מערב'],dir1_label:'אור יהודה → יהוד (מזרח)',dir2_label:'יהוד → אור יהודה (מערב)'},
  {section:'צפון',name:'כביש 22',from:'חיפה מפרץ',to:'חיפה כרמל',distance_km:12,free_flow_min:12,streets_dir1:['22 צפון'],streets_dir2:['22 דרום'],dir1_label:'דרום → צפון',dir2_label:'צפון → דרום'},
  {section:'מרכז',name:'כביש 57',from:'נתניה',to:'טול כרם',distance_km:22,free_flow_min:20,streets_dir1:['57 מזרח'],streets_dir2:['57 מערב'],dir1_label:'נתניה → מזרח',dir2_label:'מזרח → נתניה'},
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

// Region polygons (lat/lon rings) for the regional heatmap, keyed by the
// ROUTES[].section value. These follow Israel's official administrative districts
// grouped into three: צפון = Northern + Haifa, מרכז = Central + Tel Aviv + Jerusalem
// (the eastern bulge is the Jerusalem corridor), דרום = Southern (Ashdod and south).
// Simplified — coastal split lines verified against the districts (Center/North ≈
// Hadera 32.38, Center/South ≈ north of Ashdod 31.85); the eastern edge follows the
// Jordan rift rather than cutting the Green Line. Keys must match the section strings.
var REGION_GEO = {
  'צפון': [[32.38,34.88],[32.55,34.90],[32.83,34.97],[33.08,35.10],[33.28,35.58],[33.10,35.65],[32.70,35.62],[32.50,35.50],[32.38,35.28]],
  'מרכז': [[32.38,34.88],[32.38,35.28],[32.10,35.40],[31.85,35.35],[31.70,35.30],[31.70,34.95],[31.85,34.80],[31.85,34.68],[32.08,34.77],[32.33,34.85]],
  'דרום': [[31.85,34.68],[31.85,34.80],[31.70,34.95],[31.70,35.30],[31.50,35.45],[31.00,35.40],[30.10,35.10],[29.55,34.95],[30.50,34.43],[31.10,34.30],[31.55,34.45]],
};

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

// ─── Period: חירום (emergency) vs שגרה (routine) ───────────────────
// A SECOND, independent classification of every jam — by WHEN it happened, not by
// day-of-week. Unlike daytype (derived), periods are explicit date-time windows
// declared in EMERGENCY_WINDOWS below. Anything outside every window is שגרה.
// Windows are half-open [from, to) in local time (Asia/Hebron ≈ Asia/Jerusalem),
// written as 'yyyy-MM-dd HH:mm' so a plain lexicographic string compare is exact
// and timezone-clean (same fixed-width format the timestamps are formatted into).
// To declare a new emergency, add another { from, to } entry.
var PERIOD_EMERGENCY = 'חירום';
var PERIOD_ROUTINE   = 'שגרה';
var EMERGENCY_WINDOWS = [
  { from: '2026-06-07 22:00', to: '2026-06-09 06:00' },
];

// Core test: a local 'yyyy-MM-dd HH:mm' key → 'חירום' | 'שגרה'.
function _periodFromKey(key) {
  for (var i = 0; i < EMERGENCY_WINDOWS.length; i++) {
    var w = EMERGENCY_WINDOWS[i];
    if (key >= w.from && key < w.to) return PERIOD_EMERGENCY;
  }
  return PERIOD_ROUTINE;
}
// From a full timestamp (Date / millis / ISO string) — minute resolution.
function _periodFromTs(ts) {
  if (ts === null || ts === undefined || ts === '') return PERIOD_ROUTINE;
  var d = (ts instanceof Date) ? ts : new Date(ts);
  if (isNaN(d.getTime())) return PERIOD_ROUTINE;
  return _periodFromKey(Utilities.formatDate(d, 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm'));
}
// From a date (string/Date) + integer hour — for _baseline_archive, which stores
// date+hour rather than a full timestamp. Resolves at HH:00.
function _periodFromDateHour(dateStr, hour) {
  var ymd = _ymd(dateStr);
  if (!ymd) return PERIOD_ROUTINE;
  var h = parseInt(hour, 10); if (isNaN(h)) h = 0;
  return _periodFromKey(ymd + ' ' + _pad(h) + ':00');
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
// ═══ NATIONAL TRAFFIC INDEX (NCI) ═════════════════════════════════
// Reuses the exact comparison logic of "פירוט לפי שעה" (current avg-delay-per-jam
// vs the permanent _baseline_archive) but buckets hours into NCI_WINDOWS, and
// collapses each window into one jam-weighted national "deviation %" number.

// Normalize a raw date cell to 'yyyy-MM-dd' whether it's stored as a string or a Date.
function _ymd(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Jerusalem', 'yyyy-MM-dd');
  return ('' + v).slice(0, 10);
}
// Retry wrapper for transient "Service Spreadsheets timed out" errors on big docs.
function _retry(fn, tries) {
  tries = tries || 3;
  for (var i = 0; i < tries; i++) {
    try { return fn(); }
    catch (e) {
      if (i === tries - 1) throw e;
      Utilities.sleep(1500 * (i + 1));
    }
  }
}
function _nciFmtPct(p) {
  if (p == null) return '—';
  return (p >= 0 ? '+' : '') + _round1(p) + '%';
}
// Same value, isolated LTR so the minus stays on the LEFT of the number inside
// RTL HTML (email). Use only in HTML contexts — not in plain sheet cells.
function _nciFmtPctLtr(p) {
  if (p == null) return '—';
  return '<span dir="ltr" style="unicode-bidi:isolate">' + _nciFmtPct(p) + '</span>';
}
function _winRangeLabel(key) {
  for (var i = 0; i < NCI_WINDOWS.length; i++) {
    if (NCI_WINDOWS[i].key === key) {
      var h = NCI_WINDOWS[i].hours;
      return _pad(h[0]) + ':00–' + _pad(h[h.length - 1] + 1) + ':00';
    }
  }
  return key;
}
// Deviation% → { label (status word), head (slide headline), bg, fg }
function _nciStatus(dev) {
  if (dev == null)  return { label: 'אין נתונים',  head: 'אין מספיק נתונים',        bg: '#F1F5F9', fg: '#475569' };
  if (dev > 50)     return { label: 'חריג מאוד',    head: 'תנועה כבדה מאוד מהרגיל',  bg: C_RED_BG,  fg: C_RED_FG };
  if (dev > 25)     return { label: 'עמוס',         head: 'תנועה כבדה מהרגיל',        bg: C_ORG_BG,  fg: C_ORG_FG };
  if (dev > 10)     return { label: 'מתון',         head: 'כבד מעט מהרגיל',           bg: C_YEL_BG,  fg: C_YEL_FG };
  if (dev >= -10)   return { label: 'תקין',         head: 'תואם שגרה',                bg: C_GRN_BG,  fg: C_GRN_FG };
  return                   { label: 'טוב מהרגיל',   head: 'קל מהרגיל',                bg: '#DBEAFE', fg: '#1E40AF' };
}

// A national index built on too few jams/routes is statistically noisy — the
// number is jam-weighted, so on a thin window a handful of quiet routes can swing
// it ±30% (e.g. a quiet Saturday evening with ~50 jams). Returns a caveat string
// when the sample is thin, or '' when it's solid. Used as a side-note NEXT TO the
// number — never to hide it. Weekend windows are inherently lighter, so they trip
// this more often than weekdays, which is intended.
var NCI_MIN_JAMS = 80, NCI_MIN_ROUTES = 12;
function _nciLowConfidence(win) {
  if (!win || win.indexPct == null) return '';
  if (win.nJams < NCI_MIN_JAMS || win.nRoutes < NCI_MIN_ROUTES) {
    return 'מדגם קטן (' + win.nJams + ' פקקים, ' + win.nRoutes + ' מסלולים) — המדד עשוי להיות פחות יציב';
  }
  return '';
}

// Core: read the latest date's window jams (filter-independent), compute the
// per-route deviation and the jam-weighted national index per window.
function _nciData(ss, baselines) {
  var s = ss.getSheetByName(RAW_SHEET);
  if (!s || s.getLastRow() < 2) return null;
  var n = s.getLastRow() - 1;
  // Normally the latest date's rows live at the bottom, so reading the tail avoids
  // a full scan. But raw_data can fall out of chronological order (e.g. a manual
  // column sort), which would strand an old row at the bottom. So don't trust the
  // last physical row — scan the read window for the MAX date and use that.
  var readN = Math.min(n, 30000);
  var values = _retry(function() {
    return s.getRange(2 + (n - readN), 1, readN, RAW_COLS.length).getValues();
  });

  var date = '', dayName = '';
  for (var di = 0; di < readN; di++) {
    var d = _ymd(values[di][2]);
    if (d > date) { date = d; dayName = values[di][3]; }
  }
  var daytype = _dayType(dayName);

  var winByHour = {};
  NCI_WINDOWS.forEach(function(w) { w.hours.forEach(function(h) { winByHour[h] = w.key; }); });

  var routeMeta = {};
  ROUTES.forEach(function(rt) {
    routeMeta[rt.name + '::1'] = { section: rt.section, route: rt.name, dir: rt.dir1_label, dirIx: 1 };
    if (rt.dir2_label) routeMeta[rt.name + '::2'] = { section: rt.section, route: rt.name, dir: rt.dir2_label, dirIx: 2 };
  });

  var groups = {};   // window → routeKey → { meta, hours: { hour: {jams, sumDelay} } }
  for (var i = 0; i < readN; i++) {
    var r = values[i];
    if (_ymd(r[2]) !== date) continue;
    var wk = winByHour[r[4]];
    if (!wk) continue;
    var routeName = r[15] || '';
    if (!routeName) continue;
    var meta = routeMeta[routeName + '::' + r[16]];
    if (!meta) continue;
    var g = groups[wk] || (groups[wk] = {});
    var cell = g[routeName + '::' + r[16]] || (g[routeName + '::' + r[16]] = { meta: meta, hours: {} });
    var hh = cell.hours[r[4]] || (cell.hours[r[4]] = { jams: 0, sumDelay: 0 });
    hh.jams++;
    hh.sumDelay += (+r[10] || 0);
  }

  // National index = jam-weighted average of the per-HOUR "deviation %" values
  // (identical to "פירוט לפי שעה", via avgPerJamAtHour), bucketed into the window.
  var windows = {}, order = [];
  NCI_WINDOWS.forEach(function(w) {
    var g = groups[w.key];
    if (!g) return;
    var rows = [], totDevNum = 0, totDevJams = 0, contribCells = 0, totJams = 0;
    Object.keys(g).forEach(function(rk) {
      var cell = g[rk];
      var rDevNum = 0, rDevJams = 0, rJams = 0, srcCounts = {};
      Object.keys(cell.hours).forEach(function(hk) {
        var o = cell.hours[hk];
        rJams += o.jams;
        var base = baselines && baselines.avgPerJamAtHour(cell.meta.route, cell.meta.dirIx, parseInt(hk, 10), daytype);
        if (base && base.avg > 0) {
          var dev = (o.sumDelay / o.jams - base.avg) / base.avg * 100;
          rDevNum += dev * o.jams; rDevJams += o.jams;
          srcCounts[base.source] = (srcCounts[base.source] || 0) + o.jams;
        }
      });
      var rDev = rDevJams > 0 ? _round1(rDevNum / rDevJams) : null;
      var src = '—', bestN = 0;
      Object.keys(srcCounts).forEach(function(s) { if (srcCounts[s] > bestN) { src = s; bestN = srcCounts[s]; } });
      rows.push({
        section: cell.meta.section, route: cell.meta.route, dir: cell.meta.dir,
        jams: rJams, devPct: rDev, source: rDev != null ? src : '—',
      });
      if (rDev != null) { totDevNum += rDevNum; totDevJams += rDevJams; contribCells++; }
      totJams += rJams;
    });
    rows.sort(function(a, b) {
      var x = a.devPct == null ? -1e9 : a.devPct, y = b.devPct == null ? -1e9 : b.devPct;
      return y - x;
    });
    windows[w.key] = {
      key: w.key, hours: w.hours, rows: rows,
      indexPct: totDevJams > 0 ? _round1(totDevNum / totDevJams) : null,
      nJams: totJams, nRoutes: contribCells,
      period: _periodFromDateHour(date, w.hours[0]),   // חירום/שגרה for this window
    };
    order.push(w.key);
  });

  var headline = order.length ? windows[order[order.length - 1]] : null;
  return {
    date: date, dayName: dayName, daytype: daytype,
    period: headline ? headline.period : _periodFromDateHour(date, 12),
    windows: windows, order: order,
    headline: headline,
  };
}

// ─── Heatmap: per-region color data + interactive map ───

// Collapse win.rows to one deviation value per REGION (section: דרום/מרכז/צפון).
// Jam-weighted average across the region's routes — consistent with how the
// national index itself is weighted in _nciData. devPct is null when no route in
// the region had a usable baseline that window (→ rendered grey). Returns:
//   { 'מרכז': { devPct: <num|null>, jams: <int> }, ... }
function _nciRegionDeviations(win) {
  var acc = {};
  ((win && win.rows) || []).forEach(function(r) {
    var sec = r.section || '—';
    var a = acc[sec] || (acc[sec] = { num: 0, wj: 0, jams: 0, anyDev: false });
    a.jams += (r.jams || 0);
    if (r.devPct != null) { a.num += r.devPct * (r.jams || 1); a.wj += (r.jams || 1); a.anyDev = true; }
  });
  var out = {};
  Object.keys(acc).forEach(function(sec) {
    var a = acc[sec];
    out[sec] = { devPct: a.anyDev && a.wj > 0 ? _round1(a.num / a.wj) : null, jams: a.jams };
  });
  return out;
}

// Build a Geoapify Static Maps URL: region polygons filled by their deviation
// color (via _nciStatus) over a real OSM map. Geoapify free tier needs only a key
// (no credit card). Returns '' when no key is configured. `names` limits which
// regions to draw (defaults to all) — used by the per-region diagnostic.
function _nciStaticMapUrl(win, names) {
  var key = _getConfig(NCI_MAP_KEY_CONFIG_KEY);
  if (!key) return '';
  var devs = _nciRegionDeviations(win);
  var geoms = [];
  (names || Object.keys(REGION_GEO)).forEach(function(name) {
    var ring = REGION_GEO[name];
    if (!ring || !ring.length) return;
    var dev = devs[name] ? devs[name].devPct : null;
    var hex = _nciStatus(dev).fg.replace('#', '').toLowerCase();   // 6 lowercase hex digits (Geoapify is strict)
    var pts = ring.map(function(p) { return p[1] + ',' + p[0]; });   // Geoapify wants lon,lat
    pts.push(ring[0][1] + ',' + ring[0][0]);                          // close the ring
    geoms.push('polygon:' + pts.join(',') +
               ';linewidth:2;linecolor:%23' + hex + ';fillcolor:%23' + hex + ';fillopacity:0.45');
  });
  if (!geoms.length) return '';
  // Keep delimiters literal (matches Geoapify's docs format and stays short — full
  // encoding tripled the length and the last polygon got truncated). Only '#'→%23
  // and the '|' separator→%7C are encoded, which both UrlFetchApp and Geoapify accept.
  // Portrait image + an area rect framing all of Israel (Eilat→Metula). Israel is
  // tall and narrow, so a fixed center+zoom on a landscape image cropped the ends.
  // area=rect is NW(lon,lat) then SE(lon,lat); it overrides center/zoom.
  var params = [
    'style=osm-bright', 'width=520', 'height=640', 'scaleFactor=2',
    'area=rect:34.0,33.5,36.0,29.3',
    'geometry=' + geoms.join('%7C'),
    'apiKey=' + encodeURIComponent(key),
  ];
  return 'https://maps.geoapify.com/v1/staticmap?' + params.join('&');
}

// Fetch the static region map as a PNG blob, or null if no key / non-200 / error.
function _nciStaticMapBlob(win) {
  var url = _nciStaticMapUrl(win);
  if (!url) return null;
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    return resp.getBlob().setName('congestion_map.png');
  } catch (e) { return null; }
}

// Dev preview: log the static-map URL for the latest window to open in a browser.
function _previewStaticMapUrl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(_nciStaticMapUrl(_latestWindowWithRows(ss) || { rows: [] }) || '(אין מפתח Geoapify מוגדר)');
}

function menuSetNCIMapApiKey() {
  var ui = SpreadsheetApp.getUi();
  var cur = _getConfig(NCI_MAP_KEY_CONFIG_KEY);
  var res = ui.prompt('מפתח Geoapify למפת החום במייל',
    'הזן מפתח Geoapify (חינמי, הרשמה ללא כרטיס אשראי ב-geoapify.com). ' +
    'השאר ריק כדי לבטל — המייל יישלח ללא התמונה (הקישור למפה האינטראקטיבית יישאר).\n\n' +
    'נוכחי: ' + (cur ? cur.slice(0, 6) + '…' : '(לא מוגדר)'),
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  _setConfig(NCI_MAP_KEY_CONFIG_KEY, res.getResponseText().trim());
  ui.alert(res.getResponseText().trim() ? 'מפתח Geoapify נשמר — המייל יכלול תמונת מפה.'
                                        : 'המפתח נוקה — המייל יישלח ללא תמונה.');
}

// Diagnose the email map image: fetches the Geoapify URL and reports the result.
function menuTestNCIMapImage() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!_getConfig(NCI_MAP_KEY_CONFIG_KEY)) {
    ui.alert('אין מפתח Geoapify מוגדר.\nהגדר אותו דרך "🔑 הגדר מפתח מפת חום למייל (Geoapify)".');
    return;
  }
  var win = _latestWindowWithRows(ss) || { rows: [] };
  function probe(url) {
    try {
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var code = resp.getResponseCode();
      return { code: code, body: code === 200 ? (resp.getBlob().getBytes().length + ' bytes') : resp.getContentText().slice(0, 300), len: url.length };
    } catch (e) { return { code: 'EXC', body: e.message, len: url.length }; }
  }
  var url = _nciStaticMapUrl(win);
  if (!url) { ui.alert('לא נוצר URL לתמונה.'); return; }
  var full = probe(url);
  if (full.code === 200) {
    ui.alert('✅ התמונה נוצרה תקין.\nHTTP 200 · ' + full.body + '\nאורך URL: ' + full.len);
    return;
  }
  // Isolate: probe each region on its own so we know exactly which polygon breaks.
  var lines = ['⚠️ הכל ביחד: HTTP ' + full.code + ' (URL ' + full.len + ')', full.body, '', 'לכל אזור בנפרד:'];
  Object.keys(REGION_GEO).forEach(function(name) {
    var r = probe(_nciStaticMapUrl(win, [name]));
    lines.push('• ' + name + ': HTTP ' + r.code + (r.code === 200 ? ' ✅ (' + r.body + ')' : ' ✗'));
  });
  ui.alert(lines.join('\n'));
}

// Most recent NCI window that actually has per-route rows (a live recompute;
// history rows are empty so they would render all-grey). Returns null if none.
function _latestWindowWithRows(ss) {
  try {
    var nci = _nciData(ss, _computeBaselines(ss));
    if (nci && nci.order.length) {
      for (var i = nci.order.length - 1; i >= 0; i--) {
        var w = nci.windows[nci.order[i]];
        if (w && w.rows && w.rows.length) return w;
      }
    }
  } catch (_) {}
  return null;
}

// Working web-app URL for menu/email links: the hardcoded anonymous /exec
// (falls back to getUrl()), plus a cache-busting token so a stale edge-cached
// response is never served. `token` makes the link stable per send when given.
function _nciMapUrl(token) {
  var base = NCI_MAP_WEBAPP_URL;
  if (!base) { try { base = ScriptApp.getService().getUrl() || ''; } catch (_) { base = ''; } }
  if (!base) return '';
  if (!token) {
    try { token = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jerusalem', 'yyyyMMddHHmm'); }
    catch (_) { token = 'v'; }
  }
  return base + (base.indexOf('?') >= 0 ? '&' : '?') + 't=' + encodeURIComponent(token);
}

// ─── Interactive heatmap (Web App) ───
// Served at the web-app URL after Deploy ▸ New deployment ▸ Web app.
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var win = _latestWindowWithRows(ss);
  var payload = {
    regions: REGION_GEO,
    devs: win ? _nciRegionDeviations(win) : {},
    meta: win ? { key: win.key, indexPct: win.indexPct, range: _winRangeLabel(win.key) } : null,
  };
  var tmpl = HtmlService.createTemplateFromFile('Map');
  tmpl.data = JSON.stringify(payload);
  return tmpl.evaluate()
    .setTitle('מפת חום — מדד תנועה ארצי')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function menuOpenCongestionMap() {
  var ui = SpreadsheetApp.getUi();
  var url = _nciMapUrl();
  if (!url) {
    ui.alert('המפה האינטראקטיבית עדיין לא פורסמה.\n\nבעורך הסקריפט: Deploy ▸ New deployment ▸ Web app, ואז נסה שוב.');
    return;
  }
  // A real <a> the user clicks — an auto window.open() on dialog-load is blocked
  // by the popup blocker (no user gesture). The link opens in a new tab on click.
  var safe = url.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
  var html = HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font-family:Arial,Helvetica,sans-serif;text-align:center;padding:18px 16px">' +
    '<a href="' + safe + '" target="_blank" rel="noopener" ' +
    'style="display:inline-block;background:#1F3864;color:#fff;text-decoration:none;' +
    'padding:12px 24px;border-radius:8px;font-weight:bold;font-size:15px">🗺️ פתח את מפת החום</a>' +
    '<p style="color:#64748b;font-size:12px;margin:14px 0 4px">נפתח בלשונית חדשה. אם לא — העתק את הכתובת:</p>' +
    '<input readonly onclick="this.select()" value="' + safe + '" ' +
    'style="width:100%;box-sizing:border-box;font-size:11px;padding:6px;border:1px solid #cbd5e1;border-radius:6px">' +
    '</div>')
    .setWidth(340).setHeight(170);
  ui.showModalDialog(html, 'מפת חום ארצית');
}

// ─── _nci_history (permanent log of every reading) ───
function _ensureNCIHistory(ss) {
  var s = ss.getSheetByName(NCI_HISTORY_SHEET);
  if (!s) {
    s = ss.insertSheet(NCI_HISTORY_SHEET);
    try { s.hideSheet(); } catch(e) {}
    s.getRange(1, 1, 1, NCI_HISTORY_COLS.length).setValues([NCI_HISTORY_COLS])
     .setBackground('#1F3864').setFontColor('#fff').setFontWeight('bold');
    s.setFrozenRows(1);
  }
  // Self-heal: pin the numeric columns to number formats. appendRow can inherit a
  // stray date format, and a date-formatted numeric cell makes getValues() return a
  // Date object — which surfaced as nonsense "1901-…" jam counts in the trend table.
  // Re-applying a number format reinterprets the stored serial as the original integer.
  var last = s.getLastRow();
  if (last >= 1) {
    s.getRange(1, 5, last, 1).setNumberFormat('0.0');  // index_pct
    s.getRange(1, 6, last, 2).setNumberFormat('0');     // n_jams, n_routes
  }
  return s;
}
function _appendNCIHistory(ss, win, date, daytype) {
  var s = _ensureNCIHistory(ss);
  var n = s.getLastRow() - 1;
  if (n > 0) {   // dedup: one row per (date, window)
    var keys = s.getRange(2, 2, n, 2).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (_ymd(keys[i][0]) === date && keys[i][1] === win.key) return false;
    }
  }
  s.appendRow([new Date(), date, win.key, daytype, win.indexPct, win.nJams, win.nRoutes,
               win.period || _periodFromDateHour(date, (win.hours && win.hours[0]) || 12)]);
  return true;
}
function _readNCIHistory(ss, limit) {
  var s = _ensureNCIHistory(ss);     // also repairs any date-formatted numeric cells
  if (s.getLastRow() < 2) return [];
  SpreadsheetApp.flush();            // commit the format repair before reading values
  var n = s.getLastRow() - 1;
  var start = Math.max(0, n - limit);
  return s.getRange(2 + start, 1, n - start, NCI_HISTORY_COLS.length).getValues().map(function(r) {
    return { run_ts: r[0], date: r[1], window: r[2], daytype: r[3],
             index_pct: Number(r[4]), n_jams: Number(r[5]), n_routes: Number(r[6]),
             period: r[7] || PERIOD_ROUTINE };
  });
}

// ─── Tab: פירוט לפי מסגרת זמן (window-level detail that feeds the index) ───
function _sheetTimeframes(ss, nci, baselines) {
  var ws = _newSheet(ss, NCI_TIMEFRAME_SHEET, '#2E75B6');
  var cols = ['אזור','מסלול','כיוון','מסגרת זמן','סוג יום','תקופה','מס\' פקקים','סטייה %','מקור השוואה','סטטוס'];
  _tabIntro(ws, cols.length, 'אותה השוואה כמו "פירוט לפי שעה", אך השעות מקובצות למסגרות (בוקר 06–10 / ערב 16–19). "מדד התנועה הארצי" = הממוצע המשוקלל לפי מס\' פקקים של עמודת "סטייה %" בכל מסגרת. חיובי = כבד מהשגרה, שלילי = קל. השוואה מול _baseline_archive (כל ההיסטוריה).', 1);
  _hdrRow(ws, cols, 2);
  var row = 3;
  var dtLabel = nci.daytype === 'weekend' ? 'סופ"ש' : 'חול';
  nci.order.forEach(function(wk) {
    var win = nci.windows[wk];
    var st = _nciStatus(win.indexPct);
    var lc = _nciLowConfidence(win);
    ws.getRange(row, 1, 1, cols.length).merge()
      .setValue('🚦 מסגרת ' + wk + ' (' + _winRangeLabel(wk) + ')  ·  מדד ארצי משוקלל: ' +
                _nciFmtPct(win.indexPct) + '  ·  ' + st.label + '  ·  ' + win.nJams + ' פקקים, ' + win.nRoutes + ' מסלולים' +
                (lc ? '  ·  ⚠️ ' + lc : ''))
      .setBackground(C_SEC_BG).setFontColor(C_SEC_FG).setFontWeight('bold').setHorizontalAlignment('right');
    row++;
    win.rows.forEach(function(rw) {
      var rst = _nciStatus(rw.devPct);
      _dataRow(ws, row, [rw.section, rw.route, rw.dir, wk, dtLabel, win.period, rw.jams,
                         rw.devPct != null ? _nciFmtPct(rw.devPct) : '—', rw.source, rst.label],
               row % 2 === 0);
      if (rw.devPct != null) _colorDeviation(ws, row, 8, rw.devPct);
      _colorStatus(ws, row, 10, rst.label);
      row++;
    });
  });
  if (!nci.order.length) {
    ws.getRange(3, 1, 1, cols.length).merge()
      .setValue('אין עדיין נתונים בחלונות הבוקר/ערב בתאריך ' + nci.date + '.')
      .setHorizontalAlignment('center').setFontColor('#475569');
  }
  _autoWidth(ws, cols.length);
  ws.setFrozenRows(2);
}

// ─── Tab: 🚦 מדד ארצי (the slide) ───
function _sheetNCI(ss, nci, history) {
  var ws = _newSheet(ss, NCI_SHEET, '#C00000');
  var nc = 6;
  var head = nci.headline;
  var hasVal = head && head.indexPct != null;
  var st = _nciStatus(hasVal ? head.indexPct : null);
  var dtLabel = nci.daytype === 'weekend' ? 'סופ"ש' : 'חול';

  ws.getRange(1, 1, 1, nc).merge().setValue('🚦 מדד התנועה הארצי — סטייה מהשגרה')
    .setFontSize(20).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F3864')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 46);

  var winLbl = head ? (head.key + ' ' + _winRangeLabel(head.key)) : '—';
  ws.getRange(2, 1, 1, nc).merge()
    .setValue('תאריך ' + nci.date + '   •   חלון ' + winLbl + '   •   ' + dtLabel +
              '   •   ' + (nci.period === PERIOD_EMERGENCY ? '🔴 חירום' : 'שגרה') +
              '   •   עודכן ' + Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM HH:mm'))
    .setFontSize(11).setFontColor('#475569').setBackground('#F1F5F9').setHorizontalAlignment('center');

  _tabIntro(ws, nc, 'מספר אחד למצב התנועה הארצי: הממוצע המשוקלל (לפי מס\' פקקים) של "סטייה %" מול ה-baseline ההיסטורי, למסגרת הזמן שזה עתה נסגרה. חיובי = כבד מהשגרה · שלילי = קל מהשגרה. מתעדכן אוטומטית בוקר וערב. פירוט מלא בלשונית "פירוט לפי מסגרת זמן".', 3);

  ws.getRange(4, 1, 3, nc).merge().setValue(hasVal ? _nciFmtPct(head.indexPct) : 'אין נתונים')
    .setFontSize(58).setFontWeight('bold').setFontColor(st.fg).setBackground(st.bg)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(4, 38); ws.setRowHeight(5, 38); ws.setRowHeight(6, 38);

  ws.getRange(7, 1, 1, nc).merge().setValue(st.head)
    .setFontSize(22).setFontWeight('bold').setFontColor(st.fg).setBackground(st.bg)
    .setHorizontalAlignment('center');
  ws.setRowHeight(7, 36);

  var sub = hasVal
    ? ('העומס הארצי ' + (head.indexPct >= 0 ? 'גבוה' : 'נמוך') + ' ב-' + Math.abs(_round1(head.indexPct)) +
       '% משגרת ' + head.key + '־' + (dtLabel === 'חול' ? 'חול' : 'סופ"ש') +
       '   ·   ' + head.nRoutes + ' מסלולים   ·   ' + head.nJams + ' פקקים')
    : 'אין עדיין נתונים לחלון הנוכחי בתאריך ' + nci.date + '.';
  ws.getRange(8, 1, 1, nc).merge().setValue(sub).setFontSize(11).setFontColor('#475569').setHorizontalAlignment('center');

  // Thin-sample caveat shown right under the headline number (row 9, otherwise blank).
  var lcHead = _nciLowConfidence(head);
  if (lcHead) {
    ws.getRange(9, 1, 1, nc).merge().setValue('⚠️ ' + lcHead)
      .setFontSize(11).setFontWeight('bold').setFontColor('#92400E').setBackground('#FEF3C7')
      .setHorizontalAlignment('center');
  }

  var row = 10;
  _secRow(ws, row, nc, 'מדד לפי חלון — ' + nci.date); row++;
  _hdrRow(ws, ['חלון','טווח','מדד %','מס\' פקקים','מסלולים','סטטוס'], row); row++;
  nci.order.forEach(function(wk) {
    var w = nci.windows[wk];
    _dataRow(ws, row, [wk, _winRangeLabel(wk), _nciFmtPct(w.indexPct), w.nJams, w.nRoutes, _nciStatus(w.indexPct).label], row % 2 === 0);
    _colorDeviation(ws, row, 3, w.indexPct);
    row++;
  });

  row++;
  _secRow(ws, row, nc, 'מגמת ' + history.length + ' קריאות אחרונות   ·   גבוה = כבד מהשגרה'); row++;
  _hdrRow(ws, ['תאריך','חלון','סוג יום','מדד %','פקקים','‎'], row); row++;
  var trendStart = row;
  history.forEach(function(h) {
    _dataRow(ws, row, [_ymd(h.date), h.window, (h.daytype === 'weekend' ? 'סופ"ש' : 'חול'), '', h.n_jams, ''], row % 2 === 0);
    ws.getRange(row, 4).setValue(_round1(h.index_pct))
      .setNumberFormat('+0.0"%";-0.0"%";0"%"').setHorizontalAlignment('center').setFontWeight('bold');
    _colorDeviation(ws, row, 4, h.index_pct);
    row++;
  });
  var trendEnd = row - 1;
  // Pin the jams column to an integer format so a value never renders as a date.
  if (history.length) ws.getRange(trendStart, 5, history.length, 1).setNumberFormat('0');

  if (history.length > 0) {
    ws.getRange(row, 1, 1, nc).merge()
      .setFormula('=SPARKLINE(D' + trendStart + ':D' + trendEnd +
        ',{"charttype","column";"color","#C00000";"negcolor","#1E40AF";"axis",true;"axiscolor","#94A3B8"})');
    ws.setRowHeight(row, 64);
    row++;
  }

  row++;
  ws.getRange(row, 1, 1, nc).merge()
    .setValue('מתודולוגיה: ממוצע משוקלל (לפי מס\' פקקים) של סטיית ההשהיה-לפקק מול _baseline_archive (כל ההיסטוריה), לפי מסגרות זמן. תזמון אוטומטי: ' +
              NCI_WINDOWS[0].runAtHour + ':00 ו-' + NCI_WINDOWS[1].runAtHour + ':00.')
    .setFontSize(9).setFontStyle('italic').setFontColor('#94A3B8').setHorizontalAlignment('right').setWrap(true);
  ws.setRowHeight(row, 30);

  for (var c = 1; c <= nc; c++) ws.setColumnWidth(c, 128);
  try { ss.setActiveSheet(ws); ss.moveActiveSheet(2); } catch(e) {}
}

// ─── Scheduled run (morning ~10:00 / evening ~20:00) ───
function _scheduledNCI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var step = 'התחלה';
  try {
    step = 'קריאת בייסליין מ-_baseline_archive';
    var baselines = _retry(function() { return _computeBaselines(ss); });
    step = 'קריאת raw_data';
    var nci = _retry(function() { return _nciData(ss, baselines); });
    if (!nci) return;
    step = 'רישום היסטוריה';
    var nowH = parseInt(Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'H'), 10);
    var key = nowH < 13 ? 'בוקר' : 'ערב';
    var win = nci.windows[key];
    if (win && win.indexPct != null) _appendNCIHistory(ss, win, nci.date, nci.daytype);
    var history = _readNCIHistory(ss, NCI_TREND_POINTS);
    step = 'בניית לשונית "פירוט לפי מסגרת זמן"';
    _sheetTimeframes(ss, nci, baselines);
    step = 'בניית לשונית "מדד ארצי"';
    _sheetNCI(ss, nci, history);
    step = 'שליחת מייל מדד ארצי';
    try { _sendNCIEmail(nci, win, history); } catch (mailErr) {
      // A mail failure must never abort the index run or its sheets.
      try { console.error('NCI email failed: ' + mailErr.message); } catch (_) {}
    }
  } catch (e) {
    // Don't hard-block on the mere existence of a job record (it persists while
    // idle/orphaned). Only surface it as a hint if the run actually failed.
    var hint = '';
    try { if (_getArchiveJob()) hint = '\n\n(ייתכן שעבודת ארכוב פעילה ברקע נועלת את המסמך — המתן דקה ונסה שוב, או סיים אותה דרך מתקדם ← "▶️ המשך ייצוא תקוע".)'; } catch (_) {}
    throw new Error('נכשל בשלב [' + step + ']: ' + e.message + hint);
  }
}
function _hasNCITrigger() {
  return ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === NCI_TRIGGER_FN; });
}
function installNCITriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === NCI_TRIGGER_FN) ScriptApp.deleteTrigger(t);
  });
  NCI_WINDOWS.forEach(function(w) {
    ScriptApp.newTrigger(NCI_TRIGGER_FN).timeBased().everyDays(1).atHour(w.runAtHour).create();
  });
}
function menuInstallNCITriggers() {
  installNCITriggers();
  _setConfig('nci_schedule', 'on');
  SpreadsheetApp.getUi().alert('תזמון "מדד ארצי" הותקן: בוקר ' + NCI_WINDOWS[0].runAtHour +
    ':00, ערב ' + NCI_WINDOWS[1].runAtHour + ':00 (כל יום).');
}
function menuRunNCINow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    _scheduledNCI();
    var s = ss.getSheetByName(NCI_SHEET);
    if (s) s.activate();
  } catch (e) {
    SpreadsheetApp.getUi().alert('חישוב המדד הארצי נכשל:\n' + e.message +
      '\n\nאם זו שגיאת timeout של שירות הגיליונות — נפוץ במסמכים גדולים או כשפעולה אחרת רצה במקביל. המתן רגע ונסה שוב.');
  }
}

// ═══ NCI EMAIL DIGEST ═════════════════════════════════════════════
// Sent at the end of every _scheduledNCI run (morning ~10:00, evening ~20:00),
// so recipients get one mail per rush window — twice a day. Default ON; the
// recipient defaults to the sheet owner unless overridden in _config.

function _nciEmailRecipients() {
  var raw = _getConfig(NCI_EMAIL_TO_KEY) || _currentUserEmail();
  return raw.split(/[,;]+/).map(function(s) { return s.trim(); }).filter(function(s) { return s; });
}

// One coloured "+12.3% · עמוס" pill, used in the headline and the trend.
function _nciEmailPill(pct) {
  var st = _nciStatus(pct);
  return '<span style="display:inline-block;padding:3px 10px;border-radius:12px;font-weight:bold;' +
         'background:' + st.bg + ';color:' + st.fg + '">' + _nciFmtPctLtr(pct) + ' · ' + st.label + '</span>';
}

function _nciEmailHtml(nci, win, history, mapUrl, hasMap) {
  var st = _nciStatus(win.indexPct);
  var dtLabel = nci.daytype === 'weekend' ? 'סופ"ש' : 'חול';
  var h = [];
  h.push('<div dir="rtl" style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:0 auto;color:#1f2937">');

  // Headline card
  h.push('<div style="background:' + st.bg + ';border-radius:12px;padding:18px 20px;text-align:center">');
  h.push('<div style="font-size:15px;color:' + st.fg + ';font-weight:bold">🚦 מדד תנועה ארצי · מסגרת ' +
         win.key + ' (' + _winRangeLabel(win.key) + ')</div>');
  h.push('<div style="font-size:40px;font-weight:bold;color:' + st.fg + ';margin:6px 0">' + _nciFmtPctLtr(win.indexPct) + '</div>');
  h.push('<div style="font-size:18px;color:' + st.fg + ';font-weight:bold">' + st.head + '</div>');
  var ctx = [nci.dayName, dtLabel, (win.period === PERIOD_EMERGENCY ? '🔴 חירום' : 'שגרה'),
             nci.date, win.nJams + ' פקקים', win.nRoutes + ' מסלולים']
              .filter(function(x) { return x; }).join(' · ');
  h.push('<div style="font-size:13px;color:#475569;margin-top:6px">' + ctx + '</div>');
  var lc = _nciLowConfidence(win);
  if (lc) {
    h.push('<div style="font-size:12px;color:#92400E;background:#FEF3C7;border-radius:8px;' +
           'padding:6px 10px;margin-top:8px;display:inline-block">⚠️ ' + lc + '</div>');
  }
  h.push('</div>');

  // Regional heatmap image (Geoapify PNG, attached inline as cid:congestionmap)
  if (hasMap) {
    h.push('<div style="text-align:center;margin:16px 0 6px">');
    h.push('<img src="cid:congestionmap" alt="מפת חום אזורית" ' +
           'style="width:100%;max-width:420px;border-radius:10px;border:1px solid #e2e8f0">');
    h.push('</div>');
  }

  // Link to the free interactive heatmap (Leaflet web app)
  if (mapUrl) {
    h.push('<div style="text-align:center;margin:' + (hasMap ? '6px' : '16px') + ' 0 18px">' +
           '<a href="' + mapUrl + '" style="display:inline-block;background:#1F3864;color:#fff;' +
           'text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:bold">' +
           '🗺️ פתח מפה אינטראקטיבית</a></div>');
  }

  // Worst routes table (rows already sorted descending by deviation)
  var worst = (win.rows || []).filter(function(r) { return r.devPct != null; }).slice(0, 5);
  if (worst.length) {
    h.push('<h3 style="font-size:15px;margin:20px 0 8px">הכבדים ביותר במסגרת זו</h3>');
    h.push('<table style="width:100%;border-collapse:collapse;font-size:13px">');
    h.push('<tr style="background:#1F3864;color:#fff;font-weight:bold">' +
           '<td style="padding:6px 8px;text-align:right">מסלול</td>' +
           '<td style="padding:6px 8px;text-align:right">כיוון</td>' +
           '<td style="padding:6px 8px;text-align:center">פקקים</td>' +
           '<td style="padding:6px 8px;text-align:center">סטייה</td></tr>');
    worst.forEach(function(r, i) {
      var rst = _nciStatus(r.devPct);
      h.push('<tr style="background:' + (i % 2 ? '#f8fafc' : '#fff') + '">' +
             '<td style="padding:6px 8px;text-align:right">' + r.route + '</td>' +
             '<td style="padding:6px 8px;text-align:right;color:#64748b">' + r.dir + '</td>' +
             '<td style="padding:6px 8px;text-align:center">' + r.jams + '</td>' +
             '<td style="padding:6px 8px;text-align:center;font-weight:bold;color:' + rst.fg + '">' +
             _nciFmtPctLtr(r.devPct) + '</td></tr>');
    });
    h.push('</table>');
  }

  // Recent trend for THIS window
  var trend = (history || []).filter(function(r) { return r.window === win.key && r.index_pct != null; })
                             .slice(-NCI_EMAIL_TREND);
  if (trend.length > 1) {
    h.push('<h3 style="font-size:15px;margin:20px 0 8px">מגמה אחרונה (' + win.key + ')</h3>');
    h.push('<table style="border-collapse:collapse;font-size:13px"><tr>');
    trend.forEach(function(r) {
      h.push('<td style="padding:6px 10px;text-align:center;border-bottom:2px solid #e2e8f0">' +
             '<div style="color:#64748b;font-size:11px">' + _ymd(r.date).slice(5) + '</div>' +
             '<div style="margin-top:3px">' + _nciEmailPill(r.index_pct) + '</div></td>');
    });
    h.push('</tr></table>');
  }

  h.push('<p style="font-size:11px;color:#94a3b8;margin-top:24px;border-top:1px solid #e2e8f0;padding-top:10px">' +
         'נשלח אוטומטית ע"י "מדד תנועה ארצי". להפסקה/שינוי נמענים: תפריט 🚦 Waze ← מתקדם, או הסר את הטריגרים.</p>');
  h.push('</div>');
  return h.join('');
}

function _sendNCIEmail(nci, win, history) {
  if (_getConfig(NCI_EMAIL_CONFIG_KEY) === 'off') return false;
  if (!win || win.indexPct == null) return false;
  var to = _nciEmailRecipients();
  if (!to.length) return false;
  // Pull a few more readings than the slide keeps, so the same-window trend has depth.
  var hist = history;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    hist = _readNCIHistory(ss, NCI_EMAIL_TREND * 2 + 6);
  } catch (_) {}
  var st = _nciStatus(win.indexPct);
  var subject = '🚦 מדד ארצי — ' + win.key + ' ' + nci.date + ': ⁦' + _nciFmtPct(win.indexPct) + '⁩ · ' + st.label;
  var mapUrl = _nciMapUrl(nci.date + '-' + win.key);   // stable per send, cache-busted
  var mapBlob = _nciStaticMapBlob(win);   // null when no Geoapify key / fetch failed → email sent without the image
  var msg = {
    to: to.join(','),
    subject: subject,
    htmlBody: _nciEmailHtml(nci, win, hist, mapUrl, !!mapBlob),
    name: 'מדד תנועה ארצי',
  };
  if (mapBlob) msg.inlineImages = { congestionmap: mapBlob };   // key must match cid:congestionmap
  MailApp.sendEmail(msg);
  return true;
}

// ─── Menu actions for the email digest ───
function menuToggleNCIEmail() {
  var ui = SpreadsheetApp.getUi();
  var on = _getConfig(NCI_EMAIL_CONFIG_KEY) !== 'off';   // default on
  _setConfig(NCI_EMAIL_CONFIG_KEY, on ? 'off' : 'on');
  var to = _nciEmailRecipients();
  ui.alert(on ? 'מייל "מדד ארצי" כובה.' :
    'מייל "מדד ארצי" הופעל — יישלח אחרי כל חישוב (בוקר+ערב) אל:\n' +
    (to.length ? to.join(', ') : '(בעל הגיליון)'));
}
function menuSetNCIEmailRecipients() {
  var ui = SpreadsheetApp.getUi();
  var cur = _getConfig(NCI_EMAIL_TO_KEY) || _currentUserEmail();
  var res = ui.prompt('נמעני מייל "מדד ארצי"',
    'הזן כתובת אחת או כמה, מופרדות בפסיק. השאר ריק כדי לחזור לבעל הגיליון.\n\nנוכחי: ' + cur,
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  _setConfig(NCI_EMAIL_TO_KEY, res.getResponseText().trim());
  _setConfig(NCI_EMAIL_CONFIG_KEY, 'on');   // setting recipients implies wanting the mail
  var to = _nciEmailRecipients();
  ui.alert('נמענים עודכנו:\n' + (to.length ? to.join(', ') : '(בעל הגיליון)'));
}
// Build a minimal {nci, win} from the most recent _nci_history reading, so the
// test mail works even when today's raw_data has no fresh rush-window rows.
// History lacks per-route detail, so win.rows is empty (the worst-routes table
// is simply omitted) — everything else mirrors a real reading.
function _nciLatestFromHistory(ss) {
  var hist = _readNCIHistory(ss, NCI_EMAIL_TREND * 2 + 6);
  for (var i = hist.length - 1; i >= 0; i--) {
    if (hist[i].index_pct != null && hist[i].index_pct !== '') {
      var r = hist[i];
      var win = { key: r.window, indexPct: _round1(+r.index_pct), nJams: r.n_jams, nRoutes: r.n_routes, rows: [] };
      var nci = { date: _ymd(r.date), dayName: '', daytype: r.daytype, order: [r.window], windows: {} };
      nci.windows[r.window] = win;
      return { nci: nci, win: win, hist: hist };
    }
  }
  return null;
}

function menuSendNCIEmailTest() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var nci = null, win = null, hist = null, fromHistory = false;
    // Prefer a live recompute; fall back to the latest stored reading.
    try {
      var baselines = _computeBaselines(ss);
      nci = _nciData(ss, baselines);
    } catch (_) {}
    if (nci && nci.order.length) {
      for (var i = nci.order.length - 1; i >= 0; i--) {
        if (nci.windows[nci.order[i]].indexPct != null) { win = nci.windows[nci.order[i]]; break; }
      }
    }
    if (!win) {
      var fb = _nciLatestFromHistory(ss);
      if (fb) { nci = fb.nci; win = fb.win; hist = fb.hist; fromHistory = true; }
    }
    if (!win) {
      ui.alert('אין עדיין מדד לשליחה — לא בנתונים החיים ולא בהיסטוריה.\n\nהרץ קודם "🚦 חשב מדד ארצי עכשיו". אם גם זה נכשל, ייתכן שאין נתוני שעות שיא בתאריך האחרון או שחסר בייסליין ב-_baseline_archive.');
      return;
    }
    var prev = _getConfig(NCI_EMAIL_CONFIG_KEY);
    _setConfig(NCI_EMAIL_CONFIG_KEY, 'on');   // force-send even if currently off
    var sent = _sendNCIEmail(nci, win, hist || _readNCIHistory(ss, NCI_EMAIL_TREND * 2 + 6));
    if (prev === 'off') _setConfig(NCI_EMAIL_CONFIG_KEY, 'off');
    ui.alert(!sent ? 'לא נשלח (אין נמענים).' :
      'נשלח מייל בדיקה אל:\n' + _nciEmailRecipients().join(', ') +
      (fromHistory ? '\n\n(נשלח לפי הקריאה האחרונה מההיסטוריה, כי אין מדד חי כרגע.)' : ''));
  } catch (e) {
    ui.alert('שליחת מייל הבדיקה נכשלה:\n' + e.message);
  }
}

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
    { type: ROW, a: 'daytype', b: 'סוג יום: weekday (חול) או weekend (סופ"ש) — נגזר מיום בשבוע.' },
    { type: ROW, a: 'period (תקופה)', b: 'חירום או שגרה — לפי חלוני חירום מוגדרים בקוד (EMERGENCY_WINDOWS). כרגע: 7.6.2026 22:00 → 9.6.2026 06:00. שעות חירום מוחרגות מחישוב ה-baseline כדי שהסטייה תימדד מול תנועת שגרה.' },
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
  var ncW = 10;
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

  var cols = ['תאריך','סוג יום','תקופה','מסלול','כיוון','שעה','מס\' פקקים',
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
      _periodFromDateHour(dateStr, hour),
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
