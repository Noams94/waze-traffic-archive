# Waze Traffic Archive & Analyzer

Google Sheets + Apps Script tool for archiving and analyzing raw Waze API JSON snapshots — Hebrew RTL, no backend, shareable via a Google Sheets link.

## What it does

- **Ingests** raw Waze API JSON (manual upload or scheduled URL fetch)
- **Accumulates** a 30-day rolling raw archive (`raw_data`) + a **permanent aggregated archive** (`_baseline_archive`) that survives pruning
- **Filters** the archive by date range, hour-of-day, and day-of-week
- **Computes baselines** from the permanent archive (per route × direction × **hour** × **weekday/weekend**) with neighbor-window fallback (±1h, ±2h). When no historical data exists for a time-point, reports "אין מספיק נתונים" rather than comparing against an irrelevant baseline.
- **Renders** 10 sheets driven by an interactive sidebar:
  - 🎯 לוח מחוונים (Dashboard) — KPI strip, top 10 worst routes, best 5, section breakdown, analysis quality indicator
  - סיכום מסלולים — per-route summary with deviation vs historical average
  - פירוט לפי שעה — hour-by-hour breakdown with same-time comparisons (weekday vs weekend separated)
  - השוואת כיוונים — direction comparison
  - חריגות — anomaly detection (1.5σ + low speed + high level)
  - פירוט פקקים — full jam log
  - אגרגציה לאורך זמן — long-term trends from the permanent archive
  - 🚦 מדד ארצי — National Traffic Index (NCI) one-pager: one weighted number per rush window + recent-readings trend
  - פירוט לפי מסגרת זמן — the per-route detail behind the index (morning/evening windows)
  - מקרא ומתודולוגיה — legend + full methodology documentation
- **Computes a National Traffic Index (NCI)** twice a day — morning window 06:00–10:00 (runs ~10:00) and evening 16:00–19:00 (runs ~20:00): jam-weighted average of per-route deviation % vs the permanent baseline, logged forever to `_nci_history`, with a low-confidence caveat on thin windows (< 80 jams or < 12 routes)
- **Emails a digest** after each NCI run — headline index, 5 worst routes, recent trend, an inline regional heatmap image (Geoapify Static Maps, optional free key) and a link to an interactive Leaflet heatmap served as an Apps Script web app (`Map.html`). Recipients and on/off are menu-configurable
- **Classifies every jam** as חירום (emergency) / שגרה (routine) via explicit date-time windows (`EMERGENCY_WINDOWS` in `Code.gs`); emergency hours are **excluded from baselines**, so deviation is always measured against routine traffic
- **Auto-fetches** from a configured URL on a schedule (30 min / hourly / 4 hr / daily) — schedule config persists in the spreadsheet itself, with a flow for re-activating the trigger when switching Google accounts
- **Auto-archives daily** at 03:00 (incremental CSV export to Drive + pruning) via chunked resumable background jobs, with watchdog triggers and self-healing retries when the sheet approaches Google's 10M-cell cap

## Files

| File | Purpose |
|------|---------|
| `Code.gs` | Apps Script source — paste into the Apps Script editor as the main file |
| `Sidebar.html` | Apps Script HTML — paste into a new HTML file named `Sidebar` |
| `Map.html` | Apps Script HTML — the interactive regional heatmap (paste into a new HTML file named `Map`; served via a Web App deployment) |
| `waze_analyzer.py` | Original Python CLI version (still works standalone — `python waze_analyzer.py input.json output.xlsx`) |
| `waze_dashboard.jsx` | Earlier React/Recharts dashboard (reference; superseded by the Sheets dashboard) |
| `waze_routes_v2.xlsx` | Sample Excel output from the Python version |

## Setup

1. Create a new Google Sheet
2. Open `Extensions → Apps Script`
3. Replace the default `Code.gs` content with the contents of `Code.gs` from this repo
4. Click `+` → `HTML` → name it `Sidebar` → paste in the contents of `Sidebar.html`
5. Click `+` → `HTML` → name it `Map` → paste in the contents of `Map.html`
6. Save (`Cmd+S`) and reload the Sheet
7. The menu `🚦 Waze` will appear → `פתח סרגל צד...`
8. *(Optional, for the interactive heatmap link)* In the script editor: `Deploy ▸ New deployment ▸ Web app` (access: anyone), then set the resulting `/exec` URL as `NCI_MAP_WEBAPP_URL` in `Code.gs`
9. *(Optional, for the heatmap image in the email)* Menu `🚦 Waze → 🔑 הגדר מפתח מפת חום למייל (Geoapify)` — free key, no credit card

## Usage

### Manual upload
- Sidebar → 📤 העלאה tab → drag a Waze JSON or TXT file → "הוסף לארכיון"

### Scheduled auto-fetch
- Sidebar → ⚡ אוטומציה tab → paste URL (and optional `Authorization` headers as JSON) → choose interval → "💾 שמור תזמון"
- The trigger is bound to the Google account that created it. When opening the Sheet from a different account, click "🔁 הפעל תזמון בחשבון הזה"

### Filter & rebuild analysis
- Sidebar → 🔍 סינון tab → pick date range / hours / days → "החל סינון ובנה גיליונות"
- The 9 analysis sheets rebuild from the filtered archive

### National Traffic Index (NCI)
- Menu `🚦 Waze → ⏰ התקן תזמון מדד ארצי` installs the twice-daily triggers (also auto-installed on open); `🚦 חשב מדד ארצי עכשיו` runs it on demand
- Email digest: on by default after every run — toggle with `📧 הפעל/כבה מייל מדד ארצי`, set recipients with `✉️ הגדר נמעני מייל מדד...`, verify with `📤 שלח מייל מדד עכשיו (בדיקה)` and `🧪 בדוק תמונת מפה למייל`
- `🗺️ פתח מפת חום אינטראקטיבית` opens the Leaflet web-app map

### Reset
- Menu `🚦 Waze → מחק את כל הנתונים` (with confirmation)

## Data model

Sheets fall into three categories: **rebuilt on filter apply**, **rebuilt on upload**, and **persistent** (accumulating, never cleared by user actions).

### Rebuilt on every "החל סינון" (cleared and redrawn from filtered `raw_data`)

| Sheet | Purpose |
|-------|---------|
| 🎯 לוח מחוונים | One-page dashboard — KPIs, top 10 worst, best 5, section breakdown |
| סיכום מסלולים | Per-route summary table |
| פירוט לפי שעה | Hour-by-hour breakdown (weekday vs weekend separated) |
| השוואת כיוונים | Direction-vs-direction for two-way routes |
| חריגות | Local outlier detection (1.5σ / speed / level) |
| פירוט פקקים | Full jam log |
| מקרא ומתודולוגיה | Legend + methodology guide (first tab) |
| 🚦 מדד ארצי | NCI one-pager — also refreshed by the twice-daily NCI triggers; filter-independent (always latest window vs full history) |
| פירוט לפי מסגרת זמן | Window-level per-route detail feeding the index — same refresh behavior as the NCI tab |

### Rebuilt only on upload (not on filter apply)

| Sheet | Purpose |
|-------|---------|
| אגרגציה לאורך זמן | View of `_baseline_archive` — all-time trends, sorted newest first |

### Persistent — accumulate over time, untouched by filters

| Sheet | Persistence | Purpose |
|-------|-------------|---------|
| `raw_data` | append-only, auto-pruned > 30 days (plus a row-cap safety prune) | one row per jam per snapshot. Extra cols: `route_name`, `dir_ix`, `archived` (flag), `period` (חירום/שגרה) |
| `_baseline_archive` | hidden, **never pruned** | aggregated counters per `(route, dir, date, hour)` — `n, sum_delay_s, sum_speed, sum_level`. Source of truth for all historical baselines. |
| `_nci_history` | hidden, **never pruned** | one row per NCI reading — `run_ts, date, window, daytype, index_pct, n_jams, n_routes, period`. Feeds the trend on the slide and in the email. |
| `מקור` | append-only, auto-pruned > 30 days | log of every upload (`startTime`, jam count) |
| `ארכיון נפרד` | append-only | log of CSV exports to Drive (filename, date range, link) |
| `_config` | hidden | URL, headers, interval, trigger owner, NCI email/map settings — persists across users/devices |
| `_fetch_log` | hidden | log of every scheduled fetch (timestamp, user, status, error) |
| `_archive_slice_log` | hidden | diagnostics for the chunked exporter (chunks, timings, exit reason per slice) |
| `_filter` | hidden, overwritten each filter | last applied filter (JSON blob) |

**Key principle:** filtering only affects which `raw_data` rows feed the analysis tabs. Baselines for deviation calculations always come from the full `_baseline_archive` regardless of the filter — so "what I'm seeing in my chosen range" is always compared against "everything I know historically."

### Migrating from a pre-archive snapshot

If you're upgrading an existing sheet that has `raw_data` but no `_baseline_archive`: run **🚦 Waze → 🔄 העבר נתונים קיימים לארכיון אגרגטיבי** from the menu. It walks `raw_data` once, resolves each row's route+direction, populates `_baseline_archive`, and flags the rows so they aren't double-counted later.

## Methodology

This section mirrors the in-sheet "מקרא ומתודולוגיה" tab. It explains how the analysis numbers are computed and what they mean.

### Tab guide — what each sheet shows

| Sheet | What it shows and how to read it |
|-------|----------------------------------|
| 🎯 לוח מחוונים | One-page overview. KPI strip counts routes by status (תקין / מתון / עמוס / חריג מאוד / טוב מהרגיל / אין נתונים). Top 10 worst and best 5 routes ranked by deviation %. Section breakdown by region (דרום / מרכז / צפון). "איכות הניתוח" line reports how many routes are using each fallback tier. Good for "what's happening right now". |
| סיכום מסלולים | One row per route, both directions side-by-side. "סטייה כ1/כ2" columns compare against the **all-hours** historical average for the direction — coarse view. For per-hour resolution, go to the next tab. |
| פירוט לפי שעה | The core analysis. One row per (route × direction × hour × daytype). "סטייה %" tells you if the situation is worse (positive) or better (negative) than history at that time. "מקור השוואה" tells you how precise the comparison is: "שעה זו" = exact-hour match; "±1 שעות" / "±2 שעות" = expanded window because samples were sparse; "—" = no historical data. `n` is the historical sample count behind the baseline. |
| השוואת כיוונים | Dual-direction routes only. Side-by-side comparison. The busier direction is highlighted in red; "יחס" shows how many times one direction is heavier than the other. |
| חריגות | Individual jams that stand out **within the current filter**. Criterion is local statistics (1.5σ above mean for the route×direction, or speed < 5 km/h, or level ≥ 4) — not historical. Sorted by severity (קריטי → גבוה → בינוני) then by delay size. |
| פירוט פקקים | Full raw log of every jam in the current filter — one row per jam, sorted route → direction → time. For deep-dive investigation. |
| אגרגציה לאורך זמן | Direct view of `_baseline_archive` — full history, not just 30 days. One row per (date × route × direction × hour). Sorted newest first. Capped at 5,000 most recent rows in the display. Use for long-term trend analysis. |
| 🚦 מדד ארצי | The NCI slide: one big weighted deviation number for the latest closed rush window, per-window table, recent-readings trend with a sparkline, and a thin-sample caveat when the window is sparse. |
| פירוט לפי מסגרת זמן | Same comparison as פירוט לפי שעה, but hours are bucketed into the rush windows (בוקר 06–10 / ערב 16–19). One row per route × direction per window, with a summary row showing the window's weighted index. Filter-independent. |
| מקרא ומתודולוגיה | This methodology document, available in-sheet. |

### How baselines are computed

The baseline for any comparison is a historical average of `delay_s` per jam, keyed by:

```
route × direction × hour-of-day × daytype
```

Where `daytype ∈ { weekday (א'–ה'), weekend (ו'–ש') }`. The aggregation happens in the permanent `_baseline_archive` sheet — every new jam upserts into the matching `(route, dir, date, hour)` cell, accumulating `n` and `sum_delay_s` forever. Daytype is derived from date at read time.

A cell is considered "usable" only when **n ≥ 3** historical samples have accumulated for it.

**Emergency exclusion:** hours falling inside a declared emergency window (`EMERGENCY_WINDOWS` in `Code.gs`, half-open `[from, to)` local time) are skipped when baselines are built, so deviation is always measured against routine traffic. Every jam and every NCI reading also carries a `period` label (חירום/שגרה); existing rows can be re-labeled after declaring a new window via `מתקדם → 🔴 סווג חירום/שגרה לנתונים קיימים`.

### Fallback policy (neighbor-window expansion)

For a query `(route, dir, hour=H, daytype=D)`:

1. Try exact cell `(H, D)`. If `n ≥ 3` → **"שעה זו"** (most precise).
2. Expand to ±1 hour: `(H−1, H, H+1)` within the same daytype. If combined `n ≥ 3` → **"±1 שעות"** (3-hour window).
3. Expand to ±2 hours: `(H−2..H+2)`. If combined `n ≥ 3` → **"±2 שעות"** (5-hour window).
4. Otherwise → **"אין מספיק נתונים"** — no comparison, no status. The cell is shown but flagged as uncomparable.

**Important constraint:** weekday and weekend baselines never share samples, even during window expansion. Hours wrap modulo 24, so hour 23 shares with 22, 23, 0, 1.

The previous fallback to "all-day average" and "free-flow theoretical" was removed because those produced systematic false positives in rush hour and false negatives at night.

### Deviation formula

```
deviation% = (current_avg − historical_avg) / historical_avg × 100
```

Where both averages are `delay_s` per jam in seconds. Display converts to minutes; the calculation runs in seconds. Positive = worse than usual at that time; negative = better than usual.

### Status thresholds

| Status | Range |
|--------|-------|
| 🟢 תקין | `|deviation|` ≤ 10% |
| 🟡 מתון | +10% to +25% |
| 🟠 עמוס | +25% to +50% |
| 🔴 חריג מאוד | > +50% |
| 🔵 טוב מהרגיל | < −10% |
| ⚪ אין מספיק נתונים | No baseline available — cannot assign a status |

### Anomaly detection (separate from baselines)

The "חריגות" tab uses a different methodology — **local statistics within the current filter**, not historical comparison. A jam is flagged if any of:

- `delay_s > mean + 1.5σ` for that route × direction in the current filter
- `speed < 5` km/h
- `level ≥ 4` (Waze's own "heavy" / "standstill" classification)

Severity:
- **קריטי** — `speed < 3` or deviation from local mean > 200%
- **גבוה** — `level ≥ 4` or deviation > 100%
- **בינוני** — otherwise

The baseline tabs answer *"is this worse than usual?"*. The anomaly tab answers *"which jams stand out from the rest of what I'm seeing right now?"*. Complementary, not redundant.

### National Traffic Index (NCI)

One national number per rush window: the **jam-weighted average** of the per-route/per-hour deviation % (same comparison as פירוט לפי שעה) across all routes, bucketed into the window. Positive = heavier than routine, negative = lighter.

- **Windows:** בוקר 06:00–10:00 (computed ~10:00) · ערב 16:00–19:00 (computed ~20:00), daily triggers.
- **History:** every reading is appended to the permanent `_nci_history` sheet (deduped per date × window) and drives the trend on the slide and in the email.
- **Low confidence:** a window with < 80 jams or < 12 contributing routes gets a "מדגם קטן" caveat next to the number (common on weekends) — shown, never hidden.
- **Email digest:** sent after each run (default on; recipients default to the sheet owner) — headline pill, 5 worst routes, same-window trend, inline regional heatmap image (Geoapify) and a link to the interactive map.
- **Regional heatmap:** the three regions (צפון/מרכז/דרום, drawn from Israel's official administrative districts grouped into three) are colored by the region's jam-weighted deviation for the window. Interactive version = `Map.html` web app; static version = Geoapify Static Maps PNG embedded in the email.

### Color scale for deviation columns

| Color | Meaning |
|-------|---------|
| 🟢 Green | −10% to +10% (normal) |
| 🟡 Yellow | +10% to +25% |
| 🟠 Orange | +25% to +50% |
| 🔴 Red | > +50% |
| 🔵 Blue | < −10% (better than usual) |

### Methodology limits

- **First few weeks**: most cells will fall back to ±1 or ±2 windows. As the archive matures, more cells will hit "שעה זו". The fallback source column is transparent about this.
- **No holiday/event handling**: a holiday day's data becomes a regular sample in the statistics. There's no calendar of exceptional days.
- **No recency weighting**: a sample from a year ago counts equally to one from this week.
- `free_flow_min` (theoretical jam-free travel time) is a manual estimate per route, kept as a reference column but **not used as a baseline**. It answers "how long would this take with no jams?" — not "is this worse than usual?".

### Refresh flow

- **On every upload** (manual or scheduled): the JSON is parsed, deduplicated against `מקור` by `startTime`, appended to `raw_data`, and immediately upserted into `_baseline_archive`. The "אגרגציה לאורך זמן" tab is rebuilt.
- **On every "החל סינון"**: the 9 analysis tabs are wiped and redrawn from the filtered `raw_data`. Baselines are read from the full `_baseline_archive` (not the filter).
- **Twice a day** (~10:00 / ~20:00): the NCI is computed, appended to `_nci_history`, both NCI tabs are rebuilt and the email digest is sent.
- **Daily at 03:00**: incremental CSV export to Drive + pruning (toggle via `🗓️ הפעל/כבה גיבוי יומי`).
- **Pruning**: `raw_data` rows older than 30 days are exported to Drive as CSV (resumable, chunked) and then deleted; a row-cap safety prune also fires as the sheet nears Google's 10M-cell limit. `_baseline_archive` and `_nci_history` are never pruned.
- **Migration**: existing data from before the archive existed can be backfilled via **🚦 Waze → 🔄 העבר נתונים קיימים לארכיון אגרגטיבי**.

### Glossary

| Term | Meaning |
|------|---------|
| daytype | `weekday` (א'–ה') or `weekend` (ו'–ש') |
| period | `חירום` (emergency) or `שגרה` (routine) — set by explicit date-time windows in `EMERGENCY_WINDOWS`; emergency hours are excluded from baselines |
| NCI | National Traffic Index — jam-weighted national deviation % for a rush window (בוקר/ערב) |
| jam | A single Waze jam event — one row in `raw_data` |
| snapshot | One full pull from the Waze API (dozens to hundreds of jams) |
| baseline | The historical average that current values are compared against |
| upsert | "Update or insert" — if the key exists, increment counters; otherwise create the row |
| prune | Auto-deletion of `raw_data` rows older than the retention window |
| level | Waze's congestion classification: 1=free flow, 2=light, 3=moderate, 4=heavy, 5=standstill |

## Deduplication

Each upload is matched against existing snapshots in the `מקור` sheet by `startTime`. Re-uploading the same JSON returns a "duplicate" status without modifying the archive — guarantees the longitudinal aggregation isn't double-counted.

## Why not a hosted web app?

Considered Vercel + Next.js + Python serverless first. Pivoted to Sheets because:
- Zero hosting cost
- Inherent sharing (anyone with the Sheet link)
- The output is naturally a spreadsheet
- Triggers run in Google's infrastructure — no separate scheduler needed

The Vercel-based plan is preserved at `~/.claude/plans/see-those-files-i-transient-brooks.md` for reference.

## Author

Noam Keshet · [noamkeshet.com](https://noamkeshet.com)
