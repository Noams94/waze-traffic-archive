# Waze Traffic Archive & Analyzer

Google Sheets + Apps Script tool for archiving and analyzing raw Waze API JSON snapshots — Hebrew RTL, no backend, shareable via a Google Sheets link.

## What it does

- **Ingests** raw Waze API JSON (manual upload or scheduled URL fetch)
- **Accumulates** a 30-day rolling raw archive (`raw_data`) + a **permanent aggregated archive** (`_baseline_archive`) that survives pruning
- **Filters** the archive by date range, hour-of-day, and day-of-week
- **Computes baselines** from the permanent archive (per route × direction × **hour** × **weekday/weekend**) with neighbor-window fallback (±1h, ±2h). When no historical data exists for a time-point, reports "אין מספיק נתונים" rather than comparing against an irrelevant baseline.
- **Renders** 8 sheets driven by an interactive sidebar:
  - 🎯 לוח מחוונים (Dashboard) — KPI strip, top 10 worst routes, best 5, section breakdown, analysis quality indicator
  - סיכום מסלולים — per-route summary with deviation vs historical average
  - פירוט לפי שעה — hour-by-hour breakdown with same-time comparisons (weekday vs weekend separated)
  - השוואת כיוונים — direction comparison
  - חריגות — anomaly detection (1.5σ + low speed + high level)
  - פירוט פקקים — full jam log
  - אגרגציה לאורך זמן — long-term trends from the permanent archive
  - מקרא ומתודולוגיה — legend + full methodology documentation
- **Auto-fetches** from a configured URL on a schedule (30 min / hourly / 4 hr / daily) — schedule config persists in the spreadsheet itself, with a flow for re-activating the trigger when switching Google accounts

## Files

| File | Purpose |
|------|---------|
| `Code.gs` | Apps Script source — paste into the Apps Script editor as the main file |
| `Sidebar.html` | Apps Script HTML — paste into a new HTML file named `Sidebar` |
| `waze_analyzer.py` | Original Python CLI version (still works standalone — `python waze_analyzer.py input.json output.xlsx`) |
| `waze_dashboard.jsx` | Earlier React/Recharts dashboard (reference; superseded by the Sheets dashboard) |
| `waze_routes_v2.xlsx` | Sample Excel output from the Python version |

## Setup

1. Create a new Google Sheet
2. Open `Extensions → Apps Script`
3. Replace the default `Code.gs` content with the contents of `Code.gs` from this repo
4. Click `+` → `HTML` → name it `Sidebar` → paste in the contents of `Sidebar.html`
5. Save (`Cmd+S`) and reload the Sheet
6. The menu `🚦 Waze` will appear → `פתח סרגל צד...`

## Usage

### Manual upload
- Sidebar → 📤 העלאה tab → drag a Waze JSON or TXT file → "הוסף לארכיון"

### Scheduled auto-fetch
- Sidebar → ⚡ אוטומציה tab → paste URL (and optional `Authorization` headers as JSON) → choose interval → "💾 שמור תזמון"
- The trigger is bound to the Google account that created it. When opening the Sheet from a different account, click "🔁 הפעל תזמון בחשבון הזה"

### Filter & rebuild analysis
- Sidebar → 🔍 סינון tab → pick date range / hours / days → "החל סינון ובנה גיליונות"
- The 8 analysis sheets rebuild from the filtered archive

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

### Rebuilt only on upload (not on filter apply)

| Sheet | Purpose |
|-------|---------|
| אגרגציה לאורך זמן | View of `_baseline_archive` — all-time trends, sorted newest first |

### Persistent — accumulate over time, untouched by filters

| Sheet | Persistence | Purpose |
|-------|-------------|---------|
| `raw_data` | append-only, auto-pruned > 30 days | one row per jam per snapshot. Extra cols: `route_name`, `dir_ix`, `archived` (flag) |
| `_baseline_archive` | hidden, **never pruned** | aggregated counters per `(route, dir, date, hour)` — `n, sum_delay_s, sum_speed, sum_level`. Source of truth for all historical baselines. |
| `מקור` | append-only, auto-pruned > 30 days | log of every upload (`startTime`, jam count) |
| `ארכיון נפרד` | append-only | log of CSV exports to Drive (filename, date range, link) |
| `_config` | hidden | URL, headers, interval, trigger owner — persists across users/devices |
| `_fetch_log` | hidden | log of every scheduled fetch (timestamp, user, status, error) |
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
| מקרא ומתודולוגיה | This methodology document, available in-sheet. |

### How baselines are computed

The baseline for any comparison is a historical average of `delay_s` per jam, keyed by:

```
route × direction × hour-of-day × daytype
```

Where `daytype ∈ { weekday (א'–ה'), weekend (ו'–ש') }`. The aggregation happens in the permanent `_baseline_archive` sheet — every new jam upserts into the matching `(route, dir, date, hour)` cell, accumulating `n` and `sum_delay_s` forever. Daytype is derived from date at read time.

A cell is considered "usable" only when **n ≥ 3** historical samples have accumulated for it.

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
- **On every "החל סינון"**: the 7 analysis tabs are wiped and redrawn from the filtered `raw_data`. Baselines are read from the full `_baseline_archive` (not the filter).
- **Pruning**: `raw_data` rows older than 30 days are exported to Drive as CSV (resumable, chunked) and then deleted. `_baseline_archive` is never pruned.
- **Migration**: existing data from before the archive existed can be backfilled via **🚦 Waze → 🔄 העבר נתונים קיימים לארכיון אגרגטיבי**.

### Glossary

| Term | Meaning |
|------|---------|
| daytype | `weekday` (א'–ה') or `weekend` (ו'–ש') |
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
