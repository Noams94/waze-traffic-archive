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

| Sheet | Persistence | Purpose |
|-------|-------------|---------|
| `raw_data` | append-only, auto-pruned > 30 days | one row per jam per snapshot. Extra cols: `route_name`, `dir_ix`, `archived` (flag) |
| `_baseline_archive` | hidden, **permanent** | aggregated counters per `(route, dir, date, hour)` — `n, sum_delay_s, sum_speed, sum_level`. Source of truth for all historical baselines. |
| `מקור` | append-only, auto-pruned | log of every upload (`startTime`, jam count) |
| `_config` | hidden | URL, headers, interval, trigger owner — persists across users/devices |
| `_fetch_log` | hidden | log of every scheduled fetch (timestamp, user, status, error) |
| `_filter` | hidden | last applied filter |
| Analysis sheets (8) | rebuilt on each filter apply | derived views |

### Migrating from a pre-archive snapshot

If you're upgrading an existing sheet that has `raw_data` but no `_baseline_archive`: run **🚦 Waze → 🔄 העבר נתונים קיימים לארכיון אגרגטיבי** from the menu. It walks `raw_data` once, resolves each row's route+direction, populates `_baseline_archive`, and flags the rows so they aren't double-counted later.

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
