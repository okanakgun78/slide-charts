# Slide Charts — project brief

PowerPoint add-in (Office.js web add-in) that draws think-cell-style consulting charts as native,
grouped PowerPoint shapes. Owner uses PowerPoint for Mac (Microsoft 365, 16.114+), installed by
sideloading `manifest.xml` into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`.
Hosting: GitHub Pages at https://okanakgun78.github.io/slide-charts/ — every push to `main` builds,
tests and deploys (`.github/workflows/deploy.yml`). Updates reach the user on reopening the panel.

## Architecture
- `src/engine/` — pure layout: spec + frame + style + measure → primitives (rect/text/line/shape).
  `category.ts` (column/bar stacked|clustered|percent, line, combo), `gantt.ts`, `calendar.ts`
  (timeline header over free-text cells), `timeline.ts` (shared tiers + header), `dates.ts`,
  `spec.ts` (types live in `types.ts`; samples, palette A, migration), `svg.ts` (preview).
- `src/ppt/render.ts` — primitives → PowerPoint shapes; group + JSON spec in tag `SLIDECHARTS`
  (`StoredChart`: spec, frame, drawn bbox). Redraw = read tag → refit frame to the group's current
  bounds → delete group → draw → reselect. `probes.ts` = Diagnostics panel checks.
- `src/ui/` — data sheets (`sheet.ts` category, `ganttSheet.ts`, `calendarSheet.ts`), parsing (TSV with
  quoted multi-line cells), settings (localStorage).
- `src/taskpane.ts` — panel orchestration: insert, selection-driven editor, debounced live redraw.

## Facts measured on the user's Mac (keep these)
- All PowerPointApi sets 1.1–1.10 supported. No native chart API, no arrowheads, no shape events.
- **Add options with width/height 0 become 72 pt** → lines use `LINE_EPS` (0.01); diagonal lines are
  flat lines rotated about their centre (`lineGeometry`).
- Deleting a group by id works (`removeGroup` keeps fallbacks). Grouping, tags (60 KB ok), rotation,
  shape adjustments (pie: degrees clockwise from 3 o'clock, normalised to -180..180) all work.
- Text: PowerPoint width ≈ 1.04 × canvas measure for the theme font; calibrated once per session.
- Office.js constants (`PowerPoint.X.y`) are NOT available at module load — use string literals.
- Draw ~52 shapes ≈ 0.18 s; redraw ≈ 0.22 s.
- Line breaks inside a text box: the API does not document the separator; `detectLineBreak` tests
  "\n", "\r", "\v" once per session (in `ensureCalibrated`) and the renderer uses the winner.

## Conventions
- First data-sheet row = top segment in columns (think-cell); first series leftmost in bars.
- Palette A: navy ramp `#003760 → #BFD0E2`, greys, highlights gold `#CD9700` / red `#C00000`.
- No value axis; labels inside segments, totals at stack ends, series names right or legend on top.
- Gantt header: any combination of tiers (year/quarter/month/week/day, or auto); finest tier sets the
  grid. Default "filled" style follows the user's sample slide (solid navy cells, white bold text,
  white separators, days as "02 / Mon"); weeks show ISO numbers when days are shown.

## Commands
- `npm test` (vitest, engine + parsing), `npx tsc` (types against @types/office-js),
  `npm run e2e` (builds, then drives the panel in Chromium against `test/e2e/mock-office.js`,
  which enforces load-before-read and the Mac 72 pt quirk), `npm run gallery` (all chart types → HTML).
- Run all three before pushing. Bump `APP_VERSION` in `src/ppt/render.ts` and `package.json`.

## Scope
Done (Phase 1): C1 column/bar, C2 line + combo, C7 Gantt, F1 labels/totals/legend, D1 data sheet.
Added on request: Gantt timeline tiers in the user's sample style (v0.5); Calendar chart (v0.6) —
date range, any tiers, rows with height/style/fill/alignment, cells keyed by ISO date of the column.
Remaining (Phase 2): C3 waterfall, C4 Mekko (% and units), F2 CAGR arrows, F3 difference arrows,
F4 value lines / category gaps / axis breaks, F5 sorting + same scale, F6 better label placement,
E1 Harvey balls + checkboxes, E3 agenda slides. Deferred: live Excel link.
