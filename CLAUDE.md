# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-page static website ("Contributing Factor Coding Review") that presents 376
aviation accident contributing factors from the KNKT investigation database, each coded
to the CICTT *Air Traffic Causal & Contributory Factors* taxonomy (Domain / Discipline /
Element). It renders a filterable table plus six charts. Deployed to GitHub Pages;
`main` branch, `/` root, `.nojekyll` present so files are served verbatim.

The repo is the deployable bundle only. The source workbook, taxonomy PDF, and backup
copies live one level up in `../Document/` (not in git). See
`memory/investigation-database-factor-coding.md` for how columns N–P were coded.

## Files

| File | Role |
|------|------|
| `index.html` | Entire app — inline CSS + vanilla JS, no build step, no framework, no external JS/CSS except Google Fonts. |
| `data.json` | The 376 factor records. **This is the file swapped on every data update.** |
| `build-data.ps1` | Regenerates `data.json` (and the fallback copy inside `index.html`) from the Excel workbook via Excel COM automation. |

## Data update workflow

1. Add rows to `../Document/Investigation Database - Contributing Factors Update 2026.xlsx`
   — one row per contributing factor. Columns A/B/E/G/I/K/L are source data; **N/O/P**
   (Domain/Discipline/Element) are hand-coded to the CICTT taxonomy.
2. Run `.\build-data.ps1` from this folder (requires Microsoft Excel installed; close the
   workbook first). It reads sheet 1, header row 2, data from row 3, columns A–P; rewrites
   `data.json`; patches the `<script id="data">` block in `index.html`; prints a
   domain-distribution and year-range summary to check.
3. Commit `data.json` **and** `index.html`, push. GitHub Pages redeploys in ~1 min.

`build-data.ps1 -Workbook "path"` points at a workbook elsewhere. `Clean()` in the script
collapses all whitespace runs, which also repairs multi-line cells in the workbook.

## Architecture notes

- **Data loading** (`index.html` bottom script): on a web server the page `fetch`es
  `data.json` (`cache: no-store`); opened as a bare `file://` it falls back to the
  embedded `<script id="data" type="application/json">` copy. Both paths call
  `initReview(ROWS)`. Keep the two data copies in sync — `build-data.ps1` does this.
- **Record fields** (`build-data.ps1`): `row, date, year, operator, reg` (workbook col F,
  Registration), `location` (col H), `actype, occ, factor, ident, domain, discipline, element`.
  The table's Aircraft/occurrence cell renders `actype (reg)`, e.g. `Cessna C172 (PK-APA)`;
  Location is its own column after Operator.
- **`initReview(ROWS)`** (`index.html:735`) builds everything: the domain count tiles,
  the table (with search/filter/highlight), and the six charts.
- **Events vs. factors**: rows sharing `[date|year|operator|actype|occ]` (`evKey`/`gKey`)
  are one "event". Result counts and several charts count distinct events (~183), not
  rows. The table bands each event with one of two alternating background tints and
  recomputes banding on every filter so dividers land on the first visible row.
  Non-contiguous rows for the same accident show as two bands but count once.
- **Charts** (`// ===== overview charts`): six dependency-free HTML/CSS charts (no
  library, no canvas/SVG) — domain mix, discipline ranking, factors-per-event histogram,
  factors by year, occurrence-type × domain, Identification × Domain heatmap. All follow
  the active filter and rebuild in `apply()`; clicking a bar/segment sets the domain tile
  or search box. `DOMAIN_ORDER` fixes domain sort order and maps to the domain CSS vars.
- **Theming**: light/dark toggle, choice saved to `localStorage` key `fcr-theme`, applied
  by a no-flash script in `<head>` (`data-theme` attribute); follows OS theme when unset.
  Charts are theme-aware through CSS custom properties.

## Working on index.html

The file is ~215 KB — most of it is the inline `data.json` fallback on the
`<script id="data">` line and the CSS. When editing JS or CSS, target specific line
ranges rather than reading the whole file. Never hand-edit the `<script id="data">`
content — regenerate it with `build-data.ps1`.

README.md is written in Indonesian and is aimed at a non-technical maintainer publishing
via the GitHub web UI; keep it that way.
