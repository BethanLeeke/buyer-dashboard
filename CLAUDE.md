# Leekes Buyer Dashboard — Claude Code guide

## What this project is
A self-contained, offline, single-page web dashboard for Leekes buyers. Users drag-and-drop an Excel export file; the dashboard detects each sheet type and renders one tab per recognised purpose. No server, no build step, no npm.

## Working location
**Always edit in `C:\dev\buyer-dashboard`** — not the OneDrive folder. The OneDrive folder (`C:\Users\933401\OneDrive - Leekes\Documents\Leekes Buyer Dashboard`) only holds reference files and must not be edited.

## Deployment
Push to `master` → GitHub Pages auto-redeploys to https://bethanleeke.github.io/buyer-dashboard/ in ~1–2 minutes.

```
git add -A
git commit -m "describe the change"
git push
```

## File structure
```
index.html          page shell — topbar, dropzone, tab bar, tab content
styles.css          all CSS; Leekes brand variables at :root
app.js              all logic — no other JS files
lib/
  xlsx.full.min.js  SheetJS — vendored, do not replace
  chart.umd.min.js  Chart.js 4.4.3 — vendored, do not replace
```

## Do not
- **Never commit `.xlsx`, `.xls`, `.xlsm`, or `.csv` files.** They contain commercially sensitive sales, margin, vendor, and return data. They are excluded via `.gitignore`.
- Do not add npm, a bundler, a build step, or any external runtime dependency. The dashboard must run offline by opening `index.html` directly in a browser.
- Do not add a backend or server-side component.
- Do not replace the vendored libraries in `lib/` with CDN links — the dashboard must work without internet.

## Key patterns in app.js

### tableFromSheet(ws)
Parses a worksheet into `{ headers, rows, idx, criteria, get(), has(), hasAny() }`.
- Skips any "Query Criteria Used" row at the top.
- Finds the first row with ≥ 2 non-blank cells as the header row.
- `get(row, ...columnNames)` — returns the cell value for the first matching column name (normalised: lowercase, alphanumeric only).
- `has(...names)` — true if ALL named columns exist.
- `hasAny(...names)` — true if ANY named column exists.

### pickRenderer(t)
Detects which renderer to use from column presence:
- `renderMagento` — has `Published` + `Image Count` or `discontinued`
- `renderWebSales` — has `qtysold` + `valuesold`
- `renderReturns` — has `ReturnRate` + vendor column
- `renderItemInventory` — has `LLan_SalesQty_DateRange` (or similar store qty columns)

### RENDERER_META
A `Map` keyed to renderer functions. Holds the tab title (`purpose`) and task descriptions. **No Summary sheet is needed.** To add a new sheet type: write a renderer function, add a `pickRenderer` detection rule, and add an entry to `RENDERER_META` and `RENDERER_ORDER`.

### buildTable(container, options)
Reusable sortable/searchable/filterable table component.
- `options.columns` — array of `{ key, label, num?, money?, pct?, fmt?, wrap? }`
- `options.initialSort` — `{ key, desc }`
- `options.filters` — array of `{ label, values, test(row, value) }`
- `options.footer(filteredRows)` — callback returning an array of cell values/nodes for a sticky `<tfoot>` grand-total row
- `options.rowAlert(row)` — callback returning true to highlight a row

### barChart(canvas, labels, values, label, color, horizontal, moneyAxis)
Always pass real string labels as the `labels` array — Chart.js uses them directly on the category axis. The `catCb` inside truncates labels longer than 28 characters with an ellipsis.

## Brand colours (CSS variables)
```
--red: #c4122e       primary brand red (topbar border, active tab)
--navy: #133250      primary navy (KPI card top borders, charts)
--dark: #1c1918      near-black for body text and logo fill
--green: #477628
--amber: #c07600
```

## Collaboration
Two developers work from separate clones of `github.com/BethanLeeke/buyer-dashboard`.
- Create a feature branch for each change: `git checkout -b feature/my-change`
- Open a PR on GitHub → merge to `master` → auto-deploys
- Do not push directly to `master` for anything beyond trivial fixes
