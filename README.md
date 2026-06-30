# Leekes Buyer Dashboard

A self-contained, offline web dashboard for buyers. Drag-drop a buyer export
Excel file and it reads the **Summary** sheet, then builds one tab per
*Purpose*, each with KPIs, charts and action tables tailored to that sheet.

## Run it

No build step, no server. Just open `index.html` in a browser and drag your
`.xlsx` file onto the page. Works fully offline — the parsing (SheetJS) and
charting (Chart.js) libraries are vendored in `lib/`.

## Project structure

| File | Purpose |
|------|---------|
| `index.html` | Page shell, drop zone, tab containers |
| `styles.css` | All styling |
| `app.js`     | All logic: Excel parsing, Summary → tabs, the per-sheet renderers |
| `lib/`       | Vendored `xlsx.full.min.js` + `chart.umd.min.js` (committed so it runs offline) |

### How `app.js` is organised
- **Utilities** — number/currency/percent parsing & formatting
- **`tableFromSheet`** — finds the header row beneath the "Query Criteria" line, returns `{headers, rows, get(), has()}`
- **Renderers** — `renderMagento`, `renderWebSales`, `renderReturns`, `renderItemInventory`, plus a generic fallback. `pickRenderer()` chooses one by detecting columns.
- **`parseSummary` / `matchSheet`** — reads the Summary sheet, groups by Purpose, fuzzy-matches each *Extract* to a worksheet name (e.g. `Item Inventory Extract` → `Item Inventory Extract V4`)
- **`buildDashboard` / `renderTab`** — wires tabs to renderers
- **`buildTable`** — the reusable sortable / searchable / filterable table component

To add support for a new sheet type: write a `renderX(root, table)` function and add a detection rule to `pickRenderer()`.

## Data — do not commit

Buyer export files (`*.xlsx`) contain sensitive sales, margin, vendor and
return data. They are `.gitignore`d. Keep your own data file locally; never
push it to the remote.

## Collaboration workflow

This repo is the shared source of truth — **do not** keep it in a
OneDrive/SharePoint-synced folder (sync corrupts `.git`). Use the Git remote
to share instead.

```bash
# one-time
git clone <remote-url>
cd dashboard

# each change
git checkout -b feature/short-description
# ...edit, test in browser...
git add -A
git commit -m "Describe the change"
git push -u origin feature/short-description
# then open a Pull Request for review before merging to main
```

Keep `main` deployable; review changes via PRs.
