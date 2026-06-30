/* ============================================================
   Leekes Buyer Dashboard
   Drag-drop a buyer export Excel file → the dashboard detects
   each sheet's type and builds one tab per recognised purpose.
   No Summary sheet required.
   ============================================================ */

'use strict';

// Leekes brand palette for charts
const PALETTE = [
  '#133250', // navy
  '#c4122e', // red
  '#477628', // green
  '#c07600', // amber
  '#1c3a5e', // navy-mid
  '#91081c', // red-dark
  '#2e4c1a', // green-dark
  '#e07a3f', // orange
  '#5a7fa8', // steel blue
  '#8f3a1c', // rust
  '#688e4f', // sage
  '#4a3f6b', // purple
];

const charts = [];
let WB = null;

/* ---------------------------------------------------------------
   Utilities
   --------------------------------------------------------------- */
const norm = s => String(s == null ? '' : s).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
const $ = sel => document.querySelector(sel);

function num(v) {
  if (v == null || v === '') return NaN;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[£$,%\s]/g, ''));
  return isNaN(n) ? NaN : n;
}
function n0(v) { const x = num(v); return isNaN(x) ? 0 : x; }

function fmtInt(n) { return isFinite(n) ? Math.round(n).toLocaleString('en-GB') : '–'; }
function fmtMoney(n) { return isFinite(n) ? '£' + Math.round(n).toLocaleString('en-GB') : '–'; }
function fmtPct(n, dp = 1) { return isFinite(n) ? n.toFixed(dp) + '%' : '–'; }
function fmtCell(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) return v.toLocaleDateString('en-GB');
  return String(v);
}

/* ---------------------------------------------------------------
   Sheet → table object
   --------------------------------------------------------------- */
function tableFromSheet(ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  let criteria = '';
  let h = 0;
  for (let i = 0; i < Math.min(aoa.length, 6); i++) {
    const row = aoa[i] || [];
    const first = String(row[0] == null ? '' : row[0]);
    if (first.toLowerCase().includes('query criteria')) { criteria = first; continue; }
    if (row.filter(c => c !== null && c !== '').length >= 2) { h = i; break; }
  }
  const headers = (aoa[h] || []).map(c => String(c == null ? '' : c).trim());
  const rows = aoa.slice(h + 1).filter(r => r && r.some(c => c !== null && c !== ''));

  const idx = {};
  headers.forEach((hd, i) => { const k = norm(hd); if (k && !(k in idx)) idx[k] = i; });

  const get = (row, ...names) => {
    for (const nm of names) { const k = norm(nm); if (k in idx) return row[idx[k]]; }
    return undefined;
  };
  const has = (...names) => names.every(nm => norm(nm) in idx);
  const hasAny = (...names) => names.some(nm => norm(nm) in idx);

  return { headers, rows, idx, criteria, get, has, hasAny };
}

/* ---------------------------------------------------------------
   Aggregation helpers
   --------------------------------------------------------------- */
function groupSum(rows, keyFn, valFn) {
  const m = new Map();
  for (const r of rows) {
    const k = keyFn(r);
    if (k == null || k === '') continue;
    m.set(k, (m.get(k) || 0) + (valFn(r) || 0));
  }
  return [...m.entries()].map(([key, value]) => ({ key, value }));
}
function topN(arr, n, desc = true) {
  return [...arr].sort((a, b) => desc ? b.value - a.value : a.value - b.value).slice(0, n);
}

/* ---------------------------------------------------------------
   Chart helpers
   --------------------------------------------------------------- */
function makeChart(canvas, type, labels, datasets, opts = {}) {
  const ctx = canvas.getContext('2d');
  const c = new Chart(ctx, {
    type,
    data: { labels, datasets },
    options: Object.assign({
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: datasets.length > 1, labels: { boxWidth: 12, font: { size: 11 } } },
        tooltip: { callbacks: opts.tooltip || {} }
      },
      scales: opts.scales
    }, opts.options || {})
  });
  charts.push(c);
  return c;
}
function clearCharts() { while (charts.length) { try { charts.pop().destroy(); } catch (e) {} } }

function barChart(canvas, labels, values, label, color, horizontal, moneyAxis) {
  const fmt = moneyAxis ? fmtMoney : fmtInt;
  return makeChart(canvas, 'bar', labels, [{
    label, data: values, backgroundColor: color || PALETTE[0], borderRadius: 4
  }], {
    options: { indexAxis: horizontal ? 'y' : 'x' },
    tooltip: { label: c => `${label}: ${fmt(c.parsed[horizontal ? 'x' : 'y'])}` },
    scales: {
      x: { ticks: { font: { size: 10 }, callback: horizontal ? (v => fmt(v)) : undefined } },
      y: { ticks: { font: { size: 10 }, callback: horizontal ? undefined : (v => fmt(v)) } }
    }
  });
}

/* ---------------------------------------------------------------
   Reusable sortable/searchable/filterable table
   --------------------------------------------------------------- */
function buildTable(container, options) {
  const { columns } = options;
  let data = options.rows.slice();
  let pageSize = options.pageSize || 25;
  let shown = pageSize;
  let sortKey = options.initialSort ? options.initialSort.key : null;
  let sortDesc = options.initialSort ? options.initialSort.desc !== false : true;
  let searchTerm = '';
  const filterState = {};

  const wrap = document.createElement('div');
  const toolbar = document.createElement('div');
  toolbar.className = 'table-toolbar';

  if (options.search !== false) {
    const inp = document.createElement('input');
    inp.type = 'search'; inp.placeholder = 'Search…';
    inp.addEventListener('input', () => { searchTerm = inp.value.toLowerCase(); shown = pageSize; render(); });
    toolbar.appendChild(inp);
  }
  (options.filters || []).forEach((f, fi) => {
    const sel = document.createElement('select');
    const opt0 = document.createElement('option'); opt0.value = ''; opt0.textContent = f.label; sel.appendChild(opt0);
    f.values.forEach(v => { const o = document.createElement('option'); o.value = v; o.textContent = v; sel.appendChild(o); });
    sel.addEventListener('change', () => { filterState[fi] = sel.value; shown = pageSize; render(); });
    toolbar.appendChild(sel);
  });
  const count = document.createElement('span'); count.className = 'table-count';
  toolbar.appendChild(count);
  wrap.appendChild(toolbar);

  const scroll = document.createElement('div'); scroll.className = 'table-scroll';
  const table = document.createElement('table'); table.className = 'data';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  table.appendChild(thead); table.appendChild(tbody);
  scroll.appendChild(table); wrap.appendChild(scroll);

  const moreWrap = document.createElement('div'); moreWrap.className = 'show-more';
  const moreBtn = document.createElement('button');
  moreBtn.addEventListener('click', () => { shown += pageSize * 2; render(); });
  moreWrap.appendChild(moreBtn); wrap.appendChild(moreWrap);

  const htr = document.createElement('tr');
  columns.forEach(col => {
    const th = document.createElement('th');
    th.textContent = col.label;
    if (col.num || col.money || col.pct) th.className = 'num';
    th.addEventListener('click', () => {
      if (sortKey === col.key) sortDesc = !sortDesc; else { sortKey = col.key; sortDesc = true; }
      render();
    });
    htr.appendChild(th);
  });
  thead.appendChild(htr);

  function filtered() {
    let out = data;
    (options.filters || []).forEach((f, fi) => { const v = filterState[fi]; if (v) out = out.filter(r => f.test(r, v)); });
    if (searchTerm) {
      out = out.filter(r => columns.some(c => { const val = r[c.key]; return val != null && String(val).toLowerCase().includes(searchTerm); }));
    }
    if (sortKey) {
      const col = columns.find(c => c.key === sortKey);
      const isNum = col && (col.num || col.money || col.pct);
      out = out.slice().sort((a, b) => {
        let av = a[sortKey], bv = b[sortKey];
        if (isNum) { av = n0(av); bv = n0(bv); return sortDesc ? bv - av : av - bv; }
        av = String(av == null ? '' : av).toLowerCase(); bv = String(bv == null ? '' : bv).toLowerCase();
        return sortDesc ? bv.localeCompare(av) : av.localeCompare(bv);
      });
    }
    return out;
  }

  function render() {
    [...htr.children].forEach((th, i) => {
      const base = columns[i].label;
      th.textContent = columns[i].key === sortKey ? base + (sortDesc ? ' ▾' : ' ▴') : base;
    });
    const rows = filtered();
    tbody.innerHTML = '';
    for (const r of rows.slice(0, shown)) {
      const tr = document.createElement('tr');
      if (options.rowAlert && options.rowAlert(r)) tr.className = 'row-alert';
      columns.forEach(col => {
        const td = document.createElement('td');
        if (col.num || col.money || col.pct) td.className = 'num';
        if (col.wrap) td.className = (td.className + ' wrap').trim();
        let content;
        if (col.fmt) content = col.fmt(r[col.key], r);
        else if (col.money) content = fmtMoney(num(r[col.key]));
        else if (col.pct) content = fmtPct(num(r[col.key]));
        else if (col.num) content = fmtInt(num(r[col.key]));
        else content = fmtCell(r[col.key]);
        if (content instanceof Node) td.appendChild(content); else td.innerHTML = content;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
    count.textContent = `${fmtInt(rows.length)} rows`;
    moreWrap.style.display = rows.length > shown ? 'block' : 'none';
    moreBtn.textContent = `Show more (${fmtInt(rows.length - shown)} hidden)`;
  }

  render();
  container.appendChild(wrap);
}

function flag(text, cls) { const s = document.createElement('span'); s.className = 'flag ' + cls; s.textContent = text; return s; }
function el(tag, cls, html) { const e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; }

function kpiRow(container, items) {
  const row = el('div', 'kpi-row');
  items.forEach(it => {
    const c = el('div', 'kpi' + (it.tone ? ' ' + it.tone : ''));
    c.appendChild(el('div', 'kpi-label', it.label));
    c.appendChild(el('div', 'kpi-value', it.value));
    if (it.sub) c.appendChild(el('div', 'kpi-sub', it.sub));
    row.appendChild(c);
  });
  container.appendChild(row);
}

function chartCard(grid, title, sub, span) {
  const card = el('div', 'card' + (span ? ' span2' : ''));
  card.appendChild(el('h3', null, title));
  if (sub) card.appendChild(el('div', 'card-sub', sub));
  const wrap = el('div', 'chart-wrap');
  const canvas = document.createElement('canvas');
  wrap.appendChild(canvas); card.appendChild(wrap);
  grid.appendChild(card);
  return canvas;
}

function tableCard(container, title, sub) {
  const card = el('div', 'card');
  if (title) card.appendChild(el('h3', null, title));
  if (sub) card.appendChild(el('div', 'card-sub', sub));
  container.appendChild(card);
  return card;
}

/* ===============================================================
   RENDERERS — one per sheet type
   =============================================================== */

/* 1. Publishing review (Magento2 Data) */
function renderMagento(root, t) {
  const rows = t.rows.map(r => ({
    item: t.get(r, 'Item No.', 'Item No', 'ItemNo'),
    desc: t.get(r, 'Till Description', 'Web Product Name', 'Item_Desc') || '',
    vendor: t.get(r, 'Vendor No.', 'Vendor No', 'Vendor_No') || '—',
    images: n0(t.get(r, 'Image Count')),
    descLen: n0(t.get(r, 'Description Length')),
    published: String(t.get(r, 'Published') || '').trim().toLowerCase(),
    discontinued: n0(t.get(r, 'discontinued item', 'discontinued'))
  })).filter(r => r.item != null && r.item !== '');

  const live = rows.filter(r => r.discontinued === 0);
  const pub = live.filter(r => ['yes', 'y', '1', 'true'].includes(r.published));
  const unpub = live.filter(r => !['yes', 'y', '1', 'true'].includes(r.published));
  const pctUnpub = live.length ? (unpub.length / live.length) * 100 : 0;
  const noImg = unpub.filter(r => r.images === 0).length;
  const shortDesc = unpub.filter(r => r.descLen < 250).length;

  kpiRow(root, [
    { label: 'Total items', value: fmtInt(rows.length) },
    { label: 'Live lines', value: fmtInt(live.length), sub: 'not discontinued' },
    { label: 'Published online', value: fmtInt(pub.length), tone: 'good' },
    { label: 'Not published', value: fmtInt(unpub.length), sub: fmtPct(pctUnpub) + ' of live lines', tone: unpub.length ? 'alert' : 'good' },
    { label: 'Unpublished — no image', value: fmtInt(noImg), tone: noImg ? 'warn' : 'good' }
  ]);

  const grid = el('div', 'chart-grid');
  const c1 = chartCard(grid, 'Published vs not published', 'Live lines only');
  makeChart(c1, 'doughnut', ['Published', 'Not published'],
    [{ data: [pub.length, unpub.length], backgroundColor: [PALETTE[2], PALETTE[1]] }],
    { tooltip: { label: c => `${c.label}: ${fmtInt(c.parsed)}` } });

  const byVendor = topN(groupSum(unpub, r => r.vendor, () => 1), 15);
  barChart(chartCard(grid, 'Unpublished live lines by vendor', 'Top 15 vendors'),
    byVendor.map(d => d.key), byVendor.map(d => d.value), 'Unpublished', PALETTE[0], true);

  barChart(chartCard(grid, 'Likely publishing blockers', 'Among unpublished live lines'),
    ['No image', 'Short description (<250)', 'Has image & full desc'],
    [noImg, shortDesc, unpub.filter(r => r.images > 0 && r.descLen >= 250).length],
    'Items', PALETTE[1], false);
  root.appendChild(grid);

  root.appendChild(el('div', 'section-title', 'Unpublished live lines — action list'));
  root.appendChild(el('div', 'note', 'These items are live (not discontinued) but not published online. Flags suggest why: missing imagery or a thin product description usually block publishing.'));
  const vendors = [...new Set(unpub.map(r => r.vendor))].sort();
  buildTable(tableCard(root, ''), {
    rows: unpub, initialSort: { key: 'images', desc: false },
    filters: [{ label: 'All vendors', values: vendors, test: (r, v) => r.vendor === v }],
    rowAlert: r => r.images === 0,
    columns: [
      { key: 'item', label: 'Item No.' },
      { key: 'desc', label: 'Description', wrap: true },
      { key: 'vendor', label: 'Vendor' },
      { key: 'images', label: 'Images', num: true },
      { key: 'descLen', label: 'Desc length', num: true },
      { key: 'flags', label: 'Blocker', fmt: (_, r) => {
          const f = el('span');
          if (r.images === 0) f.appendChild(flag('No image', 'red'));
          if (r.descLen < 250) f.appendChild(flag('Short desc', 'amber'));
          if (r.images > 0 && r.descLen >= 250) f.appendChild(flag('Ready', 'green'));
          return f;
        } }
    ]
  });
}

/* 2. Web sales (Web Sales Top X) */
function renderWebSales(root, t) {
  const rows = t.rows.map(r => {
    const qty = n0(t.get(r, 'qtysold'));
    const value = n0(t.get(r, 'valuesold'));
    const cost = n0(t.get(r, 'costamount'));
    return {
      item: t.get(r, 'ItemNo', 'Item No'),
      desc: t.get(r, 'ItemDescription') || '',
      section: t.get(r, 'sectionname', 'Section') || '—',
      qty, value, cost,
      gp: value ? ((value - cost) / value) * 100 : 0
    };
  }).filter(r => r.item != null && r.item !== '');

  const tQty = rows.reduce((s, r) => s + r.qty, 0);
  const tVal = rows.reduce((s, r) => s + r.value, 0);
  const tCost = rows.reduce((s, r) => s + r.cost, 0);
  const gp = tVal ? ((tVal - tCost) / tVal) * 100 : 0;
  const sd = t.rows[0] ? t.get(t.rows[0], 'startdate') : null;
  const ed = t.rows[0] ? t.get(t.rows[0], 'enddate') : null;

  kpiRow(root, [
    { label: 'Units sold', value: fmtInt(tQty) },
    { label: 'Sales value', value: fmtMoney(tVal), tone: 'good' },
    { label: 'Cost of sales', value: fmtMoney(tCost) },
    { label: 'Gross margin', value: fmtPct(gp), sub: fmtMoney(tVal - tCost) + ' profit' },
    { label: 'Lines sold', value: fmtInt(rows.length), sub: (sd ? fmtCell(sd) + ' → ' + fmtCell(ed) : '') }
  ]);

  const grid = el('div', 'chart-grid');
  const topQ = topN(rows.map(r => ({ key: r.desc || r.item, value: r.qty })), 12);
  barChart(chartCard(grid, 'Top sellers by units', 'Top 12'), topQ.map(d => d.key), topQ.map(d => d.value), 'Units', PALETTE[0], true);

  const topV = topN(rows.map(r => ({ key: r.desc || r.item, value: r.value })), 12);
  barChart(chartCard(grid, 'Top sellers by value', 'Top 12'), topV.map(d => d.key), topV.map(d => d.value), 'Value', PALETTE[1], true, true);

  const bySec = topN(groupSum(rows, r => r.section, r => r.value), 12);
  barChart(chartCard(grid, 'Sales value by section', ''), bySec.map(d => d.key), bySec.map(d => d.value), 'Value', PALETTE[0], true, true);

  const bySecQ = topN(groupSum(rows, r => r.section, r => r.qty), 12);
  barChart(chartCard(grid, 'Units by section', ''), bySecQ.map(d => d.key), bySecQ.map(d => d.value), 'Units', PALETTE[2], true);
  root.appendChild(grid);

  root.appendChild(el('div', 'section-title', 'Web sales by line'));
  root.appendChild(el('div', 'note', 'Sorted by sales value. GP% below the overall margin (' + fmtPct(gp) + ') is highlighted — review pricing or supplier cost on those lines.'));
  const sections = [...new Set(rows.map(r => r.section))].sort();
  buildTable(tableCard(root, ''), {
    rows, initialSort: { key: 'value', desc: true },
    filters: [{ label: 'All sections', values: sections, test: (r, v) => r.section === v }],
    rowAlert: r => r.gp < gp,
    columns: [
      { key: 'item', label: 'Item No.' },
      { key: 'desc', label: 'Description', wrap: true },
      { key: 'section', label: 'Section' },
      { key: 'qty', label: 'Units', num: true },
      { key: 'value', label: 'Value', money: true },
      { key: 'cost', label: 'Cost', money: true },
      { key: 'gp', label: 'GP %', pct: true }
    ]
  });
}

/* 3. Returns (Return Rates Summary) */
function renderReturns(root, t) {
  const rows = t.rows.map(r => ({
    vendor: t.get(r, 'vendorname') || t.get(r, 'vendorno') || '—',
    vendorNo: t.get(r, 'vendorno') || '',
    received: n0(t.get(r, 'qtyreceived')),
    returned: n0(t.get(r, 'returnedqty')),
    rate: n0(t.get(r, 'ReturnRate')),
    faulty: n0(t.get(r, 'faultysupplierreturns')),
    faultyRate: n0(t.get(r, 'FaultySupplierReturnRate')),
    supErr: n0(t.get(r, 'SupplierErrorReturns')),
    leekesErr: n0(t.get(r, 'LeekesErrorReturns')),
    web: n0(t.get(r, 'webreturns'))
  })).filter(r => r.vendor && r.vendor !== '—');

  const tRec = rows.reduce((s, r) => s + r.received, 0);
  const tRet = rows.reduce((s, r) => s + r.returned, 0);
  const tFaulty = rows.reduce((s, r) => s + r.faulty, 0);
  const overall = tRec ? (tRet / tRec) * 100 : 0;
  const faultyRate = tRec ? (tFaulty / tRec) * 100 : 0;

  kpiRow(root, [
    { label: 'Units received', value: fmtInt(tRec) },
    { label: 'Units returned', value: fmtInt(tRet) },
    { label: 'Overall return rate', value: fmtPct(overall), tone: overall > 5 ? 'alert' : 'good' },
    { label: 'Faulty (supplier)', value: fmtInt(tFaulty), sub: fmtPct(faultyRate) + ' of received', tone: 'warn' },
    { label: 'Vendors', value: fmtInt(rows.length) }
  ]);

  const grid = el('div', 'chart-grid');
  const sig = rows.filter(r => r.received >= 20);
  const topRate = topN(sig.map(r => ({ key: r.vendor, value: r.rate })), 12);
  barChart(chartCard(grid, 'Highest return-rate vendors', 'Vendors with ≥20 units received'),
    topRate.map(d => d.key), topRate.map(d => d.value), 'Return rate %', PALETTE[1], true);

  const topRet = rows.slice().sort((a, b) => b.returned - a.returned).slice(0, 10);
  const c2 = chartCard(grid, 'Return reasons by vendor', 'Top 10 by units returned', true);
  makeChart(c2, 'bar', topRet.map(r => r.vendor), [
    { label: 'Faulty (supplier)', data: topRet.map(r => r.faulty), backgroundColor: PALETTE[1] },
    { label: 'Supplier error', data: topRet.map(r => r.supErr), backgroundColor: PALETTE[3] },
    { label: 'Leekes error', data: topRet.map(r => r.leekesErr), backgroundColor: PALETTE[7] },
    { label: 'Web returns', data: topRet.map(r => r.web), backgroundColor: PALETTE[0] }
  ], {
    options: { scales: { x: { stacked: true, ticks: { font: { size: 10 } } }, y: { stacked: true, ticks: { font: { size: 10 } } } } }
  });
  root.appendChild(grid);

  root.appendChild(el('div', 'section-title', 'Return rates by vendor'));
  root.appendChild(el('div', 'note', 'Rows above the overall return rate (' + fmtPct(overall) + ') are highlighted. The reason columns split faults between the supplier, supplier admin errors, Leekes errors and web returns.'));
  buildTable(tableCard(root, ''), {
    rows, initialSort: { key: 'rate', desc: true },
    rowAlert: r => r.rate > overall && r.received >= 20,
    columns: [
      { key: 'vendor', label: 'Vendor' },
      { key: 'received', label: 'Received', num: true },
      { key: 'returned', label: 'Returned', num: true },
      { key: 'rate', label: 'Return %', pct: true },
      { key: 'faulty', label: 'Faulty (sup.)', num: true },
      { key: 'faultyRate', label: 'Faulty %', pct: true },
      { key: 'supErr', label: 'Sup. error', num: true },
      { key: 'leekesErr', label: 'Leekes error', num: true },
      { key: 'web', label: 'Web', num: true }
    ]
  });
}

/* 4. Item sales analysis (Item Inventory Extract) */
function renderItemInventory(root, t) {
  const rows = t.rows.map(r => {
    const salesQty = n0(t.get(r, 'SalesQty'));
    const salesNet = n0(t.get(r, 'SalesNet'));
    const stock = n0(t.get(r, 'total_Stock', 'totalStock'));
    return {
      item: t.get(r, 'Item_No', 'ItemNo'),
      desc: t.get(r, 'Item_Desc') || '',
      section: t.get(r, 'Section_Name') || '—',
      vendor: t.get(r, 'Vendor_Name') || '—',
      season: String(t.get(r, 'Season_Code') || '').trim() || '—',
      retail: n0(t.get(r, 'Retail_Price')),
      gp: n0(t.get(r, 'GP_Percentage')),
      stock, salesQty, salesNet,
      salesCost: n0(t.get(r, 'SalesCost')),
      invCost: n0(t.get(r, 'InventoryatCost', 'InventoryAtCost')),
      invRetail: n0(t.get(r, 'InventoryAtRetail'))
    };
  }).filter(r => r.item != null && r.item !== '');

  const tNet = rows.reduce((s, r) => s + r.salesNet, 0);
  const tQty = rows.reduce((s, r) => s + r.salesQty, 0);
  const tInvCost = rows.reduce((s, r) => s + r.invCost, 0);
  const stockOuts = rows.filter(r => r.stock <= 0 && r.salesQty > 0);
  const marginFlags = rows.filter(r => r.gp !== 0 && (r.gp < 35 || r.gp > 70));

  kpiRow(root, [
    { label: 'Net sales', value: fmtMoney(tNet), tone: 'good' },
    { label: 'Units sold', value: fmtInt(tQty) },
    { label: 'Inventory at cost', value: fmtMoney(tInvCost) },
    { label: 'Lines', value: fmtInt(rows.length) },
    { label: 'Selling but out of stock', value: fmtInt(stockOuts.length), tone: stockOuts.length ? 'alert' : 'good' },
    { label: 'Margin outside 35–70%', value: fmtInt(marginFlags.length), tone: marginFlags.length ? 'warn' : 'good' }
  ]);

  const grid = el('div', 'chart-grid');
  const topItems = topN(rows.map(r => ({ key: r.desc || r.item, value: r.salesNet })), 12);
  barChart(chartCard(grid, 'Best-selling items by net sales', 'Top 12'), topItems.map(d => d.key), topItems.map(d => d.value), 'Net sales', PALETTE[0], true, true);

  const bySection = topN(groupSum(rows, r => r.section, r => r.salesNet), 12);
  barChart(chartCard(grid, 'Net sales by section', ''), bySection.map(d => d.key), bySection.map(d => d.value), 'Net sales', PALETTE[0], true, true);

  const byVendor = topN(groupSum(rows, r => r.vendor, r => r.salesNet), 12);
  barChart(chartCard(grid, 'Best-performing vendors', 'Top 12 by net sales'), byVendor.map(d => d.key), byVendor.map(d => d.value), 'Net sales', PALETTE[2], true, true);

  const lowM = rows.filter(r => r.gp !== 0 && r.gp < 35).length;
  const midM = rows.filter(r => r.gp >= 35 && r.gp <= 70).length;
  const hiM = rows.filter(r => r.gp > 70).length;
  const c4 = chartCard(grid, 'Margin check', 'GP% banding across lines');
  makeChart(c4, 'doughnut', ['< 35% (low)', '35–70% (healthy)', '> 70% (check)'],
    [{ data: [lowM, midM, hiM], backgroundColor: [PALETTE[1], PALETTE[2], PALETTE[3]] }],
    { tooltip: { label: c => `${c.label}: ${fmtInt(c.parsed)} lines` } });
  root.appendChild(grid);

  const sections = [...new Set(rows.map(r => r.section))].sort();
  const vendors = [...new Set(rows.map(r => r.vendor))].sort();

  root.appendChild(el('div', 'section-title', 'Item-level sales & stock'));
  root.appendChild(el('div', 'note', 'Sort any column to find best sellers by item, section or vendor. Lines with GP% under 35% or over 70% are highlighted for a margin review.'));
  buildTable(tableCard(root, ''), {
    rows, initialSort: { key: 'salesNet', desc: true }, pageSize: 25,
    filters: [
      { label: 'All sections', values: sections, test: (r, v) => r.section === v },
      { label: 'All vendors', values: vendors, test: (r, v) => r.vendor === v }
    ],
    rowAlert: r => r.gp !== 0 && (r.gp < 35 || r.gp > 70),
    columns: [
      { key: 'item', label: 'Item No.' },
      { key: 'desc', label: 'Description', wrap: true },
      { key: 'section', label: 'Section' },
      { key: 'vendor', label: 'Vendor' },
      { key: 'season', label: 'Season' },
      { key: 'salesQty', label: 'Units', num: true },
      { key: 'salesNet', label: 'Net sales', money: true },
      { key: 'gp', label: 'GP %', pct: true },
      { key: 'stock', label: 'Stock', num: true }
    ]
  });

  root.appendChild(el('div', 'section-title', 'Gap analysis — selling but out of stock'));
  root.appendChild(el('div', 'note', 'These lines sold during the period but currently show zero (or negative) total stock — potential lost sales. Prioritise re-ordering the top sellers.'));
  buildTable(tableCard(root, ''), {
    rows: stockOuts, initialSort: { key: 'salesQty', desc: true }, pageSize: 15,
    columns: [
      { key: 'item', label: 'Item No.' },
      { key: 'desc', label: 'Description', wrap: true },
      { key: 'vendor', label: 'Vendor' },
      { key: 'salesQty', label: 'Units sold', num: true },
      { key: 'salesNet', label: 'Net sales', money: true },
      { key: 'stock', label: 'Stock', num: true }
    ]
  });

  const bySeason = new Map();
  for (const r of rows) {
    const k = r.season;
    if (!bySeason.has(k)) bySeason.set(k, { season: k, lines: 0, stock: 0, salesQty: 0, salesNet: 0, invCost: 0 });
    const g = bySeason.get(k);
    g.lines++; g.stock += r.stock; g.salesQty += r.salesQty; g.salesNet += r.salesNet; g.invCost += r.invCost;
  }
  const seasonRows = [...bySeason.values()].map(g => {
    const sellThrough = (g.stock + g.salesQty) > 0 ? (g.salesQty / (g.stock + g.salesQty)) * 100 : 0;
    let strategy, cls;
    if (g.salesQty === 0 && g.stock > 0) { strategy = 'Clear — markdown'; cls = 'red'; }
    else if (sellThrough < 25 && g.stock > 0) { strategy = 'Slow — promote / markdown'; cls = 'amber'; }
    else if (sellThrough > 75) { strategy = 'Strong — reorder'; cls = 'green'; }
    else { strategy = 'Monitor'; cls = 'grey'; }
    return Object.assign(g, { sellThrough, strategy, cls });
  });
  root.appendChild(el('div', 'section-title', 'Clearance strategy by season code'));
  root.appendChild(el('div', 'note', 'Sell-through = units sold ÷ (units sold + stock on hand). Low sell-through with stock remaining flags clearance candidates; high sell-through suggests reorder.'));
  buildTable(tableCard(root, ''), {
    rows: seasonRows, initialSort: { key: 'invCost', desc: true }, pageSize: 20, search: false,
    columns: [
      { key: 'season', label: 'Season code' },
      { key: 'lines', label: 'Lines', num: true },
      { key: 'stock', label: 'Stock units', num: true },
      { key: 'salesQty', label: 'Units sold', num: true },
      { key: 'salesNet', label: 'Net sales', money: true },
      { key: 'invCost', label: 'Stock @ cost', money: true },
      { key: 'sellThrough', label: 'Sell-through', pct: true },
      { key: 'strategy', label: 'Suggested strategy', fmt: (_, r) => flag(r.strategy, r.cls) }
    ]
  });
}

/* Generic fallback */
function renderGeneric(root, t) {
  root.appendChild(el('div', 'note', 'No specialised view matched this sheet. All data is shown below with sortable columns.'));
  const objRows = t.rows.map(r => { const o = {}; t.headers.forEach((h, i) => o['c' + i] = r[i]); return o; });
  const numericCols = t.headers.map((h, i) => {
    let count = 0;
    for (const r of t.rows.slice(0, 200)) { const v = num(r[i]); if (!isNaN(v)) count++; }
    return { i, h, numeric: count > Math.min(t.rows.length, 200) * 0.6 };
  });
  buildTable(tableCard(root, ''), {
    rows: objRows, pageSize: 30,
    columns: numericCols.filter(c => c.h).map(c => ({ key: 'c' + c.i, label: c.h, num: c.numeric }))
  });
}

/* Renderer detection */
function pickRenderer(t) {
  if (t.has('Published') && (t.has('Image Count') || t.hasAny('discontinued item', 'discontinued'))) return renderMagento;
  if (t.has('qtysold') && t.has('valuesold')) return renderWebSales;
  if (t.has('ReturnRate') && t.hasAny('vendorname', 'vendorno')) return renderReturns;
  if (t.hasAny('SalesQty') && t.hasAny('total_Stock', 'GP_Percentage', 'Season_Code')) return renderItemInventory;
  return renderGeneric;
}

/* ===============================================================
   Fixed tab metadata — keyed to each renderer function.
   No Summary sheet needed.
   =============================================================== */
const RENDERER_META = new Map([
  [renderMagento, {
    purpose: 'Review live lines not published online',
    tasks: ['Review the amount of live lines not published online']
  }],
  [renderWebSales, {
    purpose: 'Understand web sales',
    tasks: [
      'Web sales analysis — top sellers and best performing vendors',
      'Gap analysis for web, stock outs'
    ]
  }],
  [renderReturns, {
    purpose: 'Understand the reason for faults and % return rate',
    tasks: ['Return summary by vendor']
  }],
  [renderItemInventory, {
    purpose: 'Item sales analysis',
    tasks: [
      'Item level analysis — best sellers by item, section & vendor',
      'Gap analysis for item sales, stock outs, margin check <35% >70%',
      'Clearance item strategy — review sales on season code, suggest best sales strategy'
    ]
  }]
]);

const RENDERER_ORDER = [renderMagento, renderWebSales, renderReturns, renderItemInventory];

/* ===============================================================
   Build the dashboard — scan each sheet directly
   =============================================================== */
function buildDashboard(wb, fileName) {
  clearCharts();
  WB = wb;

  const recognised = [];
  const generic = [];

  for (const sheetName of wb.SheetNames) {
    if (norm(sheetName) === 'summary') continue;
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const table = tableFromSheet(ws);
    const renderer = pickRenderer(table);
    const meta = RENDERER_META.get(renderer);
    const tab = { sheetName, table, renderer, purpose: meta ? meta.purpose : sheetName, tasks: meta ? meta.tasks : [] };
    if (meta) recognised.push(tab); else generic.push(tab);
  }

  recognised.sort((a, b) => RENDERER_ORDER.indexOf(a.renderer) - RENDERER_ORDER.indexOf(b.renderer));
  const tabs = [...recognised, ...generic];

  if (!tabs.length) {
    showError('No data sheets found in this file.');
    return;
  }

  const tabBar = $('#tabBar');
  const content = $('#tabContent');
  tabBar.innerHTML = ''; content.innerHTML = '';

  tabs.forEach((tab, i) => {
    const btn = el('div', 'tab' + (i === 0 ? ' active' : ''), tab.purpose);
    btn.title = tab.purpose;
    btn.addEventListener('click', () => {
      [...tabBar.children].forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      renderTab(content, tab);
    });
    tabBar.appendChild(btn);
  });

  renderTab(content, tabs[0]);
  $('#fileMeta').textContent = fileName + '  •  ' + tabs.length + (tabs.length === 1 ? ' tab' : ' tabs');
  $('#loadNewBtn').classList.remove('hidden');
  $('#landing').classList.add('hidden');
  $('#dashboard').classList.remove('hidden');
}

function renderTab(content, tab) {
  clearCharts();
  content.innerHTML = '';
  const intro = el('div', 'tab-intro');
  intro.appendChild(el('h1', null, tab.purpose));
  intro.appendChild(el('div', 'purpose', 'Source: ' + tab.sheetName));
  if (tab.tasks && tab.tasks.length) {
    const ul = el('ul');
    tab.tasks.forEach(t => ul.appendChild(el('li', null, t)));
    intro.appendChild(ul);
  }
  if (tab.table && tab.table.criteria) intro.appendChild(el('div', 'criteria', tab.table.criteria));
  content.appendChild(intro);

  try { tab.renderer(content, tab.table); }
  catch (e) { console.error(e); showError('Error rendering this tab: ' + e.message); }
  window.scrollTo({ top: 0 });
}

/* ===============================================================
   File handling + drag/drop
   =============================================================== */
function showError(msg) {
  const box = $('#errorBox');
  box.textContent = msg; box.classList.remove('hidden');
  setTimeout(() => box.classList.add('hidden'), 7000);
}

function handleFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
      buildDashboard(wb, file.name);
    } catch (err) {
      console.error(err);
      showError('Could not read that file — is it a valid .xlsx? (' + err.message + ')');
    }
  };
  reader.onerror = () => showError('Failed to read the file.');
  reader.readAsArrayBuffer(file);
}

function wire() {
  const fileInput = $('#fileInput');

  fileInput.addEventListener('change', e => handleFile(e.target.files[0]));
  $('#loadNewBtn').addEventListener('click', () => {
    $('#dashboard').classList.add('hidden');
    $('#landing').classList.remove('hidden');
    $('#loadNewBtn').classList.add('hidden');
    $('#fileMeta').textContent = '';
    fileInput.value = '';
  });

  $('#dropZone').addEventListener('click', () => fileInput.click());

  let depth = 0;
  ['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => {
    e.preventDefault();
    if (ev === 'dragenter') depth++;
    $('#dragOverlay').classList.remove('hidden');
    $('#dropZone').classList.add('drag');
  }));
  document.addEventListener('dragleave', e => {
    e.preventDefault();
    if (--depth <= 0) { depth = 0; $('#dragOverlay').classList.add('hidden'); $('#dropZone').classList.remove('drag'); }
  });
  document.addEventListener('drop', e => {
    e.preventDefault();
    depth = 0; $('#dragOverlay').classList.add('hidden'); $('#dropZone').classList.remove('drag');
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
}

document.addEventListener('DOMContentLoaded', wire);
