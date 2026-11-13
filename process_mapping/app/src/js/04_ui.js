/* =====================================================================
   UI: app shell, views, editors
   ===================================================================== */
let PRJ = null;
const UI = { tab: 'dashboard', state: 'current', sel: null, selType: null, zoom: {}, connectFrom: null, spMode: 'select', drawCat: 'walk', draft: [], matrix: false, actFilter: 'All', confirm: null };
const TABS = [
  ['dashboard', 'Dashboard'], ['flow', 'Process Flow'], ['vsm', 'Value Stream'], ['spaghetti', 'Spaghetti'], ['swimlane', 'Swimlane'], ['pam', 'Activity Map'],
  ['data', 'Step Data'], ['opps', 'Waste & Opportunities'], ['actions', 'Actions'], ['finance', 'Finance Validation'], ['compare', 'Current vs Future'],
  ['checklist', 'Floor-Walk Checklist'], ['formulas', 'Formulas & Sources'], ['io', 'Lines & Export'],
];
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
const LINE = () => PRJ.lines.find((l) => l.id === PRJ.activeLine) || PRJ.lines[0];
const ST = () => lineState(LINE(), UI.state);
const PAL = () => PALETTES[isDark() ? 'dark' : 'light'];
function isDark() { const t = document.documentElement.getAttribute('data-theme'); return t ? t === 'dark' : window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches; }

function chip(s, text) { const c = s === 'VER' ? 'ok' : s === 'EST' ? 'est' : 'nyc'; return `<span class="chip ${c}">${esc(text || statusLabel(s))}</span>`; }
function fieldChip(f) { const s = fstat(f); return `<span class="chip ${s === 'VER' ? 'ok' : s === 'EST' ? 'est' : 'nyc'}">${s === 'NYC' ? NYC : STATUS[s]}</span>`; }
function opt(list, cur) { return list.map((o) => { const [v, l] = Array.isArray(o) ? o : [o, o]; return `<option value="${esc(v)}"${String(v) === String(cur) ? ' selected' : ''}>${esc(l)}</option>`; }).join(''); }
function toast(msg, kind = '') { const t = $('#toast'); t.textContent = msg; t.className = 'toast show ' + kind; clearTimeout(toast._t); toast._t = setTimeout(() => (t.className = 'toast'), 3200); }
function stepName(id) { for (const l of PRJ.lines) for (const k of ['current', 'future']) { const s = l[k] && l[k].steps.find((x) => x.id === id); if (s) return s.name; } return '—'; }

function changed(rer = true) { PRJ.updated = new Date().toISOString(); Store.schedule(); if (rer) render(); }

/* ---------------- shell ---------------- */
function renderShell() {
  const l = LINE();
  $('#lineSel').innerHTML = opt(PRJ.lines.map((x) => [x.id, x.name]), l.id);
  $('#stateSel').innerHTML = opt([['current', 'Current state'], ['future', l.future ? 'Future state' : 'Future state (not created)']], UI.state);
  $('#demoBadge').hidden = !PRJ.demo;
  $('#nav').innerHTML = TABS.map(([k, t]) => `<button type="button" class="tab${UI.tab === k ? ' on' : ''}" data-tab="${k}" id="tab-${k}">${t}</button>`).join('');
}
function render() {
  renderShell();
  const v = $('#view');
  const fn = VIEWS[UI.tab] || VIEWS.dashboard;
  if (UI.state === 'future' && !LINE().future && !['dashboard', 'io', 'formulas', 'opps', 'actions', 'finance', 'checklist', 'compare'].includes(UI.tab)) v.innerHTML = noFuture();
  else v.innerHTML = fn();
  afterRender();
  renderPanel();
}
function noFuture() { return `<section class="card empty"><h2>No future state yet for ${esc(LINE().name)}</h2><p>Build the future state from a copy of the current state, then change the steps, layout and data to show the proposed process.</p><button class="btn primary" data-action="createFuture" id="btnCreateFuture">Create future state from current</button></section>`; }

/* ---------------- dashboard ---------------- */
function kpiTile(label, r, f, unit, note) {
  const v = r && r.v != null ? f(r.v) : 'NYC';
  return `<div class="kpi"><div class="kpi-l">${esc(label)}</div><div class="kpi-v${v === 'NYC' ? ' nycv' : ''}">${esc(v)}${v !== 'NYC' && unit ? `<small>${esc(unit)}</small>` : ''}</div>${chip(r ? r.s : 'NYC', r && r.s === 'NYC' && r.v != null ? 'Partial' : undefined)}${note ? `<div class="kpi-n">${esc(note)}</div>` : ''}</div>`;
}
const VIEWS = {};
VIEWS.dashboard = function () {
  const l = LINE(), c = calcState(l, UI.state === 'future' && l.future ? 'future' : 'current'), comp = c.completeness;
  const bn = c.bottleneck;
  const tiles = [
    kpiTile('Total lead time', c.leadTime, (v) => fmt(v / 60, 1), 'hr'),
    kpiTile('Value-added time', c.vaTime, (v) => fmt(v, 1), 'min'),
    kpiTile('Process cycle efficiency', c.pce, (v) => pct(v, 2), ''),
    kpiTile('Takt time', c.takt, (v) => fmt(v, 2), 's/unit'),
    kpiTile('Bottleneck', c.bnECT, (v) => fmt(v, 2), 's', bn ? bn.name + (c.ctCoverage.measured < c.ctCoverage.total ? ` · provisional (${c.ctCoverage.measured}/${c.ctCoverage.total} steps timed)` : '') : 'No cycle times collected'),
    kpiTile('Line capacity', c.capacity, (v) => fmt(v, 0), l.unit + '/hr'),
    kpiTile('Rolled throughput yield', c.rty, (v) => pct(v, 1), ''),
    kpiTile('Labor productivity', c.productivity, (v) => fmt(v, 1), l.unit + '/labor-hr'),
    kpiTile('Changeover time', c.coHours, (v) => fmt(v, 1), 'hr/week'),
    kpiTile('Travel labor', c.travelHrs, (v) => fmt(v, 0), 'hr/yr'),
  ].join('');
  const opps = PRJ.opps.filter((o) => o.lineId === l.id);
  const est = opps.filter((o) => o.fin.status !== 'Validated by Finance' && o.fin.status !== 'Rejected').reduce((a, o) => a + (+o.estAnnual || 0), 0);
  const valid = opps.filter((o) => o.fin.status === 'Validated by Finance').reduce((a, o) => a + (+o.fin.validated || 0), 0);
  const acts = PRJ.actions.filter((a) => opps.some((o) => o.id === a.oppId));
  const overdue = acts.filter((a) => a.status !== 'Done' && a.due && a.due < todayISO()).length;
  const pctGot = comp.tot ? comp.got / comp.tot : 0, pctVer = comp.tot ? comp.ver / comp.tot : 0;
  return `
  <section class="head-row"><div><h1>${esc(l.name)} <span class="muted">· ${UI.state === 'future' && l.future ? 'Future state' : 'Current state'}</span></h1><p class="lede">${esc(l.notes || '')}</p></div>
  <div class="head-actions">${PRJ.demo ? '' : `<button class="btn" data-action="loadDemo" id="btnDemo1">Load demonstration data</button>`}<button class="btn primary" data-tab="checklist">Open floor-walk checklist</button></div></section>
  <section class="card"><div class="card-h"><h2>Data readiness</h2><span class="muted">${comp.got} of ${comp.tot} data points collected · ${comp.ver} verified</span></div>
  <div class="bar"><span class="bar-est" style="width:${(pctGot * 100).toFixed(1)}%"></span><span class="bar-ver" style="width:${(pctVer * 100).toFixed(1)}%"></span></div>
  <div class="legend"><span><i class="sw ver"></i>Verified</span><span><i class="sw est"></i>Estimate</span><span><i class="sw nyc"></i>Not Yet Collected</span></div></section>
  <section class="kpis">${tiles}</section>
  <section class="grid2">
    <div class="card"><div class="card-h"><h2>Effective cycle time by step</h2><span class="muted">vs takt time</span></div><div class="scrollx">${bnChart(l, c)}</div></div>
    <div class="card"><div class="card-h"><h2>Improvement pipeline</h2></div>
      <div class="mini-stats"><div><b>${opps.length}</b><span>opportunities</span></div><div><b>${money(est)}</b><span>estimated / yr, not validated</span></div><div><b>${money(valid)}</b><span>validated by Finance / yr</span></div><div><b class="${overdue ? 'bad' : ''}">${overdue}</b><span>overdue actions</span></div></div>
      <div class="scrollx"><table class="tbl"><thead><tr><th>Waste</th><th>Opportunity</th><th>Finance status</th></tr></thead><tbody>${opps.slice(0, 6).map((o) => `<tr><td>${esc(o.waste)}</td><td>${esc(o.desc)}</td><td>${finChip(o.fin.status)}</td></tr>`).join('') || `<tr><td colspan="3" class="muted">No opportunities yet. Add them in Waste &amp; Opportunities, or use the suggestions generated from your data.</td></tr>`}</tbody></table></div>
    </div>
  </section>`;
};
function finChip(s) { const c = s === 'Validated by Finance' ? 'ok' : s === 'Rejected' ? 'bad' : s === 'Not started' ? 'nyc' : 'est'; return `<span class="chip ${c}">${esc(s)}</span>`; }
function bnChart(l, c) {
  const P = PAL(), steps = applicable(lineState(l, c.key).steps); const rows = steps.map((s) => [s, c.steps[s.id] && c.steps[s.id].ect]);
  const max = Math.max(1, ...rows.map((r) => (r[1] && r[1].v) || 0), c.takt.v || 0) * 1.15;
  const W = 560, rh = 24, H = rows.length * rh + 40, lx = 170, cw = W - lx - 60;
  let b = '';
  rows.forEach(([s, r], i) => {
    const y = 10 + i * rh, v = r && r.v; const isBn = c.bottleneck && c.bottleneck.stepId === s.id;
    b += `<text x="${lx - 8}" y="${y + 15}" text-anchor="end" font-size="11" fill="${P.ink}">${esc(s.name)}</text>`;
    if (v != null) b += `<rect x="${lx}" y="${y + 4}" width="${((v / max) * cw).toFixed(1)}" height="15" rx="2" fill="${isBn ? P.red : r.s === 'VER' ? P.green : P.accent}"/><text x="${lx + (v / max) * cw + 5}" y="${y + 15}" font-size="10" fill="${P.muted}">${fmt(v, 2)} s</text>`;
    else b += `<text x="${lx + 4}" y="${y + 15}" font-size="10" fill="${P.amber}">Not Yet Collected</text>`;
  });
  if (c.takt.v) { const x = lx + (c.takt.v / max) * cw; b += `<line x1="${x}" y1="4" x2="${x}" y2="${H - 26}" stroke="${P.amber}" stroke-width="2" stroke-dasharray="5 3"/><text x="${x}" y="${H - 10}" text-anchor="middle" font-size="10" fill="${P.amber}">Takt ${fmt(c.takt.v, 2)} s</text>`; }
  return `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px" role="img" aria-label="Effective cycle time by step">${b}</svg>`;
}

/* ---------------- diagram views ---------------- */
function diagramView(kind, title, desc, tools, extra = '') {
  return `<section class="head-row"><div><h1>${title}</h1><p class="lede">${desc}</p></div></section>
  <div class="toolbar">${tools}<span class="sp"></span><button class="btn sm" data-action="zoom" data-z="-1" title="Zoom out" aria-label="Zoom out">−</button><button class="btn sm" data-action="zoom" data-z="0" title="Fit">Fit</button><button class="btn sm" data-action="zoom" data-z="1" title="Zoom in" aria-label="Zoom in">+</button><button class="btn sm" data-action="exportSVG" data-d="${kind}">SVG</button><button class="btn sm" data-action="exportPNG" data-d="${kind}">PNG</button></div>
  <div class="canvas" id="canvas-${kind}" data-kind="${kind}"></div>${extra}`;
}
VIEWS.flow = function () {
  const tools = `<label class="inl">Add <select id="addKind">${opt(Object.entries(FLOW_KINDS).map(([k, v]) => [k, v]), 'operation')}</select></label><button class="btn sm" data-action="flowAdd">Add block</button>
  <button class="btn sm${UI.connectFrom !== null ? ' on' : ''}" data-action="flowConnect" id="btnConnect">${UI.connectFrom !== null ? (UI.connectFrom ? 'Pick target…' : 'Pick source…') : 'Connect'}</button>
  <button class="btn sm" data-action="delSel">Delete selected</button><button class="btn sm" data-action="flowLayout">Auto-layout</button><button class="btn sm" data-action="flowRebuild">Rebuild from steps</button>`;
  const legend = `<div class="legend wrap">${Object.entries(FLOW_KINDS).filter(([k]) => k !== 'end').map(([k, v]) => `<span><i class="sym sym-${k}"></i>${v}</span>`).join('')}<span class="sp"></span>${DEPTS.map((d) => `<span><i class="dot" style="background:${DEPT_COLORS[d]}"></i>${d}</span>`).join('')}</div>`;
  return diagramView('flow', 'Process flow map', 'Every operation from receiving to shipping, with inspections, decisions, rework and scrap. Drag blocks to move them; select a block to edit it and its owner.', tools, legend);
};
VIEWS.vsm = function () {
  return diagramView('vsm', 'Value stream map', 'Material flow (solid), information flow (dashed), data boxes and the lead-time ladder. Select a process box to edit its data. Red outline = bottleneck; star = kaizen opportunity (future state).', `<button class="btn sm" data-tab="data">Edit step data</button>`, vsmTable());
};
function vsmTable() {
  const l = LINE(), c = calcState(l, UI.state), steps = applicable(ST().steps);
  return `<section class="card"><div class="card-h"><h2>Value-added vs non-value-added</h2>${chip(c.leadTime.s)}</div><div class="scrollx"><table class="tbl"><thead><tr><th>Step</th><th>Class</th><th>Queue (min)</th><th>Processing (min)</th><th>Eff. C/T (s)</th></tr></thead><tbody>${steps.map((s) => { const r = c.steps[s.id] || {}; return `<tr><td>${esc(s.name)}</td><td>${esc(VACLASS[s.vaClass])}</td><td class="num">${r.queue && r.queue.v != null ? fmt(r.queue.v, 1) : '<span class="nyc-t">NYC</span>'}</td><td class="num">${val(s.f.processTime) != null ? fmt(val(s.f.processTime), 1) : '<span class="nyc-t">NYC</span>'}</td><td class="num">${r.ect && r.ect.v != null ? fmt(r.ect.v, 2) : '<span class="nyc-t">NYC</span>'}</td></tr>`; }).join('')}
  <tr class="tot"><td>Total lead time</td><td></td><td colspan="2" class="num">${fmt(c.leadTime.v, 1)} min ${c.leadTime.s === 'NYC' ? '(incomplete)' : ''}</td><td></td></tr><tr class="tot"><td>Process cycle efficiency</td><td></td><td colspan="2" class="num">${pct(c.pce.v, 2)}</td><td></td></tr></tbody></table></div></section>`;
}
VIEWS.spaghetti = function () {
  const l = LINE(), S = ST(), c = calcState(l, UI.state);
  const tools = `<label class="inl">Mode <select id="spMode">${opt([['select', 'Select / move zones'], ['draw', 'Draw a route']], UI.spMode)}</select></label>
  ${UI.spMode === 'draw' ? `<label class="inl">Category <select id="drawCat">${opt(Object.entries(ROUTE_CATS).map(([k, v]) => [k, v.label]), UI.drawCat)}</select></label><button class="btn sm primary" data-action="finishRoute">Finish route (${UI.draft.length} pts)</button><button class="btn sm" data-action="cancelRoute">Cancel</button>` : `<button class="btn sm" data-action="addZone">Add zone</button><button class="btn sm" data-action="delSel">Delete selected</button>`}`;
  const routes = (S.routes || []).map((r, i) => { const x = c.routes[r.id]; return `<tr class="${UI.sel === r.id ? 'selrow' : ''}"><td>${i + 1}</td><td><select data-route="${r.id}|cat">${opt(Object.entries(ROUTE_CATS).map(([k, v]) => [k, v.label]), r.cat)}</select></td><td><input data-route="${r.id}|name" value="${esc(r.name)}"></td><td class="num">${fmt(x.dist, 0)}</td><td><input type="number" min="0" step="any" data-route="${r.id}|trips.v" value="${r.trips.v ?? ''}" placeholder="NYC"></td><td><input type="number" min="0" step="any" data-route="${r.id}|people.v" value="${r.people.v ?? ''}" placeholder="NYC"></td><td><select data-route="${r.id}|status">${opt([['EST', 'Estimate'], ['VER', 'Verified']], r.status)}</select></td><td class="num">${x.distShift != null ? fmt(x.distShift, 0) : 'NYC'}</td><td class="num">${x.timeShift != null ? fmt(x.timeShift, 1) : 'NYC'}</td><td class="num">${x.hrsYear.v != null ? fmt(x.hrsYear.v, 0) : 'NYC'}</td><td class="num">${x.costYear.v != null ? money(x.costYear.v) : 'NYC'}</td><td><button class="btn xs" data-action="delRoute" data-id="${r.id}" aria-label="Delete route">✕</button></td></tr>`; }).join('');
  const extra = `<section class="card"><div class="card-h"><h2>Movement routes</h2><span class="muted">Distance from the drawn route × layout scale (1 unit = 1 ft). Annual cost ${c.travelCost.v != null ? '= ' + money(c.travelCost.v) : 'needs labor rate, trips, people, shifts and days'} ${chip(c.travelCost.s)}</span></div>
  <div class="scrollx"><table class="tbl edit"><thead><tr><th>#</th><th>Category</th><th>Route</th><th>ft/trip</th><th>Trips/shift</th><th>People</th><th>Data</th><th>ft/shift</th><th>min/shift</th><th>hr/yr</th><th>$/yr</th><th></th></tr></thead><tbody>${routes || '<tr><td colspan="12" class="muted">No routes yet. Switch Mode to “Draw a route”, pick a category, then tap points on the layout and press Finish.</td></tr>'}</tbody></table></div></section>
  <section class="card"><div class="card-h"><h2>Layout zones</h2><span class="muted">Feet. ${esc(S.layout.note)}</span></div>
  <div class="row-fields"><label>Building width (ft) <input type="number" data-layout="width" value="${S.layout.width}"></label><label>Building depth (ft) <input type="number" data-layout="height" value="${S.layout.height}"></label><label class="grow">Layout note <input data-layout="note" value="${esc(S.layout.note)}"></label></div>
  <div class="scrollx"><table class="tbl edit"><thead><tr><th>Zone</th><th>X</th><th>Y</th><th>Width</th><th>Depth</th><th>Color</th><th></th></tr></thead><tbody>${S.layout.zones.map((z) => `<tr class="${UI.sel === z.id ? 'selrow' : ''}"><td><input data-zone="${z.id}|name" value="${esc(z.name)}"></td>${['x', 'y', 'w', 'h'].map((k) => `<td><input type="number" data-zone="${z.id}|${k}" value="${z[k]}"></td>`).join('')}<td><input type="color" data-zone="${z.id}|color" value="${z.color}"></td><td><button class="btn xs" data-action="delZone" data-id="${z.id}" aria-label="Delete zone">✕</button></td></tr>`).join('')}</tbody></table></div></section>`;
  return diagramView('spaghetti', 'Spaghetti diagram', 'Walking, forklift, ingredient and product movement on a 2D layout. Draw each route once; distance, time and labor cost are calculated from trips, people and speed.', tools, extra);
};
VIEWS.swimlane = function () {
  const tools = `<label class="inl">Lane <select id="swLane">${opt(DEPTS, 'Production')}</select></label><label class="inl">Type <select id="swKind">${opt(Object.entries(SWIM_KINDS), 'task')}</select></label><button class="btn sm" data-action="swAdd">Add activity</button>
  <button class="btn sm${UI.connectFrom !== null ? ' on' : ''}" data-action="swConnect">${UI.connectFrom !== null ? (UI.connectFrom ? 'Pick target…' : 'Pick source…') : 'Connect'}</button><button class="btn sm" data-action="delSel">Delete selected</button>`;
  const items = ST().swim || [];
  const flagged = items.filter((i) => Object.values(i.flags || {}).some(Boolean));
  const extra = `<section class="card"><div class="card-h"><h2>Flagged handoffs, approvals and delays</h2><span class="muted">Select an activity to flag duplicated work, unnecessary approvals, delays or communication gaps.</span></div>
  <div class="scrollx"><table class="tbl"><thead><tr><th>Lane</th><th>Activity</th><th>Flags</th></tr></thead><tbody>${flagged.map((i) => `<tr><td>${esc(i.lane)}</td><td>${esc(i.label)}</td><td>${Object.entries(i.flags).filter((x) => x[1]).map((x) => `<span class="chip bad">${SWIM_FLAGS[x[0]]}</span>`).join(' ')}</td></tr>`).join('') || '<tr><td colspan="3" class="muted">Nothing flagged yet. The starting sequence must be confirmed with each department.</td></tr>'}</tbody></table></div></section>`;
  return diagramView('swimlane', 'Swimlane process map', 'Who does what across Production, Quality, Maintenance, Sanitation, Warehouse, Planning and Supervisors. Dashed blue arrows are cross-department handoffs. The starting sequence is a template to verify.', tools, extra);
};

/* ---------------- PAM ---------------- */
VIEWS.pam = function () {
  const l = LINE(), S = ST(), c = calcState(l, UI.state), p = c.pam;
  const rows = S.acts.map((a) => `<tr><td><select data-act="${a.id}|stepId">${opt(S.steps.map((s) => [s.id, s.name]), a.stepId)}</select></td><td><input data-act="${a.id}|desc" value="${esc(a.desc)}"></td><td><select data-act="${a.id}|type" class="type-${a.type}">${opt(Object.entries(ACT_TYPES).map(([k, v]) => [k, k + ' · ' + v]), a.type)}</select></td>
    <td><input type="number" step="any" min="0" data-act="${a.id}|time" value="${a.time ?? ''}" placeholder="NYC"></td><td><input type="number" step="any" min="0" data-act="${a.id}|dist" value="${a.dist ?? ''}" placeholder="—"></td><td><input type="number" step="any" min="0" data-act="${a.id}|people" value="${a.people ?? ''}" placeholder="—"></td>
    <td><select data-act="${a.id}|va">${opt([['VA', 'VA'], ['NNVA', 'NNVA'], ['NVA', 'NVA']], a.va)}</select></td><td><select data-act="${a.id}|status">${opt([['NYC', NYC], ['EST', 'Estimate'], ['VER', 'Verified']], a.time == null || a.time === '' ? 'NYC' : a.status)}</select></td><td><input data-act="${a.id}|opportunity" value="${esc(a.opportunity)}" placeholder="Improvement idea"></td><td><button class="btn xs" data-action="delAct" data-id="${a.id}" aria-label="Delete activity">✕</button></td></tr>`).join('');
  const max = Math.max(1, ...Object.values(p.byType).map((b) => b.time));
  const bars = Object.entries(ACT_TYPES).map(([k, v]) => `<div class="hbar"><span class="hb-l"><i class="actype type-${k}">${k}</i>${v}</span><span class="hb-t"><span style="width:${((p.byType[k].time / max) * 100).toFixed(1)}%" class="type-bg-${k}"></span></span><span class="hb-v">${fmt(p.byType[k].time, 1)} min · ${p.byType[k].n}</span></div>`).join('');
  return `<section class="head-row"><div><h1>Detailed process activity map</h1><p class="lede">Every activity classified as Operation, Transportation, Inspection, Delay or Storage, with time, distance and people. Blank time = Not Yet Collected.</p></div><div class="head-actions"><button class="btn" data-action="addAct">Add activity</button></div></section>
  <section class="grid2"><div class="card"><div class="card-h"><h2>Time by activity type</h2>${chip(p.s, p.missing ? `${p.missing} activities not timed` : undefined)}</div>${bars}</div>
  <div class="card"><div class="mini-stats"><div><b>${fmt(p.time, 1)}</b><span>total min</span></div><div><b>${fmt(p.nva, 1)}</b><span>non-value-added min</span></div><div><b>${p.time ? pct(p.va / p.time, 1) : '—'}</b><span>value-added share</span></div><div><b>${fmt(p.dist, 0)}</b><span>ft travelled</span></div><div><b>${fmt(p.peopleMin, 0)}</b><span>people-minutes</span></div></div></div></section>
  <section class="card"><div class="scrollx"><table class="tbl edit"><thead><tr><th>Step</th><th>Activity</th><th>Type</th><th>Time (min)</th><th>Distance (ft)</th><th>People</th><th>Class</th><th>Data</th><th>Improvement opportunity</th><th></th></tr></thead><tbody>${rows}</tbody></table></div></section>`;
};

/* ---------------- step data ---------------- */
VIEWS.data = function () {
  const l = LINE(), S = ST();
  const settings = SETTINGS.map((sd) => { const f = l.settings[sd.k]; return `<tr><td>${esc(sd.label)}<div class="muted sm">${esc(sd.how)}</div></td><td class="unit">${esc(sd.unit.replace('units', l.unit + 's'))}</td><td><input type="number" step="any" data-set="${sd.k}|v" value="${f.v ?? ''}" placeholder="${NYC}"></td><td>${statusSelect('data-set', sd.k, f)}</td><td><input data-set="${sd.k}|src" value="${esc(f.src)}" placeholder="Source: who / when / how"></td></tr>`; }).join('');
  const steps = S.steps.map((s, i) => `<tr class="${s.applic === 'Not applicable' ? 'na' : ''}${UI.sel === s.id ? ' selrow' : ''}"><td class="ord"><button class="btn xs" data-action="moveStep" data-id="${s.id}" data-d="-1" aria-label="Move up">↑</button><button class="btn xs" data-action="moveStep" data-id="${s.id}" data-d="1" aria-label="Move down">↓</button></td>
    <td><input data-step="${s.id}|name" value="${esc(s.name)}"></td><td><select data-step="${s.id}|dept">${opt(DEPTS, s.dept)}</select></td><td><input data-step="${s.id}|owner" value="${esc(s.owner)}" placeholder="${NYC}"></td><td><input data-step="${s.id}|equipment" value="${esc(s.equipment)}" placeholder="${NYC}"></td>
    <td><select data-step="${s.id}|applic">${opt(APPLIC, s.applic)}</select></td><td><select data-step="${s.id}|vaClass">${opt(Object.entries(VACLASS), s.vaClass)}</select></td><td><select data-step="${s.id}|kind">${opt([['operation', 'Operation'], ['storage', 'Storage'], ['inspection', 'Inspection']], s.kind)}</select></td>
    <td class="num">${FIELDS.filter((fd) => fstat(s.f[fd.k]) !== 'NYC').length}/${FIELDS.length}</td><td class="nowrap"><button class="btn xs primary" data-action="selStep" data-id="${s.id}">Data</button><button class="btn xs" data-action="addStep" data-id="${s.id}" title="Insert step after" aria-label="Insert step after">＋</button><button class="btn xs" data-action="delStep" data-id="${s.id}" aria-label="Delete step">✕</button></td></tr>`).join('');
  const matrix = `<div class="scrollx"><table class="tbl matrix"><thead><tr><th>Step</th>${FIELDS.map((fd) => `<th title="${esc(fd.how)}">${esc(fd.label)}<div class="muted sm">${esc(fd.unit.replace('units', l.unit + 's'))}</div></th>`).join('')}</tr></thead><tbody>${applicable(S.steps).map((s) => `<tr><td><button class="link" data-action="selStep" data-id="${s.id}">${esc(s.name)}</button></td>${FIELDS.map((fd) => { const f = s.f[fd.k], st = fstat(f); return `<td class="m-${st}" data-action="selStep" data-id="${s.id}">${st === 'NYC' ? 'NYC' : fmt(+f.v, 2)}</td>`; }).join('')}</tr>`).join('')}</tbody></table></div>`;
  return `<section class="head-row"><div><h1>Step data</h1><p class="lede">Add, remove, rename and reorder steps for ${esc(l.name)}. Every value carries a status and a source. Blank = Not Yet Collected; nothing is filled in by the system.</p></div>
  <div class="head-actions"><label class="inl">Unit of measure <input id="unitIn" data-line="unit" value="${esc(l.unit)}" size="8"></label></div></section>
  <section class="card"><div class="card-h"><h2>Line settings &amp; assumptions</h2><span class="muted">Used by takt, lead time, savings and travel calculations</span></div><div class="scrollx"><table class="tbl edit"><thead><tr><th>Setting</th><th>Unit</th><th>Value</th><th>Status</th><th>Source</th></tr></thead><tbody>${settings}</tbody></table></div></section>
  <section class="card"><div class="card-h"><h2>Process steps</h2><div class="seg"><button class="btn sm${UI.matrix ? '' : ' on'}" data-action="matrix" data-v="0">Steps</button><button class="btn sm${UI.matrix ? ' on' : ''}" data-action="matrix" data-v="1">Data matrix</button></div></div>
  ${UI.matrix ? matrix : `<div class="scrollx"><table class="tbl edit"><thead><tr><th></th><th>Step</th><th>Department</th><th>Process owner</th><th>Equipment</th><th>Applies to line?</th><th>Class</th><th>Type</th><th>Data</th><th></th></tr></thead><tbody>${steps}</tbody></table></div><button class="btn" data-action="addStep" data-id="">Add step at end</button>`}</section>`;
};
function statusSelect(attr, key, f) { return `<select ${attr}="${key}|s"${val(f) == null ? ' disabled title="Enter a value first"' : ''}>${opt(val(f) == null ? [['NYC', NYC]] : [['EST', 'Estimate'], ['VER', 'Verified']], fstat(f))}</select>`; }

/* ---------------- opportunities ---------------- */
function suggestions(l) {
  const out = [], c = calcState(l, 'current'), S = l.current;
  if (c.bottleneck) out.push({ stepId: c.bottleneck.stepId, waste: 'Waiting', desc: `${c.bottleneck.name} has the highest effective cycle time (${fmt(c.bottleneck.ect.v, 2)} s)${c.takt.v ? (c.bottleneck.ect.v > c.takt.v ? ', above takt — demand at risk' : ', below takt') : ''}. Investigate downtime and changeover.` });
  let hq = null; applicable(S.steps).forEach((s) => { const q = c.steps[s.id] && c.steps[s.id].queue; if (q && q.v != null && (!hq || q.v > hq.v)) hq = { s, v: q.v }; });
  if (hq) out.push({ stepId: hq.s.id, waste: 'Inventory', desc: `Largest queue is before ${hq.s.name} (${fmt(hq.v, 0)} min of WIP). Check batch sizes and pull signals.` });
  let hs = null; applicable(S.steps).forEach((s) => { const v = val(s.f.scrap); if (v != null && (!hs || v > hs.v)) hs = { s, v }; });
  if (hs) out.push({ stepId: hs.s.id, waste: 'Defects', desc: `Highest recorded scrap is at ${hs.s.name} (${fmt(hs.v, 1)}%). Pareto the scrap causes.` });
  let hc = null; applicable(S.steps).forEach((s) => { const v = c.steps[s.id] && c.steps[s.id].coH; if (v != null && (!hc || v > hc.v)) hc = { s, v }; });
  if (hc) out.push({ stepId: hc.s.id, waste: 'Waiting', desc: `${hc.s.name} spends ${fmt(hc.v, 1)} hr/week in changeover. Candidate for a SMED project.` });
  let lr = null; (S.routes || []).forEach((r) => { const x = c.routes[r.id]; if (x.distShift != null && (!lr || x.distShift > lr.v)) lr = { r, v: x.distShift }; });
  if (lr) out.push({ stepId: '', waste: lr.r.cat === 'walk' ? 'Motion' : 'Transportation', desc: `Longest movement: “${lr.r.name}” (${fmt(lr.v, 0)} ft/shift). Consider point-of-use storage or layout change.` });
  (S.swim || []).forEach((i) => Object.entries(i.flags || {}).forEach(([k, on]) => { if (on) out.push({ stepId: '', waste: k === 'duplicate' || k === 'unnecessary' ? 'Over-processing' : k === 'delay' ? 'Waiting' : 'Skills', desc: `${i.lane}: “${i.label}” flagged as ${SWIM_FLAGS[k].toLowerCase()}.` }); }));
  return out.filter((sg) => !PRJ.opps.some((o) => o.lineId === l.id && o.desc === sg.desc));
}
VIEWS.opps = function () {
  const l = LINE(), opps = PRJ.opps.filter((o) => o.lineId === l.id), sug = suggestions(l);
  const rows = opps.map((o) => `<tr><td><select data-opp="${o.id}|stepId">${opt([['', '— Line-wide —']].concat(l.current.steps.map((s) => [s.id, s.name])), o.stepId)}</select></td><td><select data-opp="${o.id}|waste">${opt(WASTES.map((w) => w.k), o.waste)}</select></td><td><textarea rows="2" data-opp="${o.id}|desc">${esc(o.desc)}</textarea></td><td><select data-opp="${o.id}|priority">${opt(['High', 'Medium', 'Low'], o.priority)}</select></td><td><input data-opp="${o.id}|owner" value="${esc(o.owner)}" placeholder="Owner"></td>
    <td><input type="number" step="any" data-opp="${o.id}|estAnnual" value="${o.estAnnual ?? ''}" placeholder="Not estimated"></td><td><input data-opp="${o.id}|estBasis" value="${esc(o.estBasis)}" placeholder="How estimated"></td><td>${finChip(o.fin.status)}</td><td class="num">${PRJ.actions.filter((a) => a.oppId === o.id).length}</td><td><button class="btn xs" data-action="delOpp" data-id="${o.id}" aria-label="Delete opportunity">✕</button></td></tr>`).join('');
  const counts = WASTES.map((w) => ({ w, n: opps.filter((o) => o.waste === w.k).length }));
  return `<section class="head-row"><div><h1>Waste &amp; opportunity register</h1><p class="lede">Classify each opportunity by the eight wastes. Estimated savings stay estimates until Finance validates them.</p></div><div class="head-actions"><button class="btn primary" data-action="addOpp">Add opportunity</button></div></section>
  <section class="wastes">${counts.map(({ w, n }) => `<div class="waste"><b>${n}</b><span>${w.k}</span><small>${w.d}</small></div>`).join('')}</section>
  <section class="card"><div class="scrollx"><table class="tbl edit"><thead><tr><th>Step</th><th>Waste</th><th>Opportunity</th><th>Priority</th><th>Owner</th><th>Est. $/yr</th><th>Estimate basis</th><th>Finance</th><th>Actions</th><th></th></tr></thead><tbody>${rows || '<tr><td colspan="10" class="muted">No opportunities recorded yet.</td></tr>'}</tbody></table></div></section>
  <section class="card"><div class="card-h"><h2>Suggested from your data</h2><span class="muted">Generated from collected values only. Review before adding.</span></div>${sug.length ? `<ul class="sugg">${sug.map((s, i) => `<li><span class="chip est">${esc(s.waste)}</span> ${esc(s.desc)} <button class="btn xs" data-action="addSug" data-i="${i}">Add to register</button></li>`).join('')}</ul>` : '<p class="muted">No suggestions yet. They appear once cycle times, WIP, scrap, changeovers, routes or swimlane flags are entered.</p>'}</section>`;
};

/* ---------------- actions ---------------- */
VIEWS.actions = function () {
  const l = LINE(), opps = PRJ.opps.filter((o) => o.lineId === l.id);
  let acts = PRJ.actions.filter((a) => opps.some((o) => o.id === a.oppId));
  if (UI.actFilter !== 'All') acts = acts.filter((a) => a.status === UI.actFilter);
  const rows = acts.map((a) => { const od = a.status !== 'Done' && a.due && a.due < todayISO(); return `<tr class="${od ? 'overdue' : ''}"><td><select data-actrow="${a.id}|oppId">${opt(opps.map((o) => [o.id, o.desc.slice(0, 60)]), a.oppId)}</select></td><td><textarea rows="2" data-actrow="${a.id}|action">${esc(a.action)}</textarea></td><td><input data-actrow="${a.id}|owner" value="${esc(a.owner)}" placeholder="Owner"></td><td><input type="date" data-actrow="${a.id}|due" value="${esc(a.due)}">${od ? '<span class="chip bad">Overdue</span>' : ''}</td><td><select data-actrow="${a.id}|status">${opt(ACTION_STATUS, a.status)}</select></td><td><input data-actrow="${a.id}|notes" value="${esc(a.notes)}"></td><td><button class="btn xs" data-action="delAction" data-id="${a.id}" aria-label="Delete action">✕</button></td></tr>`; }).join('');
  return `<section class="head-row"><div><h1>Improvement action tracker</h1><p class="lede">Every action has an owner, a due date and a status, linked to an opportunity.</p></div><div class="head-actions"><label class="inl">Show <select id="actFilter">${opt(['All'].concat(ACTION_STATUS), UI.actFilter)}</select></label><button class="btn primary" data-action="addAction"${opps.length ? '' : ' disabled title="Add an opportunity first"'}>Add action</button></div></section>
  <section class="card"><div class="scrollx"><table class="tbl edit"><thead><tr><th>Opportunity</th><th>Action</th><th>Owner</th><th>Due</th><th>Status</th><th>Notes</th><th></th></tr></thead><tbody>${rows || '<tr><td colspan="7" class="muted">No actions yet.</td></tr>'}</tbody></table></div></section>`;
};

/* ---------------- finance ---------------- */
VIEWS.finance = function () {
  const l = LINE(), opps = PRJ.opps.filter((o) => o.lineId === l.id);
  const by = (s) => opps.filter((o) => o.fin.status === s);
  const est = opps.filter((o) => !['Validated by Finance', 'Rejected'].includes(o.fin.status)).reduce((a, o) => a + (+o.estAnnual || 0), 0);
  const valid = by('Validated by Finance').reduce((a, o) => a + (+o.fin.validated || 0), 0);
  const hard = by('Validated by Finance').filter((o) => o.fin.category === 'Hard savings').reduce((a, o) => a + (+o.fin.validated || 0), 0);
  const rows = opps.map((o) => `<tr><td>${esc(o.desc)}<div class="muted sm">${esc(stepName(o.stepId))} · ${esc(o.waste)}</div></td><td><select data-opp="${o.id}|fin.category">${opt(FIN_CATS, o.fin.category)}</select></td><td><input data-opp="${o.id}|fin.baseline" value="${esc(o.fin.baseline)}" placeholder="Baseline & data source"></td><td><input data-opp="${o.id}|fin.method" value="${esc(o.fin.method)}" placeholder="Valuation method"></td><td class="num">${o.estAnnual != null && o.estAnnual !== '' ? money(+o.estAnnual) : '—'}</td><td><select data-opp="${o.id}|fin.status">${opt(FIN_STATUS, o.fin.status)}</select></td><td><input type="number" step="any" data-opp="${o.id}|fin.validated" value="${o.fin.validated ?? ''}" placeholder="—"${o.fin.status === 'Validated by Finance' ? '' : ' disabled title="Set status to Validated by Finance first"'}></td><td><input data-opp="${o.id}|fin.reviewer" value="${esc(o.fin.reviewer)}" placeholder="Finance reviewer"></td><td><input type="date" data-opp="${o.id}|fin.date" value="${esc(o.fin.date)}"></td></tr>`).join('');
  return `<section class="head-row"><div><h1>Financial savings validation</h1><p class="lede">Savings are reported only after Finance validates them against actual cost results. Agree the baseline and method at the start; value labor only when paid hours fall; value scrap at cost, not selling price.</p></div></section>
  <section class="mini-stats card"><div><b>${money(est)}</b><span>estimated, not yet validated</span></div><div><b>${money(valid)}</b><span>validated by Finance (all categories)</span></div><div><b>${money(hard)}</b><span>validated hard savings</span></div><div><b>${by('In verification').length + by('Method agreed with Finance').length}</b><span>in validation</span></div></section>
  <section class="card"><div class="scrollx"><table class="tbl edit"><thead><tr><th>Opportunity</th><th>Benefit type</th><th>Baseline</th><th>Method</th><th>Estimated $/yr</th><th>Status</th><th>Validated $/yr</th><th>Reviewer</th><th>Date</th></tr></thead><tbody>${rows || '<tr><td colspan="9" class="muted">Opportunities appear here once added to the register.</td></tr>'}</tbody></table></div></section>`;
};

/* ---------------- compare ---------------- */
VIEWS.compare = function () {
  const l = LINE();
  if (!l.future) return noFuture();
  const x = calcCompare(l), c = x.c, f = x.f;
  const row = (label, a, b, fm, better = 'lower') => { const d = a.v != null && b.v != null ? b.v - a.v : null; const good = d == null ? '' : (better === 'lower' ? d < 0 : d > 0) ? 'good' : d === 0 ? '' : 'bad'; return `<tr><td>${label}</td><td class="num">${a.v != null ? fm(a.v) : 'NYC'}</td><td class="num">${b.v != null ? fm(b.v) : 'NYC'}</td><td class="num ${good}">${d != null ? (d > 0 ? '+' : '') + fm(d) : '—'}</td><td>${chip(worst([a.s, b.s]))}</td></tr>`; };
  return `<section class="head-row"><div><h1>Current vs future state — ${esc(l.name)}</h1><p class="lede">Side-by-side KPIs and estimated savings. Every saving here is an estimate until Finance validates it.</p></div><div class="head-actions"><button class="btn" data-action="exportCompare">Comparison report (PDF)</button></div></section>
  <section class="card"><div class="scrollx"><table class="tbl"><thead><tr><th>KPI</th><th>Current</th><th>Future</th><th>Change</th><th>Data</th></tr></thead><tbody>
  ${row('Total lead time (hr)', c.leadTime, f.leadTime, (v) => fmt(v / 60, 2))}${row('Value-added time (min)', c.vaTime, f.vaTime, (v) => fmt(v, 1), 'higher')}${row('Process cycle efficiency', c.pce, f.pce, (v) => pct(v, 2), 'higher')}
  ${row('Bottleneck effective C/T (s)', c.bnECT, f.bnECT, (v) => fmt(v, 2))}${row('Line capacity (' + l.unit + '/hr)', c.capacity, f.capacity, (v) => fmt(v, 0), 'higher')}${row('Rolled throughput yield', c.rty, f.rty, (v) => pct(v, 2), 'higher')}
  ${row('Changeover (hr/week)', c.coHours, f.coHours, (v) => fmt(v, 1))}${row('Total labor (people)', c.labor, f.labor, (v) => fmt(v, 1))}${row('Travel labor (hr/yr)', c.travelHrs, f.travelHrs, (v) => fmt(v, 0))}${row('Travel labor cost ($/yr)', c.travelCost, f.travelCost, money)}
  </tbody></table></div></section>
  <section class="kpis">${kpiTile('Estimated travel labor savings', x.travelSavings, money, '/yr', 'Becomes hard savings only if paid hours come out')}${kpiTile('Changeover hours freed', x.coHrsYr, (v) => fmt(v, 0), 'hr/yr')}${kpiTile('Changeover labor value', x.coValue, money, '/yr', 'Capacity unless Finance confirms')}${kpiTile('Scrap reduction (at cost)', x.scrapSavings, money, '/yr')}${kpiTile('Throughput potential', x.throughputGain, (v) => pct(v, 1), '')}</section>
  <section class="grid2"><div class="card"><div class="card-h"><h2>Current state layout</h2></div><div class="scrollx mini">${svgSpaghetti(l, 'current', PAL(), { demo: PRJ.demo }).svg}</div></div><div class="card"><div class="card-h"><h2>Future state layout</h2></div><div class="scrollx mini">${svgSpaghetti(l, 'future', PAL(), { demo: PRJ.demo }).svg}</div></div></section>`;
};

/* ---------------- checklist ---------------- */
function checklistItems(l) {
  const items = [], S = l.current;
  SETTINGS.forEach((sd) => { if (fstat(l.settings[sd.k]) === 'NYC') items.push({ k: `${l.id}:set:${sd.k}`, area: 'Line settings (Planning / Finance)', item: sd.label, how: sd.how }); });
  S.steps.forEach((s) => {
    if (s.applic === 'To verify') items.push({ k: `${l.id}:app:${s.id}`, area: s.name, item: 'Confirm whether this step exists on the line, and its equipment', how: 'Walk the line with the supervisor; record equipment name and asset number.' });
    if (s.applic === 'Not applicable') return;
    if (!s.owner) items.push({ k: `${l.id}:own:${s.id}`, area: s.name, item: 'Process owner', how: 'Confirm with the area supervisor.' });
    FIELDS.forEach((fd) => { if (fstat(s.f[fd.k]) === 'NYC' && !(fd.k === 'oee' && s.kind === 'storage')) items.push({ k: `${l.id}:f:${s.id}:${fd.k}`, area: s.name, item: `${fd.label} (${fd.unit.replace('units', l.unit + 's')})`, how: fd.how }); });
    const fs = s.fsq; const fsx = [fs.foodSafety && 'food-safety controls', fs.allergen && 'allergen controls', fs.trace && 'traceability records', fs.qc && 'quality checks', fs.sanitation && 'sanitation requirements'].filter(Boolean);
    if (fsx.length) items.push({ k: `${l.id}:fsq:${s.id}`, area: s.name, item: 'Verify ' + fsx.join(', '), how: 'Compare the template notes against the food safety plan, allergen matrix and QC records; correct the step notes.' });
  });
  items.push({ k: `${l.id}:sp:layout`, area: 'Spaghetti diagram', item: 'Replace placeholder layout with plant drawing dimensions', how: 'Use the facility drawing; measure zone positions in feet.' });
  items.push({ k: `${l.id}:sp:routes`, area: 'Spaghetti diagram', item: 'Trace walking, forklift, ingredient and product routes for one full shift', how: 'Follow each role with a layout printout; count trips; measure routes with a measuring wheel.' });
  items.push({ k: `${l.id}:sw:verify`, area: 'Swimlane', item: 'Confirm the handoff and approval sequence with each department', how: 'Short interviews with Planning, Warehouse, Sanitation, Quality, Maintenance and the supervisors; mark duplicated work and waits.' });
  items.push({ k: `${l.id}:al:changeover`, area: 'Food safety', item: 'Observe one allergen changeover and the sanitation verification', how: 'Record steps, times and sign-offs; compare with the allergen changeover procedure.' });
  return items;
}
VIEWS.checklist = function () {
  const l = LINE(), items = checklistItems(l); const done = items.filter((i) => PRJ.checks[i.k]).length;
  const groups = {}; items.forEach((i) => (groups[i.area] = groups[i.area] || []).push(i));
  return `<section class="head-row"><div><h1>Floor-walk data collection checklist</h1><p class="lede">Generated from what is still Not Yet Collected for ${esc(l.name)} (current state). Tick items as they are measured, then enter the values in Step Data with their status and source.</p></div><div class="head-actions"><span class="muted">${done} of ${items.length} done</span></div></section>
  <div class="bar"><span class="bar-ver" style="width:${items.length ? ((done / items.length) * 100).toFixed(1) : 0}%"></span></div>
  <section class="card"><p class="muted">Bring: stopwatch or phone timer, measuring wheel, layout printout, clipboard, scale access, the food safety plan and allergen matrix. Time each cycle-time step at least 10 consecutive cycles in 3 periods. Record who measured, when and how in the Source field.</p></section>
  ${Object.entries(groups).map(([g, list]) => `<section class="card"><div class="card-h"><h2>${esc(g)}</h2><span class="muted">${list.filter((i) => PRJ.checks[i.k]).length}/${list.length}</span></div><ul class="checks">${list.map((i) => `<li><label><input type="checkbox" data-check="${esc(i.k)}"${PRJ.checks[i.k] ? ' checked' : ''}> <span><b>${esc(i.item)}</b><br><span class="muted">${esc(i.how)}</span></span></label></li>`).join('')}</ul></section>`).join('')}`;
};

/* ---------------- formulas ---------------- */
VIEWS.formulas = function () {
  const l = LINE(), c = calcState(l, UI.state);
  const trace = [['Available time per day', c.avail, 'min'], ['Takt time', c.takt, 's/unit'], ['Total lead time', c.leadTime, 'min'], ['Value-added time', c.vaTime, 'min'], ['Process cycle efficiency', c.pce, '%'], ['Bottleneck effective C/T', c.bnECT, 's'], ['Line capacity', c.capacity, l.unit + '/hr'], ['Rolled throughput yield', c.rty, '%'], ['Changeover time', c.coHours, 'hr/week'], ['Travel labor', c.travelHrs, 'hr/yr'], ['Travel labor cost', c.travelCost, '$/yr']];
  const assumptions = SETTINGS.filter((sd) => fstat(l.settings[sd.k]) === 'EST').map((sd) => `<li><b>${esc(sd.label)}</b>: ${fmt(val(l.settings[sd.k]), 2)} ${esc(sd.unit)} — ${esc(l.settings[sd.k].src || 'no source recorded')}</li>`).join('');
  return `<section class="head-row"><div><h1>Formulas, assumptions &amp; data sources</h1><p class="lede">How every number is calculated. A result is “Verified inputs” only when every input is verified; “Estimate” when any input is an estimate; “Incomplete” when an input is Not Yet Collected.</p></div></section>
  <section class="card"><div class="card-h"><h2>Calculation trace — ${esc(l.name)}, ${UI.state} state</h2></div><div class="scrollx"><table class="tbl"><thead><tr><th>Result</th><th>Value</th><th>Data</th><th>Missing inputs</th><th>Formula</th></tr></thead><tbody>${trace.map(([n, r, u]) => `<tr><td>${n}</td><td class="num">${r.v != null ? (u === '%' ? pct(r.v, 2) : fmt(r.v, 2) + ' ' + u) : 'NYC'}</td><td>${chip(r.s)}</td><td class="sm">${esc((r.missing || []).slice(0, 6).join('; '))}${(r.missing || []).length > 6 ? ' …' : ''}</td><td class="sm">${esc(r.formula)}</td></tr>`).join('')}</tbody></table></div></section>
  <section class="card"><div class="card-h"><h2>Formula reference</h2></div><div class="scrollx"><table class="tbl"><thead><tr><th>Calculation</th><th>Formula</th><th>Unit</th></tr></thead><tbody>${FORMULAS.map((f) => `<tr><td>${f[0]}</td><td>${esc(f[1])}</td><td>${esc(f[2])}</td></tr>`).join('')}</tbody></table></div></section>
  <section class="card"><div class="card-h"><h2>Assumptions in use</h2></div>${assumptions ? `<ul>${assumptions}</ul>` : '<p class="muted">No estimated settings.</p>'}<p class="muted">Data status: <b>Verified</b> = measured on the floor or taken from a system of record, with source recorded. <b>Estimate</b> = assumption, planning value or engineering estimate. <b>Not Yet Collected</b> = blank; never filled in by the system.</p></section>`;
};

/* ---------------- lines & export ---------------- */
VIEWS.io = function () {
  return `<section class="head-row"><div><h1>Lines &amp; export</h1><p class="lede">Excel is the primary data format: export the workbook, collect data in it, and import it back. Your working copy also saves automatically ${Store.mode === 'db' ? 'to your account' : Store.mode === 'local' ? 'in this browser' : '(autosave unavailable here — export to Excel to keep your work)'}.</p></div></section>
  <section class="grid2"><div class="card"><div class="card-h"><h2>Export</h2></div><div class="btn-col">
    <button class="btn primary" data-action="exportXLSX" id="btnXlsx">Excel workbook (.xlsx) — all lines, data, calculations, actions</button>
    <button class="btn" data-action="exportPPTX" id="btnPptx">Editable PowerPoint process maps (.pptx)</button>
    <button class="btn" data-action="exportPDF" id="btnPdf">PDF report — ${esc(LINE().name)}</button>
    <button class="btn" data-action="exportCompare">Current vs future comparison report (PDF)</button>
    <button class="btn" data-action="exportJSON">Backup (.json)</button></div><p class="muted sm">SVG and PNG exports are on each diagram's toolbar.</p></div>
  <div class="card"><div class="card-h"><h2>Import</h2></div><label class="file">Excel workbook (.xlsx) <input type="file" accept=".xlsx" id="impXlsx"></label><label class="file">Backup (.json) <input type="file" accept=".json,application/json" id="impJson"></label><p class="muted sm">Importing replaces the data in this app with the workbook's contents.</p>
  <div class="card-h"><h2>Demonstration data</h2></div><p class="muted sm">Fills the Nugget Line with illustrative numbers so you can see every calculation work. Clearly marked DEMO everywhere, including exports.</p>
  ${PRJ.demo ? `<button class="btn" data-action="clearDemo">Remove demonstration data (reset to blank templates)</button>` : `<button class="btn" data-action="loadDemo">Load demonstration data</button>`}</div></section>
  <section class="card"><div class="card-h"><h2>Production lines</h2><span class="muted">Every line has its own steps, equipment, layout and data</span></div>
  <div class="scrollx"><table class="tbl edit"><thead><tr><th>Line</th><th>Steps</th><th>Future state</th><th>Notes</th><th></th></tr></thead><tbody>${PRJ.lines.map((l) => `<tr><td><input data-lineid="${l.id}|name" value="${esc(l.name)}"></td><td class="num">${applicable(l.current.steps).length}</td><td>${l.future ? 'Yes' : 'No'}</td><td><input data-lineid="${l.id}|notes" value="${esc(l.notes)}"></td><td class="nowrap"><button class="btn xs" data-action="dupLine" data-id="${l.id}">Duplicate</button>${PRJ.lines.length > 1 ? `<button class="btn xs" data-action="delLine" data-id="${l.id}">Delete</button>` : ''}</td></tr>`).join('')}</tbody></table></div>
  <div class="row-fields"><label class="grow">New line name <input id="newLineName" placeholder="e.g. Biscuit Line #1"></label><button class="btn" data-action="addLine">Add line from plant template</button></div></section>
  ${UI.confirm ? '' : ''}
  <section class="card"><p class="muted sm">C.H. Guenther & Son — Duncanville. Branding: no official logo has been supplied, so the company name is shown as text. Version ${APP_VERSION}.</p></section>`;
};

/* ---------------- side panel ---------------- */
function renderPanel() {
  const p = $('#panel'); const id = UI.sel, t = UI.selType;
  let h = '';
  const S = ST();
  if (t === 'step') { const s = S.steps.find((x) => x.id === id) || LINE().current.steps.find((x) => x.id === id); if (s) h = stepPanel(s); }
  else if (t === 'node') { const n = S.flow.nodes.find((x) => x.id === id); if (n) h = nodePanel(n); }
  else if (t === 'edge') { const e = S.flow.edges.find((x) => x.id === id); if (e) h = `<h3>Connector</h3><label>Label (e.g. Yes / No) <input data-edge="${e.id}|label" value="${esc(e.label)}"></label><p class="muted sm">${esc(S.flow.nodes.find((n) => n.id === e.from)?.label)} → ${esc(S.flow.nodes.find((n) => n.id === e.to)?.label)}</p><button class="btn" data-action="delSel">Delete connector</button>`; }
  else if (t === 'swim') { const it = S.swim.find((x) => x.id === id); if (it) h = swimPanel(it); }
  else if (t === 'zone') { const z = S.layout.zones.find((x) => x.id === id); if (z) h = `<h3>Zone</h3><label>Name <input data-zone="${z.id}|name" value="${esc(z.name)}"></label><div class="row-fields">${['x', 'y', 'w', 'h'].map((k) => `<label>${{ x: 'X (ft)', y: 'Y (ft)', w: 'Width (ft)', h: 'Depth (ft)' }[k]} <input type="number" data-zone="${z.id}|${k}" value="${z[k]}"></label>`).join('')}</div><button class="btn" data-action="delSel">Delete zone</button>`; }
  else if (t === 'route') { const r = S.routes.find((x) => x.id === id); if (r) { const c = calcState(LINE(), UI.state).routes[r.id]; h = `<h3>Route</h3><label>Name <input data-route="${r.id}|name" value="${esc(r.name)}"></label><label>Category <select data-route="${r.id}|cat">${opt(Object.entries(ROUTE_CATS).map(([k, v]) => [k, v.label]), r.cat)}</select></label><p>Distance: <b>${fmt(c.dist, 0)} ft</b> per trip</p><label>Trips per shift <input type="number" data-route="${r.id}|trips.v" value="${r.trips.v ?? ''}" placeholder="${NYC}"></label><label>People <input type="number" data-route="${r.id}|people.v" value="${r.people.v ?? ''}" placeholder="${NYC}"></label><button class="btn" data-action="delSel">Delete route</button>`; } }
  p.hidden = !h; p.innerHTML = h ? `<button class="btn xs close" data-action="closePanel" aria-label="Close panel">✕ Close</button>${h}` : '';
  document.body.classList.toggle('panel-open', !!h);
}
function stepPanel(s) {
  const l = LINE();
  return `<h3>${esc(s.name)}</h3><p class="muted sm">${esc(s.dept)} · ${esc(s.applic)} · ${UI.state} state</p>
  <div class="pfields">${FIELDS.map((fd) => { const f = s.f[fd.k]; return `<div class="pf"><label for="pf-${fd.k}">${esc(fd.label)} <span class="muted">(${esc(fd.unit.replace('units', l.unit + 's'))})</span></label><div class="pf-row"><input id="pf-${fd.k}" type="number" step="any" data-field="${s.id}|${fd.k}|v" value="${f.v ?? ''}" placeholder="${NYC}">${statusSelect('data-field', s.id + '|' + fd.k, f)}</div><input class="src" data-field="${s.id}|${fd.k}|src" value="${esc(f.src)}" placeholder="Source: who / when / how"><details><summary>How to measure</summary><p class="sm">${esc(fd.how)}</p></details></div>`; }).join('')}</div>
  <h4>Food safety &amp; quality</h4>${[['foodSafety', 'Food safety / CCP'], ['allergen', 'Allergen management'], ['sanitation', 'Sanitation'], ['trace', 'Traceability'], ['qc', 'Quality control points']].map(([k, lb]) => `<label>${lb} <textarea rows="2" data-step="${s.id}|fsq.${k}">${esc(s.fsq[k])}</textarea></label>`).join('')}
  <label>Notes <textarea rows="2" data-step="${s.id}|notes">${esc(s.notes)}</textarea></label>`;
}
function nodePanel(n) {
  const S = ST();
  return `<h3>Process block</h3><label>Label <input data-node="${n.id}|label" value="${esc(n.label)}"></label><label>Symbol <select data-node="${n.id}|kind">${opt(Object.entries(FLOW_KINDS), n.kind)}</select></label><label>Owner department <select data-node="${n.id}|dept">${opt([['', '—']].concat(DEPTS), n.dept)}</select></label>
  <label>Linked step (data) <select data-node="${n.id}|stepId">${opt([['', '— none —']].concat(S.steps.map((s) => [s.id, s.name])), n.stepId || '')}</select></label>${n.stepId ? `<button class="btn" data-action="selStep" data-id="${n.stepId}">Open step data</button>` : ''}
  <p class="muted sm">To add a decision branch: add a Decision block, use Connect to draw each branch, then select the connector and label it Yes / No.</p><button class="btn" data-action="delSel">Delete block</button>`;
}
function swimPanel(it) {
  const S = ST();
  return `<h3>Swimlane activity</h3><label>Activity <input data-swim="${it.id}|label" value="${esc(it.label)}"></label><label>Lane <select data-swim="${it.id}|lane">${opt(DEPTS, it.lane)}</select></label><label>Type <select data-swim="${it.id}|kind">${opt(Object.entries(SWIM_KINDS), it.kind)}</select></label><label>Sequence column <input type="number" min="0" data-swim="${it.id}|col" value="${it.col}"></label>
  <h4>Flags</h4>${Object.entries(SWIM_FLAGS).map(([k, lb]) => `<label class="chk"><input type="checkbox" data-swimflag="${it.id}|${k}"${it.flags && it.flags[k] ? ' checked' : ''}> ${lb}</label>`).join('')}
  <h4>Next activities</h4>${(it.next || []).map((nid) => `<div class="row-fields"><span>${esc(S.swim.find((x) => x.id === nid)?.label || '?')}</span><button class="btn xs" data-action="swUnlink" data-a="${it.id}" data-b="${nid}">Remove</button></div>`).join('') || '<p class="muted sm">None</p>'}<button class="btn" data-action="delSel">Delete activity</button>`;
}
