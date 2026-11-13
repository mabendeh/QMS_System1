/* =====================================================================
   I/O: autosave, file save, Excel import/export, PowerPoint, PDF, SVG/PNG
   ===================================================================== */
const LIBS = {
  xlsx: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  pptx: 'https://cdn.jsdelivr.net/npm/pptxgenjs@4.0.1/dist/pptxgen.bundle.js',
  jszip: 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
  jspdf: 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',
  autotable: 'https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js',
};
const _loaded = {};
function loadScript(key) {
  if (_loaded[key]) return _loaded[key];
  _loaded[key] = new Promise((res, rej) => { const s = document.createElement('script'); s.src = LIBS[key]; s.onload = res; s.onerror = () => { delete _loaded[key]; rej(new Error('Could not load the ' + key + ' library. Check the internet connection and try again.')); }; document.head.appendChild(s); });
  return _loaded[key];
}
function stamp() { return new Date().toISOString().slice(0, 10); }
function fileBase() { return 'CHG_Duncanville_' + LINE().name.replace(/[^A-Za-z0-9]+/g, '_').replace(/_+$/, ''); }

/* ---------------- autosave ---------------- */
function migrate(p) {
  if (!p || !p.lines) return null;
  p.opps = p.opps || []; p.actions = p.actions || []; p.checks = p.checks || {};
  p.lines.forEach((l) => { l.settings = Object.assign(defaultSettings(), l.settings || {}); ['current', 'future'].forEach((k) => { const S = l[k]; if (!S) return; S.routes = S.routes || []; S.swim = S.swim || []; S.acts = S.acts || []; S.layout = S.layout || defaultLayout(); S.flow = S.flow || defaultFlow(S.steps); S.steps.forEach((s) => { FIELDS.forEach((fd) => (s.f[fd.k] = s.f[fd.k] || F())); s.fsq = Object.assign({ foodSafety: '', allergen: '', sanitation: '', trace: '', qc: '' }, s.fsq || {}); }); }); });
  p.opps.forEach((o) => (o.fin = Object.assign({ category: 'Not yet classified', baseline: '', method: '', status: 'Not started', validated: null, reviewer: '', date: '' }, o.fin || {})));
  if (!p.lines.some((l) => l.id === p.activeLine)) p.activeLine = p.lines[0].id;
  return p;
}
const Store = {
  mode: 'none', db: null, base: '', last: {}, timer: null, busy: false, again: false,
  loadLocal() { try { const s = localStorage.getItem('chgpm.project'); if (s) { this.mode = 'local'; return migrate(JSON.parse(s)); } } catch (_) { } return null; },
  schedule() { clearTimeout(this.timer); this.status('Unsaved changes…'); this.timer = setTimeout(() => this.save(), 1500); },
  status(t) { const el = $('#saveState'); if (el) el.textContent = t; },
  async save() {
    if (this.busy) { this.again = true; return; }
    this.busy = true;
    try {
      const json = JSON.stringify(PRJ);
      try { localStorage.setItem('chgpm.project', json); if (this.mode === 'none') this.mode = 'local'; } catch (_) { }
      if (this.db) {
        const meta = Object.assign({}, PRJ, { lines: PRJ.lines.map((l) => l.id) });
        const docs = [['mapper_meta', JSON.stringify(meta)]].concat(PRJ.lines.map((l) => ['mapper_line_' + l.id, JSON.stringify(l)]));
        for (const [name, body] of docs) {
          if (this.last[name] === body) continue;
          if (body.length > 250000) { toast('This line is too large to autosave to your account. Export to Excel to keep it.', 'warn'); continue; }
          await this.db.doc(this.base + name).set({ json: body, updated: PRJ.updated });
          this.last[name] = body;
        }
      }
      this.status(this.mode === 'db' ? 'Saved to your account' : this.mode === 'local' ? 'Saved in this browser' : 'Not saved — export to Excel');
    } catch (e) { console.warn('save failed', e); this.status('Autosave failed — export to Excel'); }
    this.busy = false; if (this.again) { this.again = false; this.save(); }
  },
  async init() {
    try {
      if (!window.claude || !claude.use) { this.status(this.mode === 'local' ? 'Saved in this browser' : 'Autosave in this browser'); return; }
      const [db, user] = await Promise.all([claude.use('db'), claude.use('user')]);
      if (!db || !user) { this.status(this.mode === 'local' ? 'Saved in this browser' : 'Autosave in this browser'); return; }
      const id = await user.id(); if (!id) return;
      this.db = db; this.base = 'data/users/' + id + '/';
      const m = await db.doc(this.base + 'mapper_meta').get();
      if (m.exists) {
        const meta = JSON.parse(m.data().json);
        const lines = [];
        for (const lid of meta.lines) { const d = await db.doc(this.base + 'mapper_line_' + lid).get(); if (d.exists) { this.last['mapper_line_' + lid] = d.data().json; lines.push(JSON.parse(d.data().json)); } }
        this.last.mapper_meta = m.data().json;
        const remote = migrate(Object.assign(meta, { lines }));
        if (remote && (!PRJ.updated || remote.updated >= PRJ.updated)) { PRJ = remote; render(); }
        this.mode = 'db'; this.status('Saved to your account');
        if (PRJ !== remote) this.schedule();
      } else { this.mode = 'db'; await this.save(); }
    } catch (e) { console.warn('store init', e); }
  },
};

/* ---------------- saving files ---------------- */
let _dl;
async function saveFile(name, data) {
  try {
    if (window.claude && claude.use) {
      if (_dl === undefined) _dl = await Promise.race([claude.use('downloads'), new Promise((r) => setTimeout(() => r(null), 2500))]);
      if (_dl) { await _dl.save({ filename: name, data }); toast('Saved ' + name); return; }
    }
  } catch (e) {
    if (e && e.code === 'declined') { toast('Save cancelled'); return; }
    if (e && e.code === 'rate_limited') { toast('A save prompt is already open'); return; }
    if (e && e.code && !['unavailable', 'not_granted', 'capability_disabled', 'capability_removed'].includes(e.code)) { toast('Could not save: ' + (e.message || e.code), 'warn'); return; }
  }
  const blob = data instanceof Blob ? data : new Blob([data]);
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); toast('Downloaded ' + name);
}
async function busy(label, fn) { toast(label + '…'); try { await fn(); } catch (e) { console.error(e); toast(e.message || String(e), 'warn'); } }

/* ---------------- Excel ---------------- */
const S2T = { VER: 'Verified', EST: 'Estimate', NYC: 'Not Yet Collected' };
const t2s = (t, v) => (v === null || v === undefined || v === '' ? 'NYC' : /verif/i.test(t || '') ? 'VER' : 'EST');
const SHEET_HEADERS = {
  Lines: ['Line ID', 'Line', 'Unit', 'Has future state', 'Notes'],
  Settings: ['Line ID', 'Key', 'Setting', 'Unit', 'Value', 'Status', 'Source', 'How to collect'],
  Steps: ['Line ID', 'State', 'Step ID', 'Order', 'Step', 'Department', 'Process owner', 'Equipment', 'Applies to line', 'VA class', 'Type', 'Food safety', 'Allergen', 'Sanitation', 'Traceability', 'Quality control', 'Notes'],
  Measurements: ['Line ID', 'State', 'Step ID', 'Step', 'Field key', 'Measure', 'Unit', 'Value', 'Status', 'Source', 'How to measure'],
  Activities: ['Line ID', 'State', 'Activity ID', 'Step ID', 'Step', 'Activity', 'Type', 'Time (min)', 'Distance (ft)', 'People', 'Class', 'Status', 'Source', 'Improvement opportunity'],
  Layouts: ['Line ID', 'State', 'Width (ft)', 'Depth (ft)', 'Note'],
  Zones: ['Line ID', 'State', 'Zone ID', 'Zone', 'X (ft)', 'Y (ft)', 'Width (ft)', 'Depth (ft)', 'Color'],
  Routes: ['Line ID', 'State', 'Route ID', 'Category', 'Route', 'Points (x,y ft)', 'Trips per shift', 'People', 'Status', 'Distance per trip (ft)', 'Distance per shift (ft)', 'Time per shift (min)', 'Labor hr/yr', 'Labor cost $/yr'],
  'Flow Nodes': ['Line ID', 'State', 'Node ID', 'Symbol', 'Label', 'X', 'Y', 'Department', 'Step ID', 'Attached to'],
  'Flow Edges': ['Line ID', 'State', 'Edge ID', 'From node', 'To node', 'Label'],
  Swimlane: ['Line ID', 'State', 'Item ID', 'Lane', 'Column', 'Activity', 'Type', 'Flags', 'Next items'],
  Opportunities: ['Opp ID', 'Line ID', 'Step ID', 'Step', 'Waste', 'Opportunity', 'Priority', 'Owner', 'Est. annual savings ($)', 'Estimate basis', 'Source'],
  Actions: ['Action ID', 'Opp ID', 'Action', 'Owner', 'Due date', 'Status', 'Notes'],
  'Finance Validation': ['Opp ID', 'Opportunity', 'Benefit type', 'Baseline', 'Valuation method', 'Estimated ($/yr)', 'Validation status', 'Validated ($/yr)', 'Finance reviewer', 'Date'],
  Checklist: ['Key', 'Line', 'Area', 'Item', 'How to measure', 'Done'],
};
const nv = (v) => (v === null || v === undefined || v === '' ? '' : v);
function buildRows() {
  const R = {}; Object.keys(SHEET_HEADERS).forEach((k) => (R[k] = [SHEET_HEADERS[k]]));
  PRJ.lines.forEach((l) => {
    R.Lines.push([l.id, l.name, l.unit, l.future ? 'Yes' : 'No', l.notes]);
    SETTINGS.forEach((sd) => { const f = l.settings[sd.k]; R.Settings.push([l.id, sd.k, sd.label, sd.unit, nv(f.v), S2T[fstat(f)], f.src, sd.how]); });
    ['current', 'future'].forEach((key) => {
      const S = l[key]; if (!S) return; const st = key === 'current' ? 'Current' : 'Future'; const c = calcState(l, key);
      S.steps.forEach((s, i) => {
        R.Steps.push([l.id, st, s.id, i + 1, s.name, s.dept, s.owner, s.equipment, s.applic, s.vaClass, s.kind, s.fsq.foodSafety, s.fsq.allergen, s.fsq.sanitation, s.fsq.trace, s.fsq.qc, s.notes]);
        FIELDS.forEach((fd) => { const f = s.f[fd.k]; R.Measurements.push([l.id, st, s.id, s.name, fd.k, fd.label, fd.unit.replace('units', l.unit + 's'), nv(f.v), S2T[fstat(f)], f.src, fd.how]); });
      });
      S.acts.forEach((a) => R.Activities.push([l.id, st, a.id, a.stepId, stepName(a.stepId), a.desc, a.type, nv(a.time), nv(a.dist), nv(a.people), a.va, S2T[a.time == null || a.time === '' ? 'NYC' : a.status], a.src, a.opportunity]));
      R.Layouts.push([l.id, st, S.layout.width, S.layout.height, S.layout.note]);
      S.layout.zones.forEach((z) => R.Zones.push([l.id, st, z.id, z.name, z.x, z.y, z.w, z.h, z.color]));
      S.routes.forEach((r) => { const x = c.routes[r.id]; R.Routes.push([l.id, st, r.id, r.cat, r.name, r.points.map((p) => p.join(',')).join('; '), nv(r.trips.v), nv(r.people.v), S2T[r.status], +x.dist.toFixed(1), x.distShift == null ? '' : +x.distShift.toFixed(0), x.timeShift == null ? '' : +x.timeShift.toFixed(1), x.hrsYear.v == null ? '' : +x.hrsYear.v.toFixed(1), x.costYear.v == null ? '' : +x.costYear.v.toFixed(0)]); });
      S.flow.nodes.forEach((n) => R['Flow Nodes'].push([l.id, st, n.id, n.kind, n.label, n.x, n.y, n.dept, n.stepId || '', n.side || '']));
      S.flow.edges.forEach((e) => R['Flow Edges'].push([l.id, st, e.id, e.from, e.to, e.label]));
      S.swim.forEach((i) => R.Swimlane.push([l.id, st, i.id, i.lane, i.col, i.label, i.kind, Object.keys(i.flags || {}).filter((k) => i.flags[k]).join(', '), (i.next || []).join(', ')]));
    });
    checklistItems(l).forEach((i) => R.Checklist.push([i.k, l.name, i.area, i.item, i.how, PRJ.checks[i.k] ? 'Yes' : '']));
  });
  PRJ.opps.forEach((o) => { R.Opportunities.push([o.id, o.lineId, o.stepId, stepName(o.stepId), o.waste, o.desc, o.priority, o.owner, nv(o.estAnnual), o.estBasis, o.source]); R['Finance Validation'].push([o.id, o.desc, o.fin.category, o.fin.baseline, o.fin.method, nv(o.estAnnual), o.fin.status, nv(o.fin.validated), o.fin.reviewer, o.fin.date]); });
  PRJ.actions.forEach((a) => R.Actions.push([a.id, a.oppId, a.action, a.owner, a.due, a.status, a.notes]));
  return R;
}
function kpiRows() {
  const rows = [['Line', 'State', 'KPI', 'Value', 'Unit', 'Data status', 'Missing inputs', 'Formula']];
  PRJ.lines.forEach((l) => ['current', 'future'].forEach((key) => {
    if (!l[key]) return; const c = calcState(l, key);
    [['Available time per day', c.avail, 'min/day'], ['Takt time', c.takt, 's/unit'], ['Total lead time', c.leadTime, 'min'], ['Value-added time', c.vaTime, 'min'], ['Non-value-added time', c.nvaTime, 'min'], ['Process cycle efficiency', c.pce, 'fraction'], ['Bottleneck effective cycle time', c.bnECT, 's/unit'], ['Line capacity', c.capacity, l.unit + '/hr'], ['Rolled throughput yield', c.rty, 'fraction'], ['Changeover time', c.coHours, 'hr/week'], ['Total labor', c.labor, 'people'], ['Labor productivity', c.productivity, l.unit + '/labor-hr'], ['Travel labor', c.travelHrs, 'hr/yr'], ['Travel labor cost', c.travelCost, '$/yr']]
      .forEach(([n, r, u]) => rows.push([l.name, key, n, r.v == null ? '' : +(+r.v).toFixed(4), u, statusLabel(r.s), (r.missing || []).join('; '), r.formula]));
    if (key === 'current' && c.bottleneck) rows.push([l.name, key, 'Bottleneck step', c.bottleneck.name, '', statusLabel(c.bnECT.s), '', FORMULAS[3][1]]);
  }));
  return rows;
}
function compareRows() {
  const rows = [['Line', 'Measure', 'Current', 'Future', 'Change', 'Data status', 'Note']];
  PRJ.lines.filter((l) => l.future).forEach((l) => {
    const x = calcCompare(l); const P2 = (n, a, b) => rows.push([l.name, n, a.v ?? '', b.v ?? '', a.v != null && b.v != null ? b.v - a.v : '', statusLabel(worst([a.s, b.s])), '']);
    P2('Total lead time (min)', x.c.leadTime, x.f.leadTime); P2('Value-added time (min)', x.c.vaTime, x.f.vaTime); P2('PCE', x.c.pce, x.f.pce); P2('Bottleneck effective C/T (s)', x.c.bnECT, x.f.bnECT); P2('Capacity (' + l.unit + '/hr)', x.c.capacity, x.f.capacity); P2('Rolled throughput yield', x.c.rty, x.f.rty); P2('Changeover (hr/week)', x.c.coHours, x.f.coHours); P2('Travel labor (hr/yr)', x.c.travelHrs, x.f.travelHrs);
    [['Estimated travel labor savings ($/yr)', x.travelSavings], ['Changeover hours freed (hr/yr)', x.coHrsYr], ['Changeover labor value ($/yr)', x.coValue], ['Scrap reduction at cost ($/yr)', x.scrapSavings], ['Throughput improvement potential', x.throughputGain]].forEach(([n, r]) => rows.push([l.name, n, '', '', r.v ?? '', statusLabel(r.s), 'Estimate — not validated by Finance. ' + (r.note || '')]));
  });
  return rows;
}
async function exportXLSX() {
  await busy('Building Excel workbook', async () => {
    await loadScript('xlsx');
    const wb = XLSX.utils.book_new(); const R = buildRows();
    const readme = [['CHG Process Mapper — data workbook'], ['C.H. Guenther & Son — Duncanville, TX'], ['Exported', new Date().toLocaleString()], ['App version', APP_VERSION], [''],
      [PRJ.demo ? 'WARNING: this workbook contains DEMONSTRATION DATA. Values marked DEMO are illustrative and are not plant measurements.' : 'Contains only values entered by the team. Blank values are Not Yet Collected.'], [''],
      ['Data status', 'Verified = measured on the floor or from a system of record, source recorded. Estimate = assumption or engineering estimate. Not Yet Collected = blank.'],
      ['Edit', 'Edit the data sheets (Settings, Steps, Measurements, Activities, Zones, Routes, Flow, Swimlane, Opportunities, Actions, Finance Validation), then import the workbook back into the app.'],
      ['Do not edit', 'KPIs and Compare are calculated by the app on export.'], ['IDs', 'Keep the ID columns unchanged so rows stay linked.']];
    const add = (name, rows, widths) => { const ws = XLSX.utils.aoa_to_sheet(rows); ws['!cols'] = (widths || rows[0].map((h) => ({ wch: Math.min(48, Math.max(10, String(h).length + 4)) }))); if (rows.length > 1) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length - 1, c: rows[0].length - 1 } }) }; XLSX.utils.book_append_sheet(wb, ws, name); };
    add('README', readme, [{ wch: 22 }, { wch: 110 }]);
    Object.keys(SHEET_HEADERS).forEach((k) => add(k, R[k]));
    add('KPIs', kpiRows()); add('Compare', compareRows()); add('Formulas', [['Calculation', 'Formula', 'Unit']].concat(FORMULAS));
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
    await saveFile(`CHG_Duncanville_Process_Mapping_${stamp()}.xlsx`, new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  });
}
function readSheet(wb, name) { const ws = wb.Sheets[name]; if (!ws) return []; return XLSX.utils.sheet_to_json(ws, { defval: '', raw: true }); }
const numOrNull = (x) => (x === '' || x === null || x === undefined || isNaN(+x) ? null : +x);
function excelDate(x) { if (typeof x === 'number') { const d = new Date(Math.round((x - 25569) * 86400000)); return d.toISOString().slice(0, 10); } return String(x || ''); }
async function importXLSX(file) {
  await busy('Importing ' + file.name, async () => {
    await loadScript('xlsx');
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const linesRows = readSheet(wb, 'Lines'); if (!linesRows.length) throw new Error('This workbook has no Lines sheet. Use a workbook exported from this app or the CHG process mapping template.');
    const p = { app: 'CHG Process Mapper', version: APP_VERSION, site: 'C.H. Guenther & Son — Duncanville, TX', demo: false, updated: new Date().toISOString(), lines: [], opps: [], actions: [], checks: {} };
    const S = (x) => String(x ?? '');
    linesRows.forEach((r) => {
      const l = { id: S(r['Line ID']), name: S(r['Line']), unit: S(r['Unit']) || 'unit', notes: S(r['Notes']), settings: defaultSettings(), current: null, future: null };
      readSheet(wb, 'Settings').filter((x) => S(x['Line ID']) === l.id).forEach((x) => { const v = numOrNull(x['Value']); l.settings[S(x['Key'])] = { v, s: t2s(x['Status'], v), src: S(x['Source']) }; });
      ['Current', 'Future'].forEach((st) => {
        const steps = readSheet(wb, 'Steps').filter((x) => S(x['Line ID']) === l.id && S(x['State']) === st).sort((a, b) => a['Order'] - b['Order']);
        if (!steps.length) return;
        const state = { steps: [], flow: { nodes: [], edges: [] }, acts: [], layout: { width: 420, height: 260, note: '', zones: [] }, routes: [], swim: [] };
        steps.forEach((x) => { const s = mkStep(S(x['Step']), S(x['Department']) || 'Production'); Object.assign(s, { id: S(x['Step ID']) || s.id, owner: S(x['Process owner']), equipment: S(x['Equipment']), applic: S(x['Applies to line']) || 'To verify', vaClass: S(x['VA class']) || 'VA', kind: S(x['Type']) || 'operation', notes: S(x['Notes']) }); s.fsq = { foodSafety: S(x['Food safety']), allergen: S(x['Allergen']), sanitation: S(x['Sanitation']), trace: S(x['Traceability']), qc: S(x['Quality control']) }; state.steps.push(s); });
        readSheet(wb, 'Measurements').filter((x) => S(x['Line ID']) === l.id && S(x['State']) === st).forEach((x) => { const s = state.steps.find((y) => y.id === S(x['Step ID'])); if (s && FIELD[S(x['Field key'])]) { const v = numOrNull(x['Value']); s.f[S(x['Field key'])] = { v, s: t2s(x['Status'], v), src: S(x['Source']) }; } });
        readSheet(wb, 'Activities').filter((x) => S(x['Line ID']) === l.id && S(x['State']) === st).forEach((x) => { const t = numOrNull(x['Time (min)']); state.acts.push({ id: S(x['Activity ID']) || uid('a'), stepId: S(x['Step ID']), desc: S(x['Activity']), type: S(x['Type']) || 'O', time: t, dist: numOrNull(x['Distance (ft)']), people: numOrNull(x['People']), va: S(x['Class']) || 'NVA', status: t2s(x['Status'], t), src: S(x['Source']), opportunity: S(x['Improvement opportunity']) }); });
        const lay = readSheet(wb, 'Layouts').find((x) => S(x['Line ID']) === l.id && S(x['State']) === st); if (lay) Object.assign(state.layout, { width: +lay['Width (ft)'] || 420, height: +lay['Depth (ft)'] || 260, note: S(lay['Note']) });
        readSheet(wb, 'Zones').filter((x) => S(x['Line ID']) === l.id && S(x['State']) === st).forEach((x) => state.layout.zones.push({ id: S(x['Zone ID']) || uid('z'), name: S(x['Zone']), x: +x['X (ft)'] || 0, y: +x['Y (ft)'] || 0, w: +x['Width (ft)'] || 20, h: +x['Depth (ft)'] || 20, color: S(x['Color']) || '#5B7FA6' }));
        readSheet(wb, 'Routes').filter((x) => S(x['Line ID']) === l.id && S(x['State']) === st).forEach((x) => { const pts = S(x['Points (x,y ft)']).split(';').map((q) => q.trim().split(',').map(Number)).filter((q) => q.length === 2 && q.every(isFinite)); const stt = /verif/i.test(S(x['Status'])) ? 'VER' : 'EST'; const tr = numOrNull(x['Trips per shift']), pe = numOrNull(x['People']); state.routes.push({ id: S(x['Route ID']) || uid('r'), cat: ROUTE_CATS[S(x['Category'])] ? S(x['Category']) : 'walk', name: S(x['Route']), points: pts, trips: { v: tr, s: tr == null ? 'NYC' : stt, src: '' }, people: { v: pe, s: pe == null ? 'NYC' : stt, src: '' }, status: stt }); });
        readSheet(wb, 'Flow Nodes').filter((x) => S(x['Line ID']) === l.id && S(x['State']) === st).forEach((x) => state.flow.nodes.push({ id: S(x['Node ID']), kind: FLOW_KINDS[S(x['Symbol'])] ? S(x['Symbol']) : 'operation', label: S(x['Label']), x: +x['X'] || 0, y: +x['Y'] || 0, dept: S(x['Department']), stepId: S(x['Step ID']), side: S(x['Attached to']) }));
        readSheet(wb, 'Flow Edges').filter((x) => S(x['Line ID']) === l.id && S(x['State']) === st).forEach((x) => state.flow.edges.push({ id: S(x['Edge ID']) || uid('e'), from: S(x['From node']), to: S(x['To node']), label: S(x['Label']) }));
        if (!state.flow.nodes.length) state.flow = defaultFlow(state.steps);
        readSheet(wb, 'Swimlane').filter((x) => S(x['Line ID']) === l.id && S(x['State']) === st).forEach((x) => { const flags = {}; S(x['Flags']).split(',').map((q) => q.trim()).filter(Boolean).forEach((k) => (flags[k] = true)); state.swim.push({ id: S(x['Item ID']) || uid('w'), lane: DEPTS.includes(S(x['Lane'])) ? S(x['Lane']) : 'Production', col: +x['Column'] || 0, label: S(x['Activity']), kind: S(x['Type']) || 'task', flags, next: S(x['Next items']).split(',').map((q) => q.trim()).filter(Boolean) }); });
        l[st === 'Current' ? 'current' : 'future'] = state;
      });
      if (!l.current) l.current = newState('generic');
      p.lines.push(l);
    });
    const fin = {}; readSheet(wb, 'Finance Validation').forEach((x) => (fin[S(x['Opp ID'])] = x));
    readSheet(wb, 'Opportunities').forEach((x) => { const f = fin[S(x['Opp ID'])] || {}; p.opps.push({ id: S(x['Opp ID']) || uid('o'), lineId: S(x['Line ID']), stepId: S(x['Step ID']), waste: S(x['Waste']) || 'Waiting', desc: S(x['Opportunity']), priority: S(x['Priority']) || 'Medium', owner: S(x['Owner']), estAnnual: numOrNull(x['Est. annual savings ($)']), estBasis: S(x['Estimate basis']), source: S(x['Source']), fin: { category: S(f['Benefit type']) || 'Not yet classified', baseline: S(f['Baseline']), method: S(f['Valuation method']), status: S(f['Validation status']) || 'Not started', validated: numOrNull(f['Validated ($/yr)']), reviewer: S(f['Finance reviewer']), date: excelDate(f['Date']) } }); });
    readSheet(wb, 'Actions').forEach((x) => p.actions.push({ id: S(x['Action ID']) || uid('t'), oppId: S(x['Opp ID']), action: S(x['Action']), owner: S(x['Owner']), due: excelDate(x['Due date']), status: S(x['Status']) || 'Not started', notes: S(x['Notes']) }));
    readSheet(wb, 'Checklist').forEach((x) => { if (/^y/i.test(S(x['Done']))) p.checks[S(x['Key'])] = true; });
    p.demo = JSON.stringify(p).includes('DEMO');
    p.activeLine = p.lines[0].id; PRJ = migrate(p); UI.state = 'current'; UI.sel = null; changed(); toast(`Imported ${p.lines.length} line(s) from ${file.name}`);
  });
}
async function exportJSON() { await saveFile(`CHG_Duncanville_Process_Mapping_backup_${stamp()}.json`, JSON.stringify(PRJ, null, 1)); }
async function importJSON(file) { await busy('Restoring backup', async () => { const p = migrate(JSON.parse(await file.text())); if (!p) throw new Error('This file is not a process mapper backup.'); PRJ = p; UI.sel = null; UI.state = 'current'; changed(); toast('Backup restored'); }); }

/* ---------------- SVG / PNG ---------------- */
function svgToPng(svg, w, h, scale = 2) {
  return new Promise((res, rej) => {
    const img = new Image(); const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
    img.onload = () => { const c = document.createElement('canvas'); c.width = Math.round(w * scale); c.height = Math.round(h * scale); const g = c.getContext('2d'); g.drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url); c.toBlob((b) => (b ? res({ blob: b, canvas: c }) : rej(new Error('PNG render failed'))), 'image/png'); };
    img.onerror = () => { URL.revokeObjectURL(url); rej(new Error('Could not render the diagram')); }; img.src = url;
  });
}
function diagramFor(kind, key = UI.state, P = PALETTES.light) { const l = LINE(); return { flow: svgFlow, vsm: svgVSM, spaghetti: svgSpaghetti, swimlane: svgSwim }[kind](l, key, P, { demo: PRJ.demo, opps: PRJ.opps.filter((o) => o.lineId === l.id) }); }
async function exportDiagram(kind, fmtx) {
  await busy('Exporting ' + fmtx.toUpperCase(), async () => {
    const d = diagramFor(kind); const name = `${fileBase()}_${kind}_${UI.state}_${stamp()}.${fmtx}`;
    if (fmtx === 'svg') await saveFile(name, d.svg);
    else { const { blob } = await svgToPng(d.svg, d.w, d.h, 2); await saveFile(name, blob); }
  });
}

/* ---------------- PowerPoint (native, editable shapes) ---------------- */
async function fixPptx(buf) {
  // Compatibility fixes for PowerPoint desktop & mobile: unique shape ids; separate notes-master theme.
  await loadScript('jszip');
  const zip = await JSZip.loadAsync(buf);
  const names = Object.keys(zip.files);
  for (const n of names.filter((x) => /^ppt\/(slides\/slide|slideLayouts\/slideLayout|slideMasters\/slideMaster|notesSlides\/notesSlide|notesMasters\/notesMaster)\d+\.xml$/.test(x))) {
    let x = await zip.file(n).async('string'); let k = 1; x = x.replace(/<p:cNvPr id="\d+"/g, () => `<p:cNvPr id="${++k}"`);
    if (/\bc[xy]="-\d/.test(x)) throw new Error('Internal error: negative shape size in ' + n);
    zip.file(n, x);
  }
  if (zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels') && !zip.file('ppt/theme/theme2.xml')) {
    zip.file('ppt/theme/theme2.xml', await zip.file('ppt/theme/theme1.xml').async('string'));
    zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', (await zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels').async('string')).replace('../theme/theme1.xml', '../theme/theme2.xml'));
    zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml').async('string')).replace('</Types>', '<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'));
  }
  // Re-pack with files only (no folder entries — they break cloud / mobile openers), [Content_Types].xml first.
  const out = new JSZip(); const files = Object.keys(zip.files).filter((n) => !zip.files[n].dir).sort((a, b) => (a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0));
  for (const n of files) out.file(n, await zip.file(n).async('uint8array'), { createFolders: false });
  return out.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
}
async function exportPPTX() {
  await busy('Building editable PowerPoint', async () => {
    await loadScript('pptx');
    const l = LINE(), key = UI.state, pres = new PptxGenJS(); pres.layout = 'LAYOUT_WIDE'; pres.title = l.name + ' process maps'; pres.company = 'C.H. Guenther and Son';
    const NAVY = '13315C', INK = '14202E', MUTED = '5E6B7A', GREEN = '2F9E6E', AMBER = 'B9770E', RED = 'C0392B', SOFT = 'EEF2F6';
    const head = (s, title, sub) => { s.background = { color: 'FFFFFF' }; s.addText(title, { x: 0.5, y: 0.3, w: 12.3, h: 0.6, fontSize: 26, bold: true, color: NAVY, fontFace: 'Arial', margin: 0 }); if (sub) s.addText(sub, { x: 0.5, y: 0.88, w: 12.3, h: 0.35, fontSize: 12, color: MUTED, fontFace: 'Arial', margin: 0 }); s.addText('C.H. GUENTHER & SON  |  DUNCANVILLE', { x: 0.5, y: 7.05, w: 6, h: 0.3, fontSize: 9, bold: true, color: NAVY, fontFace: 'Arial', margin: 0 }); if (PRJ.demo) s.addText('DEMONSTRATION DATA — NOT PLANT MEASUREMENTS', { x: 6.6, y: 7.05, w: 6.2, h: 0.3, fontSize: 9, bold: true, color: AMBER, align: 'right', fontFace: 'Arial', margin: 0 }); };
    const fitter = (w, h, box = { x: 0.5, y: 1.35, w: 12.3, h: 5.55 }) => { const k = Math.min(box.w / w, box.h / h); const ox = box.x + (box.w - w * k) / 2, oy = box.y + (box.h - h * k) / 2; return { k, X: (x) => ox + x * k, Y: (y) => oy + y * k }; };
    const line = (s, x1, y1, x2, y2, o = {}) => s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), flipH: x2 < x1, flipV: y2 < y1, line: { color: o.color || '46566A', width: o.width || 1, dashType: o.dash || 'solid', endArrowType: o.arrow ? 'triangle' : undefined } });
    const ptz = (px, k) => Math.round(Math.max(5.5, Math.min(14, px * k * 72 * 0.95)) * 2) / 2;
    // title
    let s = pres.addSlide(); s.background = { color: NAVY };
    s.addText('C.H. GUENTHER & SON  |  DUNCANVILLE', { x: 0.6, y: 0.5, w: 8, h: 0.4, fontSize: 12, bold: true, color: 'FFFFFF', fontFace: 'Arial' });
    s.addText(l.name + ' — process maps', { x: 0.6, y: 2.3, w: 12, h: 1, fontSize: 40, bold: true, color: 'FFFFFF', fontFace: 'Arial' });
    s.addText(`${key === 'future' ? 'Future' : 'Current'} state · exported ${new Date().toLocaleDateString('en-US')} · all shapes editable`, { x: 0.6, y: 3.3, w: 12, h: 0.5, fontSize: 18, color: 'D6DEE8', fontFace: 'Arial' });
    s.addText(PRJ.demo ? 'Contains DEMONSTRATION DATA. Values are illustrative, not plant measurements.' : 'Values marked NYC are Not Yet Collected. Verified and estimated values are labeled in the data workbook.', { x: 0.6, y: 6.3, w: 12, h: 0.5, fontSize: 13, color: PRJ.demo ? 'F2C46D' : 'D6DEE8', fontFace: 'Arial' });
    // KPI table
    const c = calcState(l, key);
    s = pres.addSlide(); head(s, 'Key performance indicators', 'Calculated by the CHG Process Mapper. Data status shows whether inputs are verified, estimated or incomplete.');
    const kr = [['Available time per day', c.avail, (v) => fmt(v, 0) + ' min'], ['Takt time', c.takt, (v) => fmt(v, 2) + ' s'], ['Total lead time', c.leadTime, (v) => fmt(v / 60, 2) + ' hr'], ['Value-added time', c.vaTime, (v) => fmt(v, 1) + ' min'], ['Process cycle efficiency', c.pce, (v) => pct(v, 2)], ['Bottleneck', c.bnECT, (v) => fmt(v, 2) + ' s' + (c.bottleneck ? ' — ' + c.bottleneck.name : '')], ['Line capacity', c.capacity, (v) => fmt(v, 0) + ' ' + l.unit + '/hr'], ['Rolled throughput yield', c.rty, (v) => pct(v, 1)], ['Changeover time', c.coHours, (v) => fmt(v, 1) + ' hr/week'], ['Travel labor', c.travelHrs, (v) => fmt(v, 0) + ' hr/yr']];
    const th = (t) => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: NAVY } } });
    s.addTable([[th('KPI'), th('Value'), th('Data status'), th('Missing inputs')]].concat(kr.map(([n, r, f]) => [n, r.v != null ? f(r.v) : 'NYC', statusLabel(r.s), (r.missing || []).slice(0, 3).join('; ') + ((r.missing || []).length > 3 ? ' …' : '')])), { x: 0.5, y: 1.4, w: 12.3, colW: [3, 3.2, 1.8, 4.3], fontSize: 11, fontFace: 'Arial', color: INK, border: { type: 'solid', pt: 0.5, color: 'D6DEE8' }, rowH: 0.42 });
    // process flow
    const S = lineState(l, key);
    { const fl = S.flow; let x0 = 1e9, y0 = 1e9, x1 = -1e9, y1 = -1e9; fl.nodes.forEach((n) => { const b = nodeBox(n); x0 = Math.min(x0, b.x); y0 = Math.min(y0, b.y); x1 = Math.max(x1, b.x + b.w); y1 = Math.max(y1, b.y + b.h + (n.kind === 'storage' ? 20 : 0)); });
      s = pres.addSlide(); head(s, 'Process flow map', 'Receiving to shipping, with inspections, decisions, rework and scrap. Colored dot = owner department.');
      const f = fitter(x1 - x0 + 20, y1 - y0 + 20); const X = (x) => f.X(x - x0 + 10), Y = (y) => f.Y(y - y0 + 10);
      const SH = { start: pres.shapes.FLOWCHART_TERMINATOR, end: pres.shapes.FLOWCHART_TERMINATOR, operation: pres.shapes.FLOWCHART_PROCESS, decision: pres.shapes.FLOWCHART_DECISION, inspection: pres.shapes.FLOWCHART_PREDEFINED_PROCESS, transport: pres.shapes.PENTAGON, delay: pres.shapes.FLOWCHART_DELAY, storage: pres.shapes.FLOWCHART_MERGE, rework: pres.shapes.FLOWCHART_PROCESS, scrap: pres.shapes.FLOWCHART_PROCESS };
      fl.edges.forEach((e) => { const a = fl.nodes.find((n) => n.id === e.from), b = fl.nodes.find((n) => n.id === e.to); if (!a || !b) return; const [ax, ay] = clipToBox(a.x, a.y, b.x, b.y, nodeBox(a), a.kind), [bx, by] = clipToBox(b.x, b.y, a.x, a.y, nodeBox(b), b.kind); line(s, X(ax), Y(ay), X(bx), Y(by), { arrow: true }); if (e.label) s.addText(e.label, { x: X((ax + bx) / 2) - 0.3, y: Y((ay + by) / 2) - 0.11, w: 0.6, h: 0.22, fontSize: 8, color: MUTED, align: 'center', fill: { color: 'FFFFFF' }, fontFace: 'Arial', margin: 0 }); });
      fl.nodes.forEach((n) => { const b = nodeBox(n); const isSt = n.kind === 'storage'; const bw = isSt ? 60 : b.w, bh = isSt ? 46 : b.h, bx = isSt ? n.x - 30 : b.x;
        s.addText(isSt ? '' : n.label, { shape: SH[n.kind] || pres.shapes.FLOWCHART_PROCESS, x: X(bx), y: Y(b.y), w: bw * f.k, h: bh * f.k, fill: { color: n.kind === 'start' || n.kind === 'end' ? SOFT : 'FFFFFF' }, line: { color: n.kind === 'scrap' ? RED : n.kind === 'rework' ? AMBER : NAVY, width: 1.2, dashType: n.kind === 'rework' ? 'dash' : 'solid' }, fontSize: ptz(12, f.k), color: INK, align: 'center', valign: 'middle', fontFace: 'Arial', margin: 1 });
        if (isSt) s.addText(n.label, { x: X(n.x - 60), y: Y(b.y + 48), w: 120 * f.k, h: 34 * f.k, fontSize: ptz(11, f.k), color: INK, align: 'center', fontFace: 'Arial', margin: 0 });
        if (n.dept) s.addShape(pres.shapes.OVAL, { x: X(isSt ? n.x + 30 : b.x + 4), y: Y(b.y + 4), w: 10 * f.k, h: 10 * f.k, fill: { color: (DEPT_COLORS[n.dept] || '#5E6B7A').slice(1) }, line: { type: 'none' } }); });
      let lx = 0.5; DEPTS.forEach((d) => { s.addShape(pres.shapes.OVAL, { x: lx, y: 6.83, w: 0.12, h: 0.12, fill: { color: DEPT_COLORS[d].slice(1) }, line: { type: 'none' } }); s.addText(d, { x: lx + 0.16, y: 6.77, w: 1.5, h: 0.24, fontSize: 9, color: MUTED, fontFace: 'Arial', margin: 0 }); lx += 1.7; });
    }
    // VSM
    { const steps = applicable(S.steps); s = pres.addSlide(); head(s, 'Value stream map', 'Process boxes, data boxes, inventory, lead-time ladder. NYC = Not Yet Collected. Red outline = bottleneck.');
      const colW = 220, n = steps.length, W = 120 + n * colW + 280, H = 560; const f = fitter(W, H); const X = f.X, Y = f.Y, k = f.k; const tsz = ptz(11, k);
      const bn = c.bottleneck && c.bottleneck.stepId;
      s.addText('Suppliers', { shape: pres.shapes.RECTANGLE, x: X(20), y: Y(20), w: 120 * k, h: 50 * k, fill: { color: SOFT }, line: { color: NAVY, width: 1 }, fontSize: tsz, color: INK, align: 'center', fontFace: 'Arial' });
      const dem = val(l.settings.demandPerDay);
      s.addText('Customers\n' + (dem == null ? 'Demand NYC' : fmt(dem, 0) + ' ' + l.unit + '/day'), { shape: pres.shapes.RECTANGLE, x: X(120 + n * colW + 30), y: Y(20), w: 140 * k, h: 50 * k, fill: { color: SOFT }, line: { color: NAVY, width: 1 }, fontSize: tsz, color: INK, align: 'center', fontFace: 'Arial' });
      const pcx = 120 + (n * colW) / 2 - 90; s.addText('Production Control', { shape: pres.shapes.RECTANGLE, x: X(pcx), y: Y(20), w: 180 * k, h: 50 * k, fill: { color: 'FFFFFF' }, line: { color: NAVY, width: 1 }, fontSize: tsz, color: INK, align: 'center', fontFace: 'Arial' });
      line(s, X(120 + n * colW + 30), Y(45), X(pcx + 180), Y(45), { dash: 'dash', arrow: true }); line(s, X(pcx), Y(45), X(140), Y(45), { dash: 'dash', arrow: true });
      const by = 230, bh = 66, dh = 129;
      steps.forEach((st, i) => { const x = 120 + i * colW, r = c.steps[st.id] || {};
        line(s, X(pcx + 90), Y(70), X(x + 75), Y(by - 2), { dash: 'dash', arrow: true, color: '8A97A8' });
        s.addText(st.name + '\n' + (val(st.f.labor) == null ? 'Operators NYC' : 'Operators ' + fmt(val(st.f.labor), 1)), { shape: pres.shapes.RECTANGLE, x: X(x), y: Y(by), w: 150 * k, h: bh * k, fill: { color: 'FFFFFF' }, line: { color: st.id === bn ? RED : NAVY, width: st.id === bn ? 2.5 : 1 }, fontSize: tsz, bold: true, color: INK, align: 'center', valign: 'middle', fontFace: 'Arial', margin: 1 });
        const rows = [['C/T', vsmVal(st.f.cycleTime, 's')], ['C/O', vsmVal(st.f.changeover, 'min')], ['Uptime', vsmVal(st.f.uptime, '%')], ['OEE', vsmVal(st.f.oee, '%')], ['Scrap', vsmVal(st.f.scrap, '%')], ['Eff. C/T', r.ect && r.ect.v != null ? fmt(r.ect.v, 2) + ' s' : 'NYC'], ['P/T', vsmVal(st.f.processTime, 'min')]];
        s.addText(rows.map((rw, j) => ({ text: rw[0] + ': ' + rw[1], options: { breakLine: j < rows.length - 1, color: rw[1] === 'NYC' ? AMBER : INK } })), { shape: pres.shapes.RECTANGLE, x: X(x), y: Y(by + bh), w: 150 * k, h: dh * k, fill: { color: 'F7F9FB' }, line: { color: NAVY, width: 0.75 }, fontSize: tsz * 0.9, fontFace: 'Arial', valign: 'top', margin: 2 });
        s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: X(x - 56), y: Y(by + 16), w: 32 * k, h: 28 * k, fill: { color: SOFT }, line: { color: AMBER, width: 1 } });
        s.addText(val(st.f.wip) == null ? 'WIP NYC' : fmt(val(st.f.wip), 0), { x: X(x - 70), y: Y(by + 48), w: 60 * k, h: 16 * k, fontSize: tsz * 0.85, color: val(st.f.wip) == null ? AMBER : INK, align: 'center', fontFace: 'Arial', margin: 0 });
        if (i > 0) line(s, X(x - colW + 150), Y(by + 30), X(x - 58), Y(by + 30), { width: 2.5, dash: 'sysDash', arrow: true });
        const ly = by + bh + dh + 50; line(s, X(x - 70), Y(ly), X(x), Y(ly), { color: NAVY }); line(s, X(x), Y(ly), X(x), Y(ly + 34), { color: NAVY }); line(s, X(x), Y(ly + 34), X(x + 150), Y(ly + 34), { color: NAVY }); line(s, X(x + 150), Y(ly + 34), X(x + 150), Y(ly), { color: NAVY });
        s.addText(r.queue && r.queue.v != null ? fmt(r.queue.v, 0) + ' min' : 'NYC', { x: X(x - 75), y: Y(ly - 20), w: 80 * k, h: 16 * k, fontSize: tsz * 0.85, align: 'center', color: r.queue && r.queue.v != null ? INK : AMBER, fontFace: 'Arial', margin: 0 });
        s.addText(val(st.f.processTime) == null ? 'NYC' : fmt(val(st.f.processTime), 1) + ' min', { x: X(x + 25), y: Y(ly + 38), w: 100 * k, h: 16 * k, fontSize: tsz * 0.85, align: 'center', color: val(st.f.processTime) == null ? AMBER : st.vaClass === 'VA' ? GREEN : MUTED, fontFace: 'Arial', margin: 0 });
      });
      const sx = 120 + n * colW + 20, sy = 330; const L = (lbl, r, fn) => lbl + ': ' + (r.v != null ? fn(r.v) + (r.s === 'NYC' ? ' (incomplete)' : r.s === 'EST' ? ' (est.)' : '') : 'NYC');
      s.addText([L('Lead time', c.leadTime, (v) => fmt(v / 60, 1) + ' hr'), L('Value-added', c.vaTime, (v) => fmt(v, 1) + ' min'), L('PCE', c.pce, (v) => pct(v, 2)), L('Takt', c.takt, (v) => fmt(v, 2) + ' s'), L('Bottleneck C/T', c.bnECT, (v) => fmt(v, 2) + ' s')].join('\n'), { shape: pres.shapes.RECTANGLE, x: X(sx), y: Y(sy), w: 250 * k, h: 140 * k, fill: { color: 'FFFFFF' }, line: { color: NAVY, width: 1.2 }, fontSize: tsz, color: INK, fontFace: 'Arial', valign: 'middle', margin: 4 });
    }
    // spaghetti
    { const L = S.layout; s = pres.addSlide(); head(s, 'Spaghetti diagram', L.note); const f = fitter(L.width, L.height + 18); const X = f.X, Y = f.Y, k = f.k;
      s.addShape(pres.shapes.RECTANGLE, { x: X(0), y: Y(0), w: L.width * k, h: L.height * k, fill: { color: 'FFFFFF' }, line: { color: '46566A', width: 1.5 } });
      L.zones.forEach((z) => s.addText(z.name, { shape: pres.shapes.RECTANGLE, x: X(z.x), y: Y(z.y), w: z.w * k, h: z.h * k, fill: { color: z.color.slice(1), transparency: 85 }, line: { color: z.color.slice(1), width: 1 }, fontSize: ptz(11 / 3, k), color: INK, align: 'center', valign: 'middle', fontFace: 'Arial', margin: 1 }));
      S.routes.forEach((r) => { const col = (ROUTE_CATS[r.cat] || ROUTE_CATS.walk).color.slice(1); for (let i = 1; i < r.points.length; i++) line(s, X(r.points[i - 1][0]), Y(r.points[i - 1][1]), X(r.points[i][0]), Y(r.points[i][1]), { color: col, width: 2.5, arrow: i === r.points.length - 1 }); });
      let lx = X(0); Object.values(ROUTE_CATS).forEach((d) => { line(s, lx, Y(L.height + 10), lx + 0.35, Y(L.height + 10), { color: d.color.slice(1), width: 3 }); s.addText(d.label, { x: lx + 0.4, y: Y(L.height + 10) - 0.12, w: 2.2, h: 0.24, fontSize: 9, color: MUTED, fontFace: 'Arial', margin: 0 }); lx += 2.8; });
      const rows = [[th('#'), th('Route'), th('ft/trip'), th('Trips/shift'), th('hr/yr'), th('$/yr')]].concat(S.routes.map((r, i) => { const x = c.routes[r.id]; return [String(i + 1), r.name, fmt(x.dist, 0), r.trips.v == null ? 'NYC' : String(r.trips.v), x.hrsYear.v == null ? 'NYC' : fmt(x.hrsYear.v, 0), x.costYear.v == null ? 'NYC' : money(x.costYear.v)]; }));
      if (S.routes.length) { s = pres.addSlide(); head(s, 'Spaghetti diagram — route calculations', 'Distance from drawn routes (1 unit = 1 ft) × trips ÷ speed; annual labor = time × people × shifts × days.'); s.addTable(rows, { x: 0.5, y: 1.4, w: 12.3, colW: [0.5, 6.3, 1.2, 1.4, 1.3, 1.6], fontSize: 10, fontFace: 'Arial', color: INK, border: { type: 'solid', pt: 0.5, color: 'D6DEE8' } }); }
    }
    // swimlane
    { const items = S.swim; s = pres.addSlide(); head(s, 'Swimlane process map', 'Responsibilities by department. Dashed = cross-department handoff. Red outline = flagged (duplicate, unnecessary approval, delay, communication gap).');
      const d = svgSwim(l, key, PALETTES.light); const f = fitter(d.w, d.h - 40); const X = f.X, Y = f.Y, k = f.k;
      DEPTS.forEach((dp, i) => { s.addShape(pres.shapes.RECTANGLE, { x: X(0), y: Y(d.top + i * d.laneH), w: d.w * k, h: d.laneH * k, fill: { color: i % 2 ? 'F7F9FB' : 'FFFFFF' }, line: { color: 'D3DBE5', width: 0.5 } }); s.addText(dp, { shape: pres.shapes.RECTANGLE, x: X(0), y: Y(d.top + i * d.laneH), w: d.labW * k, h: d.laneH * k, fill: { color: SOFT }, line: { color: 'D3DBE5', width: 0.5 }, fontSize: ptz(12, k), bold: true, color: INK, align: 'center', valign: 'middle', fontFace: 'Arial' }); });
      items.forEach((it) => (it.next || []).forEach((nid) => { const t = items.find((q) => q.id === nid); if (!t) return; const a = d.pos(it), b = d.pos(t); const x1 = a.x + a.w, y1 = a.y + a.h / 2, x2 = b.x, y2 = b.y + b.h / 2, mx = Math.min(x2 - 12, x1 + 16); const cross = it.lane !== t.lane; const o = { color: cross ? NAVY : '46566A', dash: cross ? 'dash' : 'solid' }; line(s, X(x1), Y(y1), X(mx), Y(y1), o); line(s, X(mx), Y(y1), X(mx), Y(y2), o); line(s, X(mx), Y(y2), X(x2), Y(y2), Object.assign({ arrow: true }, o)); }));
      items.forEach((it) => { const p = d.pos(it), flagged = Object.values(it.flags || {}).some(Boolean); const flags = Object.keys(it.flags || {}).filter((q) => it.flags[q]).map((q) => SWIM_FLAGS[q]); s.addText(it.label + (flags.length ? '\n⚑ ' + flags.join(', ') : ''), { shape: it.kind === 'decision' ? pres.shapes.FLOWCHART_DECISION : it.kind === 'delay' ? pres.shapes.FLOWCHART_DELAY : it.kind === 'handoff' ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, x: X(p.x), y: Y(p.y), w: p.w * k, h: p.h * k, fill: { color: 'FFFFFF' }, line: { color: flagged ? RED : NAVY, width: flagged ? 2 : 1 }, fontSize: ptz(11, k), color: INK, align: 'center', valign: 'middle', fontFace: 'Arial', margin: 1 }); });
    }
    // opportunities, actions, finance
    const opps = PRJ.opps.filter((o) => o.lineId === l.id);
    if (opps.length) {
      s = pres.addSlide(); head(s, 'Waste & opportunity register', 'Estimated savings are not validated until Finance confirms them.');
      s.addTable([[th('Step'), th('Waste'), th('Opportunity'), th('Est. $/yr'), th('Finance status'), th('Validated $/yr')]].concat(opps.slice(0, 12).map((o) => [stepName(o.stepId), o.waste, o.desc, o.estAnnual == null || o.estAnnual === '' ? '—' : money(+o.estAnnual), o.fin.status, o.fin.validated == null ? '—' : money(+o.fin.validated)])), { x: 0.5, y: 1.4, w: 12.3, colW: [1.8, 1.4, 5, 1.2, 1.7, 1.2], fontSize: 10, fontFace: 'Arial', color: INK, border: { type: 'solid', pt: 0.5, color: 'D6DEE8' } });
      const acts = PRJ.actions.filter((a) => opps.some((o) => o.id === a.oppId));
      if (acts.length) { s = pres.addSlide(); head(s, 'Improvement action tracker'); s.addTable([[th('Action'), th('Owner'), th('Due'), th('Status')]].concat(acts.slice(0, 14).map((a) => [a.action, a.owner, a.due, a.status + (a.status !== 'Done' && a.due && a.due < todayISO() ? ' (overdue)' : '')])), { x: 0.5, y: 1.4, w: 12.3, colW: [7, 2.2, 1.4, 1.7], fontSize: 10, fontFace: 'Arial', color: INK, border: { type: 'solid', pt: 0.5, color: 'D6DEE8' } }); }
    }
    if (l.future) { const x = calcCompare(l); s = pres.addSlide(); head(s, 'Current vs future state', 'All savings are estimates until validated by Finance.');
      const rr = (n, a, b, fn) => [n, a.v != null ? fn(a.v) : 'NYC', b.v != null ? fn(b.v) : 'NYC', statusLabel(worst([a.s, b.s]))];
      s.addTable([[th('Measure'), th('Current'), th('Future'), th('Data status')], rr('Lead time (hr)', x.c.leadTime, x.f.leadTime, (v) => fmt(v / 60, 2)), rr('PCE', x.c.pce, x.f.pce, (v) => pct(v, 2)), rr('Bottleneck C/T (s)', x.c.bnECT, x.f.bnECT, (v) => fmt(v, 2)), rr('Capacity (' + l.unit + '/hr)', x.c.capacity, x.f.capacity, (v) => fmt(v, 0)), rr('Rolled throughput yield', x.c.rty, x.f.rty, (v) => pct(v, 2)), rr('Changeover (hr/week)', x.c.coHours, x.f.coHours, (v) => fmt(v, 1)), rr('Travel labor (hr/yr)', x.c.travelHrs, x.f.travelHrs, (v) => fmt(v, 0)),
        ['Est. travel labor savings', '', x.travelSavings.v != null ? money(x.travelSavings.v) + '/yr' : 'NYC', statusLabel(x.travelSavings.s)], ['Est. scrap reduction (at cost)', '', x.scrapSavings.v != null ? money(x.scrapSavings.v) + '/yr' : 'NYC', statusLabel(x.scrapSavings.s)], ['Throughput potential', '', x.throughputGain.v != null ? pct(x.throughputGain.v, 1) : 'NYC', statusLabel(x.throughputGain.s)]], { x: 0.5, y: 1.4, w: 12.3, colW: [4.5, 2.6, 2.6, 2.6], fontSize: 11, fontFace: 'Arial', color: INK, border: { type: 'solid', pt: 0.5, color: 'D6DEE8' } });
    }
    const buf = await pres.write({ outputType: 'arraybuffer' });
    const blob = await fixPptx(buf);
    await saveFile(`${fileBase()}_Process_Maps_${key}_${stamp()}.pptx`, blob);
  });
}

/* ---------------- PDF ---------------- */
async function exportPDF(compare) {
  await busy(compare ? 'Building comparison report' : 'Building PDF report', async () => {
    const l = LINE();
    if (compare && !l.future) throw new Error('Create a future state first (Current vs Future tab).');
    await loadScript('jspdf'); await loadScript('autotable');
    const { jsPDF } = window.jspdf; const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: [960, 540] });
    const NAVY = [19, 49, 92], MUTED = [94, 107, 122], AMBER = [185, 119, 14];
    let page = 0;
    const frame = (title, sub) => { if (page++) doc.addPage(); doc.setFont('helvetica', 'bold'); doc.setFontSize(22); doc.setTextColor(...NAVY); doc.text(title, 36, 46); if (sub) { doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor(...MUTED); doc.text(doc.splitTextToSize(sub, 880), 36, 64); } doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(...NAVY); doc.text('C.H. GUENTHER & SON  |  DUNCANVILLE  ·  ' + l.name, 36, 524); doc.setTextColor(...MUTED); doc.text('Page ' + page, 924, 524, { align: 'right' }); if (PRJ.demo) { doc.setTextColor(...AMBER); doc.text('DEMONSTRATION DATA — NOT PLANT MEASUREMENTS', 480, 524, { align: 'center' }); } };
    const img = async (kind, key, title, sub) => { const d = diagramFor(kind, key); const k = Math.min(888 / d.w, 420 / d.h); const sc = Math.min(2, Math.max(0.6, (d.w * k * 2.2) / d.w)); const { canvas } = await svgToPng(d.svg, d.w, d.h, sc); frame(title, sub); doc.addImage(canvas.toDataURL('image/jpeg', 0.88), 'JPEG', 36 + (888 - d.w * k) / 2, 82, d.w * k, d.h * k, undefined, 'FAST'); };
    const tbl = (head, body, opts = {}) => doc.autoTable(Object.assign({ head: [head], body, startY: 82, margin: { left: 36, right: 36, bottom: 30 }, styles: { font: 'helvetica', fontSize: 8.5, cellPadding: 3, textColor: [20, 32, 46] }, headStyles: { fillColor: NAVY, textColor: 255 }, alternateRowStyles: { fillColor: [247, 249, 251] }, didDrawPage: () => { } }, opts));
    // cover
    frame(compare ? `${l.name} — current vs future state` : `${l.name} — process mapping report`, `${PRJ.site} · ${new Date().toLocaleDateString('en-US')} · ${compare ? 'Comparison report' : (UI.state === 'future' ? 'Future' : 'Current') + ' state'}`);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(12); doc.setTextColor(20, 32, 46);
    doc.text(doc.splitTextToSize((PRJ.demo ? 'This report contains DEMONSTRATION DATA. Values are illustrative and are not plant measurements. ' : '') + 'Values shown as NYC are Not Yet Collected. Each result is labeled Verified inputs, Estimate or Incomplete according to the status of its inputs. Estimated savings are not validated until Finance confirms them against actual cost results.', 860), 36, 110);
    if (compare) {
      const x = calcCompare(l); const rr = (n, a, b, fn) => [n, a.v != null ? fn(a.v) : 'NYC', b.v != null ? fn(b.v) : 'NYC', a.v != null && b.v != null ? fn(b.v - a.v) : '—', statusLabel(worst([a.s, b.s]))];
      frame('KPI comparison', 'Current vs future state');
      tbl(['Measure', 'Current', 'Future', 'Change', 'Data status'], [rr('Total lead time (hr)', x.c.leadTime, x.f.leadTime, (v) => fmt(v / 60, 2)), rr('Value-added time (min)', x.c.vaTime, x.f.vaTime, (v) => fmt(v, 1)), rr('Process cycle efficiency', x.c.pce, x.f.pce, (v) => pct(v, 2)), rr('Bottleneck effective C/T (s)', x.c.bnECT, x.f.bnECT, (v) => fmt(v, 2)), rr('Line capacity (' + l.unit + '/hr)', x.c.capacity, x.f.capacity, (v) => fmt(v, 0)), rr('Rolled throughput yield', x.c.rty, x.f.rty, (v) => pct(v, 2)), rr('Changeover (hr/week)', x.c.coHours, x.f.coHours, (v) => fmt(v, 1)), rr('Total labor (people)', x.c.labor, x.f.labor, (v) => fmt(v, 1)), rr('Travel labor (hr/yr)', x.c.travelHrs, x.f.travelHrs, (v) => fmt(v, 0)), rr('Travel labor cost ($/yr)', x.c.travelCost, x.f.travelCost, money)]);
      tbl(['Estimated benefit (not validated)', 'Value', 'Data status', 'Note'], [['Travel labor savings', x.travelSavings.v != null ? money(x.travelSavings.v) + '/yr' : 'NYC', statusLabel(x.travelSavings.s), 'Hard savings only if paid hours come out'], ['Changeover hours freed', x.coHrsYr.v != null ? fmt(x.coHrsYr.v, 0) + ' hr/yr' : 'NYC', statusLabel(x.coHrsYr.s), ''], ['Changeover labor value', x.coValue.v != null ? money(x.coValue.v) + '/yr' : 'NYC', statusLabel(x.coValue.s), 'Capacity unless Finance confirms'], ['Scrap reduction (at cost)', x.scrapSavings.v != null ? money(x.scrapSavings.v) + '/yr' : 'NYC', statusLabel(x.scrapSavings.s), ''], ['Throughput potential', x.throughputGain.v != null ? pct(x.throughputGain.v, 1) : 'NYC', statusLabel(x.throughputGain.s), '']], { startY: doc.lastAutoTable.finalY + 18 });
      for (const k2 of ['current', 'future']) { await img('vsm', k2, `Value stream map — ${k2} state`, ''); await img('spaghetti', k2, `Spaghetti diagram — ${k2} state`, l[k2].layout.note); }
    } else {
      const key = UI.state, c = calcState(l, key);
      frame('Key performance indicators', 'Calculated values with data status and missing inputs');
      tbl(['KPI', 'Value', 'Data status', 'Missing inputs', 'Formula'], [['Takt time', c.takt, (v) => fmt(v, 2) + ' s'], ['Total lead time', c.leadTime, (v) => fmt(v / 60, 2) + ' hr'], ['Value-added time', c.vaTime, (v) => fmt(v, 1) + ' min'], ['Process cycle efficiency', c.pce, (v) => pct(v, 2)], ['Bottleneck effective C/T', c.bnECT, (v) => fmt(v, 2) + ' s' + (c.bottleneck ? ' (' + c.bottleneck.name + ')' : '')], ['Line capacity', c.capacity, (v) => fmt(v, 0) + ' ' + l.unit + '/hr'], ['Rolled throughput yield', c.rty, (v) => pct(v, 1)], ['Changeover time', c.coHours, (v) => fmt(v, 1) + ' hr/week'], ['Travel labor', c.travelHrs, (v) => fmt(v, 0) + ' hr/yr'], ['Travel labor cost', c.travelCost, money]].map(([n, r, fn]) => [n, r.v != null ? fn(r.v) : 'NYC', statusLabel(r.s), (r.missing || []).slice(0, 4).join('; '), r.formula]), { columnStyles: { 3: { cellWidth: 220 }, 4: { cellWidth: 300 } } });
      await img('flow', key, 'Process flow map', 'Operations, inspections, decisions, rework and scrap.');
      await img('vsm', key, 'Value stream map', 'Material and information flow with data boxes and lead-time ladder.');
      await img('spaghetti', key, 'Spaghetti diagram', lineState(l, key).layout.note);
      await img('swimlane', key, 'Swimlane process map', 'Responsibilities, handoffs and flagged approvals or delays.');
      frame('Step data', 'NYC = Not Yet Collected · (E) estimate · (V) verified');
      const keys = ['cycleTime', 'processTime', 'changeover', 'uptime', 'labor', 'throughput', 'scrap', 'wip', 'oee'];
      tbl(['Step', 'Applies'].concat(keys.map((k2) => FIELD[k2].label + ' (' + FIELD[k2].unit.replace('units', l.unit) + ')')), lineState(l, key).steps.map((s) => [s.name, s.applic].concat(keys.map((k2) => { const f = s.f[k2], st = fstat(f); return st === 'NYC' ? 'NYC' : fmt(+f.v, 2) + (st === 'VER' ? ' (V)' : ' (E)'); }))), { styles: { fontSize: 7.5, cellPadding: 2.5, textColor: [20, 32, 46] } });
      frame('Process activity map', `Total ${fmt(c.pam.time, 1)} min · NVA ${fmt(c.pam.nva, 1)} min · ${fmt(c.pam.dist, 0)} ft · ${c.pam.missing} activities not timed`);
      tbl(['Step', 'Activity', 'Type', 'Time (min)', 'Distance (ft)', 'People', 'Class', 'Data'], lineState(l, key).acts.map((a) => [stepName(a.stepId), a.desc, ACT_TYPES[a.type], a.time ?? 'NYC', a.dist ?? '', a.people ?? '', a.va, a.time == null || a.time === '' ? NYC : STATUS[a.status] || 'Estimate']));
      const opps = PRJ.opps.filter((o) => o.lineId === l.id);
      frame('Waste & opportunity register, actions, finance validation', 'Estimated savings are not validated until Finance confirms them.');
      tbl(['Step', 'Waste', 'Opportunity', 'Est. $/yr', 'Benefit type', 'Finance status', 'Validated $/yr'], opps.map((o) => [stepName(o.stepId), o.waste, o.desc, o.estAnnual == null || o.estAnnual === '' ? '—' : money(+o.estAnnual), o.fin.category, o.fin.status, o.fin.validated == null ? '—' : money(+o.fin.validated)]));
      const acts = PRJ.actions.filter((a) => opps.some((o) => o.id === a.oppId));
      if (acts.length) tbl(['Action', 'Owner', 'Due', 'Status'], acts.map((a) => [a.action, a.owner, a.due, a.status + (a.status !== 'Done' && a.due && a.due < todayISO() ? ' (overdue)' : '')]), { startY: doc.lastAutoTable.finalY + 16 });
      const items = checklistItems(l).filter((i) => !PRJ.checks[i.k]);
      frame('Floor-walk checklist — open items', `${items.length} items still to collect`);
      tbl(['Area', 'Item', 'How to measure', 'Done'], items.map((i) => [i.area, i.item, i.how, '']), { columnStyles: { 2: { cellWidth: 460 } } });
      frame('Formulas', 'How each result is calculated');
      tbl(['Calculation', 'Formula', 'Unit'], FORMULAS.map((f) => [f[0], f[1], f[2]]));
    }
    await saveFile(`${fileBase()}_${compare ? 'Current_vs_Future' : 'Process_Mapping_Report'}_${stamp()}.pdf`, doc.output('blob'));
  });
}
