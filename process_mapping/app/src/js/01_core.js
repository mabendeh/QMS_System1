/* =====================================================================
   CHG Process Mapper — core: constants, data model, templates, demo data
   Every measurement is a field {v, s, src}:
     v   = number or null (null = Not Yet Collected)
     s   = 'VER' (verified on the floor / system of record) | 'EST' (estimate / assumption)
     src = free-text data source (who, when, how)
   ===================================================================== */
'use strict';

const APP_VERSION = '0.9.0-prototype';
const NYC = 'Not Yet Collected';
const STATUS = { NYC: 'Not Yet Collected', EST: 'Estimate', VER: 'Verified' };
const DEMO_SRC = 'DEMO — illustrative value, not plant data';

const DEPTS = ['Production', 'Quality', 'Maintenance', 'Sanitation', 'Warehouse', 'Planning', 'Supervisors'];
const DEPT_COLORS = { Production: '#2F6FB0', Quality: '#2F9E6E', Maintenance: '#B9770E', Sanitation: '#6E59A5', Warehouse: '#5B7FA6', Planning: '#8A5A44', Supervisors: '#13315C' };
const APPLIC = ['To verify', 'Confirmed', 'Not applicable'];
const VACLASS = { VA: 'Value-added', NNVA: 'Necessary non-value-added', NVA: 'Non-value-added' };

const WASTES = [
  { k: 'Transportation', d: 'Moving material or product more than needed' },
  { k: 'Inventory', d: 'Ingredients, WIP or finished goods beyond what flow needs' },
  { k: 'Motion', d: 'People walking, reaching or searching' },
  { k: 'Waiting', d: 'People, product or equipment idle' },
  { k: 'Overproduction', d: 'Making more or earlier than the next step needs' },
  { k: 'Over-processing', d: 'Work or checks that add no value the customer needs' },
  { k: 'Defects', d: 'Scrap, rework, holds, giveaway' },
  { k: 'Skills', d: 'Unused talent and ideas of people' },
];

const FIELDS = [
  { k: 'cycleTime', label: 'Cycle time', unit: 's/unit', how: 'Stopwatch: time between successive units leaving the step. Time 10+ consecutive cycles in 3 different periods; record min / avg / max.' },
  { k: 'processTime', label: 'Processing time (value-added)', unit: 'min', how: 'Time the product is actually being transformed at this step (mix time, bake time, freeze dwell). From recipe/spec and confirmed by observation.' },
  { k: 'changeover', label: 'Changeover time', unit: 'min', how: 'Last good unit of product A to first good unit of product B. Observe 3+ changeovers; also pull the changeover log.' },
  { k: 'changeoversPerWeek', label: 'Changeovers per week', unit: '/week', how: 'Count from the production schedule or changeover log, last 4 weeks average.' },
  { k: 'downtime', label: 'Unplanned downtime', unit: 'min/shift', how: 'Downtime log, last 4 weeks, average minutes per shift; capture top 3 causes.' },
  { k: 'uptime', label: 'Uptime', unit: '%', how: '(Planned time − unplanned downtime − changeover) ÷ planned time, from line records.' },
  { k: 'labor', label: 'Labor', unit: 'people', how: 'Operators assigned to the step per shift (staffing plan), confirmed by observation.' },
  { k: 'throughput', label: 'Throughput', unit: 'units/hr', how: 'Count good units over at least 1 hour of normal running, or pull from line counter.' },
  { k: 'scrap', label: 'Scrap', unit: '%', how: 'Scrap weight ÷ input weight at the step, from the scrap log or a shift weigh-out.' },
  { k: 'yield', label: 'Yield', unit: '%', how: 'Good output ÷ input (weight basis), from production records.' },
  { k: 'wip', label: 'WIP before step', unit: 'units', how: 'Physical count of product or batches waiting in front of the step at the time of the walk; repeat at 2–3 times.' },
  { k: 'walkDist', label: 'Operator walking', unit: 'ft/shift', how: 'From the spaghetti diagram: measured route length × trips per shift (or pedometer).' },
  { k: 'transportDist', label: 'Transport distance', unit: 'ft/trip', how: 'Measured route from the previous step (tape / wheel / plant drawing).' },
  { k: 'waitTime', label: 'Waiting time before step', unit: 'min', how: 'Observed hold / queue time. Used for lead time only where WIP was not counted.' },
  { k: 'oee', label: 'OEE', unit: '%', how: 'Availability × Performance × Quality from line data. Equipment steps only.' },
  { k: 'opCost', label: 'Operating cost', unit: '$/hr', how: 'From Finance: labor + utilities + overhead rate for the step. Do not estimate without Finance.' },
];
const FIELD = Object.fromEntries(FIELDS.map((f) => [f.k, f]));

const SETTINGS = [
  { k: 'demandPerDay', label: 'Customer demand', unit: 'units/day', how: 'Planning: average daily requirement for this line, last 13 weeks.' },
  { k: 'shiftsPerDay', label: 'Shifts per day', unit: 'shifts', how: 'Planning / production schedule.' },
  { k: 'minPerShift', label: 'Shift length', unit: 'min', how: 'Scheduled minutes per shift.' },
  { k: 'breakMin', label: 'Planned breaks & meetings', unit: 'min/shift', how: 'Scheduled breaks, lunch, huddles, planned sanitation inside the shift.' },
  { k: 'daysPerYear', label: 'Operating days per year', unit: 'days', how: 'Planning: production calendar.' },
  { k: 'weeksPerYear', label: 'Operating weeks per year', unit: 'weeks', how: 'Planning: production calendar.' },
  { k: 'laborRate', label: 'Loaded labor rate', unit: '$/hr', how: 'Finance only. Do not use an estimate for reported savings.' },
  { k: 'materialCost', label: 'Material cost per unit', unit: '$/unit', how: 'Finance: standard material cost per unit (for scrap valuation at cost).' },
  { k: 'walkSpeed', label: 'Walking speed', unit: 'ft/min', how: 'Industrial-engineering planning value is 264 ft/min (3 mph); replace with observed values.' },
  { k: 'forkliftSpeed', label: 'Forklift / pallet-jack speed', unit: 'ft/min', how: 'Observe loaded travel; plant speed limits apply.' },
];

const ROUTE_CATS = {
  walk: { label: 'Employee walking', color: '#2F6FB0', speedKey: 'walkSpeed' },
  forklift: { label: 'Forklift', color: '#C2410C', speedKey: 'forkliftSpeed' },
  ingredient: { label: 'Ingredient movement', color: '#2F9E6E', speedKey: 'walkSpeed' },
  product: { label: 'Product transport', color: '#6E59A5', speedKey: 'walkSpeed' },
};

const FLOW_KINDS = {
  start: 'Start / end', operation: 'Operation', decision: 'Decision', inspection: 'Inspection',
  transport: 'Material transfer', delay: 'Delay / hold', storage: 'Storage', rework: 'Rework', scrap: 'Scrap / reject', end: 'End',
};
const ACT_TYPES = { O: 'Operation', T: 'Transportation', I: 'Inspection', D: 'Delay', S: 'Storage' };
const SWIM_KINDS = { task: 'Task', approval: 'Approval / sign-off', handoff: 'Handoff', delay: 'Wait / delay', decision: 'Decision' };
const SWIM_FLAGS = { duplicate: 'Duplicated work', unnecessary: 'Unnecessary approval', delay: 'Delay', gap: 'Communication gap' };
const ACTION_STATUS = ['Not started', 'In progress', 'Blocked', 'Done'];
const FIN_CATS = ['Hard savings', 'Cost avoidance', 'Capacity', 'Nonfinancial', 'Not yet classified'];
const FIN_STATUS = ['Not started', 'Method agreed with Finance', 'In verification', 'Validated by Finance', 'Rejected'];

/* ---------------- helpers ---------------- */
const uid = (p = '') => p + Math.random().toString(36).slice(2, 9);
const clone = (o) => JSON.parse(JSON.stringify(o));
const F = (v = null, s = 'EST', src = '') => ({ v, s: v == null ? 'NYC' : s, src });
const val = (f) => (f && f.v !== null && f.v !== '' && f.v !== undefined && isFinite(+f.v) ? +f.v : null);
const fstat = (f) => (val(f) == null ? 'NYC' : f.s === 'VER' ? 'VER' : 'EST');
const worst = (arr) => (arr.includes('NYC') ? 'NYC' : arr.includes('EST') ? 'EST' : arr.length ? 'VER' : 'NYC');
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = (n, d = 1) => (n == null || !isFinite(n) ? '—' : Math.abs(n) >= 1000 ? Math.round(n).toLocaleString('en-US') : (+n.toFixed(d)).toLocaleString('en-US'));
const money = (n) => (n == null || !isFinite(n) ? '—' : (n < 0 ? '−$' : '$') + Math.round(Math.abs(n)).toLocaleString('en-US'));
const pct = (n, d = 1) => (n == null || !isFinite(n) ? '—' : (n * 100).toFixed(d) + '%');
const todayISO = () => new Date().toISOString().slice(0, 10);

/* ---------------- templates ---------------- */
function mkStep(name, dept, opts = {}) {
  const f = {};
  FIELDS.forEach((fd) => (f[fd.k] = F()));
  return {
    id: uid('s'), name, dept, owner: '', equipment: opts.equipment || '', applic: opts.applic || 'To verify',
    vaClass: opts.vaClass || 'VA', kind: opts.kind || 'operation',
    fsq: Object.assign({ foodSafety: '', allergen: '', sanitation: '', trace: '', qc: '' }, opts.fsq || {}),
    f, notes: opts.notes || '',
  };
}

const VERIFY_HACCP = 'Typical control point — confirm against the plant food safety (HACCP / preventive controls) plan';

function genericSteps(kind) {
  const S = [
    mkStep('Raw Material Receiving', 'Warehouse', { kind: 'operation', vaClass: 'NNVA', fsq: { foodSafety: 'Receiving temperature for refrigerated / frozen ingredients; packaging integrity. ' + VERIFY_HACCP, allergen: 'Verify allergen labeling and segregation on receipt', trace: 'Record supplier lot numbers', qc: 'COA / spec check against approved supplier list' } }),
    mkStep('Ingredient Storage', 'Warehouse', { kind: 'storage', vaClass: 'NVA', fsq: { foodSafety: 'Storage temperature monitoring (coolers / freezers)', allergen: 'Allergen segregation and labeling in storage', trace: 'FIFO / FEFO rotation; lot visible', sanitation: 'Master sanitation schedule for storage areas' } }),
    mkStep('Pre-Weigh', 'Production', { kind: 'operation', vaClass: 'NNVA', fsq: { allergen: 'Allergen verification of ingredients against batch sheet; dedicated scoops / tools', trace: 'Ingredient lots recorded on the batch record', qc: 'Scale verification / calibration check' } }),
    mkStep('Mixing', 'Production', { kind: 'operation', vaClass: 'VA', fsq: { qc: 'Dough temperature and mix time per spec', trace: 'Batch number assigned', sanitation: 'Allergen changeover clean between incompatible formulas' } }),
    mkStep('Dividing / Forming', 'Production', { kind: 'operation', vaClass: 'VA', fsq: { qc: 'Piece weight checks per spec', foodSafety: 'Foreign material control (blades, guards, plastics)' } }),
    mkStep('Sheeting / Lamination', 'Production', { kind: 'operation', vaClass: 'VA', applic: 'To verify', notes: 'Where applicable on this line' }),
    mkStep('Proofing', 'Production', { kind: 'operation', vaClass: 'VA', applic: 'To verify', notes: 'Where applicable on this line', fsq: { qc: 'Proof time, temperature and humidity per spec' } }),
    mkStep('Baking / Par-Baking', 'Production', { kind: 'operation', vaClass: 'VA', applic: 'To verify', notes: 'Where applicable on this line', fsq: { foodSafety: 'Bake time / temperature; ' + VERIFY_HACCP } }),
    mkStep('Cooling', 'Production', { kind: 'operation', vaClass: 'NNVA', applic: 'To verify', notes: 'Where applicable on this line', fsq: { foodSafety: 'Cooling time / temperature; condensation and overhead protection' } }),
    mkStep('Spiral Freezing', 'Production', { kind: 'operation', vaClass: 'VA', equipment: 'Spiral freezer', fsq: { qc: 'Product core temperature at freezer exit', sanitation: 'Freezer sanitation and defrost schedule' } }),
    mkStep('Packaging', 'Production', { kind: 'operation', vaClass: 'VA', fsq: { foodSafety: 'Metal detection / X-ray reject verification. ' + VERIFY_HACCP, allergen: 'Label and allergen declaration verification at start-up and changeover', trace: 'Lot / date code applied and verified', qc: 'Checkweigher; seal integrity' } }),
    mkStep('Cold Storage', 'Warehouse', { kind: 'storage', vaClass: 'NVA', fsq: { foodSafety: 'Freezer temperature monitoring and alarms', trace: 'Pallet license plate / lot location' } }),
    mkStep('Shipping', 'Warehouse', { kind: 'operation', vaClass: 'NNVA', fsq: { foodSafety: 'Trailer pre-cool and temperature check', trace: 'Lot numbers on bill of lading for recall traceability' } }),
  ];
  if (kind === 'nugget') {
    S.find((s) => s.name === 'Sheeting / Lamination').notes = 'Applicability to the Nugget Line not yet verified';
  }
  return S;
}

function defaultSettings() {
  const st = {};
  SETTINGS.forEach((s) => (st[s.k] = F()));
  st.walkSpeed = F(264, 'EST', 'Planning value: 3 mph normal walking pace. Replace with observed speed.');
  st.weeksPerYear = F(52, 'EST', 'Assumption: 52 operating weeks. Confirm with Planning.');
  return st;
}

function defaultLayout() {
  // Schematic placeholder in feet. NOT the Duncanville layout.
  const Z = (name, x, y, w, h, c) => ({ id: uid('z'), name, x, y, w, h, color: c });
  return {
    width: 420, height: 260, note: 'Schematic placeholder — not to scale and not the actual Duncanville layout. Replace zone positions and sizes with the plant drawing.',
    zones: [
      Z('Receiving dock', 10, 10, 60, 40, '#5B7FA6'), Z('Dry ingredient storage', 80, 10, 80, 50, '#5B7FA6'), Z('Ingredient cooler', 170, 10, 50, 40, '#5B7FA6'),
      Z('Pre-weigh room', 230, 10, 50, 40, '#2F6FB0'), Z('Mixing', 290, 10, 60, 50, '#2F6FB0'), Z('Maintenance shop', 360, 10, 50, 40, '#B9770E'),
      Z('Make-up / forming line', 80, 90, 180, 40, '#2F6FB0'), Z('Spiral freezer', 280, 80, 70, 70, '#2F6FB0'), Z('Packaging', 280, 170, 90, 40, '#2F6FB0'),
      Z('QC lab', 10, 70, 50, 40, '#2F9E6E'), Z('Sanitation chemical room', 10, 130, 50, 40, '#6E59A5'),
      Z('Finished goods cold storage', 120, 170, 140, 70, '#5B7FA6'), Z('Shipping dock', 10, 200, 90, 45, '#5B7FA6'),
    ],
  };
}

function defaultFlow(steps) {
  // Build a process flow from the step list with standard inspection / decision branches.
  const nodes = [], edges = [];
  const N = (kind, label, extra = {}) => { const n = Object.assign({ id: uid('n'), kind, label, x: 0, y: 0, dept: '' }, extra); nodes.push(n); return n; };
  const E = (a, b, label = '') => edges.push({ id: uid('e'), from: a.id, to: b.id, label });
  const main = [];
  const start = N('start', 'Ingredients arrive'); main.push(start);
  let prev = start, rejectNode = null, holdNode = null;
  steps.forEach((s) => {
    if (s.applic === 'Not applicable') return;
    const k = s.kind === 'storage' ? 'storage' : 'operation';
    const n = N(k, s.name, { stepId: s.id, dept: s.dept });
    E(prev, n); main.push(n); prev = n;
    if (/Receiving/i.test(s.name)) {
      const d = N('decision', 'Meets spec?', { dept: 'Quality', branch: true }); E(n, d); main.push(d);
      rejectNode = N('scrap', 'Reject / return to supplier', { dept: 'Quality', side: d.id }); E(d, rejectNode, 'No');
      prev = d; d._yes = true;
    }
    if (/Ingredient Storage/i.test(s.name)) { const t = N('transport', 'Transfer to pre-weigh', { dept: 'Warehouse' }); E(n, t); main.push(t); prev = t; }
    if (/Dividing|Forming/i.test(s.name)) {
      const i = N('inspection', 'Piece weight check', { dept: 'Production' }); E(n, i); main.push(i);
      const d = N('decision', 'In spec?', { dept: 'Production' }); E(i, d); main.push(d);
      const rw = N('rework', 'Dough rework (if spec allows)', { dept: 'Production', side: d.id }); E(d, rw, 'No');
      const mix = main.find((m) => /Mixing/i.test(m.label)); if (mix) E(rw, mix, 'Return');
      prev = d;
    }
    if (/Packaging/i.test(s.name)) {
      const i = N('inspection', 'Metal detect / checkweigh', { dept: 'Quality' }); E(n, i); main.push(i);
      const d = N('decision', 'Pass?', { dept: 'Quality' }); E(i, d); main.push(d);
      holdNode = N('delay', 'Hold for QA disposition', { dept: 'Quality', side: d.id }); E(d, holdNode, 'No');
      const sc = N('scrap', 'Scrap', { dept: 'Quality', side: holdNode.id }); E(holdNode, sc, 'Reject');
      prev = d;
    }
  });
  const end = N('end', 'Shipped to customer'); E(prev, end); main.push(end);
  // label "Yes" on decision-to-next edges
  edges.forEach((e) => { const a = nodes.find((n) => n.id === e.from); if (a && a.kind === 'decision' && !e.label) e.label = 'Yes'; });
  autoLayoutFlow({ nodes, edges }, main.map((m) => m.id));
  return { nodes, edges };
}

function autoLayoutFlow(flow, mainOrder) {
  const perRow = 6, dx = 190, dy = 230, x0 = 110, y0 = 80;
  const order = mainOrder || flow.nodes.filter((n) => !n.side).map((n) => n.id);
  order.forEach((id, i) => {
    const n = flow.nodes.find((m) => m.id === id); if (!n) return;
    const r = Math.floor(i / perRow), c = i % perRow;
    n.x = x0 + (r % 2 === 0 ? c : perRow - 1 - c) * dx; n.y = y0 + r * dy;
  });
  flow.nodes.filter((n) => n.side).forEach((n) => {
    const p = flow.nodes.find((m) => m.id === n.side);
    if (p) { n.x = p.x; n.y = p.y + 105; if (flow.nodes.some((o) => o !== n && o.side === p.id && o.y === n.y)) n.x += 150; }
  });
  // second-level side nodes (e.g. scrap under hold)
  flow.nodes.filter((n) => n.side && flow.nodes.find((m) => m.id === n.side && m.side)).forEach((n) => {
    const p = flow.nodes.find((m) => m.id === n.side); n.x = p.x + 160; n.y = p.y;
  });
}

function defaultActivities(steps) {
  const A = [];
  steps.forEach((s, i) => {
    if (s.applic === 'Not applicable') return;
    const add = (type, desc, va) => A.push({ id: uid('a'), stepId: s.id, desc, type, time: null, dist: null, people: null, va, status: 'NYC', src: '', opportunity: '' });
    if (i > 0) add('T', 'Move from previous step', 'NVA');
    add('D', 'Wait / queue before ' + s.name.toLowerCase(), 'NVA');
    if (s.kind === 'storage') add('S', 'Held in ' + s.name.toLowerCase(), 'NVA');
    else add('O', s.name, s.vaClass === 'VA' ? 'VA' : 'NNVA');
    if (s.fsq.qc || s.fsq.foodSafety) add('I', 'Quality / food-safety check at ' + s.name.toLowerCase(), 'NNVA');
  });
  return A;
}

function defaultSwimlane() {
  const items = [];
  const I = (lane, col, label, kind = 'task') => { const it = { id: uid('w'), lane, col, label, kind, flags: {}, next: [] }; items.push(it); return it; };
  const a = I('Planning', 0, 'Issue production schedule');
  const b = I('Supervisors', 1, 'Assign crew & line');
  const c = I('Warehouse', 1, 'Stage ingredients', 'handoff');
  const d = I('Sanitation', 2, 'Complete sanitation & pre-op clean');
  const e = I('Quality', 3, 'Pre-op inspection sign-off', 'approval');
  const f = I('Maintenance', 3, 'Start-up equipment checks');
  const g = I('Production', 4, 'Pre-weigh & allergen verification');
  const h = I('Quality', 5, 'Start-up / label approval', 'approval');
  const i = I('Production', 6, 'Run, freeze, pack');
  const j = I('Quality', 7, 'In-process checks & release', 'approval');
  const k = I('Warehouse', 8, 'Put away to cold storage', 'handoff');
  a.next = [b.id, c.id]; b.next = [d.id]; c.next = [g.id]; d.next = [e.id, f.id]; e.next = [g.id]; f.next = [g.id]; g.next = [h.id]; h.next = [i.id]; i.next = [j.id]; j.next = [k.id];
  return items;
}

function newState(kind) {
  const steps = genericSteps(kind);
  return { steps, flow: defaultFlow(steps), acts: defaultActivities(steps), layout: defaultLayout(), routes: [], swim: defaultSwimlane() };
}

function newLine(name, kind = 'generic') {
  return { id: uid('L'), name, unit: 'case', settings: defaultSettings(), current: newState(kind), future: null, notes: kind === 'nugget' ? 'Model line. Process sequence is the plant-wide starting template and must be verified on the floor walk.' : 'Template only. Steps copied from the plant-wide sequence; verify equipment and steps for this line.' };
}

function newProject() {
  const lines = [newLine('Nugget Line', 'nugget'), newLine('Rolled Dough Line #3'), newLine('Rolled Dough Line #4')];
  return { app: 'CHG Process Mapper', version: APP_VERSION, site: 'C.H. Guenther & Son — Duncanville, TX', demo: false, updated: new Date().toISOString(), lines, activeLine: lines[0].id, opps: [], actions: [], checks: {} };
}

/* ---------------- demonstration data ---------------- */
function applyDemo(project) {
  const line = project.lines[0];
  const D = (v) => F(v, 'EST', DEMO_SRC);
  const st = line.settings;
  Object.assign(st, { demandPerDay: D(5200), shiftsPerDay: D(2), minPerShift: D(480), breakMin: D(40), daysPerYear: D(250), laborRate: D(28), materialCost: D(9.5), forkliftSpeed: D(350) });
  line.unit = 'case';
  const vals = {
    'Raw Material Receiving': { cycleTime: 4.0, processTime: 0, labor: 2, wip: 9000, uptime: 95, transportDist: 0 },
    'Ingredient Storage': { cycleTime: 2.0, processTime: 0, labor: 1, wip: 14000, transportDist: 120 },
    'Pre-Weigh': { cycleTime: 5.1, processTime: 3, labor: 2, wip: 900, uptime: 92, changeover: 25, changeoversPerWeek: 6, transportDist: 180, scrap: 0.3 },
    'Mixing': { cycleTime: 5.4, processTime: 14, labor: 2, wip: 400, uptime: 88, changeover: 45, changeoversPerWeek: 6, transportDist: 40, scrap: 0.6, oee: 71 },
    'Dividing / Forming': { cycleTime: 5.9, processTime: 2, labor: 4, wip: 150, uptime: 82, changeover: 60, changeoversPerWeek: 6, scrap: 2.4, downtime: 38, oee: 64, throughput: 610 },
    'Spiral Freezing': { cycleTime: 5.0, processTime: 32, labor: 1, wip: 600, uptime: 94, scrap: 0.2, transportDist: 25 },
    'Packaging': { cycleTime: 6.3, processTime: 1, labor: 6, wip: 250, uptime: 80, changeover: 40, changeoversPerWeek: 6, scrap: 1.1, downtime: 45, oee: 61, throughput: 571 },
    'Cold Storage': { cycleTime: 2.5, processTime: 0, labor: 1, wip: 16000, transportDist: 260 },
    'Shipping': { cycleTime: 3.0, processTime: 0, labor: 2, wip: 0, transportDist: 150 },
  };
  line.current.steps.forEach((s) => {
    if (/Sheeting|Proofing|Baking|Cooling/.test(s.name)) { s.applic = 'Not applicable'; s.notes = 'DEMO assumption: not used on this line. Verify.'; }
    else s.applic = 'Confirmed';
    const v = vals[s.name] || {};
    Object.entries(v).forEach(([k, x]) => (s.f[k] = D(x)));
    s.owner = 'DEMO owner';
  });
  line.current.flow = defaultFlow(line.current.steps);
  line.current.acts = defaultActivities(line.current.steps);
  const actVals = { T: [3, 120, 1], D: [20, 0, 0], O: [8, 0, 2], I: [2, 0, 1], S: [600, 0, 0] };
  line.current.acts.forEach((a) => { const v = actVals[a.type]; a.time = v[0]; a.dist = v[1] || null; a.people = v[2] || null; a.status = 'EST'; a.src = DEMO_SRC; });
  const zone = (n) => line.current.layout.zones.find((z) => z.name === n);
  const ctr = (z) => [z.x + z.w / 2, z.y + z.h / 2];
  const R = (cat, name, pts, trips, people) => ({ id: uid('r'), cat, name, points: pts, trips: F(trips, 'EST', DEMO_SRC), people: F(people, 'EST', DEMO_SRC), status: 'EST' });
  line.current.routes = [
    R('ingredient', 'Dry ingredients: storage → pre-weigh', [ctr(zone('Dry ingredient storage')), [120, 70], [255, 70], ctr(zone('Pre-weigh room'))], 24, 1),
    R('walk', 'Mixer operator: mixing ↔ pre-weigh (fetch weighed kits)', [ctr(zone('Mixing')), [320, 70], [255, 70], ctr(zone('Pre-weigh room'))], 30, 1),
    R('walk', 'Forming operator: line ↔ QC lab (weight check forms)', [ctr(zone('Make-up / forming line')), [70, 110], ctr(zone('QC lab'))], 16, 1),
    R('forklift', 'Finished pallets: packaging → cold storage', [ctr(zone('Packaging')), [270, 205], ctr(zone('Finished goods cold storage'))], 40, 1),
    R('product', 'Packaging supplies: storage → packaging', [ctr(zone('Dry ingredient storage')), [120, 160], [300, 160], ctr(zone('Packaging'))], 12, 1),
    R('walk', 'Packaging lead: packaging ↔ maintenance shop', [ctr(zone('Packaging')), [385, 160], ctr(zone('Maintenance shop'))], 6, 1),
  ];
  line.current.swim.forEach((it) => {
    if (it.label === 'Pre-op inspection sign-off') it.flags = { duplicate: true };
    if (it.label === 'Start-up / label approval') it.flags = { unnecessary: true, delay: true };
    if (it.label === 'Stage ingredients') it.flags = { gap: true };
  });
  // future state
  const fut = clone(line.current);
  const fs = (n) => fut.steps.find((s) => s.name === n);
  fs('Dividing / Forming').f.changeover = D(35); fs('Dividing / Forming').f.uptime = D(88); fs('Dividing / Forming').f.scrap = D(1.6); fs('Dividing / Forming').f.cycleTime = D(5.5);
  fs('Packaging').f.changeover = D(25); fs('Packaging').f.uptime = D(87); fs('Packaging').f.cycleTime = D(5.8); fs('Packaging').f.labor = D(5);
  fs('Pre-Weigh').f.wip = D(450); fs('Mixing').f.wip = D(200);
  fut.routes = fut.routes.map((r) => {
    if (r.name.startsWith('Mixer operator')) return Object.assign(r, { name: 'Mixer operator: weighed kits delivered on cart (proposed)', trips: F(8, 'EST', DEMO_SRC) });
    if (r.name.startsWith('Forming operator')) return Object.assign(r, { name: 'Forming operator: weight check at line-side station (proposed)', points: [ctr(zone('Make-up / forming line')), [150, 135]], trips: F(16, 'EST', DEMO_SRC) });
    return r;
  });
  fut.swim.forEach((it) => { if (it.label === 'Start-up / label approval') { it.label = 'Start-up check (combined with pre-op)'; it.flags = {}; } });
  line.future = fut;
  const sid = (n) => line.current.steps.find((s) => s.name === n).id;
  const O = (stepName, waste, desc, est, cat, status, valid) => ({ id: uid('o'), lineId: line.id, stepId: sid(stepName), waste, desc, source: 'DEMO', priority: 'High', owner: 'DEMO owner', estAnnual: est, estBasis: 'DEMO calculation from demonstration inputs', fin: { category: cat, baseline: 'DEMO baseline', method: 'DEMO — to be agreed with Finance', status, validated: valid, reviewer: valid ? 'Finance (DEMO)' : '', date: valid ? todayISO() : '' } });
  project.opps = [
    O('Packaging', 'Waiting', 'Packaging is the bottleneck: frequent jams and slow changeover', 18500, 'Capacity', 'In verification', null),
    O('Dividing / Forming', 'Defects', 'Piece-weight scrap 2.4% at forming; tighten divider set-up', 31000, 'Hard savings', 'Method agreed with Finance', null),
    O('Mixing', 'Motion', 'Mixer operator walks to pre-weigh for every batch', 6100, 'Capacity', 'Not started', null),
    O('Pre-Weigh', 'Inventory', 'Weighed kits staged a full shift ahead of mixing', 0, 'Nonfinancial', 'Not started', null),
    O('Packaging', 'Over-processing', 'Separate start-up label approval duplicates the pre-op check', 2400, 'Capacity', 'Validated by Finance', 1800),
  ];
  const A = (oi, action, owner, due, status) => ({ id: uid('t'), oppId: project.opps[oi].id, action, owner, due, status, notes: 'DEMO' });
  const d = (n) => { const t = new Date(); t.setDate(t.getDate() + n); return t.toISOString().slice(0, 10); };
  project.actions = [
    A(0, 'Time 5 packaging changeovers and build SMED task list', 'DEMO supervisor', d(10), 'In progress'),
    A(0, 'Pareto of packaging jams from downtime log', 'DEMO maintenance', d(-3), 'Not started'),
    A(1, 'Agree scrap baseline and valuation method with Finance', 'DEMO CI', d(7), 'Done'),
    A(2, 'Trial kit cart for weighed ingredients', 'DEMO supervisor', d(21), 'Not started'),
    A(4, 'Combine start-up label approval into pre-op sign-off (Quality approval required)', 'DEMO quality', d(14), 'Blocked'),
  ];
  project.demo = true;
  return project;
}
