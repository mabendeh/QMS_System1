/* =====================================================================
   Events: delegated inputs, actions, diagram interaction
   ===================================================================== */
function setPath(obj, path, value) { const ks = path.split('.'); let o = obj; for (let i = 0; i < ks.length - 1; i++) { o[ks[i]] = o[ks[i]] || {}; o = o[ks[i]]; } o[ks[ks.length - 1]] = value; }
function coerce(el) { if (el.type === 'number') return el.value === '' ? null : +el.value; if (el.type === 'checkbox') return el.checked; return el.value; }
function findIn(list, id) { return (list || []).find((x) => x.id === id); }

function onInput(e) {
  const el = e.target; const S = ST(), l = LINE();
  const d = el.dataset; let rer = true;
  const split = (s) => s.split('|');
  if (d.field) { const [sid, k, prop] = split(d.field); const s = findIn(S.steps, sid); if (!s) return; const f = s.f[k]; const v = coerce(el); if (prop === 'v') { f.v = v; f.s = v == null ? 'NYC' : f.s === 'NYC' ? 'EST' : f.s; } else f[prop] = v; }
  else if (d.set) { const [k, prop] = split(d.set); const f = l.settings[k]; const v = coerce(el); if (prop === 'v') { f.v = v; f.s = v == null ? 'NYC' : f.s === 'NYC' ? 'EST' : f.s; } else f[prop] = v; }
  else if (d.step) { const [sid, prop] = split(d.step); const s = findIn(S.steps, sid); if (!s) return; setPath(s, prop, coerce(el)); if (prop === 'name') { S.flow.nodes.filter((n) => n.stepId === sid).forEach((n) => (n.label = s.name)); } if (prop.startsWith('fsq') || prop === 'notes') rer = false; }
  else if (d.act) { const [id, prop] = split(d.act); const a = findIn(S.acts, id); if (!a) return; a[prop] = coerce(el); if (prop === 'time') a.status = a.time == null ? 'NYC' : a.status === 'NYC' ? 'EST' : a.status; }
  else if (d.route) { const [id, prop] = split(d.route); const r = findIn(S.routes, id); if (!r) return; const v = coerce(el); if (prop.endsWith('.v')) { const fk = prop.split('.')[0]; r[fk].v = v; r[fk].s = v == null ? 'NYC' : r.status === 'VER' ? 'VER' : 'EST'; } else r[prop] = v; if (prop === 'status') ['trips', 'people'].forEach((fk) => { if (r[fk].v != null) r[fk].s = v; }); }
  else if (d.zone) { const [id, prop] = split(d.zone); const z = findIn(S.layout.zones, id); if (!z) return; z[prop] = coerce(el); }
  else if (d.layout) { S.layout[d.layout] = coerce(el); }
  else if (d.opp) { const [id, prop] = split(d.opp); const o = findIn(PRJ.opps, id); if (!o) return; setPath(o, prop, coerce(el)); if (prop === 'desc' || prop === 'estBasis' || prop === 'owner' || prop.startsWith('fin.baseline') || prop.startsWith('fin.method') || prop === 'fin.reviewer') rer = false; }
  else if (d.actrow) { const [id, prop] = split(d.actrow); const a = findIn(PRJ.actions, id); if (!a) return; a[prop] = coerce(el); if (prop === 'action' || prop === 'notes' || prop === 'owner') rer = false; }
  else if (d.node) { const [id, prop] = split(d.node); const n = findIn(S.flow.nodes, id); if (!n) return; n[prop] = coerce(el); }
  else if (d.edge) { const [id, prop] = split(d.edge); const ed = findIn(S.flow.edges, id); if (ed) ed[prop] = coerce(el); }
  else if (d.swim) { const [id, prop] = split(d.swim); const it = findIn(S.swim, id); if (!it) return; it[prop] = prop === 'col' ? Math.max(0, Math.round(coerce(el) || 0)) : coerce(el); }
  else if (d.swimflag) { const [id, k] = split(d.swimflag); const it = findIn(S.swim, id); if (!it) return; it.flags = it.flags || {}; it.flags[k] = el.checked; }
  else if (d.check) { PRJ.checks[d.check] = el.checked; if (!el.checked) delete PRJ.checks[d.check]; }
  else if (d.line) { l[d.line] = coerce(el); }
  else if (d.lineid) { const [id, prop] = split(d.lineid); const ln = findIn(PRJ.lines, id); if (ln) ln[prop] = coerce(el); }
  else if (el.id === 'lineSel') { PRJ.activeLine = el.value; UI.sel = null; UI.selType = null; if (!LINE().future && UI.state === 'future') UI.state = 'current'; }
  else if (el.id === 'stateSel') { UI.state = el.value; UI.sel = null; UI.selType = null; }
  else if (el.id === 'spMode') { UI.spMode = el.value; UI.draft = []; render(); return; }
  else if (el.id === 'drawCat') { UI.drawCat = el.value; return; }
  else if (el.id === 'actFilter') { UI.actFilter = el.value; render(); return; }
  else if (el.id === 'impXlsx') { if (el.files[0]) importXLSX(el.files[0]); return; }
  else if (el.id === 'impJson') { if (el.files[0]) importJSON(el.files[0]); return; }
  else return;
  changed(rer);
}

const ACTIONS = {
  createFuture() { const l = LINE(); l.future = clone(l.current); UI.state = 'future'; toast('Future state created from a copy of the current state'); changed(); },
  loadDemo() { const keepIds = PRJ.lines.slice(1); PRJ = applyDemo(newProject()); PRJ.lines = [PRJ.lines[0], ...keepIds]; PRJ.activeLine = PRJ.lines[0].id; UI.sel = null; toast('Demonstration data loaded — marked DEMO everywhere'); changed(); },
  clearDemo() { if (!confirmInline('clearDemo', 'Remove demonstration data and reset all lines to blank templates?')) return; PRJ = newProject(); UI.state = 'current'; UI.sel = null; toast('Reset to blank templates'); changed(); },
  zoom(btn) { const k = UI.tab, z = +btn.dataset.z; const cur = UI.zoom[k] || 0; UI.zoom[k] = z === 0 ? 0 : Math.max(0.25, Math.min(3, (cur || fitZoom(k)) * (z > 0 ? 1.25 : 0.8))); drawCanvas(); },
  flowAdd() { const S = ST(), kind = $('#addKind').value; const c = $('#canvas-flow'); const z = UI.zoom.flow || fitZoom('flow'); const n = { id: uid('n'), kind, label: FLOW_KINDS[kind], x: Math.round((c.scrollLeft + c.clientWidth / 2) / z), y: Math.round((c.scrollTop + c.clientHeight / 2) / z), dept: '' }; S.flow.nodes.push(n); UI.sel = n.id; UI.selType = 'node'; changed(); },
  flowConnect() { UI.connectFrom = UI.connectFrom === null ? '' : null; render(); },
  swConnect() { UI.connectFrom = UI.connectFrom === null ? '' : null; render(); },
  flowLayout() { autoLayoutFlow(ST().flow); changed(); },
  flowRebuild() { if (!confirmInline('flowRebuild', 'Rebuild the flow map from the step list? Manual changes to the map will be replaced.')) return; ST().flow = defaultFlow(ST().steps); changed(); },
  delSel() {
    const S = ST(), id = UI.sel; if (!id) return toast('Select something first');
    if (UI.selType === 'node') { S.flow.nodes = S.flow.nodes.filter((n) => n.id !== id); S.flow.edges = S.flow.edges.filter((e) => e.from !== id && e.to !== id); }
    else if (UI.selType === 'edge') S.flow.edges = S.flow.edges.filter((e) => e.id !== id);
    else if (UI.selType === 'swim') { S.swim = S.swim.filter((i) => i.id !== id); S.swim.forEach((i) => (i.next = (i.next || []).filter((n) => n !== id))); }
    else if (UI.selType === 'zone') S.layout.zones = S.layout.zones.filter((z) => z.id !== id);
    else if (UI.selType === 'route') S.routes = S.routes.filter((r) => r.id !== id);
    else return;
    UI.sel = null; UI.selType = null; changed();
  },
  closePanel() { UI.sel = null; UI.selType = null; render(); },
  selStep(btn) { UI.sel = btn.dataset.id; UI.selType = 'step'; render(); },
  matrix(btn) { UI.matrix = btn.dataset.v === '1'; render(); },
  moveStep(btn) { const S = ST(), i = S.steps.findIndex((s) => s.id === btn.dataset.id), j = i + +btn.dataset.d; if (j < 0 || j >= S.steps.length) return; [S.steps[i], S.steps[j]] = [S.steps[j], S.steps[i]]; changed(); },
  addStep(btn) { const S = ST(); const s = mkStep('New step', 'Production', { applic: 'To verify' }); const i = S.steps.findIndex((x) => x.id === btn.dataset.id); S.steps.splice(i < 0 ? S.steps.length : i + 1, 0, s); UI.sel = s.id; UI.selType = 'step'; toast('Step added. Use “Rebuild from steps” on the Process Flow tab to add it to the map.'); changed(); },
  delStep(btn) { if (!confirmInline('delStep' + btn.dataset.id, 'Delete this step and its activity rows?')) return; const S = ST(), id = btn.dataset.id; S.steps = S.steps.filter((s) => s.id !== id); S.acts = S.acts.filter((a) => a.stepId !== id); S.flow.nodes.forEach((n) => { if (n.stepId === id) n.stepId = ''; }); if (UI.sel === id) { UI.sel = null; UI.selType = null; } changed(); },
  addAct() { const S = ST(); S.acts.push({ id: uid('a'), stepId: (S.steps[0] || {}).id, desc: 'New activity', type: 'O', time: null, dist: null, people: null, va: 'NVA', status: 'NYC', src: '', opportunity: '' }); changed(); },
  delAct(btn) { ST().acts = ST().acts.filter((a) => a.id !== btn.dataset.id); changed(); },
  addZone() { const S = ST(); const z = { id: uid('z'), name: 'New zone', x: 20, y: 20, w: 40, h: 30, color: '#5B7FA6' }; S.layout.zones.push(z); UI.sel = z.id; UI.selType = 'zone'; changed(); },
  delZone(btn) { ST().layout.zones = ST().layout.zones.filter((z) => z.id !== btn.dataset.id); changed(); },
  delRoute(btn) { ST().routes = ST().routes.filter((r) => r.id !== btn.dataset.id); changed(); },
  finishRoute() { if (UI.draft.length < 2) return toast('Tap at least two points on the layout'); const S = ST(); const r = { id: uid('r'), cat: UI.drawCat, name: ROUTE_CATS[UI.drawCat].label + ' route ' + (S.routes.length + 1), points: UI.draft.map((p) => [Math.round(p[0]), Math.round(p[1])]), trips: F(), people: F(), status: 'EST' }; S.routes.push(r); UI.draft = []; UI.sel = r.id; UI.selType = 'route'; toast('Route added — enter trips per shift and people'); changed(); },
  cancelRoute() { UI.draft = []; render(); },
  swAdd() { const S = ST(); const lane = $('#swLane').value; const col = Math.max(0, ...S.swim.map((i) => i.col)) + 1; const it = { id: uid('w'), lane, col, label: 'New activity', kind: $('#swKind').value, flags: {}, next: [] }; S.swim.push(it); UI.sel = it.id; UI.selType = 'swim'; changed(); },
  swUnlink(btn) { const it = findIn(ST().swim, btn.dataset.a); if (it) it.next = it.next.filter((n) => n !== btn.dataset.b); changed(); },
  addOpp() { const l = LINE(); PRJ.opps.push({ id: uid('o'), lineId: l.id, stepId: '', waste: 'Waiting', desc: 'Describe the opportunity', source: 'Team', priority: 'Medium', owner: '', estAnnual: null, estBasis: '', fin: { category: 'Not yet classified', baseline: '', method: '', status: 'Not started', validated: null, reviewer: '', date: '' } }); changed(); },
  addSug(btn) { const l = LINE(); const s = suggestions(l)[+btn.dataset.i]; if (!s) return; PRJ.opps.push({ id: uid('o'), lineId: l.id, stepId: s.stepId, waste: s.waste, desc: s.desc, source: 'Suggested from data', priority: 'Medium', owner: '', estAnnual: null, estBasis: '', fin: { category: 'Not yet classified', baseline: '', method: '', status: 'Not started', validated: null, reviewer: '', date: '' } }); changed(); },
  delOpp(btn) { if (!confirmInline('delOpp' + btn.dataset.id, 'Delete this opportunity and its actions?')) return; PRJ.opps = PRJ.opps.filter((o) => o.id !== btn.dataset.id); PRJ.actions = PRJ.actions.filter((a) => a.oppId !== btn.dataset.id); changed(); },
  addAction() { const l = LINE(); const o = PRJ.opps.find((x) => x.lineId === l.id); if (!o) return; PRJ.actions.push({ id: uid('t'), oppId: o.id, action: 'Describe the action', owner: '', due: '', status: 'Not started', notes: '' }); changed(); },
  delAction(btn) { PRJ.actions = PRJ.actions.filter((a) => a.id !== btn.dataset.id); changed(); },
  addLine() { const name = ($('#newLineName').value || '').trim(); if (!name) return toast('Enter a line name first'); const l = newLine(name); PRJ.lines.push(l); PRJ.activeLine = l.id; UI.state = 'current'; toast(name + ' added from the plant template — verify its steps'); changed(); },
  dupLine(btn) { const src = findIn(PRJ.lines, btn.dataset.id); const l = clone(src); l.id = uid('L'); l.name = src.name + ' (copy)'; PRJ.lines.push(l); changed(); },
  delLine(btn) { if (!confirmInline('delLine' + btn.dataset.id, 'Delete this line and its opportunities?')) return; PRJ.lines = PRJ.lines.filter((l) => l.id !== btn.dataset.id); const keep = PRJ.opps.filter((o) => o.lineId !== btn.dataset.id); PRJ.actions = PRJ.actions.filter((a) => keep.some((o) => o.id === a.oppId)); PRJ.opps = keep; PRJ.activeLine = PRJ.lines[0].id; changed(); },
  exportXLSX: () => exportXLSX(), exportPPTX: () => exportPPTX(), exportPDF: () => exportPDF(false), exportCompare: () => exportPDF(true), exportJSON: () => exportJSON(),
  exportSVG(btn) { exportDiagram(btn.dataset.d, 'svg'); }, exportPNG(btn) { exportDiagram(btn.dataset.d, 'png'); },
};
// In-page confirmation (the artifact viewer suppresses confirm()): first click arms, second click within 5 s confirms.
function confirmInline(key, msg) { if (UI.confirm === key) { UI.confirm = null; return true; } UI.confirm = key; toast(msg + ' Tap again to confirm.', 'warn'); setTimeout(() => { if (UI.confirm === key) UI.confirm = null; }, 5000); return false; }

function onClick(e) {
  const t = e.target.closest('[data-tab],[data-action]'); if (!t) return;
  if (t.dataset.tab && !t.dataset.action) { UI.tab = t.dataset.tab; UI.connectFrom = null; UI.draft = []; try { localStorage.setItem('chgpm.tab', UI.tab); } catch (_) { } render(); window.scrollTo(0, 0); return; }
  const fn = ACTIONS[t.dataset.action]; if (fn) { e.preventDefault(); fn(t); }
}

/* ---------------- diagram canvas ---------------- */
function currentDiagram(kind, P = PAL()) {
  const l = LINE(), opts = { selected: UI.sel, connectFrom: UI.connectFrom, demo: PRJ.demo, draft: UI.draft, opps: PRJ.opps.filter((o) => o.lineId === l.id) };
  return { flow: svgFlow, vsm: svgVSM, spaghetti: svgSpaghetti, swimlane: svgSwim }[kind](l, UI.state, P, opts);
}
function fitZoom(kind) { const c = $('#canvas-' + kind); if (!c) return 1; const d = currentDiagram(kind); const fit = (c.clientWidth - 4) / d.w; return Math.max(kind === 'vsm' || kind === 'swimlane' ? Math.min(0.6, 1) : 0.25, Math.min(1.4, fit)); }
function drawCanvas() {
  const kind = { flow: 'flow', vsm: 'vsm', spaghetti: 'spaghetti', swimlane: 'swimlane' }[UI.tab]; if (!kind) return;
  const c = $('#canvas-' + kind); if (!c) return;
  const d = currentDiagram(kind); const z = UI.zoom[kind] || fitZoom(kind);
  const sl = c.scrollLeft, stp = c.scrollTop;
  c.innerHTML = d.svg; const svg = c.querySelector('svg'); svg.setAttribute('width', Math.round(d.w * z)); svg.setAttribute('height', Math.round(d.h * z));
  c.scrollLeft = sl; c.scrollTop = stp;
  c.classList.toggle('drawing', kind === 'spaghetti' && UI.spMode === 'draw');
}
function svgPoint(svg, evt) { const pt = svg.createSVGPoint(); pt.x = evt.clientX; pt.y = evt.clientY; return pt.matrixTransform(svg.getScreenCTM().inverse()); }

function bindCanvas() {
  const c = $('.canvas'); if (!c) return;
  const kind = c.dataset.kind;
  let drag = null;
  c.onpointerdown = (e) => {
    const svg = c.querySelector('svg'); if (!svg) return;
    const g = e.target.closest('[data-id]'); const S = ST();
    if (kind === 'spaghetti' && UI.spMode === 'draw') { const p = svgPoint(svg, e), d = currentDiagram('spaghetti'); const x = (p.x - d.pad) / d.k, y = (p.y - d.pad) / d.k; if (x < 0 || y < 0 || x > S.layout.width || y > S.layout.height) return; UI.draft.push([Math.round(x), Math.round(y)]); drawCanvas(); const fb = $('[data-action="finishRoute"]'); if (fb) fb.textContent = `Finish route (${UI.draft.length} pts)`; return; }
    if (!g) { if (UI.sel) { UI.sel = null; UI.selType = null; drawCanvas(); renderPanel(); } return; }
    const id = g.dataset.id;
    if (g.classList.contains('fnode') && UI.connectFrom !== null) {
      if (!UI.connectFrom) { UI.connectFrom = id; drawCanvas(); toast('Now tap the target block'); return; }
      if (UI.connectFrom !== id) { S.flow.edges.push({ id: uid('e'), from: UI.connectFrom, to: id, label: '' }); UI.connectFrom = null; changed(); return; }
    }
    if (g.classList.contains('swim') && UI.connectFrom !== null) {
      if (!UI.connectFrom) { UI.connectFrom = id; toast('Now tap the next activity'); return; }
      if (UI.connectFrom !== id) { const a = findIn(S.swim, UI.connectFrom); if (a && !a.next.includes(id)) a.next.push(id); UI.connectFrom = null; changed(); return; }
    }
    const type = g.classList.contains('fnode') ? 'node' : g.classList.contains('fedge') ? 'edge' : g.classList.contains('vstep') ? 'step' : g.classList.contains('zone') ? 'zone' : g.classList.contains('route') ? 'route' : g.classList.contains('swim') ? 'swim' : null;
    UI.sel = id; UI.selType = type;
    if (['node', 'zone', 'swim'].includes(type)) {
      const p = svgPoint(svg, e);
      const obj = type === 'node' ? findIn(S.flow.nodes, id) : type === 'zone' ? findIn(S.layout.zones, id) : findIn(S.swim, id);
      drag = { type, obj, p0: p, o0: { x: obj.x, y: obj.y, lane: obj.lane, col: obj.col }, moved: false };
      c.setPointerCapture(e.pointerId); e.preventDefault();
    }
    drawCanvas(); renderPanel();
  };
  c.onpointermove = (e) => {
    if (!drag) return; const svg = c.querySelector('svg'); const p = svgPoint(svg, e); const dx = p.x - drag.p0.x, dy = p.y - drag.p0.y;
    if (Math.abs(dx) + Math.abs(dy) > 3) drag.moved = true; if (!drag.moved) return;
    if (drag.type === 'node') { drag.obj.x = Math.round((drag.o0.x + dx) / 10) * 10; drag.obj.y = Math.round((drag.o0.y + dy) / 10) * 10; }
    else if (drag.type === 'zone') { const d = currentDiagram('spaghetti'); drag.obj.x = Math.max(0, Math.round(drag.o0.x + dx / d.k)); drag.obj.y = Math.max(0, Math.round(drag.o0.y + dy / d.k)); }
    else if (drag.type === 'swim') { const d = currentDiagram('swimlane'); drag.obj.col = Math.max(0, Math.round(drag.o0.col + dx / d.colW)); const li = Math.max(0, Math.min(DEPTS.length - 1, DEPTS.indexOf(drag.o0.lane) + Math.round(dy / d.laneH))); drag.obj.lane = DEPTS[li]; }
    drawCanvas();
  };
  c.onpointerup = () => { if (drag && drag.moved) changed(false); if (drag && drag.moved && UI.tab === 'spaghetti') render(); drag = null; };
  c.ondblclick = () => { if (kind === 'spaghetti' && UI.spMode === 'draw' && UI.draft.length >= 2) { UI.draft.pop(); ACTIONS.finishRoute(); } };
}
function afterRender() { drawCanvas(); bindCanvas(); }

function boot() {
  try { const t = localStorage.getItem('chgpm.tab'); if (t && TABS.some((x) => x[0] === t)) UI.tab = t; } catch (_) { }
  PRJ = Store.loadLocal() || newProject();
  document.addEventListener('click', onClick);
  document.addEventListener('change', onInput);
  window.addEventListener('keydown', (e) => { if (e.key === 'Escape') { UI.connectFrom = null; UI.draft = []; render(); } });
  if (window.matchMedia) matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => afterRender());
  new MutationObserver(() => afterRender()).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
  let rt; window.addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(drawCanvas, 150); });
  render();
  Store.init();
}
