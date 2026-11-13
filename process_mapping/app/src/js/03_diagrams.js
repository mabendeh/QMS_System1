/* =====================================================================
   Diagram renderers. Each returns {svg, w, h}. Colors come from a
   palette object so exports can always use the light print palette.
   ===================================================================== */

const PALETTES = {
  light: { bg: '#FFFFFF', grid: '#E7ECF2', grid2: '#D3DBE5', ink: '#14202E', muted: '#5E6B7A', line: '#46566A', node: '#FFFFFF', stroke: '#13315C', accent: '#13315C', green: '#2F9E6E', amber: '#B9770E', red: '#C0392B', soft: '#EEF2F6', soft2: '#F7F9FB', sel: '#2F6FB0' },
  dark: { bg: '#0F1822', grid: '#1A2633', grid2: '#243344', ink: '#E6EDF5', muted: '#93A3B5', line: '#8FA3B8', node: '#162231', stroke: '#8EB4E3', accent: '#8EB4E3', green: '#4CC48F', amber: '#E3A640', red: '#E86A5A', soft: '#1B2836', soft2: '#142030', sel: '#6FA8E8' },
};

function wrapText(t, max) {
  const words = String(t || '').split(/\s+/); const lines = []; let cur = '';
  words.forEach((w) => { if ((cur + ' ' + w).trim().length > max && cur) { lines.push(cur); cur = w; } else cur = (cur + ' ' + w).trim(); });
  if (cur) lines.push(cur); return lines.slice(0, 3);
}
function textLines(x, y, lines, opts = {}) {
  const lh = opts.lh || 14, fs = opts.fs || 12, fill = opts.fill, anchor = opts.anchor || 'middle', weight = opts.weight || 400;
  const y0 = y - ((lines.length - 1) * lh) / 2;
  return `<text x="${x}" text-anchor="${anchor}" font-size="${fs}" font-weight="${weight}" fill="${fill}" font-family="IBM Plex Sans, Arial, sans-serif">` +
    lines.map((l, i) => `<tspan x="${x}" y="${(y0 + i * lh + fs * 0.35).toFixed(1)}">${esc(l)}</tspan>`).join('') + '</text>';
}
function arrowDefs(P, extra = '') {
  return `<defs><marker id="ah" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="${P.line}"/></marker>` +
    `<marker id="ahs" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="${P.sel}"/></marker>${extra}</defs>`;
}
function svgWrap(w, h, body, P, title) {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" font-family="IBM Plex Sans, Arial, sans-serif"><title>${esc(title)}</title><rect width="${w}" height="${h}" fill="${P.bg}"/>${body}</svg>`;
}
function demoWatermark(P, w, h, on) { return on ? `<text x="${w - 16}" y="${h - 14}" text-anchor="end" font-size="13" font-weight="700" fill="${P.amber}" letter-spacing="1.5">DEMONSTRATION DATA — NOT PLANT MEASUREMENTS</text>` : ''; }

/* ---------------- Process flow ---------------- */
const NODE_SIZE = { start: [150, 44], end: [150, 44], operation: [150, 58], decision: [130, 86], inspection: [130, 58], transport: [140, 52], delay: [130, 58], storage: [110, 84], rework: [150, 58], scrap: [140, 52] };
function nodeBox(n) { const [w, h] = NODE_SIZE[n.kind] || [150, 58]; return { x: n.x - w / 2, y: n.y - h / 2, w, h }; }
function clipToBox(cx, cy, tx, ty, b, kind) {
  const dx = tx - cx, dy = ty - cy; if (!dx && !dy) return [cx, cy];
  let hw = b.w / 2, hh = b.h / 2;
  if (kind === 'decision') { const t = 1 / (Math.abs(dx) / hw + Math.abs(dy) / hh); return [cx + dx * t, cy + dy * t]; }
  const t = Math.min(hw / Math.abs(dx || 1e-9), hh / Math.abs(dy || 1e-9)); return [cx + dx * t, cy + dy * t];
}
function nodeShape(n, P, selected) {
  const b = nodeBox(n), s = selected ? P.sel : n.kind === 'scrap' ? P.red : n.kind === 'rework' ? P.amber : P.stroke;
  const sw = selected ? 3 : 1.6, fill = n.kind === 'start' || n.kind === 'end' ? P.soft : P.node;
  const dash = n.kind === 'rework' ? ' stroke-dasharray="6 4"' : '';
  let shape = '', lines = wrapText(n.label, n.kind === 'decision' ? 14 : 18), ty = n.y;
  switch (n.kind) {
    case 'start': case 'end': shape = `<rect x="${b.x}" y="${b.y}" width="${b.w}" height="${b.h}" rx="${b.h / 2}" fill="${fill}" stroke="${s}" stroke-width="${sw}"/>`; break;
    case 'decision': shape = `<polygon points="${n.x},${b.y} ${b.x + b.w},${n.y} ${n.x},${b.y + b.h} ${b.x},${n.y}" fill="${fill}" stroke="${s}" stroke-width="${sw}"/>`; break;
    case 'inspection': shape = `<rect x="${b.x}" y="${b.y}" width="${b.w}" height="${b.h}" fill="${fill}" stroke="${s}" stroke-width="${sw}"/><rect x="${b.x + 4}" y="${b.y + 4}" width="${b.w - 8}" height="${b.h - 8}" fill="none" stroke="${s}" stroke-width="1"/><path d="M${b.x + b.w - 20},${b.y + 14} l4,4 l8,-9" fill="none" stroke="${P.green}" stroke-width="2"/>`; break;
    case 'transport': shape = `<polygon points="${b.x},${b.y} ${b.x + b.w - 18},${b.y} ${b.x + b.w},${n.y} ${b.x + b.w - 18},${b.y + b.h} ${b.x},${b.y + b.h}" fill="${fill}" stroke="${s}" stroke-width="${sw}"/>`; break;
    case 'delay': shape = `<path d="M${b.x},${b.y} H${b.x + b.w - b.h / 2} A${b.h / 2},${b.h / 2} 0 0 1 ${b.x + b.w - b.h / 2},${b.y + b.h} H${b.x} Z" fill="${fill}" stroke="${s}" stroke-width="${sw}"/>`; break;
    case 'storage': shape = `<polygon points="${n.x - 30},${b.y} ${n.x + 30},${b.y} ${n.x},${b.y + 46}" fill="${fill}" stroke="${s}" stroke-width="${sw}"/>`; ty = b.y + 66; lines = wrapText(n.label, 20).slice(0, 2); break;
    case 'scrap': shape = `<rect x="${b.x}" y="${b.y}" width="${b.w}" height="${b.h}" rx="4" fill="${fill}" stroke="${s}" stroke-width="${sw}"/><path d="M${b.x + 8},${b.y + 8} l8,8 m0,-8 l-8,8" stroke="${P.red}" stroke-width="2"/>`; break;
    default: shape = `<rect x="${b.x}" y="${b.y}" width="${b.w}" height="${b.h}" rx="4" fill="${fill}" stroke="${s}" stroke-width="${sw}"${dash}/>`;
  }
  const dot = n.dept ? `<circle cx="${n.kind === 'storage' ? n.x + 34 : b.x + 10}" cy="${n.kind === 'storage' ? b.y + 8 : b.y + 10}" r="5" fill="${DEPT_COLORS[n.dept] || P.muted}"><title>${esc(n.dept)}</title></circle>` : '';
  return `<g class="fnode" data-id="${n.id}" style="cursor:move">${shape}${dot}${textLines(n.x, ty, lines, { fill: P.ink, fs: 12, lh: 14 })}</g>`;
}
function svgFlow(line, key, P, opts = {}) {
  const S = lineState(line, key), fl = S.flow;
  let maxX = 600, maxY = 400;
  fl.nodes.forEach((n) => { const b = nodeBox(n); maxX = Math.max(maxX, b.x + b.w + 60); maxY = Math.max(maxY, b.y + b.h + 90); });
  let body = arrowDefs(P);
  // grid
  body += `<g opacity="0.7">`; for (let x = 0; x < maxX; x += 40) body += `<line x1="${x}" y1="0" x2="${x}" y2="${maxY}" stroke="${P.grid}" stroke-width="1"/>`; for (let y = 0; y < maxY; y += 40) body += `<line x1="0" y1="${y}" x2="${maxX}" y2="${y}" stroke="${P.grid}" stroke-width="1"/>`; body += `</g>`;
  fl.edges.forEach((e) => {
    const a = fl.nodes.find((n) => n.id === e.from), b = fl.nodes.find((n) => n.id === e.to); if (!a || !b) return;
    const [x1, y1] = clipToBox(a.x, a.y, b.x, b.y, nodeBox(a), a.kind), [x2, y2] = clipToBox(b.x, b.y, a.x, a.y, nodeBox(b), b.kind);
    const sel = opts.selected === e.id;
    body += `<g class="fedge" data-id="${e.id}" style="cursor:pointer"><line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="transparent" stroke-width="14"/><line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" stroke="${sel ? P.sel : P.line}" stroke-width="${sel ? 2.6 : 1.6}" marker-end="url(#${sel ? 'ahs' : 'ah'})"/>`;
    if (e.label) { const mx = (x1 + x2) / 2, my = (y1 + y2) / 2, w = e.label.length * 6.5 + 10; body += `<rect x="${mx - w / 2}" y="${my - 9}" width="${w}" height="18" rx="9" fill="${P.bg}" stroke="${P.grid2}"/><text x="${mx}" y="${my + 4}" text-anchor="middle" font-size="11" fill="${P.muted}">${esc(e.label)}</text>`; }
    body += `</g>`;
  });
  fl.nodes.forEach((n) => (body += nodeShape(n, P, opts.selected === n.id || opts.connectFrom === n.id)));
  body += demoWatermark(P, maxX, maxY, opts.demo);
  return { svg: svgWrap(maxX, maxY, body, P, `${line.name} — process flow (${key} state)`), w: maxX, h: maxY };
}

/* ---------------- Value stream map ---------------- */
function vsmVal(f, unit, d = 1) { const v = val(f); return v == null ? 'NYC' : fmt(v, d) + (unit ? ' ' + unit : ''); }
function svgVSM(line, key, P, opts = {}) {
  const S = lineState(line, key), calc = calcState(line, key), steps = applicable(S.steps);
  const colW = 220, x0 = 120, boxW = 150, boxY = 250, boxH = 66, dataH = 7 * 17 + 10;
  const n = steps.length, w = x0 + n * colW + 300, h = 640;
  let b = arrowDefs(P);
  const T = (x, y, t, o = {}) => `<text x="${x}" y="${y}" text-anchor="${o.a || 'middle'}" font-size="${o.fs || 11}" font-weight="${o.w || 400}" fill="${o.c || P.ink}">${esc(t)}</text>`;
  // supplier / customer / production control
  const factory = (x, y, label, sub) => `<path d="M${x},${y + 50} V${y + 14} l20,-14 v14 l20,-14 v14 l20,-14 v14 h20 v50 z" fill="${P.soft}" stroke="${P.stroke}" stroke-width="1.5"/>${T(x + 50, y + 66, label, { w: 600 })}${sub ? T(x + 50, y + 80, sub, { c: P.muted, fs: 10 }) : ''}`;
  b += factory(20, 30, 'Suppliers', 'Ingredients & packaging');
  const cx = x0 + n * colW + 40;
  const dem = val(line.settings.demandPerDay);
  b += factory(cx, 30, 'Customers', dem == null ? 'Demand: NYC' : `Demand: ${fmt(dem, 0)} ${line.unit}/day`);
  const pcx = x0 + (n * colW) / 2 - 90;
  b += `<rect x="${pcx}" y="28" width="180" height="56" fill="${P.node}" stroke="${P.stroke}" stroke-width="1.5"/>${T(pcx + 90, 52, 'Production Control', { w: 600 })}${T(pcx + 90, 68, 'Planning / scheduling', { c: P.muted, fs: 10 })}`;
  b += `<line x1="${cx}" y1="56" x2="${pcx + 180}" y2="56" stroke="${P.line}" stroke-dasharray="6 4" marker-end="url(#ah)"/>${T((cx + pcx + 180) / 2, 50, 'Customer orders', { c: P.muted, fs: 10 })}`;
  b += `<line x1="${pcx}" y1="56" x2="120" y2="56" stroke="${P.line}" stroke-dasharray="6 4" marker-end="url(#ah)"/>${T((pcx + 120) / 2, 50, 'Purchase orders', { c: P.muted, fs: 10 })}`;
  // schedule bus
  if (n) { const bx1 = x0 + boxW / 2, bx2 = x0 + (n - 1) * colW + boxW / 2; b += `<line x1="${pcx + 90}" y1="84" x2="${pcx + 90}" y2="200" stroke="${P.line}" stroke-dasharray="6 4"/><line x1="${bx1}" y1="200" x2="${bx2}" y2="200" stroke="${P.line}" stroke-dasharray="6 4"/>${T(pcx + 96, 194, 'Production schedule (method to verify)', { a: 'start', c: P.muted, fs: 10 })}`; steps.forEach((s, i) => (b += `<line x1="${x0 + i * colW + boxW / 2}" y1="200" x2="${x0 + i * colW + boxW / 2}" y2="${boxY - 2}" stroke="${P.line}" stroke-dasharray="6 4" marker-end="url(#ah)"/>`)); }
  // supplier → first, last → customer
  b += `<line x1="70" y1="112" x2="70" y2="${boxY + 30}" stroke="${P.line}" stroke-width="2"/><line x1="70" y1="${boxY + 30}" x2="${x0 - 52}" y2="${boxY + 30}" stroke="${P.line}" stroke-width="2" marker-end="url(#ah)"/>`;
  if (n) b += `<line x1="${x0 + (n - 1) * colW + boxW}" y1="${boxY + 30}" x2="${cx + 50}" y2="${boxY + 30}" stroke="${P.line}" stroke-width="2"/><line x1="${cx + 50}" y1="${boxY + 30}" x2="${cx + 50}" y2="114" stroke="${P.line}" stroke-width="2" marker-end="url(#ah)"/>`;
  const tlY = boxY + boxH + dataH + 60, hi = tlY, lo = tlY + 40;
  let ladder = `M${x0 - 70},${hi}`;
  const bnId = calc.bottleneck && calc.bottleneck.stepId;
  steps.forEach((s, i) => {
    const x = x0 + i * colW, r = calc.steps[s.id] || {}, isBn = s.id === bnId;
    // inventory triangle
    const tx = x - 40, wip = val(s.f.wip);
    b += `<polygon points="${tx - 16},${boxY + 46} ${tx + 16},${boxY + 46} ${tx},${boxY + 18}" fill="${P.soft}" stroke="${P.amber}" stroke-width="1.5"/>${T(tx, boxY + 40, 'I', { w: 700, c: P.amber, fs: 10 })}`;
    b += T(tx, boxY + 62, wip == null ? 'WIP NYC' : fmt(wip, 0) + ' ' + line.unit, { fs: 10, c: wip == null ? P.amber : P.ink });
    // push arrow from previous box
    if (i > 0) b += `<line x1="${x - colW + boxW + 2}" y1="${boxY + 30}" x2="${tx - 18}" y2="${boxY + 30}" stroke="${P.line}" stroke-width="5" stroke-dasharray="4 3" marker-end="url(#ah)"/>`;
    b += `<line x1="${tx + 18}" y1="${boxY + 30}" x2="${x - 2}" y2="${boxY + 30}" stroke="${P.line}" stroke-width="1.5" marker-end="url(#ah)"/>`;
    // process box
    b += `<g class="vstep" data-id="${s.id}" style="cursor:pointer"><rect x="${x}" y="${boxY}" width="${boxW}" height="${boxH}" fill="${P.node}" stroke="${isBn ? P.red : P.stroke}" stroke-width="${isBn ? 3 : 1.6}"/>`;
    b += textLines(x + boxW / 2, boxY + 26, wrapText(s.name, 20).slice(0, 2), { fill: P.ink, fs: 12, weight: 600 });
    const lab = val(s.f.labor); b += T(x + boxW / 2, boxY + 56, lab == null ? 'Operators: NYC' : `Operators: ${fmt(lab, 1)}`, { fs: 10, c: lab == null ? P.amber : P.muted });
    if (isBn) b += `<rect x="${x + boxW - 76}" y="${boxY - 18}" width="76" height="16" rx="8" fill="${P.red}"/>${T(x + boxW - 38, boxY - 6, 'BOTTLENECK', { fs: 9, w: 700, c: '#FFFFFF' })}`;
    b += `</g>`;
    // data box
    const rows = [['C/T', vsmVal(s.f.cycleTime, 's')], ['C/O', vsmVal(s.f.changeover, 'min')], ['Uptime', vsmVal(s.f.uptime, '%')], ['OEE', vsmVal(s.f.oee, '%')], ['Scrap', vsmVal(s.f.scrap, '%')], ['Eff. C/T', r.ect && r.ect.v != null ? fmt(r.ect.v, 2) + ' s' : 'NYC'], ['P/T', vsmVal(s.f.processTime, 'min')]];
    b += `<rect x="${x}" y="${boxY + boxH}" width="${boxW}" height="${dataH}" fill="${P.soft2}" stroke="${P.stroke}" stroke-width="1"/>`;
    rows.forEach((rw, k) => { const yy = boxY + boxH + 17 + k * 17; b += T(x + 8, yy, rw[0], { a: 'start', c: P.muted }) + T(x + boxW - 8, yy, rw[1], { a: 'end', c: rw[1] === 'NYC' ? P.amber : P.ink, w: rw[1] === 'NYC' ? 600 : 400 }); });
    // kaizen bursts
    const opps = (opts.opps || []).filter((o) => o.stepId === s.id);
    if (opps.length && key === 'future') { const kx = x + boxW - 6, ky = boxY - 34; const pts = []; for (let k = 0; k < 16; k++) { const rr = k % 2 ? 9 : 17, a = (k * Math.PI) / 8; pts.push(`${(kx + rr * Math.cos(a)).toFixed(1)},${(ky + rr * Math.sin(a)).toFixed(1)}`); } b += `<polygon points="${pts.join(' ')}" fill="${P.amber}" stroke="${P.stroke}" stroke-width="0.8"><title>${esc(opps.map((o) => o.desc).join(' | '))}</title></polygon>${T(kx, ky + 4, String(opps.length), { fs: 11, w: 700, c: '#FFFFFF' })}`; }
    // timeline
    const q = r.queue, pt = val(s.f.processTime);
    ladder += ` H${x - 2} V${lo} H${x + boxW} V${hi}`;
    b += T(x - 40, hi - 8, q && q.v != null ? fmt(q.v, 0) + ' min' : 'NYC', { fs: 10, c: q && q.v != null ? P.ink : P.amber, w: 600 });
    b += T(x + boxW / 2, lo + 16, pt == null ? 'NYC' : fmt(pt, 1) + ' min', { fs: 10, c: pt == null ? P.amber : s.vaClass === 'VA' ? P.green : P.muted, w: 600 });
  });
  b += `<path d="${ladder}" fill="none" stroke="${P.stroke}" stroke-width="1.6"/>`;
  b += T(x0 - 80, hi - 22, 'Queue (wait) time', { a: 'start', c: P.muted, fs: 10 }) + T(x0 - 80, lo + 34, 'Processing time', { a: 'start', c: P.muted, fs: 10 });
  // summary
  const sx = cx - 10, sy = tlY - 40;
  const L = (lbl, r, f) => [lbl, r && r.v != null ? f(r.v) + (r.s === 'NYC' ? ' (incomplete)' : r.s === 'EST' ? ' (est.)' : '') : 'NYC'];
  const sum = [L('Lead time', calc.leadTime, (v) => fmt(v / 60, 1) + ' hr'), L('Value-added', calc.vaTime, (v) => fmt(v, 1) + ' min'), L('PCE', calc.pce, (v) => pct(v, 2)), L('Takt', calc.takt, (v) => fmt(v, 2) + ' s'), L('Bottleneck C/T', calc.bnECT, (v) => fmt(v, 2) + ' s')];
  b += `<rect x="${sx}" y="${sy}" width="250" height="${sum.length * 22 + 34}" fill="${P.node}" stroke="${P.stroke}" stroke-width="1.6"/>${T(sx + 12, sy + 20, key === 'future' ? 'FUTURE STATE TOTALS' : 'CURRENT STATE TOTALS', { a: 'start', w: 700, fs: 11, c: P.accent })}`;
  sum.forEach((r, k) => (b += T(sx + 12, sy + 42 + k * 22, r[0], { a: 'start', c: P.muted }) + T(sx + 238, sy + 42 + k * 22, r[1], { a: 'end', w: 600, c: r[1] === 'NYC' ? P.amber : P.ink })));
  b += T(x0 - 80, h - 24, 'NYC = Not Yet Collected. Values marked (est.) include estimates; (incomplete) means some inputs are missing.', { a: 'start', c: P.muted, fs: 10 });
  b += demoWatermark(P, w, h, opts.demo);
  return { svg: svgWrap(w, h, b, P, `${line.name} — value stream map (${key} state)`), w, h };
}

/* ---------------- Spaghetti diagram ---------------- */
function svgSpaghetti(line, key, P, opts = {}) {
  const S = lineState(line, key), L = S.layout, k = 3, pad = 40;
  const w = L.width * k + pad * 2, h = L.height * k + pad * 2 + 70;
  const X = (x) => pad + x * k, Y = (y) => pad + y * k;
  let markers = '';
  Object.entries(ROUTE_CATS).forEach(([c, d]) => (markers += `<marker id="m-${c}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="${d.color}"/></marker>`));
  let b = arrowDefs(P, markers);
  for (let x = 0; x <= L.width; x += 10) b += `<line x1="${X(x)}" y1="${Y(0)}" x2="${X(x)}" y2="${Y(L.height)}" stroke="${x % 50 ? P.grid : P.grid2}"/>`;
  for (let y = 0; y <= L.height; y += 10) b += `<line x1="${X(0)}" y1="${Y(y)}" x2="${X(L.width)}" y2="${Y(y)}" stroke="${y % 50 ? P.grid : P.grid2}"/>`;
  b += `<rect x="${X(0)}" y="${Y(0)}" width="${L.width * k}" height="${L.height * k}" fill="none" stroke="${P.line}" stroke-width="2"/>`;
  L.zones.forEach((z) => {
    const sel = opts.selected === z.id;
    b += `<g class="zone" data-id="${z.id}" style="cursor:move"><rect x="${X(z.x)}" y="${Y(z.y)}" width="${z.w * k}" height="${z.h * k}" fill="${z.color}" fill-opacity="0.12" stroke="${sel ? P.sel : z.color}" stroke-width="${sel ? 3 : 1.4}"/>` +
      textLines(X(z.x + z.w / 2), Y(z.y + z.h / 2), wrapText(z.name, Math.max(8, Math.floor(z.w / 3.2))), { fill: P.ink, fs: 11, lh: 13, weight: 600 }) + `</g>`;
  });
  const calc = calcState(line, key);
  (S.routes || []).forEach((r, i) => {
    const c = ROUTE_CATS[r.cat] || ROUTE_CATS.walk, sel = opts.selected === r.id;
    const pts = r.points.map((p) => `${X(p[0]).toFixed(1)},${Y(p[1]).toFixed(1)}`).join(' ');
    b += `<g class="route" data-id="${r.id}" style="cursor:pointer"><polyline points="${pts}" fill="none" stroke="transparent" stroke-width="14"/><polyline points="${pts}" fill="none" stroke="${c.color}" stroke-width="${sel ? 5 : 3}" stroke-opacity="${sel ? 1 : 0.85}" stroke-linejoin="round" marker-end="url(#m-${r.cat})"/>`;
    const p0 = r.points[0]; b += `<circle cx="${X(p0[0])}" cy="${Y(p0[1])}" r="9" fill="${c.color}"/><text x="${X(p0[0])}" y="${Y(p0[1]) + 4}" text-anchor="middle" font-size="10" font-weight="700" fill="#FFFFFF">${i + 1}</text></g>`;
  });
  if (opts.draft && opts.draft.length) b += `<polyline points="${opts.draft.map((p) => X(p[0]) + ',' + Y(p[1])).join(' ')}" fill="none" stroke="${P.sel}" stroke-width="3" stroke-dasharray="6 4"/>` + opts.draft.map((p) => `<circle cx="${X(p[0])}" cy="${Y(p[1])}" r="4" fill="${P.sel}"/>`).join('');
  // legend + scale
  const ly = Y(L.height) + 30; let lx = X(0);
  Object.entries(ROUTE_CATS).forEach(([c, d]) => { const t = calc.travelTotals[c]; b += `<line x1="${lx}" y1="${ly}" x2="${lx + 26}" y2="${ly}" stroke="${d.color}" stroke-width="4"/><text x="${lx + 32}" y="${ly + 4}" font-size="11" fill="${P.ink}">${esc(d.label)}${t ? ` — ${fmt(t.distShift, 0)} ft/shift` : ''}</text>`; lx += 250; });
  b += `<line x1="${X(0)}" y1="${ly + 22}" x2="${X(50)}" y2="${ly + 22}" stroke="${P.ink}" stroke-width="2"/><text x="${X(52)}" y="${ly + 26}" font-size="10" fill="${P.muted}">50 ft (layout scale: grid = 10 ft)</text>`;
  b += `<text x="${X(L.width)}" y="${ly + 26}" text-anchor="end" font-size="10" fill="${P.amber}">${esc(L.note || '')}</text>`;
  b += demoWatermark(P, w, h, opts.demo);
  return { svg: svgWrap(w, h, b, P, `${line.name} — spaghetti diagram (${key} state)`), w, h, X, Y, k, pad };
}

/* ---------------- Swimlane ---------------- */
function svgSwim(line, key, P, opts = {}) {
  const S = lineState(line, key), items = S.swim || [];
  const laneH = 92, labW = 130, colW = 175, top = 30;
  const cols = Math.max(6, ...items.map((i) => i.col + 1));
  const w = labW + cols * colW + 30, h = top + DEPTS.length * laneH + 70;
  let b = arrowDefs(P);
  DEPTS.forEach((d, i) => {
    const y = top + i * laneH;
    b += `<rect x="0" y="${y}" width="${w}" height="${laneH}" fill="${i % 2 ? P.soft2 : P.bg}" stroke="${P.grid2}"/><rect x="0" y="${y}" width="${labW}" height="${laneH}" fill="${P.soft}" stroke="${P.grid2}"/><circle cx="16" cy="${y + laneH / 2}" r="5" fill="${DEPT_COLORS[d]}"/>${textLines(72, y + laneH / 2, [d], { fill: P.ink, fs: 12, weight: 600 })}`;
  });
  const pos = (it) => ({ x: labW + it.col * colW + 18, y: top + DEPTS.indexOf(it.lane) * laneH + 19, w: 140, h: 54 });
  let handoffs = 0;
  items.forEach((it) => (it.next || []).forEach((nid) => {
    const t = items.find((x) => x.id === nid); if (!t) return;
    const a = pos(it), c = pos(t);
    const x1 = a.x + a.w, y1 = a.y + a.h / 2, x2 = c.x, y2 = c.y + c.h / 2, mx = Math.min(x2 - 12, x1 + 16);
    const cross = it.lane !== t.lane; if (cross) handoffs++;
    b += `<path d="M${x1},${y1} H${mx} V${y2} H${x2 - 1}" fill="none" stroke="${cross ? P.accent : P.line}" stroke-width="${cross ? 1.8 : 1.4}" ${cross ? 'stroke-dasharray="5 3"' : ''} marker-end="url(#ah)"/>`;
  }));
  items.forEach((it) => {
    const p = pos(it), flagged = Object.values(it.flags || {}).some(Boolean), sel = opts.selected === it.id;
    const st = sel ? P.sel : flagged ? P.red : P.stroke, sw = sel || flagged ? 2.6 : 1.4;
    let shape;
    if (it.kind === 'decision') shape = `<polygon points="${p.x + p.w / 2},${p.y - 4} ${p.x + p.w},${p.y + p.h / 2} ${p.x + p.w / 2},${p.y + p.h + 4} ${p.x},${p.y + p.h / 2}" fill="${P.node}" stroke="${st}" stroke-width="${sw}"/>`;
    else if (it.kind === 'delay') shape = `<path d="M${p.x},${p.y} H${p.x + p.w - 27} A27,27 0 0 1 ${p.x + p.w - 27},${p.y + p.h} H${p.x} Z" fill="${P.node}" stroke="${st}" stroke-width="${sw}"/>`;
    else shape = `<rect x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}" rx="${it.kind === 'handoff' ? 14 : 4}" fill="${P.node}" stroke="${st}" stroke-width="${sw}"/>` + (it.kind === 'approval' ? `<circle cx="${p.x + p.w - 10}" cy="${p.y + 10}" r="6" fill="none" stroke="${P.green}" stroke-width="1.5"/><path d="M${p.x + p.w - 13},${p.y + 10} l2,2 l4,-4" stroke="${P.green}" stroke-width="1.5" fill="none"/>` : '');
    let badges = ''; let bx = p.x;
    Object.entries(it.flags || {}).forEach(([k2, on]) => { if (!on) return; const lbl = { duplicate: 'DUP', unnecessary: 'APPR', delay: 'WAIT', gap: 'GAP' }[k2]; badges += `<rect x="${bx}" y="${p.y + p.h + 3}" width="36" height="14" rx="7" fill="${k2 === 'gap' ? P.amber : P.red}"/><text x="${bx + 18}" y="${p.y + p.h + 13}" text-anchor="middle" font-size="8.5" font-weight="700" fill="#FFFFFF">${lbl}<title>${esc(SWIM_FLAGS[k2])}</title></text>`; bx += 39; });
    b += `<g class="swim" data-id="${it.id}" style="cursor:move">${shape}${textLines(p.x + p.w / 2, p.y + p.h / 2, wrapText(it.label, 19), { fill: P.ink, fs: 11, lh: 13 })}${badges}</g>`;
  });
  const flagged = items.filter((i) => Object.values(i.flags || {}).some(Boolean)).length, approvals = items.filter((i) => i.kind === 'approval').length;
  b += `<text x="${labW}" y="${h - 22}" font-size="11" fill="${P.muted}">Cross-department handoffs: ${handoffs}   ·   Approvals: ${approvals}   ·   Flagged items: ${flagged}   ·   DUP = duplicated work, APPR = unnecessary approval, WAIT = delay, GAP = communication gap</text>`;
  b += demoWatermark(P, w, h, opts.demo);
  return { svg: svgWrap(w, h, b, P, `${line.name} — swimlane (${key} state)`), w, h, stats: { handoffs, approvals, flagged }, pos, laneH, labW, colW, top };
}
