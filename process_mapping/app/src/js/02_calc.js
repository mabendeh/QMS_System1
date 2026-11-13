/* =====================================================================
   Calculation engine. Every result = {v, s, missing[], formula, note}
   s: 'VER' only when every input is verified; 'EST' if any estimate;
      'NYC' (incomplete) if any required input is not yet collected.
   Partial sums are shown as "≥ value" and flagged incomplete.
   ===================================================================== */

const FORMULAS = [
  ['Available time per day', 'Shifts/day × (Shift length − Planned breaks)', 'min/day'],
  ['Takt time', 'Available time per day × 60 ÷ Customer demand per day', 's/unit'],
  ['Effective cycle time (per step)', 'Cycle time ÷ Uptime. Where uptime is not collected, cycle time is used and the result is marked Estimate.', 's/unit'],
  ['Bottleneck', 'Step with the highest effective cycle time among steps that apply to the line', ''],
  ['Line capacity', '3600 ÷ Bottleneck effective cycle time', 'units/hr'],
  ['Queue time (per step)', 'WIP before step ÷ Demand per day × Available time per day. If WIP was not counted, the observed Waiting time is used.', 'min'],
  ['Total lead time', 'Σ Queue time + Σ Processing time over applicable steps', 'min'],
  ['Value-added time', 'Σ Processing time of steps classified Value-added', 'min'],
  ['Non-value-added time', 'Total lead time − Value-added time', 'min'],
  ['Process cycle efficiency (PCE)', 'Value-added time ÷ Total lead time', '%'],
  ['Rolled throughput yield', 'Π (Yield) over steps; where yield is missing, (1 − Scrap) is used', '%'],
  ['Changeover time per week', 'Σ (Changeover time × Changeovers per week) ÷ 60', 'hr/week'],
  ['Route distance', 'Σ segment lengths of the drawn route × layout scale (1 unit = 1 ft)', 'ft/trip'],
  ['Travel time per shift', 'Route distance × Trips per shift ÷ Speed for the movement category', 'min/shift'],
  ['Annual travel labor hours', 'Travel time per shift × People × Shifts/day × Operating days/year ÷ 60', 'hr/yr'],
  ['Annual travel labor cost', 'Annual travel labor hours × Loaded labor rate', '$/yr'],
  ['Estimated annual labor savings (travel)', 'Current annual travel labor cost − Future annual travel labor cost', '$/yr'],
  ['Changeover reduction', '(Current − Future changeover hours/week) × Operating weeks/year; labor value = Σ hours × step labor × labor rate', 'hr/yr'],
  ['Scrap reduction opportunity', '(Future RTY − Current RTY) × Demand/day × Operating days/year × Material cost/unit (valued at cost)', '$/yr'],
  ['Throughput improvement potential', '(Future capacity − Current capacity) ÷ Current capacity', '%'],
  ['Labor productivity', 'Line capacity ÷ Total labor on the line', 'units/labor-hr'],
  ['PAM totals', 'Σ activity time, distance and people-minutes by activity type (O, T, I, D, S); NVA time = Σ time of activities marked NVA', 'min / ft'],
];

function R(v, s, missing = [], formula = '', note = '') { return { v, s, missing, formula, note }; }

function lineState(line, key) { return key === 'future' && line.future ? line.future : line.current; }
function applicable(steps) { return steps.filter((s) => s.applic !== 'Not applicable'); }

function calcState(line, key = 'current') {
  const S = lineState(line, key);
  const st = line.settings;
  const out = { steps: {}, key };
  const g = (k) => val(st[k]);
  const gs = (k) => fstat(st[k]);

  // available time & takt
  const shifts = g('shiftsPerDay'), len = g('minPerShift'), brk = g('breakMin'), dem = g('demandPerDay');
  const availMiss = [['shiftsPerDay', shifts], ['minPerShift', len], ['breakMin', brk]].filter((x) => x[1] == null).map((x) => SETTINGS.find((s) => s.k === x[0]).label);
  const avail = availMiss.length ? null : shifts * (len - brk);
  out.avail = R(avail, availMiss.length ? 'NYC' : worst([gs('shiftsPerDay'), gs('minPerShift'), gs('breakMin')]), availMiss, FORMULAS[0][1]);
  out.takt = R(avail != null && dem ? (avail * 60) / dem : null, avail != null && dem != null ? worst([out.avail.s, gs('demandPerDay')]) : 'NYC', [...availMiss, ...(dem == null ? ['Customer demand'] : [])], FORMULAS[1][1]);

  const steps = applicable(S.steps);
  let ltSum = 0, ltMiss = [], ltStat = [], vaSum = 0, vaMiss = [], vaStat = [], best = null, ctCount = 0;
  let rty = 1, rtyStat = [], rtyMiss = [], coH = 0, coStat = [], laborSum = 0, laborStat = [], laborMiss = [];
  steps.forEach((s) => {
    const f = s.f, r = {};
    const ct = val(f.cycleTime), up = val(f.uptime), pt = val(f.processTime), wip = val(f.wip), wt = val(f.waitTime);
    // effective CT
    if (ct != null) {
      ctCount++;
      const ect = up ? ct / (up / 100) : ct;
      r.ect = R(ect, up ? worst([fstat(f.cycleTime), fstat(f.uptime)]) : 'EST', up ? [] : ['Uptime (assumed 100%)'], FORMULAS[2][1]);
      if (!best || ect > best.ect.v) best = { step: s, ect: r.ect };
    } else r.ect = R(null, 'NYC', ['Cycle time'], FORMULAS[2][1]);
    // queue time
    if (wip != null && dem && avail != null) r.queue = R((wip / dem) * avail, worst([fstat(f.wip), gs('demandPerDay'), out.avail.s]), [], FORMULAS[5][1]);
    else if (wt != null) r.queue = R(wt, fstat(f.waitTime), [], FORMULAS[5][1], 'Observed waiting time used (WIP not counted)');
    else r.queue = R(null, 'NYC', ['WIP before step or Waiting time' + (wip != null ? ' (demand / available time missing)' : '')], FORMULAS[5][1]);
    if (r.queue.v != null) { ltSum += r.queue.v; ltStat.push(r.queue.s); } else ltMiss.push(s.name + ': queue');
    if (pt != null) { ltSum += pt; ltStat.push(fstat(f.processTime)); if (s.vaClass === 'VA') { vaSum += pt; vaStat.push(fstat(f.processTime)); } }
    else { ltMiss.push(s.name + ': processing time'); if (s.vaClass === 'VA') vaMiss.push(s.name); }
    // yield
    const y = val(f.yield), sc = val(f.scrap);
    if (y != null) { rty *= y / 100; rtyStat.push(fstat(f.yield)); } else if (sc != null) { rty *= 1 - sc / 100; rtyStat.push(fstat(f.scrap)); } else rtyMiss.push(s.name);
    // changeover
    const co = val(f.changeover), cpw = val(f.changeoversPerWeek);
    if (co != null && cpw != null) { coH += (co * cpw) / 60; coStat.push(worst([fstat(f.changeover), fstat(f.changeoversPerWeek)])); r.coH = (co * cpw) / 60; }
    const lab = val(f.labor);
    if (lab != null) { laborSum += lab; laborStat.push(fstat(f.labor)); } else laborMiss.push(s.name);
    out.steps[s.id] = r;
  });
  out.ctCoverage = { measured: ctCount, total: steps.length };
  out.leadTime = R(ltStat.length ? ltSum : null, ltMiss.length ? 'NYC' : worst(ltStat), ltMiss, FORMULAS[6][1]);
  out.vaTime = R(vaStat.length ? vaSum : null, vaMiss.length ? 'NYC' : worst(vaStat), vaMiss, FORMULAS[7][1]);
  out.nvaTime = R(ltStat.length ? ltSum - vaSum : null, worst([out.leadTime.s, out.vaTime.s]), [...ltMiss], FORMULAS[8][1]);
  out.pce = R(ltStat.length && ltSum > 0 ? vaSum / ltSum : null, worst([out.leadTime.s, out.vaTime.s]), ltMiss, FORMULAS[9][1]);
  out.bottleneck = best ? { stepId: best.step.id, name: best.step.name, ect: best.ect } : null;
  out.bnECT = best ? best.ect : R(null, 'NYC', ['Cycle time at every step'], FORMULAS[3][1]);
  if (best && ctCount < steps.length) out.bnECT = R(best.ect.v, 'NYC', [`Cycle time at ${steps.length - ctCount} of ${steps.length} steps`], FORMULAS[3][1], 'Bottleneck is provisional until every step is timed');
  out.capacity = R(best ? 3600 / best.ect.v : null, out.bnECT.s, out.bnECT.missing, FORMULAS[4][1]);
  out.rty = R(rtyMiss.length === steps.length ? null : rty, rtyMiss.length ? 'NYC' : worst(rtyStat), rtyMiss, FORMULAS[10][1]);
  out.coHours = R(coStat.length ? coH : null, coStat.length ? worst(coStat) : 'NYC', coStat.length ? [] : ['Changeover time and changeovers/week'], FORMULAS[11][1]);
  out.labor = R(laborStat.length ? laborSum : null, laborMiss.length ? 'NYC' : worst(laborStat), laborMiss, 'Σ Labor over applicable steps');
  out.productivity = R(out.capacity.v != null && laborSum ? out.capacity.v / laborSum : null, worst([out.capacity.s, out.labor.s]), [...out.capacity.missing, ...laborMiss], FORMULAS[20][1]);

  // travel
  out.routes = {};
  const totals = {}; let travelCost = 0, travelHrs = 0, travelStat = [], travelMiss = [];
  (S.routes || []).forEach((r) => {
    const d = routeLength(r.points);
    const cat = ROUTE_CATS[r.cat] || ROUTE_CATS.walk;
    const spd = g(cat.speedKey), trips = val(r.trips), ppl = val(r.people);
    const res = { dist: d };
    res.distShift = trips != null ? d * trips : null;
    res.timeShift = res.distShift != null && spd ? res.distShift / spd : null;
    const yrMiss = [];
    if (trips == null) yrMiss.push('Trips per shift'); if (!spd) yrMiss.push(SETTINGS.find((s) => s.k === cat.speedKey).label);
    if (ppl == null) yrMiss.push('People'); if (shifts == null) yrMiss.push('Shifts per day'); if (g('daysPerYear') == null) yrMiss.push('Operating days per year');
    const stats = [r.status === 'VER' ? 'VER' : 'EST', fstat(r.trips), fstat(r.people), gs(cat.speedKey), gs('shiftsPerDay'), gs('daysPerYear')];
    res.hrsYear = R(yrMiss.length ? null : (res.timeShift * ppl * shifts * g('daysPerYear')) / 60, yrMiss.length ? 'NYC' : worst(stats), yrMiss, FORMULAS[14][1]);
    const rate = g('laborRate');
    res.costYear = R(res.hrsYear.v != null && rate != null ? res.hrsYear.v * rate : null, res.hrsYear.v != null && rate != null ? worst([res.hrsYear.s, gs('laborRate')]) : 'NYC', [...yrMiss, ...(rate == null ? ['Loaded labor rate'] : [])], FORMULAS[15][1]);
    if (res.hrsYear.v != null) travelHrs += res.hrsYear.v; if (res.costYear.v != null) { travelCost += res.costYear.v; travelStat.push(res.costYear.s); } else travelMiss.push(r.name);
    const t = (totals[r.cat] = totals[r.cat] || { dist: 0, distShift: 0, time: 0 });
    t.dist += d; t.distShift += res.distShift || 0; t.time += res.timeShift || 0;
    out.routes[r.id] = res;
  });
  out.travelTotals = totals;
  out.travelHrs = R(travelStat.length ? travelHrs : null, travelMiss.length || !(S.routes || []).length ? 'NYC' : worst(travelStat), travelMiss, FORMULAS[14][1]);
  out.travelCost = R((S.routes || []).length ? travelCost : null, travelMiss.length || !(S.routes || []).length ? 'NYC' : worst(travelStat), travelMiss, FORMULAS[15][1]);

  // PAM
  const pam = { byType: {}, time: 0, nva: 0, va: 0, dist: 0, peopleMin: 0, missing: 0, stat: [] };
  Object.keys(ACT_TYPES).forEach((k) => (pam.byType[k] = { n: 0, time: 0, dist: 0 }));
  (S.acts || []).forEach((a) => {
    const b = pam.byType[a.type]; b.n++;
    if (a.time == null || a.time === '') { pam.missing++; return; }
    const t = +a.time; b.time += t; pam.time += t; if (a.va === 'NVA') pam.nva += t; if (a.va === 'VA') pam.va += t;
    if (a.dist) { b.dist += +a.dist; pam.dist += +a.dist; }
    if (a.people) pam.peopleMin += t * +a.people;
    pam.stat.push(a.status === 'VER' ? 'VER' : 'EST');
  });
  pam.s = pam.missing ? 'NYC' : worst(pam.stat);
  out.pam = pam;

  // data completeness
  let tot = 0, got = 0, ver = 0;
  steps.forEach((s) => FIELDS.forEach((fd) => { tot++; const x = fstat(s.f[fd.k]); if (x !== 'NYC') got++; if (x === 'VER') ver++; }));
  SETTINGS.forEach((sd) => { tot++; const x = gs(sd.k); if (x !== 'NYC') got++; if (x === 'VER') ver++; });
  out.completeness = { tot, got, ver };
  return out;
}

function routeLength(pts) { let d = 0; for (let i = 1; i < (pts || []).length; i++) d += Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]); return d; }

/* current vs future comparison & estimated savings (never reported as validated) */
function calcCompare(line) {
  const c = calcState(line, 'current');
  if (!line.future) return { c, f: null };
  const f = calcState(line, 'future');
  const st = line.settings, g = (k) => val(st[k]);
  const res = { c, f };
  const both = (a, b) => worst([a.s, b.s]);
  res.travelSavings = R(c.travelCost.v != null && f.travelCost.v != null ? c.travelCost.v - f.travelCost.v : null, both(c.travelCost, f.travelCost), [...c.travelCost.missing, ...f.travelCost.missing], FORMULAS[16][1]);
  res.travelHrsSaved = R(c.travelHrs.v != null && f.travelHrs.v != null ? c.travelHrs.v - f.travelHrs.v : null, both(c.travelHrs, f.travelHrs), [], 'Current − Future annual travel labor hours');
  // changeover
  const wk = g('weeksPerYear');
  let coHrsYr = null, coValue = null, coMiss = [];
  if (c.coHours.v != null && f.coHours.v != null && wk != null) {
    coHrsYr = (c.coHours.v - f.coHours.v) * wk;
    let laborHrs = 0;
    applicable(line.current.steps).forEach((s) => {
      const fs = line.future.steps.find((x) => x.id === s.id); if (!fs) return;
      const ch = c.steps[s.id] && c.steps[s.id].coH, fh = f.steps[fs.id] && f.steps[fs.id].coH, lab = val(s.f.labor);
      if (ch != null && fh != null && lab != null) laborHrs += (ch - fh) * wk * lab;
    });
    coValue = g('laborRate') != null ? laborHrs * g('laborRate') : null;
  } else coMiss.push('Changeover data in both states and operating weeks/year');
  res.coHrsYr = R(coHrsYr, coMiss.length ? 'NYC' : worst([c.coHours.s, f.coHours.s, fstat(st.weeksPerYear)]), coMiss, FORMULAS[17][1]);
  res.coValue = R(coValue, coValue == null ? 'NYC' : worst([res.coHrsYr.s, fstat(st.laborRate)]), coValue == null ? [...coMiss, 'Loaded labor rate'] : [], FORMULAS[17][1], 'Capacity unless paid hours are removed — Finance decides');
  // scrap
  const dem = g('demandPerDay'), days = g('daysPerYear'), mc = g('materialCost');
  const scMiss = [];
  if (c.rty.v == null || f.rty.v == null) scMiss.push('Scrap / yield in both states'); if (dem == null) scMiss.push('Customer demand'); if (days == null) scMiss.push('Operating days/year'); if (mc == null) scMiss.push('Material cost per unit');
  res.scrapSavings = R(scMiss.length ? null : (f.rty.v - c.rty.v) * dem * days * mc, scMiss.length ? 'NYC' : worst([c.rty.s, f.rty.s, fstat(st.demandPerDay), fstat(st.daysPerYear), fstat(st.materialCost)]), scMiss, FORMULAS[18][1]);
  // throughput
  res.throughputGain = R(c.capacity.v && f.capacity.v ? (f.capacity.v - c.capacity.v) / c.capacity.v : null, both(c.capacity, f.capacity), [...c.capacity.missing, ...f.capacity.missing], FORMULAS[19][1]);
  return res;
}

function statusLabel(s) { return s === 'VER' ? 'Verified inputs' : s === 'EST' ? 'Estimate' : 'Incomplete'; }
