// Dump the app's blank template and demo project (plus app-calculated KPIs) to JSON for the workbook builder.
// Usage: NODE_PATH=<global node_modules> node tools/dump_project.js <outdir>
const path = require('path'), fs = require('fs'); const { chromium } = require('playwright');
(async () => {
  const out = process.argv[2]; const b = await chromium.launch(); const c = await b.newContext(); await c.route(/fonts|cdn\.jsdelivr/, (r) => r.abort());
  const p = await c.newPage(); await p.goto('file://' + path.resolve(__dirname, '..', 'app', 'CHG_Process_Mapper.html')); await p.waitForSelector('#nav .tab');
  const dump = await p.evaluate(() => {
    const kp = (prj) => prj.lines.map((l) => { const r = {}; ['current', 'future'].forEach((k) => { if (!l[k]) return; const c = calcState(l, k); r[k] = { avail: c.avail.v, takt: c.takt.v, leadTime: c.leadTime.v, vaTime: c.vaTime.v, pce: c.pce.v, bnECT: c.bnECT.v, bn: c.bottleneck && c.bottleneck.name, capacity: c.capacity.v, rty: c.rty.v, coHours: c.coHours.v, labor: c.labor.v, travelHrs: c.travelHrs.v, travelCost: c.travelCost.v, routes: Object.fromEntries(Object.entries(c.routes).map(([id, x]) => [id, x.dist])) }; }); return r; });
    const t = newProject(); const d = applyDemo(newProject());
    return { template: t, templateKpi: kp(t), demo: d, demoKpi: kp(d), checklistTemplate: t.lines.map((l) => checklistItems(l)), checklistDemo: d.lines.map((l) => checklistItems(l)), fields: FIELDS, settings: SETTINGS, formulas: FORMULAS, wastes: WASTES };
  });
  fs.writeFileSync(path.join(out, 'project_dump.json'), JSON.stringify(dump)); console.log('dumped'); await b.close();
})();
