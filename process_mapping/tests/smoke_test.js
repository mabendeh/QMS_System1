// End-to-end smoke test for the CHG Process Mapper (standalone build).
// Usage: NODE_PATH=<global node_modules>:<dir with xlsx,jszip,jspdf,pptxgenjs> node tests/smoke_test.js <libdir> <outdir>
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');
const LIBDIR = process.argv[2];
const OUT = process.argv[3] || path.join(__dirname, '..', 'exports');
const APP = 'file://' + path.resolve(__dirname, '..', 'app', 'CHG_Process_Mapper.html');
const LIBS = {
  'xlsx@0.18.5/dist/xlsx.full.min.js': 'xlsx/dist/xlsx.full.min.js',
  'pptxgenjs@4.0.1/dist/pptxgen.bundle.js': 'pptxgenjs/dist/pptxgen.bundle.js',
  'jszip@3.10.1/dist/jszip.min.js': 'jszip/dist/jszip.min.js',
  'jspdf@2.5.1/dist/jspdf.umd.min.js': 'jspdf/dist/jspdf.umd.min.js',
  'jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js': 'jspdf-autotable/dist/jspdf.plugin.autotable.min.js',
};
const TABS = ['dashboard', 'flow', 'vsm', 'spaghetti', 'swimlane', 'pam', 'data', 'opps', 'actions', 'finance', 'compare', 'checklist', 'formulas', 'io'];
let failures = 0; const ok = (c, m) => { console.log((c ? 'PASS ' : 'FAIL ') + m); if (!c) failures++; };

(async () => {
  fs.mkdirSync(OUT, { recursive: true }); const shots = path.join(OUT, 'screens'); fs.mkdirSync(shots, { recursive: true });
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true });
  await ctx.route(/cdn\.jsdelivr\.net\/npm\/.*/, (route) => { const k = Object.keys(LIBS).find((x) => route.request().url().endsWith(x)); return k ? route.fulfill({ path: path.join(LIBDIR, LIBS[k]), contentType: 'application/javascript' }) : route.abort(); });
  await ctx.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort());
  const page = await ctx.newPage();
  const errors = []; page.on('pageerror', (e) => errors.push(e.message)); page.on('console', (m) => { if (m.type() === 'error' && !/fonts|ERR_FAILED/.test(m.text())) errors.push(m.text()); });
  await page.goto(APP); await page.waitForSelector('#nav .tab');

  // 1. blank template: every tab renders, NYC visible, nothing invented
  for (const t of TABS) { await page.click(`#tab-${t}`); await page.waitForTimeout(120); }
  ok(errors.length === 0, 'blank template: all tabs render without errors ' + errors.join(' | '));
  await page.click('#tab-dashboard');
  ok((await page.locator('.kpi-v.nycv').count()) >= 8, 'blank template: KPIs show NYC (no invented numbers)');
  await page.screenshot({ path: path.join(shots, 'template_dashboard.png') });
  await page.click('#tab-checklist'); const items = await page.locator('.checks li').count(); ok(items > 100, 'floor-walk checklist generated: ' + items + ' items');

  // 1b. template import + template exports (placeholders)
  await page.click('#tab-io'); await page.setInputFiles('#impXlsx', path.resolve(__dirname, '..', 'templates', 'CHG_Nugget_Line_Process_Mapping_Workbook.xlsx')); await page.waitForTimeout(1500);
  ok(await page.evaluate(() => PRJ.lines.length === 3 && PRJ.lines[0].name === 'Nugget Line' && PRJ.lines[0].current.steps.length === 13 && !PRJ.demo), 'Excel template imports into the app');
  const grab0 = async (sel, name) => { const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 60000 }), page.click(sel)]); const p2 = path.join(OUT, name); await dl.saveAs(p2); return p2; };
  await page.click('#tab-io'); await grab0('#btnPptx', 'Nugget_Line_Current_State_TEMPLATE_process_maps.pptx'); await grab0('#btnPdf', 'Nugget_Line_Current_State_TEMPLATE_report.pdf');
  ok(errors.length === 0, 'template exports without errors ' + errors.join(' | '));
  // 2. demo data
  await page.click('#tab-io'); await page.click('[data-action="loadDemo"]'); await page.waitForTimeout(200);
  ok(await page.isVisible('#demoBadge'), 'demo badge visible');
  await page.click('#tab-dashboard'); ok((await page.locator('.kpi-v.nycv').count()) <= 2, 'demo: KPIs calculated');
  const kpis = await page.$$eval('.kpi', (els) => els.map((e) => e.innerText.replace(/\s+/g, ' ')));
  console.log('  ' + kpis.join('\n  '));
  for (const t of TABS) { await page.click(`#tab-${t}`); await page.waitForTimeout(150); await page.screenshot({ path: path.join(shots, `demo_${t}.png`), fullPage: false }); }
  ok(errors.length === 0, 'demo: all tabs render without errors ' + errors.join(' | '));

  // 7. exports
  const snap = await page.evaluate(() => ({ lines: PRJ.lines.length, steps: PRJ.lines[0].current.steps.length, routes: PRJ.lines[0].current.routes.length, opps: PRJ.opps.length, acts: PRJ.actions.length, ct: PRJ.lines[0].current.steps[4].f.cycleTime.v, lt: calcState(PRJ.lines[0], 'current').leadTime.v }));
  const grab = async (sel, tab, name) => { if (tab) await page.click('#tab-' + tab); const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 60000 }), page.click(sel)]); const p = path.join(OUT, name || dl.suggestedFilename()); await dl.saveAs(p); return p; };
  const xlsx = await grab('#btnXlsx', 'io', 'Nugget_Line_DEMO_export.xlsx'); ok(fs.statSync(xlsx).size > 10000, 'export: Excel ' + fs.statSync(xlsx).size + ' bytes');
  const pptx = await grab('#btnPptx', 'io', 'Nugget_Line_DEMO_process_maps.pptx'); ok(fs.statSync(pptx).size > 20000, 'export: PowerPoint ' + fs.statSync(pptx).size + ' bytes');
  const pdf = await grab('#btnPdf', 'io', 'Nugget_Line_DEMO_report.pdf'); ok(fs.statSync(pdf).size > 50000 && fs.statSync(pdf).size < 8e6, 'export: PDF report ' + fs.statSync(pdf).size + ' bytes');
  const cmp = await grab('[data-action="exportCompare"]', 'io', 'Nugget_Line_DEMO_current_vs_future.pdf'); ok(fs.statSync(cmp).size > 50000, 'export: comparison PDF');
  for (const k of ['flow', 'vsm', 'spaghetti', 'swimlane']) { await page.click('#tab-' + (k === 'swimlane' ? 'swimlane' : k)); const svg = await grab(`[data-action="exportSVG"][data-d="${k}"]`, null, `Nugget_Line_DEMO_${k}.svg`); const png = await grab(`[data-action="exportPNG"][data-d="${k}"]`, null, `Nugget_Line_DEMO_${k}.png`); ok(fs.statSync(svg).size > 2000 && fs.statSync(png).size > 10000, `export: ${k} SVG + PNG`); }
  ok(errors.length === 0, 'exports: no errors ' + errors.join(' | '));

  // 3. interactions: drag a flow node, connect, add block
  await page.click('#tab-flow'); await page.waitForSelector('#canvas-flow .fnode');
  const node = page.locator('#canvas-flow .fnode').nth(2); const bb = await node.boundingBox();
  const before = await page.evaluate(() => JSON.stringify(ST().flow.nodes[2]));
  await page.mouse.move(bb.x + bb.width / 2, bb.y + bb.height / 2); await page.mouse.down(); await page.mouse.move(bb.x + bb.width / 2 + 60, bb.y + bb.height / 2 + 40, { steps: 5 }); await page.mouse.up();
  const after = await page.evaluate(() => JSON.stringify(ST().flow.nodes[2]));
  ok(before !== after, 'flow: node drag moves the block');
  const edges0 = await page.evaluate(() => ST().flow.edges.length);
  await page.selectOption('#addKind', 'decision'); await page.click('[data-action="flowAdd"]');
  await page.click('#btnConnect'); const n0 = page.locator('#canvas-flow .fnode').nth(0); const n1 = page.locator('#canvas-flow .fnode').last();
  await n0.click(); await n1.click();
  ok((await page.evaluate(() => ST().flow.edges.length)) === edges0 + 1, 'flow: add decision block and connect');
  // 4. spaghetti: draw a route
  await page.click('#tab-spaghetti'); await page.selectOption('#spMode', 'draw');
  const cv = await page.locator('#canvas-spaghetti svg').boundingBox(); const r0 = await page.evaluate(() => ST().routes.length);
  await page.mouse.click(cv.x + cv.width * 0.3, cv.y + cv.height * 0.3); await page.mouse.click(cv.x + cv.width * 0.6, cv.y + cv.height * 0.35); await page.mouse.click(cv.x + cv.width * 0.6, cv.y + cv.height * 0.6);
  await page.click('[data-action="finishRoute"]');
  ok((await page.evaluate(() => ST().routes.length)) === r0 + 1, 'spaghetti: route drawn');
  const d = await page.evaluate(() => routeLength(ST().routes[ST().routes.length - 1].points)); ok(d > 20, 'spaghetti: route distance computed ' + d.toFixed(0) + ' ft');
  // 5. step editing: add, rename, enter a value
  await page.click('#tab-data'); const s0 = await page.evaluate(() => ST().steps.length);
  await page.click('[data-action="addStep"][data-id=""]'); ok((await page.evaluate(() => ST().steps.length)) === s0 + 1, 'steps: add step');
  await page.fill('#pf-cycleTime', '7.25'); await page.press('#pf-cycleTime', 'Tab');
  ok(await page.evaluate(() => { const s = ST().steps[ST().steps.length - 1]; return s.f.cycleTime.v === 7.25 && s.f.cycleTime.s === 'EST'; }), 'steps: entered value stored as Estimate by default');
  // 6. future state & compare
  await page.selectOption('#stateSel', 'future'); await page.click('#tab-compare'); await page.waitForTimeout(150);
  ok((await page.locator('table.tbl tbody tr').count()) >= 10, 'compare: KPI comparison table');
  await page.screenshot({ path: path.join(shots, 'demo_compare_future.png') });
  await page.selectOption('#stateSel', 'current');

  // 8. Excel round trip
  await page.click('#tab-io'); await page.click('[data-action="clearDemo"]'); await page.click('[data-action="clearDemo"]'); await page.waitForTimeout(150);
  await page.setInputFiles('#impXlsx', xlsx); await page.waitForTimeout(1500);
  const back = await page.evaluate(() => ({ lines: PRJ.lines.length, steps: PRJ.lines[0].current.steps.length, routes: PRJ.lines[0].current.routes.length, opps: PRJ.opps.length, acts: PRJ.actions.length, ct: PRJ.lines[0].current.steps[4].f.cycleTime.v, lt: calcState(PRJ.lines[0], 'current').leadTime.v }));
  ok(JSON.stringify(snap) === JSON.stringify(back), 'Excel round trip preserves data ' + JSON.stringify(back));
  ok(errors.length === 0, 'import: no errors ' + errors.join(' | '));

  // 9. phone width
  const ph = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 1, isMobile: true, hasTouch: true });
  await ph.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort());
  const mp = await ph.newPage(); await mp.goto(APP); await mp.waitForSelector('#nav .tab');
  await mp.evaluate(() => { PRJ = applyDemo(newProject()); render(); });
  for (const t of ['dashboard', 'flow', 'data', 'spaghetti']) { await mp.evaluate((x) => { UI.tab = x; render(); }, t); await mp.waitForTimeout(100); const sw = await mp.evaluate(() => document.documentElement.scrollWidth); ok(sw <= 392, `phone: no horizontal page scroll on ${t} (${sw}px)`); await mp.screenshot({ path: path.join(shots, `phone_${t}.png`) }); }
  await browser.close();
  console.log(failures ? `\n${failures} FAILURE(S)` : '\nALL PASSED'); process.exit(failures ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
