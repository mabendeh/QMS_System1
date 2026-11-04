const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "Skill Matrix Green",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1F2A24", lt1: "FFFFFF", dk2: "0E3B27", lt2: "EEF5F0",
    accent1: "1D6F42", accent2: "E39B2D", accent3: "5FA47A", accent4: "C0504D",
    accent5: "7A8C82", accent6: "2F5D8A", hlink: "1D6F42", folHlink: "5FA47A",
  },
};
const H = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Building a Skill Matrix in Excel";
pres.author = "Quality Team";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

const W = 13.333, M = 0.6;
const FOOT = "Building a Skill Matrix in Excel  |  C.H. Guenther";

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: H.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.2, w: 7.2, h: 1.9, fontSize: 44, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 4.3, w: 7.2, h: 1.2, fontSize: 20, color: "CFE3D6", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: H.lt1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.32, w: 8, h: 0.35, fontSize: 14, bold: true, color: C.accent2, margin: 0, charSpacing: 2 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.68, w: W - 2 * M, h: 0.8, fontSize: 36, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: M, y: 7.05, w: 8, h: 0.3, fontSize: 10, color: H.accent5, margin: 0, isTextBox: true } } },
  ],
  slideNumber: { x: W - M - 0.6, y: 7.05, w: 0.6, h: 0.3, fontSize: 10, color: H.accent5, align: "right" },
});

// ---------- Sample data ----------
const SKILLS = ["Mixer Setup", "Line Packaging", "Allergen Control", "Sanitation (SSOP)", "CCP Monitoring", "Metal Detector", "Forklift"];
const SHORT = ["Mixer", "Packing", "Allergen", "Sanit.", "CCP", "Metal Det.", "Forklift"];
const PEOPLE = ["Maria G.", "James T.", "Ana R.", "David K.", "Luis M.", "Sarah P.", "Kevin B.", "Nora H."];
const SCORES = [
  [4, 3, 4, 3, 3, 3, 0],
  [3, 4, 3, 2, 2, 3, 3],
  [2, 3, 3, 4, 1, 2, 0],
  [3, 2, 2, 3, 4, 4, 2],
  [1, 3, 2, 2, 0, 1, 4],
  [0, 2, 3, 1, 3, 2, 0],
  [2, 1, 1, 3, 0, 1, 3],
  [3, 2, 4, 2, 2, 3, 1],
];
// Plant-wide (all departments) illustrative totals
const DEPTS = ["Mixing", "Packaging", "Sanitation", "Warehouse", "Quality", "Maintenance"];
const DEPT_QUAL = [
  [6, 0, 4, 1, 2, 0, 0],
  [0, 8, 5, 1, 3, 4, 1],
  [0, 0, 3, 8, 0, 0, 0],
  [0, 0, 1, 0, 0, 0, 5],
  [1, 1, 3, 1, 2, 2, 0],
  [2, 1, 0, 0, 0, 1, 0],
];
const TARGET = [8, 12, 20, 10, 9, 6, 5];
const QUAL = SKILLS.map((_, j) => DEPT_QUAL.reduce((a, r) => a + r[j], 0));
const PDEPT = ["Mixing", "Packaging", "Packaging", "Mixing", "Warehouse", "Sanitation", "Sanitation", "Quality"];
const LEVEL_FILL = ["F4C7C3", "FCE4B6", "FFF2CC", "CFE8D5", "8CCB9F"];
const LEVEL_TXT = ["8A1C14", "7A4A00", "6B5A00", "1D5B35", "0E3B27"];
const LEVELS = [
  ["0", "Not trained", "Has not received training on this task"],
  ["1", "In training", "Learning; must be supervised at all times"],
  ["2", "Trained", "Can do the task with occasional support"],
  ["3", "Qualified", "Works independently to the SOP standard"],
  ["4", "Trainer", "Can teach others and sign off training"],
];

// ---------- Helpers ----------
const XL_HEAD = "E6E6E6", XL_LINE = "BFBFBF";
function cell(text, o = {}) {
  return { text: String(text), options: Object.assign({ fontSize: 12, color: H.dk1, border: { type: "solid", pt: 0.75, color: XL_LINE }, valign: "middle", margin: [0, 0.05, 0, 0.05] }, o) };
}
// Draws a spreadsheet-like grid with column letters and row numbers.
function excelGrid(slide, x, y, colW, rowH, rows, startRow, name, letters = "ABCDEFGHIJKLMNOP".split("")) {
  const nCols = rows[0].length;
  const header = [cell("", { fill: { color: XL_HEAD } })];
  for (let i = 0; i < nCols; i++) header.push(cell(letters[i], { fill: { color: XL_HEAD }, align: "center", color: "595959", fontSize: 11 }));
  const body = rows.map((r, i) => [cell(startRow + i, { fill: { color: XL_HEAD }, align: "center", color: "595959", fontSize: 11 }), ...r]);
  slide.addTable([header, ...body], { x, y, colW: [0.4, ...colW], rowH, objectName: name });
}
function badge(slide, x, y, n, name, size = 0.55, fill = C.accent1) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: size, h: size, fill: { color: fill }, rectRadius: 0.08, line: { color: fill }, objectName: name + " badge" });
  slide.addText(String(n), { x, y, w: size, h: size, fontSize: size > 0.5 ? 20 : 14, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: name + " number" });
}
function bulletList(slide, items, x, y, w, h, name, size = 16) {
  slide.addText(
    items.map((t, i) => {
      const runs = Array.isArray(t) ? t : [{ text: t }];
      return runs.map((r, k) => ({ text: r.text, options: Object.assign({ bullet: k === 0 ? { indent: 18 } : undefined, breakLine: k === runs.length - 1 && i < items.length - 1, paraSpaceAfter: 10 }, r.options || {}) }));
    }).flat(),
    { x, y, w, h, fontSize: size, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: name }
  );
}
// Numbered how-to steps: badge + text rows
function steps(slide, items, x, y, w, rowH, name) {
  items.forEach((t, i) => {
    const yy = y + i * rowH;
    badge(slide, x, yy + 0.02, i + 1, `${name} ${i + 1}`, 0.38, C.accent3);
    slide.addText(t, { x: x + 0.55, y: yy, w: w - 0.55, h: rowH - 0.08, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: `${name} ${i + 1} text` });
  });
}
function content(section, kicker, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function card(slide, x, y, w, h, name, fill = C.background2) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.1, objectName: name });
}
function formulaBox(slide, x, y, w, label, formula, cellRef, name, fsz = 15) {
  slide.addText(label, { x, y, w, h: 0.35, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: name + " label" });
  // Excel-style formula bar
  slide.addShape(pres.shapes.RECTANGLE, { x, y: 0.4 + y, w, h: 0.5, fill: { color: "FFFFFF" }, line: { color: XL_LINE, width: 1 }, objectName: name + " bar" });
  slide.addText(cellRef, { x: x + 0.08, y: 0.4 + y, w: 0.6, h: 0.5, fontSize: 13, color: "595959", fontFace: "Courier New", valign: "middle", margin: 0, isTextBox: true, objectName: name + " ref" });
  slide.addText("fx", { x: x + 0.72, y: 0.4 + y, w: 0.35, h: 0.5, fontSize: 13, italic: true, color: H.accent5, valign: "middle", margin: 0, isTextBox: true, objectName: name + " fx" });
  slide.addText(formula, { x: x + 1.1, y: 0.4 + y, w: w - 1.2, h: 0.5, fontSize: fsz, bold: true, color: C.accent1, fontFace: "Courier New", valign: "middle", margin: 0, isTextBox: true, objectName: name + " formula" });
}

// ================= SLIDES =================
// 1. Title
pres.addSection({ title: "Introduction" });
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
  s.addText("Building a Skill Matrix in Excel", { placeholder: "title" });
  s.addText([
    { text: "A step-by-step guide to a plant-wide matrix for C.H. Guenther", options: { breakLine: true } },
    { text: "Training & Quality  |  November 2026", options: { fontSize: 14, color: "9FC3AD" } },
  ], { placeholder: "body" });
  // decorative heat-map of skill levels
  const gx = 8.6, gy = 1.35, sz = 0.52, gap = 0.1;
  SCORES.forEach((r, i) => r.forEach((v, j) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: gx + j * (sz + gap), y: gy + i * (sz + gap), w: sz, h: sz, rectRadius: 0.06, fill: { color: LEVEL_FILL[v] }, line: { color: LEVEL_FILL[v] }, objectName: `heat ${i}-${j}` });
  }));
  s.addNotes("Today I'll walk through exactly how I built our plant-wide skill matrix in Excel, step by step: one workbook that covers every department, shift and employee in the plant.");
}

// 2. Why
{
  const s = content("Introduction", "WHY IT MATTERS", "One sheet that shows who can do what");
  const items = [
    ["See coverage at a glance", "Colour-coded levels show instantly which tasks have enough qualified people across every department and shift."],
    ["Plan training & cross-training", "Gaps become a training plan: who to train next, on what, and who can train them."],
    ["Prove competence in audits", "Gives SQF / BRCGS and customer auditors clear evidence that people are trained for their roles."],
  ];
  const cw = (W - 2 * M - 2 * 0.4) / 3;
  items.forEach(([h, b], i) => {
    const x = M + i * (cw + 0.4), y = 2.0;
    card(s, x, y, cw, 3.9, `why card ${i + 1}`);
    badge(s, x + 0.35, y + 0.4, i + 1, `why ${i + 1}`);
    s.addText(h, { x: x + 0.35, y: y + 1.2, w: cw - 0.7, h: 0.8, fontSize: 21, bold: true, color: C.text2, valign: "top", margin: 0, isTextBox: true, fontFace: THEME.headFontFace, objectName: `why head ${i + 1}` });
    s.addText(b, { x: x + 0.35, y: y + 2.05, w: cw - 0.7, h: 1.6, fontSize: 16, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: `why body ${i + 1}` });
  });
  s.addText("Plant-wide: one workbook for every department, so a gap in one area can be covered by trained people from another.", { x: M, y: 6.2, w: W - 2 * M, h: 0.4, fontSize: 15, italic: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "why tip" });
  s.addNotes("A skill matrix answers three questions: who can do each task, where are we exposed if someone is absent, and how do we prove competence to auditors.");
}

// 3. Roadmap
{
  const s = content("Introduction", "THE ROADMAP", "Eight steps from blank sheet to live matrix");
  const ST = [
    ["Scope & people", "Whole plant, every employee"],
    ["List the skills", "Tasks from SOPs & HACCP plan"],
    ["Set the rating scale", "0 = not trained … 4 = trainer"],
    ["Build the grid", "People down, skills across"],
    ["Add drop-down lists", "Data Validation keeps scores clean"],
    ["Colour the levels", "Conditional Formatting"],
    ["Add summary formulas", "COUNTIF, targets, gap flags"],
    ["Chart, protect, maintain", "Dashboard + review routine"],
  ];
  const tw = (W - 2 * M - 3 * 0.4) / 4, th = 2.1;
  ST.forEach(([h, b], i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = M + col * (tw + 0.4), y = 1.95 + row * (th + 0.4);
    card(s, x, y, tw, th, `road card ${i + 1}`, row === 0 ? C.background2 : "E3EFE7");
    badge(s, x + 0.3, y + 0.3, i + 1, `road ${i + 1}`);
    s.addText(h, { x: x + 0.3, y: y + 1.0, w: tw - 0.5, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: `road head ${i + 1}` });
    s.addText(b, { x: x + 0.3, y: y + 1.45, w: tw - 0.5, h: 0.5, fontSize: 14, color: C.text1, margin: 0, isTextBox: true, objectName: `road body ${i + 1}` });
  });
  s.addNotes("Here is the whole journey on one slide, applied to the entire plant. Steps 1 to 3 are planning, 4 to 6 build the sheet, 7 and 8 turn it into a management tool.");
}

// 4. Step 1
pres.addSection({ title: "Build the matrix" });
{
  const s = content("Build the matrix", "STEP 1", "Cover the whole plant and list everyone");
  bulletList(s, [
    [{ text: "Scope = the whole plant: ", options: { bold: true } }, { text: "Mixing, Packaging, Sanitation, Warehouse, Quality and Maintenance, all shifts." }],
    [{ text: "Pull the roster from HR ", options: { bold: true } }, { text: "so names and IDs match training records." }],
    [{ text: "Capture the basics: ", options: { bold: true } }, { text: "Employee ID, Name, Department, Role, Shift." }],
    [{ text: "Include temps and new hires ", options: { bold: true } }, { text: "– they are often your biggest gap." }],
  ], M, 1.95, 5.2, 4.5, "step1 bullets");
  const rows = [
    ["Emp ID", "Name", "Department", "Role", "Shift"].map((t) => cell(t, { bold: true, fill: { color: H.accent1 }, color: "FFFFFF" })),
    ...PEOPLE.map((p, i) => [cell(1040 + i * 7), cell(p), cell(PDEPT[i]), cell(["Operator", "Lead", "Operator", "Supervisor", "Driver", "Temp", "Sanitor", "QA Tech"][i]), cell(i % 3 === 0 ? "A" : i % 3 === 1 ? "B" : "C", { align: "center" })]),
  ];
  excelGrid(s, 6.3, 1.95, [0.9, 1.25, 1.4, 1.25, 0.7], 0.42, rows, 1, "step1 grid");
  s.addText("Sheet: “Employees”  –  one row per person in the plant (example rows shown)", { x: 6.3, y: 6.3, w: 6.4, h: 0.35, fontSize: 12, italic: true, color: H.accent5, margin: 0, isTextBox: true, objectName: "step1 caption" });
  s.addNotes("The matrix covers the whole plant. I pulled the official roster from HR for every department and shift so names and IDs match our training records, and added a Department column so we can filter and roll up later.");
}

// 5. Step 2
{
  const s = content("Build the matrix", "STEP 2", "List the skills that really matter");
  bulletList(s, [
    [{ text: "Source the tasks ", options: { bold: true } }, { text: "from SOPs, job descriptions, the HACCP / food safety plan and customer requirements." }],
    [{ text: "Plant-wide vs department skills: ", options: { bold: true } }, { text: "GMP and Allergen Control apply to everyone; Mixer Setup only to Mixing." }],
    [{ text: "Keep it focused: ", options: { bold: true } }, { text: "group by category (Food Safety, Production, Equipment, Quality) to stay readable." }],
    [{ text: "Set a target ", options: { bold: true } }, { text: "– the plant-wide minimum of qualified people each skill needs to cover all shifts." }],
  ], M, 1.95, 5.2, 4.6, "step2 bullets");
  const cats = ["Production", "Production", "Food Safety", "Food Safety", "Food Safety", "Quality", "Equipment"];
  const src = ["SOP-MX-01", "SOP-PK-04", "Allergen Plan", "SSOP-02", "HACCP Plan", "SOP-QA-07", "Safety-12"];
  const rows = [
    ["Category", "Skill", "Source", "Target"].map((t) => cell(t, { bold: true, fill: { color: H.accent1 }, color: "FFFFFF" })),
    ...SKILLS.map((k, i) => [cell(cats[i]), cell(k), cell(src[i]), cell(TARGET[i], { align: "center" })]),
  ];
  excelGrid(s, 6.3, 1.95, [1.35, 1.95, 1.4, 0.8], 0.45, rows, 1, "step2 grid");
  s.addText("Sheet: “Skills”", { x: 6.3, y: 6.15, w: 6, h: 0.35, fontSize: 12, italic: true, color: H.accent5, margin: 0, isTextBox: true, objectName: "step2 caption" });
  s.addNotes("Don't list every possible task. Focus on the tasks that affect food safety, quality and line uptime. The Target column drives the gap analysis later.");
}

// 6. Step 3 – rating scale
{
  const s = content("Build the matrix", "STEP 3", "Agree one simple rating scale");
  const tw = (W - 2 * M - 4 * 0.25) / 5;
  LEVELS.forEach(([n, h, b], i) => {
    const x = M + i * (tw + 0.25), y = 2.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: tw, h: 3.4, rectRadius: 0.1, fill: { color: LEVEL_FILL[i] }, line: { color: LEVEL_FILL[i] }, objectName: `level card ${n}` });
    s.addText(n, { x: x + 0.3, y: y + 0.25, w: tw - 0.6, h: 1.1, fontSize: 60, bold: true, color: LEVEL_TXT[i], fontFace: THEME.headFontFace, margin: 0, isTextBox: true, objectName: `level num ${n}` });
    s.addText(h, { x: x + 0.3, y: y + 1.45, w: tw - 0.6, h: 0.45, fontSize: 20, bold: true, color: LEVEL_TXT[i], margin: 0, isTextBox: true, objectName: `level head ${n}` });
    s.addText(b, { x: x + 0.3, y: y + 1.95, w: tw - 0.6, h: 1.3, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: `level body ${n}` });
  });
  s.addText([
    { text: "Rule of thumb: ", options: { bold: true } },
    { text: "only level 3 and 4 count as “qualified” in the summary. Put this legend on its own sheet called “Legend” so everyone rates the same way." },
  ], { x: M, y: 5.75, w: W - 2 * M, h: 0.8, fontSize: 15, color: C.text1, margin: 0, isTextBox: true, objectName: "level note" });
  s.addNotes("The scale is the most important decision. Agree it with supervisors first. Level 3 means the person can work alone to the SOP; level 4 means they can also train others.");
}

// 7. Step 4 – build the grid
{
  const s = content("Build the matrix", "STEP 4", "Build the grid: people down, skills across");
  const head = [cell("Name", { bold: true, fill: { color: H.accent1 }, color: "FFFFFF" }), cell("Department", { bold: true, fill: { color: H.accent1 }, color: "FFFFFF" }), ...SHORT.map((t) => cell(t, { bold: true, fill: { color: H.accent1 }, color: "FFFFFF", align: "center", fontSize: 11 }))];
  const rows = [head, ...PEOPLE.slice(0, 6).map((p, i) => [cell(p), cell(PDEPT[i]), ...SCORES[i].map((v) => cell(v, { align: "center" }))])];
  excelGrid(s, M, 1.95, [1.1, 1.15, ...SHORT.map(() => 0.75)], 0.45, rows, 4, "step4 grid");
  steps(s, [
    "Names in column A, department in B; skills across row 4 (linked to the Skills sheet).",
    "Select the header row → Home › Orientation › Rotate Text Up for narrow columns.",
    "Click C5 → View › Freeze Panes so names and skills stay visible.",
    "Ctrl + T to make it a Table, then filter column B to see one department.",
  ], 8.85, 1.95, W - M - 8.85, 1.12, "step4 steps");
  s.addText("Sheet: “Matrix”  –  scores start in cell C5", { x: M, y: 5.7, w: 7, h: 0.35, fontSize: 12, italic: true, color: H.accent5, margin: 0, isTextBox: true, objectName: "step4 caption" });
  s.addNotes("This is the heart of the workbook. Each row is a person anywhere in the plant, each column a skill, and each cell holds a 0–4 score. Freeze panes and an Excel Table make it much easier to use.");
}

// 8. Step 5 – data validation
{
  const s = content("Build the matrix", "STEP 5", "Add drop-down lists so scores stay clean");
  steps(s, [
    "Select all score cells (e.g. C5:I40).",
    "Data › Data Validation › Settings tab.",
    "Allow: List   |   Source: 0,1,2,3,4",
    "Keep “In-cell dropdown” ticked.",
    "Error Alert tab: Stop, “Enter a level from 0 to 4”.",
  ], M, 1.95, 5.6, 0.85, "step5 steps");
  // mock dialog
  const dx = 7.0, dy = 1.9, dw = 5.6, dh = 4.6;
  s.addShape(pres.shapes.RECTANGLE, { x: dx, y: dy, w: dw, h: dh, fill: { color: "FFFFFF" }, line: { color: "A6A6A6", width: 1 }, shadow: { type: "outer", blur: 8, offset: 3, angle: 90, color: "000000", opacity: 0.2 }, objectName: "dialog frame" });
  s.addShape(pres.shapes.RECTANGLE, { x: dx, y: dy, w: dw, h: 0.5, fill: { color: "F2F2F2" }, line: { color: "A6A6A6", width: 1 }, objectName: "dialog titlebar" });
  s.addText("Data Validation", { x: dx + 0.2, y: dy, w: 3, h: 0.5, fontSize: 14, color: H.dk1, valign: "middle", margin: 0, isTextBox: true, objectName: "dialog title" });
  ["Settings", "Input Message", "Error Alert"].forEach((t, i) => {
    s.addText(t, { x: dx + 0.2 + i * 1.6, y: dy + 0.65, w: 1.5, h: 0.4, fontSize: 13, bold: i === 0, color: i === 0 ? H.accent1 : H.accent5, align: "center", valign: "middle", margin: 0, isTextBox: true, line: i === 0 ? { color: H.accent1, width: 1 } : undefined, objectName: `dialog tab ${i}` });
  });
  s.addText("Allow:", { x: dx + 0.3, y: dy + 1.35, w: 2, h: 0.35, fontSize: 13, color: H.dk1, margin: 0, isTextBox: true, objectName: "dialog allow label" });
  s.addText("List                                ▾", { x: dx + 0.3, y: dy + 1.75, w: 3.2, h: 0.45, fontSize: 14, color: H.dk1, valign: "middle", margin: [0, 6, 0, 6], isTextBox: true, line: { color: "7F7F7F", width: 1 }, objectName: "dialog allow box" });
  s.addText("☑  In-cell dropdown", { x: dx + 0.3, y: dy + 2.35, w: 3, h: 0.35, fontSize: 13, color: H.dk1, margin: 0, isTextBox: true, objectName: "dialog checkbox" });
  s.addText("Source:", { x: dx + 0.3, y: dy + 2.85, w: 2, h: 0.35, fontSize: 13, color: H.dk1, margin: 0, isTextBox: true, objectName: "dialog source label" });
  s.addText("0,1,2,3,4", { x: dx + 0.3, y: dy + 3.25, w: dw - 0.6, h: 0.45, fontSize: 15, bold: true, color: H.accent1, fontFace: "Courier New", valign: "middle", margin: [0, 6, 0, 6], isTextBox: true, line: { color: H.accent2, width: 2 }, objectName: "dialog source box" });
  ["OK", "Cancel"].forEach((t, i) => {
    s.addText(t, { x: dx + dw - 2.5 + i * 1.15, y: dy + dh - 0.65, w: 1.0, h: 0.4, fontSize: 13, color: i === 0 ? "FFFFFF" : H.dk1, fill: { color: i === 0 ? H.accent1 : "F2F2F2" }, align: "center", valign: "middle", margin: 0, isTextBox: true, line: { color: "A6A6A6", width: 0.75 }, objectName: `dialog btn ${t}` });
  });
  s.addNotes("Data Validation stops typos like 33 or 'yes' getting into the matrix. Everyone picks from the same 0 to 4 list, which keeps the formulas reliable.");
}

// 9. Step 6 – conditional formatting
{
  const s = content("Build the matrix", "STEP 6", "Colour the levels with Conditional Formatting");
  const mk = (colored) => {
    const head = [cell("Name", { bold: true, fill: { color: H.accent1 }, color: "FFFFFF" }), ...SHORT.slice(0, 5).map((t) => cell(t, { bold: true, fill: { color: H.accent1 }, color: "FFFFFF", align: "center", fontSize: 11 }))];
    return [head, ...PEOPLE.slice(0, 6).map((p, i) => [cell(p), ...SCORES[i].slice(0, 5).map((v) => cell(v, colored ? { align: "center", bold: true, fill: { color: LEVEL_FILL[v] }, color: LEVEL_TXT[v] } : { align: "center" }))])];
  };
  const L6 = ["A", "C", "D", "E", "F", "G"], cw6 = [1.2, ...SHORT.slice(0, 5).map(() => 0.8)];
  s.addText("Before", { x: M, y: 1.75, w: 3, h: 0.35, fontSize: 15, bold: true, color: H.accent5, margin: 0, isTextBox: true, objectName: "cf before label" });
  excelGrid(s, M, 2.15, cw6, 0.4, mk(false), 4, "cf before grid", L6);
  s.addText("After", { x: 6.93, y: 1.75, w: 3, h: 0.35, fontSize: 15, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "cf after label" });
  excelGrid(s, 6.93, 2.15, cw6, 0.4, mk(true), 4, "cf after grid", L6);
  s.addText("→", { x: 6.25, y: 3.3, w: 0.6, h: 0.6, fontSize: 32, bold: true, color: C.accent2, align: "center", margin: 0, isTextBox: true, objectName: "cf arrow" });
  const HOW = ["Select the score cells (C5:I40)", "Home › Conditional Formatting › New Rule", "“Format only cells that contain” › Cell Value = 0 › red fill", "Repeat for levels 1, 2, 3 and 4 using the legend colours"];
  const hw = (W - 2 * M - 3 * 0.3) / 4;
  HOW.forEach((t, i) => {
    const x = M + i * (hw + 0.3), y = 5.7;
    card(s, x, y, hw, 1.15, `cf how card ${i + 1}`);
    badge(s, x + 0.2, y + 0.25, i + 1, `cf how ${i + 1}`, 0.38, C.accent3);
    s.addText(t, { x: x + 0.75, y: y + 0.18, w: hw - 0.95, h: 0.9, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: `cf how text ${i + 1}` });
  });
  s.addNotes("Colours turn a wall of numbers into a picture. Red and amber show gaps immediately. Use one rule per level so the colours match the legend exactly.");
}

// 10. Step 7 – formulas
pres.addSection({ title: "Analyse & maintain" });
{
  const s = content("Analyse & maintain", "STEP 7", "Add formulas that turn scores into gaps");
  const fw = 5.8;
  formulaBox(s, M, 1.85, fw, "Qualified people plant-wide (row 303)", '=COUNTIF(C$5:C$300,">=3")', "C303", "f1");
  formulaBox(s, M, 2.95, fw, "Gap against target (row 304)", "=C303-C302", "C304", "f2");
  formulaBox(s, M, 4.05, fw, "Status flag (row 305)", '=IF(C304<0,"GAP","OK")', "C305", "f3");
  formulaBox(s, M, 5.15, fw, "Versatility per person (column J)", '=COUNTIF(C5:I5,">=3")', "J5", "f4");
  // summary rows grid
  const lab = (t) => cell(t, { bold: true, fill: { color: "F2F2F2" } });
  const rows = [
    [cell(""), ...SHORT.map((t) => cell(t, { bold: true, fill: { color: H.accent1 }, color: "FFFFFF", align: "center", fontSize: 11 }))],
    [lab("Target"), ...TARGET.map((v) => cell(v, { align: "center" }))],
    [lab("Qualified"), ...QUAL.map((v) => cell(v, { align: "center", bold: true }))],
    [lab("Gap"), ...QUAL.map((v, j) => cell(v - TARGET[j], { align: "center", color: v - TARGET[j] < 0 ? "C0504D" : H.dk1 }))],
    [lab("Status"), ...QUAL.map((v, j) => { const g = v < TARGET[j]; return cell(g ? "GAP" : "OK", { align: "center", bold: true, fontSize: 11, fill: { color: g ? LEVEL_FILL[0] : LEVEL_FILL[3] }, color: g ? LEVEL_TXT[0] : LEVEL_TXT[3] }); })],
  ];
  excelGrid(s, 6.85, 1.85, [1.05, ...SHORT.map(() => 0.63)], 0.5, rows, 301, "f grid", "BCDEFGHIJ".split(""));
  s.addText([
    { text: "Copy across with the fill handle. ", options: { bold: true } },
    { text: "Press F4 to lock references (e.g. $C$5) when a formula must always point to the same cell." },
  ], { x: 6.85, y: 4.75, w: W - M - 6.85, h: 1.0, fontSize: 15, color: C.text1, margin: 0, isTextBox: true, objectName: "f tip" });
  s.addNotes("COUNTIF counts how many people are at level 3 or above. Subtract the target to get the gap, and an IF formula flags the skill as GAP or OK. Across the plant, Packaging, Allergen Control and CCP Monitoring are short.");
}

// Step 7b – department roll-up
{
  const s = content("Analyse & maintain", "STEP 7 (CONTINUED)", "Roll it up by department");
  formulaBox(s, M, 1.85, 5.6, "Qualified per department and skill", '=COUNTIFS(Dept,$A5,Matrix!C$5:C$300,">=3")', "B5", "fd1", 12);
  bulletList(s, [
    [{ text: "New sheet “By Department”: ", options: { bold: true } }, { text: "departments down column A, skills across row 4. Name Matrix!B5:B300 “Dept” (Formulas › Define Name)." }],
    [{ text: "COUNTIFS ", options: { bold: true } }, { text: "counts people in that department who are level 3 or 4." }],
    [{ text: "Red cells ", options: { bold: true } }, { text: "show a department relying on one person – plan cross-training from other areas." }],
    [{ text: "Quick option: ", options: { bold: true } }, { text: "Insert › PivotTable, or add a Slicer on Department to the Matrix table." }],
  ], M, 3.05, 5.6, 3.6, "fd bullets", 15);
  const fillFor = (v) => (v === 0 ? "F2F2F2" : v === 1 ? LEVEL_FILL[0] : v === 2 ? LEVEL_FILL[2] : LEVEL_FILL[3]);
  const txtFor = (v) => (v === 0 ? "A6A6A6" : v === 1 ? LEVEL_TXT[0] : v === 2 ? LEVEL_TXT[2] : LEVEL_TXT[3]);
  const rows = [
    [cell("Department", { bold: true, fill: { color: H.accent1 }, color: "FFFFFF" }), ...SHORT.map((t) => cell(t, { bold: true, fill: { color: H.accent1 }, color: "FFFFFF", align: "center", fontSize: 11 }))],
    ...DEPTS.map((d, i) => [cell(d), ...DEPT_QUAL[i].map((v) => cell(v === 0 ? "–" : v, { align: "center", bold: v > 0, fill: { color: fillFor(v) }, color: txtFor(v) }))]),
    [cell("Plant total", { bold: true, fill: { color: "F2F2F2" } }), ...QUAL.map((v) => cell(v, { align: "center", bold: true, fill: { color: "F2F2F2" } }))],
  ];
  excelGrid(s, 6.45, 1.85, [1.25, ...SHORT.map(() => 0.66)], 0.5, rows, 4, "fd grid");
  s.addText("–  skill not used in that department     red = only 1 qualified person", { x: 6.45, y: 6.5, w: 6.25, h: 0.35, fontSize: 12, italic: true, color: H.accent5, margin: 0, isTextBox: true, objectName: "fd caption" });
  s.addNotes("Because the matrix is plant-wide, one COUNTIFS formula rolls it up by department. This shows where a department depends on a single person, for example Maintenance on Mixer Setup, and where we can borrow trained people from another area.");
}

// 11. Step 8a – chart
{
  const s = content("Analyse & maintain", "STEP 8", "Chart the gaps for your team meeting");
  s.addChart(pres.charts.BAR, [
    { name: "Qualified (level 3–4)", labels: SKILLS, values: QUAL },
    { name: "Target", labels: SKILLS, values: TARGET },
  ], {
    x: M, y: 1.75, w: 8.2, h: 4.95, barDir: "col", barGapWidthPct: 60,
    chartColors: [H.accent1, "C9D3CD"],
    showTitle: true, title: "Plant-wide: qualified people per skill vs target", titleFontSize: 14, titleColor: H.dk2, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: H.dk1, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: "595959", valAxisLabelColor: "595959", catAxisLabelFontSize: 11, valAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: "E5E5E5", size: 0.75 }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 22, valAxisMajorUnit: 4,
    showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: H.dk1,
    objectName: "gap chart",
  });
  steps(s, [
    "Select skill names + Qualified and Target rows.",
    "Insert › Recommended Charts › Clustered Column.",
    "Add a title and data labels (Chart Design › Add Chart Element).",
    "Put it on a “Dashboard” sheet for the plant meeting.",
  ], 9.25, 1.95, W - M - 9.25, 1.15, "chart steps");
  s.addNotes("A simple clustered column chart shows where qualified headcount is below target. Across the plant, Allergen Control, Packaging and CCP Monitoring need training first.");
}

// 12. Step 8b – protect & maintain
{
  const s = content("Analyse & maintain", "STEP 8 (CONTINUED)", "Protect it and keep it alive");
  const items = [
    ["Protect the sheet", "Unlock only the score cells (Ctrl + 1 › Protection), then Review › Protect Sheet with a password."],
    ["Track expiry dates", "Add a “Last trained” sheet:  =IF(TODAY()-D5>365,\"Retrain\",\"OK\")  to flag annual refreshers."],
    ["Review on a schedule", "Each department supervisor updates their rows monthly; plant leadership reviews gaps quarterly."],
    ["Control the document", "One master file on SharePoint / Teams for the plant; record revision date and approver."],
  ];
  const cw = (W - 2 * M - 0.4) / 2, ch = 2.1;
  items.forEach(([h, b], i) => {
    const x = M + (i % 2) * (cw + 0.4), y = 1.85 + Math.floor(i / 2) * (ch + 0.35);
    card(s, x, y, cw, ch, `maint card ${i + 1}`);
    badge(s, x + 0.35, y + 0.35, i + 1, `maint ${i + 1}`);
    s.addText(h, { x: x + 1.15, y: y + 0.35, w: cw - 1.5, h: 0.55, fontSize: 20, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true, objectName: `maint head ${i + 1}` });
    s.addText(b, { x: x + 1.15, y: y + 0.95, w: cw - 1.5, h: 1.0, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: `maint body ${i + 1}` });
  });
  s.addNotes("A matrix is only useful if it is current. Protect the formulas, flag expired training automatically, and make updating it part of the supervisors' monthly routine.");
}

// 13. Shortcuts
{
  const s = content("Analyse & maintain", "QUICK REFERENCE", "Excel shortcuts that save the most time");
  const SC = [
    ["Ctrl + T", "Turn a range into a Table"],
    ["Ctrl + Shift + L", "Turn filters on / off"],
    ["Alt, A, V, V", "Open Data Validation"],
    ["Alt, H, L", "Open Conditional Formatting"],
    ["Ctrl + 1", "Format Cells (incl. lock / unlock)"],
    ["F4", "Toggle absolute reference ($A$1)"],
    ["Ctrl + D", "Fill a formula down"],
    ["Alt, W, F, F", "Freeze Panes"],
  ];
  const cw = (W - 2 * M - 0.4) / 2;
  SC.forEach(([k, d], i) => {
    const x = M + (i % 2) * (cw + 0.4), y = 1.9 + Math.floor(i / 2) * 1.15;
    card(s, x, y, cw, 0.9, `sc card ${i + 1}`);
    s.addText(k, { x: x + 0.25, y: y + 0.17, w: 2.4, h: 0.56, fontSize: 16, bold: true, color: "FFFFFF", fill: { color: H.accent1 }, align: "center", valign: "middle", fontFace: "Courier New", margin: 0, isTextBox: true, objectName: `sc key ${i + 1}` });
    s.addText(d, { x: x + 2.95, y: y + 0.1, w: cw - 3.2, h: 0.7, fontSize: 16, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: `sc desc ${i + 1}` });
  });
  s.addNotes("These are the shortcuts I used most while building the matrix. Keep this slide handy as a cheat sheet.");
}

// 14. Closing
pres.addSection({ title: "Wrap-up" });
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Wrap-up" });
  s.addText("One plant, one skill matrix", { placeholder: "title" });
  s.addText("A living plant-wide matrix shows coverage in every department, drives one training plan and proves competence in every audit.", { placeholder: "body" });
  const NX = [
    ["Roll out", "Load every department into the matrix this quarter"],
    ["Rate", "Each supervisor scores their team on the 0–4 legend"],
    ["Act", "Turn plant-wide GAP flags into one training plan"],
  ];
  NX.forEach(([h, b], i) => {
    const x = 8.6, y = 1.5 + i * 1.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 4.13, h: 1.3, rectRadius: 0.1, fill: { color: "1A4D35" }, line: { color: "1A4D35" }, objectName: `next card ${i + 1}` });
    badge(s, x + 0.3, y + 0.37, i + 1, `next ${i + 1}`, 0.55, C.accent2);
    s.addText(h, { x: x + 1.1, y: y + 0.2, w: 2.9, h: 0.4, fontSize: 18, bold: true, color: "FFFFFF", margin: 0, isTextBox: true, objectName: `next head ${i + 1}` });
    s.addText(b, { x: x + 1.1, y: y + 0.6, w: 2.9, h: 0.6, fontSize: 14, color: "CFE3D6", valign: "top", margin: 0, isTextBox: true, objectName: `next body ${i + 1}` });
  });
  s.addNotes("To wrap up: load every department, have each supervisor rate with the same legend, and use the plant-wide GAP flags to build one training plan. Questions?");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
