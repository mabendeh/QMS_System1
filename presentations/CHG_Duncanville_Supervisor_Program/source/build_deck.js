// Supervisor Productivity & Engagement Program — kickoff deck builder (v2, executive revision)
// C.H. Guenther & Son, Duncanville. Run: PPTX_SKILL=<skill dir> node build_deck.js <outdir>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const { applyTheme } = require(path.join(process.env.PPTX_SKILL, "scripts/apply_theme.js"));
const OUT = process.argv[2] || ".";
const FILE = path.join(OUT, "CHG_Duncanville_Supervisor_Productivity_Program_Kickoff.pptx");

const THEME = {
  name: "CHG Duncanville CI",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1A2433", lt1: "FFFFFF", dk2: "0B2545", lt2: "EEF2F6",
    accent1: "13315C", accent2: "2F9E6E", accent3: "5B7FA6", accent4: "8A97A8",
    accent5: "B9770E", accent6: "D6DEE8", hlink: "13315C", folHlink: "5B7FA6",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Continuous Improvement — Duncanville";
pres.company = "C.H. Guenther &amp; Son";
pres.title = "Supervisor Productivity & Engagement Program — Kickoff";
pres.subject = "Program launch, November 17, 2026";
const C = pres.SchemeColor;
const NAVY = C.text2, INK = C.text1, WHITE = C.background1, LIGHT = C.background2;
const MID = C.accent1, GREEN = C.accent2, STEEL = C.accent3, GRAY = C.accent4, AMBER = C.accent5, PALE = C.accent6;

// ---------- icons ----------
const iconCache = {};
async function icon(name, hex) {
  const key = name + hex;
  if (iconCache[key]) return iconCache[key];
  if (!fa[name]) throw new Error("missing icon " + name);
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + hex, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return (iconCache[key] = "image/png;base64," + buf.toString("base64"));
}
const shadow = () => ({ type: "outer", color: "0B2545", opacity: 0.14, blur: 6, offset: 2, angle: 90 });

// ---------- layouts (brand frame) ----------
const BRAND = "C.H. GUENTHER & SON  |  DUNCANVILLE";
const PROGRAM = "Supervisor Productivity & Engagement Program  ·  Kickoff, November 17, 2026";
const LOGO = "Logo placeholder (replace with official C.H. Guenther logo)";
pres.defineSlideMaster({
  title: "CONTENT_LIGHT",
  background: { color: "FFFFFF" },
  objects: [
    { text: { text: BRAND, options: { x: 0.6, y: 7.02, w: 5, h: 0.3, fontSize: 9, bold: true, color: NAVY, charSpacing: 1, margin: 0, objectName: LOGO } } },
    { text: { text: PROGRAM, options: { x: 5.2, y: 7.02, w: 6.6, h: 0.3, fontSize: 9, color: GRAY, align: "right", margin: 0 } } },
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.32, w: 12.1, h: 0.34, fontSize: 12, bold: true, color: GREEN, charSpacing: 2, margin: 0, valign: "middle" }, text: "SECTION" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.66, w: 12.1, h: 0.8, fontSize: 32, bold: true, color: NAVY, margin: 0, valign: "middle", align: "left" }, text: "Headline" } },
  ],
  slideNumber: { x: 12.15, y: 7.02, w: 0.6, h: 0.3, fontSize: 9, color: GRAY, align: "right" },
});
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { path: path.join(__dirname, "hero.png") },
  objects: [
    { text: { text: BRAND, options: { x: 0.6, y: 0.45, w: 6, h: 0.35, fontSize: 12, bold: true, color: WHITE, charSpacing: 2, margin: 0, objectName: LOGO } } },
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 1.15, w: 8.5, h: 0.4, fontSize: 14, bold: true, color: GREEN, charSpacing: 2, margin: 0 }, text: "KICKER" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.6, w: 8.6, h: 1.6, fontSize: 46, bold: true, color: WHITE, margin: 0, valign: "top", align: "left" }, text: "Title" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 0.6, y: 3.25, w: 8.6, h: 0.5, fontSize: 22, color: PALE, margin: 0 }, text: "Subtitle" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION_DARK",
  background: { path: path.join(__dirname, "section.png") },
  objects: [
    { text: { text: BRAND, options: { x: 0.6, y: 7.02, w: 5, h: 0.3, fontSize: 9, bold: true, color: PALE, charSpacing: 1, margin: 0, objectName: LOGO } } },
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.5, w: 12.1, h: 0.34, fontSize: 12, bold: true, color: GREEN, charSpacing: 2, margin: 0 }, text: "KICKER" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.86, w: 12.1, h: 0.9, fontSize: 36, bold: true, color: WHITE, margin: 0, valign: "middle", align: "left" }, text: "Title" } },
  ],
  slideNumber: { x: 12.15, y: 7.02, w: 0.6, h: 0.3, fontSize: 9, color: PALE, align: "right" },
});

// ---------- helpers ----------
function newSlide(layout, section, kicker, title) {
  const s = pres.addSlide({ masterName: layout, sectionTitle: section });
  if (kicker) s.addText(kicker, { placeholder: "kicker" });
  if (title) s.addText(title, { placeholder: "title" });
  return s;
}
function T(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, margin: 0, color: INK, fontSize: 14, valign: "top" }, o)); }
function card(s, x, y, w, h, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.08, fill: { color: o.fill || LIGHT }, line: o.line ? { color: o.line, width: o.lineW || 1, dashType: o.dash || "solid" } : { type: "none" } }, o.shadow ? { shadow: shadow() } : {}, o.name ? { objectName: o.name } : {}));
}
async function iconCircle(s, name, x, y, d, fill, iconHex = "FFFFFF") {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  const p = d * 0.26;
  s.addImage({ data: await icon(name, iconHex), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: name.replace(/^Fa/, "") + " icon" });
}
function pill(s, x, y, w, text, fill, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: o.h || 0.3, rectRadius: (o.h || 0.3) / 2, fill: { color: fill }, line: { type: "none" } });
  T(s, text, { x, y, w, h: o.h || 0.3, fontSize: o.fontSize || 10, bold: true, color: o.color || WHITE, align: "center", valign: "middle", charSpacing: 1 });
}
function illus(s, x, y, w = 3.3, label = "ILLUSTRATIVE EXAMPLE — NOT PLANT DATA") { pill(s, x, y, w, label, AMBER); }
function arrowRight(s, x, y, w, color = GRAY) { s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width: 2, endArrowType: "triangle" } }); }
function bullets(items, o = {}) { return items.map((t, k) => ({ text: t, options: Object.assign({ bullet: true, breakLine: k < items.length - 1 }, o) })); }

const DEPTS = [["FaIndustry", "Production"], ["FaClipboardCheck", "Quality"], ["FaWrench", "Maintenance"], ["FaDollarSign", "Finance"], ["FaSyncAlt", "Continuous Improvement"]];

(async () => {
  // ================= 1. TITLE =================
  pres.addSection({ title: "Launch" });
  let s = newSlide("TITLE_DARK", "Launch", "PROGRAM LAUNCH  ·  TUESDAY, NOVEMBER 17, 2026", "Supervisor Productivity & Engagement Program");
  s.addText("Small Improvements. Stronger Teams. Measurable Results.", { placeholder: "subtitle" });
  T(s, "C.H. Guenther & Son — Duncanville", { x: 0.6, y: 3.8, w: 8.6, h: 0.4, fontSize: 16, bold: true, color: WHITE });
  card(s, 8.95, 0.42, 3.8, 0.62, { fill: NAVY, name: "Department icon strip" });
  for (let i = 0; i < 5; i++) await iconCircle(s, DEPTS[i][0], 9.12 + i * 0.72, 0.5, 0.46, i === 4 ? GREEN : MID);
  s.addNotes(`(About 1 minute.)

Good morning, everyone, and thanks for being here. I know what it takes to pull supervisors, Quality, Maintenance and Finance into the same room at the same time, so I'll make these 45 minutes worth it.

Today we're launching the Supervisor Productivity and Engagement Program. Small improvements, stronger teams, measurable results.

Let me say up front what this is NOT. It's not another report to fill out, and it's not CI telling you how to run your line. It's about giving supervisors the support, tools and time to fix the things that frustrate you and your crews every shift, and making sure the people who do that work get seen for it.

The five icons up top are the five groups in this room: Production, Quality, Maintenance, Finance and CI. Every one of them has a defined role, and by the end of today nobody should be unsure what theirs is.

Ask the room: who has something on their line they've wanted to fix for months but never had the time or backing for? Hold on to it. We'll use it later.`);

  // ================= 2. THE DEAL =================
  pres.addSection({ title: "Program Overview" });
  s = newSlide("CONTENT_LIGHT", "Program Overview", "WHAT YOU'LL LEAVE WITH TODAY", "Lead one improvement. We back you. Leadership sees it.");
  const deal = [
    ["FaHardHat", "Our expectations", "of each supervisor", NAVY, [["FaFlag", "Lead one improvement per cell"], ["FaRulerCombined", "Measure before and after"], ["FaUsers", "Involve your crew in the fix"], ["FaMicrophone", "Present results at WOr"]]],
    ["FaHandsHelping", "Your support", "from CI, partners, leadership", MID, [["FaSyncAlt", "Weekly CI coaching at the line"], ["FaClipboardCheck", "Quality and Maintenance in every cell"], ["FaDollarSign", "Finance helps value the results"], ["FaUserTie", "Leadership protects time, removes barriers"]]],
    ["FaTrophy", "Your recognition", "for results your team delivers", GREEN, [["FaUserCheck", "Contributors credited by name"], ["FaUserTie", "Your team presents to leadership"], ["FaShareAlt", "Wins shared across all shifts"], ["FaAward", "Day-90 team recognition"]]],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15, d = deal[i];
    card(s, x, 1.7, 3.8, 4.95, { fill: LIGHT, name: d[1] + " column" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.7, w: 3.8, h: 1.2, rectRadius: 0.08, fill: { color: d[3] }, line: { type: "none" }, objectName: d[1] + " header" });
    await iconCircle(s, d[0], x + 0.25, 1.92, 0.76, WHITE, i === 2 ? HEX.accent2 : i === 1 ? HEX.accent1 : HEX.dk2);
    T(s, d[1], { x: x + 1.2, y: 1.92, w: 2.5, h: 0.42, fontSize: 20, bold: true, color: WHITE });
    T(s, d[2], { x: x + 1.2, y: 2.34, w: 2.5, h: 0.35, fontSize: 13, color: PALE });
    for (let j = 0; j < 4; j++) {
      const y = 3.15 + j * 0.84;
      s.addImage({ data: await icon(d[4][j][0], i === 2 ? HEX.accent2 : HEX.accent1), x: x + 0.3, y: y + 0.1, w: 0.38, h: 0.38, altText: d[4][j][1] });
      T(s, d[4][j][1], { x: x + 0.9, y, w: 2.75, h: 0.6, fontSize: 15, bold: j === 0, color: INK, valign: "middle" });
    }
  }
  s.addNotes(`(About 2 minutes.)

Before we get into any detail, here is the whole program on one slide. If you only remember one slide from today, make it this one.

Left column, what we expect from each supervisor: lead one improvement per cell, measure it before and after, involve your crew in the fix, and present the result at the WOr leadership meeting. That's it. One project at a time.

Middle column, how you'll be supported. I'll coach every cell weekly, at the line, not in a conference room. Quality and Maintenance are built into every cell, so you're not chasing them. Randy in Finance will help put a value on your results. And plant leadership has committed to protecting your time and clearing barriers.

Right column, how you'll be recognized. The people who contribute get credited by name. Your team presents its own results to leadership. Wins get shared across all three shifts. And at day 90 we recognize the teams that delivered.

Everything else today is the detail behind these three columns.

Ask the room: looking at the left column, what worries you most? Time? Data? Getting your crew on board? Let's get that on the table now.`);

  // ================= 3. WHY =================
  s = newSlide("CONTENT_LIGHT", "Program Overview", "WHY WE ARE DOING THIS", "Continuous improvement belongs to everyone");
  const goals = [
    ["FaHardHat", "Empower supervisors", "Supervisors lead real improvements on their own lines.", "Measured by: projects led and presented", NAVY],
    ["FaUsers", "Engage our people", "Operators, Quality and Maintenance solve problems together.", "Measured by: ideas raised, people involved", MID],
    ["FaChartLine", "Improve performance", "Better productivity, quality and cost, proven with data.", "Measured by: minutes, scrap, validated $", GREEN],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15;
    card(s, x, 1.75, 3.8, 4.6, { fill: LIGHT, name: "Goal " + (i + 1) });
    await iconCircle(s, goals[i][0], x + 0.35, 2.05, 1.1, goals[i][4]);
    T(s, "0" + (i + 1), { x: x + 2.3, y: 2.1, w: 1.2, h: 0.9, fontSize: 44, bold: true, color: PALE, align: "right" });
    T(s, goals[i][1], { x: x + 0.35, y: 3.4, w: 3.2, h: 0.5, fontSize: 22, bold: true, color: NAVY });
    T(s, goals[i][2], { x: x + 0.35, y: 3.95, w: 3.15, h: 1.2, fontSize: 18 });
    s.addShape(pres.shapes.LINE, { x: x + 0.35, y: 5.3, w: 3.1, h: 0, line: { color: PALE, width: 1 } });
    T(s, goals[i][3], { x: x + 0.35, y: 5.45, w: 3.2, h: 0.6, fontSize: 13, italic: true, color: STEEL });
  }
  s.addNotes(`(About 1.5 minutes.)

Three goals, and each one has a way we'll measure it.

One: empower supervisors. You know your lines and your people better than anyone. We want supervisors leading improvements, with CI supporting, not the other way around. We'll measure that simply: projects led and presented.

Two: engage our people. The operators see the waste every day: the walking, the waiting, the searching. When we ask them and act on what they say, people take ownership. We'll measure ideas raised and people involved.

Three: improve performance. This has to show up in productivity, quality and cost: changeover minutes, downtime, scrap, and dollars that Finance has validated.

And the headline: continuous improvement belongs to everyone. CI is a support function. The improvement happens on your floor.

Ask the room: what does improvement look like on your shift today? Formal, informal, or not happening? What gets in the way? (Write the barriers down. They're useful later.)`);

  // ================= 4. VISION =================
  s = newSlide("CONTENT_LIGHT", "Program Overview", "OUR VISION", "One repeatable cycle, and every loop builds a leader");
  const vision = [["FaSearch", "Identify", "A real problem on the floor"], ["FaHandshake", "Engage", "The crew helps design the fix"], ["FaCogs", "Improve", "Test, adjust, implement"], ["FaChartLine", "Measure", "Before vs. after data"], ["FaLock", "Sustain", "Standard work so it sticks"], ["FaTrophy", "Recognize", "Credit the team publicly"]];
  const vw = 1.75, vg = (12.1 - 6 * vw) / 5;
  for (let i = 0; i < 6; i++) {
    const x = 0.6 + i * (vw + vg);
    await iconCircle(s, vision[i][0], x + (vw - 1.1) / 2, 1.85, 1.1, i === 5 ? GREEN : i % 2 ? MID : NAVY);
    T(s, vision[i][1], { x, y: 3.1, w: vw, h: 0.42, fontSize: 19, bold: true, color: NAVY, align: "center" });
    T(s, vision[i][2], { x, y: 3.52, w: vw, h: 0.65, fontSize: 13, align: "center" });
    if (i < 5) arrowRight(s, x + vw / 2 + 0.62, 2.4, vg + vw - 1.24, STEEL);
  }
  s.addShape(pres.shapes.LINE, { x: 1.475, y: 4.45, w: 10.35, h: 0, line: { color: PALE, width: 2, dashType: "dash", beginArrowType: "triangle" } });
  T(s, "…then repeat, with a bigger problem each time", { x: 4.4, y: 4.52, w: 4.5, h: 0.3, fontSize: 12, italic: true, color: GRAY, align: "center" });
  card(s, 0.6, 5.2, 12.1, 1.25, { fill: NAVY, name: "Leadership skills" });
  T(s, "Leadership skills each loop builds", { x: 0.9, y: 5.35, w: 4.5, h: 0.35, fontSize: 15, bold: true, color: GREEN });
  const skills = ["Solving problems with data", "Coaching and involving a crew", "Presenting results to leadership"];
  for (let i = 0; i < 3; i++) {
    const x = 0.9 + i * 3.95;
    s.addImage({ data: await icon("FaCheckCircle", HEX.accent2), x, y: 5.83, w: 0.3, h: 0.3, altText: "Check" });
    T(s, skills[i], { x: x + 0.42, y: 5.78, w: 3.4, h: 0.4, fontSize: 15, color: WHITE, valign: "middle" });
  }
  s.addNotes(`(About 1.5 minutes.)

Every project follows the same simple cycle: identify, engage, improve, measure, sustain, recognize.

Identify a real problem your team lives with, not something off a corporate list. Engage the crew, so they help design the fix. Improve: test it, adjust it, implement it. Measure before and after. Sustain it with standard work so it doesn't slide back in three weeks. And recognize the team publicly.

Then you go round again, usually with a slightly bigger problem.

The navy bar is the part I care about most. Every loop is a leadership rep: solving problems with data, coaching a crew, presenting results to leadership. Those are the skills that move people up in this business. So yes, we want the productivity, but we're also building the next generation of plant leaders.

Ask the room: of these six, which one are we weakest at as a plant today? My bet is Sustain. Tell me if I'm wrong.`);

  // ================= 5. CELLS =================
  s = newSlide("CONTENT_LIGHT", "Program Overview", "STRUCTURE", "Three shift cells. CI coaches. Finance validates.");
  card(s, 0.6, 1.7, 12.1, 0.62, { fill: GREEN, name: "CI support band" });
  await iconCircle(s, "FaSyncAlt", 0.78, 1.76, 0.5, NAVY);
  T(s, [{ text: "Continuous Improvement  ", options: { bold: true } }, { text: "coaches every cell: tools, methods, removing barriers" }], { x: 1.45, y: 1.7, w: 11.0, h: 0.62, fontSize: 16, color: WHITE, valign: "middle" });
  const cells = ["A Shift", "B Shift", "C Shift"];
  const members = [["FaIndustry", "Production supervisor", "Project owner"], ["FaClipboardCheck", "Quality", "Food safety and specs"], ["FaWrench", "Maintenance", "Equipment expertise"]];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15;
    s.addShape(pres.shapes.LINE, { x: x + 1.9, y: 2.32, w: 0, h: 0.28, line: { color: GREEN, width: 2 } });
    card(s, x, 2.6, 3.8, 2.6, { fill: LIGHT, shadow: true, name: cells[i] + " improvement cell" });
    T(s, cells[i] + " Improvement Cell", { x: x + 0.25, y: 2.72, w: 3.3, h: 0.4, fontSize: 17, bold: true, color: NAVY });
    for (let j = 0; j < 3; j++) {
      const y = 3.2 + j * 0.64;
      card(s, x + 0.2, y, 3.4, 0.54, { fill: j === 0 ? PALE : WHITE });
      await iconCircle(s, members[j][0], x + 0.3, y + 0.07, 0.4, j === 0 ? NAVY : MID);
      T(s, [{ text: members[j][1], options: { bold: true, fontSize: 14, color: NAVY, breakLine: true } }, { text: members[j][2], options: { fontSize: 11, color: STEEL } }], { x: x + 0.82, y: y + 0.02, w: 2.7, h: 0.5, valign: "middle" });
    }
    s.addShape(pres.shapes.LINE, { x: x + 1.9, y: 5.2, w: 0, h: 0.75, line: { color: STEEL, width: 2 } });
  }
  card(s, 1.6, 5.37, 10.1, 0.42, { fill: WHITE, line: STEEL, dash: "dash", name: "Cross-shift note" });
  T(s, [{ text: "Cross-shift projects: ", options: { bold: true, color: NAVY } }, { text: "one named owner, actions coordinated across A, B and C shift hand-offs" }], { x: 1.6, y: 5.37, w: 10.1, h: 0.42, fontSize: 13, align: "center", valign: "middle" });
  card(s, 0.6, 5.95, 12.1, 0.62, { fill: NAVY, name: "Finance validation band" });
  await iconCircle(s, "FaDollarSign", 0.78, 6.01, 0.5, STEEL);
  T(s, [{ text: "Finance (Randy)  ", options: { bold: true } }, { text: "agrees how results will be valued, then validates the savings" }], { x: 1.45, y: 5.95, w: 11.0, h: 0.62, fontSize: 16, color: WHITE, valign: "middle" });
  s.addNotes(`(About 2 minutes.)

Here's the structure. One improvement cell per shift: A, B and C.

Each cell has three core members: the production supervisor, who owns the project; someone from Quality, who protects food safety and specs; and someone from Maintenance, who brings the equipment know-how. Three people who between them can solve most problems on a line.

Across the top, CI supports all three cells: coaching, tools and templates, teaching methods like SMED and 5S, and, most importantly, removing barriers. If you're stuck waiting on a part, a decision or a resource, that's mine to chase.

Across the bottom is Finance. Randy agrees with you up front how the results will be valued, and then validates the savings at the end. More on that shortly.

The dotted box: some problems don't belong to one shift, like a changeover that runs across shift change. For those we name ONE owner and coordinate actions across the hand-offs, so nothing falls in the gap.

Ask the room: who do you naturally go to today for Quality and Maintenance on your shift? Those may well be your cell members.`);

  // ================= 6. ROLE COMMITMENTS =================
  s = newSlide("CONTENT_LIGHT", "Program Overview", "ROLES & COMMITMENTS", "The supervisor leads. Everyone else has a defined job.");
  card(s, 0.6, 1.7, 3.55, 4.95, { fill: NAVY, name: "Supervisor role card" });
  await iconCircle(s, "FaHardHat", 0.9, 1.98, 0.9, GREEN);
  pill(s, 1.95, 2.28, 1.4, "OWNS", GREEN);
  T(s, "Supervisor", { x: 0.9, y: 3.05, w: 3.0, h: 0.5, fontSize: 24, bold: true, color: WHITE });
  T(s, "Project owner", { x: 0.9, y: 3.52, w: 3.0, h: 0.35, fontSize: 14, color: PALE });
  T(s, bullets(["Picks and charters the project", "Owns baseline, actions, results", "Leads and credits the crew", "Presents at WOr"]), { x: 0.9, y: 4.05, w: 3.05, h: 2.3, fontSize: 15, color: WHITE, paraSpaceAfter: 8 });
  const rc = [
    ["FaUsers", "Production employees", "PARTICIPATE", MID, ["Raise ideas", "Own named actions", "Help write the standard"]],
    ["FaClipboardCheck", "Quality", "SAFEGUARDS", STEEL, ["Food-safety review of changes", "Spec and sanitation sign-off", "Member of every cell"]],
    ["FaWrench", "Maintenance", "ENABLES", STEEL, ["Checks equipment feasibility", "Makes approved changes", "Member of every cell"]],
    ["FaSyncAlt", "CI Manager", "COACHES", GREEN, ["Weekly coaching at the line", "Tools, templates, training", "Escalates and removes barriers"]],
    ["FaDollarSign", "Finance (Randy)", "VALIDATES", STEEL, ["Agrees baseline and method", "Validates against actuals", "Classifies the benefit"]],
    ["FaUserTie", "Plant leadership", "SPONSORS", MID, ["Protects supervisor time", "Clears resource barriers", "Recognizes results"]],
  ];
  for (let i = 0; i < 6; i++) {
    const x = 4.45 + (i % 3) * 2.8, y = 1.7 + Math.floor(i / 3) * 2.55;
    card(s, x, y, 2.65, 2.4, { fill: LIGHT, name: rc[i][1] + " role card" });
    await iconCircle(s, rc[i][0], x + 0.18, y + 0.18, 0.55, i === 3 ? GREEN : NAVY);
    pill(s, x + 0.9, y + 0.3, 1.55, rc[i][2], rc[i][3], { fontSize: 9 });
    T(s, rc[i][1], { x: x + 0.18, y: y + 0.85, w: 2.35, h: 0.38, fontSize: 15, bold: true, color: NAVY });
    T(s, bullets(rc[i][4]), { x: x + 0.18, y: y + 1.25, w: 2.4, h: 1.05, fontSize: 12, paraSpaceAfter: 2 });
  }
  s.addNotes(`(About 2.5 minutes.)

When everyone owns something, nobody owns it. So let's be specific.

The big navy card is the supervisor, and it's big on purpose. The supervisor owns the project: picks it, charters it, owns the baseline, the actions and the results, leads and credits the crew, and presents at WOr.

Around that, everyone else has a defined job.

Production employees participate: they raise ideas, own named actions, and help write the new standard. Some of the best ideas will come from them.

Quality safeguards: they review every change for food safety, sign off on anything touching specs or sanitation, and they're in every cell. We never trade food safety for minutes.

Maintenance enables: is this feasible and safe on this machine? Then they make the approved changes.

I coach: weekly at the line, tools and templates, and I escalate and remove barriers. I will not take your project and run it for you.

Randy validates: he agrees the baseline and valuation method at the start, validates against actual results, and classifies the benefit. Randy, thank you for being part of this.

And plant leadership sponsors: protects your time, clears resource barriers, and recognizes results.

Ask the room: does anything on here surprise anyone, or feel unclear? Better to sort it out now than three weeks in.`);

  // ================= 7. RACI =================
  s = newSlide("CONTENT_LIGHT", "Program Overview", "DECISION RIGHTS", "Who does it, who owns it, who signs off");
  const cols = ["Supervisor", "Operators", "Quality", "Maint.", "CI", "Finance", "Leadership"];
  const raci = [
    ["Raise improvement ideas", "A", "R", "C", "C", "C", "", "I"],
    ["Select and charter the project", "A", "C", "C", "C", "C", "I", "I"],
    ["Measure the baseline", "A", "R", "C", "C", "C", "C", ""],
    ["Agree baseline and savings method", "R", "", "", "", "C", "A", ""],
    ["Food-safety and quality review", "R", "I", "A", "C", "C", "", ""],
    ["Equipment changes", "R", "I", "C", "A", "C", "", ""],
    ["Implement and standardize", "A", "R", "R", "R", "C", "", "I"],
    ["Validate financial savings", "R", "", "", "", "C", "A", "I"],
    ["Present results at WOr", "A", "C", "C", "C", "C", "C", "I"],
    ["Remove barriers, provide resources", "R", "", "", "", "R", "", "A"],
    ["Recognize contributions", "R", "", "", "", "R", "", "A"],
  ];
  const cellStyle = { A: ["2F9E6E", "FFFFFF"], R: ["13315C", "FFFFFF"], C: ["D6DEE8", "0B2545"], I: ["EEF2F6", "5B7FA6"], "": ["FFFFFF", "FFFFFF"] };
  const head = [{ text: "Activity", options: { bold: true, color: "FFFFFF", fill: { color: "0B2545" } } }].concat(cols.map((c, k) => ({ text: c, options: { bold: true, align: "center", fontSize: 10, color: "FFFFFF", fill: { color: k === 0 ? "2F9E6E" : "0B2545" } } })));
  const rows = [head].concat(raci.map((r) => [{ text: r[0], options: { bold: true, color: "0B2545", fill: { color: "FFFFFF" } } }].concat(r.slice(1).map((v) => ({ text: v, options: { bold: true, align: "center", color: cellStyle[v][1], fill: { color: cellStyle[v][0] } } })))));
  s.addTable(rows, { x: 0.6, y: 1.65, w: 9.1, colW: [3.15].concat(Array(7).fill(0.85)), rowH: 0.385, fontSize: 12, valign: "middle", border: { type: "solid", pt: 1, color: "FFFFFF" }, margin: [2, 6, 2, 6] });
  const leg = [["A", "Accountable", "Owns the outcome. One per activity."], ["R", "Responsible", "Does the work."], ["C", "Consulted", "Gives input before decisions."], ["I", "Informed", "Kept up to date."]];
  for (let i = 0; i < 4; i++) {
    const y = 1.65 + i * 0.82;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.05, y, w: 0.5, h: 0.5, rectRadius: 0.06, fill: { color: cellStyle[leg[i][0]][0] }, line: i === 3 ? { color: "D6DEE8", width: 1 } : { type: "none" } });
    T(s, leg[i][0], { x: 10.05, y, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: cellStyle[leg[i][0]][1], align: "center", valign: "middle" });
    T(s, [{ text: leg[i][1], options: { bold: true, color: NAVY, fontSize: 14, breakLine: true } }, { text: leg[i][2], options: { fontSize: 12, color: GRAY } }], { x: 10.7, y: y - 0.05, w: 2.05, h: 0.65, valign: "middle" });
  }
  card(s, 10.05, 5.0, 2.65, 1.3, { fill: NAVY, name: "RACI takeaway" });
  T(s, "Supervisors are accountable for 5 of 11 activities. Quality, Maintenance and Finance each own their sign-off.", { x: 10.2, y: 5.05, w: 2.4, h: 1.2, fontSize: 12, color: WHITE, valign: "middle" });
  s.addNotes(`(About 2.5 minutes.)

This is the detail behind the role cards: who does it, who owns it, and who signs off. It's a RACI. A is accountable: the one person who owns the outcome, and there's exactly one A per row. R is responsible: does the work. C is consulted before decisions. I is kept informed.

Look down the green Supervisor column. You're accountable for raising and selecting ideas, the baseline, implementation, and presenting at WOr.

Now look at where the A is NOT the supervisor. That's deliberate.

Food safety and quality review: Quality owns that sign-off. If Quality says a change isn't safe, it doesn't happen.

Equipment changes: Maintenance owns that. Nobody modifies equipment without Maintenance.

The savings method and the validation: Finance owns those. We don't grade our own homework.

And removing barriers and recognition: plant leadership owns those, with me doing a lot of the legwork.

I'll leave a printed copy of this in each cell's folder. When there's a disagreement about who decides, this is the answer.

Ask the room: is there any row where you think the A is in the wrong place? Now's the time to challenge it.`);

  // ================= 8. START SMALL =================
  pres.addSection({ title: "Focus Areas" });
  s = newSlide("CONTENT_LIGHT", "Focus Areas", "IMPROVEMENT ROADMAP", "Start with quick wins, then earn the bigger projects");
  const stages = [
    ["STAGE 1", "Quick Wins", "5S, visual management, workplace organization, simple changeover fixes", "One area  ·  every cell starts here", "FaBroom", 1.2],
    ["STAGE 2", "Process Improvements", "Recurring downtime, setup losses, waste and rework", "One process or line", "FaCogs", 2.05],
    ["STAGE 3", "Larger Projects", "Cross-functional productivity, standardization, verified cost reduction", "Cross-functional", "FaChartLine", 2.9],
  ];
  const base = 6.2;
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.0, h = stages[i][5] + 0.85, y = base - h;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.85, h, fill: { color: [STEEL, MID, NAVY][i] }, line: { type: "none" }, objectName: stages[i][0] + " step" });
    await iconCircle(s, stages[i][4], x + 0.25, y + 0.25, 0.6, i === 2 ? GREEN : WHITE, i === 2 ? "FFFFFF" : [HEX.accent3, HEX.accent1][i]);
    T(s, stages[i][0], { x: x + 1.0, y: y + 0.22, w: 2.6, h: 0.3, fontSize: 12, bold: true, color: PALE, charSpacing: 2 });
    T(s, stages[i][1], { x: x + 1.0, y: y + 0.5, w: 2.75, h: 0.42, fontSize: 19, bold: true, color: WHITE });
    T(s, stages[i][2], { x: x + 0.25, y: y + 1.05, w: 3.4, h: 0.9, fontSize: 14, color: WHITE });
    T(s, stages[i][3], { x, y: base + 0.1, w: 3.85, h: 0.3, fontSize: 12, italic: true, color: GRAY, align: "center" });
  }
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 2.75, w: 3.0, h: 1.1, flipV: true, line: { color: GREEN, width: 2.5, dashType: "dash", endArrowType: "triangle" } });
  T(s, "Scope, skills and value grow with every completed project", { x: 0.6, y: 1.75, w: 4.0, h: 0.6, fontSize: 14, bold: true, color: GREEN });
  s.addNotes(`(About 2 minutes.)

The most common way improvement programs die is starting too big. Someone picks "fix packaging" as a first project, it drags on for six months, and everyone loses interest. We're not doing that.

Think of it as a staircase.

Stage 1, quick wins: 5S, visual management, organizing a workstation, simple changeover fixes like staging parts before the line stops. Small, fast, visible. They show your crew that ideas actually get acted on. Every cell starts here, and that's not a judgment; it's how you build the muscle.

Stage 2, process improvements: the downtime that keeps coming back, setup losses, waste, rework. These need more data and more problem solving.

Stage 3, larger projects: cross-functional productivity, standardization across shifts, verified cost reduction.

You earn the next step by finishing the one before it.

Ask the room: what Stage 1 quick win could your shift finish in the next few weeks? (Capture answers on the flip chart. They're candidate first projects.)`);

  // ================= 9. CHANGEOVER =================
  s = newSlide("CONTENT_LIGHT", "Focus Areas", "FOCUS AREA 1  ·  CHANGEOVER", "Prepare before the line stops and recover minutes");
  const co = [["FaRulerCombined", "Establish the current baseline"], ["FaEye", "Observe the actual changeover"], ["FaSearch", "Find waiting, motion and prep waste"], ["FaCogs", "Apply basic SMED principles"], ["FaClipboardList", "Standardize the improved method"], ["FaCheckCircle", "Verify over repeated changeovers"]];
  for (let i = 0; i < 6; i++) {
    const y = 1.75 + i * 0.76;
    await iconCircle(s, co[i][0], 0.6, y, 0.54, i === 5 ? GREEN : NAVY);
    T(s, String(i + 1), { x: 1.27, y, w: 0.35, h: 0.54, fontSize: 18, bold: true, color: STEEL, valign: "middle" });
    T(s, co[i][1], { x: 1.62, y, w: 4.3, h: 0.54, fontSize: 16, valign: "middle" });
  }
  card(s, 6.3, 1.7, 6.4, 4.75, { fill: LIGHT, name: "Illustrative changeover chart" });
  T(s, "Changeover time, minutes", { x: 6.6, y: 1.85, w: 3.2, h: 0.35, fontSize: 15, bold: true, color: NAVY });
  illus(s, 9.2, 1.88, 3.3);
  s.addChart(pres.charts.BAR, [
    { name: "Internal work (line stopped)", labels: ["Before", "After"], values: [25, 22] },
    { name: "Searching and motion", labels: ["Before", "After"], values: [7, 6] },
    { name: "Waiting", labels: ["Before", "After"], values: [8, 4] },
  ], {
    x: 6.5, y: 2.3, w: 6.0, h: 2.85, barDir: "bar", barGrouping: "stacked", chartColors: [HEX.accent1, HEX.accent3, HEX.accent6],
    showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFontSize: 11, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMaxVal: 42,
    showLegend: true, legendPos: "b", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: HEX.dk1, barGapWidthPct: 60,
  });
  T(s, [{ text: "40 → 32 min", options: { bold: true, color: GREEN, fontSize: 24 } }, { text: "   8 minutes back per changeover in this example", options: { fontSize: 14 } }], { x: 6.6, y: 5.25, w: 5.9, h: 0.5, valign: "middle" });
  T(s, "Hypothetical numbers that show the method. Your baseline comes from your own timed changeovers.", { x: 6.6, y: 5.8, w: 5.9, h: 0.5, fontSize: 11, italic: true, color: GRAY });
  s.addNotes(`(About 2.5 minutes.)

Focus area one is changeover. Changeover time is time we're not making product, and supervisors directly control how it's run.

Six steps.

Baseline: how long does it actually take today? Measure from last good product to first good product. Not what we think it takes; what it takes.

Observe: stand there with a stopwatch, with the crew's permission and involvement. You'll see things you've never noticed.

Find the waste: who's waiting, who's walking back and forth, what prep could have happened before we stopped?

Apply SMED: the big idea is that anything that can be done while the line is still running should be done while it's still running. Stage parts, pre-set tools, have the next ingredients ready.

Standardize the better method so every crew on every shift does it the same way.

Verify over repeated changeovers. One good changeover is a lucky day. Ten in a row is a result.

The chart is illustrative only, not our numbers: 40 minutes down to 32, mostly by cutting waiting. Eight minutes back every time. And remember, that's capacity. Whether it becomes dollars is Finance's call, which we'll get to.

Ask the room: which changeover on your shift drives you the craziest?`);

  // ================= 10. 5S =================
  s = newSlide("CONTENT_LIGHT", "Focus Areas", "FOCUS AREA 2  ·  5S", "A place for everything, built to food-safety standards");
  const five = [["Sort", "Remove unneeded tools, parts and supplies", "FaTrashAlt"], ["Set in Order", "Shadow boards and labeled homes for change parts", "FaThLarge"], ["Shine", "Clean and inspect, aligned to the sanitation plan", "FaBroom"], ["Standardize", "Visual labels, plant color code, one-point lessons", "FaTags"], ["Sustain", "Routine audits and supervisor walk-throughs", "FaLock"]];
  const fw = 2.3, fg = (12.1 - 5 * fw) / 4;
  for (let i = 0; i < 5; i++) {
    const x = 0.6 + i * (fw + fg);
    card(s, x, 1.7, fw, 1.78, { fill: LIGHT, name: five[i][0] + " card" });
    await iconCircle(s, five[i][2], x + 0.2, 1.85, 0.55, i === 4 ? GREEN : NAVY);
    T(s, five[i][0], { x: x + 0.85, y: 1.85, w: fw - 0.95, h: 0.55, fontSize: 17, bold: true, color: NAVY, valign: "middle" });
    T(s, five[i][1], { x: x + 0.2, y: 2.5, w: fw - 0.35, h: 0.9, fontSize: 13 });
  }
  const bx = 0.6, by = 3.68, bw = 3.95, bh = 3.07;
  card(s, bx, by, bw, bh, { fill: LIGHT, name: "Before concept" });
  T(s, "BEFORE  ·  concept", { x: bx + 0.2, y: by + 0.12, w: 3, h: 0.3, fontSize: 12, bold: true, color: GRAY, charSpacing: 1 });
  const junk = [[0.3, 0.6, 1.1, 0.25, 20], [1.3, 0.75, 0.8, 0.3, -15], [2.3, 0.6, 1.2, 0.22, 35], [0.5, 1.3, 0.6, 0.6, 10], [1.4, 1.5, 1.3, 0.2, -30], [2.8, 1.3, 0.6, 0.45, 5], [0.4, 2.15, 1.4, 0.24, -8], [2.1, 2.05, 0.9, 0.35, 25], [3.1, 2.15, 0.45, 0.45, 0]];
  junk.forEach(([x, y, w, h, r], k) => s.addShape(k % 3 === 0 ? pres.shapes.OVAL : pres.shapes.ROUNDED_RECTANGLE, { x: bx + x, y: by + y - 0.1, w, h, rotate: (r + 360) % 360, rectRadius: 0.05, fill: { color: k % 2 ? GRAY : STEEL }, line: { type: "none" } }));
  T(s, "Time lost searching", { x: bx + 0.2, y: by + 2.7, w: 3.5, h: 0.3, fontSize: 12, italic: true, color: GRAY });
  const ax = 4.75;
  card(s, ax, by, bw, bh, { fill: LIGHT, name: "After concept" });
  T(s, "AFTER  ·  shadow board concept", { x: ax + 0.2, y: by + 0.12, w: 3.5, h: 0.3, fontSize: 12, bold: true, color: GREEN, charSpacing: 1 });
  ["Scraper", "Wrench", "Change part", "Brush", "Gauge", "Spare blade"].forEach((t, k) => {
    const cx = ax + 0.25 + (k % 3) * 1.18, cy = by + 0.55 + Math.floor(k / 3) * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w: 1.05, h: 0.6, rectRadius: 0.06, fill: { color: WHITE }, line: { color: NAVY, width: 1.25, dashType: "dash" } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx + 0.2, y: cy + 0.22, w: 0.65, h: 0.16, rectRadius: 0.05, fill: { color: k === 3 ? GREEN : MID }, line: { type: "none" } });
    T(s, t, { x: cx - 0.05, y: cy + 0.63, w: 1.15, h: 0.3, fontSize: 10, align: "center" });
  });
  T(s, "Missing items visible at a glance", { x: ax + 0.2, y: by + 2.7, w: 3.5, h: 0.3, fontSize: 12, italic: true, color: GREEN });
  card(s, 8.9, by, 3.8, bh, { fill: NAVY, name: "Food safety callout" });
  await iconCircle(s, "FaShieldAlt", 9.1, by + 0.2, 0.55, GREEN);
  T(s, "Food safety first", { x: 9.75, y: by + 0.2, w: 2.8, h: 0.55, fontSize: 18, bold: true, color: WHITE, valign: "middle" });
  T(s, bullets(["Follow the sanitation program and approved chemicals", "Keep food-contact and non-contact tools separate", "Use the plant's color code; never invent one", "Approved, cleanable materials only", "Quality reviews before anything goes on the floor"]), { x: 9.1, y: by + 0.88, w: 3.45, h: 2.1, fontSize: 12, color: WHITE, paraSpaceAfter: 3 });
  s.addNotes(`(About 2 minutes.)

Focus area two is 5S, made real for a frozen-food plant.

Sort: if it's not needed at that station, it goes. The extra wrench, the broken scraper, the parts for a product we haven't run in a year.

Set in order: everything gets a home. Shadow boards, labeled locations, so anyone on any shift can find it in seconds.

Shine: cleaning is inspecting. When you clean equipment you find the loose bolt, the leak, the worn belt. In our world, Shine lines up with the sanitation plan.

Standardize: make the right way obvious. Labels, the plant's color code, one-point lessons.

Sustain: the hard one. Routine audits and supervisor walk-throughs.

The two concept pictures make the point: on the left you lose time searching; on the right a missing tool is visible at a glance.

The navy box is non-negotiable. In food manufacturing, 5S has to support food safety, never work against it. Use the sanitation program and approved chemicals, keep food-contact tools separate, use the plant's existing color code, use approved materials, and Quality reviews it before anything goes on the floor.

Ask the room: where on your line do people lose the most time searching for something?`);

  // ================= 11. ENGAGEMENT =================
  s = newSlide("CONTENT_LIGHT", "Focus Areas", "FOCUS AREA 3  ·  EMPLOYEE ENGAGEMENT", "People support what they help improve");
  const cx0 = 3.15, cy0 = 4.1, R = 1.65;
  s.addShape(pres.shapes.OVAL, { x: cx0 - R, y: cy0 - R, w: 2 * R, h: 2 * R, fill: { type: "none" }, line: { color: PALE, width: 4 }, objectName: "Idea loop ring" });
  card(s, cx0 - 1.0, cy0 - 0.55, 2.0, 1.1, { fill: NAVY, name: "Loop center" });
  T(s, [{ text: "The idea loop", options: { bold: true, color: GREEN, fontSize: 13, breakLine: true } }, { text: "Every idea gets an answer", options: { bold: true, color: WHITE, fontSize: 14 } }], { x: cx0 - 0.95, y: cy0 - 0.55, w: 1.9, h: 1.1, align: "center", valign: "middle" });
  const loop = [["FaComments", "Ask", "“What slowed you down?”"], ["FaReply", "Respond", "Yes, or not now and why"], ["FaUserCheck", "Involve", "Operator owns an action"], ["FaShareAlt", "Share", "Across A, B and C shifts"], ["FaAward", "Recognize", "Credit by name"]];
  for (let i = 0; i < 5; i++) {
    const a = (-90 + i * 72) * Math.PI / 180, nx = cx0 + R * Math.cos(a), ny = cy0 + R * Math.sin(a);
    await iconCircle(s, loop[i][0], nx - 0.4, ny - 0.4, 0.8, i === 4 ? GREEN : i % 2 ? MID : NAVY);
    const right = Math.cos(a) > 0.2, left = Math.cos(a) < -0.2;
    const lx = i === 0 ? nx + 0.5 : right ? nx + 0.48 : left ? nx - 2.08 : nx - 0.8;
    T(s, [{ text: loop[i][1], options: { bold: true, color: NAVY, fontSize: 15, breakLine: true } }, { text: loop[i][2], options: { fontSize: 11, color: STEEL } }], { x: left ? 0.3 : lx, y: ny - 0.32, w: left ? nx - 0.78 : 1.6, h: 0.64, valign: "middle", align: left ? "right" : "left" });
  }
  card(s, 6.75, 1.7, 5.95, 3.55, { fill: LIGHT, name: "Supervisor engagement standard" });
  T(s, "Supervisor engagement standard", { x: 7.05, y: 1.85, w: 4.0, h: 0.4, fontSize: 17, bold: true, color: NAVY });
  pill(s, 11.0, 1.9, 1.5, "PROPOSED", STEEL, { fontSize: 9 });
  const std = [["FaComments", "Ask the crew every week what slowed them down"], ["FaReply", "Answer every idea within one week"], ["FaUserCheck", "Give every project at least one operator-owned action"], ["FaAward", "Credit contributors by name, in front of peers"]];
  for (let i = 0; i < 4; i++) {
    const y = 2.4 + i * 0.7;
    s.addImage({ data: await icon(std[i][0], HEX.accent2), x: 7.05, y: y + 0.1, w: 0.36, h: 0.36, altText: std[i][1] });
    T(s, std[i][1], { x: 7.6, y, w: 4.9, h: 0.56, fontSize: 15, valign: "middle" });
  }
  card(s, 6.75, 5.45, 5.95, 1.2, { fill: NAVY, name: "Engagement measures" });
  T(s, "How we'll know it's working", { x: 7.05, y: 5.55, w: 5.4, h: 0.35, fontSize: 14, bold: true, color: GREEN });
  ["Ideas raised", "% answered ≤ 1 week", "Operator actions"].forEach((t, k) => pill(s, 7.05 + k * 1.85, 6.05, 1.75, t, MID, { h: 0.38, fontSize: 11 }));
  s.addNotes(`(About 3 minutes. This is the slide that makes or breaks the program.)

The operators running the line know where the problems are. Usually they've known for a long time. The question is whether we ask, and whether we act.

On the left is the idea loop. Ask: "What slowed you down this shift?" Respond: every idea gets an answer, either yes, or "not now, and here's why." Silence kills engagement faster than a no does. Involve: the person who raised it owns an action, if they want to. Share it across all three shifts. And recognize: credit them by name.

On the right is a proposed engagement standard for supervisors. I'm calling it proposed because I want your input before we lock it in. Ask your crew every week what slowed them down. Answer every idea within a week. Every project has at least one operator-owned action. And credit contributors by name, in front of their peers.

At the bottom is how we'll know it's working: ideas raised, the percentage answered within a week, and how many operators own an action. Notice we're measuring the response, not just the ideas. That's on us, not on the operators.

The headline says it all: people support what they help improve. If we design a change in an office and push it onto the floor, it fades. If the crew builds it, they'll defend it.

Ask the room: is a one-week answer realistic on your shift? What would get in the way?`);

  // ================= 12. PROJECT SELECTION =================
  pres.addSection({ title: "Running a Project" });
  s = newSlide("CONTENT_LIGHT", "Running a Project", "PROJECT SELECTION", "Pick a first project you can finish, measure and influence");
  const filt = [["FaShieldAlt", "Safe", "No added risk to people or food safety"], ["FaCompress", "Small enough to complete", "Weeks, not months"], ["FaRulerCombined", "Measurable", "A baseline and a number to move"], ["FaCrosshairs", "Within our influence", "The team can change it directly"], ["FaGem", "Valuable to operations", "Output, quality or cost"], ["FaBoxes", "Resourced", "People, parts and time available"]];
  T(s, "ALL IDEAS", { x: 0.6, y: 1.65, w: 7.2, h: 0.3, fontSize: 12, bold: true, color: GRAY, align: "center", charSpacing: 2 });
  for (let i = 0; i < 6; i++) {
    const w = 7.2 - i * 0.5, x = 0.6 + (7.2 - w) / 2, y = 2.0 + i * 0.64;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.54, rectRadius: 0.06, fill: { color: [PALE, PALE, STEEL, STEEL, MID, MID][i] }, line: { type: "none" }, objectName: "Funnel filter " + (i + 1) });
    const dark = i >= 2;
    s.addImage({ data: await icon(filt[i][0], dark ? "FFFFFF" : HEX.dk2), x: x + 0.18, y: y + 0.13, w: 0.28, h: 0.28, altText: filt[i][1] });
    T(s, [{ text: filt[i][1] + "  ", options: { bold: true, fontSize: 14 } }, { text: filt[i][2], options: { fontSize: 12 } }], { x: x + 0.55, y, w: w - 0.65, h: 0.54, valign: "middle", color: dark ? WHITE : NAVY });
  }
  card(s, 2.6, 5.9, 3.2, 0.52, { fill: GREEN, name: "Funnel output" });
  T(s, "Our first project", { x: 2.6, y: 5.9, w: 3.2, h: 0.52, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
  card(s, 8.15, 1.7, 4.55, 2.25, { fill: LIGHT, name: "Good first projects" });
  s.addImage({ data: await icon("FaCheckCircle", HEX.accent2), x: 8.35, y: 1.88, w: 0.32, h: 0.32, altText: "Good" });
  T(s, "Good first projects", { x: 8.8, y: 1.86, w: 3.7, h: 0.36, fontSize: 16, bold: true, color: GREEN });
  T(s, bullets(["Stage change parts for one changeover", "Shadow board for one line's tools", "Stop one recurring jam on one machine"]), { x: 8.35, y: 2.35, w: 4.2, h: 1.45, fontSize: 14, paraSpaceAfter: 4 });
  card(s, 8.15, 4.15, 4.55, 2.27, { fill: LIGHT, name: "Too big for now" });
  s.addImage({ data: await icon("FaFlag", HEX.accent5), x: 8.35, y: 4.32, w: 0.3, h: 0.3, altText: "Too big" });
  T(s, "Too big for a first project", { x: 8.8, y: 4.3, w: 3.7, h: 0.36, fontSize: 16, bold: true, color: AMBER });
  T(s, bullets(["Redesign the whole packaging area", "Replace major equipment", "Anything needing capital approval"]), { x: 8.35, y: 4.8, w: 4.2, h: 1.45, fontSize: 14, paraSpaceAfter: 4 });
  s.addNotes(`(About 2.5 minutes.)

How do you pick a good first project? Run it through the funnel. Lots of ideas go in the top; one good project comes out the bottom.

Is it safe? Nothing we do adds risk to people or food safety. Non-negotiable.

Small enough to complete? Weeks, not months, for the first one.

Measurable? If we can't measure before and after, we can't prove it worked.

Within our influence? If the fix depends on a supplier, corporate or a capital request, it's not a first project.

Valuable? It has to matter to output, quality or cost.

Resourced? People, parts and time available.

On the right, good first projects are narrow: one changeover, one line's tools, one recurring jam. Too big for now: redesigning packaging, replacing equipment, anything needing capital. Those aren't bad ideas; they're Stage 3 ideas.

There's a scoring matrix in the appendix if you want it more formal.

Ask the room: take the idea you had in your head at the start and run it through the funnel. Where does it get stuck? (Work one or two examples live.)`);

  // ================= 13. JOURNEY =================
  s = newSlide("CONTENT_LIGHT", "Running a Project", "THE SUPERVISOR'S PROJECT JOURNEY", "Seven steps from problem to presented result");
  const steps = [["Identify a problem", "What's hurting us?", "FaSearch"], ["Define the baseline", "How bad is it today?", "FaRulerCombined"], ["Set a measurable target", "What does good look like?", "FaBullseye"], ["Engage the team", "Who needs to be involved?", "FaUsers"], ["Implement the improvement", "What are we changing?", "FaCogs"], ["Verify the results", "Did it really work?", "FaChartLine"], ["Present and sustain", "How do we keep it?", "FaMicrophone"]];
  const pdca = [["PLAN", 0, 4, NAVY], ["DO", 4, 1, MID], ["CHECK", 5, 1, STEEL], ["ACT", 6, 1, GREEN]];
  const sw = 12.1 / 7;
  pdca.forEach(([lab, st, n, col]) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6 + st * sw + 0.06, y: 1.7, w: n * sw - 0.12, h: 0.44, rectRadius: 0.08, fill: { color: col }, line: { type: "none" }, objectName: "PDCA " + lab });
    T(s, lab, { x: 0.6 + st * sw, y: 1.7, w: n * sw, h: 0.44, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle", charSpacing: 3 });
  });
  s.addShape(pres.shapes.LINE, { x: 0.6 + sw / 2, y: 3.0, w: 12.1 - sw, h: 0, line: { color: PALE, width: 4 } });
  for (let i = 0; i < 7; i++) {
    const cx = 0.6 + i * sw + sw / 2, col = i < 4 ? NAVY : [MID, STEEL, GREEN][i - 4];
    await iconCircle(s, steps[i][2], cx - 0.42, 2.58, 0.84, col);
    T(s, String(i + 1), { x: cx - 0.3, y: 3.52, w: 0.6, h: 0.45, fontSize: 22, bold: true, color: col, align: "center" });
    T(s, steps[i][0], { x: cx - sw / 2 + 0.06, y: 3.98, w: sw - 0.12, h: 0.7, fontSize: 15, bold: true, color: NAVY, align: "center" });
    T(s, steps[i][1], { x: cx - sw / 2 + 0.06, y: 4.68, w: sw - 0.12, h: 0.5, fontSize: 12, italic: true, color: GRAY, align: "center" });
  }
  s.addShape(pres.shapes.LINE, { x: 0.6, y: 5.3, w: 12.1, h: 0, line: { color: PALE, width: 1 } });
  T(s, "CHECKPOINTS", { x: 0.6, y: 5.42, w: 1.6, h: 0.3, fontSize: 11, bold: true, color: GRAY, charSpacing: 2 });
  const gates = [[1, "CI", GREEN, "FaSyncAlt"], [2, "CI", GREEN, "FaSyncAlt"], [5, "CI", GREEN, "FaSyncAlt"], [1, "Finance", NAVY, "FaDollarSign"], [5, "Finance", NAVY, "FaDollarSign"], [4, "Quality", STEEL, "FaClipboardCheck"], [6, "WOr", MID, "FaUserTie"]];
  const byStep = {};
  gates.forEach((g) => (byStep[g[0]] = (byStep[g[0]] || []).concat([g])));
  for (const k of Object.keys(byStep)) {
    const cx = 0.6 + Number(k) * sw + sw / 2, list = byStep[k];
    for (let j = 0; j < list.length; j++) {
      const y = 5.45 + j * 0.42;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx - 0.72, y, w: 1.44, h: 0.34, rectRadius: 0.17, fill: { color: list[j][2] }, line: { type: "none" } });
      s.addImage({ data: await icon(list[j][3], "FFFFFF"), x: cx - 0.62, y: y + 0.07, w: 0.2, h: 0.2, altText: list[j][1] + " checkpoint" });
      T(s, list[j][1], { x: cx - 0.36, y, w: 1.0, h: 0.34, fontSize: 11, bold: true, color: WHITE, valign: "middle" });
    }
  }
  s.addNotes(`(About 2 minutes.)

Every project follows the same seven steps, and if they look familiar, it's Plan, Do, Check, Act.

Plan is the first four: identify the problem, define the baseline, set a measurable target, engage the team. Most projects that fail, fail in Plan. They skip the baseline, or the target is vague. A good target sounds like: "cut changeover on this line from our measured baseline to a specific number, held over ten changeovers."

Do is implementing. Check is verifying. Act is presenting and sustaining: update the standard work, train the other shifts, bring it to WOr.

The italic questions are the questions I'll ask in coaching. If you can answer them, you're in good shape.

The pills along the bottom are the checkpoints. CI checks in at the baseline, the target and verification, because that's where projects wobble. Finance checks in twice: at the baseline, to agree how the benefit will be valued, and at verification, to validate it. Quality reviews before anything is implemented. And WOr is where you present.

Ask the room: which of these seven steps feels hardest to you right now?`);

  // ================= 14. MEASURING =================
  s = newSlide("CONTENT_LIGHT", "Running a Project", "MEASURING RESULTS", "Teams track operational gains; Finance confirms savings");
  T(s, "OPERATIONAL GAINS  ·  tracked by the team on the floor", { x: 0.6, y: 1.65, w: 7.4, h: 0.32, fontSize: 12, bold: true, color: MID, charSpacing: 1 });
  const kpis = [["FaStopwatch", "Changeover minutes", "Last good to first good product"], ["FaTachometerAlt", "Downtime minutes", "Unplanned stops, by cause"], ["FaRecycle", "Scrap and rework", "Pounds, units or cases lost"], ["FaUsers", "Employee participation", "Ideas raised and answered"], ["FaThLarge", "5S audit results", "Score against the area standard"], ["FaTasks", "Action completion", "Actions closed on time"]];
  for (let i = 0; i < 6; i++) {
    const x = 0.6 + (i % 2) * 3.75, y = 2.05 + Math.floor(i / 2) * 1.22;
    card(s, x, y, 3.55, 1.06, { fill: LIGHT, name: kpis[i][1] });
    await iconCircle(s, kpis[i][0], x + 0.15, y + 0.24, 0.58, NAVY);
    T(s, kpis[i][1], { x: x + 0.88, y: y + 0.18, w: 2.6, h: 0.35, fontSize: 15, bold: true, color: NAVY });
    T(s, kpis[i][2], { x: x + 0.88, y: y + 0.55, w: 2.6, h: 0.35, fontSize: 12, color: STEEL });
  }
  arrowRight(s, 8.0, 3.75, 0.55, GREEN);
  card(s, 8.7, 1.65, 4.0, 3.96, { fill: NAVY, name: "Financial benefits panel" });
  await iconCircle(s, "FaHandHoldingUsd", 8.95, 1.9, 0.7, GREEN);
  T(s, "Financial benefits", { x: 8.95, y: 2.75, w: 3.5, h: 0.45, fontSize: 21, bold: true, color: WHITE });
  T(s, "VALIDATED BY FINANCE", { x: 8.95, y: 3.2, w: 3.5, h: 0.32, fontSize: 12, bold: true, color: GREEN, charSpacing: 1 });
  T(s, "A better process is a real gain. It becomes reported savings only when Finance confirms that actual cost went down.", { x: 8.95, y: 3.65, w: 3.5, h: 1.7, fontSize: 14, color: PALE });
  card(s, 0.6, 5.85, 12.1, 0.75, { fill: LIGHT, name: "Distinction" });
  T(s, [{ text: "Operational gain ", options: { bold: true, color: NAVY } }, { text: "= the process got better.        " }, { text: "Realized savings ", options: { bold: true, color: GREEN } }, { text: "= actual cost went down, and Finance agrees." }], { x: 0.85, y: 5.85, w: 11.7, h: 0.75, fontSize: 16, valign: "middle", align: "center" });
  s.addNotes(`(About 2 minutes.)

How will we measure results? Honestly, and in two different ways.

On the left are the operational measures your team tracks on the floor: changeover minutes, downtime minutes by cause, scrap and rework, employee participation (ideas raised and answered), 5S audit results, and action completion. Pick the one or two that fit your project. You don't need all six. Use data we already collect wherever we can.

On the right is the financial side, and the distinction at the bottom matters. An operational gain means the process got better. Realized savings means actual cost went down, and Finance agrees.

They are not the same. If we cut changeover by eight minutes but run the same hours with the same crew, the process improved and we created capacity. We didn't save money yet.

Why does this matter? Because the fastest way to lose credibility with leadership is to add up a pile of "savings" that never shows up in the cost results. We won't do that. We'll report operational gains proudly, and we'll report savings only once Randy has validated them.

Ask the room: which of these measures could you start tracking on your line this week, with what you already have?`);

  // ================= 15. FINANCE VALIDATION =================
  s = newSlide("CONTENT_LIGHT", "Running a Project", "FINANCE VALIDATION  ·  RANDY", "Savings are reported only after Finance validates them");
  const flow = [["1", "Agree method", "Baseline, data source and valuation, at the charter"], ["2", "Implement", "Make the change and track the measure"], ["3", "Sustain", "Hold the gain over an agreed period"], ["4", "Validate", "Finance checks against actual cost results"], ["5", "Report", "By benefit type, net of one-time costs"]];
  const flw = 2.32, flg = (12.1 - 5 * flw) / 4;
  for (let i = 0; i < 5; i++) {
    const x = 0.6 + i * (flw + flg), col = i === 0 || i === 3 ? NAVY : i === 4 ? GREEN : STEEL;
    s.addShape(pres.shapes.CHEVRON, { x, y: 1.68, w: flw + (i < 4 ? 0.12 : 0), h: 0.62, fill: { color: col }, line: { type: "none" }, objectName: "Flow " + flow[i][1] });
    T(s, flow[i][0] + "  " + flow[i][1], { x: x + 0.3, y: 1.68, w: flw - 0.4, h: 0.62, fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, flow[i][2], { x: x + 0.15, y: 2.38, w: flw - 0.2, h: 0.6, fontSize: 12, color: INK, align: "center" });
    if (i === 0 || i === 3) pill(s, x + 0.5, 3.0, flw - 0.9, "FINANCE", NAVY, { fontSize: 9, h: 0.26 });
  }
  const cats = [["FaDollarSign", "Hard savings", "Actual spend falls versus the baseline: overtime, material, scrap.", "Counts as savings", GREEN], ["FaShieldAlt", "Cost avoidance", "A likely future cost that we prevented.", "Tracked separately", MID], ["FaTachometerAlt", "Capacity", "Freed time or output. Worth $ only once used.", "Reported in hours", STEEL], ["FaUsers", "Nonfinancial", "Safety, food safety, ergonomics, morale.", "Reported as impact", GRAY]];
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + i * 3.08;
    card(s, x, 3.45, 2.86, 1.9, { fill: LIGHT, name: cats[i][1] });
    await iconCircle(s, cats[i][0], x + 0.18, 3.6, 0.5, cats[i][4]);
    T(s, cats[i][1], { x: x + 0.8, y: 3.6, w: 1.95, h: 0.5, fontSize: 16, bold: true, color: NAVY, valign: "middle" });
    T(s, cats[i][2], { x: x + 0.18, y: 4.17, w: 2.55, h: 0.7, fontSize: 12 });
    pill(s, x + 0.18, 4.92, 2.5, cats[i][3], cats[i][4], { fontSize: 10, h: 0.28 });
  }
  card(s, 0.6, 5.55, 12.1, 1.12, { fill: NAVY, name: "Valuation rules" });
  T(s, "VALUATION RULES", { x: 0.85, y: 5.62, w: 3, h: 0.3, fontSize: 11, bold: true, color: GREEN, charSpacing: 2 });
  const rules = [["FaStopwatch", "Labor counts only when paid hours fall"], ["FaRecycle", "Scrap valued at cost, not selling price"], ["FaBalanceScale", "Adjusted for volume and mix"], ["FaLayerGroup", "Net of one-time costs; no double counting"]];
  for (let i = 0; i < 4; i++) {
    const x = 0.85 + i * 2.98;
    s.addImage({ data: await icon(rules[i][0], HEX.accent6), x, y: 6.04, w: 0.32, h: 0.32, altText: rules[i][1] });
    T(s, rules[i][1], { x: x + 0.42, y: 5.95, w: 2.45, h: 0.55, fontSize: 12, color: WHITE, valign: "middle" });
  }
  s.addNotes(`(About 3 minutes. Consider inviting Randy to cover part of this slide.)

The rule is simple: savings are reported only after Finance validates them.

The process across the top has five steps, and Finance is involved at two of them.

Step one happens at the START, not the end. When you write your charter, you and Randy agree the baseline, the data source, and how the benefit will be valued. That avoids arguments later about whose number is right.

Step two: implement and track. Step three: sustain the gain over an agreed verification period, because one good week doesn't count. Step four: Randy validates against actual cost results. Step five: we report by benefit type, net of any one-time costs, like the cost of a new cart or a shadow board.

The four boxes are the benefit types. Hard savings: actual spend falls compared with the baseline, such as overtime, material or scrap. Only these count as savings. Cost avoidance: we prevented a likely future cost, like an avoided repair or not needing added overtime for extra volume. Real, but tracked separately. Capacity: we freed up time or output, reported in hours. It's worth dollars only once it's used, either by removing paid hours or by producing additional product we actually sell. Nonfinancial: safety, food safety, ergonomics, morale.

And the navy bar has the valuation rules Finance will apply. Labor counts only when paid hours actually fall; saved minutes alone are capacity. Scrap is valued at cost, not selling price. Results are adjusted for volume and product mix, so a slow month doesn't look like a saving. And everything is net of one-time costs, with no double counting between projects.

Ask the room: can anyone think of a past improvement that was "savings" on paper but never showed up in the cost results? That's exactly what this process prevents.`);

  // ================= 16. PRESENTATION =================
  s = newSlide("CONTENT_LIGHT", "Running a Project", "LEADERSHIP PRESENTATION  ·  WOr MEETING", "Your project. Your results. Your presentation.");
  T(s, "Every presentation answers six questions", { x: 0.6, y: 1.7, w: 5.6, h: 0.4, fontSize: 16, bold: true, color: NAVY });
  const qs = ["What problem did we identify?", "What was the baseline?", "What did the team change?", "What were the results?", "What financial value was validated?", "How will we sustain the improvement?"];
  for (let i = 0; i < 6; i++) {
    const y = 2.25 + i * 0.62;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.46, h: 0.46, fill: { color: i === 4 ? GREEN : NAVY }, line: { type: "none" } });
    T(s, String(i + 1), { x: 0.6, y, w: 0.46, h: 0.46, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, qs[i], { x: 1.25, y, w: 4.9, h: 0.46, fontSize: 16, valign: "middle" });
  }
  ["5 minutes", "1 page", "Team named"].forEach((t, k) => pill(s, 0.6 + k * 1.9, 6.1, 1.75, t, k === 2 ? GREEN : MID, { h: 0.4, fontSize: 12 }));
  const mx = 6.55, my = 1.7, mw = 6.15, mh = 5.0;
  card(s, mx, my, mw, mh, { fill: WHITE, line: PALE, lineW: 1, shadow: true, name: "Mock one-page project summary" });
  s.addShape(pres.shapes.RECTANGLE, { x: mx, y: my, w: mw, h: 0.55, fill: { color: NAVY }, line: { type: "none" } });
  T(s, "PROJECT ONE-PAGER  ·  SAMPLE TEMPLATE", { x: mx + 0.2, y: my, w: 4.2, h: 0.55, fontSize: 12, bold: true, color: WHITE, valign: "middle", charSpacing: 1 });
  T(s, "Cell: [A Shift]   Owner: [Supervisor]", { x: mx + 3.4, y: my, w: 2.6, h: 0.55, fontSize: 10, color: PALE, valign: "middle", align: "right" });
  const boxes = [["1 · Problem", "[One sentence: what, where, impact]"], ["2 · Baseline", "[Measured value, period, data source]"], ["3 · What we changed", "[Key changes and who made them]"], ["4 · Results", ""], ["5 · Validated value", "[Benefit type and Finance-reviewed amount]"], ["6 · Sustain plan", "[Standard work, audit, owner]"]];
  const bw2 = (mw - 0.6) / 2, bh2 = 1.38;
  for (let i = 0; i < 6; i++) {
    const x = mx + 0.2 + (i % 2) * (bw2 + 0.2), y = my + 0.7 + Math.floor(i / 2) * (bh2 + 0.1);
    card(s, x, y, bw2, bh2, { fill: LIGHT });
    T(s, boxes[i][0], { x: x + 0.12, y: y + 0.08, w: bw2 - 0.24, h: 0.3, fontSize: 12, bold: true, color: i === 4 ? GREEN : NAVY });
    if (i === 3) {
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: y + 0.5, w: 1.85, h: 0.24, fill: { color: GRAY }, line: { type: "none" } });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: y + 0.84, w: 1.35, h: 0.24, fill: { color: GREEN }, line: { type: "none" } });
      T(s, "Before", { x: x + 2.08, y: y + 0.47, w: 0.6, h: 0.3, fontSize: 10, color: GRAY });
      T(s, "After", { x: x + 1.58, y: y + 0.81, w: 0.6, h: 0.3, fontSize: 10, color: GREEN });
      T(s, "Chart of your measured results", { x: x + 0.12, y: y + 1.1, w: bw2 - 0.24, h: 0.25, fontSize: 10, italic: true, color: GRAY });
    } else T(s, boxes[i][1], { x: x + 0.12, y: y + 0.42, w: bw2 - 0.24, h: 0.85, fontSize: 12, italic: true, color: STEEL });
  }
  s.addNotes(`(About 2 minutes.)

Here's where it comes together: you present your own project to leadership at the existing WOr meeting. Not CI presenting your work. Your project, your team's results, your presentation. It's how leadership sees what you can do, and it's one of the best development opportunities in the whole program.

Every presentation answers the same six questions: the problem, the baseline, what the team changed, the results, the financial value validated, and how we'll sustain it.

Three rules: five minutes, one page, team named. No thirty-slide decks. And the people who did the work are named on the page.

On the right is the one-page format. I'll send everyone an editable version, and I'm happy to do a dry run with anyone before they present.

Box five is where Randy's review shows up. If the value is capacity or nonfinancial, say so. That's completely fine, and it's honest.

You'll present at the next WOr meeting after your results are verified. I'll coordinate the slot with plant leadership so you're not scrambling.

Ask the room: how many of you have presented to the leadership team before? For those who haven't, we'll practice together. Nobody gets thrown in cold.`);

  // ================= 17. RECOGNITION =================
  s = newSlide("CONTENT_LIGHT", "Running a Project", "RECOGNITION", "Your team's results will be seen, shared and celebrated");
  const rec = [
    ["FaUserCheck", "On your shift", "Supervisor credits contributors by name at the shift huddle", "Same week", STEEL, 2.0],
    ["FaShareAlt", "Across shifts", "Wins posted and shared at A/B/C hand-offs", "As results come in", MID, 2.5],
    ["FaUserTie", "At WOr", "Your team presents its results to plant leadership", "After verification", NAVY, 3.0],
    ["FaTrophy", "Day 90", "Plant leadership recognizes successful teams (format to be confirmed)", "Week of Feb 15", GREEN, 3.5],
  ];
  const rbase = 5.75;
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + i * 3.05, h = rec[i][5], y = rbase - h;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h, fill: { color: rec[i][4] }, line: { type: "none" }, objectName: "Recognition level " + (i + 1) });
    await iconCircle(s, rec[i][0], x + 0.22, y + 0.22, 0.6, WHITE, i === 3 ? HEX.accent2 : i === 2 ? HEX.dk2 : i === 1 ? HEX.accent1 : HEX.accent3);
    T(s, rec[i][1], { x: x + 0.95, y: y + 0.22, w: 1.85, h: 0.6, fontSize: 18, bold: true, color: WHITE, valign: "middle" });
    T(s, rec[i][2], { x: x + 0.22, y: y + 0.95, w: 2.5, h: 0.62, fontSize: 13, color: WHITE });
    T(s, rec[i][3], { x, y: rbase + 0.08, w: 2.9, h: 0.3, fontSize: 12, italic: true, color: STEEL, align: "center" });
  }
  const pr = [["FaBullseye", "Specific"], ["FaRegClock", "Timely"], ["FaUsers", "Team-based"], ["FaUserCheck", "Operators credited"]];
  T(s, "Recognition principles", { x: 0.6, y: 6.2, w: 2.6, h: 0.42, fontSize: 14, bold: true, color: NAVY, valign: "middle" });
  for (let i = 0; i < 4; i++) {
    const x = 3.1 + i * 1.95;
    s.addImage({ data: await icon(pr[i][0], HEX.accent2), x, y: 6.27, w: 0.28, h: 0.28, altText: pr[i][1] });
    T(s, pr[i][1], { x: x + 0.36, y: 6.2, w: 1.55, h: 0.42, fontSize: 14, valign: "middle" });
  }
  s.addNotes(`(About 2 minutes.)

I promised you'd know how your work will be recognized, so here it is. Four levels, and they build on each other.

On your shift: the fastest, and honestly the most important. Supervisors credit the people who contributed, by name, at the shift huddle, the same week. It costs nothing and it matters more than anything else on this slide.

Across shifts: as results come in, wins get posted and shared at the A, B and C hand-offs, so the other shifts see them and can copy them.

At WOr: your team presents its own results to plant leadership. That's real visibility for you and for your crew.

Day 90: plant leadership recognizes the teams that delivered. The exact format is for leadership to confirm, and I'll share it as soon as it's set.

Along the bottom are four recognition principles. Specific: say what they did, not just "good job." Timely: the same week, not three months later. Team-based: we recognize teams, not just the supervisor. And operators credited: the people on the line get named.

Ask the room: what kind of recognition actually means something to your crews? I want to make sure Day 90 lands with them, not just with us.`);

  // ================= 18. 90 DAYS =================
  pres.addSection({ title: "Launch Plan" });
  s = newSlide("CONTENT_LIGHT", "Launch Plan", "OUR FIRST 90 DAYS  ·  NOV 17, 2026 – FEB 15, 2027", "Three projects chartered, delivered and recognized in 90 days");
  const weeks = ["11/16", "11/23", "11/30", "12/7", "12/14", "12/21", "12/28", "1/4", "1/11", "1/18", "1/25", "2/1", "2/8", "2/15"];
  const gx = 3.25, gw = (12.7 - gx) / 14, gy = 2.4, rh = 0.36;
  const phase = [["DAYS 1–30  LAUNCH", 0, 5, STEEL], ["DAYS 31–60  BUILD", 5, 4, MID], ["DAYS 61–90  SUSTAIN & GROW", 9, 5, NAVY]];
  phase.forEach(([t, st, n, col]) => {
    s.addShape(pres.shapes.RECTANGLE, { x: gx + st * gw + 0.02, y: 1.65, w: n * gw - 0.04, h: 0.36, fill: { color: col }, line: { type: "none" }, objectName: t });
    T(s, t, { x: gx + st * gw, y: 1.65, w: n * gw, h: 0.36, fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle", charSpacing: 1 });
  });
  T(s, "WEEK OF", { x: 0.6, y: 2.05, w: 2.5, h: 0.3, fontSize: 10, bold: true, color: GRAY, charSpacing: 2, valign: "middle" });
  [1, 5, 6].forEach((k) => s.addShape(pres.shapes.RECTANGLE, { x: gx + k * gw, y: 2.05, w: gw, h: gy + 8 * rh - 2.05, fill: { color: LIGHT }, line: { type: "none" }, objectName: "Holiday week" }));
  weeks.forEach((w, k) => T(s, w, { x: gx + k * gw, y: 2.05, w: gw, h: 0.3, fontSize: 10, color: [1, 5, 6].includes(k) ? AMBER : INK, bold: [1, 5, 6].includes(k), align: "center", valign: "middle" }));
  const ws = [
    ["Confirm cells and sign charters", 0, 1, NAVY],
    ["Measure baselines with Finance", 1, 3, MID],
    ["Implement one quick win per cell", 2, 5, STEEL],
    ["Verify results and standardize", 5, 9, MID],
    ["Present at WOr (after verification)", 8, 10, GREEN],
    ["Select and charter Stage 2 projects", 10, 12, NAVY],
  ];
  for (let r = 0; r < 8; r++) {
    const y = gy + r * rh;
    if (r % 2 === 0) s.addShape(pres.shapes.LINE, { x: 0.6, y: y + rh, w: 12.1, h: 0, line: { color: PALE, width: 0.5 } });
  }
  ws.forEach(([t, a, b, col], r) => {
    const y = gy + r * rh;
    T(s, t, { x: 0.6, y, w: 2.6, h: rh, fontSize: 11, color: INK, valign: "middle" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: gx + a * gw + 0.04, y: y + 0.07, w: (b - a + 1) * gw - 0.08, h: rh - 0.14, rectRadius: 0.1, fill: { color: col }, line: { type: "none" }, objectName: "Bar: " + t });
  });
  const y7 = gy + 6 * rh;
  T(s, "CI coaching, weekly at the line", { x: 0.6, y: y7, w: 2.6, h: rh, fontSize: 11, bold: true, color: GREEN, valign: "middle" });
  for (let k = 0; k < 14; k++) {
    const light = [5, 6].includes(k), d = light ? 0.12 : 0.18;
    s.addShape(pres.shapes.OVAL, { x: gx + k * gw + (gw - d) / 2, y: y7 + (rh - d) / 2, w: d, h: d, fill: { color: light ? PALE : GREEN }, line: { type: "none" } });
  }
  const y8 = gy + 7 * rh;
  T(s, "Leadership reviews", { x: 0.6, y: y8, w: 2.6, h: rh, fontSize: 11, bold: true, color: NAVY, valign: "middle" });
  [4, 8, 13].forEach((k) => s.addShape(pres.shapes.DIAMOND, { x: gx + k * gw + (gw - 0.26) / 2, y: y8 + 0.05, w: 0.26, h: 0.26, fill: { color: k === 13 ? GREEN : NAVY }, line: { type: "none" } }));
  T(s, "Shaded = holiday weeks (Thanksgiving, Christmas, New Year): lighter workload planned", { x: gx, y: gy + 8 * rh + 0.04, w: 12.7 - gx, h: 0.28, fontSize: 10, italic: true, color: AMBER, align: "right" });
  const exits = [["DAY 30  ·  DEC 17", ["3 cells confirmed, charters signed", "Baselines agreed with Finance", "Quick wins under way"], STEEL], ["DAY 60  ·  WEEK OF JAN 11", ["First results verified", "New standards drafted", "Wins shared across shifts"], MID], ["DAY 90  ·  FEB 15", ["Gains sustained and presented", "Teams recognized", "Stage 2 projects chartered"], NAVY]];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1, y = 5.6;
    card(s, x, y, 3.9, 1.15, { fill: LIGHT, name: "Exit criteria " + exits[i][0] });
    s.addShape(pres.shapes.DIAMOND, { x: x + 0.15, y: y + 0.12, w: 0.24, h: 0.24, fill: { color: i === 2 ? GREEN : NAVY }, line: { type: "none" } });
    T(s, exits[i][0], { x: x + 0.5, y: y + 0.08, w: 3.3, h: 0.32, fontSize: 11, bold: true, color: exits[i][2], charSpacing: 1, valign: "middle" });
    T(s, bullets(exits[i][1]), { x: x + 0.2, y: y + 0.42, w: 3.6, h: 0.7, fontSize: 11, paraSpaceAfter: 0 });
  }
  s.addNotes(`(About 3 minutes.)

Here's the 90-day plan, week by week, from today through February 15th. The goal is in the headline: three projects (one per shift) chartered, delivered and recognized in 90 days. That's deliberately modest. I'd rather we finish three than start ten.

Walk the bars top to bottom.

This week and next: confirm cells and sign charters.

Weeks two to four: measure baselines, with Randy agreeing the valuation method.

Weeks three to six: implement one quick win per cell.

Then, through the holidays and into January: verify results and standardize. Verification takes time on purpose. We want to see the gain hold.

Mid to late January: present at WOr, at the first meeting after your results are verified.

Late January into February: select and charter your Stage 2 projects.

Notice the shaded columns. Thanksgiving week and the two holiday weeks at the end of December are realistic about staffing and schedules, so we've planned less work there, and coaching is lighter those weeks.

The green dots are weekly CI coaching: 20 to 30 minutes per cell, at the line. The diamonds are leadership reviews at day 30, day 60 and day 90.

The three boxes at the bottom are the completion criteria for each review, so everyone knows what "on track" means. Day 30: cells confirmed, charters signed, baselines agreed with Finance, quick wins under way. Day 60: first results verified, standards drafted, wins shared. Day 90: gains sustained and presented, teams recognized, Stage 2 chartered.

Ask the room: looking at your shift's schedule for the next few months (peak runs, planned shutdowns), where does this plan collide with reality? Let's adjust it now.`);

  // ================= 19. CALL TO ACTION =================
  s = newSlide("SECTION_DARK", "Launch Plan", "CALL TO ACTION  ·  FIRST ASSIGNMENTS", "Every Shift. Every Team. Every Improvement Counts.");
  const acts = [
    ["FaUsers", "Confirm cell members", "Supervisors, with Quality and Maintenance leads", "Fri, Nov 20"],
    ["FaHardHat", "Assign a supervisor project owner", "Each cell; one owner for cross-shift work", "Fri, Nov 20"],
    ["FaCalendarCheck", "Schedule the first CI coaching session", "CI Manager, with each cell", "Tue, Nov 24"],
    ["FaLightbulb", "Identify one opportunity per cell", "Cell team, with operator input", "Wed, Dec 2"],
    ["FaRulerCombined", "Establish baseline data", "Project owner; method agreed with Finance", "Fri, Dec 11"],
  ];
  const aw = 2.3, ag = (12.1 - 5 * aw) / 4;
  for (let i = 0; i < 5; i++) {
    const x = 0.6 + i * (aw + ag);
    card(s, x, 2.05, aw, 2.95, { fill: NAVY, line: STEEL, lineW: 1, name: "Assignment " + (i + 1) });
    await iconCircle(s, acts[i][0], x + 0.22, 2.27, 0.66, GREEN);
    T(s, String(i + 1), { x: x + 1.55, y: 2.27, w: 0.55, h: 0.66, fontSize: 28, bold: true, color: STEEL, align: "right", valign: "middle" });
    T(s, acts[i][1], { x: x + 0.22, y: 3.08, w: aw - 0.4, h: 0.85, fontSize: 16, bold: true, color: WHITE });
    T(s, acts[i][2], { x: x + 0.22, y: 3.93, w: aw - 0.4, h: 0.55, fontSize: 11, color: PALE });
    pill(s, x + 0.22, 4.55, aw - 0.44, "DUE " + acts[i][3].toUpperCase(), GREEN, { h: 0.3, fontSize: 10 });
  }
  T(s, "Proposed due dates: confirm before we leave today.", { x: 0.6, y: 5.1, w: 12.1, h: 0.3, fontSize: 12, italic: true, color: PALE });
  T(s, "We are not launching another reporting exercise.", { x: 0.6, y: 5.55, w: 12.1, h: 0.5, fontSize: 22, color: PALE, align: "center" });
  T(s, "We are building leaders who improve the business.", { x: 0.6, y: 6.02, w: 12.1, h: 0.6, fontSize: 30, bold: true, color: WHITE, align: "center" });
  s.addNotes(`(About 3 minutes, then open Q&A.)

Every shift, every team, every improvement counts. Here are the first five assignments, in date order, each with an owner and a proposed due date.

One and two, by this Friday, November 20th: confirm your cell members with your Quality and Maintenance leads, and name the supervisor project owner for each cell. If a project crosses shifts, agree on one owner.

Three, by next Tuesday, November 24th: I'll get the first coaching session on the calendar with every cell. That one's on me.

Four, by Wednesday, December 2nd (we're allowing for Thanksgiving week): identify one improvement opportunity per cell. Use the funnel, and ask your operators.

Five, by Friday, December 11th: establish your baseline data, with the method agreed with Randy.

These dates are proposed. Let's confirm them before anyone leaves the room. If a date doesn't work for your shift, tell me now.

I'll leave you with this. We are not launching another reporting exercise. Nobody needs another spreadsheet. We are building leaders who improve the business. You already solve problems every shift. This program gives that work structure, support and the recognition it deserves.

Open it up: what questions do you have? And what could make this program fail? I'd genuinely rather hear it now than in month two.

Thank you. Let's go make some improvements.`);

  // ================= APPENDIX =================
  pres.addSection({ title: "Appendix" });
  s = newSlide("SECTION_DARK", "Appendix", "APPENDIX", "Tools and templates for supervisors");
  const apx = ["A1  Sample project charter", "A2  Example changeover project", "A3  Example 5S project", "A4  Suggested supervisor KPIs", "A5  Project selection scoring matrix", "A6  Example financial savings calculation"];
  for (let i = 0; i < 6; i++) {
    const x = 0.6 + (i % 2) * 6.15, y = 2.2 + Math.floor(i / 2) * 1.05;
    card(s, x, y, 5.95, 0.85, { fill: NAVY, line: STEEL, lineW: 1 });
    T(s, apx[i], { x: x + 0.3, y, w: 5.4, h: 0.85, fontSize: 19, bold: true, color: WHITE, valign: "middle" });
  }
  T(s, "All numbers in the appendix are illustrative examples, not C.H. Guenther plant data.", { x: 0.6, y: 5.6, w: 12.1, h: 0.4, fontSize: 14, italic: true, color: PALE, align: "center" });
  s.addNotes(`The appendix has practical tools supervisors can use straight away: a sample charter, worked changeover and 5S examples, suggested KPIs, the scoring matrix and a worked financial calculation.

Every number in these examples is illustrative, not our plant data. They show the method, not targets. I'll send editable versions to everyone after today.`);

  const tblHead = (arr) => arr.map((t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: "0B2545" }, fontSize: 13, valign: "middle" } }));
  const tblOpts = (y, colW, rowH) => ({ x: 0.6, y, w: colW.reduce((a, b) => a + b, 0), colW, rowH, fontSize: 13, color: INK, valign: "middle", border: { type: "solid", pt: 0.75, color: "D6DEE8" }, margin: [3, 6, 3, 6], fill: { color: "FFFFFF" } });

  // A1
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A1", "Sample project charter");
  const charter = [
    ["Project title", "[Short, specific name, e.g. “Line __ product changeover”]", "Project owner", "[Supervisor name]"],
    ["Cell / shift", "[A / B / C, or cross-shift]", "Team members", "[Operators, Quality, Maintenance]"],
    ["Problem statement", "[What is happening, where, since when, why it matters]", "CI coach", "[CI Manager]"],
    ["Baseline", "[Value, period, data source]", "Finance reviewer", "Randy (Finance)"],
    ["Target", "[Measurable goal and date]", "Valuation method", "[Agreed with Finance at charter]"],
    ["Scope: in / out", "[What is included; what is not]", "Food-safety review", "[Quality sign-off needed? Y/N]"],
    ["Expected benefit type", "[Hard savings / cost avoidance / capacity / nonfinancial]", "Key dates", "[Baseline · verify · WOr]"],
  ];
  s.addTable(charter.map((r) => r.map((c, k) => ({ text: c, options: k % 2 === 0 ? { bold: true, color: "0B2545", fill: { color: "EEF2F6" } } : { italic: c.startsWith("["), color: c.startsWith("[") ? "5B7FA6" : "1A2433" } }))), tblOpts(1.7, [2.0, 4.3, 1.9, 3.9], 0.66));
  T(s, "Fields in [brackets] are completed by the project owner. Keep the charter to one page.", { x: 0.6, y: 6.5, w: 12.1, h: 0.3, fontSize: 12, italic: true, color: GRAY });
  s.addNotes(`The sample project charter. One page on purpose: if it won't fit on one page, it's probably too big for a first project.

The fields that matter most are the problem statement, baseline and target. Two others to point out: the valuation method, which you agree with Randy when you write the charter, not at the end; and the food-safety review, so Quality signs off if the change touches food contact, sanitation or specs.`);

  // A2
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A2", "Example changeover project");
  illus(s, 9.4, 0.32, 3.3);
  const coRows = [tblHead(["Step", "What the team did (example)", "Example result"]),
    ["Baseline", "Timed 5 changeovers on one line, last good to first good product", "Average 40 min"],
    ["Observe", "Filmed one changeover with crew consent; mapped every task", "8 min of waiting found"],
    ["Separate", "Split tasks into internal (line stopped) and external (line running)", "6 tasks can move to external"],
    ["Improve", "Pre-staged parts cart, pre-set guides, Maintenance on standby", "Waiting and motion reduced"],
    ["Standardize", "One-page changeover checklist; trained all three shifts", "Same method on A/B/C"],
    ["Verify", "Timed the next 10 changeovers", "Average 32 min, held"]];
  s.addTable(coRows.map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: k === 0 ? { bold: true, color: "0B2545" } : k === 2 ? { bold: true, color: "2F9E6E" } : {} }))), tblOpts(1.7, [1.8, 6.8, 3.5], 0.58));
  card(s, 0.6, 5.95, 12.1, 0.75, { fill: LIGHT });
  T(s, [{ text: "Reported as: ", options: { bold: true, color: NAVY } }, { text: "8 min of capacity per changeover. Hard savings only for paid hours Finance confirms were removed (see A6)." }], { x: 0.85, y: 5.95, w: 11.7, h: 0.75, fontSize: 14, valign: "middle" });
  s.addNotes(`A worked changeover example, illustrative only, not our numbers.

The pattern is what matters: time several changeovers for a real baseline; watch one closely with the crew; separate internal from external tasks; move what you can to external; standardize across all three shifts; then verify over a run of changeovers.

And the bottom line: it's reported as capacity unless Finance confirms paid hours actually came out. A6 shows how that's calculated.`);

  // A3
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A3", "Example 5S project: changeover tool station");
  illus(s, 9.4, 0.32, 3.3);
  const s5 = [["Sort", "Red-tagged unused tools and obsolete change parts, with the crew", "FaTrashAlt"], ["Set in Order", "Shadow board at point of use; labeled bins for change parts", "FaThLarge"], ["Shine", "Cleaning and inspection added to the sanitation checklist (with Quality)", "FaBroom"], ["Standardize", "Photo standard posted; plant color code for food-contact tools", "FaTags"], ["Sustain", "Weekly 5-minute supervisor audit; score posted at the station", "FaLock"]];
  for (let i = 0; i < 5; i++) {
    const y = 1.7 + i * 0.86;
    card(s, 0.6, y, 7.6, 0.74, { fill: LIGHT });
    await iconCircle(s, s5[i][2], 0.72, y + 0.1, 0.54, i === 4 ? GREEN : NAVY);
    T(s, s5[i][0], { x: 1.45, y, w: 1.75, h: 0.74, fontSize: 15, bold: true, color: NAVY, valign: "middle" });
    T(s, s5[i][1], { x: 3.2, y, w: 4.85, h: 0.74, fontSize: 13, valign: "middle" });
  }
  card(s, 8.5, 1.7, 4.2, 4.18, { fill: NAVY });
  T(s, "Example measures", { x: 8.8, y: 1.9, w: 3.6, h: 0.4, fontSize: 18, bold: true, color: WHITE });
  [["Tool search time", "e.g. 4 min → under 1 min"], ["5S audit score", "e.g. 2.1 → 4.0 of 5"], ["Missing-tool events", "e.g. tracked weekly"]].forEach((m, k) => {
    T(s, m[0], { x: 8.8, y: 2.5 + k * 1.05, w: 3.6, h: 0.35, fontSize: 15, bold: true, color: GREEN });
    T(s, m[1], { x: 8.8, y: 2.85 + k * 1.05, w: 3.6, h: 0.35, fontSize: 14, color: PALE });
  });
  T(s, "Benefit type: mostly capacity and nonfinancial (safety, food safety, morale).", { x: 0.6, y: 6.15, w: 12.1, h: 0.4, fontSize: 14, italic: true, color: GRAY });
  s.addNotes(`An example 5S project, again illustrative. The scope is deliberately small: one changeover tool station.

Sort with the crew, give everything a home at the point of use, link cleaning and inspection to the sanitation checklist with Quality, post a photo standard using the plant's existing color code, and sustain it with a quick weekly audit.

Most 5S benefit is capacity and nonfinancial, and that's fine. Report it honestly.`);

  // A4
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A4", "Suggested supervisor KPIs");
  const kRows = [tblHead(["KPI", "Definition", "Source", "Suggested review"]),
    ["Changeover time", "Minutes from last good to first good product", "Line logs / timed observation", "Per changeover; weekly trend"],
    ["Unplanned downtime", "Minutes of unplanned stops, by top cause", "Downtime tracking", "Daily; weekly Pareto"],
    ["Scrap and rework", "Product lost or reworked (lb, units or cases)", "Production / Quality records", "Daily; weekly trend"],
    ["5S audit score", "Score against the area standard", "Supervisor / cell audit", "Weekly"],
    ["Ideas raised", "Improvement ideas logged by the crew", "Cell idea log", "Weekly"],
    ["Idea response rate", "% of ideas answered within one week", "Cell idea log", "Weekly"],
    ["Action completion", "% of project actions closed on time", "Project action list", "Weekly at coaching"],
    ["Validated value", "Finance-reviewed benefit, by type", "Finance (Randy)", "At project close"]];
  s.addTable(kRows.map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: k === 0 ? { bold: true, color: "0B2545" } : {} }))), tblOpts(1.7, [2.4, 4.4, 2.9, 2.4], 0.5));
  T(s, "Use the one or two KPIs that fit your project, and existing plant data sources where they exist.", { x: 0.6, y: 6.45, w: 12.1, h: 0.35, fontSize: 12, italic: true, color: GRAY });
  s.addNotes(`Suggested supervisor KPIs. Pick the one or two that match your project, and use data we already have wherever possible so we're not creating extra reporting.

Note the two engagement KPIs: ideas raised, and the idea response rate. The response rate measures us, not the operators. And validated value comes from Finance at project close.`);

  // A5
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A5", "Project selection scoring matrix");
  illus(s, 9.4, 0.32, 3.3, "ILLUSTRATIVE SCORES — EXAMPLE IDEAS");
  const crit = [["Safe (gate: must pass)", "Pass/Fail", "Pass", "Pass", "Pass"], ["Small enough to complete", "x3", "5", "3", "1"], ["Measurable", "x2", "5", "4", "3"], ["Within team's influence", "x3", "5", "4", "2"], ["Value to operations", "x2", "3", "4", "5"], ["Resources available", "x1", "4", "3", "2"]];
  const score = (col) => crit.slice(1).reduce((a, r) => a + Number(r[col]) * Number(r[1].slice(1)), 0);
  const sRows = [tblHead(["Criterion (score 1–5)", "Weight", "Idea A: stage change parts", "Idea B: fix recurring jam", "Idea C: redesign pack area"])]
    .concat(crit.map((r) => r.map((c, k) => ({ text: c, options: k === 0 ? { bold: true, color: "0B2545" } : { align: "center" } }))))
    .concat([[{ text: "Weighted total (max 55)", options: { bold: true, color: "FFFFFF", fill: { color: "13315C" } } }, { text: "", options: { fill: { color: "13315C" } } }].concat([2, 3, 4].map((k) => ({ text: String(score(k)), options: { bold: true, align: "center", color: "FFFFFF", fill: { color: k === 2 ? "2F9E6E" : "13315C" } } })))]);
  s.addTable(sRows, tblOpts(1.7, [3.3, 1.3, 2.5, 2.5, 2.5], 0.52));
  T(s, [{ text: "How to use: ", options: { bold: true, color: NAVY } }, { text: "any idea that fails the safety gate is out. Score the rest 1–5 as a team, multiply by the weight, and start with the highest total. Here, Idea A becomes the first project and Idea C is parked for Stage 3." }], { x: 0.6, y: 6.0, w: 12.1, h: 0.75, fontSize: 14 });
  s.addNotes(`If you want something more structured than the funnel, use this matrix. Safety is a gate: fail it and the idea is out.

Score each idea 1 to 5 as a team and multiply by the weight. "Small enough" and "within our influence" carry the most weight on purpose, because for first projects, finishing matters more than size.

The scores are illustrative. Notice the redesign idea scores highest on value but lowest overall. It's not a bad idea; it's a Stage 3 idea.`);

  // A6
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A6", "Example savings calculation: same gain, three outcomes");
  illus(s, 9.4, 0.32, 3.3, "ILLUSTRATIVE — ALL INPUTS HYPOTHETICAL");
  const fin = [tblHead(["Gross operational gain", "Calculation (assumed inputs)", "Result"]),
    ["Time saved per changeover", "40 min baseline − 32 min verified", "8 min"],
    ["Changeovers per year (assumed)", "5 per week × 50 weeks", "250"],
    ["Line time freed", "8 min × 250 ÷ 60", "≈ 33 line hours"],
    ["Crew hours freed (assumed crew of 6)", "33.3 h × 6 people", "≈ 200 crew hours"]];
  s.addTable(fin.map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: k === 0 ? { bold: true, color: "0B2545" } : k === 2 ? { bold: true, align: "right" } : {} }))), tblOpts(1.65, [4.3, 5.0, 2.8], 0.46));
  T(s, "What Finance reports depends on what actually changes:", { x: 0.6, y: 4.08, w: 12.1, h: 0.35, fontSize: 15, bold: true, color: NAVY });
  const sc = [
    ["A", "Paid hours unchanged", "Capacity: ≈ 200 crew hours", "Hard savings: $0", "Real value, but no cost went down.", STEEL],
    ["B", "Overtime actually reduced", "120 OT hours removed × $45/h = $5,400", "Less $800 one-time cost = $4,600", "Only hours Finance confirms in payroll count, not the full 200.", GREEN],
    ["C", "Freed time makes more product", "Value = contribution margin on extra cases sold", "Finance calculates; only if demand exists", "Unsold output is inventory, not savings.", MID],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1, y = 4.5;
    card(s, x, y, 3.9, 2.22, { fill: i === 1 ? NAVY : LIGHT, name: "Scenario " + sc[i][0] });
    s.addShape(pres.shapes.OVAL, { x: x + 0.18, y: y + 0.15, w: 0.42, h: 0.42, fill: { color: sc[i][5] }, line: { type: "none" } });
    T(s, sc[i][0], { x: x + 0.18, y: y + 0.15, w: 0.42, h: 0.42, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    const dark = i === 1;
    T(s, sc[i][1], { x: x + 0.72, y: y + 0.15, w: 3.05, h: 0.42, fontSize: 15, bold: true, color: dark ? WHITE : NAVY, valign: "middle" });
    T(s, [{ text: sc[i][2], options: { breakLine: true } }, { text: sc[i][3], options: { bold: true, color: dark ? GREEN : NAVY } }], { x: x + 0.2, y: y + 0.7, w: 3.55, h: 0.8, fontSize: 13, color: dark ? WHITE : INK });
    T(s, sc[i][4], { x: x + 0.2, y: y + 1.55, w: 3.55, h: 0.6, fontSize: 11, italic: true, color: dark ? PALE : GRAY });
  }
  s.addNotes(`A worked financial example. Every input is hypothetical: round numbers to make the logic easy to follow, not our rates, volumes or costs.

The top table is the gross operational gain: 8 minutes per changeover, about 250 changeovers a year, about 33 line hours, and with a crew of six, about 200 crew hours freed up.

Here's the key point: the same gain can be worth very different amounts, depending on what actually changes.

Scenario A: nobody's paid hours change. We report about 200 crew hours of capacity. Hard savings are zero. Still valuable, but no cost went down.

Scenario B: overtime actually comes down. Finance confirms in payroll that, say, 120 overtime hours were removed compared with the baseline. At an assumed $45 an hour that's $5,400, less an $800 one-time cost for the staging cart, so $4,600 of validated hard savings. Notice it's the 120 confirmed hours, not the theoretical 200.

Scenario C: we use the freed time to make more product. That's worth the contribution margin on the extra cases, but only if we actually sell them. Finance does that calculation. Unsold product is inventory, not savings.

That's the discipline that keeps this program credible with leadership.`);

  await pres.writeFile({ fileName: FILE });
  await applyTheme(FILE, THEME);
  console.log("wrote", FILE);
})().catch((e) => { console.error(e); process.exit(1); });
