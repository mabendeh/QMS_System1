// Supervisor Productivity & Engagement Program — kickoff deck builder
// C.H. Guenther & Son, Duncanville. Run: node build_deck.js <outdir>
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const SKILL = process.env.PPTX_SKILL;
const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));
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
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
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
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + hex, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return (iconCache[key] = "image/png;base64," + buf.toString("base64"));
}
const shadow = () => ({ type: "outer", color: "0B2545", opacity: 0.14, blur: 6, offset: 2, angle: 90 });

// ---------- layouts ----------
const FOOT = "C.H. GUENTHER & SON  |  DUNCANVILLE";
const PROGRAM = "Supervisor Productivity & Engagement Program";
pres.defineSlideMaster({
  title: "CONTENT_LIGHT",
  background: { color: "FFFFFF" },
  objects: [
    { text: { text: FOOT, options: { x: 0.6, y: 7.02, w: 5, h: 0.3, fontSize: 9, bold: true, color: STEEL, charSpacing: 1, margin: 0, objectName: "Logo placeholder (replace with official logo)" } } },
    { text: { text: PROGRAM, options: { x: 5.2, y: 7.02, w: 6.6, h: 0.3, fontSize: 9, color: GRAY, align: "right", margin: 0 } } },
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.32, w: 12.1, h: 0.34, fontSize: 12, bold: true, color: GREEN, charSpacing: 2, margin: 0, valign: "middle" }, text: "SECTION" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.66, w: 12.1, h: 0.8, fontSize: 34, bold: true, color: NAVY, margin: 0, valign: "middle", align: "left" }, text: "Slide title" } },
  ],
  slideNumber: { x: 12.15, y: 7.02, w: 0.6, h: 0.3, fontSize: 9, color: GRAY, align: "right" },
});
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { path: path.join(__dirname, "hero.png") },
  objects: [
    { text: { text: FOOT, options: { x: 0.6, y: 0.45, w: 6, h: 0.35, fontSize: 12, bold: true, color: WHITE, charSpacing: 2, margin: 0, objectName: "Logo placeholder (replace with official logo)" } } },
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 1.15, w: 8.5, h: 0.4, fontSize: 14, bold: true, color: GREEN, charSpacing: 2, margin: 0 }, text: "KICKER" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.6, w: 8.6, h: 1.6, fontSize: 46, bold: true, color: WHITE, margin: 0, valign: "top", align: "left" }, text: "Title" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 0.6, y: 3.25, w: 8.6, h: 0.5, fontSize: 22, color: PALE, margin: 0 }, text: "Subtitle" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION_DARK",
  background: { path: path.join(__dirname, "section.png") },
  objects: [
    { text: { text: FOOT, options: { x: 0.6, y: 7.02, w: 5, h: 0.3, fontSize: 9, bold: true, color: PALE, charSpacing: 1, margin: 0, objectName: "Logo placeholder (replace with official logo)" } } },
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
async function discuss(s, q, y = 6.28) {
  card(s, 0.6, y, 12.1, 0.5, { fill: WHITE, line: GREEN, lineW: 1.25, name: "Discussion question" });
  s.addImage({ data: await icon("FaComments", HEX.accent2), x: 0.78, y: y + 0.11, w: 0.28, h: 0.28, altText: "Discussion icon" });
  T(s, [{ text: "DISCUSS  ", options: { bold: true, color: GREEN, fontSize: 12, charSpacing: 1 } }, { text: q, options: { italic: true, color: NAVY, fontSize: 15 } }], { x: 1.2, y, w: 11.3, h: 0.5, valign: "middle" });
}
function illus(s, x, y, w = 3.3, label = "ILLUSTRATIVE EXAMPLE — NOT PLANT DATA") {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.3, rectRadius: 0.15, fill: { color: AMBER }, line: { type: "none" }, objectName: "Illustrative label" });
  T(s, label, { x, y, w, h: 0.3, fontSize: 10, bold: true, color: WHITE, align: "center", valign: "middle", charSpacing: 1 });
}
function arrowRight(s, x, y, w, color = GRAY) {
  s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width: 2, endArrowType: "triangle" } });
}

const DEPTS = [
  ["FaIndustry", "Production"], ["FaClipboardCheck", "Quality"], ["FaWrench", "Maintenance"],
  ["FaDollarSign", "Finance"], ["FaSyncAlt", "Continuous Improvement"],
];

(async () => {
  // =============== 1. TITLE ===============
  pres.addSection({ title: "Launch" });
  let s = newSlide("TITLE_DARK", "Launch", "PROGRAM LAUNCH  ·  TUESDAY, NOVEMBER 17, 2026", "Supervisor Productivity & Engagement Program");
  s.addText("Small Improvements. Stronger Teams. Measurable Results.", { placeholder: "subtitle" });
  T(s, "C.H. Guenther & Son — Duncanville", { x: 0.6, y: 3.8, w: 8.6, h: 0.4, fontSize: 16, bold: true, color: WHITE });
  // department icon row — top right glass panel
  card(s, 8.95, 0.42, 3.8, 0.62, { fill: NAVY, name: "Department icon strip" });
  for (let i = 0; i < 5; i++) {
    await iconCircle(s, DEPTS[i][0], 9.12 + i * 0.72, 0.5, 0.46, i === 4 ? GREEN : MID);
  }
  s.addNotes(`Good morning, everyone, and thanks for being here. I know how much it costs to pull supervisors, Quality, Maintenance and Finance into the same room at the same time, so I'll make this worth your 45 minutes.

Today we're launching something I'm genuinely excited about: the Supervisor Productivity and Engagement Program. The short version is right on the slide: small improvements, stronger teams, measurable results.

I want to be clear from the start about what this is NOT. This is not another report you have to fill out, and it's not CI showing up to tell you how to run your line. This is about giving you, the supervisors, the tools, support and time to fix the things that frustrate you and your people every shift.

The five icons up top are the five groups in this room: Production, Quality, Maintenance, Finance and CI. Every one of them has a role in this program, and we'll walk through exactly what that looks like.

By the end of today I'd like every shift to leave with three things: a cell team, a first improvement idea, and a date for our first coaching session.

Before we get going: quick show of hands. Who has an idea for something on their line that they've wanted to fix for months but never had the time or support? Keep that idea in your head. We're going to come back to it.`);

  // =============== 2. WHY ===============
  pres.addSection({ title: "Program Overview" });
  s = newSlide("CONTENT_LIGHT", "Program Overview", "WHY ARE WE HERE?", "Three goals for this program");
  const goals = [
    ["FaHardHat", "Empower supervisors", "Supervisors lead real improvements on their own lines, with coaching and tools behind them."],
    ["FaUsers", "Engage our people", "Build employee engagement and teamwork across Production, Quality and Maintenance."],
    ["FaChartLine", "Improve performance", "Deliver measurable gains in productivity, quality and cost, verified by data."],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15;
    card(s, x, 1.75, 3.8, 2.95, { fill: LIGHT, name: "Goal card " + (i + 1) });
    T(s, "0" + (i + 1), { x: x + 0.35, y: 1.95, w: 1.2, h: 0.75, fontSize: 40, bold: true, color: PALE });
    await iconCircle(s, goals[i][0], x + 2.85, 2.0, 0.65, i === 2 ? GREEN : MID);
    T(s, goals[i][1], { x: x + 0.35, y: 2.85, w: 3.2, h: 0.45, fontSize: 20, bold: true, color: NAVY });
    T(s, goals[i][2], { x: x + 0.35, y: 3.35, w: 3.15, h: 1.2, fontSize: 15 });
  }
  card(s, 0.6, 4.95, 12.1, 1.05, { fill: NAVY, name: "Central message" });
  await iconCircle(s, "FaSyncAlt", 0.95, 5.15, 0.65, GREEN);
  T(s, "Continuous improvement belongs to everyone.", { x: 1.85, y: 4.95, w: 10.6, h: 1.05, fontSize: 28, bold: true, color: WHITE, valign: "middle" });
  await discuss(s, "What does “continuous improvement” look like on your shift today?");
  s.addNotes(`So why are we here? Three reasons.

First, we want to empower supervisors. You know your lines, your people and your problems better than anyone. Right now a lot of improvement work flows through CI or engineering. We want to flip that so you're leading and we're supporting.

Second, engagement. Our operators see waste every day: the walking, the waiting, the searching for tools. When we ask them and actually act on what they say, people feel ownership. We also want Production, Quality and Maintenance working on problems together, not handing them back and forth.

Third, performance. At the end of the day this has to show up in productivity, quality and cost, and we're going to measure it honestly.

The message at the bottom is the one I want you to remember: continuous improvement belongs to everyone. It isn't a department. CI is a support function. The improvement happens on your floor.

Let me ask you this one, and I'd really like a couple of answers: what does continuous improvement look like on your shift today? Is it happening? Is it formal or informal? What gets in the way?

(Let two or three people answer. Don't defend anything; just listen and write the barriers down. They're useful later.)`);

  // =============== 3. VISION ===============
  s = newSlide("CONTENT_LIGHT", "Program Overview", "OUR VISION", "Every improvement builds a leader");
  const vision = [
    ["FaSearch", "Identify", "Spot a real problem on the floor"],
    ["FaHandshake", "Engage", "Bring the team into the fix"],
    ["FaCogs", "Improve", "Test and implement the change"],
    ["FaChartLine", "Measure", "Prove it with before/after data"],
    ["FaLock", "Sustain", "Standardize so it sticks"],
    ["FaTrophy", "Recognize", "Celebrate the team's result"],
  ];
  const vw = 1.75, vg = (12.1 - 6 * vw) / 5;
  for (let i = 0; i < 6; i++) {
    const x = 0.6 + i * (vw + vg);
    await iconCircle(s, vision[i][0], x + (vw - 1.1) / 2, 1.85, 1.1, i === 5 ? GREEN : i % 2 ? MID : NAVY);
    T(s, vision[i][1], { x, y: 3.1, w: vw, h: 0.42, fontSize: 19, bold: true, color: NAVY, align: "center" });
    T(s, vision[i][2], { x, y: 3.52, w: vw, h: 0.65, fontSize: 13, color: INK, align: "center" });
    if (i < 5) arrowRight(s, x + vw / 2 + 0.62, 2.4, vg + vw - 1.24, STEEL);
  }
  // loop back
  s.addShape(pres.shapes.LINE, { x: 1.475, y: 4.45, w: 10.35, h: 0, line: { color: PALE, width: 2, dashType: "dash", beginArrowType: "triangle" } });
  T(s, "…and repeat, with a bigger problem each time", { x: 4.4, y: 4.52, w: 4.5, h: 0.3, fontSize: 12, italic: true, color: GRAY, align: "center" });
  card(s, 0.6, 5.2, 12.1, 1.25, { fill: LIGHT, name: "Leadership skills" });
  T(s, "Leadership skills built along the way", { x: 0.9, y: 5.35, w: 4.0, h: 0.35, fontSize: 16, bold: true, color: NAVY });
  const skills = ["Problem solving with data", "Coaching and involving a team", "Presenting results to leadership"];
  for (let i = 0; i < 3; i++) {
    const x = 0.9 + i * 3.95;
    s.addImage({ data: await icon("FaCheckCircle", HEX.accent2), x, y: 5.83, w: 0.3, h: 0.3, altText: "Check" });
    T(s, skills[i], { x: x + 0.42, y: 5.78, w: 3.4, h: 0.4, fontSize: 15, valign: "middle" });
  }
  s.addNotes(`Here's the vision in one picture. Every project follows the same simple path: identify, engage, improve, measure, sustain, recognize.

Identify means you spot a real problem; not something from a corporate list, something your team actually lives with. Engage means you bring the operators, Quality and Maintenance into the fix early. Improve is where we try the change. Measure means we prove it with before and after data, not just a feeling. Sustain means we standardize it so it doesn't slide back in three weeks. And Recognize: we celebrate the team that did it.

Notice the dotted line along the bottom. When you finish one loop, you go around again, usually with a slightly bigger problem.

Here's the part I really want to land: every one of these projects is a leadership rep. When you lead one, you're practicing problem solving with data, coaching a team, and presenting results to leadership. Those are exactly the skills that move people up in this business. So yes, we want the productivity gains, but we're also building the next generation of leaders in this plant.

Question for the room: of these six steps, which one do you think we're weakest at as a plant today? My guess is Sustain, but I'd love to hear yours.`);

  // =============== 4. CELLS ===============
  s = newSlide("CONTENT_LIGHT", "Program Overview", "STRUCTURE", "How our improvement cells work");
  card(s, 0.6, 1.7, 12.1, 0.62, { fill: GREEN, name: "CI support band" });
  await iconCircle(s, "FaSyncAlt", 0.78, 1.76, 0.5, NAVY);
  T(s, [{ text: "Continuous Improvement  ", options: { bold: true } }, { text: "coaches every cell: tools, methods and removing barriers" }], { x: 1.45, y: 1.7, w: 11.0, h: 0.62, fontSize: 16, color: WHITE, valign: "middle" });
  const cells = ["A Shift", "B Shift", "C Shift"];
  const members = [["FaIndustry", "Production supervisor", "Project owner"], ["FaClipboardCheck", "Quality", "Food safety & specs"], ["FaWrench", "Maintenance", "Equipment expertise"]];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15;
    s.addShape(pres.shapes.LINE, { x: x + 1.9, y: 2.32, w: 0, h: 0.28, line: { color: GREEN, width: 2 } });
    card(s, x, 2.6, 3.8, 2.6, { fill: LIGHT, shadow: true, name: cells[i] + " improvement cell" });
    T(s, cells[i] + " Improvement Cell", { x: x + 0.25, y: 2.72, w: 3.3, h: 0.4, fontSize: 17, bold: true, color: NAVY });
    for (let j = 0; j < 3; j++) {
      const y = 3.2 + j * 0.64;
      card(s, x + 0.2, y, 3.4, 0.54, { fill: WHITE });
      await iconCircle(s, members[j][0], x + 0.3, y + 0.07, 0.4, j === 0 ? NAVY : MID);
      T(s, [{ text: members[j][1], options: { bold: true, fontSize: 14, color: NAVY, breakLine: true } }, { text: members[j][2], options: { fontSize: 11, color: GRAY } }], { x: x + 0.82, y: y + 0.02, w: 2.7, h: 0.5, valign: "middle" });
    }
    s.addShape(pres.shapes.LINE, { x: x + 1.9, y: 5.2, w: 0, h: 0.75, line: { color: STEEL, width: 2 } });
  }
  card(s, 1.6, 5.37, 10.1, 0.42, { fill: WHITE, line: STEEL, dash: "dash", name: "Cross-shift note" });
  T(s, [{ text: "Cross-shift projects: ", options: { bold: true, color: NAVY } }, { text: "one shared owner, coordinated actions handed off across A, B and C shifts" }], { x: 1.6, y: 5.37, w: 10.1, h: 0.42, fontSize: 13, align: "center", valign: "middle", color: INK });
  card(s, 0.6, 5.95, 12.1, 0.62, { fill: NAVY, name: "Finance validation band" });
  await iconCircle(s, "FaDollarSign", 0.78, 6.01, 0.5, STEEL);
  T(s, [{ text: "Finance  ", options: { bold: true } }, { text: "validates the financial results of every project" }], { x: 1.45, y: 5.95, w: 11.0, h: 0.62, fontSize: 16, color: WHITE, valign: "middle" });
  s.addNotes(`Let's talk about how this is organized. We're keeping it simple: one improvement cell per shift, A, B and C.

Each cell has three core members: the production supervisor, who owns the project; someone from Quality, who makes sure we protect food safety and specs; and someone from Maintenance, who brings the equipment know-how. Three people who together can solve most of the problems that show up on a line.

Across the top, CI supports all three cells. I'll be coaching, bringing tools and templates, teaching methods like SMED and 5S, and, honestly, the most important part of my job, removing barriers. If you're stuck waiting on a part, a decision or a resource, that's my problem to chase.

Across the bottom is Finance. Finance validates the financial results. We'll go deeper on that later, but the idea is that when we say something saved money, Finance agrees with the number.

Now look at the dotted box in the middle. Some problems don't belong to one shift. A changeover that runs across shift change, for example. For those, we'll name one shared owner and coordinate the actions across shifts so nothing falls in the gap at handoff.

Question: who do you naturally go to today on Quality and Maintenance on your shift? Those might be your first cell members.`);

  // =============== 5. WHO OWNS WHAT ===============
  s = newSlide("CONTENT_LIGHT", "Program Overview", "ROLES & ACCOUNTABILITY", "Who owns what?");
  const roles = [
    ["FaHardHat", "Supervisors", "Own the project, lead the team and present results", "OWNS", NAVY],
    ["FaUsers", "Production employees", "Identify opportunities and take part in improvements", "PARTICIPATES", MID],
    ["FaClipboardCheck", "Quality", "Protect quality and food-safety requirements in every change", "SUPPORTS", STEEL],
    ["FaWrench", "Maintenance", "Provide technical support and equipment expertise", "SUPPORTS", STEEL],
    ["FaDollarSign", "Finance (Randy)", "Review and validate financial savings", "VALIDATES", STEEL],
    ["FaSyncAlt", "CI Manager", "Coach, provide tools, teach methods, remove barriers", "COACHES", GREEN],
    ["FaUserTie", "Plant leadership", "Sponsor, recognize and support improvements", "SPONSORS", MID],
  ];
  for (let i = 0; i < roles.length; i++) {
    const y = 1.7 + i * 0.66;
    card(s, 0.6, y, 8.35, 0.56, { fill: i === 0 ? PALE : LIGHT, name: roles[i][1] + " row" });
    await iconCircle(s, roles[i][0], 0.72, y + 0.07, 0.42, i === 5 ? GREEN : NAVY);
    T(s, roles[i][1], { x: 1.3, y, w: 2.2, h: 0.56, fontSize: 15, bold: true, color: NAVY, valign: "middle" });
    T(s, roles[i][2], { x: 3.5, y, w: 3.85, h: 0.56, fontSize: 13, valign: "middle" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.45, y: y + 0.13, w: 1.35, h: 0.3, rectRadius: 0.15, fill: { color: roles[i][4] }, line: { type: "none" } });
    T(s, roles[i][3], { x: 7.45, y: y + 0.13, w: 1.35, h: 0.3, fontSize: 10, bold: true, color: WHITE, align: "center", valign: "middle", charSpacing: 1 });
  }
  card(s, 9.3, 1.7, 3.4, 4.52, { fill: NAVY, name: "Key message card" });
  await iconCircle(s, "FaHardHat", 9.6, 2.0, 0.75, WHITE, HEX.dk2);
  T(s, "The supervisor leads.", { x: 9.6, y: 2.9, w: 2.9, h: 0.95, fontSize: 28, bold: true, color: WHITE });
  await iconCircle(s, "FaSyncAlt", 9.6, 4.05, 0.75, GREEN);
  T(s, "CI coaches.", { x: 9.6, y: 4.95, w: 2.9, h: 0.5, fontSize: 28, bold: true, color: GREEN });
  T(s, "Ownership stays on the floor.", { x: 9.6, y: 5.5, w: 2.9, h: 0.5, fontSize: 14, italic: true, color: PALE });
  s.addNotes(`This slide is about clarity. When everyone owns something, nobody owns it, so let's be specific.

Supervisors own the project. You lead it, you drive the actions, and you present the results. That's the big shift.

Production employees are the eyes and ears. They identify opportunities and take part in making the change. Some of the best ideas will come from the people running the line.

Quality supports by making sure every change protects quality and food safety. We never trade food safety for minutes.

Maintenance provides technical support and equipment expertise: the "can we actually do this safely on this machine" question.

Finance — Randy — reviews and validates financial savings. Randy's job isn't to block anything; it's to make sure the numbers we report are numbers we can stand behind.

My job as CI Manager: coach, provide tools, teach methods and remove barriers. I'm not going to take your project and run it for you.

And plant leadership sponsors, recognizes and supports. They make sure you have the time and the backing.

So if you remember one line, it's the one on the right: the supervisor leads, CI coaches.

Quick check: does anything on this list surprise anyone, or feel unclear? Better to sort it out now than three weeks in.`);

  // =============== 6. START SMALL ===============
  pres.addSection({ title: "Focus Areas" });
  s = newSlide("CONTENT_LIGHT", "Focus Areas", "IMPROVEMENT ROADMAP", "Start small, grow bigger");
  const stages = [
    ["STAGE 1", "Quick Wins", "5S, visual management, workplace organization, simple changeover improvements", "Scope: one area", "FaBroom", 1.2],
    ["STAGE 2", "Process Improvements", "Recurring downtime, setup losses, waste and rework", "Scope: one process or line", "FaCogs", 2.05],
    ["STAGE 3", "Larger Projects", "Cross-functional productivity, standardization and verified cost reduction", "Scope: cross-functional", "FaChartLine", 2.9],
  ];
  const base = 5.75;
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.0, h = stages[i][5] + 0.85;
    const y = base - h;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.85, h, fill: { color: [STEEL, MID, NAVY][i] }, line: { type: "none" }, objectName: stages[i][0] + " step" });
    await iconCircle(s, stages[i][4], x + 0.25, y + 0.25, 0.6, i === 2 ? GREEN : WHITE, i === 2 ? "FFFFFF" : [HEX.accent3, HEX.accent1][i]);
    T(s, stages[i][0], { x: x + 1.0, y: y + 0.22, w: 2.6, h: 0.3, fontSize: 12, bold: true, color: PALE, charSpacing: 2 });
    T(s, stages[i][1], { x: x + 1.0, y: y + 0.5, w: 2.75, h: 0.42, fontSize: 19, bold: true, color: WHITE });
    T(s, stages[i][2], { x: x + 0.25, y: y + 1.05, w: 3.4, h: 0.9, fontSize: 14, color: WHITE });
    T(s, stages[i][3], { x, y: base + 0.08, w: 3.85, h: 0.3, fontSize: 12, italic: true, color: GRAY, align: "center" });
  }
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 3.5, w: 3.0, h: -1.1, line: { color: GREEN, width: 2.5, dashType: "dash", endArrowType: "triangle" } });
  T(s, "Skills, scope and value grow with every project", { x: 0.6, y: 1.75, w: 3.7, h: 0.6, fontSize: 14, bold: true, color: GREEN });
  await discuss(s, "What Stage 1 quick win could your shift finish in the next two weeks?");
  s.addNotes(`One of the most common mistakes in improvement programs is starting too big. Somebody picks "fix packaging" as a first project, it drags on for six months, and everyone loses interest. We're not doing that.

Think of this as a staircase.

Stage 1 is quick wins: 5S, visual management, organizing a workstation, simple changeover tweaks like staging parts before the line stops. Small, fast, visible. These build confidence and show your team that ideas actually get acted on.

Stage 2 is process improvements: the downtime that keeps coming back, setup losses, waste and rework. These take a bit more data and a bit more problem solving.

Stage 3 is the larger projects: cross-functional productivity work, standardization across shifts, and cost reduction that Finance has verified.

Everybody starts on Stage 1. That's not a judgment; it's how you build the muscle. As you and your team get comfortable, the problems get bigger and so does the value.

So here's my question: what Stage 1 quick win could your shift finish in the next two weeks? Something small enough that you're confident you could actually close it out.

(Capture answers on the flip chart. These are candidate first projects.)`);

  // =============== 7. CHANGEOVER ===============
  s = newSlide("CONTENT_LIGHT", "Focus Areas", "FOCUS AREA 1", "Changeover: get back to good product faster");
  const co = [
    ["FaRulerCombined", "Establish the current baseline"],
    ["FaEye", "Observe the actual changeover"],
    ["FaSearch", "Find waiting, motion and prep waste"],
    ["FaCogs", "Apply basic SMED principles"],
    ["FaClipboardList", "Standardize the improved method"],
    ["FaCheckCircle", "Verify over repeated changeovers"],
  ];
  for (let i = 0; i < 6; i++) {
    const y = 1.75 + i * 0.72;
    await iconCircle(s, co[i][0], 0.6, y, 0.52, i === 5 ? GREEN : NAVY);
    T(s, String(i + 1), { x: 1.25, y, w: 0.35, h: 0.52, fontSize: 18, bold: true, color: PALE, valign: "middle" });
    T(s, co[i][1], { x: 1.6, y, w: 4.3, h: 0.52, fontSize: 16, valign: "middle", color: INK });
  }
  card(s, 6.3, 1.7, 6.4, 4.45, { fill: LIGHT, name: "Illustrative changeover chart" });
  T(s, "Changeover time, minutes", { x: 6.6, y: 1.85, w: 3.2, h: 0.35, fontSize: 15, bold: true, color: NAVY });
  illus(s, 9.2, 1.88, 3.3);
  s.addChart(pres.charts.BAR, [
    { name: "Internal work", labels: ["Before", "After"], values: [25, 22] },
    { name: "Searching & motion", labels: ["Before", "After"], values: [7, 6] },
    { name: "Waiting", labels: ["Before", "After"], values: [8, 4] },
  ], {
    x: 6.5, y: 2.3, w: 6.0, h: 2.75, barDir: "bar", barGrouping: "stacked", chartColors: [HEX.accent1, HEX.accent3, HEX.accent6],
    showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFontSize: 11, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMaxVal: 42,
    showLegend: true, legendPos: "b", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: HEX.dk1, barGapWidthPct: 60,
  });
  T(s, [{ text: "40 → 32 min", options: { bold: true, color: GREEN, fontSize: 22 } }, { text: "   8 minutes back per changeover (20%) in this example", options: { fontSize: 13, color: INK } }], { x: 6.6, y: 5.15, w: 5.9, h: 0.45, valign: "middle" });
  T(s, "Hypothetical numbers to show the method. Your baseline comes from your own observed changeovers.", { x: 6.6, y: 5.6, w: 5.9, h: 0.45, fontSize: 11, italic: true, color: GRAY });
  T(s, "SMED basics: do prep while the line is still running, stage tools and parts before stopping, and cut waiting between hand-offs.", { x: 0.6, y: 6.15, w: 5.4, h: 0.6, fontSize: 12, italic: true, color: GRAY });
  s.addNotes(`Our first focus area is changeover, because changeover time is time we're not making product, and it's something supervisors can directly influence.

Here's the approach, step by step.

One: establish a baseline. How long does this changeover actually take today? Measure it from last good product to first good product, not what we think it takes.

Two: go watch it. Stand there with a stopwatch and a notepad, ideally with the crew's permission and involvement. You'll see things you've never noticed.

Three: look for waste. Who's waiting? Who's walking back and forth? What preparation could have happened before we stopped the line?

Four: apply basic SMED principles. The big idea is simple: anything that can be done while the line is still running should be done while the line is still running. Staging parts, pre-setting tools, having the next ingredients ready.

Five: standardize the better method so every crew does it the same way.

Six: verify over repeated changeovers. One good changeover is a lucky day. Five or ten in a row is a result.

The chart on the right is an illustrative example only. These aren't our numbers. It shows a changeover going from 40 minutes to 32, mostly by cutting waiting and some internal work. That's 8 minutes back every time.

Question: which changeover on your shift drives you the craziest? That might be your first project.`);

  // =============== 8. 5S ===============
  s = newSlide("CONTENT_LIGHT", "Focus Areas", "FOCUS AREA 2", "5S: a place for everything, and food safe");
  const five = [
    ["1", "Sort", "Remove unneeded tools, parts and supplies from the line", "FaTrashAlt"],
    ["2", "Set in Order", "Shadow boards and labeled homes for change parts and tools", "FaThLarge"],
    ["3", "Shine", "Clean and inspect equipment, linked to the sanitation plan", "FaBroom"],
    ["4", "Standardize", "Visual labels, color codes and one-point lessons", "FaTags"],
    ["5", "Sustain", "Routine audits and supervisor walk-throughs", "FaLock"],
  ];
  const fw = 2.3, fg = (12.1 - 5 * fw) / 4;
  for (let i = 0; i < 5; i++) {
    const x = 0.6 + i * (fw + fg);
    card(s, x, 1.7, fw, 1.78, { fill: LIGHT, name: five[i][1] + " card" });
    await iconCircle(s, five[i][3], x + 0.2, 1.85, 0.55, i === 4 ? GREEN : NAVY);
    T(s, five[i][1], { x: x + 0.85, y: 1.85, w: fw - 0.95, h: 0.55, fontSize: 17, bold: true, color: NAVY, valign: "middle" });
    T(s, five[i][2], { x: x + 0.2, y: 2.5, w: fw - 0.35, h: 1.05, fontSize: 13 });
  }
  // Before / after concept
  const bx = 0.6, by = 3.68, bw = 3.95, bh = 3.07;
  card(s, bx, by, bw, bh, { fill: LIGHT, name: "Before concept" });
  T(s, "BEFORE  ·  concept", { x: bx + 0.2, y: by + 0.12, w: 3, h: 0.3, fontSize: 12, bold: true, color: GRAY, charSpacing: 1 });
  const junk = [[0.3, 0.6, 1.1, 0.25, 20], [1.3, 0.75, 0.8, 0.3, -15], [2.3, 0.6, 1.2, 0.22, 35], [0.5, 1.3, 0.6, 0.6, 10], [1.4, 1.5, 1.3, 0.2, -30], [2.8, 1.3, 0.6, 0.45, 5], [0.4, 2.15, 1.4, 0.24, -8], [2.1, 2.05, 0.9, 0.35, 25], [3.1, 2.15, 0.45, 0.45, 0]];
  junk.forEach(([x, y, w, h, r], k) => s.addShape(k % 3 === 0 ? pres.shapes.OVAL : pres.shapes.ROUNDED_RECTANGLE, { x: bx + x, y: by + y, w, h, rotate: r, rectRadius: 0.05, fill: { color: k % 2 ? GRAY : STEEL }, line: { type: "none" } }));
  const ax = 4.75;
  card(s, ax, by, bw, bh, { fill: LIGHT, name: "After concept" });
  T(s, "AFTER  ·  shadow board concept", { x: ax + 0.2, y: by + 0.12, w: 3.5, h: 0.3, fontSize: 12, bold: true, color: GREEN, charSpacing: 1 });
  const tools = ["Scraper", "Wrench", "Change part", "Brush", "Gauge", "Spare blade"];
  tools.forEach((t, k) => {
    const cx = ax + 0.25 + (k % 3) * 1.18, cy = by + 0.55 + Math.floor(k / 3) * 1.12;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w: 1.05, h: 0.6, rectRadius: 0.06, fill: { color: WHITE }, line: { color: NAVY, width: 1.25, dashType: "dash" } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx + 0.2, y: cy + 0.22, w: 0.65, h: 0.16, rectRadius: 0.05, fill: { color: k === 3 ? GREEN : MID }, line: { type: "none" } });
    T(s, t, { x: cx - 0.05, y: cy + 0.63, w: 1.15, h: 0.3, fontSize: 10, color: INK, align: "center" });
  });
  card(s, 8.9, by, 3.8, bh, { fill: NAVY, name: "Food safety callout" });
  await iconCircle(s, "FaShieldAlt", 9.1, by + 0.2, 0.55, GREEN);
  T(s, "Food safety first", { x: 9.75, y: by + 0.2, w: 2.8, h: 0.55, fontSize: 18, bold: true, color: WHITE, valign: "middle" });
  T(s, [
    { text: "Follow our sanitation program and approved chemicals", options: { bullet: true, breakLine: true } },
    { text: "Keep food-contact and non-contact tools separate", options: { bullet: true, breakLine: true } },
    { text: "Use the plant's color code; never invent a new one", options: { bullet: true, breakLine: true } },
    { text: "Use cleanable, approved materials for boards and labels", options: { bullet: true } },
  ], { x: 9.1, y: by + 0.88, w: 3.45, h: 2.1, fontSize: 13, color: WHITE, paraSpaceAfter: 4 });
  s.addNotes(`Focus area two is 5S. Most of you have heard of it, but let's make it real for a frozen-food plant.

Sort: get rid of what doesn't belong. That extra wrench, the broken scraper, the box of parts from a product we haven't run in a year. If it's not needed at that station, it goes.

Set in order: everything gets a home. Shadow boards for tools, labeled locations for change parts, so anyone on any shift can find it in seconds.

Shine: cleaning is inspecting. When you clean a piece of equipment, you find the loose bolt, the leak, the worn belt. And in our world, Shine has to line up with the sanitation plan.

Standardize: make the standard visual. Labels, color codes, one-point lessons, so the right way is obvious.

Sustain: this is the hard one. Routine audits and supervisor walk-throughs keep it from sliding back.

The before and after pictures are just concepts: a cluttered area on the left, a shadow board on the right where you can see instantly if something's missing.

And the navy box is the non-negotiable part. In food manufacturing, 5S has to support food safety, never work against it. Use our sanitation program, keep food-contact tools separate, use the plant's existing color code, and use approved, cleanable materials. Quality is your partner here: involve them before you put anything new on the floor.

Question: where on your line do people lose the most time searching for something?`);

  // =============== 9. ENGAGEMENT ===============
  s = newSlide("CONTENT_LIGHT", "Focus Areas", "FOCUS AREA 3", "Employee engagement: the people closest to the work");
  const eng = [
    ["FaComments", "Ask what slows them down", "Start every project by asking operators where the pain is"],
    ["FaLightbulb", "Encourage suggestions", "Make it easy and safe to raise an idea, big or small"],
    ["FaAward", "Recognize contributions", "Name the people behind the idea, in front of their peers"],
    ["FaUserCheck", "Give ownership", "Let employees own specific improvement actions"],
    ["FaShareAlt", "Share across shifts", "Spread what works from A to B to C shift"],
  ];
  for (let i = 0; i < 5; i++) {
    const y = 1.7 + i * 0.86;
    card(s, 0.6, y, 7.3, 0.74, { fill: LIGHT, name: eng[i][1] });
    await iconCircle(s, eng[i][0], 0.75, y + 0.1, 0.54, NAVY);
    T(s, eng[i][1], { x: 1.5, y: y + 0.06, w: 6.2, h: 0.32, fontSize: 16, bold: true, color: NAVY });
    T(s, eng[i][2], { x: 1.5, y: y + 0.38, w: 6.2, h: 0.3, fontSize: 13 });
  }
  card(s, 8.25, 1.7, 4.45, 4.18, { fill: NAVY, name: "Engagement quote" });
  T(s, "“", { x: 8.55, y: 1.75, w: 1, h: 1, fontSize: 80, bold: true, color: GREEN });
  T(s, "People support what they help improve.", { x: 8.55, y: 2.75, w: 3.9, h: 1.7, fontSize: 30, bold: true, color: WHITE });
  T(s, "When operators shape the change, they keep it going after the project ends.", { x: 8.55, y: 4.55, w: 3.85, h: 1.0, fontSize: 14, color: PALE });
  await discuss(s, "What’s one thing that slows your team down every single shift?");
  s.addNotes(`Focus area three is engagement, and honestly this is the one that makes or breaks the other two.

The people running the line every day know where the problems are. They've usually known for a long time. The question is whether we ask, and whether we act.

So five simple habits.

Ask what slows them down. Start every project by going to the operators and asking, "What makes your job harder than it needs to be?" Then be quiet and listen.

Encourage suggestions. Make it easy and safe to raise an idea. If someone gets shut down the first time, you'll never hear from them again.

Recognize contributions. When an idea works, say whose idea it was, out loud, in front of their peers.

Give ownership. Don't just collect ideas; hand out actions. "Maria, can you own the labeling for the new shadow board?" (Use a real name from your crew.) Ownership is engagement.

And share across shifts. If A shift solves something, B and C should hear about it the same week.

The quote on the right says it all: people support what they help improve. If we design a change in an office and push it onto the floor, it'll fade. If the crew builds it, they'll defend it.

So let me ask you directly: what's one thing that slows your team down every single shift? Shout it out.`);

  // =============== 10. PROJECT SELECTION ===============
  pres.addSection({ title: "Running a Project" });
  s = newSlide("CONTENT_LIGHT", "Running a Project", "PROJECT SELECTION", "How to select a project");
  const filt = [
    ["FaShieldAlt", "Safe", "No added risk to people or food safety"],
    ["FaCompress", "Small enough to complete", "Can close in weeks, not months"],
    ["FaRulerCombined", "Measurable", "Clear baseline and a number to move"],
    ["FaCrosshairs", "Within our influence", "The team can change it directly"],
    ["FaGem", "Valuable to operations", "Output, quality or cost"],
    ["FaBoxes", "Resourced", "People, parts and time are available"],
  ];
  T(s, "ALL IDEAS", { x: 0.6, y: 1.65, w: 7.2, h: 0.3, fontSize: 12, bold: true, color: GRAY, align: "center", charSpacing: 2 });
  for (let i = 0; i < 6; i++) {
    const w = 7.2 - i * 0.5, x = 0.6 + (7.2 - w) / 2, y = 2.0 + i * 0.6;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.52, rectRadius: 0.06, fill: { color: [PALE, PALE, STEEL, STEEL, MID, MID][i] }, line: { type: "none" }, objectName: "Funnel filter " + (i + 1) });
    const dark = i >= 2;
    s.addImage({ data: await icon(filt[i][0], dark ? "FFFFFF" : HEX.dk2), x: x + 0.18, y: y + 0.12, w: 0.28, h: 0.28, altText: filt[i][1] });
    T(s, [{ text: filt[i][1] + "  ", options: { bold: true, fontSize: 14 } }, { text: filt[i][2], options: { fontSize: 12 } }], { x: x + 0.55, y, w: w - 0.65, h: 0.52, valign: "middle", color: dark ? WHITE : NAVY });
  }
  card(s, 2.6, 5.62, 3.2, 0.5, { fill: GREEN, name: "Funnel output" });
  T(s, "Our first project", { x: 2.6, y: 5.62, w: 3.2, h: 0.5, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
  // examples
  card(s, 8.15, 1.7, 4.55, 2.05, { fill: LIGHT, name: "Good first projects" });
  s.addImage({ data: await icon("FaCheckCircle", HEX.accent2), x: 8.35, y: 1.85, w: 0.32, h: 0.32, altText: "Good" });
  T(s, "Good first projects", { x: 8.8, y: 1.83, w: 3.7, h: 0.36, fontSize: 16, bold: true, color: GREEN });
  T(s, [
    { text: "Stage change parts before one product changeover", options: { bullet: true, breakLine: true } },
    { text: "Shadow board for one line’s tools", options: { bullet: true, breakLine: true } },
    { text: "Cut one recurring jam on one machine", options: { bullet: true } },
  ], { x: 8.35, y: 2.3, w: 4.2, h: 1.35, fontSize: 14, paraSpaceAfter: 4 });
  card(s, 8.15, 3.95, 4.55, 2.17, { fill: LIGHT, name: "Too big for now" });
  s.addImage({ data: await icon("FaFlag", HEX.accent5), x: 8.35, y: 4.1, w: 0.3, h: 0.3, altText: "Too big" });
  T(s, "Too big for a first project", { x: 8.8, y: 4.08, w: 3.7, h: 0.36, fontSize: 16, bold: true, color: AMBER });
  T(s, [
    { text: "Redesign the whole packaging area", options: { bullet: true, breakLine: true } },
    { text: "Replace major equipment", options: { bullet: true, breakLine: true } },
    { text: "Anything needing capital approval", options: { bullet: true } },
  ], { x: 8.35, y: 4.55, w: 4.2, h: 1.45, fontSize: 14, paraSpaceAfter: 4 });
  await discuss(s, "Run your idea through the funnel. Where does it get stuck?");
  s.addNotes(`So how do you pick a good project? Think of it as a funnel. Lots of ideas go in the top; we want one good project coming out the bottom.

Six filters.

Is it safe? Nothing we do adds risk to people or to food safety. Non-negotiable.

Is it small enough to complete? We want things we can close in weeks, not months, especially for the first one.

Is it measurable? If we can't measure a before and after, we can't prove it worked.

Is it within our influence? If the fix depends on a supplier, corporate or a capital request, it's not a good first project.

Is it valuable? It should matter to output, quality or cost, not just be busy work.

And is it resourced? Do we have the people, the parts and the time?

On the right are some examples. Good first projects are focused: staging change parts for one changeover, a shadow board for one line, one recurring jam on one machine. Too big for now: redesigning packaging, replacing equipment, anything that needs capital.

There's a scoring matrix in the appendix if you want something more formal, but honestly these six questions get you most of the way.

Let's try it. Take the idea you had in your head at the start and run it through the funnel. Where does it get stuck? (Take one or two examples and work them live.)`);

  // =============== 11. JOURNEY ===============
  s = newSlide("CONTENT_LIGHT", "Running a Project", "THE SUPERVISOR'S PROJECT JOURNEY", "Seven steps, one simple PDCA cycle");
  const steps = [
    ["Identify a problem", "What’s hurting us?", "FaSearch"],
    ["Define the baseline", "How bad is it today?", "FaRulerCombined"],
    ["Set a measurable target", "What does good look like?", "FaBullseye"],
    ["Engage the team", "Who needs to be involved?", "FaUsers"],
    ["Implement the improvement", "What are we changing?", "FaCogs"],
    ["Verify the results", "Did it really work?", "FaChartLine"],
    ["Present and sustain", "How do we keep it?", "FaMicrophone"],
  ];
  const pdca = [["PLAN", 0, 4, NAVY], ["DO", 4, 1, MID], ["CHECK", 5, 1, STEEL], ["ACT", 6, 1, GREEN]];
  const sw = 12.1 / 7;
  pdca.forEach(([lab, st, n, col]) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6 + st * sw + 0.06, y: 1.7, w: n * sw - 0.12, h: 0.44, rectRadius: 0.08, fill: { color: col }, line: { type: "none" }, objectName: "PDCA " + lab });
    T(s, lab, { x: 0.6 + st * sw, y: 1.7, w: n * sw, h: 0.44, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle", charSpacing: 3 });
  });
  s.addShape(pres.shapes.LINE, { x: 0.6 + sw / 2, y: 3.0, w: 12.1 - sw, h: 0, line: { color: PALE, width: 4 } });
  for (let i = 0; i < 7; i++) {
    const cx = 0.6 + i * sw + sw / 2;
    const col = i < 4 ? NAVY : [MID, STEEL, GREEN][i - 4];
    await iconCircle(s, steps[i][2], cx - 0.42, 2.58, 0.84, col);
    T(s, String(i + 1), { x: cx - 0.3, y: 3.52, w: 0.6, h: 0.45, fontSize: 22, bold: true, color: col, align: "center" });
    T(s, steps[i][0], { x: cx - sw / 2 + 0.06, y: 3.98, w: sw - 0.12, h: 0.7, fontSize: 15, bold: true, color: NAVY, align: "center" });
    T(s, steps[i][1], { x: cx - sw / 2 + 0.06, y: 4.7, w: sw - 0.12, h: 0.55, fontSize: 12, italic: true, color: GRAY, align: "center" });
  }
  card(s, 0.6, 5.5, 12.1, 0.95, { fill: LIGHT, name: "Coaching touchpoints" });
  await iconCircle(s, "FaSyncAlt", 0.85, 5.67, 0.6, GREEN);
  T(s, [{ text: "CI coaching at every step.  ", options: { bold: true, color: NAVY } }, { text: "Bring your project to a coaching session at the baseline, the target and the verification steps. These are the points where most projects go off track." }], { x: 1.65, y: 5.5, w: 10.85, h: 0.95, fontSize: 14, valign: "middle" });
  s.addNotes(`Here's the journey every project follows, from start to finish. Seven steps, and if they look familiar, it's because it's just Plan, Do, Check, Act.

Plan is the first four steps. Identify a problem: what's hurting us? Define the baseline: how bad is it today, in numbers? Set a measurable target: what does good look like? And engage the team: who needs to be involved?

I want to slow down on Plan. Most projects that fail, fail here. They skip the baseline, or the target is vague like "improve changeover." A good target sounds like "cut changeover on line X from our measured baseline to a specific number, held over ten changeovers."

Do is step five: implement the improvement. Try it, adjust it, try it again.

Check is step six: verify the results. Did it really work, and does it hold up over time?

Act is step seven: present it and sustain it. Update the standard work, train the other shifts, and bring it to leadership.

The italic questions under each step are the questions I'll be asking you in coaching sessions. If you can answer them, you're in good shape.

At the bottom: I'll be coaching throughout, but the three moments I really want to see you are at the baseline, the target and the verification. Those are where projects tend to wobble.

Question: which of these seven steps feels hardest to you right now?`);

  // =============== 12. MEASURING ===============
  s = newSlide("CONTENT_LIGHT", "Running a Project", "MEASURING RESULTS", "What we measure, and what counts as savings");
  T(s, "OPERATIONAL GAINS  ·  measured by the team on the floor", { x: 0.6, y: 1.65, w: 7.4, h: 0.32, fontSize: 12, bold: true, color: MID, charSpacing: 1 });
  const kpis = [
    ["FaStopwatch", "Changeover minutes", "Last good to first good product"],
    ["FaTachometerAlt", "Downtime minutes", "Unplanned stops by cause"],
    ["FaRecycle", "Scrap & rework", "Pounds, units or cases lost"],
    ["FaUsers", "Employee participation", "Ideas raised, people involved"],
    ["FaThLarge", "5S audit results", "Score against the area standard"],
    ["FaTasks", "Action completion", "Actions closed on time"],
  ];
  for (let i = 0; i < 6; i++) {
    const x = 0.6 + (i % 2) * 3.75, y = 2.05 + Math.floor(i / 2) * 1.12;
    card(s, x, y, 3.55, 0.98, { fill: LIGHT, name: kpis[i][1] });
    await iconCircle(s, kpis[i][0], x + 0.15, y + 0.2, 0.58, NAVY);
    T(s, kpis[i][1], { x: x + 0.88, y: y + 0.15, w: 2.6, h: 0.35, fontSize: 15, bold: true, color: NAVY });
    T(s, kpis[i][2], { x: x + 0.88, y: y + 0.5, w: 2.6, h: 0.35, fontSize: 12, color: GRAY });
  }
  arrowRight(s, 8.0, 3.6, 0.55, GREEN);
  card(s, 8.7, 1.65, 4.0, 3.62, { fill: NAVY, name: "Financial benefits panel" });
  await iconCircle(s, "FaHandHoldingUsd", 8.95, 1.9, 0.7, GREEN);
  T(s, "Financial benefits", { x: 8.95, y: 2.7, w: 3.5, h: 0.45, fontSize: 21, bold: true, color: WHITE });
  T(s, "Validated by Finance", { x: 8.95, y: 3.12, w: 3.5, h: 0.32, fontSize: 13, bold: true, color: GREEN });
  T(s, "Operational gains are real, but they only become reported savings once Finance confirms the money actually changed.", { x: 8.95, y: 3.55, w: 3.5, h: 1.7, fontSize: 14, color: PALE });
  card(s, 0.6, 5.47, 12.1, 0.6, { fill: LIGHT, name: "Distinction" });
  T(s, [{ text: "Operational gain ", options: { bold: true, color: NAVY } }, { text: "= the process got better.     " }, { text: "Realized savings ", options: { bold: true, color: GREEN } }, { text: "= the cost actually went down and Finance agrees." }], { x: 0.85, y: 5.47, w: 11.7, h: 0.6, fontSize: 15, valign: "middle", align: "center" });
  await discuss(s, "Which of these measures could you start tracking on your line this week?");
  s.addNotes(`Let's talk about how we'll measure results, because "it feels better" isn't going to cut it.

On the left are the operational measures, the ones your team can track on the floor. Changeover minutes. Downtime minutes, ideally by cause. Scrap and rework. Employee participation: how many ideas, how many people involved. 5S audit results. And action completion: are we closing what we said we'd close?

Pick the one or two that fit your project. You don't need all six.

On the right is the financial side, and this is where I want to be really careful. An operational gain means the process got better. A realized saving means the cost actually went down, and Finance agrees.

Those are not the same thing. If we cut changeover by eight minutes but we're still running the same hours with the same crew, the process improved, and that's valuable, but we didn't save money yet. We created capacity.

Why does this matter? Because the fastest way to lose credibility with leadership is to add up a pile of "savings" that never show up in the P&L. We are not going to do that. We'll report operational gains proudly, and we'll report financial savings only when Randy has validated them.

Question for you: which of these measures could you start tracking on your line this week, with what you already have?`);

  // =============== 13. FINANCE ===============
  s = newSlide("CONTENT_LIGHT", "Running a Project", "FINANCE VALIDATION  ·  RANDY", "Turning improvements into verified value");
  const flow = ["Baseline", "Improvement", "Measured benefit", "Finance review", "Validated savings"];
  const flw = 2.25, flg = (12.1 - 5 * flw) / 4;
  for (let i = 0; i < 5; i++) {
    const x = 0.6 + i * (flw + flg);
    s.addShape(pres.shapes.CHEVRON, { x, y: 1.68, w: flw + 0.12, h: 0.7, fill: { color: i === 4 ? GREEN : i === 3 ? NAVY : [STEEL, STEEL, MID][i] }, line: { type: "none" }, objectName: "Flow " + flow[i] });
    T(s, flow[i], { x: x + 0.3, y: 1.68, w: flw - 0.4, h: 0.7, fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
  }
  const cats = [
    ["FaDollarSign", "Hard savings", "Spending actually goes down: less overtime, material or scrap cost.", "Reported as savings", GREEN],
    ["FaShieldAlt", "Cost avoidance", "A future cost we prevented, such as avoided repairs or added spend.", "Reported separately", MID],
    ["FaTachometerAlt", "Capacity / productivity", "Time or output freed up. Value depends on how we use it.", "Reported as capacity", STEEL],
    ["FaUsers", "Nonfinancial", "Safety, food safety, ergonomics, morale, ease of work.", "Reported as impact", GRAY],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + i * 3.08;
    card(s, x, 2.65, 2.86, 2.7, { fill: LIGHT, name: cats[i][1] });
    await iconCircle(s, cats[i][0], x + 0.2, 2.82, 0.6, cats[i][4]);
    T(s, cats[i][1], { x: x + 0.2, y: 3.5, w: 2.5, h: 0.4, fontSize: 17, bold: true, color: NAVY });
    T(s, cats[i][2], { x: x + 0.2, y: 3.92, w: 2.5, h: 0.95, fontSize: 13 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: 4.9, w: 2.46, h: 0.3, rectRadius: 0.15, fill: { color: cats[i][4] }, line: { type: "none" } });
    T(s, cats[i][3], { x: x + 0.2, y: 4.9, w: 2.46, h: 0.3, fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
  }
  card(s, 0.6, 5.6, 12.1, 1.0, { fill: NAVY, name: "Saved time callout" });
  await iconCircle(s, "FaStopwatch", 0.85, 5.78, 0.64, AMBER);
  T(s, [{ text: "Saved time ≠ saved money.  ", options: { bold: true, color: WHITE, fontSize: 17 } }, { text: "Minutes saved become labor savings only when paid hours actually come out (for example less overtime). Otherwise they are capacity, which is valuable, but different.", options: { color: PALE, fontSize: 14 } }], { x: 1.7, y: 5.6, w: 10.8, h: 1.0, valign: "middle" });
  s.addNotes(`This slide is about turning improvements into verified value, and it's where Randy from Finance comes in.

The flow across the top is simple. We start with a baseline. We make an improvement. We measure the benefit. Randy reviews it. And what comes out the end is validated savings: a number Finance stands behind.

I want to thank Randy for being part of this. Having Finance involved from the start is what gives this program credibility with leadership.

(Optional: invite Randy to say a few words about what he'll look for.)

Now, not all value is the same kind of value, and the four boxes in the middle explain the difference.

Hard savings: spending actually goes down. Less overtime, less material, less scrap cost. That's what gets reported as savings.

Cost avoidance: we prevented a future cost, like avoiding a repair or not needing to add a shift. Real, but reported separately.

Capacity or productivity: we freed up time or output. That's valuable, but its dollar value depends on what we do with it.

Nonfinancial: safety, food safety, ergonomics, morale. Often the most important improvements don't have a dollar sign at all.

And the big one at the bottom: saved time does not automatically equal saved money. If we save 30 minutes a day but everyone still works the same paid hours, that's capacity, not labor savings. It becomes hard savings only when the paid hours actually come out, for example less overtime.

Question: can anyone think of an improvement from the past that was "savings" on paper but never showed up in the budget?`);

  // =============== 14. PRESENTATION EXPECTATIONS ===============
  s = newSlide("CONTENT_LIGHT", "Running a Project", "LEADERSHIP PRESENTATION  ·  WOr MEETING", "Your project. Your results. Your presentation.");
  T(s, "Every project presentation answers six questions:", { x: 0.6, y: 1.7, w: 5.6, h: 0.4, fontSize: 16, bold: true, color: NAVY });
  const qs = ["What problem did we identify?", "What was the baseline?", "What did the team change?", "What were the results?", "What financial value was validated?", "How will we sustain the improvement?"];
  for (let i = 0; i < 6; i++) {
    const y = 2.25 + i * 0.66;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.48, h: 0.48, fill: { color: i === 4 ? GREEN : NAVY }, line: { type: "none" } });
    T(s, String(i + 1), { x: 0.6, y, w: 0.48, h: 0.48, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, qs[i], { x: 1.25, y, w: 4.9, h: 0.48, fontSize: 16, valign: "middle" });
  }
  T(s, "Presented by the supervisor, at the existing WOr leadership meeting.", { x: 0.6, y: 6.25, w: 5.6, h: 0.45, fontSize: 13, italic: true, color: GRAY });
  // mock one-pager
  const mx = 6.55, my = 1.7, mw = 6.15, mh = 5.0;
  card(s, mx, my, mw, mh, { fill: WHITE, line: PALE, lineW: 1, shadow: true, name: "Mock one-page project summary" });
  s.addShape(pres.shapes.RECTANGLE, { x: mx, y: my, w: mw, h: 0.55, fill: { color: NAVY }, line: { type: "none" } });
  T(s, "PROJECT ONE-PAGER  ·  SAMPLE TEMPLATE", { x: mx + 0.2, y: my, w: 4.2, h: 0.55, fontSize: 12, bold: true, color: WHITE, valign: "middle", charSpacing: 1 });
  T(s, "Cell: [A Shift]   Owner: [Supervisor]", { x: mx + 3.4, y: my, w: 2.6, h: 0.55, fontSize: 10, color: PALE, valign: "middle", align: "right" });
  const boxes = [["1 · Problem", "[Describe the problem in one sentence]"], ["2 · Baseline", "[Measured starting point + how measured]"], ["3 · What we changed", "[Key countermeasures, who did them]"], ["4 · Results", ""], ["5 · Validated value", "[Finance-reviewed category & amount]"], ["6 · Sustain plan", "[Standard work, audit, owner]"]];
  const bw2 = (mw - 0.6) / 2, bh2 = 1.38;
  for (let i = 0; i < 6; i++) {
    const x = mx + 0.2 + (i % 2) * (bw2 + 0.2), y = my + 0.7 + Math.floor(i / 2) * (bh2 + 0.1);
    card(s, x, y, bw2, bh2, { fill: LIGHT });
    T(s, boxes[i][0], { x: x + 0.12, y: y + 0.08, w: bw2 - 0.24, h: 0.3, fontSize: 12, bold: true, color: i === 4 ? GREEN : NAVY });
    if (i === 3) {
      // mini before/after bars (illustrative)
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: y + 0.5, w: 1.85, h: 0.24, fill: { color: GRAY }, line: { type: "none" } });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: y + 0.84, w: 1.35, h: 0.24, fill: { color: GREEN }, line: { type: "none" } });
      T(s, "Before", { x: x + 2.08, y: y + 0.47, w: 0.6, h: 0.3, fontSize: 10, color: GRAY });
      T(s, "After", { x: x + 1.58, y: y + 0.81, w: 0.6, h: 0.3, fontSize: 10, color: GREEN });
      T(s, "Chart of your measured results", { x: x + 0.12, y: y + 1.1, w: bw2 - 0.24, h: 0.25, fontSize: 10, italic: true, color: GRAY });
    } else {
      T(s, boxes[i][1], { x: x + 0.12, y: y + 0.42, w: bw2 - 0.24, h: 0.85, fontSize: 12, italic: true, color: GRAY });
    }
  }
  s.addNotes(`Here's where it all comes together: you present your own project to leadership at the existing WOr meeting.

I want to be clear about why. This isn't CI presenting your work. It's your project, your team's results and your presentation. That's how leadership sees what you can do, and frankly it's one of the best development opportunities in this whole program.

Every presentation answers the same six questions. What problem did we identify? What was the baseline? What did the team change? What were the results? What financial value was validated? And how will we sustain it?

If you can answer those six clearly, you've told the whole story. Five minutes is plenty.

On the right is a mock one-page format, just a sample template. One page, six boxes, matching the six questions. No 30-slide decks. I'll share an editable version of this, and I'm happy to do a dry run with anyone before they present.

Notice box five: validated value. That's where Randy's review shows up. If the value is capacity or nonfinancial, say so. That's completely fine and honest.

And this is important: put your team's names on it. Recognition starts here.

Question: how many of you have presented to the leadership team before? (For those who haven't: we'll practice together. Nobody gets thrown in cold.)`);

  // =============== 15. 90 DAYS ===============
  pres.addSection({ title: "Launch Plan" });
  s = newSlide("CONTENT_LIGHT", "Launch Plan", "ROADMAP", "Our first 90 days");
  const phases = [
    ["DAYS 1–30", "Launch", ["Confirm cells and owners", "Select first projects", "Establish baselines", "Complete initial quick wins"], STEEL],
    ["DAYS 31–60", "Build", ["Review early results", "Strengthen problem solving", "Expand practices that work", "Share wins across shifts"], MID],
    ["DAYS 61–90", "Sustain & grow", ["Standardize improvements", "Develop larger opportunities", "Recognize successful teams", "Plan the next cycle"], NAVY],
  ];
  // timeline axis
  s.addShape(pres.shapes.LINE, { x: 0.6, y: 2.05, w: 12.1, h: 0, line: { color: PALE, width: 3 } });
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.7, w: 3.9, h: 0.7, fill: { color: phases[i][3] }, line: { type: "none" }, objectName: phases[i][0] });
    T(s, [{ text: phases[i][0] + "   ", options: { bold: true, fontSize: 13, color: PALE, charSpacing: 1 } }, { text: phases[i][1], options: { bold: true, fontSize: 19, color: WHITE } }], { x: x + 0.25, y: 1.7, w: 3.5, h: 0.7, valign: "middle" });
    card(s, x, 2.5, 3.9, 1.95, { fill: LIGHT });
    T(s, phases[i][2].map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < 3 } })), { x: x + 0.2, y: 2.62, w: 3.55, h: 1.75, fontSize: 16, paraSpaceAfter: 6 });
  }
  // checkpoints row
  T(s, "Suggested CI coaching checkpoints", { x: 0.6, y: 4.75, w: 4.5, h: 0.3, fontSize: 14, bold: true, color: GREEN });
  for (let k = 0; k < 12; k++) {
    const x = 0.75 + k * (11.8 / 12) + 0.25;
    s.addShape(pres.shapes.OVAL, { x, y: 5.15, w: 0.24, h: 0.24, fill: { color: GREEN }, line: { type: "none" } });
  }
  T(s, "Weekly, per cell (suggested): 20–30 minutes at the line", { x: 6.2, y: 4.75, w: 6.5, h: 0.3, fontSize: 12, italic: true, color: GRAY, align: "right" });
  T(s, "Leadership reviews", { x: 0.6, y: 5.65, w: 4.5, h: 0.3, fontSize: 14, bold: true, color: NAVY });
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1 + 3.9 - 0.32;
    s.addShape(pres.shapes.DIAMOND, { x, y: 6.05, w: 0.32, h: 0.32, fill: { color: NAVY }, line: { type: "none" } });
    T(s, ["Day 30 review", "Day 60 review", "Day 90 review & recognition"][i], { x: x - 2.95, y: 6.05, w: 2.85, h: 0.32, fontSize: 13, color: NAVY, align: "right", valign: "middle" });
  }
  s.addNotes(`Here's how the first 90 days look.

Days 1 to 30 are launch. We confirm the cells and owners, select first projects, establish baselines and knock out some initial quick wins. The goal for month one is momentum: every cell has a project, a baseline and at least one visible win.

Days 31 to 60 are about building. We review early results, strengthen problem solving (that's where coaching gets deeper), expand what's working and share wins across shifts.

Days 61 to 90 are sustain and grow. We standardize the improvements so they stick, start developing larger opportunities, the Stage 2 projects, and recognize the teams that delivered. Then we plan the next cycle.

The green dots are suggested coaching checkpoints: I'm proposing a short weekly touchpoint with each cell, 20 to 30 minutes, at the line rather than in a conference room. That's a suggestion, and I'm open to adjusting the rhythm to what works for each shift.

The diamonds are leadership reviews at 30, 60 and 90 days, where we look at progress as a group, remove barriers and, at day 90, recognize teams.

This is a pilot rhythm. After 90 days we'll look honestly at what worked and adjust.

Question: does a weekly 20 to 30 minute touchpoint feel realistic for your shift? What time of the shift works best?`);

  // =============== 16. CALL TO ACTION ===============
  s = newSlide("SECTION_DARK", "Launch Plan", "CALL TO ACTION  ·  FIRST ASSIGNMENTS", "Every Shift. Every Team. Every Improvement Counts.");
  const acts = [
    ["FaUsers", "Confirm cell members"],
    ["FaLightbulb", "Identify one opportunity per cell"],
    ["FaHardHat", "Assign a supervisor project owner"],
    ["FaRulerCombined", "Establish baseline data"],
    ["FaCalendarCheck", "Schedule the first CI coaching session"],
  ];
  const aw = 2.3, ag = (12.1 - 5 * aw) / 4;
  for (let i = 0; i < 5; i++) {
    const x = 0.6 + i * (aw + ag);
    card(s, x, 2.15, aw, 2.45, { fill: NAVY, line: STEEL, lineW: 1, name: "Assignment " + (i + 1) });
    await iconCircle(s, acts[i][0], x + 0.25, 2.4, 0.7, GREEN);
    T(s, String(i + 1), { x: x + 1.55, y: 2.4, w: 0.55, h: 0.7, fontSize: 30, bold: true, color: STEEL, align: "right", valign: "middle" });
    T(s, acts[i][1], { x: x + 0.25, y: 3.3, w: aw - 0.45, h: 1.1, fontSize: 17, bold: true, color: WHITE });
  }
  T(s, "We are not launching another reporting exercise.", { x: 0.6, y: 4.95, w: 12.1, h: 0.55, fontSize: 24, color: PALE, align: "center" });
  T(s, "We are building leaders who improve the business.", { x: 0.6, y: 5.45, w: 12.1, h: 0.65, fontSize: 30, bold: true, color: WHITE, align: "center" });
  s.addNotes(`So let's bring it home. Every shift, every team, every improvement counts.

Here are the first five assignments, and none of them are big.

One: confirm your cell members: your Quality and Maintenance partners.

Two: identify one improvement opportunity per cell. Just one. Use the funnel.

Three: assign a supervisor project owner. For cross-shift projects, agree on the one shared owner.

Four: establish baseline data. Before we change anything, measure where we are.

Five: schedule your first CI coaching session with me. I'll come to you.

I'd like all five done by our first checkpoint, and I'll follow up with each of you individually this week to set up coaching.

I'll leave you with this. We are not launching another reporting exercise. Nobody needs another spreadsheet. We are building leaders who improve the business. You already solve problems every single shift; this program is about giving that work structure, support and the recognition it deserves.

Before we wrap: what questions do you have? And what would make this program fail? I genuinely want to know, because I'd rather fix it now than in month two.

Thank you. Let's go make some improvements.`);

  // =============== APPENDIX ===============
  pres.addSection({ title: "Appendix" });
  s = newSlide("SECTION_DARK", "Appendix", "APPENDIX", "Tools and templates for supervisors");
  const apx = ["A1  Sample project charter", "A2  Example changeover project", "A3  Example 5S project", "A4  Suggested supervisor KPIs", "A5  Project selection scoring matrix", "A6  Example financial savings calculation"];
  for (let i = 0; i < 6; i++) {
    const x = 0.6 + (i % 2) * 6.15, y = 2.2 + Math.floor(i / 2) * 1.05;
    card(s, x, y, 5.95, 0.85, { fill: NAVY, line: STEEL, lineW: 1 });
    T(s, apx[i], { x: x + 0.3, y, w: 5.4, h: 0.85, fontSize: 19, bold: true, color: WHITE, valign: "middle" });
  }
  T(s, "All numbers in the appendix are illustrative examples, not C.H. Guenther plant data.", { x: 0.6, y: 5.6, w: 12.1, h: 0.4, fontSize: 14, italic: true, color: PALE, align: "center" });
  s.addNotes(`The appendix has practical tools you can use straight away: a sample project charter, worked examples for a changeover project and a 5S project, suggested KPIs, the project scoring matrix and a worked financial calculation.

Every number in these examples is illustrative. None of it is our plant data. They're there to show the method, not to set targets.

I'll send the editable versions to everyone after today's meeting.`);

  const tblHead = (arr) => arr.map((t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 13, valign: "middle" } }));
  const tblOpts = (y, colW, rowH) => ({ x: 0.6, y, w: colW.reduce((a, b) => a + b, 0), colW, rowH, fontSize: 13, color: INK, valign: "middle", border: { type: "solid", pt: 0.75, color: "D6DEE8" }, margin: [3, 6, 3, 6], fill: { color: "FFFFFF" } });

  // A1 Charter
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A1", "Sample project charter");
  const charter = [
    ["Project title", "[Short, specific name, e.g. “Line __ product changeover”]", "Project owner", "[Supervisor name]"],
    ["Cell / shift", "[A / B / C, or cross-shift]", "Team members", "[Production, Quality, Maintenance, operators]"],
    ["Problem statement", "[What is happening, where, since when, and why it matters]", "CI coach", "[CI Manager]"],
    ["Baseline", "[Current measured value, how and when measured]", "Finance reviewer", "Randy (Finance)"],
    ["Target", "[Measurable goal and date]", "Food-safety check", "[Quality sign-off required? Y/N]"],
    ["Scope: in / out", "[What is included; what is not]", "Key dates", "[Start · baseline · review · WOr]"],
    ["Expected benefit type", "[Hard savings / cost avoidance / capacity / nonfinancial]", "Risks & support", "[Barriers CI should help remove]"],
  ];
  s.addTable(charter.map((r) => r.map((c, k) => ({ text: c, options: k % 2 === 0 ? { bold: true, color: NAVY, fill: { color: "EEF2F6" } } : { italic: c.startsWith("["), color: c.startsWith("[") ? "5B7FA6" : "1A2433" } }))), tblOpts(1.7, [2.0, 4.3, 1.9, 3.9], 0.66));
  T(s, "Template fields in [brackets] are filled in by the project owner. Keep the whole charter to one page.", { x: 0.6, y: 6.5, w: 12.1, h: 0.3, fontSize: 12, italic: true, color: GRAY });
  s.addNotes(`This is the sample project charter. It's one page on purpose. If you can't fit your project on one page, it's probably too big for a first project.

The fields that matter most are the problem statement, the baseline and the target. Get those three right and the rest follows.

Two fields I want to point out: the food-safety check, so Quality signs off if the change touches food contact, sanitation or specs; and the expected benefit type, so we're honest from day one about whether this is hard savings, cost avoidance, capacity or nonfinancial.`);

  // A2 Changeover example
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A2", "Example changeover project");
  illus(s, 9.4, 0.32, 3.3);
  const coRows = [
    tblHead(["Step", "What the team did (example)", "Example result"]),
    ["Baseline", "Timed 5 changeovers on one line, last good to first good product", "Average 40 min"],
    ["Observe", "Filmed one changeover with crew consent; mapped every task", "Found 8 min of waiting"],
    ["Separate", "Split tasks into internal (line stopped) and external (line running)", "6 tasks movable to external"],
    ["Improve", "Pre-staged parts cart, pre-set guides, Maintenance on standby", "Waiting and motion cut"],
    ["Standardize", "One-page changeover checklist, trained all three shifts", "Same method A/B/C"],
    ["Verify", "Timed the next 10 changeovers", "Average 32 min, held"],
  ];
  s.addTable(coRows.map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: k === 0 ? { bold: true, color: NAVY } : k === 2 ? { bold: true, color: "2F9E6E" } : {} }))), tblOpts(1.7, [1.8, 6.8, 3.5], 0.58));
  card(s, 0.6, 5.95, 12.1, 0.75, { fill: LIGHT });
  T(s, [{ text: "Reported as: ", options: { bold: true, color: NAVY } }, { text: "8 min of capacity per changeover. Becomes hard savings only if Finance confirms paid hours came out (e.g. reduced overtime)." }], { x: 0.85, y: 5.95, w: 11.7, h: 0.75, fontSize: 14, valign: "middle" });
  s.addNotes(`This is a worked example of a changeover project, illustrative only. These are not our numbers.

The pattern is what matters: measure a few changeovers to get a real baseline; watch one closely with the crew; separate internal from external tasks; move what you can to external; standardize across all three shifts; then verify over a run of changeovers, not just one.

And notice the bottom line: the result is reported as capacity unless Finance confirms the hours actually came out.`);

  // A3 5S example
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A3", "Example 5S project: changeover tool station");
  illus(s, 9.4, 0.32, 3.3);
  const s5 = [
    ["Sort", "Red-tagged unused tools and obsolete change parts with the crew", "FaTrashAlt"],
    ["Set in Order", "Shadow board at point of use; labeled bins for change parts", "FaThLarge"],
    ["Shine", "Cleaning and inspection added to sanitation checklist (with Quality)", "FaBroom"],
    ["Standardize", "Photo standard posted; plant color code for food-contact tools", "FaTags"],
    ["Sustain", "Weekly 5-minute supervisor audit; score posted at the station", "FaLock"],
  ];
  for (let i = 0; i < 5; i++) {
    const y = 1.7 + i * 0.86;
    card(s, 0.6, y, 7.6, 0.74, { fill: LIGHT });
    await iconCircle(s, s5[i][2], 0.72, y + 0.1, 0.54, i === 4 ? GREEN : NAVY);
    T(s, s5[i][0], { x: 1.45, y, w: 1.75, h: 0.74, fontSize: 15, bold: true, color: NAVY, valign: "middle" });
    T(s, s5[i][1], { x: 3.2, y, w: 4.85, h: 0.74, fontSize: 13, valign: "middle" });
  }
  card(s, 8.5, 1.7, 4.2, 4.18, { fill: NAVY });
  T(s, "Example measures", { x: 8.8, y: 1.9, w: 3.6, h: 0.4, fontSize: 18, bold: true, color: WHITE });
  const m5 = [["Tool search time", "e.g. 4 min → under 1 min"], ["5S audit score", "e.g. 2.1 → 4.0 of 5"], ["Missing-tool events", "e.g. tracked weekly"]];
  m5.forEach((m, k) => {
    T(s, m[0], { x: 8.8, y: 2.5 + k * 1.05, w: 3.6, h: 0.35, fontSize: 15, bold: true, color: GREEN });
    T(s, m[1], { x: 8.8, y: 2.85 + k * 1.05, w: 3.6, h: 0.35, fontSize: 14, color: PALE });
  });
  T(s, "Benefit type: mostly capacity and nonfinancial (safety, food safety, morale).", { x: 0.6, y: 6.15, w: 12.1, h: 0.4, fontSize: 14, italic: true, color: GRAY });
  s.addNotes(`Here's an example 5S project, again illustrative. The scope is deliberately small: one changeover tool station.

Walk through it with the crew: sort what isn't needed, give everything a home at the point of use, link cleaning and inspection to the sanitation checklist with Quality's input, post a photo standard using the plant's existing color code, and sustain it with a quick weekly audit.

The example numbers on the right just show the kind of thing to measure. Most 5S benefit is capacity and nonfinancial, and that's fine. Report it honestly.`);

  // A4 KPIs
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A4", "Suggested supervisor KPIs");
  const kRows = [
    tblHead(["KPI", "Definition", "Source", "Suggested review"]),
    ["Changeover time", "Minutes from last good to first good product", "Line logs / timed observation", "Per changeover; weekly trend"],
    ["Unplanned downtime", "Minutes of unplanned stops, by top cause", "Downtime tracking", "Daily; weekly Pareto"],
    ["Scrap & rework", "Product lost or reworked (lb, units or cases)", "Production / Quality records", "Daily; weekly trend"],
    ["5S audit score", "Score against the area standard", "Supervisor / cell audit", "Weekly"],
    ["Employee participation", "Ideas raised; employees involved in actions", "Cell idea log", "Weekly"],
    ["Action completion", "% of project actions closed on time", "Project action list", "Weekly at coaching"],
    ["Validated value", "Finance-reviewed benefit by category", "Finance (Randy)", "At project close"],
  ];
  s.addTable(kRows.map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: k === 0 ? { bold: true, color: NAVY } : {} }))), tblOpts(1.7, [2.4, 4.4, 2.9, 2.4], 0.56));
  T(s, "Suggested starting set: use the one or two KPIs that fit your project, and use existing plant data sources where they exist.", { x: 0.6, y: 6.35, w: 12.1, h: 0.35, fontSize: 12, italic: true, color: GRAY });
  s.addNotes(`These are suggested KPIs for supervisors. You don't need all of them. Pick the one or two that match your project.

Wherever possible, use data sources we already have, so we aren't creating extra reporting. And notice the last row: validated value comes from Finance at project close, not from us.`);

  // A5 Scoring matrix
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A5", "Project selection scoring matrix");
  illus(s, 9.4, 0.32, 3.3, "ILLUSTRATIVE SCORES — EXAMPLE IDEAS");
  const crit = [["Safe (gate: must pass)", "Pass/Fail", "Pass", "Pass", "Pass"], ["Small enough to complete", "x3", "5", "3", "1"], ["Measurable", "x2", "5", "4", "3"], ["Within team’s influence", "x3", "5", "4", "2"], ["Value to operations", "x2", "3", "4", "5"], ["Resources available", "x1", "4", "3", "2"]];
  const score = (col) => crit.slice(1).reduce((a, r) => a + Number(r[col]) * Number(r[1].slice(1)), 0);
  const sRows = [tblHead(["Criterion (score 1–5)", "Weight", "Idea A: stage change parts", "Idea B: fix recurring jam", "Idea C: redesign pack area"])]
    .concat(crit.map((r) => r.map((c, k) => ({ text: c, options: k === 0 ? { bold: true, color: NAVY } : { align: "center" } }))))
    .concat([[{ text: "Weighted total (max 55)", options: { bold: true, color: WHITE, fill: { color: "13315C" } } }, { text: "", options: { fill: { color: "13315C" } } }].concat([2, 3, 4].map((k) => ({ text: String(score(k)), options: { bold: true, align: "center", color: WHITE, fill: { color: k === 2 ? "2F9E6E" : "13315C" } } })))]);
  s.addTable(sRows, tblOpts(1.7, [3.3, 1.3, 2.5, 2.5, 2.5], 0.52));
  T(s, [{ text: "How to use: ", options: { bold: true, color: NAVY } }, { text: "any idea that fails the safety gate is out. Score the rest 1–5 as a team, multiply by the weight, and start with the highest total. In this example, Idea A becomes the first project and Idea C is parked for Stage 3." }], { x: 0.6, y: 6.0, w: 12.1, h: 0.75, fontSize: 14 });
  s.addNotes(`If you want something more structured than the funnel, use this scoring matrix. Safety is a gate: if an idea fails it, it's out.

Then score each idea 1 to 5 as a team and multiply by the weight. "Small enough" and "within our influence" are weighted highest on purpose. For first projects, finishing matters more than size.

The three ideas and scores here are illustrative. Notice that the redesign idea scores highest on value but lowest overall. That's not a bad idea; it's a Stage 3 idea.`);

  // A6 Financial calculation
  s = newSlide("CONTENT_LIGHT", "Appendix", "APPENDIX A6", "Example financial savings calculation");
  illus(s, 9.4, 0.32, 3.3, "ILLUSTRATIVE — ALL INPUTS HYPOTHETICAL");
  const fin = [
    tblHead(["Step", "Calculation (example inputs)", "Result"]),
    ["Time saved per changeover", "40 min baseline − 32 min verified", "8 min"],
    ["Annual changeovers (assumed)", "5 per week × 50 weeks", "250"],
    ["Line time freed", "8 min × 250 ÷ 60", "≈ 33 line hours"],
    ["Crew labor hours (assumed crew of 6)", "33.3 h × 6 people", "≈ 200 labor hours"],
    ["Potential value (assumed $30/h loaded rate)", "200 h × $30", "≈ $6,000 potential"],
  ];
  s.addTable(fin.map((r, i) => i === 0 ? r : r.map((c, k) => ({ text: c, options: k === 0 ? { bold: true, color: NAVY } : k === 2 ? { bold: true, align: "right" } : {} }))), tblOpts(1.7, [4.3, 5.0, 2.8], 0.55));
  const outY = 5.15;
  card(s, 0.6, outY, 5.9, 1.55, { fill: LIGHT, line: STEEL, lineW: 1 });
  T(s, "If paid hours do NOT change", { x: 0.85, y: outY + 0.12, w: 5.4, h: 0.35, fontSize: 15, bold: true, color: STEEL });
  T(s, "Report ≈ 33 line hours of capacity. Hard savings = $0 until the capacity is used, for example to absorb volume without overtime.", { x: 0.85, y: outY + 0.5, w: 5.4, h: 0.95, fontSize: 14 });
  card(s, 6.8, outY, 5.9, 1.55, { fill: NAVY });
  T(s, "If Finance confirms hours came out", { x: 7.05, y: outY + 0.12, w: 5.4, h: 0.35, fontSize: 15, bold: true, color: GREEN });
  T(s, "e.g. overtime reduced by a matching amount: the confirmed amount is reported as validated hard savings.", { x: 7.05, y: outY + 0.5, w: 5.4, h: 0.95, fontSize: 14, color: WHITE });
  s.addNotes(`Last one: a worked financial calculation, with every input hypothetical. I picked round numbers so the logic is easy to follow; they are not our rates or volumes.

Eight minutes saved per changeover, about 250 changeovers a year, gives about 33 hours of line time. With a crew of six that's about 200 labor hours, and at an assumed rate that's roughly six thousand dollars of potential value.

Here's the key point, and why Randy's role matters: that six thousand dollars is potential. If nobody's paid hours change, what we've actually created is 33 hours of capacity, and the hard savings are zero until we use that capacity. If Finance confirms that overtime actually came down, then the confirmed amount becomes validated hard savings.

That's the discipline that keeps this program credible.`);

  await pres.writeFile({ fileName: FILE });
  await applyTheme(FILE, THEME);
  console.log("wrote", FILE);
})().catch((e) => { console.error(e); process.exit(1); });
