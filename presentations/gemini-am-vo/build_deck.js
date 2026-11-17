// Gemini Pilot — From Supervisors to Equipment Owners
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const MODE = (process.env.MODE || "AM").toUpperCase();
const AM = MODE === "AM", VO = MODE === "VO";
const OUT = process.argv[2] || (AM ? "Gemini_AM_Supervisor_Program.pptx" : "Gemini_VO_Supervisor_Program.pptx");
const PILLAR = AM ? "Autonomous Maintenance (AM)" : "Value Optimization (VO)";

const THEME = {
  name: "Gemini Industrial",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1A2433", lt1: "FFFFFF", dk2: "0B1F3A", lt2: "EEF3F7",
    accent1: "0E9C9C", accent2: "E9A23B", accent3: "D9534F",
    accent4: "2E9E5B", accent5: "3D7CC9", accent6: "5D6B7E",
    hlink: "0E9C9C", folHlink: "3D7CC9",
  },
};
const HEX = { navy: "0B1F3A", teal: "0E9C9C", amber: "E9A23B", coral: "D9534F", green: "2E9E5B", blue: "3D7CC9", slate: "5D6B7E", ink: "1A2433", light: "EEF3F7", white: "FFFFFF", grid: "D5DDE5" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "From Supervisors to Equipment Owners — Gemini Pilot — " + PILLAR;
pres.subject = PILLAR + " deployment";
pres.author = "Continuous Improvement — Duncanville";
pres.company = "C.H. Guenther &amp; Son";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const K = { navy: C.text2, ink: C.text1, white: C.background1, light: C.background2, teal: C.accent1, amber: C.accent2, coral: C.accent3, green: C.accent4, blue: C.accent5, slate: C.accent6 };
const S = pres.shapes;

// ---------- icons ----------
const iconCache = {};
async function icon(name, hex = "FFFFFF") {
  const key = name + hex;
  if (!iconCache[key]) {
    const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + hex, size: 256 }));
    const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
    iconCache[key] = "image/png;base64," + buf.toString("base64");
  }
  return iconCache[key];
}
const ICONS = ["FaHardHat","FaBrain","FaCogs","FaShieldAlt","FaSearch","FaTools","FaChartLine","FaUsers","FaClipboardCheck","FaClipboardList","FaIndustry","FaPlayCircle","FaLock","FaRedo","FaFish","FaMicroscope","FaCog","FaQuestionCircle","FaExclamationTriangle","FaBoxOpen","FaSlidersH","FaCheckCircle","FaTimesCircle","FaChalkboardTeacher","FaTag","FaBullseye","FaComments","FaLightbulb","FaEye","FaHandPaper","FaDollarSign","FaCalendarAlt","FaUserTie","FaTrophy","FaCamera","FaBolt","FaWrench","FaOilCan","FaBroom","FaTachometerAlt","FaSyncAlt","FaUserFriends","FaLayerGroup","FaFlag","FaMapMarkedAlt","FaDatabase","FaBalanceScale","FaStar","FaChartBar","FaHandshake","FaRegSquare","FaGraduationCap","FaHeadphones"];
const IC = {};

// ---------- helpers ----------
const shadow = () => ({ type: "outer", color: "0B1F3A", opacity: 0.13, blur: 6, offset: 2, angle: 90 });
let objN = 0;
function txt(s, text, o = {}) {
  s.addText(text, Object.assign({ isTextBox: true, fontSize: 14, color: K.ink, valign: "top", margin: 0.04, objectName: "Text " + ++objN }, o));
}
function box(s, x, y, w, h, o = {}) {
  const opt = { x, y, w, h, fill: { color: o.fill || K.white, transparency: o.tr || 0 }, objectName: o.name || "Card " + ++objN };
  if (o.rect) {} else opt.rectRadius = o.r === undefined ? 0.08 : o.r;
  if (o.line) opt.line = { color: o.line, width: o.lw || 1, dashType: o.dash || "solid" };
  else opt.line = { color: o.fill || K.white, width: 0, transparency: 100 };
  if (o.shadow) opt.shadow = shadow();
  s.addShape(o.rect ? S.RECTANGLE : S.ROUNDED_RECTANGLE, opt);
}
function circ(s, x, y, d, fill, o = {}) {
  const opt = { x, y, w: d, h: d, fill: { color: fill, transparency: o.tr || 0 }, objectName: "Circle " + ++objN };
  opt.line = o.line ? { color: o.line, width: o.lw || 1.5, dashType: o.dash || "solid" } : { color: fill, width: 0, transparency: 100 };
  s.addShape(S.OVAL, opt);
}
function iconCircle(s, name, x, y, d, fill, iconHex = "FFFFFF") {
  circ(s, x, y, d, fill);
  const p = d * 0.24;
  s.addImage({ data: IC[name + iconHex], x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: name.replace("Fa", "") + " icon", objectName: "Icon " + ++objN });
}
function iconOnly(s, name, x, y, d, iconHex) {
  s.addImage({ data: IC[name + iconHex], x, y, w: d, h: d, altText: name.replace("Fa", "") + " icon", objectName: "Icon " + ++objN });
}
function numCircle(s, n, x, y, d, fill, fs = 14, color = K.white) {
  circ(s, x, y, d, fill);
  txt(s, String(n), { x, y, w: d, h: d, fontSize: fs, bold: true, color, align: "center", valign: "middle", margin: 0 });
}
function badge(s, label, x, y, w = 2.3, fill = K.amber, color = K.navy) {
  box(s, x, y, w, 0.3, { fill, r: 0.15 });
  txt(s, label, { x, y, w, h: 0.3, fontSize: 10, bold: true, color, align: "center", valign: "middle", margin: 0, charSpacing: 1 });
}
function arrow(s, x1, y1, x2, y2, color = K.slate, w = 1.75) {
  const o = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001), line: { color, width: w, endArrowType: "triangle" }, objectName: "Arrow " + ++objN };
  if (x2 < x1) o.flipH = true;
  if (y2 < y1) o.flipV = true;
  s.addShape(S.LINE, o);
}
function line(s, x1, y1, x2, y2, color = K.slate, w = 1.5, dash = "solid") {
  const o = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001), line: { color, width: w, dashType: dash }, objectName: "Line " + ++objN };
  if ((x2 < x1) !== (y2 < y1)) o.flipV = true;
  s.addShape(S.LINE, o);
}
function bullets(items, o = {}) {
  return items.map((t, i) => {
    const base = typeof t === "string" ? { text: t, options: {} } : t;
    return { text: base.text, options: Object.assign({ bullet: o.bullet === false ? false : (o.num ? { type: "number" } : { indent: 14 }), breakLine: i < items.length - 1, paraSpaceAfter: o.space || 4 }, base.options) };
  });
}
// slide factories
let curSection = "";
function section(t) { if (t === curSection) return; curSection = t; pres.addSection({ title: t }); }
function content(kicker, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: curSection });
  s.addText(kicker.replace(/^PART \d+ · /, ""), { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function dark(kicker, title) {
  const s = pres.addSlide({ masterName: "EXERCISE", sectionTitle: curSection });
  s.addText(kicker.replace(/^PART \d+ · /, ""), { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function notes(s, n) { s.addNotes(n.trim().replace(/^ +/gm, "")); }

// ---------- layouts ----------
const FOOT = "C.H. Guenther & Son  |  Duncanville, Texas  |  Gemini Pilot — " + PILLAR;
const hdr = (t) => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, fontSize: 13 } });
pres.defineSlideMaster({
  title: "TITLE", background: { color: K.navy },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 1.6, w: 7.6, h: 2.1, fontSize: 42, bold: true, color: K.white, valign: "top", margin: 0, align: "left" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CLOSING", background: { color: K.navy },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 0.8, w: 11.9, h: 1.5, fontSize: 36, bold: true, color: K.white, valign: "middle", align: "center", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: K.white },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.28, w: 11.0, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.56, w: 12.1, h: 0.72, fontSize: 26, bold: true, color: K.navy, margin: 0, valign: "top", align: "left" }, text: "" } },
    { text: { text: FOOT, options: { x: 0.6, y: 7.08, w: 10, h: 0.26, fontSize: 9, color: K.slate, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.08, w: 0.5, h: 0.26, fontSize: 9, color: "5D6B7E", align: "right", margin: 0 },
});
pres.defineSlideMaster({
  title: "EXERCISE", background: { color: K.navy },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.28, w: 11.0, h: 0.3, fontSize: 11, bold: true, color: K.amber, charSpacing: 2, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.56, w: 12.1, h: 0.72, fontSize: 26, bold: true, color: K.white, margin: 0, valign: "top", align: "left" }, text: "" } },
    { text: { text: FOOT, options: { x: 0.6, y: 7.08, w: 10, h: 0.26, fontSize: 9, color: K.light, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.08, w: 0.5, h: 0.26, fontSize: 9, color: "EEF3F7", align: "right", margin: 0 },
});

async function main() {
  let s;
  for (const n of ICONS) for (const h of [HEX.white, HEX.navy, HEX.teal, HEX.coral, HEX.amber, HEX.blue, HEX.green, HEX.slate]) IC[n + h] = await icon(n, h);

  if (true) {
  // ======================= 1. TITLE =======================
  section("Kickoff");
  s = pres.addSlide({ masterName: "TITLE", sectionTitle: curSection });
  txt(s, "GEMINI PILOT  ·  " + (AM ? "AUTONOMOUS MAINTENANCE (AM) PILLAR" : "VALUE OPTIMIZATION (VO) PILLAR"), { x: 0.7, y: 1.05, w: 7.6, h: 0.35, fontSize: 13, bold: true, color: K.amber, charSpacing: 3, margin: 0 });
  s.addText("FROM SUPERVISORS TO EQUIPMENT OWNERS", { placeholder: "title" });
  txt(s, AM ? "Building High-Performance Teams Through Autonomous Maintenance and Continuous Improvement" : "Building High-Performance Teams Through Value Optimization and Continuous Improvement", { x: 0.7, y: 3.85, w: 7.4, h: 1.0, fontSize: 18, color: K.light, margin: 0 });
  txt(s, [
    { text: "C.H. Guenther & Son — Duncanville, Texas", options: { bold: true, color: K.white, breakLine: true } },
    { text: "Pilot area: Gemini Equipment — Three Machines", options: { color: K.teal } },
  ], { x: 0.7, y: 5.35, w: 7.4, h: 0.8, fontSize: 15, margin: 0 });
  for (let i = 0; i < 3; i++) {
    const y = 1.2 + i * 1.3;
    box(s, 8.9, y, 3.75, 1.1, { fill: K.white, tr: 90 });
    iconCircle(s, "FaIndustry", 9.1, y + 0.2, 0.7, K.teal);
    txt(s, "GEMINI " + (i + 1), { x: 10.0, y: y + 0.17, w: 2.5, h: 0.4, fontSize: 20, bold: true, color: K.white, margin: 0 });
    txt(s, "Cell " + (i + 1) + (AM ? "  ·  AM ownership" : "  ·  VO ownership"), { x: 10.0, y: y + 0.58, w: 2.6, h: 0.35, fontSize: 13, color: K.light, margin: 0 });
  }
  box(s, 8.9, 5.3, 3.75, 0.95, { fill: K.navy, line: K.light, dash: "dash", lw: 1 });
  txt(s, "Brand logo area — insert approved C.H. Guenther / Future Foods artwork", { x: 9.0, y: 5.35, w: 3.55, h: 0.85, fontSize: 11, color: K.light, align: "center", valign: "middle" });
  notes(s, `PURPOSE: Open the official launch of the Supervisor Productivity and Equipment Ownership Program for the three Gemini machines.
  TIMING: 2 minutes.
  SAY: "Today is not a training class on another tool. It is the start of how we will run these three machines from now on. Supervisors will become equipment owners — leading people, standards and measurable results."
  ${AM ? "SET THE FRAME: Three machines, three cross-functional cells, one common method: the four Autonomous Maintenance (AM) tools — Safety Map, CILT, Centerlining and Defect Handling. This deck covers the AM pillar only; Value Optimization (VO) is delivered in its companion deck." : "SET THE FRAME: Three machines, three cross-functional cells, one common method: the six Value Optimization (VO) steps — baseline, losses, prioritize, projects, validate savings, sustain. This deck covers the VO pillar only; Autonomous Maintenance (AM) is delivered in its companion deck."}
  FACILITATOR NOTE: Replace the logo placeholder with approved C.H. Guenther / Future Foods brand artwork before delivery. No brand artwork or Gemini photos were supplied when this deck was built; all machine visuals are illustrative until real photos are added.`);

  }
  if (true) {
  // ======================= 2. VISION =======================
  s = content("PART 1 · PROGRAM VISION", "Our vision: supervisors who lead people, equipment and results");
  box(s, 0.6, 1.5, 6.0, 5.3, { fill: K.navy });
  txt(s, "“We are not introducing another checklist.”", { x: 0.95, y: 1.8, w: 5.3, h: 1.1, fontSize: 26, bold: true, color: K.white, margin: 0, fontFace: "Arial" });
  txt(s, "We are building a culture where every team understands its equipment, recognizes abnormalities, and takes ownership of improvement.", { x: 0.95, y: 3.0, w: 5.3, h: 1.4, fontSize: 18, color: K.light, margin: 0 });
  const roles4 = ["Coach teams", "Understand equipment", "Prevent recurring losses", "Lead measurable improvement"];
  txt(s, "OUR SUPERVISORS WILL", { x: 0.95, y: 4.55, w: 5.3, h: 0.3, fontSize: 11, bold: true, color: K.amber, charSpacing: 2, margin: 0 });
  roles4.forEach((r, i) => {
    const x = 0.95 + (i % 2) * 2.7, y = 4.95 + Math.floor(i / 2) * 0.8;
    box(s, x, y, 2.55, 0.62, { fill: K.teal, r: 0.1 });
    txt(s, r, { x, y, w: 2.55, h: 0.62, fontSize: 13, bold: true, color: K.white, align: "center", valign: "middle" });
  });
  const vis = [
    ["FaUsers", "Engagement & ownership", "Teams own their machine, not just their shift"],
    ["FaHardHat", "Safety", "Every standard starts with safe access and energy control"],
    ["FaShieldAlt", "Food quality", "Process conditions that protect the product"],
    ["FaCogs", "Equipment reliability", "Basic conditions restored and maintained"],
    ["FaClipboardCheck", "Consistent standards", "The same right way, on every shift"],
    ["FaSearch", "Root-cause problem-solving", "Fix the system, not the symptom"],
    AM ? ["FaChartLine", "Measurable results", "Fewer defects, stops and adjustments — tracked weekly"] : ["FaChartLine", "Productivity & savings", "Measured, Finance-validated, sustained"],
  ];
  vis.forEach(([ic, h, d], i) => {
    const y = 1.5 + i * 0.76;
    iconCircle(s, ic, 7.0, y + 0.05, 0.56, i < 3 ? K.teal : K.navy);
    txt(s, h, { x: 7.75, y: y, w: 4.95, h: 0.32, fontSize: 15, bold: true, color: K.navy, margin: 0 });
    txt(s, d, { x: 7.75, y: y + 0.32, w: 4.95, h: 0.32, fontSize: 13, color: K.slate, margin: 0 });
  });
  notes(s, `PURPOSE: State the vision and the core message.
  TIMING: 3 minutes.
  SAY: Read the core message aloud. "We are not introducing another checklist. We are building a culture where every team understands its equipment, recognizes abnormalities, and takes ownership of improvement."
  KEY POINT: Supervisors already manage production. This program adds four capabilities: coaching teams, understanding the equipment, preventing recurring losses, and leading measurable improvement.
  WALK THE RIGHT-HAND LIST: Safety and food quality come first and are never traded for output. Reliability and standards are the foundation. ${AM ? "Root-cause thinking and measurable results show whether it is working." : "Root-cause thinking, productivity and Finance-validated savings are how we measure whether it is working."}
  ASK: "What would be different on your shift if every operator truly owned their machine?" Capture 3–4 answers on a flip chart; refer back to them in the closing.`);

  }
  if (true) {
  // ======================= 3. PROGRAM MAP =======================
  s = content("PROGRAM ROADMAP", AM ? "The AM program at a glance: learn, own, integrate, lead" : "The VO program at a glance: learn, measure, improve, lead");
  const phases = AM ? [
    ["FaBrain", "LEARN", K.blue, ["How the brain learns & why problems return", "Root-cause thinking beyond 'operator error'"], "Foundation"],
    ["FaUsers", "OWN", K.teal, ["Three cross-functional Gemini cells", "AM tools: Safety Map, CILT, Centerlining, Defect Handling"], "AM tools"],
    ["FaHandshake", "INTEGRATE", K.amber, ["AM + Planned Maintenance + Quality", "Shared action & escalation process"], "PM & Quality"],
    ["FaFlag", "LEAD & SUSTAIN", K.navy, ["Roles & RACI", "90-day AM roadmap", "Daily huddles & coaching", "WOr reviews", "AM KPI dashboard"], "Leadership"],
  ] : [
    ["FaSearch", "LEARN", K.blue, ["Root-cause thinking beyond 'operator error'", "How VO differs from AM"], "Foundation"],
    ["FaDatabase", "MEASURE", K.teal, ["Step 1: Baseline (4–8 weeks)", "Step 2: Six Big Losses & Pareto"], "VO steps 1–2"],
    ["FaChartLine", "IMPROVE", K.amber, ["Step 3: Prioritize", "Step 4: Projects", "Step 5: Validate savings with Finance"], "VO steps 3–5"],
    ["FaFlag", "LEAD & SUSTAIN", K.navy, ["Step 6: Sustain & expand", "Roles & RACI", "90-day VO roadmap", "WOr reviews", "VO KPI dashboard"], "Leadership"],
  ];
  phases.forEach(([ic, name, col, items, part], i) => {
    const x = 0.6 + i * 3.1;
    box(s, x, 1.5, 2.85, 4.25, { fill: K.light });
    iconCircle(s, ic, x + 0.25, 1.75, 0.75, col);
    txt(s, part.toUpperCase(), { x: x + 1.15, y: 1.8, w: 1.6, h: 0.28, fontSize: 10, bold: true, color: K.slate, margin: 0, charSpacing: 1 });
    txt(s, name, { x: x + 1.15, y: 2.08, w: 1.65, h: 0.4, fontSize: 18, bold: true, color: K.navy, margin: 0 });
    txt(s, bullets(items, { space: 6 }), { x: x + 0.2, y: 2.75, w: 2.5, h: 2.9, fontSize: 13, color: K.ink });
    if (i < 3) arrow(s, x + 2.88, 3.6, x + 3.08, 3.6, K.slate, 2);
  });
  box(s, 0.6, 6.0, 12.1, 0.8, { fill: K.teal, tr: 85 });
  txt(s, AM ? [
    { text: "This deck: ", options: { bold: true, color: K.navy } },
    { text: "Autonomous Maintenance — establish and keep basic equipment conditions.  " },
    { text: "Companion deck: ", options: { bold: true, color: K.navy } },
    { text: "Value Optimization — find, size, prioritize and validate losses." },
  ] : [
    { text: "This deck: ", options: { bold: true, color: K.navy } },
    { text: "Value Optimization — find, size, prioritize and validate losses.  " },
    { text: "Companion deck: ", options: { bold: true, color: K.navy } },
    { text: "Autonomous Maintenance — establish and keep basic conditions." },
  ], { x: 0.85, y: 6.0, w: 11.7, h: 0.8, fontSize: 14, valign: "middle", color: K.ink });
  notes(s, AM ? `PURPOSE: Give the audience the map of the AM deck before diving in.
  TIMING: 2 minutes.
  SAY: "We start with how people learn and why problems come back, and how to find real root causes. Then we set up ownership — three cells and four AM tools. We integrate AM with Planned Maintenance and Quality. Finally we lead and sustain it with clear roles, a 90-day plan, daily huddles, leadership reviews and a dashboard."
  SUGGESTED SESSION PLAN (adjust to your schedule):
  • Leadership kickoff (60 min): opening, three cells, what AM is, roles, roadmap, dashboard and closing.
  • Supervisor practical training, Day 1 (4 hrs): neuroscience, root cause, Safety Map and CILT.
  • Supervisor practical training, Day 2 (3 hrs): Centerlining, Defect Handling, PM & Quality integration, daily huddles and coaching.
  COMPANION DECK: Value Optimization (VO) — baseline, losses, prioritization, projects and Finance-validated savings — is delivered separately. AM and VO are different pillars with different jobs.` : `PURPOSE: Give the audience the map of the VO deck before diving in.
  TIMING: 2 minutes.
  SAY: "We start with root-cause thinking and how VO differs from AM. Then we measure: a baseline for each Gemini and a clear view of the biggest losses. We improve: prioritize, run projects, and validate savings with Finance. Finally we sustain and lead it with clear roles, a 90-day plan, leadership reviews and a dashboard."
  SUGGESTED SESSION PLAN (adjust to your schedule):
  • Leadership kickoff (45 min): opening, cells, AM vs VO, the six VO steps, roles, roadmap, dashboard and closing.
  • Supervisor practical training (3 hrs): root cause, baseline and losses, prioritization, projects, the savings-validation exercise, WOr template.
  COMPANION DECK: Autonomous Maintenance (AM) — Safety Map, CILT, Centerlining and Defect Handling — is delivered separately. AM and VO are different pillars with different jobs.`);

  }
  if (AM) {
  // ======================= 4. OPENER =======================
  section("Neuroscience & Leadership");
  s = content("PART 2 · NEUROSCIENCE & LEADERSHIP — DISCUSSION OPENER", "Why does the problem come back two weeks later?");
  txt(s, "“", { x: 0.6, y: 1.3, w: 1, h: 1, fontSize: 80, bold: true, color: K.teal, margin: 0 });
  txt(s, "We train someone today, see improvement tomorrow — and watch the same problem return two weeks later.", { x: 0.6, y: 2.1, w: 5.6, h: 1.4, fontSize: 20, bold: true, color: K.navy, margin: 0 });
  txt(s, "DISCUSS IN PAIRS", { x: 0.6, y: 3.75, w: 5.6, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  txt(s, bullets(["Think of a time training “didn’t stick.” What happened?", "What was different on the day it worked?", "What reminded people — or failed to remind them — afterward?"], { num: true, space: 8 }), { x: 0.6, y: 4.1, w: 5.6, h: 1.9, fontSize: 15 });
  badge(s, "10 MINUTES", 0.6, 6.25, 1.6, K.navy, K.white);
  const days = ["Day 0", "1", "2", "3", "5", "7", "10", "14"];
  s.addChart(pres.charts.LINE, [
    { name: "Training only", labels: days, values: [100, 68, 55, 46, 38, 33, 30, 28] },
    { name: "Training + reinforcement", labels: days, values: [100, 80, 92, 86, 93, 90, 95, 94] },
  ], {
    x: 6.6, y: 1.9, w: 6.1, h: 4.3, chartColors: [HEX.coral, HEX.teal], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 7,
    showTitle: true, title: "Conceptual retention curve", titleFontSize: 14, titleColor: HEX.navy, titleFontFace: "+mn-lt",
    showLegend: true, legendPos: "b", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: HEX.ink,
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 25, showValAxisTitle: true, valAxisTitle: "Relative recall (concept)", valAxisTitleFontSize: 10, valAxisTitleColor: HEX.slate,
    catAxisLabelColor: HEX.slate, valAxisLabelColor: HEX.slate, catAxisLabelFontSize: 10, valAxisLabelFontSize: 10, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: HEX.grid, size: 0.5 }, catGridLine: { style: "none" },
  });
  badge(s, "CONCEPTUAL — NOT MEASURED DATA", 6.6, 1.5, 3.3);
  txt(s, "Shape inspired by classic forgetting-curve research (Ebbinghaus). Real rates vary by person, task and context.", { x: 6.6, y: 6.25, w: 6.1, h: 0.5, fontSize: 10, italic: true, color: K.slate, margin: 0 });
  notes(s, `PURPOSE: Engage the room with a problem every supervisor has lived.
  TIMING: 10 minutes (2 min framing, 5 min pairs, 3 min report-out).
  FACILITATOR INSTRUCTIONS: Ask the question on the slide. Give pairs 5 minutes for the three prompts. Take 3–4 report-outs. Write recurring themes on a flip chart (typical answers: "no follow-up", "new person on shift", "the standard wasn't posted", "we were busy", "the machine was different").
  TEACHING POINT: Training alone does not guarantee sustained behavior. Without retrieval, practice and reinforcement, newly learned information fades quickly.
  ABOUT THE CHART: It is conceptual, drawn to show the general shape described in forgetting-curve research dating back to Hermann Ebbinghaus (1885). It is not measured plant data. Do not quote percentages from it.
  TRANSITION: "To fix this, we need to understand a little about how the brain learns."`);

  }
  if (AM) {
  // ======================= 5. HOW BRAIN LEARNS =======================
  s = content("PART 2 · HOW THE BRAIN LEARNS", "Learning is a process — not a single event");
  const flow = [
    ["FaEye", "1  Attention", "We only learn what we notice. Noise, fatigue and distraction limit what gets in."],
    ["FaLayerGroup", "2  Encoding", "New information is held briefly and linked to what we already know."],
    ["FaSyncAlt", "3  Consolidation", "Over hours to days — including sleep — memories are stabilized and reorganized."],
    ["FaRedo", "4  Retrieval & use", "Each time we recall and apply it, the memory becomes stronger and easier to access."],
  ];
  flow.forEach(([ic, h, d], i) => {
    const x = 0.6 + i * 3.1;
    box(s, x, 1.5, 2.85, 1.9, { fill: K.light });
    iconCircle(s, ic, x + 0.2, 1.65, 0.55, K.teal);
    txt(s, h, { x: x + 0.85, y: 1.7, w: 1.95, h: 0.45, fontSize: 15, bold: true, color: K.navy, margin: 0, valign: "middle" });
    txt(s, d, { x: x + 0.2, y: 2.3, w: 2.5, h: 1.05, fontSize: 12, color: K.ink, margin: 0 });
    if (i < 3) arrow(s, x + 2.88, 2.45, x + 3.08, 2.45, K.slate, 2);
  });
  box(s, 0.6, 3.65, 5.9, 2.3, { fill: K.white, line: K.coral, lw: 1.5 });
  iconCircle(s, "FaBolt", 0.85, 3.85, 0.55, K.coral);
  txt(s, "Working memory", { x: 1.55, y: 3.9, w: 4.8, h: 0.45, fontSize: 17, bold: true, color: K.navy, margin: 0 });
  txt(s, bullets(["Holds only a small amount — a few items — for seconds", "Easily displaced by interruptions and stress", "Where a new standard sits right after a training session"]), { x: 0.85, y: 4.5, w: 5.5, h: 1.4, fontSize: 13 });
  box(s, 6.8, 3.65, 5.9, 2.3, { fill: K.white, line: K.teal, lw: 1.5 });
  iconCircle(s, "FaDatabase", 7.05, 3.85, 0.55, K.teal);
  txt(s, "Long-term memory", { x: 7.75, y: 3.9, w: 4.8, h: 0.45, fontSize: 17, bold: true, color: K.navy, margin: 0 });
  txt(s, bullets(["Very large capacity, built gradually", "Strengthened by repetition, meaning and spaced practice", "Where standards must live to survive shift changes and pressure"]), { x: 7.05, y: 4.5, w: 5.5, h: 1.4, fontSize: 13 });
  arrow(s, 6.52, 4.8, 6.78, 4.8, K.slate, 2.5);
  box(s, 0.6, 6.15, 12.1, 0.65, { fill: K.teal, tr: 85 });
  txt(s, [{ text: "Floor implication: ", options: { bold: true, color: K.navy } }, { text: "a 20-minute training fills working memory. Only practice, retrieval and reinforcement build long-term capability." }], { x: 0.85, y: 6.15, w: 11.7, h: 0.65, fontSize: 14, valign: "middle" });
  notes(s, `PURPOSE: Explain, accurately and simply, how learning works.
  TIMING: 6 minutes.
  SAY (Attention): "If an operator is distracted by noise, a jam, or a production push, the information may never be properly encoded."
  SAY (Encoding): "New information is linked to existing knowledge. That's why connecting a standard to something people already know — 'this is like checking tire pressure' — helps."
  SAY (Consolidation): "The brain stabilizes new memories over hours and days; sleep plays an important role. One session is not enough."
  SAY (Retrieval): "Every time someone recalls and uses a standard, it gets stronger. This is called the testing or retrieval-practice effect."
  WORKING vs LONG-TERM MEMORY: Working memory is limited (research commonly estimates about four meaningful 'chunks' at a time) and short-lived. Long-term memory is large but must be built through repetition and use.
  SCIENTIFIC ACCURACY: Keep claims at this level. Avoid "we only use 10% of our brain" or "left-brain/right-brain learners" — these are myths.
  ASK: "How many new things did we ask an operator to remember at the last training? How many could they recall a week later?"`);

  }
  if (AM) {
  // ======================= 6. MODES + PLASTICITY =======================
  s = content("PART 2 · ATTENTION, HABITS & NEUROPLASTICITY", "Same brain, two modes — and connections that grow with practice");
  txt(s, "FROM EFFORTFUL TO AUTOMATIC", { x: 0.6, y: 1.5, w: 6, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  const modes = [["New standard", "Needs full attention. Slow, error-prone, easily forgotten under pressure.", K.coral], ["Practiced", "Fewer errors. Still needs cues and occasional coaching.", K.amber], ["Automatic", "Runs reliably, even when busy — good habits and bad habits alike.", K.teal]];
  modes.forEach(([h, d, col], i) => {
    const x = 0.6 + i * 2.0;
    box(s, x, 1.9, 1.85, 2.2, { fill: col, tr: 85 });
    numCircle(s, i + 1, x + 0.15, 2.05, 0.45, col, 13);
    txt(s, h, { x: x + 0.15, y: 2.6, w: 1.6, h: 0.35, fontSize: 14, bold: true, color: K.navy, margin: 0 });
    txt(s, d, { x: x + 0.15, y: 2.97, w: 1.6, h: 1.1, fontSize: 11, color: K.ink, margin: 0 });
  });
  arrow(s, 0.7, 4.3, 6.4, 4.3, K.slate, 2);
  txt(s, "repetition + feedback + time", { x: 0.6, y: 4.35, w: 5.9, h: 0.3, fontSize: 11, italic: true, color: K.slate, align: "center", margin: 0 });
  box(s, 0.6, 4.85, 5.9, 1.95, { fill: K.light });
  iconOnly(s, "FaExclamationTriangle", 0.8, 5.0, 0.4, HEX.amber);
  txt(s, [{ text: "Scientifically accurate framing: ", options: { bold: true, color: K.navy } }, { text: "“conscious” and “subconscious” are not two separate physical brain structures. They describe different modes of processing across overlapping brain networks. Deliberate control relies heavily on prefrontal networks; well-practiced routines increasingly engage networks that include the basal ganglia." }], { x: 1.3, y: 4.95, w: 5.1, h: 1.8, fontSize: 12 });
  // neuroplasticity panel
  box(s, 6.9, 1.5, 5.8, 5.3, { fill: K.navy });
  txt(s, "NEUROPLASTICITY", { x: 7.2, y: 1.7, w: 5.2, h: 0.3, fontSize: 11, bold: true, color: K.amber, charSpacing: 2, margin: 0 });
  txt(s, "The brain changes its connections with experience.", { x: 7.2, y: 2.0, w: 5.2, h: 0.4, fontSize: 15, bold: true, color: K.white, margin: 0 });
  const rows = [["Before practice", 1.25, HEX.slate, "weak, inefficient connection"], ["After repeated practice", 6, HEX.teal, "stronger, more efficient connection"]];
  rows.forEach(([lab, wgt, col, d], i) => {
    const y = 2.75 + i * 1.45;
    txt(s, lab, { x: 7.2, y, w: 5.2, h: 0.3, fontSize: 12, bold: true, color: K.light, margin: 0 });
    circ(s, 7.3, y + 0.42, 0.55, col === HEX.teal ? K.teal : K.slate);
    line(s, 7.85, y + 0.7, 9.55, y + 0.7, col === HEX.teal ? K.teal : K.slate, wgt);
    circ(s, 9.5, y + 0.55, 0.3, col === HEX.teal ? K.teal : K.slate);
    line(s, 10.0, y + 0.7, 10.6, y + 0.7, K.light, 1.5);
    circ(s, 10.55, y + 0.42, 0.55, K.blue);
    txt(s, d, { x: 11.15, y: y + 0.4, w: 1.45, h: 0.65, fontSize: 10, color: K.light, margin: 0, valign: "middle" });
  });
  txt(s, "synapse", { x: 9.6, y: 3.6, w: 1.0, h: 0.25, fontSize: 10, italic: true, color: K.amber, align: "center", margin: 0 });
  txt(s, "Repeated activation can strengthen synapses (for example, long-term potentiation). Connections that are not used can weaken. Learning a standard is literally building and strengthening networks.", { x: 7.2, y: 5.65, w: 5.25, h: 1.05, fontSize: 12, color: K.white, margin: 0 });
  notes(s, `PURPOSE: Explain conscious attention vs automatic habits, and neuroplasticity, without myths.
  TIMING: 6 minutes.
  LEFT SIDE: When a standard is new, people must think about every step — this demands attention and breaks down under pressure (a jam, a rush order). With repetition and feedback, the behavior becomes more automatic and reliable.
  IMPORTANT ACCURACY POINT: Do not describe the conscious and subconscious mind as two separate physical parts of the brain. They are modes of processing supported by overlapping, interacting networks. Research on habit learning shows prefrontal networks are heavily involved in deliberate, goal-directed control, while practiced routines increasingly involve circuits that include the basal ganglia (especially the striatum).
  WARNING: Habits form whether we design them or not. If the 'workaround' is practiced every day, the workaround becomes automatic.
  RIGHT SIDE (diagram is schematic, not to scale): A synapse is the junction where one neuron signals another. With repeated, coordinated activity, synapses can strengthen — long-term potentiation (LTP) is one well-studied mechanism. Unused connections can weaken or be pruned.
  ASK: "What workaround on the Gemini has become automatic for some operators?"`);

  }
  if (AM) {
  // ======================= 7. FORGET / STICK =======================
  s = content("PART 2 · RETENTION", "Why people forget — and what makes standards stick");
  box(s, 0.6, 1.5, 4.6, 5.3, { fill: K.light });
  txt(s, "WHY WE FORGET", { x: 0.9, y: 1.7, w: 4, h: 0.3, fontSize: 11, bold: true, color: K.coral, charSpacing: 2, margin: 0 });
  const forget = [["No use", "Memories that are never retrieved fade"], ["Interference", "Old habits and other methods compete with the new one"], ["Overload", "Too much at once exceeds working memory"], ["No cue at point of use", "Nothing at the machine triggers the right action"]];
  forget.forEach(([h, d], i) => {
    const y = 2.15 + i * 1.12;
    iconCircle(s, "FaTimesCircle", 0.9, y, 0.5, K.coral);
    txt(s, h, { x: 1.55, y: y - 0.03, w: 3.5, h: 0.32, fontSize: 15, bold: true, color: K.navy, margin: 0 });
    txt(s, d, { x: 1.55, y: y + 0.3, w: 3.5, h: 0.6, fontSize: 12, color: K.ink, margin: 0 });
  });
  txt(s, "WHAT MAKES IT STICK  →  HOW WE USE IT IN THE GEMINI PILOT", { x: 5.6, y: 1.5, w: 7.1, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  const stick = [["FaEye", "Visual standards", "CILT cards, safety maps and centerline marks at the machine"], ["FaComments", "Coaching", "Supervisors ask questions at the machine, not just give instructions"], ["FaCheckCircle", "Feedback", "Defect tags closed and verified — people see their findings matter"], ["FaCalendarAlt", "Spaced repetition", "Key standards revisited in daily huddles, not once a year"], ["FaChalkboardTeacher", "Teach-back", "Operators explain standards to each other and to new hires"]];
  stick.forEach(([ic, h, d], i) => {
    const y = 1.95 + i * 0.97;
    box(s, 5.6, y, 7.1, 0.82, { fill: K.white, line: HEX.grid, lw: 0.75 });
    iconCircle(s, ic, 5.75, y + 0.13, 0.56, K.teal);
    txt(s, h, { x: 6.45, y, w: 2.1, h: 0.82, fontSize: 15, bold: true, color: K.navy, valign: "middle", margin: 0 });
    txt(s, d, { x: 8.6, y, w: 4.0, h: 0.82, fontSize: 12, color: K.ink, valign: "middle", margin: 0 });
  });
  notes(s, `PURPOSE: Connect the science to the tools we are deploying.
  TIMING: 5 minutes.
  WHY WE FORGET: Unused memories become harder to retrieve; competing habits interfere with new ones; overload at the point of learning limits what is encoded; and when there is no cue at the moment of action, people default to habit.
  WHAT MAKES IT STICK — the five reinforcers are evidence-informed learning principles:
  • Visual standards provide retrieval cues exactly where the work happens.
  • Coaching creates active processing (the person has to think, not just listen).
  • Feedback closes the loop — people repeat behaviors they see make a difference.
  • Spaced repetition (revisiting over days and weeks) beats a single long session.
  • Teach-back is a form of retrieval practice and reveals gaps immediately.
  KEY LINE: "Every AM tool in this program is also a learning tool."
  ASK: "Which of these five do we do well today? Which one is missing on the Gemini?"`);

  }
  if (AM) {
  // ======================= 8. VIDEOS =======================
  s = content("PART 2 · RECOMMENDED VIDEOS", "Two short videos to bring the science to life");
  const vids = [
    { t: "After watching this, your brain will not be the same", src: "Dr. Lara Boyd, University of British Columbia — TEDxVancouver (2015)", dur: "About 14 minutes", url: "https://www.youtube.com/watch?v=LNHBMFCzznE", what: "An engaging, research-based talk on neuroplasticity: how the brain changes its structure and function as we learn, why practice is required, and why people learn at different rates.", use: "Show before the Power of Repetition exercise. Ask: “What did she say is the main driver of brain change?” (Behavior — practice.)" },
    { t: "2-Minute Neuroscience: Long-Term Potentiation (LTP)", src: "Neuroscientifically Challenged — animated explainer", dur: "About 2 minutes", url: "https://www.youtube.com/watch?v=-mHgPfXHzJE", what: "A scientifically grounded animation of a synapse, showing how repeated activation can strengthen the connection between neurons (NMDA-receptor-dependent LTP).", use: "Show with the neuroplasticity slide. Emphasize that LTP is a leading candidate mechanism for learning and memory at the cellular level." },
  ];
  vids.forEach((v, i) => {
    const x = 0.6 + i * 6.15;
    box(s, x, 1.5, 5.95, 4.45, { fill: K.white, line: HEX.grid, lw: 1, shadow: true });
    iconCircle(s, "FaPlayCircle", x + 0.25, 1.72, 0.7, i === 0 ? K.coral : K.teal);
    txt(s, "VIDEO " + (i + 1) + "  ·  " + v.dur.toUpperCase(), { x: x + 1.1, y: 1.72, w: 4.6, h: 0.28, fontSize: 10, bold: true, color: K.slate, charSpacing: 1, margin: 0 });
    txt(s, v.t, { x: x + 1.1, y: 1.98, w: 4.7, h: 0.6, fontSize: 15, bold: true, color: K.navy, margin: 0 });
    txt(s, v.src, { x: x + 0.25, y: 2.65, w: 5.5, h: 0.3, fontSize: 12, italic: true, color: K.slate, margin: 0 });
    txt(s, [{ text: v.url, options: { hyperlink: { url: v.url, tooltip: v.t }, color: K.teal, underline: true } }], { x: x + 0.25, y: 2.97, w: 5.5, h: 0.3, fontSize: 12, margin: 0 });
    txt(s, [{ text: "What it demonstrates: ", options: { bold: true, color: K.navy } }, { text: v.what }], { x: x + 0.25, y: 3.4, w: 5.5, h: 1.3, fontSize: 12, margin: 0 });
    txt(s, [{ text: "How to use it: ", options: { bold: true, color: K.navy } }, { text: v.use }], { x: x + 0.25, y: 4.75, w: 5.5, h: 1.1, fontSize: 12, margin: 0 });
  });
  box(s, 0.6, 6.15, 12.1, 0.65, { fill: K.amber, tr: 80 });
  txt(s, [{ text: "Use with care: ", options: { bold: true, color: K.navy } }, { text: "videos of neurons and synapses illustrate biological mechanisms. They do not by themselves prove how a specific workplace habit forms — that takes practice, coaching and measurement on the floor." }], { x: 0.85, y: 6.15, w: 11.7, h: 0.65, fontSize: 13, valign: "middle" });
  notes(s, `PURPOSE: Provide two credible, short videos.
  TIMING: 16–18 minutes if both are shown in full; or show Video 2 (2 min) plus a 3–4 minute segment of Video 1.
  LINK VERIFICATION: Both titles and YouTube video IDs were confirmed against published video listings in November 2026. YouTube was not reachable from the build environment, so open each link once on the plant network before the session. Plant networks often block YouTube — download through an approved method or use a personal hotspot if permitted by IT policy.
  VIDEO 1 — "After watching this, your brain will not be the same" by Dr. Lara Boyd (UBC), TEDxVancouver 2015: https://www.youtube.com/watch?v=LNHBMFCzznE . Key messages: the brain is changed by every experience; practice is the primary driver of change; there is no one-size-fits-all way to learn.
  VIDEO 2 — "2-Minute Neuroscience: Long-Term Potentiation (LTP)" by Neuroscientifically Challenged: https://www.youtube.com/watch?v=-mHgPfXHzJE . Companion page: https://neuroscientificallychallenged.com/blog/2-minute-neuroscience-long-term-potentiation . Animated explanation of NMDA-receptor-dependent LTP — how repeated activation strengthens a synapse.
  CAUTION TO STATE OUT LOUD: These videos explain mechanisms. They do not prove how a specific workplace habit forms. Avoid saying "this video shows your operators' habits growing." Say instead: "This is one of the mechanisms scientists think supports learning — practice is how we put it to work."
  DISCUSSION QUESTION: "If practice changes the brain, what are our operators practicing every day on the Gemini right now?"`);

  }
  if (AM) {
  // ======================= 9. EXERCISE REPETITION =======================
  s = dark("INTERACTIVE EXERCISE · 15 MINUTES", "The Power of Repetition: hear it, do it, teach it");
  const rounds = [
    ["FaHeadphones", "ROUND 1", "Hear It", "3 min", "The supervisor explains one standard once, in words only. Each team member writes the steps from memory.", "Working memory only — expect gaps"],
    ["FaHandPaper", "ROUND 2", "Do It", "5 min", "The team demonstrates the process using the visual standard. The supervisor coaches with questions, not corrections.", "Action + visual cue = stronger encoding"],
    ["FaChalkboardTeacher", "ROUND 3", "Teach It", "5 min", "A different team member teaches the standard back to the group. The group checks it against the standard.", "Retrieval + explanation builds long-term memory"],
  ];
  rounds.forEach(([ic, r, h, t, d, b], i) => {
    const x = 0.6 + i * 4.1;
    box(s, x, 1.5, 3.85, 4.1, { fill: K.white });
    iconCircle(s, ic, x + 0.25, 1.75, 0.75, [K.coral, K.amber, K.teal][i], HEX.white);
    txt(s, r + "  ·  " + t, { x: x + 1.15, y: 1.78, w: 2.5, h: 0.3, fontSize: 11, bold: true, color: K.slate, charSpacing: 1, margin: 0 });
    txt(s, h, { x: x + 1.15, y: 2.08, w: 2.5, h: 0.45, fontSize: 22, bold: true, color: K.navy, margin: 0 });
    txt(s, d, { x: x + 0.25, y: 2.8, w: 3.35, h: 1.6, fontSize: 13, color: K.ink, margin: 0 });
    box(s, x + 0.25, 4.6, 3.35, 0.75, { fill: K.teal, tr: 85 });
    txt(s, b, { x: x + 0.35, y: 4.6, w: 3.15, h: 0.75, fontSize: 12, bold: true, color: K.navy, valign: "middle", margin: 0 });
    if (i < 2) arrow(s, x + 3.87, 3.55, x + 4.08, 3.55, K.amber, 2.5);
  });
  txt(s, [
    { text: "Training introduces the standard. ", options: { color: K.light } },
    { text: "Practice develops capability. ", options: { color: K.light } },
    { text: "Leadership reinforcement makes it stick.", options: { color: K.amber, bold: true } },
  ], { x: 0.6, y: 5.85, w: 12.1, h: 0.5, fontSize: 18, bold: true, margin: 0 });
  txt(s, "Use a real, safe, non-energized standard — e.g., a GMP hand-washing step or a CILT visual card. Debrief: 2 minutes.", { x: 0.6, y: 6.4, w: 12.1, h: 0.4, fontSize: 12, italic: true, color: K.light, margin: 0 });
  notes(s, `PURPOSE: Let participants experience how retention improves from hearing → doing → teaching.
  TIMING: 15 minutes (3 + 5 + 5 + 2 debrief).
  SETUP: Teams of 4–6. Pick ONE real, safe, non-energized standard (e.g., a GMP hand-washing sequence, a PPE donning sequence, or a printed CILT visual card for a non-hazardous inspection point). Never use an energized machine task for this exercise.
  ROUND 1 — HEAR IT (3 min): Supervisor reads the standard once, no visuals. Team members write the steps from memory. Count how many steps each person got right.
  ROUND 2 — DO IT (5 min): Team walks through the steps using the visual standard (simulation or mock-up). Supervisor coaches only with questions: "What comes next?", "What does normal look like?"
  ROUND 3 — TEACH IT (5 min): A different person teaches the standard back to the group without notes. Group checks against the standard.
  DEBRIEF QUESTIONS (2 min): Which round produced the most accurate recall? What would happen if we only did Round 1 — which is how most training is delivered? How can we build Rounds 2 and 3 into daily huddles?
  CONNECT: Daily huddles = spaced repetition. Coaching at the machine = Round 2. Operators teaching new hires = Round 3. Machine ownership grows from all three.
  KEY MESSAGE (read aloud): "Training introduces the standard. Practice develops capability. Leadership reinforcement makes it stick."`);

  }
  if (true) {
  // ======================= 10. BEYOND OPERATOR ERROR =======================
  section("Root-Cause Thinking");
  s = content("PART 3 · ROOT-CAUSE THINKING", "Stop at “operator error” and the problem comes back");
  txt(s, "WHERE INVESTIGATIONS OFTEN STOP", { x: 0.6, y: 1.5, w: 4.6, h: 0.3, fontSize: 11, bold: true, color: K.coral, charSpacing: 2, margin: 0 });
  txt(s, "ASK INSTEAD: WHY DID THE SYSTEM ALLOW IT?", { x: 6.0, y: 1.5, w: 6.7, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  const stops = [["“Operator error”", "What made the error easy to make and hard to detect?"], ["“Lack of training”", "Is there a clear, visual standard — and was competency verified, not just attendance?"], ["“Employee forgot”", "What cue, check or design should have prompted the right action at the right time?"], ["“Didn’t follow instructions”", "Were instructions correct, available, practical, and reinforced by leaders on every shift?"]];
  stops.forEach(([a, b], i) => {
    const y = 1.95 + i * 1.0;
    box(s, 0.6, y, 4.6, 0.82, { fill: K.coral, tr: 85 });
    iconOnly(s, "FaTimesCircle", 0.8, y + 0.21, 0.4, HEX.coral);
    txt(s, a, { x: 1.35, y, w: 3.7, h: 0.82, fontSize: 16, bold: true, color: K.navy, valign: "middle", margin: 0 });
    arrow(s, 5.3, y + 0.41, 5.9, y + 0.41, K.slate, 2.5);
    box(s, 6.0, y, 6.7, 0.82, { fill: K.teal, tr: 87 });
    iconOnly(s, "FaSearch", 6.2, y + 0.21, 0.4, HEX.teal);
    txt(s, b, { x: 6.75, y, w: 5.85, h: 0.82, fontSize: 14, color: K.ink, valign: "middle", margin: 0 });
  });
  box(s, 0.6, 6.05, 12.1, 0.75, { fill: K.navy });
  txt(s, [{ text: "People are part of the system. ", options: { bold: true, color: K.amber } }, { text: "Fix the system so the right way is the easy way — and the wrong way is hard to do and easy to see.", options: { color: K.white } }], { x: 0.85, y: 6.05, w: 11.7, h: 0.75, fontSize: 15, valign: "middle" });
  notes(s, `PURPOSE: Shift supervisors from blame-based conclusions to system-based root causes.
  TIMING: 5 minutes.
  SAY: "When we write 'operator error' or 'retrain employee' as the root cause, we are really writing down where we stopped looking."
  WALK EACH ROW: Each common stopping point has a system question behind it. Human error is real, but it is usually a symptom of the system: unclear standards, poor visual cues, equipment conditions that make the job harder, conflicting priorities, or no verification.
  NOTE: This is not about removing accountability. Accountability is stronger when the standard is clear, the equipment is in good condition and leaders reinforce the right behavior.
  ASK: "Look at the last three incident or downtime reports for the Gemini. How many root causes say 'retrain'? Did the problem come back?"`);

  }
  if (true) {
  // ======================= 11. RCA TOOLKIT =======================
  s = content("PART 3 · THE ROOT-CAUSE TOOLKIT", "Six tools supervisors will use — and coach teams to use");
  const tools = [["FaQuestionCircle", "5 Whys", "Keep asking why until you reach a cause you can act on and verify. Support every answer with evidence."], ["FaFish", "Fishbone diagram", "Brainstorm possible causes across Machine, Method, Material, Measurement, People and Environment — then test them."], ["FaMicroscope", "Evidence-based RCA", "Facts over opinion. Go to the machine. Check data, photos, worn parts and records."], ["FaClipboardList", "Standard-work verification", "Does a standard exist? Is it correct? Is it followed on every shift? Check all three."], ["FaCog", "Equipment-condition assessment", "Look for wear, looseness, contamination, misalignment, missing parts and lubrication issues."], ["FaRedo", "Recurrence prevention", "Countermeasure + updated standard + training + verification that the problem stays fixed."]];
  tools.forEach(([ic, h, d], i) => {
    const x = 0.6 + (i % 3) * 4.1, y = 1.5 + Math.floor(i / 3) * 2.7;
    box(s, x, y, 3.85, 2.45, { fill: K.light });
    iconCircle(s, ic, x + 0.25, y + 0.25, 0.7, i < 3 ? K.navy : K.teal);
    txt(s, h, { x: x + 1.1, y: y + 0.25, w: 2.6, h: 0.7, fontSize: 16, bold: true, color: K.navy, valign: "middle", margin: 0 });
    txt(s, d, { x: x + 0.25, y: y + 1.1, w: 3.4, h: 1.25, fontSize: 13, color: K.ink, margin: 0 });
  });
  notes(s, `PURPOSE: Introduce the six root-cause tools in plain language.
  TIMING: 6 minutes.
  5 WHYS: Simple but easy to misuse. Rules: (1) start from a clearly defined problem; (2) each 'why' must be supported by evidence, not opinion; (3) stop when you reach a cause you can act on and that would prevent recurrence; (4) branches are normal — one 'why' can have two answers.
  FISHBONE (Ishikawa): Use it to brainstorm POSSIBLE causes in categories (6M). Then test the most likely ones with data or observation. A fishbone is a list of hypotheses, not findings.
  EVIDENCE-BASED RCA: Go to gemba (the machine). Look at the part, the setting, the data, the record.
  STANDARD-WORK VERIFICATION: A standard can be missing, wrong, or not followed. Each has a different fix.
  EQUIPMENT-CONDITION ASSESSMENT: Many 'people' problems are actually equipment-condition problems that people compensate for.
  RECURRENCE PREVENTION: The problem is not solved until it stays solved. Verify over a defined period.
  ASK: "Which of these do we use today? Which one do we skip most often?" (Typical answer: evidence and recurrence verification.)`);

  }
  if (true) {
  // ======================= 12. GEMINI JAM CASE =======================
  s = content("PART 3 · ILLUSTRATIVE GEMINI CASE", "Case: jams keep returning after retraining");
  box(s, 0.6, 1.5, 5.6, 0.85, { fill: K.navy });
  txt(s, [{ text: "Scenario: ", options: { bold: true, color: K.amber } }, { text: "a Gemini repeatedly jams. Operators are retrained twice. The jams continue.", options: { color: K.white } }], { x: 0.8, y: 1.5, w: 5.3, h: 0.85, fontSize: 14, valign: "middle" });
  const whys = [["Why does it jam?", "Product misaligns at the transfer point"], ["Why does it misalign?", "Guide/sensor position varies between shifts"], ["Why does it vary?", "No documented, validated centerline for it"], ["Why isn’t wear caught?", "Guide wear isn’t a CILT or PM inspection point"], ["Why not?", "Standards were never built for this area"]];
  whys.forEach(([q, a], i) => {
    const y = 2.55 + i * 0.75;
    numCircle(s, i + 1, 0.6, y + 0.07, 0.5, K.teal, 13);
    box(s, 1.2, y, 5.0, 0.64, { fill: K.light });
    txt(s, [{ text: q + "  ", options: { bold: true, color: K.navy } }, { text: a, options: { color: K.ink } }], { x: 1.3, y, w: 4.85, h: 0.64, fontSize: 12, valign: "middle", margin: 0 });
  });
  txt(s, "Possible system causes: missing centerline + missing inspection point — verify each step with evidence.", { x: 0.6, y: 6.35, w: 5.6, h: 0.45, fontSize: 12, bold: true, italic: true, color: K.teal, margin: 0 });
  // fishbone
  badge(s, "POSSIBLE CAUSES — NOT CONFIRMED FINDINGS", 7.6, 1.5, 4.3, K.amber);
  const spineY = 4.15;
  line(s, 6.7, spineY, 11.45, spineY, K.navy, 3);
  box(s, 11.45, spineY - 0.45, 1.25, 0.9, { fill: K.coral });
  txt(s, "Recurring jams", { x: 11.45, y: spineY - 0.45, w: 1.25, h: 0.9, fontSize: 12, bold: true, color: K.white, align: "center", valign: "middle" });
  const cats = [["Machine", "Worn guide or transfer component"], ["Method", "No defined centerline; settings differ by shift"], ["Maintenance", "PM task missing an inspection point"], ["Measurement", "Sensor position not verified"], ["Material", "Product or packaging variation"], ["People", "Standard unclear; retraining treated the symptom"]];
  cats.forEach(([c, d], i) => {
    const top = i < 3, bx = 6.85 + (i % 3) * 1.55;
    if (top) line(s, bx, 2.85, bx + 0.6, spineY, K.slate, 1.75);
    else line(s, bx, 5.45, bx + 0.6, spineY, K.slate, 1.75);
    const ly = top ? 1.9 : 5.5;
    txt(s, [{ text: c, options: { bold: true, color: K.navy, breakLine: true, fontSize: 13 } }, { text: d, options: { color: K.ink, fontSize: 11 } }], { x: bx - 0.3, y: ly, w: 1.5, h: 0.95, margin: 0 });
  });
  notes(s, `PURPOSE: Show what real root-cause investigation looks like on a Gemini, using an illustrative case.
  TIMING: 10 minutes (walk-through 4 min; table discussion 6 min).
  IMPORTANT: This is an ILLUSTRATIVE example. The causes shown are POSSIBLE causes for discussion. They are not confirmed findings about any Gemini machine.
  WALK THE 5 WHYS: Notice how the answer moves from the person (jams → operators) to the system (missing centerline and missing inspection point). Each 'why' would need evidence: shift-by-shift setting checks, measured guide wear, CILT and PM task lists.
  WALK THE FISHBONE: The real investigation might uncover a worn component, an incorrect centerline, a missing inspection point, or an inadequate PM task — or something else. The team must test each possibility against evidence.
  TABLE DISCUSSION (6 min): "If this were your Gemini, what evidence would you collect first, and who would you involve from Maintenance and Quality?"
  EXPECTED ANSWERS: Check the jam location and frequency by shift in Redzone; inspect the guide/transfer component for wear; compare current sensor/guide positions across shifts; review PM and CILT task lists for that area; check product/packaging lots.
  ${AM ? "LINK FORWARD: the fixes here are exactly the AM tools we are about to deploy: centerlining, CILT inspection points, and defect handling." : "LINK FORWARD: root-cause thinking is how VO projects find countermeasures that last. A loss on the Pareto is a symptom; the project must find the system cause."}`);

  }
  if (true) {
  // ======================= 13. THREE CELLS =======================
  section("Gemini Deployment Model");
  s = content("PART 4 · THE GEMINI DEPLOYMENT MODEL", "Three machines, three cross-functional improvement cells");
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1;
    box(s, x, 1.5, 3.85, 4.25, { fill: K.white, line: HEX.grid, lw: 1, shadow: true });
    iconCircle(s, "FaIndustry", x + 0.25, 1.7, 0.75, K.navy);
    txt(s, "CELL " + (i + 1), { x: x + 1.15, y: 1.72, w: 2.5, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
    txt(s, "Gemini " + (i + 1), { x: x + 1.15, y: 2.0, w: 2.5, h: 0.45, fontSize: 22, bold: true, color: K.navy, margin: 0 });
    box(s, x + 0.25, 2.65, 3.35, 0.45, { fill: K.teal, r: 0.1 });
    txt(s, AM ? "Ownership: AM" : "Ownership: VO", { x: x + 0.25, y: 2.65, w: 3.35, h: 0.45, fontSize: 13, bold: true, color: K.white, align: "center", valign: "middle" });
    txt(s, [{ text: "Supervisor leader: ", options: { bold: true, color: K.navy } }, { text: "confirm at kickoff", options: { italic: true, color: K.slate } }], { x: x + 0.25, y: 3.25, w: 3.4, h: 0.35, fontSize: 13, margin: 0 });
    [["Production", K.teal], ["Maintenance", K.amber], ["Quality", K.blue]].forEach(([r, col], j) => {
      const y = 3.75 + j * 0.5;
      circ(s, x + 0.3, y + 0.1, 0.22, col);
      txt(s, [{ text: r + "  ", options: { bold: true, color: K.ink } }, { text: "members TBD", options: { color: K.slate, italic: true } }], { x: x + 0.65, y, w: 3.0, h: 0.42, fontSize: 13, valign: "middle", margin: 0 });
    });
    txt(s, [{ text: "Shift coverage: ", options: { bold: true, color: K.navy } }, { text: "TBD", options: { italic: true, color: K.slate } }], { x: x + 0.25, y: 5.3, w: 3.4, h: 0.35, fontSize: 13, margin: 0 });
  }
  box(s, 0.6, 5.95, 12.1, 0.85, { fill: K.light });
  iconCircle(s, "FaChalkboardTeacher", 0.8, 6.06, 0.62, K.teal);
  txt(s, [{ text: "Continuous Improvement coaches and facilitates — ", options: { bold: true, color: K.navy } }, { text: AM ? "the cells own the work and the results.  Plant leadership sponsors  ·  Maintenance & Quality are full cell members." : "the cells own the work and the results.  Plant leadership sponsors  ·  Finance (Randy) validates savings." }], { x: 1.6, y: 5.95, w: 11.0, h: 0.85, fontSize: 14, valign: "middle" });
  notes(s, `PURPOSE: Introduce the deployment structure.
  TIMING: 5 minutes.
  SAY: "Each Gemini gets its own cross-functional cell. Every cell has Production, Maintenance and Quality members, and one designated supervisor leader who owns the cell."
  ROLE OF CI: Continuous Improvement coaches and facilitates. CI will NOT perform every improvement project for the teams. The goal is to build supervisor and team capability — if CI does the work, the capability leaves when CI leaves.
  ${AM ? "COMMON FRAMEWORK: All three cells use the same deployment framework and the same four AM tools, but each develops equipment-specific standards." : "COMMON FRAMEWORK: All three cells use the same six VO steps and the same savings rules, but each builds its own baseline, loss Pareto and project list."}
  DO NOT INVENT: Actual machine IDs, personnel names and shift coverage are intentionally TBD. Confirm them at kickoff and update this slide.
  ASK: "What does a good cell member from Maintenance and Quality need from the supervisor to be effective?" (Answer: clear meeting rhythm, protected time, visible follow-through.)`);

  }
  if (AM) {
  // ======================= 14. COMMON FRAMEWORK =======================
  s = content("PART 4 · COMMON FRAMEWORK, MACHINE-SPECIFIC STANDARDS", "Same tools in every cell — standards built for each Gemini");
  const fwRows = [["Safety Map", "Template, hazard categories, approval route", "Layout, hazards, E-stops, LOTO points"], ["CILT", "Card format, frequency logic, escalation", "Inspection points, tasks, tools, photos"], ["Centerlining", "Register format, change-control rules", "Parameters, targets, ranges (validated)"], ["Defect Handling", "Tag, register, severity, escalation", "Defect history, owners, repeat issues"], ["Daily management", "Huddle agenda, KPI board, WOr template", "Cell actions, machine-specific KPIs"]];
  const fwData = [[hdr("Element"), hdr("Common to all three cells"), hdr("Specific to each Gemini")]].concat(fwRows.map((r, i) => [
    { text: r[0], options: { bold: true, color: K.navy, fill: { color: i % 2 ? K.white : K.light } } },
    { text: r[1], options: { fill: { color: i % 2 ? K.white : K.light } } },
    { text: r[2], options: { fill: { color: i % 2 ? K.white : K.light } } },
  ]));
  s.addTable(fwData, { x: 0.6, y: 1.5, w: 8.0, colW: [1.9, 3.05, 3.05], rowH: 0.62, fontSize: 12, color: K.ink, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: HEX.grid }, margin: 0.08 });
  box(s, 8.9, 1.5, 3.8, 5.3, { fill: K.amber, tr: 82 });
  iconOnly(s, "FaClipboardCheck", 9.1, 1.7, 0.45, HEX.navy);
  txt(s, "Confirm at kickoff — do not assume", { x: 9.65, y: 1.68, w: 2.95, h: 0.55, fontSize: 14, bold: true, color: K.navy, margin: 0, valign: "middle" });
  txt(s, bullets(["Machine IDs / asset numbers", "Supervisor leader per cell", "Production, Maintenance, Quality members", "Shift coverage per cell", "Products and changeovers on each Gemini", "Redzone line / asset mapping", "Existing PM tasks and OEM manuals", "LOTO and safety procedures in force"], { space: 6 }), { x: 9.1, y: 2.4, w: 3.45, h: 4.3, fontSize: 13 });
  notes(s, `PURPOSE: Clarify what is common across the cells and what is unique to each machine.
  TIMING: 4 minutes.
  SAY: "We will not let three cells invent three different systems. The templates, rules and escalation paths are common. What each cell builds is the machine-specific content: its own hazards, inspection points, centerline values and loss data."
  WHY: Common templates allow comparison across Geminis, shared learning, and easy expansion to other lines later.
  KICKOFF CHECKLIST: Walk through the right-hand list. Assign an owner and date for each item. None of these details have been assumed in this deck.
  ASK: "Are the three Geminis identical in design and product mix, or do they differ? What does that mean for our standards?"`);

  }
  if (AM) {
  // ======================= 15. WHAT AM IS =======================
  section("Autonomous Maintenance");
  s = content("PART 5 · AUTONOMOUS MAINTENANCE (AM)", "AM builds basic equipment conditions — not a cleaning program");
  const amNot = ["Just cleaning the machine", "Operators doing Maintenance’s job", "A one-time 5S blitz", "More paperwork for supervisors"];
  const amIs = ["Restore, maintain and monitor basic equipment conditions", "Operators who detect abnormalities early", "Shared ownership with Maintenance and Quality", "A step-by-step build of operator capability"];
  [[amNot, "AM IS NOT", "FaTimesCircle", K.coral, HEX.coral], [amIs, "AM IS", "FaCheckCircle", K.teal, HEX.teal]].forEach(([items, h, ic, col, hx], i) => {
    const x = 0.6 + i * 6.15;
    box(s, x, 1.5, 5.95, 2.45, { fill: col, tr: 88 });
    iconOnly(s, ic, x + 0.25, 1.68, 0.45, hx);
    txt(s, h, { x: x + 0.85, y: 1.68, w: 4.8, h: 0.45, fontSize: 18, bold: true, color: K.navy, valign: "middle", margin: 0 });
    txt(s, bullets(items, { space: 5 }), { x: x + 0.25, y: 2.25, w: 5.5, h: 1.65, fontSize: 14 });
  });
  txt(s, "THE CLASSIC AM STEP PATH (TPM)", { x: 0.6, y: 4.15, w: 8, h: 0.3, fontSize: 11, bold: true, color: K.slate, charSpacing: 2, margin: 0 });
  const steps7 = ["Initial cleaning & inspection", "Eliminate contamination sources & hard-to-reach areas", "Cleaning & lubrication standards", "General inspection", "Autonomous inspection", "Standardization", "Autonomous management"];
  steps7.forEach((t, i) => {
    const x = 0.6 + i * 1.75, pilot = i < 4;
    box(s, x, 4.5, 1.6, 1.3, { fill: pilot ? K.teal : K.light, tr: pilot ? 0 : 0 });
    txt(s, "STEP " + (i + 1), { x: x + 0.1, y: 4.57, w: 1.4, h: 0.25, fontSize: 10, bold: true, color: pilot ? K.white : K.slate, margin: 0 });
    txt(s, t, { x: x + 0.1, y: 4.82, w: 1.42, h: 0.95, fontSize: 11, bold: true, color: pilot ? K.white : K.navy, margin: 0 });
  });
  ["Safety Map", "CILT", "Centerlining", "Defect Handling"].forEach((t, i) => {
    const x = 0.6 + i * 1.75;
    box(s, x, 6.0, 1.6, 0.48, { fill: K.navy, r: 0.12 });
    txt(s, t, { x, y: 6.0, w: 1.6, h: 0.48, fontSize: 12, bold: true, color: K.white, align: "center", valign: "middle" });
  });
  txt(s, "Pilot foundation: four AM tools deployed in every Gemini cell, building capability for Steps 1–4", { x: 0.6, y: 6.5, w: 7.0, h: 0.3, fontSize: 11, italic: true, color: K.teal, margin: 0 });
  txt(s, "Later waves, once Steps 1–4 are sustained", { x: 7.6, y: 6.0, w: 5.1, h: 0.48, fontSize: 12, italic: true, color: K.slate, valign: "middle", align: "center", margin: 0 });
  notes(s, `PURPOSE: Define AM correctly and position the four pilot tools.
  TIMING: 5 minutes.
  DEFINITION: Autonomous Maintenance is one of the core pillars of Total Productive Maintenance (TPM). It is a structured system to restore, maintain and monitor basic equipment conditions while building operator ownership. "Basic conditions" = clean, inspected, properly lubricated, tight, and operating to standard settings.
  WHAT IT IS NOT: It is not "operators doing maintenance work." Technical repairs and PM stay with Maintenance. Operators perform only approved, trained, safe tasks.
  STEP PATH: The classic TPM AM path has seven steps. The Gemini pilot does not attempt all seven. It deploys four foundational tools — Safety Map, CILT, Centerlining and Defect Handling — that build capability for Steps 1–4. Steps 5–7 come in later waves once the foundation is sustained.
  NOTE ON CENTERLINING: Centerlining is not one of the original seven AM steps, but it is widely used in TPM-based systems in food and consumer-goods manufacturing as part of establishing and maintaining standard operating conditions.
  ASK: "If we only cleaned the Gemini more often, what would we still miss?"`);

  }
  if (AM) {
  // ======================= 16. SAFETY MAP =======================
  s = content("AM TOOL 1 · SAFETY MAPPING", "Safety map: see every hazard before touching the machine");
  box(s, 0.6, 1.5, 7.7, 5.3, { fill: K.light });
  badge(s, "ILLUSTRATIVE — NOT A REAL GEMINI LAYOUT", 0.8, 1.62, 4.2);
  // machine drawing
  box(s, 0.9, 3.65, 1.75, 0.7, { fill: K.slate, tr: 55, rect: true });
  txt(s, "Infeed conveyor", { x: 0.9, y: 3.65, w: 1.75, h: 0.7, fontSize: 11, bold: true, color: K.navy, align: "center", valign: "middle" });
  box(s, 2.55, 2.35, 3.4, 3.3, { fill: K.white, tr: 100, line: K.coral, lw: 2, dash: "dash", r: 0.05 });
  txt(s, "LOTO boundary", { x: 2.6, y: 2.38, w: 1.6, h: 0.25, fontSize: 10, bold: true, color: K.coral, margin: 0 });
  box(s, 2.8, 2.7, 2.9, 2.7, { fill: K.navy, r: 0.05 });
  txt(s, "GEMINI\n(main unit)", { x: 2.8, y: 3.4, w: 2.9, h: 1.0, fontSize: 15, bold: true, color: K.white, align: "center", valign: "middle" });
  box(s, 5.95, 3.65, 2.05, 0.7, { fill: K.slate, tr: 55, rect: true });
  txt(s, "Outfeed conveyor", { x: 5.95, y: 3.65, w: 2.05, h: 0.7, fontSize: 11, bold: true, color: K.navy, align: "center", valign: "middle" });
  box(s, 3.2, 5.85, 1.5, 0.6, { fill: K.blue, tr: 30, rect: true });
  txt(s, "Control panel / HMI", { x: 3.2, y: 5.85, w: 1.5, h: 0.6, fontSize: 10, bold: true, color: K.white, align: "center", valign: "middle" });
  line(s, 2.8, 2.65, 5.7, 2.65, K.green, 4);
  txt(s, "guarding", { x: 4.6, y: 2.4, w: 1.0, h: 0.22, fontSize: 10, bold: true, color: K.green, margin: 0 });
  // hazards
  const hz = [[1, 2.45, 3.3, K.coral], [2, 3.2, 2.95, K.amber], [3, 4.75, 5.85, K.blue], [4, 5.2, 2.95, K.slate], [5, 4.4, 4.6, K.amber], [1, 5.75, 4.45, K.coral]];
  hz.forEach(([n, x, y, col]) => numCircle(s, n, x, y, 0.4, col, 12));
  [[2.25, 5.0], [6.2, 2.85]].forEach(([x, y]) => {
    s.addShape(S.OCTAGON, { x, y, w: 0.42, h: 0.42, fill: { color: K.coral }, line: { color: K.white, width: 1.5 }, objectName: "E-stop marker " + ++objN });
    txt(s, "E", { x, y, w: 0.42, h: 0.42, fontSize: 12, bold: true, color: K.white, align: "center", valign: "middle", margin: 0 });
  });
  s.addShape(S.RIGHT_ARROW, { x: 0.9, y: 4.75, w: 1.2, h: 0.5, fill: { color: K.green }, line: { color: K.green, width: 0 }, objectName: "Safe access arrow" });
  txt(s, "Safe access", { x: 0.85, y: 5.3, w: 1.5, h: 0.3, fontSize: 10, bold: true, color: K.green, margin: 0 });
  txt(s, "Replace with a real Gemini photo or layout. Validate with EHS and Maintenance before use.", { x: 0.8, y: 6.45, w: 2.3, h: 0.32, fontSize: 9, italic: true, color: K.slate, margin: 0 });
  // legend
  box(s, 8.55, 1.5, 4.15, 2.95, { fill: K.white, line: HEX.grid, lw: 1 });
  txt(s, "MAP LEGEND", { x: 8.75, y: 1.6, w: 3, h: 0.28, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  const leg = [["1", K.coral, "Pinch / nip point"], ["2", K.amber, "Moving parts"], ["3", K.blue, "Electrical hazard"], ["4", K.slate, "Stored energy (pneumatic, gravity…)"], ["5", K.amber, "Cleaning & inspection hazard"], ["E", K.coral, "Emergency stop"], ["", K.green, "Guarding / safe access point"], ["", K.coral, "Lockout/tagout boundary (dashed)"]];
  leg.forEach(([n, col, t], i) => {
    const y = 1.95 + i * 0.3;
    circ(s, 8.8, y + 0.03, 0.22, col);
    if (n) txt(s, n, { x: 8.8, y: y + 0.03, w: 0.22, h: 0.22, fontSize: 8, bold: true, color: K.white, align: "center", valign: "middle", margin: 0 });
    txt(s, t, { x: 9.15, y, w: 3.45, h: 0.28, fontSize: 11, color: K.ink, valign: "middle", margin: 0 });
  });
  box(s, 8.55, 4.6, 4.15, 2.2, { fill: K.navy });
  txt(s, "DELIVERABLES PER CELL", { x: 8.75, y: 4.7, w: 3.8, h: 0.28, fontSize: 11, bold: true, color: K.amber, charSpacing: 2, margin: 0 });
  txt(s, bullets(["Approved safety map", "Hazard identification register", "Corrective action tracker", "Employee training record", "Safety verification checklist"], { space: 3 }), { x: 8.75, y: 5.0, w: 3.85, h: 1.75, fontSize: 13, color: K.white });
  notes(s, `PURPOSE: Introduce AM Tool 1 — the Safety Map — the first thing each cell builds.
  TIMING: 8 minutes.
  IMPORTANT: The map shown is ILLUSTRATIVE. It is not the real Gemini layout and does not identify real hazards. Each cell must build its own map from the actual machine, using real photos where available, and have it validated by authorized personnel (EHS and Maintenance) before use.
  WHAT THE MAP MUST SHOW: machine layout; pinch points; moving parts; electrical hazards; stored energy (pneumatic, hydraulic, spring, gravity, thermal); emergency stops; guarding; safe access points; cleaning and inspection hazards; lockout/tagout boundaries.
  PROCESS: Walk the machine (cell team + EHS) → identify hazards → photograph safely → draft map → validate with EHS and Maintenance → approve → train all operators on all shifts → verify with a safety checklist → post at the machine.
  DELIVERABLES: Approved safety map; hazard identification register; corrective action tracker (for hazards that need fixing, e.g., a missing guard); employee training record; safety verification checklist.
  RULE: Safety mapping happens BEFORE CILT tasks are written — you cannot define a safe cleaning or inspection task without knowing the hazards.
  ASK: "Where on your Gemini would a new operator be most likely to get hurt during cleaning?"`);

  }
  if (AM) {
  // ======================= 17. CILT ELEMENTS =======================
  s = content("AM TOOL 2 · CILT", "CILT: Cleaning, Inspection, Lubrication, Tightening");
  const cilt = [["C", "Cleaning", "FaBroom", "Cleaning is inspection. Removing build-up exposes leaks, cracks, wear and looseness.", "Operators — approved, safe methods within sanitation rules"], ["I", "Inspection", "FaEye", "Use senses and simple checks to compare the actual condition with the defined normal condition.", "Operators — with criteria defined by Maintenance and Quality"], ["L", "Lubrication", "FaOilCan", "Right lubricant, right point, right amount, right frequency — food-grade where required.", "Only as authorized — often Maintenance"], ["T", "Tightening", "FaWrench", "Find and correct looseness before it causes vibration, misalignment or failure.", "Operators only for approved, non-critical fasteners; otherwise Maintenance"]];
  cilt.forEach(([l, h, ic, d, who], i) => {
    const x = 0.6 + i * 3.07;
    box(s, x, 1.5, 2.88, 3.75, { fill: K.light });
    numCircle(s, l, x + 0.2, 1.7, 0.8, K.navy, 28);
    iconOnly(s, ic, x + 2.15, 1.85, 0.5, HEX.teal);
    txt(s, h, { x: x + 0.2, y: 2.62, w: 2.5, h: 0.42, fontSize: 18, bold: true, color: K.navy, margin: 0 });
    txt(s, d, { x: x + 0.2, y: 3.05, w: 2.5, h: 1.2, fontSize: 12, color: K.ink, margin: 0 });
    txt(s, [{ text: "Typical owner: ", options: { bold: true, color: K.teal } }, { text: who }], { x: x + 0.2, y: 4.3, w: 2.5, h: 0.9, fontSize: 11, color: K.ink, margin: 0 });
  });
  box(s, 0.6, 5.45, 12.1, 1.35, { fill: K.coral, tr: 86 });
  iconCircle(s, "FaLock", 0.85, 5.67, 0.8, K.coral);
  txt(s, [{ text: "Authorization rule: ", options: { bold: true, color: K.navy } }, { text: "lubrication, tightening, guard access and any maintenance activity follow equipment procedures, LOTO and safety requirements. No technical maintenance task is assigned to operators without Maintenance and EHS approval and documented training." }], { x: 1.85, y: 5.45, w: 10.7, h: 1.35, fontSize: 14, valign: "middle" });
  notes(s, `PURPOSE: Explain what CILT is and who does what.
  TIMING: 6 minutes.
  DEFINITION: CILT = Cleaning, Inspection, Lubrication and Tightening. These are the routine activities that keep basic equipment conditions in place.
  KEY IDEA: "Cleaning is inspection." When operators clean with attention, they find abnormalities early: leaks, cracks, loose bolts, worn belts, product build-up.
  WHO DOES WHAT: The "typical owner" lines are guidance only. The final split between operator tasks and Maintenance tasks is decided by Maintenance and EHS for each Gemini, based on the safety map, LOTO requirements and the training operators have received.
  AUTHORIZATION RULE (read aloud): Lubrication, tightening, guarding access and any maintenance activity must be authorized according to equipment procedures and safety requirements. We do not assign technical maintenance activities to operators without approval.
  FOOD SAFETY: Lubricants on or near product zones must meet the site's food-grade lubricant requirements; cleaning methods must follow sanitation procedures and allergen controls.
  ASK: "Which CILT tasks on the Gemini are clearly safe for operators? Which ones need a Maintenance decision?"`);

  }
  if (AM) {
  // ======================= 18. CILT CARD ANATOMY =======================
  s = content("AM TOOL 2 · CILT VISUAL STANDARD", "Anatomy of a visual CILT standard");
  box(s, 0.6, 1.5, 7.5, 5.3, { fill: K.white, line: K.navy, lw: 1.5 });
  box(s, 0.6, 1.5, 7.5, 0.5, { fill: K.navy, r: 0.04 });
  txt(s, "CILT STANDARD  ·  GEMINI [#]  ·  POINT [ID]  ·  [C / I / L / T]", { x: 0.8, y: 1.5, w: 7.1, h: 0.5, fontSize: 12, bold: true, color: K.white, valign: "middle", margin: 0 });
  box(s, 0.85, 2.2, 3.1, 2.35, { fill: K.light, line: K.slate, dash: "dash", lw: 1 });
  iconOnly(s, "FaCamera", 2.1, 2.75, 0.6, HEX.slate);
  txt(s, "Photo / diagram of the inspection point (real Gemini photo)", { x: 0.95, y: 3.45, w: 2.9, h: 0.9, fontSize: 11, color: K.slate, align: "center", margin: 0 });
  const fields = ["Task description", "Expected normal condition", "Frequency", "Responsible role", "Required tools", "Safety precautions", "Abnormality response", "Verification method"];
  fields.forEach((f, i) => {
    const x = 4.15 + (i % 2) * 1.9, y = 2.2 + Math.floor(i / 2) * 0.6;
    txt(s, [{ text: f, options: { bold: true, color: K.teal, fontSize: 10, breakLine: true } }, { text: "[ TBD ]", options: { color: K.slate, fontSize: 11 } }], { x, y, w: 1.85, h: 0.56, margin: 0 });
  });
  [["NORMAL", K.green, "Photo of normal condition"], ["ABNORMAL", K.coral, "Photo of abnormal condition"]].forEach(([h, col, d], i) => {
    const x = 0.85 + i * 3.6;
    box(s, x, 4.75, 3.4, 1.3, { fill: col, tr: 88, line: col, lw: 1 });
    txt(s, [{ text: h, options: { bold: true, color: col, breakLine: true } }, { text: d, options: { color: K.slate, italic: true } }], { x: x + 0.15, y: 4.85, w: 3.1, h: 1.1, fontSize: 12, margin: 0 });
  });
  txt(s, "Approved by:  Production ____   Maintenance ____   Quality ____   EHS ____   Rev/date ____", { x: 0.85, y: 6.2, w: 7.1, h: 0.4, fontSize: 10, color: K.slate, margin: 0 });
  badge(s, "ILLUSTRATIVE TEMPLATE", 5.6, 2.2 + 4 * 0.6 - 0.02 + 0.0, 2.3);
  txt(s, "THE 10 REQUIRED FIELDS", { x: 8.45, y: 1.5, w: 4.2, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  txt(s, bullets(["Inspection or maintenance point", "Photo or diagram", "Task description", "Expected normal condition", "Frequency", "Responsible role", "Required tools", "Safety precautions", "Abnormality response", "Verification method"], { num: true, space: 2 }), { x: 8.45, y: 1.85, w: 4.25, h: 3.3, fontSize: 13 });
  box(s, 8.45, 5.25, 4.25, 1.55, { fill: K.light });
  txt(s, [{ text: "Design rules", options: { bold: true, color: K.navy, breakLine: true } }, { text: "One point per card  ·  Photo over words  ·  Normal condition defined by Maintenance/Quality  ·  Posted at the point of use", options: { color: K.ink } }], { x: 8.6, y: 5.32, w: 4.0, h: 1.4, fontSize: 12, margin: 0 });
  notes(s, `PURPOSE: Show what a good, machine-specific visual CILT standard contains.
  TIMING: 5 minutes.
  WALK THE CARD: Every CILT task card has 10 elements: the point, a photo or diagram, the task description, the expected normal condition, frequency, responsible role, required tools, safety precautions, abnormality response (what to do if it's not normal), and how completion is verified.
  "NORMAL vs ABNORMAL" PHOTOS are the most powerful part of the card — they let anyone, on any shift, recognize an abnormality in seconds.
  APPROVAL: Each card is approved by Production, Maintenance, Quality and EHS before it is posted. This prevents unsafe or unauthorized tasks from reaching the floor.
  TEMPLATE: This is an illustrative template. Field values are TBD and must come from the real Gemini, using real photos.
  ACTIVITY OPTION (10 min): Give each table a blank card. Ask them to fill it out for a simple, non-hazardous inspection point they know well, then trade cards with another table to test whether someone else could follow it.`);

  }
  if (AM) {
  // ======================= 19. CILT FLOW =======================
  s = content("AM TOOL 2 · CILT DELIVERABLES & ESCALATION", "From CILT check to work order: every abnormality is closed");
  const dels = [["FaClipboardList", "CILT master sheet"], ["FaCheckCircle", "Daily checklist"], ["FaCalendarAlt", "Weekly checklist"], ["FaMapMarkedAlt", "Inspection-point map"], ["FaChartBar", "CILT completion dashboard"], ["FaExclamationTriangle", "Abnormality escalation process"]];
  dels.forEach(([ic, t], i) => {
    const x = 0.6 + i * 2.04;
    box(s, x, 1.5, 1.9, 1.35, { fill: K.light });
    iconOnly(s, ic, x + 0.72, 1.62, 0.45, HEX.teal);
    txt(s, t, { x: x + 0.1, y: 2.12, w: 1.7, h: 0.68, fontSize: 12, bold: true, color: K.navy, align: "center", valign: "middle", margin: 0 });
  });
  const fb = (x, y, w, h, t, fill = K.navy, col = K.white) => { box(s, x, y, w, h, { fill }); txt(s, t, { x: x + 0.05, y, w: w - 0.1, h, fontSize: 12, bold: true, color: col, align: "center", valign: "middle" }); };
  fb(0.6, 3.45, 2.1, 0.95, "Operator performs CILT check");
  s.addShape(S.DIAMOND, { x: 3.0, y: 3.3, w: 1.9, h: 1.25, fill: { color: K.amber }, line: { color: K.amber, width: 0 }, objectName: "Decision normal" });
  txt(s, "Normal condition?", { x: 3.15, y: 3.3, w: 1.6, h: 1.25, fontSize: 12, bold: true, color: K.navy, align: "center", valign: "middle" });
  fb(5.2, 3.45, 2.1, 0.95, "Record completion", K.green);
  fb(7.6, 3.45, 2.3, 0.95, "CILT completion dashboard", K.green);
  fb(10.2, 3.45, 2.5, 0.95, "Repeat abnormality? → Root-cause analysis", K.coral);
  fb(3.0, 5.4, 1.9, 0.95, "Make safe, tag & report", K.coral);
  fb(5.2, 5.4, 2.1, 0.95, "Work-order request (CMMS)", K.teal);
  fb(7.6, 5.4, 2.3, 0.95, "Maintenance plans & corrects", K.teal);
  fb(10.2, 5.4, 2.5, 0.95, "Verify, close & update standard", K.teal);
  arrow(s, 2.72, 3.92, 2.98, 3.92); arrow(s, 4.92, 3.92, 5.18, 3.92); arrow(s, 7.32, 3.92, 7.58, 3.92);
  arrow(s, 3.95, 4.57, 3.95, 5.38); arrow(s, 4.92, 5.87, 5.18, 5.87); arrow(s, 7.32, 5.87, 7.58, 5.87); arrow(s, 9.92, 5.87, 10.18, 5.87); arrow(s, 11.45, 5.38, 11.45, 4.42);
  txt(s, "YES", { x: 4.85, y: 3.55, w: 0.4, h: 0.25, fontSize: 10, bold: true, color: K.green, margin: 0 });
  txt(s, "NO", { x: 4.0, y: 4.75, w: 0.4, h: 0.25, fontSize: 10, bold: true, color: K.coral, margin: 0 });
  notes(s, `PURPOSE: Show the CILT deliverables and how every abnormality found during CILT is escalated.
  TIMING: 5 minutes.
  DELIVERABLES PER CELL: CILT master sheet (all points for that Gemini); daily checklist; weekly checklist; machine inspection-point map (numbered points matching the cards); CILT completion dashboard; abnormality escalation process.
  FLOW: Operator performs the CILT check. If the condition is normal, completion is recorded and feeds the dashboard. If not normal: make safe, tag and report → a work-order request goes into the maintenance system (CMMS — confirm the system used on site) → Maintenance plans and corrects → verification and closure → update the standard if the task or normal condition needs to change. Repeat abnormalities trigger root-cause analysis.
  KEY MESSAGE: "A CILT check that finds an abnormality but goes nowhere teaches operators that CILT doesn't matter. Closing the loop is what builds ownership."
  ASK: "Today, when an operator reports a problem on the Gemini, how do they know whether it was fixed?"`);

  }
  if (AM) {
  // ======================= 20. CENTERLINING DEF + ZONES =======================
  s = content("AM TOOL 3 · CENTERLINING", "Centerlining: the right settings, every run, every shift");
  box(s, 0.6, 1.45, 12.1, 0.9, { fill: K.teal, tr: 85 });
  txt(s, [{ text: "Centerlining ", options: { bold: true, color: K.navy } }, { text: "means identifying, documenting, validating and maintaining the correct operating conditions needed to achieve consistent equipment performance." }], { x: 0.85, y: 1.45, w: 11.7, h: 0.9, fontSize: 16, valign: "middle" });
  const bY = 3.0, bX = 0.9, bW = 11.5;
  box(s, bX, bY, 2.4, 0.9, { fill: K.coral, tr: 55, rect: true });
  box(s, bX + 2.4, bY, 6.7, 0.9, { fill: K.teal, tr: 70, rect: true });
  box(s, bX + 9.1, bY, 2.4, 0.9, { fill: K.coral, tr: 55, rect: true });
  const tx = bX + 2.4 + 3.35;
  line(s, tx, bY - 0.25, tx, bY + 1.15, K.navy, 3);
  txt(s, "TARGET", { x: tx - 0.6, y: bY - 0.55, w: 1.2, h: 0.28, fontSize: 12, bold: true, color: K.navy, align: "center", margin: 0 });
  txt(s, "Out of standard", { x: bX, y: bY + 0.95, w: 2.4, h: 0.3, fontSize: 12, bold: true, color: K.coral, align: "center", margin: 0 });
  txt(s, "Approved operating range", { x: bX + 2.4, y: bY + 0.95, w: 3.0, h: 0.3, fontSize: 12, bold: true, color: K.teal, align: "center", margin: 0 });
  txt(s, "Out of standard", { x: bX + 9.1, y: bY + 0.95, w: 2.4, h: 0.3, fontSize: 12, bold: true, color: K.coral, align: "center", margin: 0 });
  [[7.6, K.green, "Actual A — within range: run & monitor"], [11.0, K.coral, "Actual B — out of standard: follow response plan"]].forEach(([x, col, t]) => {
    s.addShape(S.ISOSCELES_TRIANGLE, { x: x - 0.2, y: bY + 0.2, w: 0.4, h: 0.45, rotate: 180, fill: { color: col }, line: { color: K.white, width: 1 }, objectName: "Actual marker " + ++objN });
    txt(s, t, { x: x - 1.3, y: bY + 1.3, w: 2.6, h: 0.55, fontSize: 11, bold: true, color: col, align: "center", valign: "top", margin: 0 });
  });
  const zones = [["Target setting", "The validated best setting for this product and condition", K.navy], ["Approved range", "Limits within which performance and quality are proven acceptable", K.teal], ["Actual condition", "What the machine is really running at, right now", K.green], ["Out-of-standard", "Outside the approved range — triggers a defined response and escalation", K.coral]];
  zones.forEach(([h, d, col], i) => {
    const x = 0.6 + i * 3.07;
    box(s, x, 5.0, 2.88, 1.8, { fill: K.light });
    circ(s, x + 0.2, 5.17, 0.3, col);
    txt(s, h, { x: x + 0.6, y: 5.12, w: 2.2, h: 0.4, fontSize: 15, bold: true, color: K.navy, valign: "middle", margin: 0 });
    txt(s, d, { x: x + 0.2, y: 5.6, w: 2.55, h: 1.1, fontSize: 12, color: K.ink, margin: 0 });
  });
  notes(s, `PURPOSE: Define centerlining and the four key concepts. This is one of the two most important hands-on sections of the program.
  TIMING: 8 minutes.
  DEFINITION (read aloud): Centerlining means identifying, documenting, validating and maintaining the correct operating conditions needed to achieve consistent equipment performance.
  THE FOUR CONCEPTS:
  • TARGET — the validated best setting for a given product/condition.
  • APPROVED OPERATING RANGE — the limits within which we have evidence that performance and quality are acceptable.
  • ACTUAL — what the machine is really running at right now (measured, not assumed).
  • OUT-OF-STANDARD — actual outside the approved range. This triggers a defined response, not an improvised adjustment.
  DIAGRAM: The band is conceptual. Actual A is inside the range — keep running and monitoring. Actual B is outside — stop adjusting by feel and follow the response plan (contain, notify, return to standard under authorization, verify product).
  ANALOGY: Like a recipe with oven temperature and time. A cook who 'adjusts by feel' every time produces inconsistent results — and when it goes wrong, nobody knows why.
  ASK: "For your Gemini, do we know the target and range for every critical setting? Where are they written down?"`);

  }
  if (AM) {
  // ======================= 21. RUNS BUT NOT RIGHT =======================
  s = content("AM TOOL 3 · CENTERLINING", "A running machine is not the same as a machine running right");
  box(s, 0.6, 1.5, 3.6, 5.3, { fill: K.navy });
  txt(s, "WHAT YOU SEE", { x: 0.85, y: 1.75, w: 3.1, h: 0.3, fontSize: 11, bold: true, color: K.amber, charSpacing: 2, margin: 0 });
  iconCircle(s, "FaIndustry", 1.65, 2.4, 1.5, K.teal);
  txt(s, "The Gemini is running.\nProduct is coming off the line.", { x: 0.85, y: 4.15, w: 3.1, h: 1.0, fontSize: 16, bold: true, color: K.white, align: "center", margin: 0 });
  txt(s, "So why are we missing the plan?", { x: 0.85, y: 5.5, w: 3.1, h: 0.8, fontSize: 14, italic: true, color: K.light, align: "center", margin: 0 });
  txt(s, "WHAT MAY BE HIDDEN", { x: 4.5, y: 1.5, w: 8, h: 0.3, fontSize: 11, bold: true, color: K.coral, charSpacing: 2, margin: 0 });
  const hid = [["FaBoxOpen", "Excess scrap"], ["FaHandPaper", "Frequent minor stops"], ["FaSlidersH", "Constant adjustments"], ["FaShieldAlt", "Inconsistent quality"], ["FaTachometerAlt", "Reduced speed"], ["FaRedo", "Long, variable startups"]];
  hid.forEach(([ic, t], i) => {
    const x = 4.5 + (i % 3) * 2.75, y = 1.9 + Math.floor(i / 3) * 1.25;
    box(s, x, y, 2.55, 1.05, { fill: K.coral, tr: 88 });
    iconOnly(s, ic, x + 0.2, y + 0.3, 0.45, HEX.coral);
    txt(s, t, { x: x + 0.8, y, w: 1.7, h: 1.05, fontSize: 14, bold: true, color: K.navy, valign: "middle", margin: 0 });
  });
  box(s, 4.5, 4.55, 8.2, 2.25, { fill: K.light });
  txt(s, "COMMON REASONS", { x: 4.75, y: 4.7, w: 7, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  const why = ["Settings drift slowly over time", "Undocumented “tribal knowledge” adjustments", "Each shift runs the machine differently", "Adjustments compensate for worn parts"];
  why.forEach((t, i) => {
    const x = 4.75 + (i % 2) * 4.0, y = 5.1 + Math.floor(i / 2) * 0.8;
    iconOnly(s, "FaExclamationTriangle", x, y + 0.12, 0.35, HEX.amber);
    txt(s, t, { x: x + 0.5, y, w: 3.4, h: 0.6, fontSize: 14, color: K.ink, valign: "middle", margin: 0 });
  });
  notes(s, `PURPOSE: Explain why a machine can run but still produce excessive scrap, stops or inconsistent quality.
  TIMING: 4 minutes.
  SAY: "'It's running' is not the same as 'it's running right.' Many of the biggest losses on a packaging or forming machine are hidden inside 'running' time: minor stops, reduced speed, adjustments, scrap and slow startups."
  COMMON REASONS: Settings drift over time; people make undocumented adjustments that 'worked once'; each shift develops its own way of running; and operators often compensate for worn parts by adjusting settings — which hides the real problem from Maintenance.
  KEY POINT: When settings are not centerlined, troubleshooting becomes guesswork because nobody knows what 'normal' looks like.
  ASK: "On your Gemini, is there a setting that each shift runs differently? Who decided which one is right?"`);

  }
  if (AM) {
  // ======================= 22. CENTERLINE REGISTER =======================
  s = content("AM TOOL 3 · CENTERLINE REGISTER", "Centerline register: one per Gemini, per product where needed");
  const ch = ["Parameter", "Target", "Approved range", "Actual", "Status", "Validated by"].map(hdr);
  const params = ["Machine speed", "Sensor alignment / position", "Timing", "Temperature (if applicable)", "Pressure (if applicable)", "Conveyor settings", "Product-specific adjustments", "Startup conditions", "Changeover settings"];
  const reg = [ch].concat(params.map((p, i) => {
    const f = { fill: { color: i % 2 ? K.white : K.light } };
    return [{ text: p, options: Object.assign({ bold: true, color: K.navy }, f) }, { text: "TBD", options: Object.assign({ color: K.slate }, f) }, { text: "TBD", options: Object.assign({ color: K.slate }, f) }, { text: "TBD", options: Object.assign({ color: K.slate }, f) }, { text: "Verify", options: { bold: true, color: K.navy, fill: { color: K.amber }, align: "center" } }, { text: "Maint + Quality", options: Object.assign({ color: K.ink }, f) }];
  }));
  s.addTable(reg, { x: 0.6, y: 1.5, w: 8.7, colW: [2.5, 1.1, 1.5, 1.1, 1.0, 1.5], rowH: 0.45, fontSize: 12, color: K.ink, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: HEX.grid }, margin: 0.07 });
  badge(s, "ILLUSTRATIVE — NO REAL SETTINGS", 0.6, 6.15, 3.3);
  txt(s, "Parameters are potential examples; the real list depends on the Gemini’s design. Enter values only after validation.", { x: 4.05, y: 6.08, w: 5.25, h: 0.6, fontSize: 11, italic: true, color: K.slate, margin: 0 });
  box(s, 9.6, 1.5, 3.1, 5.3, { fill: K.navy });
  txt(s, "WHERE VALID SETTINGS COME FROM", { x: 9.8, y: 1.65, w: 2.8, h: 0.55, fontSize: 11, bold: true, color: K.amber, charSpacing: 1, margin: 0 });
  txt(s, bullets(["OEM manual and specifications", "Maintenance expertise", "Quality and product requirements", "Data from proven good runs", "Structured, approved trials"], { space: 6 }), { x: 9.8, y: 2.3, w: 2.8, h: 2.8, fontSize: 13, color: K.white });
  box(s, 9.8, 5.2, 2.7, 1.4, { fill: K.coral, tr: 20 });
  txt(s, [{ text: "Never: ", options: { bold: true } }, { text: "copied from memory or “what worked last time.”" }], { x: 9.9, y: 5.2, w: 2.5, h: 1.4, fontSize: 13, color: K.white, valign: "middle" });
  notes(s, `PURPOSE: Introduce the centerline register — the controlled document for each Gemini.
  TIMING: 5 minutes.
  STRUCTURE: For every critical parameter: Target, Approved range, Actual (checked at defined times — e.g., startup, after changeover, each shift), Status (in / out of standard), and who validated it.
  POTENTIAL PARAMETERS (subject to the actual equipment design): operating speed, temperature, pressure, timing, alignment, sensor position, conveyor settings, product-specific adjustments, startup conditions, changeover settings.
  DO NOT INVENT VALUES: All values are TBD. Real targets and ranges must come from the OEM documentation, Maintenance, Quality requirements, data from proven good runs and structured, approved trials.
  PER PRODUCT: Where the Gemini runs several products, some parameters will need product-specific targets. Keep one register per Gemini with product columns, or one sheet per product — the cell decides with CI.
  VISUAL CONTROLS: Once validated, mark settings at the machine (e.g., colored range marks on gauges, position marks on guides, HMI recipe locks where available).
  ASK: "Which three parameters on your Gemini cause the most adjustment today? Start your register with those."`);

  }
  if (AM) {
  // ======================= 23. BENEFITS + CHANGE CONTROL =======================
  s = content("AM TOOL 3 · BENEFITS & CHANGE CONTROL", "Why centerlining pays off — and how changes are controlled");
  const ben = [["Consistent startup", "Same settings, same result, every shift"], ["Faster changeovers", "Go to the known setting instead of searching"], ["Reduced adjustment losses", "Less time and product lost to tweaking"], ["Better product quality", "Stable conditions, stable product"], ["Reduced scrap", "Fewer out-of-condition runs"], ["More repeatable performance", "Shift-to-shift differences shrink"], ["Easier troubleshooting", "Compare actual to standard — fast"]];
  ben.forEach(([h, d], i) => {
    const y = 1.5 + i * 0.76;
    numCircle(s, i + 1, 0.6, y + 0.08, 0.5, K.teal, 13);
    txt(s, h, { x: 1.25, y, w: 4.6, h: 0.35, fontSize: 15, bold: true, color: K.navy, margin: 0 });
    txt(s, d, { x: 1.25, y: y + 0.34, w: 4.6, h: 0.32, fontSize: 12, color: K.slate, margin: 0 });
  });
  const cx = 6.4;
  box(s, cx, 1.5, 6.3, 0.95, { fill: K.navy });
  txt(s, "Need to adjust a setting?  Is it within the approved range AND are you authorized?", { x: cx + 0.15, y: 1.5, w: 6.0, h: 0.95, fontSize: 14, bold: true, color: K.white, valign: "middle" });
  txt(s, "YES", { x: cx, y: 2.55, w: 2.9, h: 0.28, fontSize: 11, bold: true, color: K.green, align: "center", margin: 0 });
  txt(s, "NO", { x: cx + 3.4, y: 2.55, w: 2.9, h: 0.28, fontSize: 11, bold: true, color: K.coral, align: "center", margin: 0 });
  box(s, cx, 2.9, 2.9, 0.85, { fill: K.green });
  txt(s, "Adjust, record in log, verify output", { x: cx + 0.1, y: 2.9, w: 2.7, h: 0.85, fontSize: 12, bold: true, color: K.white, align: "center", valign: "middle" });
  const noSteps = ["STOP — don’t adjust. Escalate to supervisor", "Change request: Maintenance + Quality review", "Approved trial with verification", "Update register & standard; train all shifts"];
  noSteps.forEach((t, i) => {
    const y = 2.9 + i * 1.0;
    box(s, cx + 3.4, y, 2.9, 0.8, { fill: i === 0 ? K.coral : K.teal });
    txt(s, t, { x: cx + 3.5, y, w: 2.7, h: 0.8, fontSize: 12, bold: true, color: K.white, align: "center", valign: "middle" });
    if (i < 3) arrow(s, cx + 4.85, y + 0.82, cx + 4.85, y + 0.98, K.slate, 1.75);
  });
  box(s, cx, 3.95, 2.9, 2.85, { fill: K.coral, tr: 87 });
  txt(s, [{ text: "Why unapproved “tweaks” hurt", options: { bold: true, color: K.navy, breakLine: true } }, { text: "Each adjustment adds variation. The next shift inherits an unknown condition. The real cause — often wear — stays hidden.", options: { color: K.ink } }], { x: cx + 0.15, y: 4.05, w: 2.65, h: 2.65, fontSize: 12 });
  notes(s, `PURPOSE: Show the benefits of centerlining and the change-control rule.
  TIMING: 6 minutes.
  BENEFITS (left): 1) consistent startup, 2) faster changeovers, 3) reduced adjustment losses, 4) better product quality, 5) reduced scrap, 6) more repeatable performance, 7) easier troubleshooting. These are expected benefits — we will measure them through VO; we do not claim them in advance.
  CHANGE CONTROL (right): Centerline changes must follow defined authorization and verification rules. Within the approved range AND authorized → adjust, record, verify output. Outside the range, or not authorized → stop, escalate, change request reviewed by Maintenance and Quality, approved trial with verification, then update the register and standard and train all shifts.
  WHY TWEAKS HURT: Every unapproved adjustment adds variation, leaves the next shift with an unknown condition, and hides the true cause (e.g., a worn guide compensated for by moving a sensor).
  NOTE: Authorization levels (who may adjust which parameter) must be defined per Gemini by Maintenance and Quality and documented on the register.
  ASK: "Who is authorized today to change speed on your Gemini? Is that written anywhere?"`);

  }
  if (AM) {
  // ======================= 24. CENTERLINE DETECTIVE =======================
  s = dark("INTERACTIVE EXERCISE · 20 MINUTES", "The Centerline Detective");
  box(s, 0.6, 1.5, 4.6, 5.3, { fill: K.white });
  iconOnly(s, "FaSearch", 0.85, 1.7, 0.45, HEX.navy);
  txt(s, "CASE FILE — HYPOTHETICAL GEMINI X", { x: 1.4, y: 1.7, w: 3.7, h: 0.45, fontSize: 12, bold: true, color: K.navy, valign: "middle", margin: 0 });
  txt(s, bullets(["Minor stops have increased on 2nd shift since a product changeover", "Product appearance is inconsistent", "Operators increased speed to recover output", "Two operators moved a sensor bracket “to make it run”", "Maintenance replaced a component last week"], { space: 6 }), { x: 0.85, y: 2.3, w: 4.15, h: 2.9, fontSize: 13 });
  box(s, 0.85, 5.3, 4.1, 1.3, { fill: K.amber, tr: 30 });
  txt(s, "What would you check BEFORE changing any setting?", { x: 0.95, y: 5.3, w: 3.9, h: 1.3, fontSize: 15, bold: true, color: K.navy, valign: "middle" });
  const clues = [["Register", "Compare actual settings to the centerline register. What is out?"], ["Change", "What changed — product, material, changeover, maintenance work, shift?"], ["Sensor", "Is the sensor at its marked standard position?"], ["Condition", "Any wear, looseness or build-up on the replaced or nearby parts?"], ["Quality", "What do quality checks and scrap data show — and since when?"], ["Authority", "Who approved the speed and bracket changes? Were they recorded?"]];
  clues.forEach(([h, d], i) => {
    const x = 5.5 + (i % 3) * 2.42, y = 1.5 + Math.floor(i / 3) * 2.45;
    box(s, x, y, 2.28, 2.25, { fill: K.teal });
    txt(s, "CLUE " + (i + 1), { x: x + 0.15, y: y + 0.15, w: 2, h: 0.28, fontSize: 10, bold: true, color: K.light, charSpacing: 2, margin: 0 });
    txt(s, h, { x: x + 0.15, y: y + 0.42, w: 2, h: 0.42, fontSize: 18, bold: true, color: K.white, margin: 0 });
    txt(s, d, { x: x + 0.15, y: y + 0.9, w: 2.0, h: 1.3, fontSize: 12, color: K.white, margin: 0 });
  });
  txt(s, [{ text: "Rule: ", options: { bold: true, color: K.amber } }, { text: "diagnose first, adjust last — and only within approved, authorized limits.", options: { color: K.white } }], { x: 5.5, y: 6.4, w: 7.2, h: 0.4, fontSize: 14, margin: 0 });
  notes(s, `PURPOSE: Practice centerline thinking: investigate before adjusting.
  TIMING: 20 minutes (3 min brief, 10 min table work, 7 min debrief).
  SETUP: Tables of 4–6. Show only the CASE FILE first (cover the clue cards or print them as face-down cards). Each table lists what they would check, in order, before touching any setting.
  THIS IS HYPOTHETICAL: "Gemini X" is not a real machine and the facts are invented for training.
  REVEAL THE CLUES after 7 minutes, and let tables compare their list to the six clues.
  MODEL ANSWER:
  1. Register — compare actual settings to the centerline register for this product. (If there is no register, that is finding #1.)
  2. Change — something changed: the changeover, the replaced component, the speed increase, the bracket move. Each is a suspect.
  3. Sensor — the bracket was moved twice by different people; is it at the standard position?
  4. Condition — check the replaced component (correct part, correct installation?) and nearby parts for wear or looseness.
  5. Quality — use quality checks and scrap data to pinpoint when the problem started — before or after the changeover / repair?
  6. Authority — the speed and bracket changes were unapproved adjustments. They add variation and may now be the main cause.
  DEBRIEF: "Why can unapproved adjustments make it worse?" — they stack new variation on top of the original cause and destroy the evidence. Centerline changes must follow defined authorization and verification rules.
  LIKELY RESOLUTION PATH: Return to validated centerline (authorized person), verify the replaced component, verify product, then investigate root cause of the original stops.`);

  }
  if (AM) {
  // ======================= 25. DEFECT TYPES =======================
  s = content("AM TOOL 4 · DEFECT HANDLING", "Not all defects are the same — know which one you are facing");
  const dtypes = [["FaCog", "Equipment abnormality", K.amber, "Machine condition differs from normal", "Loose bolt, leak, noise, wear, build-up", "Maintenance (via tag)"], ["FaBoxOpen", "Product-quality defect", K.blue, "Product does not meet specification", "Weight, seal, appearance, foreign-material risk", "Quality — contain product"], ["FaExclamationTriangle", "Safety defect", K.coral, "Condition that could injure someone", "Damaged guard, failed interlock, slip hazard, exposed wiring", "Stop & make safe — EHS"], ["FaSlidersH", "Process deviation", K.teal, "Operating outside the approved standard or centerline", "Setting out of range, skipped step, wrong changeover setup", "Supervisor — Quality if product affected"], ["FaRedo", "Recurring equipment failure", K.navy, "Same problem returns after being “fixed”", "Repeat jams, repeat part failures", "Cell RCA team"]];
  dtypes.forEach(([ic, h, col, d, ex, who], i) => {
    const x = 0.6 + i * 2.45;
    box(s, x, 1.5, 2.3, 4.1, { fill: K.light });
    iconCircle(s, ic, x + 0.2, 1.7, 0.65, col);
    txt(s, h, { x: x + 0.2, y: 2.45, w: 1.95, h: 0.65, fontSize: 14, bold: true, color: K.navy, margin: 0 });
    txt(s, d, { x: x + 0.2, y: 3.1, w: 1.95, h: 0.75, fontSize: 12, color: K.ink, margin: 0 });
    txt(s, [{ text: "Examples: ", options: { bold: true, color: K.slate } }, { text: ex }], { x: x + 0.2, y: 3.85, w: 1.95, h: 0.95, fontSize: 11, color: K.ink, margin: 0 });
    txt(s, [{ text: "First call: ", options: { bold: true, color: K.teal } }, { text: who }], { x: x + 0.2, y: 4.85, w: 1.95, h: 0.7, fontSize: 11, color: K.ink, margin: 0 });
  });
  box(s, 0.6, 5.8, 12.1, 1.0, { fill: K.navy });
  txt(s, [{ text: "One event can be several types. ", options: { bold: true, color: K.amber } }, { text: "A loose guard shedding metal is a safety defect AND a food-safety risk. Always handle the highest risk first.", options: { color: K.white } }], { x: 0.85, y: 5.8, w: 11.7, h: 1.0, fontSize: 15, valign: "middle" });
  notes(s, `PURPOSE: Teach supervisors to classify defects correctly — this drives who responds and how fast. This is the second major hands-on section.
  TIMING: 6 minutes.
  FIVE TYPES:
  • Equipment abnormality — any deviation from normal machine condition (the 'normal condition' on the CILT card).
  • Product-quality defect — product does not meet specification. Quality is involved and product is contained per site procedure.
  • Safety defect — any condition that could injure someone. Stop and make safe first; EHS involved.
  • Process deviation — operating outside the approved standard or centerline.
  • Recurring equipment failure — the same failure returns after being 'fixed'. This requires root-cause analysis, not another quick fix.
  KEY RULE: One event can be several types at once. Always address the highest risk first: people, then food safety and product, then equipment, then output.
  ASK: "Give me an example from the Gemini of each type." Capture on a flip chart — these become the starting point for each cell's defect register.`);

  }
  if (AM) {
  // ======================= 26. DEFECT CYCLE =======================
  s = content("AM TOOL 4 · THE DEFECT-HANDLING CYCLE", "From detection to prevention — nine steps, every time");
  const cyc = ["Detect", "Make safe / Contain", "Tag", "Report", "Assign", "Correct", "Verify", "Close", "Prevent recurrence"];
  const cX = 6.67, cY = 4.15, rx = 4.55, ry = 2.05;
  s.addShape(S.OVAL, { x: cX - rx, y: cY - ry, w: 2 * rx, h: 2 * ry, fill: { color: K.white, transparency: 100 }, line: { color: K.teal, width: 2, dashType: "dash" }, objectName: "Cycle loop" });
  cyc.forEach((t, i) => {
    const a = -Math.PI / 2 + (i * 2 * Math.PI) / cyc.length;
    const x = cX + rx * Math.cos(a) - 0.95, y = cY + ry * Math.sin(a) - 0.37;
    const col = i === 1 ? K.coral : i === 8 ? K.amber : K.navy;
    box(s, x, y, 1.9, 0.74, { fill: col, shadow: true });
    txt(s, [{ text: (i + 1) + "  ", options: { color: i === 8 ? K.navy : K.amber, bold: true } }, { text: t, options: { color: i === 8 ? K.navy : K.white, bold: true } }], { x: x + 0.05, y, w: 1.8, h: 0.74, fontSize: 13, align: "center", valign: "middle" });
  });
  box(s, cX - 2.1, cY - 0.75, 4.2, 1.5, { fill: K.teal, tr: 85 });
  txt(s, "A defect isn’t closed when it’s fixed. It’s closed when the fix is verified — and the cause is prevented from returning.", { x: cX - 2.0, y: cY - 0.75, w: 4.0, h: 1.5, fontSize: 14, bold: true, color: K.navy, align: "center", valign: "middle" });
  notes(s, `PURPOSE: Teach the full defect-handling cycle.
  TIMING: 6 minutes.
  THE NINE STEPS:
  1. Detect — during CILT, operation, changeover or a quality check.
  2. Make safe / Contain — protect people and product first. Stop the machine if needed; contain suspect product per Quality procedure.
  3. Tag — physically tag the location (where safe) so the defect is visible.
  4. Report — enter it in the defect register with a photo.
  5. Assign — a named owner and due date, based on severity.
  6. Correct — the responsible department fixes it (Maintenance for technical work).
  7. Verify — someone other than the fixer confirms the condition is back to normal (and product is OK if relevant).
  8. Close — the record is closed only after verification.
  9. Prevent recurrence — for repeat or significant defects: root cause, update CILT/PM/centerline standard, train.
  KEY LINE (read aloud): "A defect isn't closed when it's fixed. It's closed when the fix is verified and the cause is prevented from returning."
  ASK: "Which step do we skip most often today?" (Usual answers: verify and prevent recurrence.)`);

  }
  if (AM) {
  // ======================= 27. TAG + RECORD =======================
  s = content("AM TOOL 4 · DEFECT TAG & RECORD", "Tag it, log it, own it: the defect record");
  box(s, 0.9, 1.5, 3.7, 4.2, { fill: K.coral, shadow: true });
  circ(s, 2.55, 1.65, 0.4, K.white);
  txt(s, "DEFECT TAG  ·  No. [____]", { x: 1.05, y: 2.15, w: 3.4, h: 0.4, fontSize: 14, bold: true, color: K.white, align: "center", margin: 0 });
  box(s, 1.1, 2.6, 3.3, 2.95, { fill: K.white });
  const tagF = ["Machine:", "Location / point:", "Defect:", "Found by / date / shift:", "Severity:  S1  S2  S3  S4", "Owner:", "Due date:"];
  txt(s, tagF.map((f, i) => ({ text: f + (i === 4 ? "" : "  ________"), options: { breakLine: i < tagF.length - 1, bold: i === 4 } })), { x: 1.2, y: 2.68, w: 3.1, h: 2.8, fontSize: 12, color: K.navy, paraSpaceAfter: 5 });
  badge(s, "ILLUSTRATIVE — ALIGN WITH SITE STANDARD", 0.6, 5.85, 4.3);
  txt(s, [{ text: "Red tag: ", options: { bold: true, color: K.coral } }, { text: "Maintenance action   " }, { text: "Blue tag: ", options: { bold: true, color: K.blue } }, { text: "operator action (approved tasks only)" }], { x: 0.6, y: 6.25, w: 4.4, h: 0.55, fontSize: 11, margin: 0 });
  const recF = [["Machine number", "Gemini ID as confirmed at kickoff"], ["Location", "Point ID from the inspection-point map"], ["Defect description", "What, where, how much — facts, not blame"], ["Photo", "Taken safely; shows the condition clearly"], ["Date identified", "Date and shift found"], ["Severity", "Per escalation matrix (S1–S4)"], ["Responsible department", "Production, Maintenance, Quality or EHS"], ["Assigned owner", "A named person, not a department"], ["Due date", "Set by severity rules"], ["Status", "Open · In progress · Corrected · Verified · Closed"], ["Corrective action", "What was done, incl. root cause if repeat"], ["Closure verification", "Who verified, how, and when"]];
  const recData = [[hdr("Record field"), hdr("What good looks like")]].concat(recF.map((r, i) => [{ text: r[0], options: { bold: true, color: K.navy, fill: { color: i % 2 ? K.white : K.light } } }, { text: r[1], options: { fill: { color: i % 2 ? K.white : K.light } } }]));
  s.addTable(recData, { x: 5.3, y: 1.5, w: 7.4, colW: [2.5, 4.9], rowH: 0.405, fontSize: 12, color: K.ink, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: HEX.grid }, margin: 0.06 });
  notes(s, `PURPOSE: Show the visual defect tag and the 12 fields of every defect record.
  TIMING: 5 minutes.
  TAG (left): The physical tag makes the defect visible at the machine. Colors are illustrative — a common TPM convention is red tags for Maintenance action and blue (or white) tags for operator action. Align colors with any existing site tagging standard. Safety-critical defects follow the site safety procedure immediately — a tag never replaces making the area safe.
  TAG SAFETY: Tags are placed only where it is safe to do so and must not interfere with guarding, LOTO devices or product zones. Tags used near product must meet food-safety requirements (e.g., detectable/secured, no loose parts) — confirm with Quality.
  RECORD (right): Every defect in the register has the 12 fields shown. Emphasize: a NAMED owner (not "Maintenance"), and closure VERIFICATION by someone other than the person who fixed it.
  ASK: "What happens today when a defect has no named owner?"`);

  }
  if (AM) {
  // ======================= 28. ESCALATION =======================
  s = content("AM TOOL 4 · ESCALATION", "Escalation: safety and food safety come first — always");
  const esc = [
    ["S1  Safety-critical / food-safety risk", "Guard or interlock failure, exposed energy, foreign-material risk, allergen or sanitation concern", "STOP. Make safe. Contain affected product. Do not restart until released.", "Supervisor + EHS and/or Quality immediately; Maintenance", "Immediate", K.coral],
    ["S2  High", "Condition likely to cause breakdown or quality loss soon", "Contain risk. Run only under an approved procedure, if permitted.", "Supervisor, Maintenance, Quality as relevant", "Within shift (TBD)", K.amber],
    ["S3  Medium", "Abnormality with no immediate risk", "Tag, log, schedule correction", "Maintenance planner / cell owner", "TBD per site", K.blue],
    ["S4  Low", "Cosmetic issue or improvement opportunity", "Tag, log, plan with the cell", "Cell owner", "TBD per site", K.slate],
  ];
  const escData = [[hdr("Severity"), hdr("Examples"), hdr("Immediate action"), hdr("Notify"), hdr("Response")]].concat(esc.map((r) => [
    { text: r[0], options: { bold: true, color: K.white, fill: { color: r[5] } } },
    { text: r[1] }, { text: r[2], options: { bold: r[5] === K.coral } }, { text: r[3] }, { text: r[4], options: { bold: true, color: K.navy } },
  ]));
  s.addTable(escData, { x: 0.6, y: 1.5, w: 12.1, colW: [2.4, 3.3, 2.9, 2.2, 1.3], rowH: [0.45, 1.05, 0.95, 0.75, 0.75], fontSize: 12, color: K.ink, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: HEX.grid }, margin: 0.08 });
  box(s, 0.6, 5.7, 12.1, 1.1, { fill: K.light });
  iconOnly(s, "FaShieldAlt", 0.85, 5.95, 0.6, HEX.coral);
  txt(s, [{ text: "Draft for validation: ", options: { bold: true, color: K.navy } }, { text: "examples and response times must be confirmed with EHS, Quality and Maintenance. Site safety and food-safety procedures always govern — this matrix never overrides them." }], { x: 1.65, y: 5.7, w: 10.9, h: 1.1, fontSize: 14, valign: "middle" });
  notes(s, `PURPOSE: Define how defects are escalated, with safety-critical and food-safety-related defects at the top.
  TIMING: 5 minutes.
  S1 (SAFETY-CRITICAL / FOOD-SAFETY): Stop, make safe, contain affected product, notify the supervisor and EHS and/or Quality immediately. The machine does not restart until released by the authorized person. Examples: failed interlock, missing guard, exposed energy, foreign-material risk (e.g., metal fragments, broken plastic near product), allergen or sanitation concerns.
  S2 (HIGH): Likely to cause breakdown or quality loss soon. Contain the risk; continue only if an approved procedure permits.
  S3 / S4: Tag, log, schedule or plan.
  RESPONSE TIMES are TBD — they must be set with EHS, Quality and Maintenance and align with existing site procedures (food safety plan, HACCP/PCP where applicable, safety program).
  KEY MESSAGE: "Safety and compliance are never traded for output — or for savings."
  ASK: "Does everyone on your shift know they have the authority to stop the Gemini for an S1 condition?"`);

  }
  if (AM) {
  // ======================= 29. DEFECT EXERCISE =======================
  s = dark("INTERACTIVE EXERCISE · 20 MINUTES", "Stop, contain, escalate — or continue?");
  const scen = [["A", "During CILT, an operator finds fresh metal shavings near a product-contact area on Gemini X."], ["B", "A bearing on Gemini X has made a growing noise for three days. No heat, no product impact, no safety exposure."], ["C", "An interlocked guard door is found taped open “to speed up jam clearing.”"], ["D", "The same jam has occurred four times this week. Each time it is cleared using the approved jam-clearing procedure."]];
  scen.forEach(([l, t], i) => {
    const x = 0.6 + (i % 2) * 4.3, y = 1.5 + Math.floor(i / 2) * 2.4;
    box(s, x, y, 4.1, 2.2, { fill: K.white });
    numCircle(s, l, x + 0.2, y + 0.2, 0.6, K.amber, 18, K.navy);
    txt(s, "SCENARIO " + l, { x: x + 0.95, y: y + 0.27, w: 2.8, h: 0.45, fontSize: 12, bold: true, color: K.slate, charSpacing: 2, valign: "middle", margin: 0 });
    txt(s, t, { x: x + 0.2, y: y + 0.9, w: 3.7, h: 1.2, fontSize: 13, color: K.ink, margin: 0 });
  });
  box(s, 9.3, 1.5, 3.4, 4.6, { fill: K.teal });
  txt(s, "YOUR DECISION", { x: 9.5, y: 1.65, w: 3, h: 0.3, fontSize: 11, bold: true, color: K.white, charSpacing: 2, margin: 0 });
  ["STOP", "CONTAIN", "ESCALATE", "CONTINUE*"].forEach((t, i) => {
    box(s, 9.5, 2.05 + i * 0.6, 3.0, 0.48, { fill: K.white, r: 0.12 });
    txt(s, t, { x: 9.5, y: 2.05 + i * 0.6, w: 3.0, h: 0.48, fontSize: 14, bold: true, color: K.navy, align: "center", valign: "middle" });
  });
  txt(s, bullets(["Is anyone at risk?", "Is product at risk?", "Is there an approved procedure?", "Is it recurring?"], { space: 2 }), { x: 9.5, y: 4.5, w: 3.05, h: 1.5, fontSize: 12, color: K.white });
  txt(s, "* Continue only under an approved procedure — and still tag, report and assign.", { x: 0.6, y: 6.4, w: 12.1, h: 0.4, fontSize: 13, italic: true, color: K.light, margin: 0 });
  notes(s, `PURPOSE: Practice defect decisions: stop, contain, escalate, or continue under an approved procedure.
  TIMING: 20 minutes (2 min setup, 10 min table decisions, 8 min debrief).
  SETUP: All scenarios are hypothetical. Tables decide for each scenario: which actions (more than one can apply), defect type(s), severity (S1–S4), who is notified, and what happens to product.
  MODEL ANSWERS:
  A — Metal shavings near product contact: STOP, CONTAIN product (since last good check, per Quality procedure), ESCALATE immediately to Quality and Maintenance. Severity S1 (food-safety risk + equipment abnormality). Do not restart until released. Then root cause (what is wearing?).
  B — Growing bearing noise, no other symptoms: CONTINUE under approved procedure, but TAG, report, and raise a work order; Maintenance assesses (vibration/temperature check). Severity S2/S3 depending on Maintenance assessment. Escalate immediately if noise, heat or vibration changes.
  C — Interlocked guard taped open: STOP, MAKE SAFE (remove bypass under proper procedure), ESCALATE to EHS. Severity S1 safety defect. Also ask WHY: the jam frequency is creating pressure to bypass — that is a system problem to solve, not just a discipline issue.
  D — Repeat jam cleared with approved procedure: CONTINUE using the approved procedure, but TAG as a RECURRING failure and launch root-cause analysis (centerline check, condition check — connect to the Centerline Detective and the jam case).
  DEBRIEF: How did Maintenance, Production and Quality each contribute to the decision? Production detects and contains; Maintenance assesses and corrects equipment; Quality decides product disposition; EHS governs safety.`);

  }
  if (AM) {
  // ======================= 30. AM/PM/QUALITY =======================
  section("PM & Quality Integration");
  s = content("PART 6 · PLANNED MAINTENANCE & QUALITY INTEGRATION", "AM, PM and Quality: three roles, one equipment system");
  const ven = [["AM", "Production", 0.75, 1.5, K.teal, "Recognize abnormalities; maintain approved basic conditions"], ["PM", "Maintenance", 3.25, 1.5, K.amber, "Planned technical work; restore condition; prevent breakdowns"], ["QUALITY", "Quality", 2.0, 3.55, K.blue, "Define requirements; verify process conditions; control nonconforming product"]];
  ven.forEach(([h, who, x, y, col]) => circ(s, x, y, 3.1, col, { tr: 72 }));
  txt(s, [{ text: "AM", options: { bold: true, fontSize: 20, color: K.navy, breakLine: true } }, { text: "Recognize abnormalities; maintain approved basic conditions", options: { fontSize: 11, color: K.ink } }], { x: 0.95, y: 2.05, w: 1.9, h: 1.4, align: "center", margin: 0 });
  txt(s, [{ text: "PM", options: { bold: true, fontSize: 20, color: K.navy, breakLine: true } }, { text: "Planned technical work; restore condition; prevent breakdowns", options: { fontSize: 11, color: K.ink } }], { x: 4.25, y: 2.05, w: 1.9, h: 1.4, align: "center", margin: 0 });
  txt(s, [{ text: "QUALITY", options: { bold: true, fontSize: 18, color: K.navy, breakLine: true } }, { text: "Define requirements; verify process conditions; control nonconforming product", options: { fontSize: 11, color: K.ink } }], { x: 2.55, y: 5.1, w: 2.0, h: 1.4, align: "center", margin: 0 });
  txt(s, "Stable, safe, capable Gemini", { x: 2.85, y: 3.7, w: 1.45, h: 0.75, fontSize: 11, bold: true, color: K.navy, align: "center", valign: "middle", margin: 0 });
  txt(s, "HOW INFORMATION FLOWS", { x: 7.0, y: 1.5, w: 5.7, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  const flows = [["CILT finding", "Defect tag → work-order request → PM planning & backlog", K.teal], ["PM finding", "New inspection point or centerline update → operator training", K.amber], ["Quality trend", "e.g., rising scrap on one Gemini → equipment-condition check → RCA", K.blue]];
  flows.forEach(([h, d, col], i) => {
    const y = 1.88 + i * 0.95;
    box(s, 7.0, y, 5.7, 0.82, { fill: col, tr: 85 });
    txt(s, h, { x: 7.15, y, w: 1.6, h: 0.82, fontSize: 13, bold: true, color: K.navy, valign: "middle", margin: 0 });
    txt(s, d, { x: 8.8, y, w: 3.8, h: 0.82, fontSize: 12, color: K.ink, valign: "middle", margin: 0 });
  });
  txt(s, "SHARED ACTION & ESCALATION PROCESS", { x: 7.0, y: 4.85, w: 5.7, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  const escSteps = ["Detect", "Triage", "Assign", "Review", "Unblock"];
  escSteps.forEach((t, i) => s.addShape(i === 0 ? S.PENTAGON : S.CHEVRON, { x: 7.0 + i * 1.15, y: 5.25, w: 1.2, h: 1.0, fill: { color: i === 4 ? K.navy : K.teal }, line: { color: K.white, width: 1 }, objectName: "Escalation step " + (i + 1) }));
  escSteps.forEach((t, i) => txt(s, t, { x: 7.0 + i * 1.15 + (i === 0 ? 0.08 : 0.3), y: 5.25, w: 0.72, h: 1.0, fontSize: 10, bold: true, color: K.white, align: "center", valign: "middle", margin: 0 }));
  txt(s, "Open actions live on one shared cell action list.", { x: 7.0, y: 6.4, w: 5.7, h: 0.35, fontSize: 12, italic: true, color: K.slate, margin: 0 });
  notes(s, `PURPOSE: Show how AM, PM and Quality support one another.
  TIMING: 6 minutes.
  AM (Production): Operators recognize abnormalities and maintain approved basic conditions.
  PM (Maintenance): Maintenance performs planned technical work, restores equipment condition and prevents breakdowns.
  QUALITY: Defines product requirements, verifies process conditions and controls nonconforming product.
  THE OVERLAP is where the Gemini becomes stable, safe and capable — no single function can achieve this alone.
  INFORMATION FLOWS:
  • CILT findings feed the maintenance work-order process (via defect tags) and the PM backlog.
  • PM findings feed back into AM — e.g., Maintenance identifies a new inspection point or a setting that needs a centerline.
  • Quality trends can reveal equipment problems — e.g., rising scrap or seal defects on one Gemini may indicate a worn part or drifted setting.
  SHARED ACTION & ESCALATION: Detect → supervisor triage (severity, type) → owner assigned (Production, Maintenance or Quality) → weekly cell review of open actions → plant leadership removes barriers that the cell cannot.
  ASK: "How does Maintenance find out today about what operators see during cleaning?"`);

  }
  if (VO) {
  // ======================= 31. AM vs VO =======================
  section("Value Optimization");
  s = content("PART 7 · VALUE OPTIMIZATION (VO)", "Two pillars, two jobs: AM stabilizes, VO targets the losses");
  const pil = [["AUTONOMOUS MAINTENANCE", "FaCogs", K.teal, [["Purpose", "Establish, restore and maintain basic equipment conditions; prevent abnormalities"], ["Key question", "Is the machine in its normal, safe, standard condition?"], ["Outputs", "Safety maps, CILT, centerlines, defect closure"], ["Led by", "Cell supervisor + Production, with Maintenance & Quality"]]], ["VALUE OPTIMIZATION", "FaChartLine", K.amber, [["Purpose", "Identify, quantify, prioritize and validate opportunities to eliminate losses"], ["Key question", "Where are we losing, what is it worth, and what do we fix first?"], ["Outputs", "Baseline, loss Pareto, project pipeline, Finance-validated savings"], ["Led by", "Cell supervisor, with CI coaching and Finance validation"]]]];
  pil.forEach(([h, ic, col, rowsP], i) => {
    const x = i === 0 ? 0.6 : 7.7;
    box(s, x, 1.5, 5.0, 4.05, { fill: col, tr: 87 });
    iconCircle(s, ic, x + 0.25, 1.65, 0.65, col);
    txt(s, h, { x: x + 1.05, y: 1.65, w: 3.8, h: 0.65, fontSize: 16, bold: true, color: K.navy, valign: "middle", margin: 0 });
    rowsP.forEach(([k, v], j) => txt(s, [{ text: k + ": ", options: { bold: true, color: K.navy } }, { text: v }], { x: x + 0.25, y: 2.5 + j * 0.75, w: 4.55, h: 0.72, fontSize: 12, margin: 0 }));
  });
  const loop = ["AM stabilizes conditions", "VO measures & finds losses", "VO prioritizes & validates", "AM sustains the new standard"];
  loop.forEach((t, i) => {
    const y = 1.55 + i * 1.0;
    box(s, 5.8, y, 1.7, 0.78, { fill: i % 3 === 0 ? K.teal : K.amber });
    txt(s, t, { x: 5.82, y, w: 1.66, h: 0.78, fontSize: 11, bold: true, color: K.white, align: "center", valign: "middle", margin: 0 });
    if (i < 3) arrow(s, 6.65, y + 0.8, 6.65, y + 0.98, K.slate, 1.75);
  });
  box(s, 0.6, 5.75, 12.1, 1.05, { fill: K.navy });
  txt(s, [{ text: "Not interchangeable. ", options: { bold: true, color: K.amber } }, { text: "AM without VO keeps machines stable but may miss the biggest losses. VO without AM finds losses that return, because basic conditions are not in place.", options: { color: K.white } }], { x: 0.85, y: 5.75, w: 11.7, h: 1.05, fontSize: 14, valign: "middle" });
  notes(s, `PURPOSE: Introduce VO as the second pillar and make the AM/VO distinction crystal clear.
  TIMING: 5 minutes.
  AM: Establishes equipment conditions and prevents abnormalities.
  VO: Identifies where losses occur, quantifies their impact and supports improvement prioritization — and makes sure benefits are measured and validated.
  HOW THEY WORK TOGETHER (middle): AM stabilizes conditions → VO measures and finds where the losses are → VO prioritizes and validates improvements → AM sustains the new standard (updated CILT, centerline, etc.).
  CONSISTENT TERMINOLOGY: Use "AM" and "VO" as the two pillars throughout. Do not use them interchangeably.
  ASK: "If we only did AM, how would we know whether we were working on the right losses? If we only did VO, why might savings disappear?"`);

  }
  if (VO) {
  // ======================= 32. VO STEPS + STEP 1 =======================
  s = content("PART 7 · VO STEP 1", "Establish the baseline: 4–8 weeks of data for each Gemini");
  const vo = ["1 Baseline", "2 Losses", "3 Prioritize", "4 Projects", "5 Validate savings", "6 Sustain"];
  vo.forEach((t, i) => {
    s.addShape(i === 0 ? S.PENTAGON : S.CHEVRON, { x: 0.6 + i * 2.02, y: 1.45, w: 2.1, h: 0.65, fill: { color: i === 0 ? K.amber : K.light }, line: { color: K.white, width: 1.5 }, objectName: "VO step " + (i + 1) });
    txt(s, t, { x: 0.6 + i * 2.02 + 0.3, y: 1.45, w: 1.6, h: 0.65, fontSize: 12, bold: true, color: K.navy, align: "center", valign: "middle", margin: 0 });
  });
  box(s, 0.6, 2.35, 7.8, 4.45, { fill: K.light });
  iconOnly(s, "FaDatabase", 0.85, 2.55, 0.42, HEX.teal);
  txt(s, "Collect for each Gemini — from Redzone and validated production records", { x: 1.4, y: 2.5, w: 6.9, h: 0.5, fontSize: 14, bold: true, color: K.navy, valign: "middle", margin: 0 });
  const dataItems = ["OEE", "Availability", "Performance", "Quality", "Downtime", "Minor stops", "Changeover time", "Scrap", "Rework", "Production output", "Planned vs actual production", "Labor hours", "Maintenance losses", "Product mix", "Run hours"];
  dataItems.forEach((t, i) => {
    const x = 0.85 + Math.floor(i / 5) * 2.5, y = 3.2 + (i % 5) * 0.68;
    iconOnly(s, "FaCheckCircle", x, y + 0.1, 0.3, HEX.teal);
    txt(s, t, { x: x + 0.4, y, w: 2.05, h: 0.5, fontSize: 13, color: K.ink, valign: "middle", margin: 0 });
  });
  box(s, 8.7, 2.35, 4.0, 4.45, { fill: K.amber, tr: 82 });
  iconOnly(s, "FaExclamationTriangle", 8.9, 2.55, 0.42, HEX.navy);
  txt(s, "Know your data limits", { x: 9.45, y: 2.5, w: 3.1, h: 0.5, fontSize: 15, bold: true, color: K.navy, valign: "middle", margin: 0 });
  txt(s, bullets(["Manually entered downtime reasons may be incomplete or miscoded", "Short stops may be under-recorded", "Product mix shifts distort comparisons", "Agree the comparison period with Finance up front", "Document every assumption"], { space: 6 }), { x: 8.9, y: 3.15, w: 3.65, h: 3.55, fontSize: 13 });
  notes(s, `PURPOSE: Explain VO Step 1 — establishing a credible baseline.
  TIMING: 5 minutes.
  THE SIX VO STEPS (top ribbon): Baseline → Losses → Prioritize → Projects → Validate savings → Sustain.
  STEP 1: For each Gemini, collect 4–8 weeks of available performance data, with an agreed comparison period. Use Redzone and validated production records where available.
  DATA TO COLLECT: OEE and its three factors (availability, performance, quality), downtime, minor stops, changeover time, scrap, rework, production output, planned vs actual production, labor hours, maintenance losses, product mix and run hours.
  DATA-QUALITY LIMITATIONS: Be explicit. Manually entered downtime or production information can be incomplete, late or miscoded; short stops are often under-recorded; product mix changes can make before/after comparisons misleading. Write these limitations down in the baseline report.
  FINANCE: Randy reviews the baseline assumptions before any improvement work is measured against them.
  NO DATA IN THIS DECK: No actual Gemini performance results are shown or implied. The baseline will be built during Days 1–30.
  ASK: "How confident are we in the downtime reasons currently entered for the Gemini? Who enters them, and when?"`);

  }
  if (VO) {
  // ======================= 33. STEP 2 LOSSES =======================
  s = content("PART 7 · VO STEP 2", "Identify the losses: Six Big Losses and who influences them");
  const groups = [["AVAILABILITY", K.coral, [["Equipment failures", "Unplanned stops & breakdowns", "AM · PM"], ["Setup & adjustments", "Changeovers, startup adjustments", "AM (centerline) · Process · Scheduling"]]], ["PERFORMANCE", K.amber, [["Idling & minor stops", "Short stops, jams, blocked/starved", "AM · Process"], ["Reduced speed", "Running below standard rate", "AM (centerline) · PM · Process"]]], ["QUALITY", K.blue, [["Process defects", "Scrap and rework during steady run", "Quality · AM · Process"], ["Reduced yield", "Startup and changeover rejects", "AM (centerline) · Quality · Scheduling"]]]];
  groups.forEach(([g, col, ls], i) => {
    const x = 0.6 + i * 2.25;
    box(s, x, 1.5, 2.1, 0.5, { fill: col, r: 0.1 });
    txt(s, g, { x, y: 1.5, w: 2.1, h: 0.5, fontSize: 12, bold: true, color: K.white, align: "center", valign: "middle", charSpacing: 1 });
    ls.forEach(([h, d, inf], j) => {
      const y = 2.15 + j * 2.25;
      box(s, x, y, 2.1, 2.1, { fill: col, tr: 88 });
      txt(s, h, { x: x + 0.12, y: y + 0.1, w: 1.9, h: 0.55, fontSize: 13, bold: true, color: K.navy, margin: 0 });
      txt(s, d, { x: x + 0.12, y: y + 0.65, w: 1.9, h: 0.6, fontSize: 11, color: K.ink, margin: 0 });
      txt(s, [{ text: "Influenced by: ", options: { bold: true, color: K.teal, breakLine: true } }, { text: inf }], { x: x + 0.12, y: y + 1.25, w: 1.9, h: 0.8, fontSize: 10, color: K.ink, margin: 0 });
    });
  });
  const pl = ["Minor stops", "Changeovers", "Equip. failures", "Scrap & rework", "Reduced speed", "Startup losses"];
  const pv = [120, 90, 70, 45, 30, 20];
  let acc = 0; const tot = pv.reduce((a, b) => a + b, 0); const cum = pv.map((v) => Math.round(((acc += v) / tot) * 100));
  s.addChart([
    { type: pres.charts.BAR, data: [{ name: "Loss (illustrative units)", labels: pl, values: pv }], options: { chartColors: [HEX.teal], barGapWidthPct: 45 } },
    { type: pres.charts.LINE, data: [{ name: "Cumulative %", labels: pl, values: cum }], options: { chartColors: [HEX.amber], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 6, secondaryValAxis: true, secondaryCatAxis: true } },
  ], {
    x: 7.45, y: 1.85, w: 5.25, h: 4.45, showTitle: true, title: "Example loss Pareto", titleFontSize: 14, titleColor: HEX.navy, titleFontFace: "+mn-lt",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendFontFace: "+mn-lt", legendColor: HEX.ink,
    catAxisLabelColor: HEX.slate, catAxisLabelFontSize: 10, catAxisLabelFontFace: "+mn-lt",
    valAxes: [{ showValAxisTitle: true, valAxisTitle: "Loss (units)", valAxisTitleFontSize: 10, valAxisTitleColor: HEX.slate, valAxisLabelColor: HEX.slate, valAxisLabelFontSize: 10, valAxisMinVal: 0, valGridLine: { color: HEX.grid, size: 0.5 } }, { showValAxisTitle: true, valAxisTitle: "Cumulative %", valAxisTitleFontSize: 10, valAxisTitleColor: HEX.slate, valAxisLabelColor: HEX.slate, valAxisLabelFontSize: 10, valAxisMinVal: 0, valAxisMaxVal: 100, valGridLine: { style: "none" } }],
    catAxes: [{ catAxisTitle: "" }, { catAxisHidden: true }],
  });
  badge(s, "ILLUSTRATIVE — NOT GEMINI DATA", 7.45, 1.5, 3.2);
  txt(s, "Replace with each Gemini’s real Pareto from the baseline.", { x: 7.45, y: 6.4, w: 5.25, h: 0.35, fontSize: 10, italic: true, color: K.slate, margin: 0 });
  notes(s, `PURPOSE: Explain the Six Big Losses and how to separate losses by who can influence them.
  TIMING: 7 minutes.
  SIX BIG LOSSES OF TPM, grouped by the OEE factor they reduce:
  • Availability: equipment failures (breakdowns) and setup & adjustments (changeovers, startup adjustments).
  • Performance: idling & minor stops (jams, short stops, blocked/starved) and reduced speed.
  • Quality: process defects (scrap and rework during steady running) and reduced yield (startup/changeover rejects).
  WHO CAN INFLUENCE EACH LOSS: Separate losses that can be influenced by AM, PM, Quality, process improvements and scheduling. This avoids assigning, for example, a scheduling-driven changeover loss to the AM team.
  PARETO (right): ILLUSTRATIVE ONLY — the values are invented to show how a Pareto works. They are NOT Gemini data. Each cell builds its own Pareto from the baseline. The bars show loss size, largest first; the line shows cumulative percentage — typically a few categories drive most of the loss.
  ASK: "Without looking at data, what do you think is the biggest loss on your Gemini?" Then: "We'll check that against the baseline — intuition is often wrong."`);

  }
  if (VO) {
  // ======================= 34. STEP 3 PRIORITIZE =======================
  s = content("PART 7 · VO STEP 3", "Prioritize: safety gate first, then impact versus ease");
  box(s, 0.6, 1.5, 3.3, 5.3, { fill: K.coral });
  iconCircle(s, "FaLock", 1.75, 1.75, 1.0, K.white, HEX.coral);
  txt(s, "GATE 1", { x: 0.8, y: 2.95, w: 2.9, h: 0.3, fontSize: 12, bold: true, color: K.white, align: "center", charSpacing: 2, margin: 0 });
  txt(s, "Safety, food safety & compliance", { x: 0.8, y: 3.25, w: 2.9, h: 0.75, fontSize: 17, bold: true, color: K.white, align: "center", margin: 0 });
  txt(s, "Required work is done regardless of financial return. These items are never traded against savings.", { x: 0.85, y: 4.15, w: 2.8, h: 1.6, fontSize: 13, color: K.white, align: "center", margin: 0 });
  arrow(s, 3.95, 4.15, 4.35, 4.15, K.slate, 2.5);
  const mx = 4.9, my = 1.5, mw = 4.6, mh = 4.7;
  const quads = [["Major projects", "Plan & resource", K.blue, 0, 0], ["Quick wins", "Do first", K.teal, 1, 0], ["Deprioritize", "Revisit later", K.slate, 0, 1], ["Fill-ins", "When capacity allows", K.amber, 1, 1]];
  quads.forEach(([h, d, col, cx2, cy2]) => {
    const x = mx + cx2 * (mw / 2 + 0.05), y = my + cy2 * (mh / 2 + 0.05);
    box(s, x, y, mw / 2 - 0.05, mh / 2 - 0.05, { fill: col, tr: cx2 === 1 && cy2 === 0 ? 0 : 80 });
    const dk = cx2 === 1 && cy2 === 0;
    txt(s, [{ text: h, options: { bold: true, fontSize: 17, color: dk ? K.white : K.navy, breakLine: true } }, { text: d, options: { fontSize: 13, color: dk ? K.white : K.ink } }], { x: x + 0.15, y: y + 0.15, w: mw / 2 - 0.35, h: 1.2, margin: 0 });
  });
  txt(s, "IMPACT  →", { x: 4.4, y: 3.85, w: 0.4, h: 1.4, fontSize: 11, bold: true, color: K.slate, vert: "vert270", align: "center", margin: 0 });
  txt(s, "EASE OF IMPLEMENTATION  →", { x: mx, y: 6.3, w: mw, h: 0.3, fontSize: 11, bold: true, color: K.slate, align: "center", margin: 0 });
  box(s, 9.8, 1.5, 2.9, 5.3, { fill: K.light });
  txt(s, "SCORING CRITERIA", { x: 10.0, y: 1.65, w: 2.6, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  txt(s, bullets([{ text: "Safety & food-safety risk (gate)", options: { bold: true } }, "Financial impact", "Frequency", "Severity", "Ease of implementation", "Available resources", "Expected sustainability"], { space: 5 }), { x: 10.0, y: 2.05, w: 2.6, h: 3.3, fontSize: 13 });
  txt(s, "Score each 1–5. The cell agrees scores with CI and Maintenance.", { x: 10.0, y: 5.6, w: 2.6, h: 1.0, fontSize: 12, italic: true, color: K.slate, margin: 0 });
  notes(s, `PURPOSE: Show how to prioritize improvement opportunities practically.
  TIMING: 5 minutes.
  GATE 1 — SAFETY, FOOD SAFETY & COMPLIANCE: Anything required for safety, food safety or regulatory compliance is done regardless of financial return. These items are NEVER traded against financial benefit and never "scored out" of the plan.
  THEN THE MATRIX: Impact (financial impact, frequency, severity) vs Ease (ease of implementation, available resources, expected sustainability).
  • Quick wins (high impact, easy): do first — ideal starter projects for supervisor-led cells.
  • Major projects (high impact, hard): plan and resource with leadership.
  • Fill-ins (low impact, easy): do when capacity allows.
  • Deprioritize (low impact, hard): revisit later.
  SCORING: Score each criterion 1–5. The cell agrees scores together with CI and Maintenance, so priorities are owned, not imposed.
  ASK: "What makes an improvement 'sustainable'?" (Answer: it changes a standard, not just a one-time fix; it does not depend on one person.)`);

  }
  if (VO) {
  // ======================= 35. STEP 4 PROJECTS =======================
  s = content("PART 7 · VO STEP 4", "Develop projects: start small, prove ownership, then scale");
  const waves = [["WAVE 1", "Starter projects — cell-led (Days 31–90)", ["Reduce recurring minor stops", "Improve changeover consistency", "Reduce adjustment losses", "Improve CILT completion", "Eliminate repeat equipment defects", "Improve centerline adherence", "Reduce product scrap", "Reduce unplanned downtime"], 2.55, K.teal], ["WAVE 2", "Expand — after measurable results", ["Structured changeover reduction (SMED)", "PM optimization with Maintenance", "Quality-loss RCA (DMAIC)", "Cross-shift standard harmonization"], 2.05, K.amber], ["WAVE 3", "Scale", ["Standardize best practices across all three Geminis", "Next-wave AM + VO deployment to other lines", "Experienced supervisors coach new cells"], 1.55, K.navy]];
  waves.forEach(([w, sub, items, y, col], i) => {
    const x = 0.6 + i * 4.1, h = 6.8 - y;
    box(s, x, y, 3.85, h, { fill: col, tr: i === 2 ? 0 : 85 });
    const dk = i === 2;
    txt(s, w, { x: x + 0.25, y: y + 0.2, w: 3.4, h: 0.3, fontSize: 12, bold: true, color: dk ? K.amber : col, charSpacing: 2, margin: 0 });
    txt(s, sub, { x: x + 0.25, y: y + 0.5, w: 3.4, h: 0.65, fontSize: 15, bold: true, color: dk ? K.white : K.navy, margin: 0 });
    txt(s, bullets(items, { space: 4 }), { x: x + 0.25, y: y + 1.2, w: 3.4, h: h - 1.35, fontSize: 13, color: dk ? K.white : K.ink });
  });
  txt(s, "Gate to next wave: CILT and centerline adherence sustained + first Finance-validated result", { x: 0.6, y: 1.5, w: 7.9, h: 0.4, fontSize: 12, italic: true, bold: true, color: K.slate, margin: 0 });
  notes(s, `PURPOSE: Show how improvement projects grow in complexity as capability grows.
  TIMING: 4 minutes.
  WAVE 1 (Days 31–90): Manageable, cell-led projects that build confidence: reduce recurring minor stops; improve changeover consistency; reduce adjustment losses; improve CILT completion; eliminate repeat equipment defects; improve centerline adherence; reduce product scrap; reduce unplanned downtime. Each cell picks 1–2 based on its own Pareto, not all eight.
  WAVE 2: More complex projects once supervisors have demonstrated ownership and measurable results — e.g., structured changeover reduction (SMED), PM optimization with Maintenance, DMAIC-based quality-loss RCA.
  WAVE 3: Scale — standardize across all three Geminis and deploy to the next lines, with experienced supervisors coaching new cells.
  GATE: Move to the next wave only when the basics are sustained (CILT and centerline adherence) and at least one result has been validated by Finance.
  ASK: "Which Wave 1 project would make the biggest difference to your team's daily frustration?"`);

  }
  if (VO) {
  // ======================= 36. STEP 5 SAVINGS =======================
  s = content("PART 7 · VO STEP 5", "Validate savings with Finance (Randy)");
  const sv = ["Baseline", "Improvement", "Comparable results", "Financial calculation", "Finance approval", "Sustained results"];
  sv.forEach((t, i) => {
    s.addShape(i === 0 ? S.PENTAGON : S.CHEVRON, { x: 0.6 + i * 2.02, y: 1.45, w: 2.1, h: 0.85, fill: { color: i === 4 ? K.amber : K.navy }, line: { color: K.white, width: 1.5 }, objectName: "Savings step " + (i + 1) });
    txt(s, t, { x: 0.6 + i * 2.02 + 0.3, y: 1.45, w: 1.6, h: 0.85, fontSize: 11, bold: true, color: i === 4 ? K.navy : K.white, align: "center", valign: "middle", margin: 0 });
  });
  const st = [["Estimated savings", "Projected before or early in a project", "No — pipeline only"], ["Validated savings", "Calculated from comparable before/after data; approved by Finance", "Yes"], ["Cost avoidance", "Prevents a future cost not in current spend (e.g., avoided failure)", "Report separately"], ["Actual realized savings", "Visible in actual financial results", "Yes — confirmed"], ["One-time vs recurring", "One-time counted once; recurring annualized only per Finance rules", "Per Finance rules"]];
  const stData = [[hdr("Savings type"), hdr("Meaning"), hdr("Counts as validated?")]].concat(st.map((r, i) => [{ text: r[0], options: { bold: true, color: K.navy, fill: { color: i % 2 ? K.white : K.light } } }, { text: r[1], options: { fill: { color: i % 2 ? K.white : K.light } } }, { text: r[2], options: { bold: true, fill: { color: i % 2 ? K.white : K.light }, color: r[2].startsWith("Yes") ? K.green : K.ink } }]));
  s.addTable(stData, { x: 0.6, y: 2.6, w: 8.0, colW: [2.1, 4.1, 1.8], rowH: 0.7, fontSize: 12, color: K.ink, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: HEX.grid }, margin: 0.08 });
  box(s, 8.9, 2.6, 3.8, 4.2, { fill: K.navy });
  iconOnly(s, "FaBalanceScale", 9.1, 2.8, 0.45, HEX.amber);
  txt(s, "Avoid double-counting", { x: 9.65, y: 2.78, w: 2.95, h: 0.5, fontSize: 15, bold: true, color: K.white, valign: "middle", margin: 0 });
  txt(s, bullets(["One improvement, one savings record", "Map each saving to a single loss category", "If two projects touch the same loss, Finance allocates", "Recalculate when mix or volume changes", { text: "No projected benefit is reported as achieved", options: { bold: true, color: K.amber } }], { space: 6 }), { x: 9.1, y: 3.45, w: 3.45, h: 3.25, fontSize: 13, color: K.white });
  notes(s, `PURPOSE: Define how savings are validated so results are credible.
  TIMING: 5 minutes.
  WORKFLOW: Baseline → Improvement → Comparable results (same products, similar volumes, agreed period) → Financial calculation (using Finance-approved rates) → Finance approval (Randy) → Sustained results (the gain holds over the agreed period).
  RANDY'S ROLE (Finance): Review baseline assumptions, validate financial calculations, confirm realized benefits and prevent double-counting.
  DEFINITIONS — use these words consistently:
  • Estimated savings — a projection; pipeline only.
  • Validated savings — calculated from comparable before/after data and approved by Finance.
  • Cost avoidance — prevents a future cost that is not in current spend; reported separately.
  • Actual realized savings — visible in actual financial results.
  • One-time vs recurring — one-time benefits are counted once; recurring benefits are annualized only per Finance rules.
  DOUBLE-COUNTING: If a CILT project and a centerline project both reduce minor stops on Gemini 2, the reduction is counted once. Finance allocates.
  CRITICAL RULE: Do not present projected benefits as achieved results — in WOr reviews, dashboards or anywhere else.`);

  }
  if (VO) {
  // ======================= 136. SAVINGS EXERCISE =======================
  s = dark("INTERACTIVE EXERCISE · 20 MINUTES", "Savings or not? Classify the benefit");
  const sc = [["A", "Cell 2 projects a large downtime reduction from a new CILT point implemented last week."], ["B", "Minor stops on Gemini 1 stayed below the agreed baseline for the agreed period, with comparable product mix. Finance reviewed and approved the calculation."], ["C", "A worn gearbox was replaced before it failed. A failure would have caused a long unplanned outage."], ["D", "The same scrap reduction on Gemini 3 is claimed by both the centerline project and the changeover project."]];
  sc.forEach(([l, t], i) => {
    const x = 0.6 + (i % 2) * 4.3, y = 1.5 + Math.floor(i / 2) * 2.4;
    box(s, x, y, 4.1, 2.2, { fill: K.white });
    numCircle(s, l, x + 0.2, y + 0.2, 0.6, K.amber, 18, K.navy);
    txt(s, "SCENARIO " + l, { x: x + 0.95, y: y + 0.27, w: 2.8, h: 0.45, fontSize: 12, bold: true, color: K.slate, charSpacing: 2, valign: "middle", margin: 0 });
    txt(s, t, { x: x + 0.2, y: y + 0.9, w: 3.7, h: 1.25, fontSize: 13, color: K.ink, margin: 0 });
  });
  box(s, 9.3, 1.5, 3.4, 4.6, { fill: K.teal });
  txt(s, "CLASSIFY AS", { x: 9.5, y: 1.65, w: 3, h: 0.3, fontSize: 11, bold: true, color: K.white, charSpacing: 2, margin: 0 });
  ["Estimated", "Validated", "Cost avoidance", "Realized", "Double-count risk"].forEach((t, i) => {
    box(s, 9.5, 2.05 + i * 0.6, 3.0, 0.48, { fill: K.white, r: 0.12 });
    txt(s, t, { x: 9.5, y: 2.05 + i * 0.6, w: 3.0, h: 0.48, fontSize: 14, bold: true, color: K.navy, align: "center", valign: "middle" });
  });
  txt(s, "Then: what would make it validated?", { x: 9.5, y: 5.2, w: 3.05, h: 0.8, fontSize: 13, bold: true, color: K.white, margin: 0 });
  txt(s, "All scenarios are hypothetical. No real Gemini results or dollar values are implied.", { x: 0.6, y: 6.4, w: 12.1, h: 0.4, fontSize: 13, italic: true, color: K.light, margin: 0 });
  notes(s, `PURPOSE: Practice using the savings vocabulary correctly — the most common source of disputes with Finance.
  TIMING: 20 minutes (2 min setup, 10 min tables, 8 min debrief).
  SETUP: Tables classify each scenario and answer "what evidence would make it validated?" All scenarios are hypothetical.
  MODEL ANSWERS:
  A — ESTIMATED. Implemented last week; no comparable results yet. It belongs in the pipeline only. To validate: run for the agreed period, compare to baseline with comparable mix, calculate with Finance-approved rates, get Randy's approval.
  B — VALIDATED. Comparable before/after data, agreed period, Finance approval. It becomes REALIZED when it shows up in actual financial results.
  C — COST AVOIDANCE. The cost was never in current spend. Report it separately from savings, using Finance's rules for avoided failures.
  D — DOUBLE-COUNT RISK. One loss reduction, two claims. Count it once; Finance allocates between projects (or credits it jointly).
  DEBRIEF: "Why does this matter?" — credibility. One overstated claim makes leadership doubt every result the program reports. Never present projected benefits as achieved.`);

  }
  if (VO) {
  // ======================= 37. STEP 6 SUSTAIN =======================
  s = content("PART 7 · VO STEP 6", "Sustain and expand: prove it holds, then spread it");
  const sus = [["FaClipboardList", "Documented standards"], ["FaGraduationCap", "Training & competency"], ["FaClipboardCheck", "Audits"], ["FaChartLine", "Periodic performance reviews"], ["FaLayerGroup", "Expand to next wave"]];
  sus.forEach(([ic, t], i) => {
    const x = 0.6 + i * 2.5;
    box(s, x, 1.6, 2.2, 2.1, { fill: K.light });
    iconCircle(s, ic, x + 0.7, 1.8, 0.8, i === 4 ? K.amber : K.teal);
    txt(s, t, { x: x + 0.1, y: 2.75, w: 2.0, h: 0.85, fontSize: 14, bold: true, color: K.navy, align: "center", valign: "middle", margin: 0 });
    if (i < 4) arrow(s, x + 2.23, 2.65, x + 2.47, 2.65, K.slate, 2);
  });
  txt(s, "PROOF OF SUSTAINMENT", { x: 0.6, y: 4.05, w: 8, h: 0.3, fontSize: 11, bold: true, color: K.teal, charSpacing: 2, margin: 0 });
  [["FaUsers", "Across shifts", "The same result on every shift, not just the supervisor’s shift"], ["FaSyncAlt", "Across products", "Holds through product changes and changeovers"], ["FaCalendarAlt", "Over time", "Holds for an agreed period (TBD with Finance) before it is declared sustained"]].forEach(([ic, h, d], i) => {
    const x = 0.6 + i * 4.1;
    box(s, x, 4.45, 3.85, 2.35, { fill: K.navy });
    iconOnly(s, ic, x + 0.25, 4.65, 0.5, HEX.amber);
    txt(s, h, { x: x + 0.9, y: 4.65, w: 2.8, h: 0.5, fontSize: 17, bold: true, color: K.white, valign: "middle", margin: 0 });
    txt(s, d, { x: x + 0.25, y: 5.35, w: 3.4, h: 1.3, fontSize: 13, color: K.light, margin: 0 });
  });
  notes(s, `PURPOSE: Show how gains are sustained and expanded.
  TIMING: 3 minutes.
  SUSTAIN CYCLE: Documented standards (updated CILT, centerline, SOPs) → training and verified competency → audits (e.g., layered process audits by supervisors, managers and CI) → periodic performance reviews (WOr) → expand to the next wave.
  PROOF OF SUSTAINMENT: Demonstrate that the improvement continues across shifts and product changes, and over an agreed period (to be set with Finance — e.g., several weeks of stable results).
  ASK: "What usually causes an improvement to fade on our lines?" (Typical answers: people change, the standard isn't updated, nobody checks.) Each of those is addressed by the cycle on this slide.`);

  }
  if (true) {
  // ======================= 38. RACI =======================
  section("Roles & Roadmap");
  s = content("PART 8 · ROLES & RESPONSIBILITIES", AM ? "AM deployment responsibility matrix (RACI)" : "VO deployment responsibility matrix (RACI)");
  const rh = AM ? ["Activity", "Supervisor", "Production", "Maintenance", "Quality", "CI Manager", "Plant Leadership"] : ["Activity", "Supervisor", "Production", "Maintenance", "Quality", "CI Manager", "Finance (Randy)", "Plant Leadership"];
  const raci = AM ? [
    ["Lead cell & conduct daily huddles", "A/R", "R", "C", "C", "C", "I"],
    ["Create & approve safety map (EHS approves)", "A", "R", "R", "C", "C", "I"],
    ["Develop CILT standards", "A", "R", "R", "C", "C", "I"],
    ["Execute approved AM / CILT tasks", "A", "R", "C", "I", "I", "–"],
    ["Technical repairs & PM", "I", "I", "A/R", "C", "I", "–"],
    ["Define & validate centerlines", "A", "C", "R", "R", "C", "I"],
    ["Follow centerline standards", "A", "R", "C", "C", "I", "–"],
    ["Detect, tag & report defects", "A", "R", "R", "R", "I", "–"],
    ["Resolve equipment abnormalities", "C", "C", "A/R", "C", "I", "–"],
    ["Product disposition & quality verification", "C", "I", "C", "A/R", "I", "I"],
    ["Root-cause analysis of repeat defects", "A", "R", "R", "R", "C", "I"],
    ["Methodology, templates & tracking", "C", "I", "I", "I", "A/R", "I"],
    ["Resources & barrier removal", "C", "–", "C", "C", "C", "A/R"],
    ["Present AM results at WOr", "A/R", "R", "C", "C", "C", "I"],
  ] : [
    ["Lead cell & VO cadence", "A/R", "R", "C", "C", "C", "–", "I"],
    ["Provide accurate downtime & production data", "A", "R", "R", "R", "C", "–", "I"],
    ["Baseline data & loss analysis", "A", "C", "C", "C", "R", "C", "I"],
    ["Review baseline assumptions", "C", "–", "–", "–", "R", "A", "I"],
    ["Prioritize opportunities", "A", "C", "C", "C", "R", "C", "I"],
    ["Lead improvement projects", "A/R", "R", "R", "R", "C", "C", "I"],
    ["Measure results vs baseline", "A", "I", "C", "C", "R", "C", "I"],
    ["Validate financial benefits", "C", "–", "–", "–", "R", "A", "I"],
    ["Prevent double-counting", "I", "–", "–", "–", "R", "A", "I"],
    ["Sustain: standards, audits, reviews", "A", "R", "R", "R", "C", "–", "I"],
    ["Methodology, templates & tracking", "C", "I", "I", "I", "A/R", "I", "I"],
    ["Resources & barrier removal", "C", "–", "C", "C", "C", "–", "A/R"],
    ["Present VO results at WOr", "A/R", "R", "C", "C", "C", "C", "I"],
  ];
  const cellFill = (v) => (v === "A/R" ? K.teal : v === "A" ? K.navy : v === "R" ? K.teal : v === "C" ? K.light : K.white);
  const cellTr = (v) => (v === "R" ? 60 : 0);
  const raciData = [rh.map((t) => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, fontSize: 12, align: t === "Activity" ? "left" : "center" } }))].concat(raci.map((r) => r.map((v, j) => j === 0 ? { text: v, options: { bold: true, color: K.navy } } : { text: v, options: { align: "center", bold: true, color: v === "A" || v === "A/R" ? K.white : K.navy, fill: { color: cellFill(v), transparency: cellTr(v) } } })));
  const roleW = (12.1 - 4.2) / (rh.length - 1);
  s.addTable(raciData, { x: 0.6, y: 1.42, w: 12.1, colW: [4.2].concat(Array(rh.length - 1).fill(roleW)), rowH: 0.35, fontSize: 11, color: K.ink, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: HEX.grid }, margin: [0.02, 0.06, 0.02, 0.06] });
  txt(s, "R = Responsible   A = Accountable   C = Consulted   I = Informed   ·   Draft — validate at kickoff." + (AM ? " EHS approves safety maps and safety-critical tasks." : " Finance (Randy) is Accountable for every reported saving."), { x: 0.6, y: 6.72, w: 12.1, h: 0.3, fontSize: 10, color: K.slate, margin: 0 });
  notes(s, AM ? `PURPOSE: The AM deployment responsibility matrix — who is Responsible, Accountable, Consulted and Informed for each key AM activity.
  TIMING: 6 minutes (leadership kickoff: walk the A column for each role).
  READING THE MATRIX: R = does the work; A = owns the outcome and signs off; C = consulted before decisions; I = kept informed.
  KEY POINTS:
  • Supervisors are Accountable for most cell activities — this is what "equipment owner" means.
  • Maintenance is Accountable and Responsible for technical repairs and PM, and Responsible for validating centerlines (with Quality).
  • Quality is Accountable for product disposition and quality verification.
  • The CI Manager owns the methodology and tracking but does not own the cell's work.
  • Plant Leadership is Accountable for resources and barrier removal.
  • EHS (not shown as a column) approves safety maps and any safety-critical task assignment.
  DRAFT: Validate with all functions at kickoff.` : `PURPOSE: The VO deployment responsibility matrix.
  TIMING: 6 minutes.
  KEY POINTS:
  • Supervisors are Accountable for their cell's baseline, priorities and projects.
  • Production, Maintenance and Quality are Responsible for accurate data — the baseline is only as good as the downtime and scrap records behind it.
  • The CI Manager is Responsible for the baseline analysis and measurement method, coaching the cell rather than doing the projects.
  • Finance (Randy) is Accountable for reviewing baseline assumptions, validating financial benefits and preventing double-counting.
  • Plant Leadership is Accountable for resources and barrier removal.
  DRAFT: Validate with all functions at kickoff.`);

  }
  if (true) {
  // ======================= 39. ROLE CARDS =======================
  s = content("PART 8 · ROLES & RESPONSIBILITIES", "What each role owns");
  const rc = AM ? [
    ["FaUserTie", "Supervisors", K.navy, ["Lead the assigned Gemini cell", "Conduct daily huddles", "Reinforce standards", "Monitor CILT & centerline adherence", "Escalate defects", "Coach employees", "Lead AM improvement actions", "Present results to leadership"]],
    ["FaUsers", "Production", K.teal, ["Execute approved AM tasks", "Identify abnormal conditions", "Follow centerline standards", "Report defects", "Participate in RCA and improvements"]],
    ["FaWrench", "Maintenance", K.amber, ["Own technical repairs and PM", "Validate maintenance-related standards", "Support centerline development", "Resolve equipment abnormalities", "Train operators on approved activities"]],
    ["FaShieldAlt", "Quality", K.blue, ["Define product requirements", "Support centerline validation", "Investigate quality losses", "Verify product disposition and quality improvement"]],
    ["FaChalkboardTeacher", "Continuous Improvement Manager", K.teal, ["Own the AM deployment methodology", "Coach supervisors and teams; facilitate problem-solving", "Develop standard templates", "Track deployment progress; report program progress"]],
    ["FaFlag", "Plant Leadership", K.coral, ["Provide resources", "Remove barriers", "Review progress", "Reinforce accountability"]],
  ] : [
    ["FaUserTie", "Supervisors", K.navy, ["Own the cell's loss baseline", "Lead loss analysis & prioritization", "Lead improvement projects", "Track results vs baseline", "Coach team problem-solving", "Present results to leadership"]],
    ["FaUsers", "Production", K.teal, ["Record accurate downtime & production data", "Participate in loss analysis", "Execute project actions", "Sustain new standards"]],
    ["FaWrench", "Maintenance", K.amber, ["Analyze equipment-related losses", "Lead PM-related countermeasures", "Validate technical solutions", "Record maintenance losses accurately"]],
    ["FaShieldAlt", "Quality", K.blue, ["Quantify scrap & rework losses", "Investigate quality losses", "Verify product impact of improvements", "Verify product disposition"]],
    ["FaChalkboardTeacher", "Continuous Improvement Manager", K.teal, ["Own the VO methodology", "Coach supervisors; facilitate problem-solving", "Support VO analysis and savings measurement", "Track and report program progress"]],
    ["FaDollarSign", "Finance — Randy", K.green, ["Review baseline assumptions", "Validate financial calculations", "Confirm realized benefits", "Prevent double-counting"]],
    ["FaFlag", "Plant Leadership", K.coral, ["Provide resources", "Remove barriers", "Review progress", "Reinforce accountability"]],
  ];
  const nb = rc.length - 4, bw = (12.1 - (nb - 1) * 0.2) / nb;
  rc.forEach(([ic, h, col, items], i) => {
    const top = i < 4;
    const w = top ? 2.9 : bw, gap = top ? 0.167 : 0.2;
    const x = 0.6 + (top ? i : i - 4) * (w + gap), y = top ? 1.45 : 4.35, hh = top ? 2.75 : 2.45;
    box(s, x, y, w, hh, { fill: K.light });
    iconCircle(s, ic, x + 0.15, y + 0.15, 0.5, col);
    txt(s, h, { x: x + 0.75, y: y + 0.15, w: w - 0.85, h: 0.5, fontSize: 13, bold: true, color: K.navy, valign: "middle", margin: 0 });
    txt(s, bullets(items, { space: 1 }), { x: x + 0.12, y: y + 0.72, w: w - 0.22, h: hh - 0.8, fontSize: 11 });
  });
  notes(s, `PURPOSE: Detail what each role owns in the Gemini ${AM ? "AM" : "VO"} deployment.
  TIMING: 5 minutes.
  WALK EACH CARD and ask each function: "What do you need from the other roles to deliver your part?" Record commitments.
  ${AM ? "Finance is not shown here: Finance validation of savings belongs to the VO pillar (companion deck)." : "Finance (Randy) reviews baseline assumptions, validates financial calculations, confirms realized benefits and prevents double-counting."}`);

  }
  if (true) {
  // ======================= 40. GANTT =======================
  s = content("PART 9 · 90-DAY DEPLOYMENT ROADMAP", AM ? "90-day AM roadmap: foundation, implementation, sustainment" : "90-day VO roadmap: foundation, implementation, sustainment");
  const gx = 4.0, gw = 8.7, wk = gw / 13, gy = 1.45;
  [["FOUNDATION · DAYS 1–30", 1, 4, K.teal], ["IMPLEMENTATION · DAYS 31–60", 5, 8, K.amber], ["SUSTAINMENT · DAYS 61–90", 9, 13, K.navy]].forEach(([t, a, b, col]) => {
    box(s, gx + (a - 1) * wk + 0.02, gy, (b - a + 1) * wk - 0.04, 0.34, { fill: col, r: 0.05 });
    txt(s, t, { x: gx + (a - 1) * wk, y: gy, w: (b - a + 1) * wk, h: 0.34, fontSize: 10, bold: true, color: col === K.amber ? K.navy : K.white, align: "center", valign: "middle", margin: 0 });
  });
  for (let w = 1; w <= 13; w++) {
    txt(s, "W" + w, { x: gx + (w - 1) * wk, y: gy + 0.38, w: wk, h: 0.24, fontSize: 10, bold: true, color: K.slate, align: "center", margin: 0 });
    if (w > 1) line(s, gx + (w - 1) * wk, gy + 0.62, gx + (w - 1) * wk, 6.75, HEX.grid, 0.5);
  }
  const tasks = AM ? [
    ["Confirm cell ownership & kickoff", 1, 1, 0], ["Machine assessments", 1, 3, 0], ["Create safety maps", 2, 4, 0], ["Identify CILT points", 2, 4, 0], ["Establish defect registers", 2, 3, 0], ["Gather settings, OEM & PM records", 2, 4, 0], ["Supervisor coaching (ongoing)", 1, 13, 0],
    ["Launch approved CILT standards", 5, 6, 1], ["Validate centerline parameters", 5, 8, 1], ["Train employees", 5, 8, 1], ["Implement defect handling", 5, 6, 1], ["Integrate PM & Quality", 6, 8, 1], ["Start cell AM improvement actions", 7, 13, 1],
    ["Audit deployment", 9, 12, 2], ["Review defect closure", 9, 13, 2], ["Verify centerline adherence", 9, 12, 2], ["Review AM metric trends", 11, 12, 2], ["Present results (WOr)", 13, 13, 2], ["Next-wave AM plan", 12, 13, 2],
  ] : [
    ["Confirm cell ownership & kickoff", 1, 1, 0], ["Map Redzone data & reason codes", 1, 2, 0], ["Collect baseline data", 1, 4, 0], ["Data-quality review", 2, 3, 0], ["Agree baseline with Finance", 3, 4, 0], ["Supervisor coaching (ongoing)", 1, 13, 0],
    ["Build loss Paretos", 5, 6, 1], ["Separate losses by owner", 5, 6, 1], ["Prioritize opportunities", 6, 7, 1], ["Charter starter projects", 7, 7, 1], ["Execute starter projects", 7, 13, 1], ["Track results weekly vs baseline", 8, 13, 1],
    ["Sustainment checks (shifts, products)", 9, 13, 2], ["Compare performance to baseline", 11, 12, 2], ["Financial calculation", 12, 12, 2], ["Finance validation (Randy)", 12, 13, 2], ["Present results (WOr)", 13, 13, 2], ["Next-wave VO plan", 12, 13, 2],
  ];
  const rowH = Math.min(0.24, 4.4 / tasks.length), ty = gy + 0.68;
  const pc = [K.teal, K.amber, K.navy];
  tasks.forEach(([t, a, b, p], i) => {
    const y = ty + i * rowH;
    if (i % 2 === 0) box(s, 0.6, y, 12.1, rowH, { fill: K.light, rect: true });
    txt(s, t, { x: 0.65, y, w: 3.3, h: rowH, fontSize: 10, color: K.ink, valign: "middle", margin: 0 });
    box(s, gx + (a - 1) * wk + 0.04, y + 0.04, (b - a + 1) * wk - 0.08, rowH - 0.08, { fill: pc[p], r: 0.04 });
  });
  [4, 8, 13].forEach((w) => {
    s.addShape(S.DIAMOND, { x: gx + w * wk - 0.2, y: ty + tasks.length * rowH + 0.03, w: 0.2, h: 0.2, fill: { color: K.coral }, line: { color: K.white, width: 1 }, objectName: "WOr milestone W" + w });
  });
  txt(s, "◆ WOr leadership checkpoints (proposed W4, W8, W13)  ·  Durations are planning targets — confirm with cell leaders at kickoff", { x: 0.6, y: 6.78, w: 11.5, h: 0.26, fontSize: 10, color: K.slate, margin: 0 });
  notes(s, `PURPOSE: The detailed 90-day ${AM ? "AM" : "VO"} roadmap — the same plan for all three Geminis, run in parallel.
  TIMING: 6 minutes.
  ${AM ? "DAYS 1–30 FOUNDATION: confirm cell ownership; machine assessments; safety maps; CILT points; defect registers; gather current settings, OEM manuals and PM records; begin supervisor coaching (continues throughout).\n  DAYS 31–60 IMPLEMENTATION: launch approved CILT; validate centerline parameters; train employees; implement defect handling; integrate PM and Quality; start cell improvement actions.\n  DAYS 61–90 SUSTAINMENT: audit deployment; review defect closure; verify centerline adherence; review AM metric trends; present results; plan the next wave." : "DAYS 1–30 FOUNDATION: confirm cell ownership; map Redzone data and reason codes; collect 4–8 weeks of baseline data; data-quality review; agree baseline and comparison period with Finance; begin supervisor coaching.\n  DAYS 31–60 IMPLEMENTATION: build loss Paretos; separate losses by owner (AM, PM, Quality, process, scheduling); prioritize; charter and start starter projects; track results weekly.\n  DAYS 61–90 SUSTAINMENT: sustainment checks across shifts and products; compare performance to baseline; financial calculation and Finance validation; present results; plan the next wave."}
  MILESTONES: Red diamonds mark proposed WOr leadership checkpoints (Weeks 4, 8, 13).
  EDITABLE: Bars are standard shapes — drag to adjust once dates are confirmed.
  CAUTION: This roadmap shows planned activities only. It does not present projected benefits as achieved results.`);

  }
  if (true) {
  // ======================= 41. WEEKLY MILESTONES =======================
  s = content("PART 9 · 90-DAY DEPLOYMENT ROADMAP", "Weekly milestones, owners and evidence");
  const wm = AM ? [
    ["W1", "Kickoff; confirm cells, members, machine IDs, shifts", "Plant Leadership + CI", "Signed cell charter; contact list"],
    ["W2", "Machine assessments; safety-map walk; gather settings & PM records", "Supervisors + Maintenance + EHS", "Assessment checklist; draft safety map"],
    ["W3", "Hazard register; CILT point identification; defect register live", "Cell teams + EHS", "Hazard register; inspection-point map; register"],
    ["W4", "Safety maps approved; draft CILT cards; WOr #1", "Supervisor + EHS + CI", "Approved maps; draft CILT cards"],
    ["W5", "Launch approved CILT (daily / weekly); defect tagging live", "Supervisor + Production", "CILT checklists in use; tags in register"],
    ["W6", "Centerline parameter list & draft registers; PM/Quality links", "Maintenance + Quality", "Draft registers; work-order flow agreed"],
    ["W7", "Training & competency checks; first cell AM actions", "Supervisor + Maintenance", "Training records; action list"],
    ["W8", "Centerlines validated & posted; WOr #2", "Maintenance + Quality + Supervisor", "Validated registers; WOr deck"],
    ["W9", "Deployment audit #1; centerline adherence checks", "Supervisor + CI", "Audit results; adherence log"],
    ["W10", "Defect closure review; repeat-defect RCAs", "Cell teams", "Closure rate; RCA records"],
    ["W11", "AM metric trend review (CILT, centerline, defects)", "CI + Supervisor", "AM dashboard"],
    ["W12", "Deployment audit #2; next-wave AM draft", "Supervisor + CI", "Audit results; draft plan"],
    ["W13", "Final WOr: AM status, risks, next-wave plan", "Supervisors", "WOr deck; approved next-wave plan"],
  ] : [
    ["W1", "Kickoff; confirm cells, members, Redzone line mapping", "Plant Leadership + CI", "Signed cell charter; data map"],
    ["W2", "Pull 4–8 weeks of history; review reason codes", "CI + Supervisors", "Raw baseline extract; code list"],
    ["W3", "Data-quality review; fill gaps from production records", "CI + Production + Maintenance", "Data-quality notes"],
    ["W4", "Baseline & comparison period agreed with Finance; WOr #1", "Supervisor + CI + Finance", "Finance-reviewed baseline"],
    ["W5", "Loss Paretos per Gemini; Six Big Losses split", "CI + Supervisor", "Pareto charts"],
    ["W6", "Losses assigned by owner (AM/PM/Quality/process/scheduling)", "Cell teams", "Loss-owner map"],
    ["W7", "Prioritization matrix; starter projects chartered", "Supervisor + CI", "Scored matrix; project charters"],
    ["W8", "Projects under way; weekly tracking live; WOr #2", "Supervisors", "Tracker; WOr deck"],
    ["W9", "Sustainment checks across shifts", "Supervisor + CI", "Shift comparison"],
    ["W10", "Mid-point results review; adjust countermeasures", "Cell teams", "Review notes"],
    ["W11", "Performance vs baseline (comparable period)", "CI + Supervisor", "Before/after comparison"],
    ["W12", "Savings calculation & Finance review; next-wave draft", "CI + Finance (Randy)", "Finance validation sheet"],
    ["W13", "Final WOr: validated results, risks, next-wave plan", "Supervisors", "WOr deck; approved next-wave plan"],
  ];
  const wmData = [[hdr("Week"), hdr("Milestone"), hdr("Owner"), hdr("Evidence / deliverable")]].concat(wm.map((r, i) => {
    const ph = i < 4 ? K.teal : i < 8 ? K.amber : K.navy;
    const f = { fill: { color: i % 2 ? K.white : K.light } };
    return [{ text: r[0], options: { bold: true, color: ph === K.amber ? K.navy : K.white, fill: { color: ph }, align: "center" } }, { text: r[1], options: f }, { text: r[2], options: Object.assign({ color: K.navy, bold: true }, f) }, { text: r[3], options: f }];
  }));
  s.addTable(wmData, { x: 0.6, y: 1.42, w: 12.1, colW: [0.75, 5.0, 2.85, 3.5], rowH: 0.375, fontSize: 11, color: K.ink, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: HEX.grid }, margin: [0.02, 0.07, 0.02, 0.07] });
  notes(s, `PURPOSE: Weekly milestones with owners and evidence — the operational version of the Gantt.
  TIMING: 4 minutes (leadership kickoff); use as a weekly tracking tool afterwards.
  HOW TO USE: Review this table every week in the CI deployment meeting. A milestone is complete only when the evidence in the right-hand column exists.
  PHASE COLORS: Teal = Foundation (W1–4), Amber = Implementation (W5–8), Navy = Sustainment (W9–13).
  EVIDENCE: Evidence is objective (a document, record, audit result or approved deck) — not "we talked about it."
  ADJUST: Owners and timing are proposed. Confirm at kickoff. If one cell falls behind, the CI Manager escalates to plant leadership rather than doing the work for the cell.
  IMPORTANT: Results reviewed in Weeks 11–13 reflect real measured data only; nothing in this plan presents projected benefits as achieved.`);

  }
  if (AM) {
  // ======================= 42. HUDDLE =======================
  section("Daily Management & WOr");
  s = content("PART 10 · DAILY MANAGEMENT", "The daily huddle: nine topics, same order, every day");
  const hud = [["FaHardHat", "Safety", "Any incidents, near misses or hazards?"], ["FaShieldAlt", "Food quality", "Any holds, complaints or quality risks?"], ["FaCogs", "Machine condition", "How is the Gemini running?"], ["FaClipboardCheck", "CILT completion", "Was CILT done? What did it find?"], ["FaSlidersH", "Centerline abnormalities", "Anything out of range?"], ["FaTag", "Open defects", "What’s open, overdue, verified?"], ["FaChartBar", "Yesterday’s biggest loss", "What cost us most — and why?"], ["FaBullseye", "Today’s improvement action", "One action, one owner, today"], ["FaStar", "Recognition", "Who did it right? Say it out loud."]];
  hud.forEach(([ic, h, q], i) => {
    const x = 0.6 + (i % 3) * 2.6, y = 1.5 + Math.floor(i / 3) * 1.8;
    box(s, x, y, 2.45, 1.65, { fill: i < 2 ? K.coral : i === 8 ? K.amber : K.teal, tr: 86 });
    numCircle(s, i + 1, x + 0.15, y + 0.15, 0.45, K.navy, 13);
    iconOnly(s, ic, x + 1.85, y + 0.17, 0.4, HEX.navy);
    txt(s, h, { x: x + 0.15, y: y + 0.65, w: 2.2, h: 0.45, fontSize: 14, bold: true, color: K.navy, margin: 0 });
    txt(s, q, { x: x + 0.15, y: y + 1.1, w: 2.2, h: 0.5, fontSize: 11, color: K.ink, margin: 0 });
  });
  box(s, 8.5, 1.5, 4.2, 5.3, { fill: K.navy });
  iconOnly(s, "FaBrain", 8.75, 1.72, 0.5, HEX.amber);
  txt(s, "Why the same order, every day?", { x: 9.35, y: 1.68, w: 3.2, h: 0.6, fontSize: 15, bold: true, color: K.white, valign: "middle", margin: 0 });
  txt(s, bullets(["Spaced repetition moves standards from working memory into long-term memory", "A fixed structure frees attention for the real problems", "The visual board is a retrieval cue at the point of use", "Recognition reinforces the behavior you want repeated"], { space: 8 }), { x: 8.75, y: 2.5, w: 3.75, h: 3.3, fontSize: 13, color: K.white });
  box(s, 8.75, 5.85, 3.7, 0.7, { fill: K.amber, r: 0.1 });
  txt(s, "5–10 min · at the Gemini · start of shift", { x: 8.75, y: 5.85, w: 3.7, h: 0.7, fontSize: 13, bold: true, color: K.navy, align: "center", valign: "middle" });
  notes(s, `PURPOSE: Introduce the daily management model and the nine-topic huddle.
  TIMING: 5 minutes.
  THE NINE TOPICS (in order): 1) Safety, 2) Food quality, 3) Machine condition, 4) CILT completion, 5) Centerline abnormalities, 6) Open defects, 7) Yesterday's biggest loss, 8) Today's improvement action, 9) Recognition and reinforcement.
  NEUROSCIENCE LINK: Repeating important standards in short daily huddles is spaced repetition — the same principle from the Power of Repetition exercise. A fixed agenda becomes automatic for the team, which frees attention for the actual problems. The board is a visual retrieval cue. Recognition reinforces the behaviors we want repeated.
  FORMAT: 5–10 minutes, standing, at the Gemini (or at the cell board), at the start of each shift. Same order every day.
  SUPERVISOR TIP: Ask before you tell. Let operators report CILT findings and defects — it builds ownership.
  ASK: "Which of these nine topics is missing from today's shift start-up meetings?"`);

  }
  if (AM) {
  // ======================= 43. SAMPLE HUDDLE =======================
  s = content("PART 10 · DAILY MANAGEMENT", "Sample five-minute huddle — Gemini cell, start of shift");
  const seg = [["0:00", 0.75, "Safety & food quality", K.coral], ["0:45", 1.0, "Machine & CILT", K.teal], ["1:45", 0.75, "Centerline & defects", K.blue], ["2:30", 1.25, "Biggest loss", K.amber], ["3:45", 0.75, "Today’s action", K.navy], ["4:30", 0.5, "Recognition", K.green]];
  let sx = 0.6;
  seg.forEach(([t, d, lab, col]) => {
    const w = (12.1 * d) / 5;
    box(s, sx + 0.02, 1.5, w - 0.04, 0.6, { fill: col, r: 0.05 });
    txt(s, lab, { x: sx, y: 1.5, w, h: 0.6, fontSize: 11, bold: true, color: col === K.amber ? K.navy : K.white, align: "center", valign: "middle", margin: 0 });
    txt(s, t, { x: sx, y: 2.12, w: 0.6, h: 0.25, fontSize: 10, color: K.slate, margin: 0 });
    sx += w;
  });
  txt(s, "5:00", { x: 12.2, y: 2.12, w: 0.5, h: 0.25, fontSize: 10, color: K.slate, align: "right", margin: 0 });
  const script = [
    ["0:00–0:45", "Safety & food quality", "“Any injuries, near misses or hazards since last shift? Any holds or quality concerns?” Thank the person who raises one."],
    ["0:45–1:45", "Machine condition & CILT", "“CILT on Gemini [#] was completed at [__]%. Two abnormalities were found and tagged. [Name], what did you see?”"],
    ["1:45–2:30", "Centerline & open defects", "“One setting was out of range at startup — returned to standard and logged. Three defects open, one overdue. Maintenance, update?”"],
    ["2:30–3:45", "Yesterday’s biggest loss", "“Biggest loss yesterday: [minor stops at the transfer point], [__] minutes. What do you think is driving it?” Ask — don’t tell."],
    ["3:45–4:30", "Today’s improvement action", "“Today, [Name] checks guide condition with Maintenance at [time]. We’ll review results tomorrow.”"],
    ["4:30–5:00", "Recognition", "“[Name] spotted a loose bracket during CILT before it caused a jam. That’s ownership.”"],
  ];
  const scData = [[hdr("Time"), hdr("Topic"), hdr("What the supervisor says or does (example)")]].concat(script.map((r, i) => {
    const f = { fill: { color: i % 2 ? K.white : K.light } };
    return [{ text: r[0], options: Object.assign({ bold: true, color: K.navy }, f) }, { text: r[1], options: Object.assign({ bold: true, color: K.navy }, f) }, { text: r[2], options: Object.assign({ italic: true }, f) }];
  }));
  s.addTable(scData, { x: 0.6, y: 2.5, w: 12.1, colW: [1.3, 2.6, 8.2], rowH: [0.4, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6], fontSize: 12, color: K.ink, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: HEX.grid }, margin: 0.07 });
  badge(s, "ILLUSTRATIVE SCRIPT — [ ] = FILL IN", 0.6, 6.62, 3.3);
  notes(s, `PURPOSE: Give supervisors a concrete, five-minute huddle they can use tomorrow.
  TIMING: 4 minutes, then let one volunteer run it live (5 min).
  THE SCRIPT: Example language only. Bracketed items [ ] are placeholders to be filled with real information from that shift — the deck does not contain real Gemini data.
  TECHNIQUES TO POINT OUT:
  • Open with safety and food quality — always.
  • Let operators report their own CILT findings ("[Name], what did you see?").
  • Use the defect register — named owners give updates.
  • On the biggest loss, ASK what's driving it rather than announcing the answer.
  • End with one specific action, one owner, one time — and a commitment to review tomorrow.
  • Close with specific recognition ("spotted a loose bracket during CILT") rather than generic praise ("good job").
  LIVE PRACTICE: Ask a volunteer supervisor to run the huddle with 4–5 participants acting as the team. Time it. Debrief: what kept it to five minutes?`);

  }
  if (AM) {
  // ======================= 44. COACHING EXERCISE =======================
  s = dark("INTERACTIVE EXERCISE · 15 MINUTES", "Coach, don’t just tell: a role-play");
  box(s, 0.6, 1.5, 5.8, 2.15, { fill: K.white });
  txt(s, [{ text: "TELLING", options: { bold: true, color: K.coral, breakLine: true, fontSize: 12, charSpacing: 2 } }, { text: "“Run the speed back to standard and stop adjusting the sensor.”", options: { italic: true, color: K.navy, fontSize: 15 } }], { x: 0.85, y: 1.65, w: 5.4, h: 1.0, margin: 0 });
  txt(s, "Fast — but the operator learns only what not to do, and the reason behind the adjustment stays hidden.", { x: 0.85, y: 2.7, w: 5.4, h: 0.85, fontSize: 12, color: K.slate, margin: 0 });
  box(s, 0.6, 3.85, 5.8, 2.95, { fill: K.teal });
  txt(s, [{ text: "COACHING", options: { bold: true, color: K.white, breakLine: true, fontSize: 12, charSpacing: 2 } }, { text: "“What does the centerline say for this product? What did you see that made you adjust it? What could we check together before changing anything?”", options: { italic: true, color: K.white, fontSize: 15 } }], { x: 0.85, y: 4.0, w: 5.4, h: 1.8, margin: 0 });
  txt(s, "Slower at first — but builds understanding, surfaces the real cause, and lasts after you walk away.", { x: 0.85, y: 5.85, w: 5.4, h: 0.85, fontSize: 12, color: K.white, margin: 0 });
  txt(s, "HOW TO RUN IT", { x: 6.8, y: 1.5, w: 5.9, h: 0.3, fontSize: 11, bold: true, color: K.amber, charSpacing: 2, margin: 0 });
  const how = [["Pair up", "One supervisor, one ‘operator’"], ["Scenario", "Operator keeps moving a sensor bracket to stop jams"], ["Round A — 3 min", "Supervisor tells"], ["Round B — 3 min", "Supervisor coaches: Ask → Listen → Show the standard → Agree action → Follow up"], ["Debrief — 5 min", "Which round surfaced the real cause? Which will last?"]];
  how.forEach(([h, d], i) => {
    const y = 1.9 + i * 0.72;
    numCircle(s, i + 1, 6.8, y + 0.05, 0.45, K.amber, 13, K.navy);
    txt(s, [{ text: h + ":  ", options: { bold: true, color: K.white } }, { text: d, options: { color: K.light } }], { x: 7.4, y, w: 5.3, h: 0.66, fontSize: 13, valign: "middle", margin: 0 });
  });
  txt(s, "COACHING QUESTION BANK", { x: 6.8, y: 5.6, w: 5.9, h: 0.3, fontSize: 11, bold: true, color: K.amber, charSpacing: 2, margin: 0 });
  ["What’s normal here?", "What changed?", "What does the standard say?", "What would you try first?", "How will we know it worked?"].forEach((q, i) => {
    const x = 6.8 + (i % 3) * 1.98, y = 5.95 + Math.floor(i / 3) * 0.45;
    box(s, x, y, 1.9, 0.38, { fill: K.white, tr: 85, r: 0.15 });
    txt(s, q, { x, y, w: 1.9, h: 0.38, fontSize: 10, bold: true, color: K.white, align: "center", valign: "middle", margin: 0 });
  });
  notes(s, `PURPOSE: Let supervisors practice coaching rather than simply giving instructions.
  TIMING: 15 minutes (1 min setup, 3 + 3 min rounds, 3 min swap optional, 5 min debrief).
  SETUP: Pairs. One plays the supervisor, one plays an operator who has been repeatedly moving a sensor bracket to stop jams (connects to the Centerline Detective and jam case). Brief the 'operator' privately: "You moved it because the guide looks worn and nobody fixed it after you reported it."
  ROUND A — TELL: Supervisor gives instructions only.
  ROUND B — COACH: Supervisor uses the sequence Ask → Listen → Show the standard → Agree an action → Follow up. Use the question bank.
  DEBRIEF QUESTIONS: In which round did the supervisor learn about the worn guide? Which approach would change behavior on the next shift when the supervisor isn't there? How does coaching connect to the neuroscience (active processing and retrieval)?
  KEY POINT: Telling is sometimes necessary — for an immediate safety risk, tell and stop the job. Coaching is for building capability and finding causes.`);

  }
  if (true) {
  // ======================= 45. WOr STRUCTURE =======================
  section("Daily Management & WOr");
  s = content("PART 11 · WOr LEADERSHIP PRESENTATIONS", "WOr: every cell tells its story with the same structure");
  const wor = AM ? [["FaIndustry", "Machine ownership", 0], ["FaUsers", "Team members", 0], ["FaMapMarkedAlt", "Safety Map status", 1], ["FaClipboardCheck", "CILT status", 1], ["FaSlidersH", "Centerline status", 1], ["FaTag", "Defect Handling status", 1], ["FaTools", "Improvement actions", 0], ["FaCamera", "Before & after photos", 0], ["FaExclamationTriangle", "Open risks", 0], ["FaFlag", "Next actions", 0]]
    : [["FaIndustry", "Machine ownership", 0], ["FaUsers", "Team members", 0], ["FaChartBar", "Baseline performance", 1], ["FaSearch", "Top losses", 1], ["FaTools", "Improvement actions", 0], ["FaCamera", "Before & after photos", 0], ["FaChartLine", "OEE & quality trends", 1], ["FaDollarSign", "Financial benefits (Finance-validated)", 1], ["FaExclamationTriangle", "Open risks", 0], ["FaFlag", "Next actions", 0]];
  const wc = 5, ww = (12.1 - (wc - 1) * 0.15) / wc;
  wor.forEach(([ic, t, hl], i) => {
    const x = 0.6 + (i % wc) * (ww + 0.15), y = 1.5 + Math.floor(i / wc) * 2.05;
    box(s, x, y, ww, 1.85, { fill: hl ? K.teal : K.light, tr: hl ? 85 : 0 });
    txt(s, String(i + 1).padStart(2, "0"), { x: x + 0.12, y: y + 0.1, w: 0.6, h: 0.3, fontSize: 12, bold: true, color: K.slate, margin: 0 });
    iconOnly(s, ic, x + ww / 2 - 0.25, y + 0.45, 0.5, HEX.navy);
    txt(s, t, { x: x + 0.08, y: y + 1.05, w: ww - 0.16, h: 0.75, fontSize: 13, bold: true, color: K.navy, align: "center", valign: "middle", margin: 0 });
  });
  box(s, 0.6, 5.7, 12.1, 1.1, { fill: K.navy });
  txt(s, [{ text: "Proposed cadence: ", options: { bold: true, color: K.amber } }, { text: "Weeks 4, 8 and 13 · 10 minutes per cell + 5 minutes Q&A · led by the supervisor, with team members presenting their own sections. ", options: { color: K.white } }, { text: "Goal: build analytical, communication and leadership capability.", options: { color: K.white, bold: true } }], { x: 0.85, y: 5.7, w: 11.7, h: 1.1, fontSize: 14, valign: "middle" });
  notes(s, `PURPOSE: Introduce the standardized WOr leadership presentation for the ${AM ? "AM" : "VO"} pillar.
  TIMING: 4 minutes.
  NOTE ON THE TERM: "WOr" is the site's name for the periodic leadership review. Confirm the local definition and keep the term consistent.
  THE 10 ELEMENTS (always in this order): ${wor.map((w) => w[1]).join("; ")}. Highlighted tiles are the ${AM ? "four AM tools — every cell reports on all four" : "VO measures — baseline, losses, trends and Finance-validated benefits"}.
  CADENCE (proposed): end of each 30-day phase (Weeks 4, 8, 13); 10 minutes per cell plus 5 minutes Q&A.
  PURPOSE: Develop supervisors' analytical, communication and leadership capability — not create a reporting burden.
  ${VO ? "RULE: Only Finance-validated savings appear as results; estimates are labeled as estimates. Never present projected benefits as achieved." : "In a combined review with VO, append the VO section from the companion deck after element 6."}`);

  }
  if (true) {
  // ======================= 46. WOr TEMPLATE =======================
  s = content("TEMPLATE · DUPLICATE FOR EACH CELL", AM ? "WOr AM cell report: Gemini [#]" : "WOr VO cell report: Gemini [#]");
  box(s, 0.6, 1.4, 12.1, 0.45, { fill: K.light });
  txt(s, "Cell: [  ]   ·   Gemini ID: [  ]   ·   Supervisor: [  ]   ·   Team — Prod: [  ]  Maint: [  ]  Qual: [  ]   ·   Period: [  ]   ·   Review #: [  ]", { x: 0.75, y: 1.4, w: 11.8, h: 0.45, fontSize: 11, color: K.navy, valign: "middle", margin: 0 });
  const tl = AM ? [["Safety Map", "Draft / Approved"], ["CILT", "Compliance: [__]%"], ["Centerline", "Adherence: [__]%"], ["Defects", "Open [_] · Overdue [_] · Verified [_]"]] : [["Baseline", "Agreed with Finance: Y / N"], ["Top loss", "[  ] — [__] min/week"], ["Projects", "Active [_] · Closed [_]"], ["Validated savings", "$[  ] (Finance)"]];
  tl.forEach(([h, v], i) => {
    const x = 0.6 + i * 1.53;
    box(s, x, 2.0, 1.43, 1.1, { fill: AM ? K.teal : K.amber, tr: 85 });
    txt(s, [{ text: h, options: { bold: true, color: K.navy, breakLine: true, fontSize: 12 } }, { text: v, options: { color: K.ink, fontSize: 10 } }], { x: x + 0.08, y: 2.05, w: 1.3, h: 1.0, margin: 0 });
  });
  const km = AM ? ["CILT compliance", "Centerline adherence", "Open / overdue defects", "Repeat abnormalities"] : ["OEE", "Unplanned downtime", "Minor stops", "Scrap & rework"];
  const kd = [[hdr("Metric"), hdr("Baseline"), hdr("Current"), hdr("Trend")]].concat(km.map((m) => [{ text: m, options: { bold: true, color: K.navy } }, "TBD", "TBD", "TBD"]));
  s.addTable(kd, { x: 6.8, y: 2.0, w: 5.9, colW: [2.3, 1.2, 1.2, 1.2], rowH: 0.36, fontSize: 11, color: K.ink, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: HEX.grid }, margin: 0.06 });
  const sect = (x, y, w, h, title, body, fill = K.light) => { box(s, x, y, w, h, { fill }); txt(s, [{ text: title, options: { bold: true, color: K.navy, breakLine: true, fontSize: 12 } }, { text: body, options: { color: K.slate, fontSize: 11 } }], { x: x + 0.12, y: y + 0.08, w: w - 0.24, h: h - 0.16, margin: 0 }); };
  sect(0.6, 3.25, 2.95, 1.75, AM ? "Top 3 repeat abnormalities" : "Top 3 losses", "1. [  ]\n2. [  ]\n3. [  ]");
  sect(3.65, 3.25, 2.95, 1.75, "Improvement actions", "Action · owner · due · status");
  if (AM) sect(6.8, 4.0, 2.9, 1.4, "Training & competency", "Operators verified on CILT / centerline / defect handling: [__]%");
  else sect(6.8, 4.0, 2.9, 1.4, "Financial benefits", "Finance-validated only (Randy). Estimates labeled as estimates.");
  sect(9.8, 4.0, 2.9, 1.4, "Open risks", "Safety · quality · resource · schedule");
  sect(6.8, 5.55, 5.9, 1.25, "Next actions & help needed from leadership", "[  ]");
  [["BEFORE", 0.6], ["AFTER", 3.65]].forEach(([t, x]) => {
    box(s, x, 5.15, 2.95, 1.65, { fill: K.white, line: K.slate, dash: "dash", lw: 1 });
    iconOnly(s, "FaCamera", x + 1.25, 5.45, 0.45, HEX.slate);
    txt(s, t + " photo", { x, y: 6.0, w: 2.95, h: 0.4, fontSize: 11, bold: true, color: K.slate, align: "center", margin: 0 });
  });
  notes(s, `PURPOSE: A fillable one-page WOr template for each Gemini cell (${AM ? "AM" : "VO"} pillar).
  TIMING: 3 minutes to explain; cells complete it before each WOr review.
  HOW TO USE: Duplicate this slide for each cell. Replace every [ ] and TBD with real, validated information.
  RULES:
  ${AM ? "• Status tiles cover all four AM tools — never skip one.\n  • 'Repeat abnormalities' are defects that return after closure; each needs an RCA." : "• Financial benefits show only Finance-validated savings; any estimate is labeled 'estimate'.\n  • Trends use the agreed baseline and comparable periods."}
  • Photos must be taken safely and follow site photography rules in production areas.
  • Never present projected benefits as achieved results.`);

  }
  if (true) {
  // ======================= 47. EXEC DASHBOARD =======================
  section("Success Metrics & Close");
  s = content("PART 12 · SUCCESS METRICS", AM ? "AM executive dashboard — all three Geminis" : "VO executive dashboard — all three Geminis");
  const tiles = AM ? [["Safety maps approved", "AM"], ["CILT compliance", "AM"], ["Centerline adherence", "AM"], ["Open/overdue defects", "AM"], ["Verified closure", "AM"], ["Repeat abnormalities", "AM"]] : [["OEE", "VO"], ["Unplanned downtime", "VO"], ["Minor stops", "VO"], ["Changeover losses", "VO"], ["Scrap & rework", "VO"], ["Validated savings", "VO"]];
  for (let g = 0; g < 3; g++) {
    const x = 0.6 + g * 4.1;
    box(s, x, 1.45, 3.85, 0.5, { fill: K.navy, r: 0.08 });
    txt(s, "GEMINI " + (g + 1) + "  ·  CELL " + (g + 1), { x, y: 1.45, w: 3.85, h: 0.5, fontSize: 13, bold: true, color: K.white, align: "center", valign: "middle", charSpacing: 2 });
    tiles.forEach(([t, p], i) => {
      const tx2 = x + (i % 2) * 1.95, ty2 = 2.1 + Math.floor(i / 2) * 1.38;
      box(s, tx2, ty2, 1.9, 1.28, { fill: K.light });
      txt(s, t, { x: tx2 + 0.1, y: ty2 + 0.06, w: 1.3, h: 0.42, fontSize: 10, bold: true, color: K.slate, margin: 0 });
      box(s, tx2 + 1.45, ty2 + 0.08, 0.38, 0.22, { fill: p === "AM" ? K.teal : K.amber, r: 0.1 });
      txt(s, p, { x: tx2 + 1.45, y: ty2 + 0.08, w: 0.38, h: 0.22, fontSize: 8, bold: true, color: K.white, align: "center", valign: "middle", margin: 0 });
      txt(s, "TBD", { x: tx2 + 0.1, y: ty2 + 0.48, w: 1.7, h: 0.45, fontSize: 22, bold: true, color: K.navy, margin: 0 });
      circ(s, tx2 + 0.12, ty2 + 1.0, 0.14, K.slate);
      txt(s, "Target: TBD", { x: tx2 + 0.32, y: ty2 + 0.93, w: 1.5, h: 0.28, fontSize: 10, color: K.slate, margin: 0 });
    });
  }
  box(s, 0.6, 6.3, 12.1, 0.5, { fill: K.amber, tr: 80 });
  txt(s, AM ? "All values and targets are TBD until AM deployment data is collected and leadership approves targets." : "All values and targets are TBD until the baseline is locked and leadership approves targets. Savings appear only after Finance (Randy) validation.", { x: 0.8, y: 6.3, w: 11.7, h: 0.5, fontSize: 12, bold: true, color: K.navy, valign: "middle", margin: 0 });
  notes(s, `PURPOSE: The executive ${AM ? "AM" : "VO"} dashboard template for all three Geminis.
  TIMING: 3 minutes.
  HOW IT WORKS: One column per Gemini, six headline tiles per machine: ${tiles.map((t) => t[0]).join(", ")}.
  STATUS DOT: Gray = baseline pending. Once targets are approved, use green / amber / red — define the bands with leadership.
  EDITABLE: Every tile is a text box. Update weekly from the KPI register (next slide).
  TARGETS: TBD until ${AM ? "deployment data is collected" : "baseline data is locked"} and leadership approves. No values are shown because no actual results exist yet.`);

  }
  if (true) {
  // ======================= 48. KPI REGISTER =======================
  s = content("PART 12 · SUCCESS METRICS", AM ? "KPI register: AM deployment and leadership metrics" : "KPI register: VO performance and leadership metrics");
  const kpis = [
    ["AM", "Safety maps approved", "Approved maps ÷ 3 Geminis", "Cell records", "Weekly"],
    ["AM", "CILT compliance", "Tasks completed on time ÷ tasks scheduled", "CILT checklists", "Daily"],
    ["AM", "Centerline adherence", "Parameters in range ÷ parameters checked", "Centerline checks", "Per shift"],
    ["AM", "Open equipment defects", "Count of open defect tags", "Defect register", "Daily"],
    ["AM", "Overdue defects", "Open tags past due date", "Defect register", "Daily"],
    ["AM", "Verified defect closure", "Closed & verified ÷ total closed", "Defect register", "Weekly"],
    ["AM", "Repeat abnormalities", "Defects recurring within [X] days of closure", "Defect register", "Weekly"],
    ["VO", "OEE", "Availability × Performance × Quality", "Redzone", "Daily / weekly"],
    ["VO", "Unplanned downtime", "Minutes of unplanned stops", "Redzone", "Daily"],
    ["VO", "Minor stops", "Count and minutes of short stops (threshold TBD)", "Redzone", "Daily"],
    ["VO", "Changeover losses", "Changeover minutes vs standard", "Redzone / logs", "Per changeover"],
    ["VO", "Scrap & rework", "Scrap and rework as % of output", "Production / Quality records", "Daily"],
    ["VO", "Production productivity", "Output per labor hour", "Production & labor records", "Weekly"],
    ["VO", "Validated savings", "$ approved by Finance", "Finance (Randy)", "Monthly"],
    ["Lead", "Huddle completion", "Huddles held ÷ scheduled", "Huddle log", "Weekly"],
    ["Lead", "Employee engagement", "Participation: tags raised, ideas, pulse survey (TBD)", "Cell records", "Monthly"],
    ["Lead", "Training competency", "Verified competent ÷ required", "Training records", "Monthly"],
    ["Lead", "Supervisor-led projects", "Projects led and closed by supervisors", "Project tracker", "Monthly"],
    ["Lead", "Improvement sustainability", "Gains held for the agreed period", "WOr reviews", "Monthly"],
  ];
  const kSel = kpis.filter((r) => (AM ? r[0] !== "VO" : r[0] === "VO" || ["Supervisor-led projects", "Improvement sustainability"].includes(r[1])));
  const pCol = { AM: K.teal, VO: K.amber, Lead: K.blue };
  const kData = [[hdr("Pillar"), hdr("Metric"), hdr("Definition / calculation"), hdr("Source"), hdr("Frequency"), hdr("Target")]].concat(kSel.map((r) => [{ text: r[0], options: { bold: true, color: r[0] === "VO" ? K.navy : K.white, fill: { color: pCol[r[0]] }, align: "center" } }, { text: r[1], options: { bold: true, color: K.navy } }, r[2], r[3], r[4], { text: "TBD", options: { color: K.slate, align: "center" } }]));
  s.addTable(kData, { x: 0.6, y: 1.4, w: 12.1, colW: [0.75, 2.45, 4.3, 2.4, 1.4, 0.8], rowH: Math.min(0.4, 5.2 / (kSel.length + 1)), fontSize: kSel.length > 12 ? 11 : 12, color: K.ink, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: HEX.grid }, margin: [0.01, 0.06, 0.01, 0.06] });
  notes(s, `PURPOSE: Define every KPI so all three cells measure the same way.
  TIMING: 4 minutes (leadership); reference document afterwards.
  ${AM ? "AM DEPLOYMENT METRICS: safety maps approved; CILT compliance; centerline adherence; open equipment defects; overdue defects; verified defect closure; repeat abnormalities.\n  LEADERSHIP METRICS: huddle completion; employee engagement; training competency; supervisor-led projects; improvement sustainability.\n  DEFINITIONS: Proposed — confirm with CI, Maintenance and Quality.\n  TARGETS: TBD until deployment data and leadership approval are available.\n  NOTE: Performance outcomes (OEE, downtime, scrap, savings) are tracked in the VO companion deck." : "VO PERFORMANCE METRICS: OEE; unplanned downtime; minor stops; changeover losses; scrap and rework; production productivity; financial savings validated by Finance.\n  LEADERSHIP METRICS: supervisor-led projects; improvement sustainability.\n  DEFINITIONS: Proposed — confirm with CI, Finance and Quality, and match Redzone's calculation logic for OEE and stops (e.g., the threshold that defines a minor stop).\n  TARGETS: TBD until baseline data and leadership approval are available."}`);

  }
  if (true) {
  // ======================= 49. SITE DATA NEEDED =======================
  s = content("FINALIZING THE DEPLOYMENT", "Site-specific data still needed to finalize deployment");
  const need = AM ? [
    ["FaCogs", "Equipment & safety", K.navy, ["Confirmed Gemini machine IDs / asset numbers", "Real Gemini photos and layout drawings", "OEM manuals and specifications", "Existing PM tasks and CMMS records", "LOTO procedures and energy sources per machine", "Current guarding / interlock status"]],
    ["FaUsers", "People & organization", K.teal, ["Supervisor leader for each cell", "Production, Maintenance, Quality members", "Shift coverage per cell", "EHS contact for approvals", "Training schedule windows", "Local definition of “WOr” review"]],
    ["FaSlidersH", "Standards & settings", K.amber, ["Existing centerline / settings documentation", "Current SOPs and sanitation procedures", "Food-grade lubricant requirements", "CMMS work-order request process", "Existing tagging color standard", "Site severity & response-time standards"]],
    ["FaShieldAlt", "Quality & brand", K.blue, ["Quality specs for Gemini products", "Food-safety controls linked to the Gemini", "Product disposition procedure", "Approved C.H. Guenther / Future Foods brand assets", "Training & competency record format", "Photography rules in production areas"]],
  ] : [
    ["FaUsers", "People & organization", K.teal, ["Supervisor leader for each cell", "Production, Maintenance, Quality members", "Shift coverage per cell", "Finance (Randy) review cadence", "CI coaching capacity", "Local definition of “WOr” review"]],
    ["FaDatabase", "Data & performance", K.amber, ["Redzone line / asset mapping", "4–8 weeks of baseline data per Gemini", "Downtime reason codes & data-quality review", "Product mix and changeover matrix", "Labor hours by line and shift", "Agreed comparison period"]],
    ["FaDollarSign", "Finance assumptions", K.green, ["Cost basis for downtime hours", "Scrap & rework cost basis", "Labor rates and overtime rules", "Annualization rules for recurring benefits", "Cost-avoidance reporting rules", "Savings-reporting format"]],
    ["FaClipboardList", "Operations records", K.blue, ["Production plan vs actual records", "Maintenance loss records", "Quality hold, scrap & rework records", "List of existing improvement projects (avoid double-counting)", "Changeover standards & times", "Approved brand assets"]],
  ];
  need.forEach(([ic, h, col, items], i) => {
    const x = 0.6 + i * 3.07;
    box(s, x, 1.5, 2.88, 5.3, { fill: K.light });
    iconCircle(s, ic, x + 0.2, 1.68, 0.6, col);
    txt(s, h, { x: x + 0.9, y: 1.68, w: 1.9, h: 0.6, fontSize: 14, bold: true, color: K.navy, valign: "middle", margin: 0 });
    items.forEach((t, j) => {
      const y = 2.5 + j * 0.7;
      iconOnly(s, "FaRegSquare", x + 0.2, y + 0.08, 0.26, HEX.slate);
      txt(s, t, { x: x + 0.55, y, w: 2.25, h: 0.65, fontSize: 12, color: K.ink, margin: 0 });
    });
  });
  notes(s, `PURPOSE: List the site-specific data still needed to finalize the ${AM ? "AM" : "VO"} deployment. Nothing on this list has been assumed or invented in this deck.
  TIMING: 3 minutes; assign an owner and due date to each item at kickoff.
  ${need.map((n) => n[1].toUpperCase() + ": " + n[3].join("; ") + ".").join("\n  ")}
  ALSO NEEDED TO FINISH THIS DECK: the previous PowerPoint referenced as the foundation and the brand reference artwork were not available when this version was built.`);

  }
  if (true) {
  // ======================= 50. CLOSING =======================
  s = pres.addSlide({ masterName: "CLOSING", sectionTitle: curSection });
  s.addText("THREE MACHINES. THREE TEAMS. ONE CULTURE OF OWNERSHIP.", { placeholder: "title" });
  txt(s, AM ? "We are not asking supervisors to do more paperwork. We are giving them the tools, knowledge, and ownership to make their machines safer, more reliable, and more productive." : "We are not asking supervisors to do more paperwork. We are giving them the data, methods, and ownership to find the biggest losses, fix what matters most, and prove the results.", { x: 1.4, y: 2.45, w: 10.5, h: 1.1, fontSize: 18, color: K.light, align: "center", margin: 0 });
  (AM ? [["FaCogs", "Know your machine"], ["FaClipboardCheck", "Own your standards"], ["FaChartLine", "Improve every day"]] : [["FaSearch", "Know your losses"], ["FaBullseye", "Fix what matters most"], ["FaBalanceScale", "Prove it, then sustain it"]]).forEach(([ic, t], i) => {
    const x = 1.15 + i * 3.85;
    box(s, x, 3.95, 3.35, 2.55, { fill: K.white, tr: 90 });
    iconCircle(s, ic, x + 1.17, 4.15, 1.0, K.teal);
    txt(s, String(i + 1), { x: x + 0.2, y: 4.1, w: 0.6, h: 0.6, fontSize: 28, bold: true, color: K.amber, margin: 0 });
    txt(s, t, { x: x + 0.1, y: 5.3, w: 3.15, h: 0.8, fontSize: 22, bold: true, color: K.white, align: "center", valign: "middle", margin: 0 });
  });
  txt(s, "C.H. Guenther & Son — Duncanville  ·  Gemini Pilot: " + PILLAR, { x: 0.7, y: 6.85, w: 11.9, h: 0.35, fontSize: 11, color: K.light, align: "center", margin: 0 });
  notes(s, `PURPOSE: Close with energy and commitment.
  TIMING: 3 minutes.
  SAY: "Three machines. Three teams. One culture of ownership." Pause.
  READ THE MESSAGE on the slide aloud.
  THREE COMMITMENTS — ask the room to say them together:
  ${AM ? "1. Know your machine — its hazards, its normal condition, its centerline.\n  2. Own your standards — CILT, centerlines, defect handling, every shift.\n  3. Improve every day — one abnormality, one action, one owner at every huddle." : "1. Know your losses — a credible baseline and Pareto for your Gemini.\n  2. Fix what matters most — safety gate first, then impact versus ease.\n  3. Prove it, then sustain it — Finance-validated, across shifts and products."}
  CALL BACK: Refer to the flip chart from slide 2 ("What would be different if every operator owned their machine?") — "That's what we're building."
  NEXT STEPS: ${AM ? "Week 1 kickoff actions — confirm cells, members, machine IDs and shift coverage; schedule safety-map walks." : "Week 1 kickoff actions — confirm cells and members; map Redzone data; start the baseline data pull; book the baseline review with Finance."}`);

  }
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
