/**
 * Starting configuration. These areas and checklists are INITIAL ASSUMPTIONS and
 * must be verified against the actual Duncanville plant layout and procedures.
 * HR / administrators can change all of it in Administration → Areas & Skills.
 *
 * Format per competency: [name, description, flags] where flags may contain
 * "critical" (needs practical evaluation + evidence before Competent).
 */
type C = [string, string, string?];
export interface AreaDef {
  code: string;
  name: string;
  kind: "PRODUCTION" | "PLANT_WIDE";
  active: boolean;
  description: string;
  competencies: C[];
}

export const AREAS: AreaDef[] = [
  {
    code: "PW", name: "Pre-Weigh", kind: "PRODUCTION", active: true, description: "Ingredient weigh-up and staging.",
    competencies: [
      ["Ingredient identification", "Identifies every ingredient by label, code and appearance."],
      ["Recipe reading", "Reads the current recipe/batch sheet correctly, including units and version."],
      ["Scale operation", "Operates, zeroes and tares scales; performs the daily scale check."],
      ["Ingredient weighing accuracy", "Weighs within the tolerance stated on the batch sheet.", "critical"],
      ["Lot identification and traceability", "Records every ingredient lot used against the batch.", "critical"],
      ["Allergen control", "Segregates allergens, uses dedicated tools, prevents cross-contact.", "critical"],
      ["Material staging", "Stages weighed kits in the right place and order for mixing."],
      ["Barcode scanning", "Scans ingredients and kits; resolves scan errors correctly."],
      ["Ingredient reconciliation", "Reconciles usage against the work order and reports variances."],
      ["Documentation accuracy", "Completes records in real time with good documentation practices."],
      ["Food safety compliance", "Follows GMPs and food safety rules in the weigh area."],
      ["Handling incorrect ingredients", "Stops, isolates and escalates wrong or suspect ingredients.", "critical"],
      ["Container labeling", "Labels every container with contents, lot, date and allergen status."],
      ["Material movement", "Moves materials safely and keeps them identified."],
      ["Work order verification", "Confirms the work order matches the recipe before starting."],
    ],
  },
  {
    code: "MX", name: "Mixing", kind: "PRODUCTION", active: true,
    description: "Normal production method is AUTO operation unless an approved procedure explicitly authorizes another mode.",
    competencies: [
      ["Mixer startup", "Performs pre-start checks and starts the mixer per the approved procedure."],
      ["Automatic operation", "Runs the mixer in AUTO, the standard production mode."],
      ["Recipe selection", "Selects and confirms the correct recipe on the HMI."],
      ["Ingredient dosing", "Confirms automatic and hand-added ingredients are dosed correctly."],
      ["Mixing time control", "Verifies mix times match the recipe."],
      ["Mixing speed control", "Verifies mix speeds match the recipe."],
      ["Dough temperature monitoring", "Checks and records dough temperature; reacts to out-of-range results."],
      ["Batch quality verification", "Checks batch against the quality standard before release."],
      ["HMI operation", "Navigates HMI screens for recipes, status and alarms."],
      ["Alarm identification", "Identifies each alarm and what it means."],
      ["Safe alarm response", "Responds to alarms per procedure without bypassing safeguards or interlocks.", "critical"],
      ["Batch documentation", "Completes batch records accurately and on time."],
      ["Allergen changeover", "Performs and documents allergen changeover and verification.", "critical"],
      ["Cleaning and sanitation", "Cleans the mixer per the sanitation procedure, using lockout/tagout where required."],
      ["Equipment shutdown", "Shuts down the mixer safely per procedure."],
      ["Abnormal condition escalation", "Stops and escalates abnormal conditions to the lead or supervisor.", "critical"],
      ["Process troubleshooting", "Follows approved troubleshooting procedures. Manual mode only when authorized, trained and covered by an approved procedure; never bypasses interlocks or lockout/tagout."],
    ],
  },
  {
    code: "BL", name: "Blending", kind: "PRODUCTION", active: true, description: "Dry and wet blend preparation.",
    competencies: [
      ["Recipe verification", "Confirms recipe and version before blending."],
      ["Ingredient sequencing", "Adds ingredients in the specified order."],
      ["Blending equipment setup", "Sets up blending equipment per procedure."],
      ["Batch weight verification", "Verifies total batch weight within tolerance."],
      ["Blend time monitoring", "Monitors blend time against the specification."],
      ["Material transfer", "Transfers blended material without loss or contamination."],
      ["Product consistency checks", "Performs consistency checks and records results."],
      ["Lot traceability", "Records lots and links them to the blend batch.", "critical"],
      ["Allergen management", "Prevents allergen cross-contact during blending.", "critical"],
      ["Equipment cleaning", "Cleans blending equipment per the sanitation procedure."],
      ["Abnormal condition response", "Stops and escalates abnormal conditions."],
      ["Documentation", "Completes blending records accurately."],
    ],
  },
  ...(
    [
      ["DD", "Dough Dividing", true, ["Divider setup", "Piece weight checks", "Divider operation in AUTO", "Jam clearing per procedure (with lockout/tagout where required)", "Changeover and cleaning", "Documentation"]],
      ["DF", "Dough Forming", true, ["Former setup", "Shape and size checks", "Forming operation", "Reject handling", "Changeover and cleaning", "Documentation"]],
      ["SH", "Sheeting", true, ["Sheeter setup", "Sheet thickness checks", "Dusting control", "Sheeter operation", "Changeover and cleaning", "Documentation"]],
      ["LM", "Laminating", true, ["Laminator setup", "Fat/dough layer checks", "Laminator operation", "Changeover and cleaning", "Documentation"]],
      ["PF", "Proofing", false, ["Proofer setup", "Temperature and humidity checks", "Proof time control", "Cleaning", "Documentation"]],
      ["BK", "Baking", false, ["Oven start-up", "Bake profile verification", "Product temperature checks", "Oven shutdown", "Documentation"]],
      ["CL", "Cooling", true, ["Cooling conveyor operation", "Product temperature checks before freezing", "Belt monitoring", "Documentation"]],
      ["SF", "Spiral Freezing", true, ["Freezer start-up checklist", "Dwell time and belt speed", "Core temperature verification", "Frost and belt monitoring", "Freezer alarm response", "Cold environment safety"]],
      ["PK", "Packaging", true, ["Bagger / form-fill-seal operation", "Seal integrity checks", "Date and lot code verification", "Label and film verification", "Checkweigher verification", "Case packer operation", "Packaging changeover and line clearance"]],
      ["MD", "Metal Detection / Inspection", true, ["Metal detector test-piece checks", "Reject mechanism confirmation", "Failed-check response and product hold", "X-ray / vision inspection checks", "Inspection documentation"]],
      ["PL", "Palletizing", true, ["Palletizer operation", "Pallet pattern and height", "Stretch wrapper operation", "Pallet labeling", "Fault clearing per procedure"]],
      ["SN", "Sanitation", true, ["Chemical handling and SDS", "Lockout/tagout for cleaning", "Equipment teardown and reassembly", "Pre-op inspection", "Sanitation documentation"]],
      ["MH", "Material Handling", true, ["Powered industrial truck operation (licensed)", "Pallet jack operation", "FIFO / FEFO stock rotation", "Freezer put-away", "Inventory accuracy"]],
      ["LS", "Line Setup / Changeover", true, ["Line start-up sequence", "Product changeover", "Allergen changeover verification", "Line clearance sign-off", "Shutdown sequence"]],
    ] as [string, string, boolean, string[]][]
  ).map(([code, name, active, comps]) => ({
    code, name, kind: "PRODUCTION" as const, active,
    description: active ? "Starter checklist – verify against plant procedures." : "Configured but inactive (frozen operations scope). Enable in Administration if needed.",
    competencies: comps.map((c): C => {
      const critical = /test-piece|Reject mechanism|Failed-check|Allergen|Date and lot|Label and film|Core temperature|Lockout|licensed/i.test(c);
      return [c, "Starter competency – confirm the standard with the area owner.", critical ? "critical" : undefined];
    }),
  })),
  {
    code: "FS", name: "Food Safety & GMP", kind: "PLANT_WIDE", active: true, description: "Mandatory food safety training for all production roles.",
    competencies: [
      ["GMPs and personal hygiene", "Follows GMP, hygiene and PPE rules."],
      ["Allergen awareness", "Knows plant allergens and cross-contact controls.", "critical"],
      ["HACCP / CCP awareness", "Knows the CCPs in their area and the response to a deviation.", "critical"],
      ["Foreign material control", "Applies glass, brittle plastic and tool control rules."],
      ["Hold and non-conforming product", "Tags and isolates suspect product; never releases without QA."],
    ],
  },
  {
    code: "SA", name: "Safety", kind: "PLANT_WIDE", active: true, description: "Mandatory safety training for all production roles.",
    competencies: [
      ["Lockout/tagout awareness", "Knows authorized vs affected roles; never removes another person's lock.", "critical"],
      ["Machine guarding and E-stops", "Knows E-stop locations; never bypasses guards or interlocks."],
      ["Chemical safety (HazCom)", "Reads labels and SDS; uses the right PPE."],
      ["Cold environment safety", "Uses cold PPE, warm-up breaks and freezer door releases."],
      ["Emergency response and evacuation", "Knows alarms, evacuation routes and muster points."],
    ],
  },
];

/** Job titles and the areas whose competencies they must hold. */
export const JOB_ROLES: { title: string; areas: string[] }[] = [
  { title: "Pre-Weigh Operator", areas: ["PW", "FS", "SA"] },
  { title: "Mixer Operator", areas: ["MX", "FS", "SA"] },
  { title: "Blending Operator", areas: ["BL", "FS", "SA"] },
  { title: "Forming Operator", areas: ["DD", "DF", "FS", "SA"] },
  { title: "Freezer Operator", areas: ["SF", "FS", "SA"] },
  { title: "Packaging Operator", areas: ["PK", "MD", "FS", "SA"] },
  { title: "Palletizer Operator", areas: ["PL", "MH", "FS", "SA"] },
  { title: "Sanitation Technician", areas: ["SN", "FS", "SA"] },
  { title: "Line Lead", areas: ["LS", "FS", "SA"] },
  { title: "Production Associate", areas: ["FS", "SA"] },
];

export const PRIMARY_AREA_FOR_ROLE: Record<string, string> = {
  "Pre-Weigh Operator": "PW", "Mixer Operator": "MX", "Blending Operator": "BL", "Forming Operator": "DF",
  "Freezer Operator": "SF", "Packaging Operator": "PK", "Palletizer Operator": "PL", "Sanitation Technician": "SN",
  "Line Lead": "LS", "Production Associate": "PK",
};

export const CAREER_LEVELS = [
  { rank: 0, name: "Trainee", description: "Learning basic production operations." },
  { rank: 1, name: "Qualified Operator", description: "Independently qualified in the assigned production area." },
  { rank: 2, name: "Multi-Skilled Operator", description: "Independently qualified in at least two approved production areas." },
  { rank: 3, name: "Advanced Operator", description: "Advanced technical competencies and independent problem-solving in approved areas." },
  { rank: 4, name: "Certified Trainer", description: "Qualified to train and evaluate others under an approved training program." },
];
