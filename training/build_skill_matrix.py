"""Builds the Frozen Production Skill Matrix workbook (no baking operations).

Run:  python training/build_skill_matrix.py
Output: training/Frozen_Production_Skill_Matrix.xlsx
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).with_name("Frozen_Production_Skill_Matrix.xlsx")
FONT = "Arial"
N_EMP = 50                      # employee rows on the matrix
FIRST_ROW = 6                   # first employee / role data row
SKILL_COL = 10                  # column J = first skill column

# ---------------------------------------------------------------- roles
ROLES = [
    ("PA", "Production Associate"),
    ("BT", "Batching / Ingredient Room"),
    ("MX", "Mixer Operator"),
    ("FO", "Forming / Sheeting Operator"),
    ("FR", "Freezer Operator"),
    ("PK", "Packaging Operator"),
    ("WH", "Palletizer / Freezer Warehouse"),
    ("LL", "Line Lead"),
    ("SV", "Production Supervisor"),
]
ROLE_NAMES = [r[1] for r in ROLES]

# ---------------------------------------------------------------- skills
# (id, skill, what "Qualified (3)" looks like, training method, verification,
#  refresher, required level per role in ROLES order: PA BT MX FO FR PK WH LL SV)
CATEGORIES = [
    ("Food Safety & Quality", "1F4E79", [
        ("FS01", "GMPs & personal hygiene",
         "Follows hairnet/beard net, jewelry, handwashing, glove and smock rules without reminders; corrects others.",
         "Classroom + floor walk", "Written quiz + GMP observation", "Annual",
         [3, 3, 3, 3, 3, 3, 3, 4, 4]),
        ("FS02", "Allergen control & allergen changeover",
         "Knows which products/ingredients contain allergens on their line; performs and documents allergen clean & verification.",
         "Classroom + OJT", "Quiz + observed changeover", "Annual",
         [2, 3, 3, 3, 2, 3, 2, 4, 3]),
        ("FS03", "HACCP & CCP awareness",
         "Can name the CCPs on their line, the critical limit, and what to do when a limit is missed.",
         "Classroom", "Quiz + verbal check", "Annual",
         [2, 2, 2, 2, 3, 3, 2, 3, 3]),
        ("FS04", "Foreign material control",
         "Applies glass & brittle plastic, hard plastic, tool and blade accountability rules; reports damage immediately.",
         "Classroom + OJT", "Observation", "Annual",
         [3, 3, 3, 3, 3, 3, 3, 4, 3]),
        ("FS05", "Metal detector / X-ray CCP checks",
         "Runs test pieces (Fe / Non-Fe / SS) at required frequency, verifies reject works, and holds product back to last good check on failure.",
         "OJT", "Observed check x3 + record review", "Annual",
         [0, 0, 0, 0, 0, 3, 0, 4, 3]),
        ("FS06", "Food defense awareness",
         "Recognizes and reports unknown people, tampering signs, and unsecured ingredients.",
         "Classroom", "Quiz", "Annual",
         [2, 2, 2, 2, 2, 2, 2, 3, 3]),
        ("FS07", "Hold & non-conforming product",
         "Tags, segregates and logs suspect product; never releases held product without QA.",
         "Classroom + OJT", "Observation + record review", "Annual",
         [2, 3, 3, 3, 3, 3, 3, 4, 4]),
        ("FS08", "Good documentation practices",
         "Records in real time, in ink, single-line corrections with initials/date; no pre-filling or blanks.",
         "Classroom", "Record review", "Annual",
         [2, 3, 3, 3, 3, 3, 3, 4, 4]),
        ("FS09", "Environmental monitoring & zoning",
         "Follows traffic flow, color-coded tools and zone rules; understands why Listeria control matters in a frozen plant.",
         "Classroom", "Quiz + observation", "Annual",
         [2, 2, 2, 2, 3, 2, 2, 3, 3]),
        ("FS10", "Pre-op inspection & clean-as-you-go",
         "Completes pre-op checklist for own area; keeps area clean during the run without contaminating product.",
         "OJT", "Observation + pre-op record", "Annual",
         [3, 3, 3, 3, 3, 3, 3, 4, 3]),
    ]),
    ("Safety", "C00000", [
        ("SA01", "Lockout / Tagout (LOTO)",
         "Affected: knows never to remove locks. Authorized: applies and verifies zero energy per equipment-specific procedure.",
         "Classroom + practical", "Practical demo per machine", "Annual",
         [2, 3, 3, 3, 3, 3, 3, 3, 3]),
        ("SA02", "Machine guarding & E-stops",
         "Locates every E-stop on the line, never bypasses guards or interlocks, reports missing guards.",
         "OJT", "Walk-through test", "Annual",
         [3, 3, 3, 3, 3, 3, 3, 4, 3]),
        ("SA03", "Cold environment & freezer safety",
         "Wears required cold PPE, follows warm-up breaks, knows freezer door inside release and the buddy rule.",
         "Classroom + walk-through", "Verbal check + observation", "Annual",
         [3, 2, 2, 2, 3, 3, 3, 3, 3]),
        ("SA04", "Ammonia awareness & evacuation",
         "Recognizes ammonia odor/alarm, knows evacuation routes and muster point; does not respond to leaks.",
         "Classroom + drill", "Drill participation", "Annual",
         [3, 3, 3, 3, 3, 3, 3, 3, 4]),
        ("SA05", "Chemical safety / HazCom (SDS)",
         "Reads labels and SDS, uses correct PPE for sanitation chemicals and CO2/LN2 where used.",
         "Classroom", "Quiz", "Annual",
         [2, 3, 2, 2, 3, 2, 2, 3, 3]),
        ("SA06", "Ergonomics & manual handling",
         "Uses lift aids and team lifts, rotates tasks as scheduled, reports discomfort early.",
         "Classroom + OJT", "Observation", "Annual",
         [3, 3, 3, 3, 3, 3, 3, 3, 3]),
        ("SA07", "Powered industrial truck (forklift / pallet jack)",
         "Holds current PIT license; completes pre-use inspection; drives safely in freezer conditions.",
         "Certified course + practical", "Practical evaluation (licensed)", "Every 3 years",
         [0, 2, 0, 0, 0, 0, 3, 2, 0]),
        ("SA08", "PPE & hearing conservation",
         "Wears correct PPE for the task (cut gloves, ear plugs, safety glasses, freezer gear) at all times.",
         "Classroom", "Observation", "Annual",
         [3, 3, 3, 3, 3, 3, 3, 3, 3]),
    ]),
    ("Ingredient Room / Batching", "7030A0", [
        ("BA01", "Reading batch sheets & formulas",
         "Picks the correct formula version, follows ingredient order and units, flags any mismatch.",
         "OJT", "Observed batch x3", "On change",
         [0, 3, 2, 0, 0, 0, 0, 3, 2]),
        ("BA02", "Scale operation & daily verification",
         "Performs daily scale check with test weights, records it, and removes an out-of-tolerance scale from use.",
         "OJT", "Observed check + record", "Annual",
         [0, 3, 2, 0, 0, 1, 0, 3, 2]),
        ("BA03", "Lot traceability & FIFO/FEFO",
         "Records every ingredient lot used; rotates stock oldest/first-expiring first; labels opened materials.",
         "Classroom + OJT", "Record review", "Annual",
         [1, 3, 3, 1, 0, 2, 3, 3, 3]),
        ("BA04", "Allergen staging & segregation",
         "Stores allergens in designated areas, uses dedicated scoops/containers, prevents cross-contact.",
         "OJT", "Observation", "Annual",
         [0, 3, 2, 0, 0, 0, 2, 3, 2]),
        ("BA05", "Minor ingredient weigh-up",
         "Weighs minors (salt, leavening, spices, enzymes) within tolerance, labels and verifies each kit.",
         "OJT", "Observed weigh-up x3", "Annual",
         [0, 3, 2, 0, 0, 0, 0, 2, 1]),
    ]),
    ("Mixing", "C55A11", [
        ("MX01", "Mixer setup & operation",
         "Sets mixer per recipe (speed/time), adds ingredients in order, runs safely and to target.",
         "OJT", "Observed batches x5", "Annual",
         [0, 1, 3, 1, 0, 0, 0, 3, 2]),
        ("MX02", "Dough / batter temperature control",
         "Calculates water temperature, uses ice/chilled water/CO2 as specified, hits target dough temp.",
         "OJT", "Temperature records review", "Annual",
         [0, 0, 3, 1, 0, 0, 0, 3, 2]),
        ("MX03", "Mix time & consistency checks",
         "Checks consistency/development against standard; adjusts within allowed limits; escalates outside limits.",
         "OJT", "Observation + records", "Annual",
         [0, 0, 3, 1, 0, 0, 0, 3, 2]),
        ("MX04", "Rework handling",
         "Uses rework only per procedure (same allergen profile, within time/percentage limits) and records it.",
         "Classroom + OJT", "Record review", "Annual",
         [0, 1, 3, 2, 0, 0, 0, 3, 3]),
    ]),
    ("Forming / Sheeting / Depositing", "548235", [
        ("FO01", "Sheeter / laminator setup",
         "Sets roll gaps, flour dusting and belt speeds to spec; produces in-spec sheet within startup window.",
         "OJT", "Observed startup x3", "Annual",
         [1, 0, 1, 3, 0, 0, 0, 3, 2]),
        ("FO02", "Cutter / die changeover",
         "Changes cutters/dies safely (LOTO), checks for damage and accounts for all parts.",
         "OJT", "Observed changeover x2", "Annual",
         [0, 0, 0, 3, 0, 0, 0, 3, 2]),
        ("FO03", "Depositor / filler operation",
         "Runs depositor/filler to target weight and placement; clears jams safely.",
         "OJT", "Observed run", "Annual",
         [1, 0, 0, 3, 0, 0, 0, 3, 2]),
        ("FO04", "Piece weight & dimension checks",
         "Takes samples at frequency, records weights/dimensions, reacts to trends before going out of spec.",
         "OJT", "Record review", "Annual",
         [2, 0, 1, 3, 0, 2, 0, 3, 3]),
        ("FO05", "Conveyor & transfer troubleshooting",
         "Spots and clears product build-up, mistracking and transfer issues safely (with LOTO when required).",
         "OJT", "Observation", "Annual",
         [1, 0, 0, 3, 2, 2, 0, 3, 2]),
    ]),
    ("Freezing", "2E75B6", [
        ("FR01", "Spiral / tunnel freezer start-up & shutdown",
         "Follows start-up/shutdown checklist; confirms air temperature at setpoint before product enters.",
         "OJT", "Observed start-up x3", "Annual",
         [0, 0, 0, 1, 3, 0, 0, 3, 2]),
        ("FR02", "Dwell time & belt speed control",
         "Sets belt speed for the product's required dwell time; adjusts when line rate changes.",
         "OJT", "Observation + records", "Annual",
         [0, 0, 0, 1, 3, 0, 0, 3, 2]),
        ("FR03", "Product core temperature checks",
         "Uses calibrated probe correctly, checks exit core temp at frequency, holds product that is not at spec.",
         "OJT", "Observed check x3 + records", "Annual",
         [1, 0, 0, 1, 3, 3, 2, 3, 3]),
        ("FR04", "Frost, defrost & belt monitoring",
         "Monitors coil frost and belt condition; follows defrost schedule; reports damaged belt links.",
         "OJT", "Observation", "Annual",
         [0, 0, 0, 0, 3, 0, 0, 3, 2]),
        ("FR05", "Freezer alarm response",
         "Knows each alarm, the first response, and who to call (maintenance / refrigeration).",
         "OJT", "Verbal scenario check", "Annual",
         [0, 0, 0, 0, 3, 0, 1, 3, 3]),
    ]),
    ("Packaging", "BF8F00", [
        ("PK01", "Bagger / form-fill-seal operation",
         "Threads film, sets jaws and temperatures, runs at target speed with minimal waste.",
         "OJT", "Observed run x3", "Annual",
         [1, 0, 0, 0, 0, 3, 0, 3, 2]),
        ("PK02", "Seal integrity checks",
         "Performs seal checks at frequency; recognizes and reacts to leaking/weak seals.",
         "OJT", "Observed check + records", "Annual",
         [2, 0, 0, 0, 0, 3, 0, 3, 2]),
        ("PK03", "Date / lot code setup & verification",
         "Sets and verifies code at start-up, changeover and frequency; holds product with wrong/missing codes.",
         "OJT", "Observed check + records", "Annual",
         [2, 0, 0, 0, 0, 3, 2, 4, 3]),
        ("PK04", "Label & packaging material verification",
         "Confirms film, labels and cases match the product and allergen statement before running.",
         "OJT", "Observed check + records", "Annual",
         [2, 0, 0, 0, 0, 3, 2, 4, 3]),
        ("PK05", "Checkweigher operation & verification",
         "Verifies checkweigher with test weights, confirms rejects, responds to underweight trends.",
         "OJT", "Observed check + records", "Annual",
         [1, 0, 0, 0, 0, 3, 0, 3, 2]),
        ("PK06", "Case packer / cartoner / case sealer",
         "Runs, adjusts and clears jams on secondary packaging equipment safely.",
         "OJT", "Observed run", "Annual",
         [2, 0, 0, 0, 0, 3, 1, 3, 2]),
        ("PK07", "Packaging changeover & line clearance",
         "Removes all previous film/labels/cases and product; documents line clearance before restart.",
         "OJT", "Observed changeover + record", "Annual",
         [1, 0, 0, 0, 0, 3, 0, 4, 3]),
    ]),
    ("Palletizing & Cold Storage", "3B3838", [
        ("WH01", "Palletizer & stretch wrapper",
         "Builds pallets to pattern and height, operates wrapper, clears faults safely.",
         "OJT", "Observed run", "Annual",
         [1, 0, 0, 0, 0, 1, 3, 3, 2]),
        ("WH02", "Pallet labeling & production reporting",
         "Applies correct pallet tag, scans/records production counts accurately.",
         "OJT", "Record review", "Annual",
         [0, 0, 0, 0, 0, 1, 3, 3, 3]),
        ("WH03", "Freezer put-away & temperature monitoring",
         "Moves product to freezer within time limit, checks and records storage temperature, keeps doors closed.",
         "OJT", "Observation + records", "Annual",
         [0, 0, 0, 0, 1, 0, 3, 3, 2]),
        ("WH04", "Inventory counts & stock rotation",
         "Performs accurate cycle counts; rotates finished goods FIFO; reports damaged pallets.",
         "OJT", "Count accuracy check", "Annual",
         [0, 1, 0, 0, 0, 0, 3, 2, 2]),
    ]),
    ("Line Leadership & CI", "404040", [
        ("LL01", "Line start-up & shutdown sequence",
         "Coordinates the full line start-up/shutdown in the correct order with all stations.",
         "OJT", "Observed start-up x3", "Annual",
         [0, 0, 1, 1, 1, 1, 0, 4, 3]),
        ("LL02", "Changeover planning (SMED)",
         "Plans product/allergen changeovers, stages parts and people, hits target changeover time.",
         "Classroom + OJT", "Observed changeover", "Annual",
         [0, 0, 1, 1, 0, 1, 0, 3, 3]),
        ("LL03", "Troubleshooting & root cause (5 Why)",
         "Leads a 5-Why for line problems and records the cause and fix.",
         "Classroom + OJT", "Completed 5-Why review", "Annual",
         [0, 0, 1, 1, 1, 1, 0, 3, 4]),
        ("LL04", "Downtime & OEE tracking",
         "Records downtime with correct reason codes; understands availability, performance and quality losses.",
         "Classroom", "Record review", "Annual",
         [0, 0, 1, 1, 1, 1, 1, 3, 4]),
        ("LL05", "On-the-job trainer",
         "Trains new people with the standard work and signs off levels 1-3 honestly.",
         "Train-the-trainer", "Supervisor observation", "Every 2 years",
         [0, 0, 0, 0, 0, 0, 0, 3, 4]),
        ("LL06", "Shift handover communication",
         "Gives/receives a complete handover: safety, quality holds, downtime, schedule status.",
         "OJT", "Observation", "Annual",
         [1, 1, 2, 2, 2, 2, 2, 4, 4]),
        ("LL07", "Autonomous maintenance (CIL)",
         "Completes clean-inspect-lube checklists for own equipment and writes clear work requests.",
         "OJT", "Checklist review", "Annual",
         [1, 1, 2, 2, 2, 2, 1, 3, 2]),
    ]),
]

SKILLS = []  # flat list of dicts
for cat, color, items in CATEGORIES:
    for sid, name, qual, method, verify, refresh, req in items:
        assert len(req) == len(ROLES), sid
        SKILLS.append(dict(id=sid, cat=cat, color=color, name=name, qual=qual,
                           method=method, verify=verify, refresh=refresh, req=req))
N_SKILLS = len(SKILLS)
LAST_SKILL_COL = SKILL_COL + N_SKILLS - 1
LSC = get_column_letter(LAST_SKILL_COL)
FSC = get_column_letter(SKILL_COL)
LAST_EMP_ROW = FIRST_ROW + N_EMP - 1

LEVELS = [
    (0, "Not trained", "Has not started training on this skill.", "F2F2F2", "7F7F7F"),
    (1, "In training", "Learning. Works only under direct supervision of a trainer.", "FFF2CC", "7F6000"),
    (2, "Can do with help", "Does the task but still needs support or checks from a qualified person.", "FCE4D6", "843C0C"),
    (3, "Qualified", "Works independently to standard. Signed off by trainer AND supervisor.", "C6EFCE", "006100"),
    (4, "Expert / Trainer", "Can troubleshoot, train others and sign off levels 1-3.", "548235", "FFFFFF"),
]
GAP_FILL, GAP_FONT = "FF7C80", "9C0006"

# ---------------------------------------------------------------- styles
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
INPUT_FILL = PatternFill("solid", fgColor="FFFFE0")
HEAD_FILL = PatternFill("solid", fgColor="1F3864")


def f(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


def title(ws, text, sub, width_to):
    ws["A1"] = text
    ws["A1"].font = f(16, True, "1F3864")
    ws["A2"] = sub
    ws["A2"].font = f(10, italic=True, color="595959")
    ws.row_dimensions[1].height = 26


def head(cell, text, color="1F3864", rotate=False):
    cell.value = text
    cell.font = f(10, True, "FFFFFF")
    cell.fill = fill(color)
    cell.border = BORDER
    cell.alignment = Alignment(horizontal="center", vertical="bottom" if rotate else "center",
                               wrap_text=True, text_rotation=90 if rotate else 0)


wb = Workbook()

# ================================================================ How To Use
ws = wb.active
ws.title = "How To Use"
ws.sheet_view.showGridLines = False
title(ws, "C.H. Guenther – Frozen Production Skill Matrix",
      "Frozen operations only (batching, mixing, forming, freezing, packaging, palletizing). Baking / oven skills are intentionally excluded.", 6)
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 24
ws.column_dimensions["C"].width = 22
ws.column_dimensions["D"].width = 80

r = 4
ws.cell(r, 2, "1. Rating scale – use the same scale for every skill").font = f(12, True, "1F3864")
r += 1
for c, h in zip((2, 3, 4), ("Level", "Name", "What it means on the floor")):
    head(ws.cell(r, c), h)
for lvl, name, desc, bg, fg in LEVELS:
    r += 1
    a = ws.cell(r, 2, lvl); a.fill = fill(bg); a.font = f(11, True, fg); a.alignment = CENTER; a.border = BORDER
    b = ws.cell(r, 3, name); b.font = f(10, True); b.alignment = LEFT; b.border = BORDER
    d = ws.cell(r, 4, desc); d.font = f(10); d.alignment = LEFT; d.border = BORDER
r += 1
a = ws.cell(r, 2, "RED"); a.fill = fill(GAP_FILL); a.font = f(11, True, GAP_FONT); a.alignment = CENTER; a.border = BORDER
b = ws.cell(r, 3, "Gap"); b.font = f(10, True); b.border = BORDER
d = ws.cell(r, 4, "On the Skill Matrix, any cell turns RED when the person is BELOW the level their role requires "
                 "(see the Role Requirements tab). Red = needs a training plan.")
d.font = f(10); d.alignment = LEFT; d.border = BORDER
ws.row_dimensions[r].height = 30

r += 2
ws.cell(r, 2, "2. How to fill it in (5 steps)").font = f(12, True, "1F3864")
steps = [
    ("Step 1", "Skill Matrix tab", "Delete the 3 EXAMPLE rows. Enter each person's ID, name, role (drop-down), shift (drop-down) and line/area in the light-yellow cells."),
    ("Step 2", "Skill Matrix tab", "For each skill, pick the person's current level 0–4 from the drop-down. Leave blank or 0 if not trained. Only enter 3 or 4 when a sign-off exists."),
    ("Step 3", "Skill Matrix tab", "Read the RED cells and the 'Gaps' / '% Ready' columns. These are calculated automatically – do not type in them."),
    ("Step 4", "Training Plan tab", "For every red cell, add a line: person, skill ID, target level, trainer and due date. Update status as training progresses."),
    ("Step 5", "Shift Coverage tab", "Check every skill has at least the minimum number of qualified people (level 3+) on each shift. 'HIGH RISK' = only one or zero qualified people in the whole team."),
]
r += 1
for c, h in zip((2, 3, 4), ("Step", "Where", "What to do")):
    head(ws.cell(r, c), h)
for s in steps:
    r += 1
    for c, v in zip((2, 3, 4), s):
        x = ws.cell(r, c, v); x.font = f(10, c == 2); x.alignment = LEFT; x.border = BORDER
    ws.row_dimensions[r].height = 32

r += 2
ws.cell(r, 2, "3. Tabs in this workbook").font = f(12, True, "1F3864")
tabs = [
    ("Skill Matrix", "Main grid. One row per person, one column per skill. Live gap highlighting and readiness %."),
    ("Role Requirements", "Minimum level each role must reach for each skill. Edit here to match your line – the matrix updates automatically."),
    ("Skills Library", "Definition of every skill: what 'Qualified' looks like, how it is trained, how it is verified, refresher frequency, SOP reference."),
    ("Shift Coverage", "Number of qualified people and trainers per skill, per shift, with OK / GAP / HIGH RISK flags."),
    ("Training Plan", "Action list to close gaps, with trainer and supervisor sign-off dates."),
    ("Lists", "Drop-down values (roles, shifts, status). Edit shift names here if needed."),
]
r += 1
for name, desc in tabs:
    b = ws.cell(r, 2, name); b.font = f(10, True); b.border = BORDER; b.alignment = LEFT
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    d = ws.cell(r, 3, desc); d.font = f(10); d.alignment = LEFT; d.border = BORDER
    ws.row_dimensions[r].height = 28
    r += 1

r += 1
ws.cell(r, 2, "4. Rules for sign-off").font = f(12, True, "1F3864")
rules = [
    "Level 3 requires: completed training + observed performance (see 'Verification' in Skills Library) + trainer AND supervisor signatures on the Training Plan.",
    "Level 4 requires: Level 3 for 6+ months, completed On-the-Job Trainer skill (LL05), and supervisor approval.",
    "Food safety and CCP skills (FS02, FS03, FS05, FR03, PK03, PK04) must be verified by QA in addition to the supervisor.",
    "Review the matrix monthly at the production meeting and after every new hire, transfer, new product or new equipment.",
    "Keep signed training records per the site document control and record retention procedure – this workbook is the summary, not the record.",
    "Yellow cells = you type here. White / grey cells = formulas or reference text.",
]
for t in rules:
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    c = ws.cell(r, 2, "•  " + t); c.font = f(10); c.alignment = LEFT
    ws.row_dimensions[r].height = 28

# ================================================================ Lists
wl = wb.create_sheet("Lists")
wl["A1"] = "Roles"; wl["B1"] = "Code"; wl["D1"] = "Shifts"; wl["F1"] = "Status"; wl["H1"] = "Levels"
for c in ("A1", "B1", "D1", "F1", "H1"):
    head(wl[c], wl[c].value)
for i, (code, name) in enumerate(ROLES, start=2):
    wl.cell(i, 1, name).font = f(); wl.cell(i, 2, code).font = f()
for i, s in enumerate(["1st Shift", "2nd Shift", "3rd Shift"], start=2):
    c = wl.cell(i, 4, s); c.font = f(); c.fill = INPUT_FILL
for i, s in enumerate(["Not started", "In progress", "Ready for sign-off", "Complete", "On hold"], start=2):
    wl.cell(i, 6, s).font = f()
for i, (lvl, *_r) in enumerate(LEVELS, start=2):
    wl.cell(i, 8, lvl).font = f()
wl["D6"] = "Shift names are editable (yellow)."; wl["D6"].font = f(9, italic=True, color="595959")
for col, w in zip("ABCDEFGH", (32, 8, 3, 14, 3, 20, 3, 8)):
    wl.column_dimensions[col].width = w
ROLE_LIST = f"Lists!$A$2:$A${1 + len(ROLES)}"
SHIFT_LIST = "Lists!$D$2:$D$4"

# ================================================================ Role Requirements
wr = wb.create_sheet("Role Requirements", 1)
title(wr, "Role Requirements – minimum level required per role",
      "0 or blank = not required for this role. Edit the yellow cells to match your line; the Skill Matrix gap colors update automatically.", 0)


def skill_headers(ws):
    # row 3 = category bands, row 4 = skill id, row 5 = skill name (rotated)
    col = SKILL_COL
    for cat, color, items in CATEGORIES:
        start = col
        for sk in items:
            head(ws.cell(4, col), sk[0], color)
            head(ws.cell(5, col), sk[1], color, rotate=True)
            ws.column_dimensions[get_column_letter(col)].width = 6
            col += 1
        ws.merge_cells(start_row=3, start_column=start, end_row=3, end_column=col - 1)
        head(ws.cell(3, start), cat, color)
        for cc in range(start, col):
            ws.cell(3, cc).border = BORDER
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[5].height = 190


skill_headers(wr)
for c, h in ((2, "Code"), (3, "Role")):
    wr.merge_cells(start_row=3, start_column=c, end_row=5, end_column=c)
    head(wr.cell(3, c), h)
wr.merge_cells(start_row=3, start_column=4, end_row=5, end_column=SKILL_COL - 1)
head(wr.cell(3, 4), "# skills required")
wr.column_dimensions["A"].width = 2
wr.column_dimensions["B"].width = 7
wr.column_dimensions["C"].width = 30
for cc in range(4, SKILL_COL):
    wr.column_dimensions[get_column_letter(cc)].width = 3
for i, (code, name) in enumerate(ROLES):
    rr = FIRST_ROW + i
    wr.cell(rr, 2, code).font = f(10, True)
    wr.cell(rr, 3, name).font = f(10, True)
    wr.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=SKILL_COL - 1)
    cnt = wr.cell(rr, 4, f'=COUNTIF({FSC}{rr}:{LSC}{rr},">0")')
    cnt.font = f(10, True); cnt.alignment = CENTER
    for c in (2, 3, 4):
        wr.cell(rr, c).border = BORDER
        wr.cell(rr, c).alignment = CENTER if c != 3 else LEFT
    for j, sk in enumerate(SKILLS):
        v = sk["req"][i]
        c = wr.cell(rr, SKILL_COL + j, v if v else None)
        c.font = f(10, True); c.alignment = CENTER; c.border = BORDER; c.fill = INPUT_FILL
    wr.row_dimensions[rr].height = 22
REQ_LAST = FIRST_ROW + len(ROLES) - 1
REQ_GRID = f"'Role Requirements'!${FSC}${FIRST_ROW}:${LSC}${REQ_LAST}"
REQ_ROLES = f"'Role Requirements'!$C${FIRST_ROW}:$C${REQ_LAST}"
req_rng = f"{FSC}{FIRST_ROW}:{LSC}{REQ_LAST}"
for lvl, _n, _d, bg, fg in LEVELS[1:]:
    wr.conditional_formatting.add(req_rng, CellIsRule(operator="equal", formula=[str(lvl)],
                                                      fill=fill(bg), font=Font(name=FONT, bold=True, color=fg)))
dv = DataValidation(type="whole", operator="between", formula1="0", formula2="4", allow_blank=True,
                    error="Enter a level from 0 to 4", errorTitle="Invalid level")
wr.add_data_validation(dv); dv.add(req_rng)
note = REQ_LAST + 2
wr.cell(note, 3, "Starting values are a recommended baseline for a frozen dough/food line. Adjust to your actual equipment and job descriptions.").font = f(9, italic=True, color="595959")
wr.freeze_panes = wr.cell(FIRST_ROW, SKILL_COL)

# ================================================================ Skill Matrix
wm = wb.create_sheet("Skill Matrix", 1)
title(wm, "Frozen Production Skill Matrix",
      "Pick a level 0–4 in each cell.  RED = below what the role requires.  0 Not trained · 1 In training · 2 Can do with help · 3 Qualified · 4 Expert/Trainer", 0)
skill_headers(wm)
fixed = [("A", "Emp ID", 9), ("B", "Name", 24), ("C", "Role", 26), ("D", "Shift", 10),
         ("E", "Line / Area", 12), ("F", "Skills required", 6), ("G", "Skills met", 6),
         ("H", "Gaps", 6), ("I", "% Ready", 8)]
for col, h, w in fixed:
    ci = wm[col + "3"].column
    wm.merge_cells(start_row=3, start_column=ci, end_row=5, end_column=ci)
    head(wm.cell(3, ci), h, "1F3864" if ci < 6 else "404040", rotate=ci >= 6)
    wm.column_dimensions[col].width = w

examples = {
    0: ("E-0001", "EXAMPLE – delete me", "Packaging Operator", "1st Shift", "Line 1"),
    1: ("E-0002", "EXAMPLE – delete me", "Freezer Operator", "2nd Shift", "Line 1"),
    2: ("E-0003", "EXAMPLE – delete me", "Line Lead", "1st Shift", "Line 1"),
}


def example_levels(role_idx, seed):
    out = []
    for j, sk in enumerate(SKILLS):
        req = sk["req"][role_idx]
        v = req
        if (j + seed) % 7 == 0 and req > 0:
            v = req - 1          # deliberate gap so the red highlight is visible
        elif (j + seed) % 5 == 0 and req < 4:
            v = req + 1          # some cross-training
        out.append(max(0, min(4, v)))
    return out


role_index = {n: i for i, n in enumerate(ROLE_NAMES)}
for k in range(N_EMP):
    rr = FIRST_ROW + k
    ex = examples.get(k)
    for ci in range(1, 6):
        c = wm.cell(rr, ci, ex[ci - 1] if ex else None)
        c.fill = INPUT_FILL; c.border = BORDER; c.font = f(10, italic=bool(ex), color="595959" if ex else "000000")
        c.alignment = LEFT if ci in (2, 3) else CENTER
    req_row = f"INDEX({REQ_GRID},MATCH($C{rr},{REQ_ROLES},0),0)"
    lv = f"${FSC}{rr}:${LSC}{rr}"
    wm.cell(rr, 6, f'=IF($C{rr}="","",SUMPRODUCT(--({req_row}>0)))')
    wm.cell(rr, 7, f'=IF($C{rr}="","",SUMPRODUCT(({req_row}>0)*({lv}>={req_row})))')
    wm.cell(rr, 8, f'=IF($C{rr}="","",F{rr}-G{rr})')
    wm.cell(rr, 9, f'=IF(OR($C{rr}="",F{rr}=0),"",G{rr}/F{rr})')
    for ci in range(6, 10):
        c = wm.cell(rr, ci); c.font = f(10, True); c.alignment = CENTER; c.border = BORDER
        c.fill = fill("F2F2F2")
    wm.cell(rr, 9).number_format = "0%"
    lv_vals = example_levels(role_index[ex[2]], k) if ex else None
    for j in range(N_SKILLS):
        c = wm.cell(rr, SKILL_COL + j, lv_vals[j] if ex else None)
        c.alignment = CENTER; c.border = BORDER; c.font = f(10, True)
    wm.row_dimensions[rr].height = 18

grid = f"{FSC}{FIRST_ROW}:{LSC}{LAST_EMP_ROW}"
tl = f"{FSC}{FIRST_ROW}"
req_cell = f"INDEX('Role Requirements'!{FSC}${FIRST_ROW}:{FSC}${REQ_LAST},MATCH($C{FIRST_ROW},{REQ_ROLES},0))"
gap_formula = f'AND($C{FIRST_ROW}<>"",IFERROR({req_cell},0)>0,IF(ISNUMBER({tl}),{tl},0)<IFERROR({req_cell},0))'
wm.conditional_formatting.add(grid, FormulaRule(formula=[gap_formula], stopIfTrue=True,
                                                fill=fill(GAP_FILL), font=Font(name=FONT, bold=True, color=GAP_FONT)))
for lvl, _n, _d, bg, fg in LEVELS:
    wm.conditional_formatting.add(grid, FormulaRule(formula=[f'AND(ISNUMBER({tl}),{tl}={lvl})'],
                                                    fill=fill(bg), font=Font(name=FONT, bold=True, color=fg)))
# % ready traffic light
pr = f"I{FIRST_ROW}:I{LAST_EMP_ROW}"
wm.conditional_formatting.add(pr, FormulaRule(formula=[f'AND(ISNUMBER(I{FIRST_ROW}),I{FIRST_ROW}>=1)'], fill=fill("C6EFCE"), font=Font(name=FONT, bold=True, color="006100")))
wm.conditional_formatting.add(pr, FormulaRule(formula=[f'AND(ISNUMBER(I{FIRST_ROW}),I{FIRST_ROW}>=0.8)'], fill=fill("FFEB9C"), font=Font(name=FONT, bold=True, color="7F6000")))
wm.conditional_formatting.add(pr, FormulaRule(formula=[f'ISNUMBER(I{FIRST_ROW})'], fill=fill("FFC7CE"), font=Font(name=FONT, bold=True, color="9C0006")))
gp = f"H{FIRST_ROW}:H{LAST_EMP_ROW}"
wm.conditional_formatting.add(gp, FormulaRule(formula=[f'AND(ISNUMBER(H{FIRST_ROW}),H{FIRST_ROW}>0)'], font=Font(name=FONT, bold=True, color="C00000")))

dv_lvl = DataValidation(type="list", formula1='"0,1,2,3,4"', allow_blank=True,
                        error="Pick a level from 0 to 4", errorTitle="Invalid level",
                        prompt="0 Not trained | 1 In training | 2 Can do with help | 3 Qualified | 4 Expert/Trainer",
                        promptTitle="Skill level")
dv_role = DataValidation(type="list", formula1=f"={ROLE_LIST}", allow_blank=True)
dv_shift = DataValidation(type="list", formula1=f"={SHIFT_LIST}", allow_blank=True)
for d in (dv_lvl, dv_role, dv_shift):
    wm.add_data_validation(d)
dv_lvl.add(grid)
dv_role.add(f"C{FIRST_ROW}:C{LAST_EMP_ROW}")
dv_shift.add(f"D{FIRST_ROW}:D{LAST_EMP_ROW}")

# team totals rows
tr = LAST_EMP_ROW + 2
wm.cell(tr, 2, "Team totals").font = f(11, True, "1F3864")
labels = [("# Qualified (3+)", ">=3"), ("# Trainers (4)", "=4")]
for i, (lab, crit) in enumerate(labels):
    rr = tr + 1 + i
    wm.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=SKILL_COL - 1)
    c = wm.cell(rr, 2, lab); c.font = f(10, True); c.alignment = Alignment(horizontal="right")
    for j in range(N_SKILLS):
        col = get_column_letter(SKILL_COL + j)
        x = wm.cell(rr, SKILL_COL + j, f'=COUNTIF({col}{FIRST_ROW}:{col}{LAST_EMP_ROW},"{crit}")')
        x.font = f(10, True); x.alignment = CENTER; x.border = BORDER; x.fill = fill("F2F2F2")
rr = tr + 3
wm.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=SKILL_COL - 1)
c = wm.cell(rr, 2, "Team readiness (avg % Ready)"); c.font = f(10, True); c.alignment = Alignment(horizontal="right")
x = wm.cell(rr, SKILL_COL, f'=IFERROR(AVERAGE(I{FIRST_ROW}:I{LAST_EMP_ROW}),"")')
x.number_format = "0%"; x.font = f(11, True, "1F3864")
wm.merge_cells(start_row=rr, start_column=SKILL_COL, end_row=rr, end_column=SKILL_COL + 2)

wm.freeze_panes = wm.cell(FIRST_ROW, SKILL_COL)
wm["C3"].comment = Comment("Role drives the RED gap highlighting. Requirements per role are on the 'Role Requirements' tab.", "Skill Matrix")
wm.auto_filter.ref = f"A5:{LSC}{LAST_EMP_ROW}"

# ================================================================ Skills Library
wsl = wb.create_sheet("Skills Library", 3)
title(wsl, "Skills Library – what each skill means and how it is verified",
      "Use this to train and sign off consistently. Add the site SOP / work instruction number in the yellow column.", 0)
cols = [("ID", 7), ("Category", 22), ("Skill", 32), ("What 'Qualified' (Level 3) looks like", 60),
        ("Training method", 20), ("How it is verified", 24), ("Refresher", 13), ("SOP / WI reference", 18),
        ("# roles requiring", 10)]
for i, (h, w) in enumerate(cols, start=1):
    head(wsl.cell(4, i), h)
    wsl.column_dimensions[get_column_letter(i)].width = w
wsl.row_dimensions[4].height = 30
for k, sk in enumerate(SKILLS):
    rr = 5 + k
    vals = [sk["id"], sk["cat"], sk["name"], sk["qual"], sk["method"], sk["verify"], sk["refresh"], None]
    for i, v in enumerate(vals, start=1):
        c = wsl.cell(rr, i, v); c.font = f(10, i in (1, 3)); c.alignment = LEFT if i in (2, 3, 4, 5, 6) else CENTER
        c.border = BORDER
    wsl.cell(rr, 1).fill = fill(sk["color"]); wsl.cell(rr, 1).font = f(10, True, "FFFFFF")
    wsl.cell(rr, 8).fill = INPUT_FILL
    col = get_column_letter(SKILL_COL + k)
    c = wsl.cell(rr, 9, f"=COUNTIF('Role Requirements'!{col}${FIRST_ROW}:{col}${REQ_LAST},\">0\")")
    c.font = f(10); c.alignment = CENTER; c.border = BORDER
    wsl.row_dimensions[rr].height = 42
wsl.freeze_panes = "D5"
wsl.auto_filter.ref = f"A4:I{4 + N_SKILLS}"

# ================================================================ Shift Coverage
wc = wb.create_sheet("Shift Coverage", 4)
title(wc, "Shift Coverage – do we have enough qualified people on every shift?",
      "Counts people at level 3+ on the Skill Matrix. Set the minimum per shift in the yellow column (default 2).", 0)
ccols = [("ID", 7), ("Category", 22), ("Skill", 34), ("Min per shift", 9),
         ("=Lists!$D$2", 10), ("=Lists!$D$3", 10), ("=Lists!$D$4", 10),
         ("Total qualified", 10), ("Trainers (Level 4)", 10), ("Status", 14)]
for i, (h, w) in enumerate(ccols, start=1):
    head(wc.cell(4, i), h)
    wc.column_dimensions[get_column_letter(i)].width = w
wc.row_dimensions[4].height = 34
mgrid = f"'Skill Matrix'!${FSC}${FIRST_ROW}:${LSC}${LAST_EMP_ROW}"
mshift = f"'Skill Matrix'!$D${FIRST_ROW}:$D${LAST_EMP_ROW}"
mids = f"'Skill Matrix'!${FSC}$4:${LSC}$4"
for k, sk in enumerate(SKILLS):
    rr = 5 + k
    col = get_column_letter(SKILL_COL + k)
    rng = f"'Skill Matrix'!${col}${FIRST_ROW}:${col}${LAST_EMP_ROW}"
    wc.cell(rr, 1, sk["id"]); wc.cell(rr, 2, sk["cat"]); wc.cell(rr, 3, sk["name"])
    wc.cell(rr, 4, 2).fill = INPUT_FILL
    for i, sc in enumerate(("E", "F", "G")):
        wc.cell(rr, 5 + i, f'=COUNTIFS({mshift},{sc}$4,{rng},">=3")')
    wc.cell(rr, 8, f'=COUNTIF({rng},">=3")')
    wc.cell(rr, 9, f'=COUNTIF({rng},4)')
    wc.cell(rr, 10, f'=IF(H{rr}<=1,"HIGH RISK",IF(MIN(E{rr}:G{rr})>=D{rr},"OK","GAP"))')
    for i in range(1, 11):
        c = wc.cell(rr, i); c.border = BORDER; c.font = f(10, i in (1, 10))
        c.alignment = LEFT if i in (2, 3) else CENTER
    wc.cell(rr, 1).fill = fill(sk["color"]); wc.cell(rr, 1).font = f(10, True, "FFFFFF")
srng = f"J5:J{4 + N_SKILLS}"
wc.conditional_formatting.add(srng, CellIsRule(operator="equal", formula=['"OK"'], fill=fill("C6EFCE"), font=Font(name=FONT, bold=True, color="006100")))
wc.conditional_formatting.add(srng, CellIsRule(operator="equal", formula=['"GAP"'], fill=fill("FFEB9C"), font=Font(name=FONT, bold=True, color="7F6000")))
wc.conditional_formatting.add(srng, CellIsRule(operator="equal", formula=['"HIGH RISK"'], fill=fill("FFC7CE"), font=Font(name=FONT, bold=True, color="9C0006")))
shrng = f"E5:G{4 + N_SKILLS}"
wc.conditional_formatting.add(shrng, FormulaRule(formula=["E5<$D5"], font=Font(name=FONT, bold=True, color="C00000"), fill=fill("FFF2CC")))
# summary
sr = 6 + N_SKILLS
wc.cell(sr, 2, "Summary").font = f(11, True, "1F3864")
for i, lab in enumerate(("OK", "GAP", "HIGH RISK")):
    wc.cell(sr + 1 + i, 2, lab).font = f(10, True)
    c = wc.cell(sr + 1 + i, 3, f'=COUNTIF({srng},"{lab}")'); c.font = f(10, True); c.alignment = Alignment(horizontal="left")
wc.freeze_panes = "D5"

# ================================================================ Training Plan
wt = wb.create_sheet("Training Plan", 5)
title(wt, "Training Plan & Sign-off Log",
      "One line per gap. Skill name fills in automatically from the Skill ID. Level 3+ needs trainer AND supervisor sign-off (plus QA for CCP / allergen skills).", 0)
tcols = [("Employee name", 24), ("Skill ID", 9), ("Skill (auto)", 32), ("Current level", 9), ("Target level", 9),
         ("Trainer", 18), ("Start date", 12), ("Due date", 12), ("Status", 16),
         ("Trainer sign-off date", 13), ("Supervisor sign-off date", 13), ("QA sign-off date (if req.)", 13),
         ("Notes", 30)]
for i, (h, w) in enumerate(tcols, start=1):
    head(wt.cell(4, i), h)
    wt.column_dimensions[get_column_letter(i)].width = w
wt.row_dimensions[4].height = 40
lib_ids = f"'Skills Library'!$A$5:$A${4 + N_SKILLS}"
lib_names = f"'Skills Library'!$C$5:$C${4 + N_SKILLS}"
N_PLAN = 100
for k in range(N_PLAN):
    rr = 5 + k
    for i in range(1, 14):
        c = wt.cell(rr, i); c.border = BORDER; c.font = f(10)
        c.alignment = LEFT if i in (1, 3, 6, 13) else CENTER
        if i != 3:
            c.fill = INPUT_FILL
    wt.cell(rr, 3, f'=IF(B{rr}="","",IFERROR(INDEX({lib_names},MATCH(B{rr},{lib_ids},0)),"Unknown ID"))')
    wt.cell(rr, 3).fill = fill("F2F2F2")
    for i in (7, 8, 10, 11, 12):
        wt.cell(rr, i).number_format = "yyyy-mm-dd"
ex = ["EXAMPLE – delete me", "PK05", None, 2, 3, "Line Lead name", "2026-11-09", "2026-11-30", "In progress", None, None, None,
      "3 observed checkweigher verifications needed"]
from datetime import date  # noqa: E402
for i, v in enumerate(ex, start=1):
    if i == 3:
        continue
    if i in (7, 8) and v:
        v = date.fromisoformat(v)
    c = wt.cell(5, i, v); c.font = f(10, italic=True, color="595959")
tdv_skill = DataValidation(type="list", formula1=f"={lib_ids}", allow_blank=True)
tdv_lvl = DataValidation(type="list", formula1='"0,1,2,3,4"', allow_blank=True)
tdv_status = DataValidation(type="list", formula1="=Lists!$F$2:$F$6", allow_blank=True)
for d in (tdv_skill, tdv_lvl, tdv_status):
    wt.add_data_validation(d)
tdv_skill.add(f"B5:B{4 + N_PLAN}")
tdv_lvl.add(f"D5:E{4 + N_PLAN}")
tdv_status.add(f"I5:I{4 + N_PLAN}")
st = f"I5:I{4 + N_PLAN}"
for val, bg, fg in (("Complete", "C6EFCE", "006100"), ("In progress", "FFEB9C", "7F6000"),
                    ("Ready for sign-off", "DDEBF7", "1F4E79"), ("On hold", "D9D9D9", "404040")):
    wt.conditional_formatting.add(st, CellIsRule(operator="equal", formula=[f'"{val}"'], fill=fill(bg), font=Font(name=FONT, bold=True, color=fg)))
# overdue due-date
wt.conditional_formatting.add(f"H5:H{4 + N_PLAN}", FormulaRule(
    formula=['AND(ISNUMBER(H5),H5<TODAY(),$I5<>"Complete")'], fill=fill("FFC7CE"), font=Font(name=FONT, bold=True, color="9C0006")))
wt.freeze_panes = "B5"
wt.auto_filter.ref = f"A4:M{4 + N_PLAN}"

# ---------------------------------------------------------------- page setup
for s in wb.worksheets:
    s.sheet_properties.pageSetUpPr.fitToPage = True
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
    s.print_options.horizontalCentered = True
wm.print_title_rows = "3:5"
wm.print_title_cols = "A:C"
wm.page_setup.paperSize = wm.PAPERSIZE_TABLOID
for name, color in (("How To Use", "1F3864"), ("Skill Matrix", "548235"), ("Role Requirements", "7030A0"),
                    ("Skills Library", "2E75B6"), ("Shift Coverage", "BF8F00"), ("Training Plan", "C55A11"),
                    ("Lists", "808080")):
    wb[name].sheet_properties.tabColor = color
wb.move_sheet("Lists", offset=len(wb.sheetnames))
wb.active = 0
wb.save(OUT)
print(f"Saved {OUT} ({N_SKILLS} skills, {len(ROLES)} roles)")
